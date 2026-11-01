#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""合同审核预检、一次性转换/提取与本地缓存。

该脚本把旧版 ``.doc`` 转换和正文提取合并成一个可复用步骤：

* 对源文件计算 SHA-256；同一源文件再次审核时复用 ``base.docx``、
  ``blocks.json`` 和 ``contract.txt``；
* 预先识别原文已有修订/批注、空白字段、金额/电话多口径、旧日期、旧
  “合同法”表述、停服/责任上限/数据条款等固定模式；
* 只输出索引和命中摘要，不把整份合同打印到终端；
* 将每个阶段耗时写入 ``work/preflight.json``，方便后续复盘。

脚本不会修改用户原始文件。它只调用 ``convert_input.py`` 和
``docx_extract.py`` 的公共函数，避免工作流先转 TXT、再重复转换 DOCX。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import time
from collections import Counter
from datetime import date
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from convert_input import convert_to_docx  # noqa: E402
from docx_extract import extract_docx  # noqa: E402
from stage_mark import record_stage_mark  # noqa: E402


WORKFLOW_VERSION = "preflight-1.6"

# 预检缓存必须随影响结果的规则、模板和脚本变化失效。仅比较源合同哈希会
# 导致“代码已更新但沿用旧 blocks/context”的隐性错误，因此把这些小型输入
# 固化为运行指纹写入 manifest。
_CACHE_INPUTS = (
    "preflight_contract.py",
    "docx_extract.py",
    "convert_input.py",
    "lib/ooxml.py",
    "lib/kb.py",
    "cite_law.py",
    "assets/schemas/review-draft.schema.json",
    "references/policies/review-quickref.md",
    "references/policies/group-redlines.md",
    "references/workflows/review.md",
)

_LEGAL_CELL_RE = re.compile(
    r"(?:甲\s*方|乙\s*方|地\s*址|联系人|电\s*话|传\s*真|开户行|帐号|账号|税号|开票|签署|盖章|代表|生效|"
    r"元|¥|￥|万元|总价|总费用|费用|价格|小计|合计|总计|单价|金额|价款|折扣|优惠|付款|支付|账期|发票|税率|增值税|预付|尾款|质保金|保证金|"
    r"交付|验收|上线|试运行|终验|初验|里程碑|工期|人/天|人天|延期|逾期|故障|响应|修复|维护|维保|质保|"
    r"违约|赔偿|责任|上限|直接损失|间接损失|解除|终止|中止|停止|暂停|停服|锁定|"
    r"知识产权|著作权|专利|源代码|许可|保密|商业秘密|个人信息|数据|安全|漏洞|分包|转包|贿赂|廉洁|不可抗力|管辖|诉讼|仲裁|送达|"
    r"(?:19|20)\d{2}[年./-])",
    re.I,
)

_CONTRACT_TYPE_RULES: list[dict[str, Any]] = [
    {
        "code": "tech",
        "name": "技术开发 / 软件外包 / SaaS / 技术服务",
        "keywords": ["软件", "系统", "技术开发", "技术服务", "实施服务", "e-cology", "SaaS", "云服务", "信息化", "源代码", "上线", "并发", "用户许可", "运维", "算法"],
        "checklists": [("C0202", "02_技术开发与技术服务（含软件外包_SaaS）审查清单")],
        "sectors": [],
        "group_sections": [],
        "templates": ["T0105"],
        "preset_laws": "民法典:851,852,853,854,858,859,860,861;著作权法:19;网络安全法:24,61,62;网络数据安全管理条例:9,10,11,12,16,21",
    },
    {
        "code": "sale",
        "name": "买卖 / 采购 / 设备采购安装",
        "keywords": ["买卖", "采购", "购销", "供货", "设备", "服务器", "安装调试", "购买", "产品购买", "交货", "质保期"],
        "checklists": [("C0201", "01_买卖与采购合同审查清单")],
        "sectors": [],
        "group_sections": ["十"],
        "templates": ["T0108"],
        "preset_laws": "民法典:596,600,615,620,621,622,623,624;产品质量法:26,40;反不正当竞争法:8,15,24,31;最高人民法院关于适用《中华人民共和国民法典》侵权责任编的解释（一）:19",
    },
    {
        "code": "medical-procurement",
        "name": "医疗设备采购维保 / 药品耗材 SPD",
        "keywords": ["医疗设备", "医疗器械", "耗材", "SPD", "医院采购", "体外诊断"],
        "checklists": [("C0201", "01_买卖与采购合同审查清单")],
        "sectors": [("S0202", "大型医疗设备采购维保与药品耗材SPD供应链合规审查清单")],
        "group_sections": ["十"],
        "templates": ["T0108"],
        "preset_laws": "医疗卫生机构医药采购与医疗损害责任核心法条汇编:1;反不正当竞争法:8,24",
    },
    {
        "code": "food-procurement",
        "name": "中央厨房 / 生鲜冷链 / 食材采购",
        "keywords": ["食材", "生鲜", "冷链", "中央厨房", "食品采购", "食品安全", "农产品采购"],
        "checklists": [("C0201", "01_买卖与采购合同审查清单")],
        "sectors": [("S0301", "中央厨房生鲜冷链采购、食品安全“退一赔十”防控与灵活用工合规清单")],
        "group_sections": ["十"],
        "templates": ["T0108"],
        "preset_laws": "食品安全法:41,53,81,124,132,148;最高人民法院关于审理食品药品惩罚性赔偿纠纷案件适用法律若干问题的解释:1,2;农产品质量安全法:36",
    },
    {
        "code": "commodity-trade",
        "name": "大宗商品贸易",
        "keywords": ["大宗商品", "贸易", "钢材", "煤炭", "有色金属", "化工品", "仓单", "供应链贸易", "融资性贸易"],
        "checklists": [("C0201", "01_买卖与采购合同审查清单")],
        "sectors": [("S0402", "大宗商品贸易合规审查清单（严防融资性贸易、循环走单与虚开风险）")],
        "group_sections": ["十"],
        "templates": [],
        "preset_laws": "民法典:533,596,604,641;产品质量法:26,40",
    },
    {
        "code": "intl-trade",
        "name": "国际货物买卖（进出口 / FOB / CIF / 信用证）",
        "keywords": ["进口", "出口", "外贸", "FOB", "CIF", "信用证", "Incoterms", "国际货物", "提单"],
        "checklists": [("C0201", "01_买卖与采购合同审查清单")],
        "sectors": [("S0401", "国际货物买卖（进出口外贸 FOB_CIF_信用证）审查清单与中英双语核心条款")],
        "group_sections": ["十"],
        "templates": [],
        "preset_laws": "民法典:596,603,604,620,621",
    },
    {
        "code": "lease",
        "name": "房屋 / 商业地产 / 产业园区租赁",
        "keywords": ["租赁", "出租", "承租", "租金", "写字楼", "商铺", "厂房", "园区", "公寓", "宿舍", "免租期", "押金", "腾退"],
        "checklists": [("C0204", "04_房屋与商业地产租赁合同审查清单")],
        "sectors": [],
        "group_sections": ["九"],
        "templates": ["T0103"],
        "preset_laws": "民法典:705,715,716,717,719,725,726;住房租赁条例:7,8,10,12,19,39,43;最高人民法院关于审理城镇房屋租赁合同纠纷案件具体应用法律若干问题的解释:2,3,7,12,13,14;物业管理条例:41,44;最高人民法院关于审理预付式消费民事纠纷案件适用法律若干问题的解释:6",
    },
    {
        "code": "property-mgmt",
        "name": "物业管理服务",
        "keywords": ["物业服务", "物业管理", "物业费", "前期物业", "公共能耗", "转供电"],
        "checklists": [("C0204", "04_房屋与商业地产租赁合同审查清单")],
        "sectors": [("S0702", "商业物业管理服务合同与MCN直播带货_品牌营销代运营合同审查清单")],
        "group_sections": ["九"],
        "templates": ["T0103"],
        "preset_laws": "民法典:942,944;物业管理条例:41,44;最高人民法院关于审理物业服务及建筑物区分所有权纠纷案件适用法律若干问题的解释:2",
    },
    {
        "code": "construction",
        "name": "建设工程施工（总包 / 专业分包 / 劳务分包）",
        "keywords": ["施工", "总承包", "专业分包", "劳务分包", "工程款", "竣工", "进度款", "结算", "优先受偿权", "工期"],
        "checklists": [],
        "sectors": [("S0102", "建设工程施工合同（总包_专业分包_劳务分包）审查清单与核心条款范本")],
        "group_sections": ["十一"],
        "templates": [],
        "preset_laws": "民法典:791,793,799,807;最高人民法院关于审理建设工程施工合同纠纷案件适用法律问题的解释（一）:1,9,14,17,19,21,35,38;最高人民法院关于审理建设工程施工合同纠纷案件适用法律问题的解释（二）:1,5,6,7,8,13,14,16,17;安全生产法:49",
    },
    {
        "code": "epc",
        "name": "EPC 工程总承包 / 勘察设计 / 大宗物业收购",
        "keywords": ["EPC", "工程总承包", "勘察", "设计", "物业收购", "资产收购", "在建工程"],
        "checklists": [],
        "sectors": [("S0101", "EPC工程总承包与设计勘察及大宗物业收购合规审查清单")],
        "group_sections": ["十一"],
        "templates": [],
        "preset_laws": "民法典:791,795,807;最高人民法院关于审理建设工程施工合同纠纷案件适用法律问题的解释（一）:1,19,35;最高人民法院关于审理建设工程施工合同纠纷案件适用法律问题的解释（二）:1,5,13,16",
    },
    {
        "code": "outsourcing",
        "name": "业务外包（BPO）/ 劳务外包 / 劳务派遣",
        "keywords": ["业务外包", "BPO", "劳务外包", "劳务派遣", "人力外包", "灵活用工", "驻场服务"],
        "checklists": [("C0303", "03_企业劳动用工全流程合规审查清单")],
        "sectors": [("S0701", "业务外包（BPO）_劳务外包与劳务派遣合规边界审查清单（防假外包真派遣）")],
        "group_sections": [],
        "templates": [],
        "preset_laws": "劳动合同法:58,59,62,66;最高人民法院关于审理劳动争议案件适用法律问题的解释（二）:1,2,3",
    },
    {
        "code": "labor",
        "name": "劳动合同 / 竞业限制 / 用工文书",
        "keywords": ["劳动合同", "试用期", "竞业限制", "服务期", "解除劳动合同", "员工", "入职", "离职", "经济补偿"],
        "checklists": [
            ("C0303", "03_企业劳动用工全流程合规审查清单"),
            ("C0203", "03_保密协议（NDA）与竞业限制审查清单"),
        ],
        "sectors": [],
        "group_sections": [],
        "templates": ["T0102"],
        "preset_laws": "劳动合同法:19,22,23,24,38,39,40;最高人民法院关于审理劳动争议案件适用法律问题的解释（二）:10,12,13,14,15,19",
    },
    {
        "code": "nda",
        "name": "保密协议（NDA）",
        "keywords": ["保密协议", "NDA", "保密信息", "披露方", "接收方"],
        "checklists": [("C0203", "03_保密协议（NDA）与竞业限制审查清单")],
        "sectors": [],
        "group_sections": [],
        "templates": ["T0104"],
        "preset_laws": "民法典:501;反不正当竞争法:10,22,39",
    },
    {
        "code": "data",
        "name": "数据委托处理 / DPA / 数据跨境",
        "keywords": ["数据处理协议", "DPA", "委托处理", "数据出境", "数据跨境", "敏感个人信息"],
        "checklists": [("C0302", "02_企业个人信息保护与数据合规自查清单")],
        "sectors": [],
        "group_sections": [],
        "templates": ["T0101"],
        "preset_laws": "个人信息保护法:21,28,29,30,38,51;网络数据安全管理条例:9,10,11,12,21,31,35,37;数据安全法:27,29;促进和规范数据跨境流动规定:3,4,5",
    },
    {
        "code": "equity-transfer",
        "name": "股权转让",
        "keywords": ["股权转让", "转让股权", "受让方", "出资", "认缴", "目标公司"],
        "checklists": [("C0301", "01_新《公司法》公司章程及治理结构合规自查清单")],
        "sectors": [("S0502", "集团投资并购法律尽职调查（LDD）标准核查清单")],
        "group_sections": [],
        "templates": ["T0107"],
        "preset_laws": "公司法:47,54,84,86,88;最高人民法院关于适用《中华人民共和国公司法》时间效力的若干规定:1,4;全国法院民商事审判工作会议纪要:1",
    },
    {
        "code": "capital-increase",
        "name": "增资扩股 / 股东协议 / 对赌",
        "keywords": ["增资", "投资协议", "股东协议", "对赌", "回购", "反稀释", "一票否决", "领售", "随售", "业绩补偿"],
        "checklists": [("C0301", "01_新《公司法》公司章程及治理结构合规自查清单")],
        "sectors": [
            ("S0501", "增资扩股与股东协议（含对赌回购、一票否决、反稀释、领售随售）审查清单与范本"),
            ("S0502", "集团投资并购法律尽职调查（LDD）标准核查清单"),
        ],
        "group_sections": [],
        "templates": ["T0106"],
        "preset_laws": "公司法:15,47,54,88,162,224;全国法院民商事审判工作会议纪要:1;最高人民法院关于适用《中华人民共和国公司法》时间效力的若干规定:1",
    },
    {
        "code": "articles",
        "name": "公司章程 / 股东会与董事会决议",
        "keywords": ["章程", "股东会决议", "董事会决议", "催缴", "失权", "减资", "法定代表人"],
        "checklists": [("C0301", "01_新《公司法》公司章程及治理结构合规自查清单")],
        "sectors": [],
        "group_sections": [],
        "templates": ["T0106"],
        "preset_laws": "公司法:10,11,15,47,51,52,54,59,66,67,224;最高人民法院关于适用《中华人民共和国公司法》时间效力的若干规定:1",
    },
    {
        "code": "franchise",
        "name": "特许经营 / 连锁加盟",
        "keywords": ["特许经营", "加盟", "连锁", "品牌授权", "加盟费", "特许人", "被特许人"],
        "checklists": [],
        "sectors": [("S0302", "餐饮品牌特许经营（连锁加盟）合同审查清单与标准范本")],
        "group_sections": [],
        "templates": [],
        "preset_laws": "商业特许经营管理条例:7,8,11,12,13,21,22,23;最高人民法院关于审理预付式消费民事纠纷案件适用法律若干问题的解释:5,6,23;食品安全法:41,148;反不正当竞争法:8,10",
    },
    {
        "code": "oem",
        "name": "OEM / ODM 委托加工 / 代工生产",
        "keywords": ["OEM", "ODM", "委托加工", "代工", "模具", "贴牌", "定作"],
        "checklists": [("C0201", "01_买卖与采购合同审查清单")],
        "sectors": [("S0601", "OEM_ODM委托加工与代工生产合同审查清单及模具归属_质量召回核心范本")],
        "group_sections": ["十"],
        "templates": [],
        "preset_laws": "民法典:770,774,781,787;产品质量法:26,40;最高人民法院关于适用《中华人民共和国民法典》侵权责任编的解释（一）:19",
    },
    {
        "code": "ehs",
        "name": "外包施工入厂 / EHS 安全管理协议",
        "keywords": ["安全生产", "EHS", "入厂", "外包施工", "动火", "安全管理协议", "有限空间", "交叉作业"],
        "checklists": [],
        "sectors": [("S0602", "工厂EHS安全生产管理及外包施工入厂安全管理协议标准范本")],
        "group_sections": [],
        "templates": [],
        "preset_laws": "安全生产法:48,49,103;最高人民法院关于适用《中华人民共和国民法典》侵权责任编的解释（一）:15,16,17,18",
    },
    {
        "code": "mcn",
        "name": "MCN 直播带货 / 品牌营销代运营",
        "keywords": ["MCN", "直播", "带货", "代运营", "达人", "主播", "坑位费", "投流"],
        "checklists": [],
        "sectors": [("S0702", "商业物业管理服务合同与MCN直播带货_品牌营销代运营合同审查清单")],
        "group_sections": [],
        "templates": [],
        "preset_laws": "反不正当竞争法:8,9,11,24;个人信息保护法:21;民法典:919,928,933",
    },
    {
        "code": "clinical",
        "name": "临床科研合作（GCP / CRO）/ 医疗纠纷和解",
        "keywords": ["临床试验", "GCP", "CRO", "科研合作", "医疗纠纷", "和解协议", "受试者", "患方"],
        "checklists": [],
        "sectors": [("S0201", "医院临床科研合作（GCP_CRO）、九项准则反腐及医疗纠纷和解协议模板")],
        "group_sections": [],
        "templates": [],
        "preset_laws": "医疗纠纷预防和处理条例:13,15,22,23,31;医疗卫生机构医药采购与医疗损害责任核心法条汇编:1",
    },
    {
        "code": "agri-land",
        "name": "农村土地经营权流转 / 订单农业 / 农资种业采购",
        "keywords": ["土地流转", "土地经营权", "订单农业", "种子", "化肥", "农资", "规模化种植"],
        "checklists": [],
        "sectors": [("S0801", "农村土地经营权流转（规模化种植）_订单农业与农资种业采购合规审查清单")],
        "group_sections": [],
        "templates": [],
        "preset_laws": "农产品质量安全法:27,29,36;粮食安全保障法:18,24;民法典:339,340,341",
    },
    {
        "code": "grain-storage",
        "name": "政策性储备粮承储 / 委托代储 / 轮换购销",
        "keywords": ["储备粮", "代储", "承储", "轮换", "粮库", "粮食收购", "政策性粮食"],
        "checklists": [],
        "sectors": [("S0802", "政策性储备粮承储_委托代储与轮换购销合同审查清单及标准范本")],
        "group_sections": [],
        "templates": [],
        "preset_laws": "粮食流通管理条例:19,22,26,49;政府储备粮食仓储管理办法:14,19,29;粮食安全保障法:29,30,31",
    },
]

_BASE_PRESET_LAWS = "民法典:470,496,497,501,506,509,511,527,528,562,563,565,566,577,581,584,585,590;最高人民法院关于适用《民法典》合同编通则若干问题的解释:9,10,64,65"


_CATALOG_HEADER_RE = re.compile(r"功能描述|功能清单|优势功能|模块功能|产品功能|技术参数说明")
_COMMERCIAL_HEADER_RE = re.compile(
    r"单位\s*[：:]\s*元|费用|价格|金额|小计|合计|总计|付款|开户行|甲\s*方|乙\s*方|工期|上线标准|验收标准"
)
_CATALOG_KEEP_CELL_RE = re.compile(
    r"(?:甲\s*方|乙\s*方|本合同|非标|另行\s*(?:收费|计费|采购|购买|协商|签署)|不含在|需第三方|"
    r"不承担|不负.*责任|免责|概不负责|自行承担|风险由.*承担|赔偿|违约|责任上限|"
    r"质保期|维保期|免费维护|视为.*合格|著作权|知识产权|源代码|开源|GPL|AGPL|"
    r"付款条件|付款方式|付款节点|逾期付款|预付|尾款|"
    r"(?<!\d)\d+(?:,\d{3})*(?:\.\d+)?\s*(?:元|万元)|[¥￥]\s*\d)",
    re.I,
)


def build_compact_contract_text(blocks: list[dict[str, Any]]) -> tuple[str, dict[str, Any]]:
    """生成折叠产品功能大表后的紧凑合同视图 (contract_compact.txt)。

    规则：
    1. 所有正文段落（p-xxxx 及页眉页脚）100% 保留；
    2. 非产品功能描述类表格（如商务/签署/价格/付款/工期/工程量/验收/附件表）100% 保留；
    3. 产品功能/技术参数描述清单表（is_catalog 为真）：
       - 保留表头行（r0、以及 r0 为单格横幅时的 r1）；
       - 保留含金额、付款、质保、权责、免责、知产/开源、非标声明或既有修订/批注的整行；
       - 其余纯功能描述行折叠为 1 行表格摘要，提示可用 docx_extract.py --table 展开。
    """
    table_cell_re = re.compile(r"^(?P<table>t\d+)-r(?P<row>\d+)-c(?P<col>\d+)")
    tables: dict[str, dict[int, list[dict[str, Any]]]] = {}
    table_order: list[str] = []

    for b in blocks:
        text_val = str(b.get("text") or "").strip()
        label_val = str(b.get("list_label") or "").strip()
        if not text_val and not label_val:
            continue
        m = table_cell_re.match(str(b.get("block_id") or ""))
        if not m:
            continue
        t_id = m.group("table")
        r_idx = int(m.group("row"))
        if t_id not in tables:
            tables[t_id] = {}
            table_order.append(t_id)
        tables[t_id].setdefault(r_idx, []).append(b)

    retained_rows: dict[str, set[int]] = {}
    fold_summaries: dict[str, str] = {}
    folded_table_count = 0
    folded_cell_count = 0
    total_table_cells = 0
    fold_manifest: list[dict[str, Any]] = []

    for t_id in table_order:
        row_map = tables[t_id]
        sorted_rows = sorted(row_map.keys())
        non_empty_cells = sum(len(cells) for cells in row_map.values())
        total_table_cells += non_empty_cells

        header_rows = sorted_rows[:2]
        header_text = " ".join(
            str(c.get("text") or "").strip()
            for r in header_rows
            for c in row_map[r]
        )
        is_catalog = bool(_CATALOG_HEADER_RE.search(header_text)) and not bool(
            _COMMERCIAL_HEADER_RE.search(header_text)
        )

        # 严格仅对“产品功能/技术参数描述清单表”执行折叠；其余商务/工程/履约表格 100% 保留
        if not is_catalog:
            retained_rows[t_id] = set(sorted_rows)
            continue

        keep: set[int] = set()
        if sorted_rows:
            keep.add(sorted_rows[0])
            if len(sorted_rows) > 1 and len(row_map[sorted_rows[0]]) == 1:
                keep.add(sorted_rows[1])

        matcher = _CATALOG_KEEP_CELL_RE
        for r_idx in sorted_rows:
            if r_idx in keep:
                continue
            row_cells = row_map[r_idx]
            row_joined = " ".join(str(c.get("text") or "") for c in row_cells)
            has_rev = any(bool(c.get("has_existing_revisions") or c.get("has_comments")) for c in row_cells)
            if has_rev or matcher.search(row_joined):
                keep.add(r_idx)

        kept_cells = sum(len(row_map[r]) for r in keep)
        omitted_cells = non_empty_cells - kept_cells
        if omitted_cells <= 6:
            retained_rows[t_id] = set(sorted_rows)
            continue

        retained_rows[t_id] = keep
        folded_table_count += 1
        folded_cell_count += omitted_cells
        sample_titles: list[str] = []
        for r_idx in sorted_rows:
            if r_idx in keep:
                continue
            first_cell = str(row_map[r_idx][0].get("text") or "").strip().replace("\n", " ")
            if first_cell and len(first_cell) <= 30 and first_cell not in sample_titles:
                sample_titles.append(first_cell)
            if len(sample_titles) >= 5:
                break
        preview = "、".join(sample_titles) if sample_titles else "常规功能/参数描述"
        fold_summaries[t_id] = (
            f"[{t_id}|折叠表格摘要] 表 {t_id} 共 {non_empty_cells} 个非空单元格，"
            f"已保留表头及关键行（{kept_cells} 格），折叠常规功能/描述 {omitted_cells} 格"
            f"（含：{preview} 等；如需展开请运行 docx_extract.py <blocks.json> --table {t_id}）"
        )
        fold_manifest.append({
            "table_id": t_id,
            "total_non_empty_cells": non_empty_cells,
            "kept_rows": sorted(keep),
            "omitted_rows": [r for r in sorted_rows if r not in keep],
            "kept_cells": kept_cells,
            "omitted_cells": omitted_cells,
            "reason": "catalog_table_non_key_rows_folded",
            "expand_command": f"docx_extract.py <blocks.json> --table {t_id}",
        })

    lines: list[str] = []
    emitted_summary: set[str] = set()
    for b in blocks:
        text_val = str(b.get("text") or "")
        label_val = str(b.get("list_label") or "")
        if not text_val.strip() and not label_val.strip():
            continue
        bid = str(b.get("block_id") or "")
        m = table_cell_re.match(bid)
        if m:
            t_id = m.group("table")
            r_idx = int(m.group("row"))
            if r_idx not in retained_rows.get(t_id, {r_idx}):
                if t_id in fold_summaries and t_id not in emitted_summary:
                    lines.append(fold_summaries[t_id])
                    emitted_summary.add(t_id)
                continue
        clause_tag = b.get("clause_ref") or label_val or ""
        single_line = text_val.replace("\r\n", "↵").replace("\n", "↵").replace("\r", "↵")
        lines.append(f"[{bid}|{clause_tag}] {single_line}")

    compact_text = "\n".join(lines) + ("\n" if lines else "")
    stats = {
        "compact_chars": len(compact_text),
        "compact_lines": len(lines),
        "folded_tables": folded_table_count,
        "folded_cells": folded_cell_count,
        "total_table_cells": total_table_cells,
        "fold_manifest": fold_manifest,
    }
    return compact_text, stats


def _infer_routing(
    blocks: list[dict[str, Any]],
    contract_text: str,
    check_result: dict[str, Any],
) -> dict[str, Any]:
    """根据合同首部与正文关键词确定性推断合同类型、必查清单与候选法条。"""
    head_text = "\n".join(str(b.get("text") or "") for b in blocks[:30])
    scored: list[tuple[int, dict[str, Any], list[str], list[dict[str, Any]]]] = []
    for rule in _CONTRACT_TYPE_RULES:
        score = 0
        matched_kws: list[str] = []
        evidence: list[dict[str, Any]] = []
        for kw in rule["keywords"]:
            in_head = kw.lower() in head_text.lower()
            in_body = kw.lower() in contract_text.lower()
            if in_head:
                score += 4
                matched_kws.append(kw)
                block = next((b for b in blocks[:30] if kw.lower() in str(b.get("text") or "").lower()), {})
                evidence.append({"keyword": kw, "location": "head", "block_id": block.get("block_id")})
            elif in_body:
                score += 1
                matched_kws.append(kw)
                block = next((b for b in blocks if kw.lower() in str(b.get("text") or "").lower()), {})
                evidence.append({"keyword": kw, "location": "body", "block_id": block.get("block_id")})
        if score > 0:
            scored.append((score, rule, matched_kws, evidence))

    scored.sort(key=lambda x: x[0], reverse=True)
    selected_rules: list[dict[str, Any]] = []
    if scored:
        top_score = scored[0][0]
        for idx, (s, r, _kws, _evidence) in enumerate(scored[:3]):
            has_head_hit = any(kw.lower() in head_text.lower() for kw in r["keywords"])
            if idx == 0 or (s >= 4 and (has_head_hit or s >= max(6, top_score // 2))):
                selected_rules.append(r)
    if not selected_rules:
        selected_rules = [_CONTRACT_TYPE_RULES[0]]

    checklists: list[dict[str, str]] = [
        {"id": "group-redlines", "title": "集团法务强制红线速查表（group-redlines.md）", "reason": "全合同必查底座"},
        {"id": "C0102", "title": "集团法务合同审核要点", "reason": "全合同必查底座"},
        {"id": "C0101", "title": "通用商事合同审查总清单（主审与复核双视角）", "reason": "全合同必查底座"},
    ]
    seen_ids = {c["id"] for c in checklists}

    def add_chk(cid: str, title: str, reason: str) -> None:
        if cid not in seen_ids:
            seen_ids.add(cid)
            checklists.append({"id": cid, "title": title, "reason": reason})

    preset_law_specs = [_BASE_PRESET_LAWS]
    group_sections: list[str] = []
    for rule in selected_rules:
        for cid, ctitle in rule.get("checklists", []):
            add_chk(cid, ctitle, f"命中合同类型 {rule['code']}（{rule['name']}）")
        for sid, stitle in rule.get("sectors", []):
            add_chk(sid, stitle, f"命中板块专属库 {rule['code']}")
        for gsec in rule.get("group_sections", []):
            if gsec not in group_sections:
                group_sections.append(gsec)
        if rule.get("preset_laws"):
            preset_law_specs.append(rule["preset_laws"])

    # 专项触发层
    data_check = next((c for c in check_result.get("checks", []) if c.get("code") == "DATA_TERMS"), {})
    if data_check.get("severity") in ("high", "review"):
        add_chk("C0302", "02_企业个人信息保护与数据合规自查清单", "触发个人信息与数据安全专项")
        preset_law_specs.append("个人信息保护法:21,51;网络数据安全管理条例:9,10,11,12,21;网络安全法:61,62")

    has_non_compete = bool(re.search(r"竞业限制|劳动合同|试用期|离职补偿", contract_text))
    if re.search(r"保密|商业秘密|源代码|竞业", contract_text):
        add_chk("C0203", "03_保密协议（NDA）与竞业限制审查清单", "含保密/商业秘密/不招揽条款")
        if has_non_compete:
            preset_law_specs.append("反不正当竞争法:10,22,39;劳动合同法:23,24")
        else:
            preset_law_specs.append("反不正当竞争法:10,22,39")

    if re.search(r"派驻|驻场|外包人员|劳务派遣|劳务外包|灵活用工", contract_text):
        add_chk("C0303", "03_企业劳动用工全流程合规审查清单", "含驻场/外包用工边界")
        add_chk("S0701", "业务外包（BPO）_劳务外包与劳务派遣合规边界审查清单（防假外包真派遣）", "含驻场/外包用工边界")

    if re.search(r"对外担保|担保人|保证人|连带责任保证|一般保证|抵押|质押|债务加入|股东会决议|董事会决议", contract_text):
        add_chk("C0301", "01_新《公司法》公司章程及治理结构合规自查清单", "含担保/公司治理触发词")
        preset_law_specs.append("公司法:15,23,47,54,88")

    top_score = scored[0][0] if scored else 0
    second_score = scored[1][0] if len(scored) > 1 else 0
    margin = top_score - second_score
    confidence_score = round(
        min(1.0, top_score / 12.0)
        * (1.0 if not second_score else min(1.0, margin / max(top_score, 1))),
        3,
    )
    requires_confirmation = (not scored) or top_score < 4 or margin < 2
    if not scored:
        confidence_reason = "未命中合同类型关键词，已采用通用兜底路由"
    elif top_score < 4:
        confidence_reason = "最高命中分低于可靠阈值"
    elif margin < 2:
        confidence_reason = "第一、第二候选路由分差过小"
    else:
        confidence_reason = "首部关键词命中且与次优路由有足够分差"

    return {
        "primary_type": selected_rules[0]["code"],
        "primary_type_name": selected_rules[0]["name"],
        "matched_types": [
            {"code": r["code"], "name": r["name"], "score": s, "keywords": kws[:6], "evidence": evidence[:8]}
            for s, r, kws, evidence in scored[:3]
        ],
        "confidence_score": confidence_score,
        "confidence_level": "high" if not requires_confirmation and confidence_score >= 0.65 else ("medium" if confidence_score >= 0.35 else "low"),
        "routing_confirmation_required": requires_confirmation,
        "confidence_reason": confidence_reason,
        "checklists": checklists,
        "group_sections": group_sections,
        "has_non_compete": has_non_compete,
        "preset_law_spec": ";".join(preset_law_specs),
    }


def _extract_c0102_special_sections(group_sections: list[str]) -> list[str]:
    """按命中的合同类型抽取 C0102《集团法务合同审核要点》第九（租赁）、十（采购）、十一（建工）专项章节完整原文。"""
    if not group_sections:
        return []
    try:
        from lib.kb import entry_by_id, read_entry_text
        ent = entry_by_id("C0102")
        if not ent:
            return []
        raw = read_entry_text(ent)
    except (FileNotFoundError, ValueError):
        raise
    except Exception:
        return []

    section_headers = {
        "九": "## 九、",
        "十": "## 十、",
        "十一": "## 十一、",
    }
    lines = raw.splitlines()
    out: list[str] = []
    for sec_key in group_sections:
        prefix = section_headers.get(sec_key)
        if not prefix:
            continue
        capturing = False
        sec_lines: list[str] = []
        for line in lines:
            if line.startswith(prefix):
                capturing = True
                sec_lines.append(f"### `C0102` 专项章节：{line.lstrip('#').strip()}")
                continue
            if capturing:
                if line.startswith("## "):
                    break
                sec_lines.append(line)
        if sec_lines:
            out.extend(sec_lines)
            out.append("")
    return out


def _extract_checklist_digest(
    checklist_ids: list[str],
    group_sections: list[str] | None = None,
    *,
    has_non_compete: bool = True,
) -> list[str]:
    """无损内联命中的专类/专项审查清单与板块专属库（保留全部核查要点、甲乙方修改策略列及底部核心条款范本）。"""
    try:
        from lib.kb import entry_by_id, read_entry_text
    except Exception:
        return []

    digest_lines: list[str] = []
    if group_sections:
        digest_lines.extend(_extract_c0102_special_sections(group_sections))

    for cid in checklist_ids:
        if cid in ("group-redlines", "C0101", "C0102"):
            continue
        try:
            ent = entry_by_id(cid)
            if not ent:
                continue
            raw = read_entry_text(ent)
        except (FileNotFoundError, ValueError):
            raise
        except Exception:
            continue

        digest_lines.append(f"### `{cid}` {ent.get('title', cid)}（完整核查表与修改策略/范本）")
        started_body = False
        for line in raw.splitlines():
            s = line.strip()
            if not started_body:
                # 跳过文件顶部的 H1 标题与适用范围/法源导语引用块，从第一个二级标题或表格开始完整保留
                if s.startswith("## ") or s.startswith("|"):
                    started_body = True
                else:
                    continue
            # 普通 B2B 商事合同未涉及员工离职竞业限制时，C0203 仅需保留第一部分（商业保密协议 NDA 核查表）
            if cid == "C0203" and not has_non_compete and s.startswith("## 二、"):
                break
            if line.startswith("## "):
                digest_lines.append(f"#### {line[3:].strip()}")
            else:
                digest_lines.append(line)
        digest_lines.append("")
    return digest_lines


def _prefetch_candidate_laws(spec: str) -> tuple[dict[str, Any], list[str]]:
    """根据路由预设批量提取候选法条原文摘要，供模型直接引用而无需逐条调用 cite_law.py。"""
    try:
        from cite_law import cite_law_batch
    except Exception:
        error = {"status": "error", "error": "cite_law_batch 不可用"}
        return error, ["- **法条预取失败**：本地 cite_law 模块不可用，所有法源需人工核验。"]

    # 合并同一部法规的条号，避免重复查询
    merged: dict[str, list[str]] = {}
    for part in spec.split(";"):
        part = part.strip()
        if not part or ":" not in part:
            continue
        law_k, arts = part.split(":", 1)
        law_k = law_k.strip()
        for a in arts.split(","):
            a = a.strip()
            if a and a not in merged.setdefault(law_k, []):
                merged[law_k].append(a)

    normalized_spec = ";".join(f"{k}:{','.join(v)}" for k, v in merged.items())
    try:
        _code, batch_payload = cite_law_batch([normalized_spec])
    except (FileNotFoundError, ValueError):
        # 知识库缺失或条目 SHA256 不一致属于环境/数据完整性错误，必须
        # 阻断预检，不能降级成“库外法源”继续生成审查结论。
        raise
    except Exception as exc:
        error = {"status": "error", "error": f"{type(exc).__name__}: {exc}"}
        return error, ["- **法条预取失败**：本地知识库查询异常，严禁据此编造法条原文；请人工核验。"]

    card_lines: list[str] = []
    seen_keys: set[tuple[str, str]] = set()
    for res in batch_payload.get("results") or []:
        if res.get("status") == "out_of_kb":
            q = res.get("query") or {}
            card_lines.append(
                f"- **《{q.get('law')}》第{q.get('article')}条【库外法源·待人工核实】**：库外法源，仅允许点名条号并带此标注，严禁编造原文。"
            )
            continue
        for art in res.get("articles") or []:
            key = (str(art.get("law_id")), str(art.get("label")))
            if key in seen_keys:
                continue
            seen_keys.add(key)
            law_title = str(art.get("law_title") or "").split("_")[0]
            label = str(art.get("label") or "")
            raw_txt = re.sub(r"\s+", " ", str(art.get("text") or "")).strip()
            preview = raw_txt[:600] + ("…" if len(raw_txt) > 600 else "")
            excerpt_tag = "（专题汇编节选）" if art.get("excerpt") else ""
            # 这是节选卡片，不把截断文本描述为“完整原文”。完整原文仍须
            # 通过 cite_law.py 的单条查询取得并记录出处。
            card_lines.append(f"- **《{_safe_md_inline(law_title, 100)}》{_safe_md_inline(label, 60)}**{excerpt_tag}：原文节选（最多600字，仅供定位）：{_safe_md_inline(preview, 600)}")
    return batch_payload, card_lines


def _safe_md_inline(value: Any, limit: int = 240) -> str:
    """将合同/法条中的不可信文本安全地放入 Markdown 行内代码。"""
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    text = text.replace("`", "'")
    return text[:limit] + ("…" if len(text) > limit else "")


def _build_context_pack(
    work: Path,
    source: Path,
    blocks: list[dict[str, Any]],
    contract_text: str,
    compact_stats: dict[str, Any],
    check_result: dict[str, Any],
    contract_summary: dict[str, Any],
    candidate_findings: list[dict[str, Any]],
    routing: dict[str, Any],
) -> dict[str, str]:
    """生成一站式审查上下文包 work/review_context.md 与 work/prefetched_laws.json。"""
    prefetched_payload, law_card_lines = _prefetch_candidate_laws(routing.get("preset_law_spec", ""))
    prefetched_path = work / "prefetched_laws.json"
    # 即使没有命中条文也写入明确的空/错误载荷，避免下一次运行误把
    # “文件不存在”当成可复用缓存。
    _write_json(prefetched_path, prefetched_payload or {"status": "empty", "results": []})

    fold_manifest_path = work / "fold_manifest.json"
    fold_manifest = compact_stats.get("fold_manifest") or []
    _write_json(fold_manifest_path, {
        "schema_version": 1,
        "source": str(source),
        "folded_tables": fold_manifest,
        "note": "任何折叠表格均可按 expand_command 从 blocks.json 无损展开。",
    })

    chk_ids = [c["id"] for c in routing.get("checklists", [])]
    digest_lines = _extract_checklist_digest(
        chk_ids,
        routing.get("group_sections"),
        has_non_compete=bool(routing.get("has_non_compete")),
    )

    parties = contract_summary.get("parties") or {}
    amounts = contract_summary.get("amounts") or {}
    has_rev = contract_summary.get("document", {}).get("has_existing_revisions", False)

    md: list[str] = [
        f"# 合同审查上下文包（`{source.name}`）",
        "",
        "> **快速审查指引**：本文件由 `preflight_contract.py` 自动汇总。模型阅读本文件、[review-quickref.md] 与 `work/contract_compact.txt` 即可直接开展分析并编写 `work/review-draft.json`。",
        "",
        "## 一、 合同基础识别与自动路由结论",
        "",
        f"- **甲方**：`{_safe_md_inline(parties.get('甲方', '未识别/空白待填'))}` ｜ **乙方**：`{_safe_md_inline(parties.get('乙方', '未识别/空白待填'))}`",
        f"- **原文既有修订痕迹**：`{'是（涉及既有修订的段落建议优先用 comment_only 或在报告中提示人工确认）' if has_rev else '否'}`",
        f"- **金额核对**：正文标示总价 `{amounts.get('explicit_total')}` ｜ 付款计划合计 `{amounts.get('payment_sum')}` ｜ 分项合计 `{amounts.get('final_component_sum')}` ｜ 识别到的金额数字 `{', '.join(amounts.get('variants', [])[:12])}`",
        f"- **文本压缩统计**：原文 `{len(contract_text)}` 字 ➔ `work/contract_compact.txt` `{compact_stats.get('compact_chars')}` 字（已折叠 `{compact_stats.get('folded_tables')}` 张功能描述大表共 `{compact_stats.get('folded_cells')}` 格；折叠明细见 `work/fold_manifest.json`，可无损展开）",
        f"- **推荐合同类型路由**：`{routing.get('primary_type')}`（{routing.get('primary_type_name')}）",
        f"- **路由置信度**：`{routing.get('confidence_level')}`（{routing.get('confidence_score')}；{routing.get('confidence_reason')}）",
        f"- **路由确认门禁**：`{'需要在 review-draft.json.routing_confirmation 中确认' if routing.get('routing_confirmation_required') else '自动路由可继续，但仍须逐项核查清单'}`",
        "- **本次必查清单覆盖（可直接用于 `checklist_coverage`）**：",
    ]
    for c in routing.get("checklists", []):
        md.append(f"  - `{c['id']}`：{c['title']}（{c['reason']}）")

    md.extend([
        "",
        "## 二、 确定性预检命中与高优风险线索（`candidate_findings`）",
        "",
    ])
    if candidate_findings:
        for cf in candidate_findings:
            md.append(f"### [{cf['severity'].upper()}] `{cf['code']}` {cf['title']}（命中 {cf['hit_count']} 处）")
            md.append(f"- **处理建议**：{cf['advice']}")
            for h in (cf.get("hits") or [])[:5]:
                md.append(f"  - `{_safe_md_inline(h.get('block_id'), 80)}` ({_safe_md_inline(h.get('clause_ref') or '—', 80)})：`{_safe_md_inline(h.get('text'), 120)}`")
            md.append("")
    else:
        md.append("- 未命中固定规则告警，请结合合同正文逐条排查。")
        md.append("")

    if digest_lines:
        md.extend([
            "## 三、 命中专类与专项清单核心要点速览",
            "",
            *digest_lines,
        ])

    if law_card_lines:
        md.extend([
            "## 四、 已预取核验的常用法条卡片（原文节选，已由 `cite_law.py` 定位）",
            "",
            "> **反锚定与精准引用铁律**：下表仅是定位节选，严禁将截断内容当作完整法条。凡审查点涉及下表未覆盖的具体法律制度或特定行业条款，**必须**调用 `cite_law.py --batch \"法名:条号1,条号2\"`（不确定条号时先调 `kb_search.py`）补充核取完整原文。",
            "",
            *law_card_lines,
            "",
        ])

    context_path = work / "review_context.md"
    context_path.write_text("\n".join(md) + "\n", encoding="utf-8")
    return {
        "review_context_md": str(context_path),
        "prefetched_laws_json": str(prefetched_path) if prefetched_payload else "",
        "fold_manifest_json": str(fold_manifest_path),
    }


_FIELD_LABEL_RE = re.compile(
    r"(?:甲\s*方|乙\s*方|地\s*址|联系人|电\s*话|传\s*真|开户行|帐号|账号|银行帐号|公司税号|开票地址|生效日期|签署日期)"
)
_PHONE_RE = re.compile(r"(?<!\d)(?:1[3-9]\d{9}|0\d{2,3}[-－]?\d{7,8})(?!\d)")


def _collect_blank_field_hits(blocks: list[dict[str, Any]], limit: int = 50) -> list[dict[str, Any]]:
    """只检查已知签署/付款字段，避免把普通标题冒号误报为空白字段。"""
    hits: list[dict[str, Any]] = []
    table_rows = _table_rows(blocks)
    table_pattern = re.compile(r"^t(?P<table>\d+)-r(?P<row>\d+)-c(?P<col>\d+)")
    for block in blocks:
        text = str(block.get("text") or "")
        table_match = table_pattern.match(str(block.get("block_id") or ""))
        row_siblings: list[dict[str, Any]] = []
        if table_match:
            row_siblings = table_rows.get(
                (table_match.group("table"), int(table_match.group("row"))), []
            )
        matches = list(_FIELD_LABEL_RE.finditer(text))
        for idx, label_match in enumerate(matches):
            colon = re.match(r"\s*[：:]", text[label_match.end() :])
            if not colon:
                continue
            value_start = label_match.end() + colon.end()
            next_label = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
            value = text[value_start:next_label].strip()
            if value and not re.fullmatch(r"(?:_|…|待填|待补|待确认|空白|\[\s*[●*]+\s*\])+(?:\s*)", value, re.I):
                continue
            # Word 表格经常把“联系人：/电话：”与实际值拆到相邻单元格。
            # 标签单元格为空时，先检查同一行是否存在非空值单元格，避免把
            # 已填写的主体、联系人和电话重复报为空白字段。
            if row_siblings:
                current_col = int(table_match.group("col")) if table_match else -1
                sibling_values = [
                    str(sibling.get("text") or "").strip()
                    for sibling in row_siblings
                    if int(sibling.get("_col", -1)) != current_col
                    and str(sibling.get("text") or "").strip()
                ]
                if sibling_values:
                    continue
            hits.append(
                {
                    "block_id": block.get("block_id"),
                    "clause_ref": block.get("clause_ref") or block.get("list_label") or "",
                    "text": text[:180],
                    "match": f"{label_match.group(0)}：",
                    "field": label_match.group(0),
                }
            )
            if len(hits) >= limit:
                return hits
    return hits


def _table_role_columns(blocks: list[dict[str, Any]]) -> dict[str, dict[str, list[int]]]:
    """推断表格中甲方/乙方标签所在列，供电话口径比对使用。"""
    table_roles: dict[str, dict[str, list[int]]] = {}
    pattern = re.compile(r"^t(?P<table>\d+)-r\d+-c(?P<col>\d+)")
    for block in blocks:
        block_id = str(block.get("block_id") or "")
        match = pattern.match(block_id)
        if not match:
            continue
        table = match.group("table")
        col = int(match.group("col"))
        text = str(block.get("text") or "")
        roles = table_roles.setdefault(table, {"甲方": [], "乙方": []})
        if re.search(r"甲\s*方", text):
            roles["甲方"].append(col)
        if re.search(r"乙\s*方", text):
            roles["乙方"].append(col)
    return table_roles


def _collect_phone_mentions(blocks: list[dict[str, Any]]) -> tuple[list[str], dict[str, list[dict[str, Any]]], list[dict[str, Any]]]:
    """提取手机和座机，按表格列/文本上下文推断主体并比较同主体多口径。"""
    table_roles = _table_role_columns(blocks)
    values: set[str] = set()
    grouped: dict[str, list[dict[str, Any]]] = {}
    mentions: list[dict[str, Any]] = []
    table_pattern = re.compile(r"^t(?P<table>\d+)-r\d+-c(?P<col>\d+)")
    for block in blocks:
        text = str(block.get("text") or "")
        block_id = str(block.get("block_id") or "")
        matches = list(_PHONE_RE.finditer(text))
        if not matches:
            continue
        prefix = text[: matches[0].start()]
        role = "未分组"
        if re.search(r"甲\s*方", prefix):
            role = "甲方"
        elif re.search(r"乙\s*方", prefix):
            role = "乙方"
        table_match = table_pattern.match(block_id)
        if table_match:
            table = table_match.group("table")
            col = int(table_match.group("col"))
            roles = table_roles.get(table) or {}
            distances = [
                (abs(col - role_col), role_name)
                for role_name, role_cols in roles.items()
                for role_col in role_cols
            ]
            if distances and role == "未分组":
                role = min(distances, key=lambda x: x[0])[1]
        for phone_match in matches:
            normalized = re.sub(r"\D", "", phone_match.group(0))
            values.add(normalized)
            mention = {
                "block_id": block_id,
                "role": role,
                "text": text[:180],
                "match": phone_match.group(0),
                "normalized": normalized,
            }
            mentions.append(mention)
            grouped.setdefault(role, []).append(mention)
    return sorted(values), grouped, mentions


def _extract_amount(text: str) -> float | None:
    cleaned = text.replace(",", "").replace("，", "")
    match = re.search(r"(?<!\d)(\d+(?:\.\d+)?)", cleaned)
    return float(match.group(1)) if match else None


def _table_rows(blocks: list[dict[str, Any]]) -> dict[tuple[str, int], list[dict[str, Any]]]:
    """按 ``t<表>-r<行>`` 聚合表格单元格，支持标签和值分列的合同表格。"""
    rows: dict[tuple[str, int], list[dict[str, Any]]] = {}
    pattern = re.compile(r"^t(?P<table>\d+)-r(?P<row>\d+)-c(?P<col>\d+)")
    for block in blocks:
        match = pattern.match(str(block.get("block_id") or ""))
        if not match:
            continue
        key = (match.group("table"), int(match.group("row")))
        item = dict(block)
        item["_col"] = int(match.group("col"))
        rows.setdefault(key, []).append(item)
    for cells in rows.values():
        cells.sort(key=lambda item: int(item.get("_col", 0)))
    return rows


def _amount_tokens(text: str) -> list[float]:
    """提取金额表格中的数字，保留小数并忽略千分位逗号。"""
    values: list[float] = []
    for raw in re.findall(r"(?<![A-Za-z0-9])\d{1,3}(?:[,，]\d{3})+(?:\.\d+)?|(?<![A-Za-z0-9])\d+(?:\.\d+)?", text):
        try:
            values.append(float(raw.replace(",", "").replace("，", "")))
        except ValueError:
            continue
    return values


def _amount_reconciliation(blocks: list[dict[str, Any]], contract_text: str) -> dict[str, Any]:
    """对常见合同总价、付款计划和最终分项价做确定性算术核对。"""
    total_match = re.search(r"本合同标的总费用\s*[：:]\s*([\d,，]+(?:\.\d+)?)\s*元?", contract_text)
    explicit_total = _extract_amount(total_match.group(1)) if total_match else None

    payment_amounts: list[float] = []
    component_values: dict[str, list[float]] = {}
    component_labels = (
        "平台费用合计",
        "平台费用最终优惠价格",
        "用户许可费用价格合计",
        "用户使用许可费用最终优惠价格",
        "服务费用合计",
        "实施费用总计",
    )
    rows = _table_rows(blocks)
    for (table, row), cells in rows.items():
        row_text = " ".join(str(cell.get("text") or "") for cell in cells)
        row_numbers = _amount_tokens(row_text)
        if not row_numbers:
            continue
        for label in component_labels:
            if label in row_text:
                # 标签和值常位于相邻单元格；取该行最后一个数字，避免把项目序号
                # 或标签中的版本编号误当成金额。
                component_values.setdefault(label, []).append(row_numbers[-1])

    for block in blocks:
        block_id = str(block.get("block_id") or "")
        text = str(block.get("text") or "")
        if re.match(r"^t9-r[2-5]-c2(?:-p\d+)?$", block_id):
            value = _extract_amount(text)
            if value is not None:
                payment_amounts.append(value)
        # 兼容未带可解析表格编号的段落式金额表达。
        value = _extract_amount(text)
        if value is None:
            continue
        for label in component_labels:
            if label in text and not any(label in " ".join(str(c.get("text") or "") for c in cells) for cells in rows.values()):
                component_values.setdefault(label, []).append(value)

    payment_sum = sum(payment_amounts) if payment_amounts else None
    final_component_labels = ("平台费用最终优惠价格", "用户使用许可费用最终优惠价格", "实施费用总计")
    final_component_values = [
        values[-1]
        for label, values in component_values.items()
        if label in final_component_labels and values
    ]
    final_component_sum = sum(final_component_values) if final_component_values else None
    comparisons: list[dict[str, Any]] = []
    for name, value in (("付款计划合计", payment_sum), ("最终分项合计", final_component_sum)):
        if explicit_total is not None and value is not None:
            comparisons.append({"name": name, "expected": explicit_total, "actual": value, "matches": value == explicit_total})
    mismatched_component_labels = [
        label for label, values in component_values.items() if len(set(values)) > 1
    ]
    comparable = bool(comparisons)
    applicable = explicit_total is not None or payment_sum is not None or final_component_sum is not None
    mismatches = [item for item in comparisons if not item["matches"]]
    status = "pass"
    if not applicable or not comparable:
        status = "not_applicable"
    elif mismatches or mismatched_component_labels:
        status = "review"
    return {
        "code": "AMOUNT_RECONCILIATION",
        "title": "金额算术核对",
        "severity": "warning" if status == "review" else "pass",
        "status": status,
        "explicit_total": explicit_total,
        "payment_sum": payment_sum,
        "final_component_sum": final_component_sum,
        "comparable": comparable,
        "comparisons": comparisons,
        "mismatched_component_labels": mismatched_component_labels,
        "advice": "签署前以盖章价格确认表锁定唯一总价、分项价和付款计划，并核对各项合计。",
    }


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _runtime_fingerprint() -> str:
    """返回影响预检语义的本地输入指纹。

    指纹只覆盖工作流代码、规则和 schema，不扫描整个法规库，避免每次预检
    都重新读取数百个文件。知识库本身通过 ``lib.kb`` 的版本/索引指纹参与
    法条预取结果；若无法取得则明确记录 ``missing``，让缓存无法静默复用。
    """
    h = hashlib.sha256()
    root = SCRIPT_DIR.parent
    for rel in _CACHE_INPUTS:
        path = root / rel
        h.update(rel.encode("utf-8"))
        h.update(b"\0")
        if path.is_file():
            h.update(sha256_file(path).encode("ascii"))
        else:
            h.update(b"missing")
        h.update(b"\0")
    try:
        from lib.kb import kb_fingerprint, kb_source_consistency
        h.update(str(kb_fingerprint()).encode("utf-8"))
        h.update(json.dumps(kb_source_consistency(), ensure_ascii=False, sort_keys=True).encode("utf-8"))
        source_manifest = root / "assets" / "kb_source_manifest.json"
        h.update(sha256_file(source_manifest).encode("ascii") if source_manifest.is_file() else b"source-manifest-missing")
    except Exception as exc:
        h.update(f"kb-error:{type(exc).__name__}".encode("utf-8"))
    return h.hexdigest()


def _write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    # 替换写避免进程中断留下半截 manifest，半截 manifest 会让后续运行误判
    # 为可复用缓存并掩盖真正的失败原因。
    tmp = path.with_name(f".{path.name}.{time.time_ns()}.tmp")
    try:
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        tmp.replace(path)
    finally:
        if tmp.exists():
            tmp.unlink()


def _line_hit(block: dict[str, Any], pattern: re.Pattern[str], max_chars: int = 180) -> dict[str, Any] | None:
    text = str(block.get("text") or "")
    match = pattern.search(text)
    if not match:
        return None
    return {
        "block_id": block.get("block_id"),
        "clause_ref": block.get("clause_ref") or block.get("list_label") or "",
        "text": text[:max_chars],
        "match": match.group(0)[:80],
    }


def _collect_hits(blocks: list[dict[str, Any]], pattern: re.Pattern[str], limit: int = 30) -> list[dict[str, Any]]:
    hits: list[dict[str, Any]] = []
    for block in blocks:
        found = _line_hit(block, pattern)
        if found:
            hits.append(found)
            if len(hits) >= limit:
                break
    return hits


def _run_checks(blocks: list[dict[str, Any]], contract_text: str) -> dict[str, Any]:
    """执行低成本、确定性的预检，不代替法律判断。"""
    current_year = date.today().year
    checks: list[dict[str, Any]] = []

    def add(code: str, title: str, severity: str, pattern: re.Pattern[str], advice: str) -> None:
        hits = _collect_hits(blocks, pattern)
        checks.append({
            "code": code,
            "title": title,
            "severity": severity if hits else "pass",
            "hit_count": len(hits),
            "hits": hits,
            "advice": advice,
        })

    blank_hits = _collect_blank_field_hits(blocks)
    checks.append(
        {
            "code": "BLANK_FIELD",
            "title": "空白/待补字段",
            "severity": "warning" if blank_hits else "pass",
            "hit_count": len(blank_hits),
            "hits": blank_hits,
            "advice": "签署前补齐主体、账户、发票、联系人、送达地址和生效日期等字段。",
        }
    )
    add(
        "STOP_SERVICE",
        "停服/停用许可权",
        "high",
        re.compile(r"(?:停止|暂停|终止).{0,20}(?:服务|许可|使用)|停服|锁定.{0,10}数据", re.I),
        "区分无争议欠款、催告和补救期；不得停用核心系统、锁定或删除数据。",
    )
    add(
        "LIABILITY_CAP",
        "责任上限/直接损失限制",
        "high",
        re.compile(r"(?:责任|赔偿).{0,20}(?:上限|不高于|不超过)|仅限.{0,10}直接损失|间接性损失", re.I),
        "核对知识产权、数据、保密、欺诈、故意或重大过失、监管罚款和第三方索赔是否排除在上限外。",
    )
    add(
        "OLD_CONTRACT_LAW",
        "旧《合同法》表述",
        "warning",
        re.compile(r"(?<!民法典)合同法(?:的规定|规定|第|相关|所称)?"),
        "一般商事合同应改为现行《民法典》及具体适用条款，并核对引用法源。",
    )
    add(
        "STALE_DATE",
        "疑似过时日期",
        "warning",
        re.compile(r"(?:19|20)\d{2}[年./-]\d{1,2}[月./-]\d{1,2}日?|(?:19|20)\d{2}年"),
        "核对日期是否属于当前项目；附件中的旧实施日期、版本和里程碑应更新。",
    )

    # 既抓“15000元”这类带单位金额，也抓表格中位于“费用/合计/单价”等
    # 金额语境后的裸数字；不把手机号、年份和人天数字混入金额清单。
    amount_values: set[str] = set()
    for raw in re.findall(r"(?<!\d)\d+(?:,\d{3})*(?:\.\d+)?\s*元", contract_text):
        amount_values.add(raw.replace(" ", "").replace(",", "").replace("元", ""))
    amount_context = re.compile(
        r"(?:总价|总费用|费用|价格|小计|合计|单价|价款|金额|收费).{0,24}?"
        r"(?P<num>\d{2,}(?:,\d{3})*(?:\.\d+)?)",
        re.I,
    )
    for match in amount_context.finditer(contract_text):
        num = match.group("num").replace(",", "")
        try:
            if float(num) >= 100:
                amount_values.add(num)
        except ValueError:
            continue
    amounts = sorted(amount_values, key=lambda x: float(x))
    amount_check = {
        "code": "AMOUNT_VARIANTS",
        "title": "金额口径",
        "severity": "warning" if len(amounts) > 1 else "pass",
        "hit_count": len(amounts),
        "values": amounts[:30],
        "advice": "核对总价、分项价、折扣价、税额和付款节点，建议以盖章价格确认表作为唯一口径。",
    }
    checks.append(amount_check)
    checks.append(_amount_reconciliation(blocks, contract_text))

    phones, phone_groups, phone_mentions = _collect_phone_mentions(blocks)
    phone_mismatches = {
        role: sorted({item["normalized"] for item in mentions})
        for role, mentions in phone_groups.items()
        if role != "未分组" and len({item["normalized"] for item in mentions}) > 1
    }
    checks.append({
        "code": "PHONE_VARIANTS",
        "title": "联系电话口径",
        "severity": "warning" if phone_mismatches else "pass",
        "hit_count": len(phone_mismatches),
        "values": phones[:30],
        "mismatches": phone_mismatches,
        "mentions": phone_mentions[:50],
        "advice": "核对首部、签署页、通知条款和附件中的联系人及电话。",
    })

    article_labels: list[str] = []
    for block in blocks:
        text = str(block.get("text") or "")
        # 目录行通常带页码或制表符，不应与正文条款标题混为重复编号。
        if "\t" in text or re.search(r"\s+\d{1,3}\s*$", text.strip()):
            continue
        article_labels.extend(re.findall(r"^(第[一二三四五六七八九十百千万〇零0-9]+条)", text.strip()))
    duplicates = sorted(k for k, v in Counter(article_labels).items() if v > 1)
    checks.append({
        "code": "DUPLICATE_CLAUSE",
        "title": "条款编号重复",
        "severity": "warning" if duplicates else "pass",
        "hit_count": len(duplicates),
        "values": duplicates,
        "advice": "统一正文及附件的条款编号，避免引用和验收记录无法定位。",
    })

    explicit_data_terms = re.compile(r"个人信息|数据处理|网络安全|数据安全|DPA|委托处理", re.I)
    context_data_terms = re.compile(
        r"个人信息|数据处理|网络安全|数据安全|DPA|委托处理|员工|客户|人力资源|移动办公|合同信息|权限|登录日志",
        re.I,
    )
    data_hits = _collect_hits(blocks, explicit_data_terms)
    data_context_hits = _collect_hits(blocks, context_data_terms)
    data_severity = "high" if data_hits else ("review" if data_context_hits else "pass")
    checks.append({
        "code": "DATA_TERMS",
        "title": "数据处理条款触发",
        "severity": data_severity,
        "confirmed": bool(data_hits),
        "triggered": bool(data_context_hits),
        "hit_count": len(data_hits),
        "hits": data_hits,
        "context_hit_count": len(data_context_hits),
        "context_hits": data_context_hits,
        "advice": "如系统处理员工、客户或合同数据，应补充处理目的、范围、访问控制、事件通报、转委托、导出和删除条款；仅出现业务数据上下文时需人工确认是否触发DPA。",
    })

    return {
        "checks": checks,
        "high_findings": [c["code"] for c in checks if c.get("severity") == "high"],
        "warning_findings": [c["code"] for c in checks if c.get("severity") == "warning"],
        "review_findings": [c["code"] for c in checks if c.get("severity") == "review"],
        "counts": {
            "blocks": len(blocks),
            "non_empty_blocks": sum(1 for b in blocks if str(b.get("text") or "").strip()),
            "contract_chars": len(contract_text),
            "amount_variants": len(amounts),
            "phone_variants": len(phones),
            "phone_mismatch_groups": len(phone_mismatches),
            "duplicate_clause_labels": len(duplicates),
            "current_year": current_year,
        },
    }


def _summary_hits(blocks: list[dict[str, Any]], pattern: re.Pattern[str], limit: int = 8) -> list[dict[str, Any]]:
    """Return compact, review-ready hits for the human/LLM review stage."""
    return _collect_hits(blocks, pattern, limit=limit)


def _build_contract_summary(
    blocks: list[dict[str, Any]], contract_text: str, check_result: dict[str, Any]
) -> dict[str, Any]:
    """Build a deterministic summary so review work starts from indexed facts.

    The summary is deliberately conservative: it records text evidence and does not
    infer legal conclusions.  It is safe to use as an LLM review input or as a UI
    preview without printing the whole contract.
    """
    party_pattern = re.compile(r"^(?P<role>甲\s*方|乙\s*方)\s*[：:]\s*(?P<name>[^\n]{2,80})$")
    parties: dict[str, str] = {}
    # 首选正文/表格中“甲方：公司名”同一块的精确标签，避免把“甲方向乙方购买”
    # 之类叙述误识别为主体名称。
    for block in blocks:
        text = str(block.get("text") or "").strip()
        match = party_pattern.match(text)
        if match:
            role = re.sub(r"\s+", "", match.group("role"))
            name = match.group("name").strip(" \t:：")
            if name and role not in parties:
                parties[role] = name

    # 旧版 Word 常把标签和公司名拆成同一表格行的相邻单元格；按表格行重建。
    row_cells: dict[tuple[str, str], list[tuple[int, str]]] = {}
    cell_re = re.compile(r"^t(?P<table>\d+)-r(?P<row>\d+)-c(?P<col>\d+)")
    for block in blocks:
        match = cell_re.match(str(block.get("block_id") or ""))
        if not match:
            continue
        key = (match.group("table"), match.group("row"))
        row_cells.setdefault(key, []).append((int(match.group("col")), str(block.get("text") or "").strip()))
    for cells in row_cells.values():
        cells.sort(key=lambda item: item[0])
        for idx, (col, text) in enumerate(cells):
            role_match = re.match(r"^(甲\s*方|乙\s*方)\s*[：:]?$", text)
            if not role_match:
                continue
            role = re.sub(r"\s+", "", role_match.group(1))
            following = next((value for next_col, value in cells[idx + 1:] if next_col > col and value), "")
            if following and role not in parties:
                parties[role] = following.strip(" \t:：")
    # Word tables commonly split “甲方：” and the company name into adjacent
    # cells. Associate the label with the first non-empty sibling in its row.
    for cells in _table_rows(blocks).values():
        for cell in cells:
            cell_text = str(cell.get("text") or "").strip()
            role_match = re.search(r"(甲\s*方|乙\s*方)\s*[：:]?", cell_text)
            if not role_match:
                continue
            role = re.sub(r"\s+", "", role_match.group(1))
            inline_name = cell_text[role_match.end() :].strip(" \t:：")
            candidates = [inline_name] if inline_name else []
            current_col = int(cell.get("_col", -1))
            following = sorted(
                (
                    sibling
                    for sibling in cells
                    if sibling is not cell
                    and int(sibling.get("_col", -1)) > current_col
                    and str(sibling.get("text") or "").strip()
                ),
                key=lambda sibling: int(sibling.get("_col", -1)),
            )
            candidates.extend(str(sibling.get("text") or "").strip(" \t:：") for sibling in following)
            candidates.extend(
                str(sibling.get("text") or "").strip(" \t:：")
                for sibling in cells
                if sibling is not cell and sibling not in following and str(sibling.get("text") or "").strip()
            )
            name = next(
                (
                    value
                    for value in candidates
                    if len(value) >= 2 and not re.fullmatch(r"(?:甲|乙)\s*方\s*[：:]?", value)
                ),
                None,
            )
            if name and role not in parties:
                parties[role] = name

    amount_check = next((c for c in check_result.get("checks", []) if c.get("code") == "AMOUNT_RECONCILIATION"), {})
    variant_check = next((c for c in check_result.get("checks", []) if c.get("code") == "AMOUNT_VARIANTS"), {})
    blank_check = next((c for c in check_result.get("checks", []) if c.get("code") == "BLANK_FIELD"), {})
    phone_check = next((c for c in check_result.get("checks", []) if c.get("code") == "PHONE_VARIANTS"), {})

    evidence_patterns = {
        "payment": re.compile(r"付款|支付|账期|发票|开票", re.I),
        "acceptance": re.compile(r"验收|上线|试运行|测试|交付", re.I),
        "stop_service": re.compile(r"停服|停止.{0,20}(?:服务|许可)|暂停.{0,20}(?:服务|许可)|锁定.{0,10}数据", re.I),
        "liability": re.compile(r"责任上限|赔偿.{0,20}(?:直接损失|间接损失|不高于|不超过)|直接损失", re.I),
        "data": re.compile(r"个人信息|数据处理|网络安全|数据安全|委托处理|权限|登录日志|员工|客户", re.I),
        "ip": re.compile(r"知识产权|著作权|源代码|二次开发|软件许可|商业秘密|侵权", re.I),
        "attachment_dates": re.compile(r"(?:19|20)\d{2}年(?:\d{1,2}月)?(?:\d{1,2}日)?", re.I),
    }
    evidence = {key: _summary_hits(blocks, pattern) for key, pattern in evidence_patterns.items()}

    # Only retain fields useful for routing and review prioritisation.
    return {
        "parties": parties,
        "document": {
            "block_count": len(blocks),
            "contract_chars": len(contract_text),
            "has_existing_revisions": any(bool(block.get("has_existing_revisions")) for block in blocks),
        },
        "amounts": {
            "explicit_total": amount_check.get("explicit_total"),
            "payment_sum": amount_check.get("payment_sum"),
            "final_component_sum": amount_check.get("final_component_sum"),
            "comparisons": amount_check.get("comparisons", []),
            "variants": variant_check.get("values", []),
        },
        "completion": {
            "blank_fields": blank_check.get("hit_count", 0),
            "phone_mismatch_groups": phone_check.get("hit_count", 0),
            "stale_date_hits": len(evidence["attachment_dates"]),
        },
        "evidence": evidence,
    }


def _build_candidate_findings(
    blocks: list[dict[str, Any]], check_result: dict[str, Any]
) -> list[dict[str, Any]]:
    """Map deterministic preflight triggers to actionable review candidates."""
    advice_by_code = {
        "STOP_SERVICE": "区分无争议欠款、催告和补救期；限制停服，不得锁定/删除核心系统数据。",
        "LIABILITY_CAP": "核对知识产权、保密、数据安全、故意或重大过失、人身损害和第三方索赔例外。",
        "DATA_TERMS": "确认是否受托处理个人信息/重要数据，并补充目的、范围、转委托、事件通报、导出删除和日志留存。",
        "AMOUNT_VARIANTS": "以盖章价格确认表锁定总价、分项价、折扣、税率和付款计划。",
        "PHONE_VARIANTS": "统一首部、签署页、通知条款和附件的联系人及电话。",
        "BLANK_FIELD": "签署前补齐主体、账户、发票、联系人、送达地址和生效日期。",
        "STALE_DATE": "核对附件日期是否属于当前项目，并更新旧实施计划。",
        "OLD_CONTRACT_LAW": "将旧《合同法》表述改为现行《民法典》及具体适用条款。",
    }
    candidates: list[dict[str, Any]] = []
    for check in check_result.get("checks", []):
        severity = str(check.get("severity") or "pass")
        if severity == "pass":
            continue
        code = str(check.get("code") or "")
        hits = check.get("hits") or check.get("context_hits") or []
        candidates.append({
            "code": code,
            "severity": severity,
            "title": check.get("title", code),
            "hit_count": check.get("hit_count", 0),
            "hits": hits[:8],
            "advice": advice_by_code.get(code, check.get("advice", "请结合合同类型清单人工复核。")),
        })
    # Stable priority: high first, then review, then warning, preserving source order.
    rank = {"high": 0, "review": 1, "warning": 2}
    candidates.sort(key=lambda item: (rank.get(item["severity"], 9), item["code"]))
    return candidates


def _load_cached(
    work: Path,
    source_sha256: str,
    *,
    cache_key: str | None = None,
    context_pack: bool = False,
) -> dict[str, Any] | None:
    manifest_path = work / "preflight.json"
    base = work / "base.docx"
    blocks = work / "blocks.json"
    contract = work / "contract.txt"
    if not manifest_path.is_file() or not base.is_file() or not blocks.is_file() or not contract.is_file():
        return None
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception:
        return None
    if manifest.get("workflow_version") != WORKFLOW_VERSION:
        return None
    if manifest.get("source", {}).get("sha256") != source_sha256:
        return None
    if cache_key is not None and manifest.get("cache_key") != cache_key:
        return None
    artifacts = manifest.get("artifacts") or {}
    work_root = work.resolve()
    for name in ("base_docx", "blocks_json", "contract_txt", "contract_compact_txt", "review_context_md", "prefetched_laws_json", "fold_manifest_json"):
        raw_path = artifacts.get(name)
        if not raw_path:
            continue
        candidate = Path(str(raw_path)).resolve()
        try:
            candidate.relative_to(work_root)
        except ValueError:
            return None
        if not candidate.is_file():
            return None
        expected = artifacts.get(f"{name}_sha256")
        if expected and sha256_file(candidate) != expected:
            return None
    if context_pack:
        required = ("review_context_md", "prefetched_laws_json", "fold_manifest_json")
        if any(not artifacts.get(name) or not Path(str(artifacts[name])).is_file() for name in required):
            return None
    return manifest


def preflight_contract(
    input_path: str | Path,
    out_dir: str | Path,
    *,
    force: bool = False,
    existing_revisions: str = "accept-view",
    context_pack: bool = True,
    run_id: str | None = None,
) -> dict[str, Any]:
    source = Path(input_path).resolve()
    if not source.is_file():
        raise FileNotFoundError(f"输入合同不存在: {source}")
    if existing_revisions not in {"accept-view", "comment-only", "fail"}:
        raise ValueError("existing_revisions 必须为 accept-view、comment-only 或 fail")

    out = Path(out_dir).resolve()
    work = out / "work"
    work.mkdir(parents=True, exist_ok=True)
    started = time.perf_counter()
    source_hash = sha256_file(source)
    runtime_fingerprint = _runtime_fingerprint()
    try:
        from lib.kb import kb_source_consistency
        kb_consistency = kb_source_consistency()
    except Exception as exc:
        kb_consistency = {"status": "error", "error": str(exc)}
    if kb_consistency.get("status") in {"mismatch", "error"}:
        raise RuntimeError(
            "知识库源文件与构建快照不一致，已拒绝继续预检；请先重新构建 Skill: "
            + json.dumps(kb_consistency, ensure_ascii=False)
        )
    cache_key = hashlib.sha256(
        json.dumps(
            {
                "source_sha256": source_hash,
                "runtime_fingerprint": runtime_fingerprint,
                "existing_revisions": existing_revisions,
                "context_pack": context_pack,
            },
            sort_keys=True,
        ).encode("utf-8")
    ).hexdigest()

    cached = None if force else _load_cached(
        work,
        source_hash,
        cache_key=cache_key,
        context_pack=context_pack,
    )
    if cached:
        cached["cache"] = {"hit": True, "reason": "source_and_runtime_inputs_match", "cache_key": cache_key}
        _write_json(work / "preflight.json", cached)
        record_stage_mark(work, "preflight_done", note="命中预检缓存", run_id=run_id, actor="script")
        return cached

    timings: dict[str, float] = {}
    step = time.perf_counter()
    base_docx = work / "base.docx"
    conversion = convert_to_docx(source, base_docx)
    timings["convert_or_copy_ms"] = round((time.perf_counter() - step) * 1000, 1)

    step = time.perf_counter()
    extraction = extract_docx(base_docx, work)
    timings["extract_ms"] = round((time.perf_counter() - step) * 1000, 1)

    blocks_data = json.loads((work / "blocks.json").read_text(encoding="utf-8"))
    contract_text = (work / "contract.txt").read_text(encoding="utf-8")
    blocks = blocks_data.get("blocks") or []
    if extraction.get("has_existing_revisions") and existing_revisions == "fail":
        raise RuntimeError("检测到原文已有修订痕迹；existing_revisions=fail，已停止后续流程")

    step = time.perf_counter()
    check_result = _run_checks(blocks, contract_text)
    timings["deterministic_checks_ms"] = round((time.perf_counter() - step) * 1000, 1)
    step = time.perf_counter()
    contract_summary = _build_contract_summary(blocks, contract_text, check_result)
    candidate_findings = _build_candidate_findings(blocks, check_result)
    timings["summary_and_candidates_ms"] = round((time.perf_counter() - step) * 1000, 1)

    step = time.perf_counter()
    compact_text, compact_stats = build_compact_contract_text(blocks)
    compact_path = work / "contract_compact.txt"
    compact_path.write_text(compact_text, encoding="utf-8")
    routing = _infer_routing(blocks, compact_text, check_result)
    context_artifacts: dict[str, str] = {}
    if context_pack:
        context_artifacts = _build_context_pack(
            work,
            source,
            blocks,
            contract_text,
            compact_stats,
            check_result,
            contract_summary,
            candidate_findings,
            routing,
        )
    timings["context_pack_ms"] = round((time.perf_counter() - step) * 1000, 1)

    artifacts_dict: dict[str, Any] = {
        "base_docx": str(base_docx),
        "base_docx_sha256": sha256_file(base_docx),
        "blocks_json": str(work / "blocks.json"),
        "blocks_json_sha256": sha256_file(work / "blocks.json"),
        "contract_txt": str(work / "contract.txt"),
        "contract_txt_sha256": sha256_file(work / "contract.txt"),
        "contract_compact_txt": str(compact_path),
        "contract_compact_txt_sha256": sha256_file(compact_path),
    }
    if context_artifacts.get("review_context_md"):
        artifacts_dict["review_context_md"] = context_artifacts["review_context_md"]
        artifacts_dict["review_context_md_sha256"] = sha256_file(Path(context_artifacts["review_context_md"]))
    if context_artifacts.get("prefetched_laws_json"):
        artifacts_dict["prefetched_laws_json"] = context_artifacts["prefetched_laws_json"]
        artifacts_dict["prefetched_laws_json_sha256"] = sha256_file(Path(context_artifacts["prefetched_laws_json"]))
    if context_artifacts.get("fold_manifest_json"):
        artifacts_dict["fold_manifest_json"] = context_artifacts["fold_manifest_json"]
        artifacts_dict["fold_manifest_json_sha256"] = sha256_file(Path(context_artifacts["fold_manifest_json"]))

    manifest: dict[str, Any] = {
        "workflow_version": WORKFLOW_VERSION,
        "cache_key": cache_key,
        "runtime_fingerprint": runtime_fingerprint,
        "kb_source_consistency": kb_consistency,
        "source": {
            "path": str(source),
            "name": source.name,
            "suffix": source.suffix.lower(),
            "size": source.stat().st_size,
            "mtime_ns": source.stat().st_mtime_ns,
            "sha256": source_hash,
        },
        "artifacts": artifacts_dict,
        "conversion": conversion,
        "extraction": {
            "block_count": extraction.get("block_count"),
            "has_existing_revisions": extraction.get("has_existing_revisions"),
            "has_comments": extraction.get("has_comments"),
            "existing_revisions_strategy": existing_revisions,
            "compact": compact_stats,
        },
        "routing": routing,
        "checks": check_result,
        "summary": contract_summary,
        "candidate_findings": candidate_findings,
        "timings_ms": timings,
        "cache": {"hit": False, "reason": "fresh_run"},
    }
    manifest["timings_ms"]["total_ms"] = round((time.perf_counter() - started) * 1000, 1)
    _write_json(work / "preflight.json", manifest)
    record_stage_mark(work, "preflight_done", note=f"预检与上下文包生成完成（{manifest['timings_ms']['total_ms']}ms）", run_id=run_id, actor="script")
    return manifest


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(description="合同审核预检、一次性转换/提取、自动路由与审查上下文包生成")
    parser.add_argument("input", help="合同文件（.doc/.wps/.rtf/.docx）")
    parser.add_argument("--out-dir", required=True, help="本次审核输出目录，将生成 work/")
    parser.add_argument("--force", action="store_true", help="忽略缓存，重新转换并提取")
    parser.add_argument(
        "--context-pack",
        dest="context_pack",
        action="store_true",
        default=True,
        help="生成 work/review_context.md 与 work/contract_compact.txt（默认开启）",
    )
    parser.add_argument(
        "--no-context-pack",
        dest="context_pack",
        action="store_false",
        help="仅生成压缩文本，跳过 review_context.md 与预取法条",
    )
    parser.add_argument(
        "--existing-revisions",
        choices=["accept-view", "comment-only", "fail"],
        default="accept-view",
        help="原文已有修订时的处理策略；默认按当前接受视图提取",
    )
    args = parser.parse_args(argv)

    try:
        result = preflight_contract(
            args.input,
            args.out_dir,
            force=args.force,
            existing_revisions=args.existing_revisions,
            context_pack=args.context_pack,
        )
    except Exception as exc:
        print(json.dumps({"error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1

    summary = {
        "source": result["source"]["path"],
        "source_sha256": result["source"]["sha256"],
        "cache_hit": result.get("cache", {}).get("hit", False),
        "base_docx": result["artifacts"]["base_docx"],
        "contract_compact_txt": result["artifacts"].get("contract_compact_txt"),
        "review_context_md": result["artifacts"].get("review_context_md"),
        "block_count": result["extraction"].get("block_count"),
        "compact": result["extraction"].get("compact"),
        "primary_type": (result.get("routing") or {}).get("primary_type"),
        "has_existing_revisions": result["extraction"].get("has_existing_revisions"),
        "high_findings": result["checks"].get("high_findings", []),
        "warning_findings": result["checks"].get("warning_findings", []),
        "review_findings": result["checks"].get("review_findings", []),
        "timings_ms": result.get("timings_ms", {}),
        "report": str(Path(args.out_dir).resolve() / "work" / "preflight.json"),
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
