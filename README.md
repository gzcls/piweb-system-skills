# PIWeb System Skills

This repository contains 28 deployable system-level File Skill packages. It includes the original 21-package snapshot exported from the local PIWeb runtime on 2026-08-04 plus independently maintained additions.

## Skills

- `any-to-markdown`
- `contract-auditor`
- `compare-contracts`
- `data-analysis`
- `group-legal-contract`
- `data-to-docx`
- `data-to-pdf`
- `data-to-xlsx`
- `docx`
- `edge-tts`
- `frontend-design`
- `html-to-pdf`
- `huaxin-pptx`
- `image-toolkit`
- `lucanet-excel-data-extraction`
- `mx-data` — 妙想金融数据
- `mx-search` — 妙想资讯搜索
- `mx-xuangu` — 妙想智能选股
- `mx-moni` — 妙想模拟组合管理
- `mx-zixuan` — 妙想自选股管理
- `office-convert`
- `pdf`
- `pdf-ocr`
- `pdf-to-image`
- `pdf-toolkit`
- `qrcode`
- `video-toolkit`
- `xlsx`

## Production installation

Install through PIWeb's governed File Skill ingestion endpoint. Do not copy these directories directly into the production runtime projection.

```http
POST /api/skills/install
Content-Type: application/json

{
  "package": "https://github.com/gzcls/piweb-system-skills.git",
  "scope": "system"
}
```

The production installation creates new installation and revision identities. Compatibility findings must be reviewed in the File Skill approval queue before affected packages become active.

Runtime dependencies such as LibreOffice, Chromium, FFmpeg, Python packages, and global Node packages belong in the production sandbox image; they are not provisioned by this repository.

## 妙想 Skills

The five `mx-*` packages are included as top-level skill directories so PIWeb can discover them when installing this repository. Each package contains `SKILL.md`, its Python helper, and upstream release metadata in `_meta.json`. Release versions: `mx-data` 1.0.5, `mx-search` 1.0.5, `mx-xuangu` 1.0.4, `mx-moni` 1.0.4, and `mx-zixuan` 1.0.0.

Source: [东方财富妙想 Skills](https://marketing.dfcfs.com/views/finskillshub/index) and its QQK000 ClawHub releases. These are the existing PIWeb-adapted packages; upstream releases now use `eastmoney-*` names.

Configure `MX_APIKEY` through PIWeb credential configuration. Do not commit credentials to this repository. The sandbox needs outbound access to `https://mkapi2.dfcfs.com`, Python with `requests`, and, for `mx-data` Excel exports, `pandas` and `openpyxl`; `mx-moni` also uses `curl`. Dependencies belong in the sandbox image.

After an earlier installation of this repository, install the repository again to discover newly added packages. Follow the File Skill approval flow before activation.

## Licensing

Some packages contain their own license files. Those licenses apply to their respective package contents. No repository-wide license is granted for files that do not include an explicit license.
