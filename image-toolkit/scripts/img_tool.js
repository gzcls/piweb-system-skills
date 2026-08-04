// img_tool.js —— 图片处理工具箱，基于 sharp（Node）。
// 子命令：resize / convert / compress / thumbnail / watermark / info
// 单文件：
//   node img_tool.js resize /workspace/in.png -o /workspace/out.png --width 800
// 批量（--outdir，对多个文件/通配符逐个处理，输出同名文件到目录）：
//   node img_tool.js resize "/workspace/imgs/*.png" --outdir /workspace/out --width 800
//   node img_tool.js convert "/workspace/imgs/*.png" --outdir /workspace/out --ext webp --quality 70
// 若找不到全局模块，脚本会自动从 `npm root -g` 解析 sharp。

const fs = require("fs");
const path = require("path");

function load(name) {
  try { return require(name); }
  catch (e) {
    try {
      const groot = require("child_process").execSync("npm root -g").toString().trim();
      return require(path.join(groot, name));
    } catch (e2) {
      console.error("无法加载模块 " + name + "：沙箱镜像应已全局安装 sharp；若仍缺失，请重建镜像，或设置 NODE_PATH=$(npm root -g)");
      throw e;
    }
  }
}
const sharp = load("sharp");

// 极简参数解析
function parseArgs(argv) {
  const out = { _: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "-o" || a === "--out") out.out = argv[++i];
    else if (a.startsWith("--")) {
      const k = a.slice(2);
      const v = (argv[i + 1] && !argv[i + 1].startsWith("--")) ? argv[++i] : "true";
      out[k] = v;
    } else out._.push(a);
  }
  return out;
}

// 展开通配符（shell 未展开时的兜底；支持 dir/*.ext 这类）
function expand(inputs) {
  const out = [];
  for (const p of inputs) {
    if (p.includes("*")) {
      const dir = path.dirname(p), base = path.basename(p);
      const re = new RegExp("^" + base.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*") + "$");
      try {
        for (const f of fs.readdirSync(dir)) if (re.test(f)) out.push(path.join(dir, f));
      } catch (e) { /* 目录不存在则忽略 */ }
    } else out.push(p);
  }
  return out;
}

function esc(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

async function processOne(cmd, src, outFile, a) {
  if (!fs.existsSync(src)) { console.error("文件不存在: " + src); return false; }

  if (cmd === "info") {
    const m = await sharp(src).metadata();
    console.log(src + ": " + JSON.stringify({ format: m.format, width: m.width, height: m.height, space: m.space, hasAlpha: m.hasAlpha }));
    return true;
  }

  let img = sharp(src);
  const w = a.width ? parseInt(a.width) : null;
  const h = a.height ? parseInt(a.height) : null;
  const q = a.quality ? parseInt(a.quality) : 80;

  if (cmd === "resize") {
    img = img.resize(w, h, { fit: a.fit || "inside", withoutEnlargement: true });
  } else if (cmd === "thumbnail") {
    img = img.resize(w || 200, h || null, { fit: "cover", position: "centre" });
  } else if (cmd === "convert" || cmd === "compress") {
    // 仅靠下方按扩展名设置格式/质量
  } else if (cmd === "watermark") {
    const text = a.text || "样本";
    const meta = await sharp(src).metadata();
    const W = meta.width || 800, H = meta.height || 600;
    const fontSize = Math.max(18, Math.round(W / 18));
    const svg = `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">
      <text x="${W - 16}" y="${H - 16}" text-anchor="end"
        font-family="Noto Sans CJK SC, Microsoft YaHei, sans-serif"
        font-size="${fontSize}" fill="rgba(0,0,0,0.35)"
        stroke="rgba(255,255,255,0.35)" stroke-width="0.6">${esc(text)}</text></svg>`;
    img = img.composite([{ input: Buffer.from(svg), gravity: "southeast" }]);
  } else {
    console.error("未知子命令: " + cmd); process.exit(2);
  }

  const ext = path.extname(outFile).toLowerCase();
  if (ext === ".jpg" || ext === ".jpeg") img = img.jpeg({ quality: q });
  else if (ext === ".webp") img = img.webp({ quality: q });
  else if (ext === ".png") img = img.png({ compressionLevel: 9 });
  else if (ext === ".avif") img = img.avif({ quality: q });

  fs.mkdirSync(path.dirname(path.resolve(outFile)), { recursive: true });
  await img.toFile(outFile);
  console.log("[OK] 已生成 " + path.resolve(outFile));
  return true;
}

async function main() {
  const [cmd, ...rest] = process.argv.slice(2);
  const a = parseArgs(rest);
  const inputs = expand(a._);

  if (!cmd) {
    console.error("用法: node img_tool.js <resize|convert|compress|thumbnail|watermark|info> <in...> (-o <out> | --outdir <dir>) [--opts]");
    process.exit(2);
  }

  // 批量模式：给了 --outdir，对所有输入逐个处理
  if (a.outdir) {
    if (!inputs.length) { console.error("批量模式下没有匹配到输入文件"); process.exit(2); }
    let ok = 0;
    for (const src of inputs) {
      let name = path.basename(src);
      if (cmd === "convert" && a.ext) {
        name = name.replace(/\.[^.]+$/, "." + String(a.ext).replace(/^\./, ""));
      }
      const outFile = path.join(a.outdir, name);
      if (await processOne(cmd, src, outFile, a)) ok++;
    }
    console.log(`[OK] 批量完成，共处理 ${ok}/${inputs.length} 个文件`);
    return;
  }

  // 单文件模式
  const src = inputs[0] || a._[0];
  if (cmd === "info") {
    if (!src) { console.error("缺少输入文件"); process.exit(2); }
    await processOne("info", src, null, a);
    return;
  }
  if (!src || !a.out) {
    console.error("单文件模式需要 <in> 与 -o <out>；批量请用 --outdir");
    process.exit(2);
  }
  await processOne(cmd, src, a.out, a);
}

main().catch((e) => { console.error(e); process.exit(1); });
