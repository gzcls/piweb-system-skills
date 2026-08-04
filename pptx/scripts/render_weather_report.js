#!/usr/bin/env node

const fs = require("node:fs");
const path = require("node:path");
const PptxGenJS = require("pptxgenjs");

const SKILL_ROOT = path.resolve(__dirname, "..");
const TITLE_BG = path.join(SKILL_ROOT, "assets", "huaxin_bg_title.jpeg");
const CONTENT_BG = path.join(SKILL_ROOT, "assets", "huaxin_bg_content.jpeg");
const SLIDE_W = 10;
const SLIDE_H = 5.625;
const FONT = "Microsoft YaHei";
const COLORS = {
  ink: "25354D",
  muted: "667085",
  primary: "374B9B",
  red: "E63237",
  orange: "ED7D31",
  green: "70AD47",
  blue: "2E75B6",
  cyan: "3AA6B9",
  paleBlue: "EAF1FB",
  paleRed: "FDECEE",
  paleGreen: "EEF6E9",
  paleOrange: "FFF3E8",
  border: "D7DEE8",
  white: "FFFFFF",
};

function fail(message) {
  throw new Error(`weather report input: ${message}`);
}

function requireString(value, field, maxLength = 200) {
  if (typeof value !== "string" || !value.trim()) fail(`${field} must be a non-empty string`);
  const normalized = value.trim();
  if (normalized.length > maxLength) fail(`${field} is longer than ${maxLength} characters`);
  return normalized;
}

function optionalString(value, field, maxLength = 200) {
  if (value == null || value === "") return undefined;
  return requireString(value, field, maxLength);
}

function requireNumber(value, field, min, max) {
  if (typeof value !== "number" || !Number.isFinite(value) || value < min || value > max) {
    fail(`${field} must be a number between ${min} and ${max}`);
  }
  return value;
}

function requireUrl(value, field) {
  const text = requireString(value, field, 2_000);
  let parsed;
  try {
    parsed = new URL(text);
  } catch {
    fail(`${field} must be an absolute HTTP(S) URL`);
  }
  if (!new Set(["http:", "https:"]).has(parsed.protocol)) fail(`${field} must use HTTP(S)`);
  return text;
}

function requireIsoDate(value, field) {
  const text = requireString(value, field, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text) || Number.isNaN(Date.parse(`${text}T00:00:00Z`))) {
    fail(`${field} must use YYYY-MM-DD`);
  }
  return text;
}

function normalizeSource(source, field) {
  if (!source || typeof source !== "object" || Array.isArray(source)) fail(`${field} is required`);
  return {
    name: requireString(source.name, `${field}.name`, 80),
    url: requireUrl(source.url, `${field}.url`),
  };
}

function normalizeInput(raw) {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) fail("root must be an object");
  if (!Array.isArray(raw.days) || raw.days.length !== 7) fail("days must contain exactly 7 forecasts");

  const seenDates = new Set();
  const days = raw.days.map((day, index) => {
    if (!day || typeof day !== "object" || Array.isArray(day)) fail(`days[${index}] must be an object`);
    const date = requireIsoDate(day.date, `days[${index}].date`);
    if (seenDates.has(date)) fail(`duplicate forecast date: ${date}`);
    seenDates.add(date);
    const lowC = requireNumber(day.lowC, `days[${index}].lowC`, -60, 60);
    const highC = requireNumber(day.highC, `days[${index}].highC`, -60, 60);
    if (highC < lowC) fail(`days[${index}].highC must be >= lowC`);
    return {
      date,
      weekday: optionalString(day.weekday, `days[${index}].weekday`, 12) || weekdayFor(date),
      condition: requireString(day.condition, `days[${index}].condition`, 40),
      lowC,
      highC,
      wind: optionalString(day.wind, `days[${index}].wind`, 60) || "未提供",
      precipProbability: day.precipProbability == null
        ? undefined
        : requireNumber(day.precipProbability, `days[${index}].precipProbability`, 0, 100),
      humidity: day.humidity == null
        ? undefined
        : requireNumber(day.humidity, `days[${index}].humidity`, 0, 100),
      uv: optionalString(day.uv, `days[${index}].uv`, 30),
    };
  }).sort((a, b) => a.date.localeCompare(b.date));

  const alerts = raw.alerts == null ? [] : raw.alerts;
  if (!Array.isArray(alerts) || alerts.length > 5) fail("alerts must be an array with at most 5 items");

  const retrievedAt = requireString(raw.retrievedAt, "retrievedAt", 80);
  if (Number.isNaN(Date.parse(retrievedAt))) fail("retrievedAt must be an ISO-8601 timestamp");

  return {
    city: requireString(raw.city, "city", 40),
    reportTitle: optionalString(raw.reportTitle, "reportTitle", 80) || `${raw.city}未来一周天气报告`,
    reporter: optionalString(raw.reporter, "reporter", 40) || "华新智能小助手",
    retrievedAt,
    source: normalizeSource(raw.source, "source"),
    secondarySource: raw.secondarySource ? normalizeSource(raw.secondarySource, "secondarySource") : undefined,
    days,
    alerts: alerts.map((alert, index) => {
      if (!alert || typeof alert !== "object" || Array.isArray(alert)) fail(`alerts[${index}] must be an object`);
      return {
        title: requireString(alert.title, `alerts[${index}].title`, 80),
        description: requireString(alert.description, `alerts[${index}].description`, 240),
        sourceUrl: requireUrl(alert.sourceUrl, `alerts[${index}].sourceUrl`),
      };
    }),
  };
}

function weekdayFor(date) {
  return new Intl.DateTimeFormat("zh-CN", { weekday: "short", timeZone: "Asia/Shanghai" })
    .format(new Date(`${date}T12:00:00+08:00`));
}

function shortDate(date) {
  const [, month, day] = date.split("-");
  return `${Number(month)}/${Number(day)}`;
}

function chineseDate(date) {
  const [year, month, day] = date.split("-");
  return `${year}年${Number(month)}月${Number(day)}日`;
}

function periodText(data) {
  return `${chineseDate(data.days[0].date)} - ${chineseDate(data.days[data.days.length - 1].date)}`;
}

function conditionStyle(condition) {
  if (/雷|暴|storm/i.test(condition)) return { color: COLORS.red, fill: COLORS.paleRed, label: "强对流" };
  if (/雨|雪|rain|shower/i.test(condition)) return { color: COLORS.cyan, fill: "E9F7F9", label: "降水" };
  if (/晴|sun|clear/i.test(condition)) return { color: COLORS.orange, fill: COLORS.paleOrange, label: "晴热" };
  return { color: COLORS.primary, fill: COLORS.paleBlue, label: "多云" };
}

function deriveMetrics(data) {
  const rainDays = data.days.filter((day) => /雨|雪|雷|rain|shower|storm/i.test(day.condition)).length;
  const hottest = data.days.reduce((best, day) => day.highC > best.highC ? day : best);
  const coolest = data.days.reduce((best, day) => day.lowC < best.lowC ? day : best);
  const averageHigh = Math.round(data.days.reduce((sum, day) => sum + day.highC, 0) / data.days.length);
  const averageLow = Math.round(data.days.reduce((sum, day) => sum + day.lowC, 0) / data.days.length);
  const largestSwing = data.days.reduce((best, day) => {
    const swing = day.highC - day.lowC;
    return swing > best.swing ? { day, swing } : best;
  }, { day: data.days[0], swing: data.days[0].highC - data.days[0].lowC });
  return { rainDays, hottest, coolest, averageHigh, averageLow, largestSwing };
}

function addBackground(slide, image) {
  slide.addImage({ path: image, x: 0, y: 0, w: SLIDE_W, h: SLIDE_H });
}

function addContentTitle(slide, title) {
  slide.addText(title, {
    x: 0.5, y: 0.1, w: 8.9, h: 0.42,
    fontFace: FONT, fontSize: 20, bold: true, color: COLORS.white,
    margin: 0, breakLine: false, valign: "mid",
  });
}

function addFooter(slide, data, page, total) {
  const sourceNames = [data.source.name, data.secondarySource?.name].filter(Boolean).join(" / ");
  slide.addText(`来源：${sourceNames}  |  获取时间：${data.retrievedAt}`, {
    x: 0.5, y: 5.36, w: 8.25, h: 0.14,
    fontFace: FONT, fontSize: 7.5, color: COLORS.muted, margin: 0,
    fit: "shrink",
  });
  slide.addText(`${page}/${total}`, {
    x: 9.0, y: 5.34, w: 0.5, h: 0.16,
    fontFace: FONT, fontSize: 8, color: COLORS.muted, align: "right", margin: 0,
  });
}

function addCard(slide, x, y, w, h, options = {}) {
  slide.addShape("roundRect", {
    x, y, w, h,
    rectRadius: 0.06,
    fill: { color: options.fill || COLORS.white, transparency: options.transparency || 0 },
    line: { color: options.line || COLORS.border, width: options.lineWidth || 1 },
    shadow: options.shadow === false ? undefined : { type: "outer", color: "B8C1CE", opacity: 0.16, blur: 1, angle: 45, distance: 1 },
  });
}

function addMetric(slide, x, label, value, note, color, fill) {
  addCard(slide, x, 1.42, 2.0, 1.35, { fill, line: color, shadow: false });
  slide.addText(value, {
    x: x + 0.15, y: 1.63, w: 1.7, h: 0.48,
    fontFace: FONT, fontSize: 25, bold: true, color, align: "center", margin: 0,
  });
  slide.addText(label, {
    x: x + 0.15, y: 2.15, w: 1.7, h: 0.22,
    fontFace: FONT, fontSize: 11, bold: true, color: COLORS.ink, align: "center", margin: 0,
  });
  slide.addText(note, {
    x: x + 0.15, y: 2.42, w: 1.7, h: 0.16,
    fontFace: FONT, fontSize: 8, color: COLORS.muted, align: "center", margin: 0,
  });
}

function addDailyCard(slide, day, x, y, w, h) {
  const style = conditionStyle(day.condition);
  addCard(slide, x, y, w, h, { fill: COLORS.white, line: style.color });
  slide.addShape("rect", { x, y, w, h: 0.12, fill: { color: style.color }, line: { color: style.color } });
  slide.addText(`${day.weekday}  ${shortDate(day.date)}`, {
    x: x + 0.14, y: y + 0.25, w: w - 0.28, h: 0.28,
    fontFace: FONT, fontSize: 13, bold: true, color: COLORS.ink, margin: 0,
  });
  slide.addText(day.condition, {
    x: x + 0.14, y: y + 0.66, w: w - 0.28, h: 0.4,
    fontFace: FONT, fontSize: 19, bold: true, color: style.color, margin: 0,
    fit: "shrink",
  });
  slide.addText(`${day.highC}°`, {
    x: x + 0.14, y: y + 1.2, w: 0.85, h: 0.48,
    fontFace: FONT, fontSize: 26, bold: true, color: COLORS.red, margin: 0,
  });
  slide.addText(`/ ${day.lowC}°C`, {
    x: x + 0.95, y: y + 1.36, w: w - 1.1, h: 0.25,
    fontFace: FONT, fontSize: 12, color: COLORS.blue, margin: 0,
  });
  const details = [
    `风力：${day.wind}`,
    day.precipProbability == null ? "降水概率：未提供" : `降水概率：${day.precipProbability}%`,
    day.humidity == null ? undefined : `湿度：${day.humidity}%`,
    day.uv ? `紫外线：${day.uv}` : undefined,
  ].filter(Boolean);
  slide.addText(details.join("\n"), {
    x: x + 0.14, y: y + 1.84, w: w - 0.28, h: h - 2.02,
    fontFace: FONT, fontSize: 9.5, color: COLORS.muted, margin: 0,
    breakLine: false, valign: "top", fit: "shrink",
  });
}

function buildRecommendations(data, metrics) {
  const items = [];
  if (metrics.rainDays > 0) {
    items.push({ title: "通勤防雨", body: `本周有 ${metrics.rainDays} 天出现降水类天气，随身携带雨具并预留通勤时间。`, color: COLORS.cyan, fill: "E9F7F9" });
  } else {
    items.push({ title: "出行安排", body: "本周预报未出现降水类天气，仍建议出发前查看临近预报。", color: COLORS.green, fill: COLORS.paleGreen });
  }
  if (metrics.hottest.highC >= 35) {
    items.push({ title: "防暑补水", body: `${shortDate(metrics.hottest.date)} 最高 ${metrics.hottest.highC}°C，减少午后长时间户外活动。`, color: COLORS.orange, fill: COLORS.paleOrange });
  } else {
    items.push({ title: "温度管理", body: `最高温约 ${metrics.hottest.highC}°C，室内外温差场景注意及时增减衣物。`, color: COLORS.orange, fill: COLORS.paleOrange });
  }
  const strongUvDays = data.days.filter((day) => day.uv && /强|高|very high|extreme/i.test(day.uv)).length;
  items.push({
    title: strongUvDays > 0 ? "紫外线防护" : "动态复核",
    body: strongUvDays > 0
      ? `${strongUvDays} 天紫外线描述偏强，户外活动使用遮阳与防晒措施。`
      : "天气和预警会动态更新，重要行程请在出发前复核官方临近预报。",
    color: COLORS.primary,
    fill: COLORS.paleBlue,
  });
  return items.slice(0, 3);
}

function createDeck(data) {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "HUAXIN_16X9", width: SLIDE_W, height: SLIDE_H });
  pptx.layout = "HUAXIN_16X9";
  pptx.author = data.reporter;
  pptx.company = "华新集团";
  pptx.subject = `${data.city}七天天气预报`;
  pptx.title = data.reportTitle;
  pptx.lang = "zh-CN";
  pptx.theme = {
    headFontFace: FONT,
    bodyFontFace: FONT,
    lang: "zh-CN",
  };
  pptx.defineSlideMaster({
    title: "HUAXIN_CONTENT",
    background: { color: COLORS.white },
    objects: [],
    slideNumber: { x: 9, y: 5.34, color: COLORS.muted, fontFace: FONT, fontSize: 8 },
  });

  const total = 7;
  const metrics = deriveMetrics(data);

  // 1. Cover
  let slide = pptx.addSlide();
  addBackground(slide, TITLE_BG);
  slide.addText(data.reportTitle, {
    x: 2.62, y: 0.9, w: 5.9, h: 0.95,
    fontFace: FONT, fontSize: 34, bold: true, color: COLORS.white,
    margin: 0, valign: "mid", fit: "shrink",
  });
  slide.addText(periodText(data), {
    x: 2.62, y: 2.0, w: 5.4, h: 0.38,
    fontFace: FONT, fontSize: 15, color: "DCE4F2", margin: 0,
  });
  slide.addShape("roundRect", {
    x: 2.62, y: 2.75, w: 2.2, h: 0.45,
    fill: { color: COLORS.white, transparency: 82 },
    line: { color: COLORS.white, transparency: 65 },
  });
  slide.addText("7 DAY WEATHER OUTLOOK", {
    x: 2.82, y: 2.88, w: 1.8, h: 0.14,
    fontFace: "Arial", fontSize: 8.5, bold: true, color: COLORS.white,
    align: "center", margin: 0, charSpacing: 1.2,
  });
  slide.addText(`汇报人：${data.reporter}   |   汇报日期：${chineseDate(data.retrievedAt.slice(0, 10))}`, {
    x: 4.6, y: 5.12, w: 4.9, h: 0.2,
    fontFace: FONT, fontSize: 9, color: "44546A", align: "right", margin: 0,
  });

  // 2. Overview
  slide = pptx.addSlide();
  addBackground(slide, CONTENT_BG);
  addContentTitle(slide, "一周天气概览");
  addMetric(slide, 0.65, "平均最高温", `${metrics.averageHigh}°C`, "七日算术平均", COLORS.red, COLORS.paleRed);
  addMetric(slide, 2.9, "平均最低温", `${metrics.averageLow}°C`, "七日算术平均", COLORS.blue, COLORS.paleBlue);
  addMetric(slide, 5.15, "降水天气", `${metrics.rainDays} 天`, "按天气描述统计", COLORS.cyan, "E9F7F9");
  addMetric(slide, 7.4, "本周高点", `${metrics.hottest.highC}°C`, `${shortDate(metrics.hottest.date)} ${metrics.hottest.weekday}`, COLORS.orange, COLORS.paleOrange);
  addCard(slide, 0.65, 3.15, 8.75, 1.25, { fill: "F7F9FC", line: COLORS.border, shadow: false });
  slide.addText("关键观察", {
    x: 0.9, y: 3.46, w: 1.0, h: 0.28,
    fontFace: FONT, fontSize: 13, bold: true, color: COLORS.primary, margin: 0,
  });
  slide.addText(
    `最高温出现在 ${shortDate(metrics.hottest.date)}（${metrics.hottest.highC}°C）；最低温出现在 ${shortDate(metrics.coolest.date)}（${metrics.coolest.lowC}°C）；最大日温差为 ${metrics.largestSwing.swing}°C。`,
    {
      x: 1.85, y: 3.38, w: 7.2, h: 0.48,
      fontFace: FONT, fontSize: 12, color: COLORS.ink, margin: 0, valign: "mid", fit: "shrink",
    },
  );
  addFooter(slide, data, 2, total);

  // 3. Temperature trend
  slide = pptx.addSlide();
  addBackground(slide, CONTENT_BG);
  addContentTitle(slide, "七日温度趋势");
  const chart = { x: 0.85, y: 1.35, w: 8.3, h: 3.15 };
  addCard(slide, 0.55, 1.05, 8.9, 3.75, { fill: COLORS.white, line: COLORS.border, shadow: false });
  const allTemps = data.days.flatMap((day) => [day.lowC, day.highC]);
  const minTemp = Math.floor((Math.min(...allTemps) - 2) / 5) * 5;
  const maxTemp = Math.ceil((Math.max(...allTemps) + 2) / 5) * 5;
  const range = Math.max(5, maxTemp - minTemp);
  const yFor = (temp) => chart.y + chart.h - ((temp - minTemp) / range) * chart.h;
  for (let i = 0; i <= 4; i += 1) {
    const temp = minTemp + (range * i) / 4;
    const y = yFor(temp);
    slide.addShape("line", { x: chart.x, y, w: chart.w, h: 0, line: { color: "E5EAF1", width: 0.8 } });
    slide.addText(`${Math.round(temp)}°`, {
      x: 0.5, y: y - 0.08, w: 0.3, h: 0.16,
      fontFace: FONT, fontSize: 8, color: COLORS.muted, align: "right", margin: 0,
    });
  }
  const step = chart.w / (data.days.length - 1);
  for (const [field, color] of [["highC", COLORS.red], ["lowC", COLORS.blue]]) {
    for (let i = 0; i < data.days.length - 1; i += 1) {
      const x1 = chart.x + i * step;
      const y1 = yFor(data.days[i][field]);
      const x2 = chart.x + (i + 1) * step;
      const y2 = yFor(data.days[i + 1][field]);
      slide.addShape("line", { x: x1, y: y1, w: x2 - x1, h: y2 - y1, line: { color, width: 2.4, beginArrowType: "none", endArrowType: "none" } });
    }
    data.days.forEach((day, i) => {
      const x = chart.x + i * step;
      const y = yFor(day[field]);
      slide.addShape("ellipse", { x: x - 0.07, y: y - 0.07, w: 0.14, h: 0.14, fill: { color }, line: { color: COLORS.white, width: 1 } });
      slide.addText(`${day[field]}°`, {
        x: x - 0.25, y: y + (field === "highC" ? -0.35 : 0.14), w: 0.5, h: 0.18,
        fontFace: FONT, fontSize: 8.5, bold: true, color, align: "center", margin: 0,
      });
    });
  }
  data.days.forEach((day, i) => {
    const x = chart.x + i * step;
    slide.addText(`${day.weekday}\n${shortDate(day.date)}`, {
      x: x - 0.35, y: 4.55, w: 0.7, h: 0.35,
      fontFace: FONT, fontSize: 8.5, color: COLORS.ink, align: "center", margin: 0, fit: "shrink",
    });
  });
  slide.addShape("line", { x: 7.4, y: 0.84, w: 0.38, h: 0, line: { color: COLORS.red, width: 2.5 } });
  slide.addText("最高温", { x: 7.84, y: 0.76, w: 0.6, h: 0.18, fontFace: FONT, fontSize: 8.5, color: COLORS.ink, margin: 0 });
  slide.addShape("line", { x: 8.45, y: 0.84, w: 0.38, h: 0, line: { color: COLORS.blue, width: 2.5 } });
  slide.addText("最低温", { x: 8.89, y: 0.76, w: 0.6, h: 0.18, fontFace: FONT, fontSize: 8.5, color: COLORS.ink, margin: 0 });
  addFooter(slide, data, 3, total);

  // 4. Daily details, first four days
  slide = pptx.addSlide();
  addBackground(slide, CONTENT_BG);
  addContentTitle(slide, "逐日预报 · 前四日");
  data.days.slice(0, 4).forEach((day, index) => addDailyCard(slide, day, 0.5 + index * 2.32, 1.15, 2.08, 3.95));
  addFooter(slide, data, 4, total);

  // 5. Daily details, final three days and alerts
  slide = pptx.addSlide();
  addBackground(slide, CONTENT_BG);
  addContentTitle(slide, "逐日预报 · 后三日");
  data.days.slice(4).forEach((day, index) => addDailyCard(slide, day, 0.55 + index * 3.02, 1.12, 2.72, 3.2));
  const alertText = data.alerts.length > 0
    ? data.alerts.map((alert) => `${alert.title}：${alert.description}`).join("  |  ")
    : "本报告未录入生效中的预警信息；天气和预警可能动态更新，请以官方实时发布为准。";
  addCard(slide, 0.55, 4.5, 8.78, 0.55, { fill: data.alerts.length ? COLORS.paleRed : "F7F9FC", line: data.alerts.length ? COLORS.red : COLORS.border, shadow: false });
  slide.addText(data.alerts.length ? `预警  ${alertText}` : `提示  ${alertText}`, {
    x: 0.75, y: 4.67, w: 8.38, h: 0.2,
    fontFace: FONT, fontSize: 9, bold: data.alerts.length > 0, color: data.alerts.length ? COLORS.red : COLORS.muted,
    margin: 0, fit: "shrink",
  });
  addFooter(slide, data, 5, total);

  // 6. Recommendations and sources
  slide = pptx.addSlide();
  addBackground(slide, CONTENT_BG);
  addContentTitle(slide, "出行建议与数据说明");
  buildRecommendations(data, metrics).forEach((item, index) => {
    const x = 0.55 + index * 3.02;
    addCard(slide, x, 1.15, 2.72, 1.7, { fill: item.fill, line: item.color, shadow: false });
    slide.addText(item.title, {
      x: x + 0.2, y: 1.42, w: 2.32, h: 0.3,
      fontFace: FONT, fontSize: 14, bold: true, color: item.color, margin: 0,
    });
    slide.addText(item.body, {
      x: x + 0.2, y: 1.9, w: 2.32, h: 0.65,
      fontFace: FONT, fontSize: 10.5, color: COLORS.ink, margin: 0, valign: "top", fit: "shrink",
    });
  });
  addCard(slide, 0.55, 3.15, 8.78, 1.55, { fill: "F7F9FC", line: COLORS.border, shadow: false });
  slide.addText("数据来源", {
    x: 0.8, y: 3.43, w: 0.9, h: 0.25,
    fontFace: FONT, fontSize: 12, bold: true, color: COLORS.primary, margin: 0,
  });
  const sourceLines = [
    `${data.source.name}：${data.source.url}`,
    data.secondarySource ? `${data.secondarySource.name}：${data.secondarySource.url}` : undefined,
    ...data.alerts.map((alert) => `${alert.title}：${alert.sourceUrl}`),
    `获取时间：${data.retrievedAt}`,
    "说明：温度统计和建议由上述七日预报字段确定性计算；未录入的数据不作推断。",
  ].filter(Boolean);
  slide.addText(sourceLines.join("\n"), {
    x: 1.75, y: 3.35, w: 7.25, h: 1.05,
    fontFace: FONT, fontSize: 8.5, color: COLORS.ink, margin: 0, breakLine: false, fit: "shrink",
  });
  addFooter(slide, data, 6, total);

  // 7. End
  slide = pptx.addSlide();
  addBackground(slide, TITLE_BG);
  slide.addText("谢谢", {
    x: 3.2, y: 1.45, w: 3.6, h: 0.9,
    fontFace: FONT, fontSize: 38, bold: true, color: COLORS.white, align: "center", margin: 0,
  });
  slide.addText("天气持续变化，请关注官方临近预报与预警", {
    x: 2.5, y: 2.45, w: 5.0, h: 0.35,
    fontFace: FONT, fontSize: 13, color: "DCE4F2", align: "center", margin: 0,
  });
  slide.addText(`汇报人：${data.reporter}   |   数据获取：${data.retrievedAt}`, {
    x: 4.0, y: 5.12, w: 5.5, h: 0.2,
    fontFace: FONT, fontSize: 9, color: "44546A", align: "right", margin: 0, fit: "shrink",
  });

  return pptx;
}

async function main() {
  const [inputArg, outputArg] = process.argv.slice(2);
  if (!inputArg || !outputArg) {
    console.error("Usage: node render_weather_report.js <weather.json> <output.pptx>");
    process.exit(2);
  }
  const inputPath = path.resolve(inputArg);
  const outputPath = path.resolve(outputArg);
  const raw = JSON.parse(fs.readFileSync(inputPath, "utf8"));
  const data = normalizeInput(raw);
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  const pptx = createDeck(data);
  await pptx.writeFile({ fileName: outputPath });
  console.log(`PPTX=${outputPath}`);
  console.log(`SLIDES=7`);
  console.log(`SOURCE=${data.source.url}`);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
