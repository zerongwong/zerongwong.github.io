// 信心与信任：从消费看中国经济形势与市场监管的作为 —— 讲座课件生成脚本（第二版）
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));
const fs = require("fs");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";
const FONT = "Microsoft YaHei";

// ---------- 主题与配色 ----------
const THEME = {
  name: "Xinxin Xinren",
  headFontFace: FONT,
  bodyFontFace: FONT,
  colors: {
    dk1: "22272E", lt1: "FFFFFF", dk2: "0B3B3C", lt2: "EEF5F4",
    accent1: "00897B", accent2: "C8442D", accent3: "4A63A6", accent4: "B07D0E",
    accent5: "4B5563", accent6: "E0B252", hlink: "4A63A6", folHlink: "4B5563",
  },
};
const HX = {
  ink: "22272E", jade: "0B3B3C", tint: "EEF5F4", white: "FFFFFF",
  teal: "00897B", red: "C8442D", blue: "4A63A6", gold: "B07D0E",
  muted: "4B5563", goldLight: "E0B252", redTint: "FBEFEC", gray: "F3F4F6",
  grid: "E5E7EB", axis: "4B5563", context: "A9B4BE", line: "D1D5DB", darkText: "D7E4E2", slate: "6B7280",
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "信心与信任：从消费看中国经济形势与市场监管的作为";
pres.subject = "全省市场监管系统干部专题讲座";
pres.author = "浙江财经大学经济学院";
pres.company = "浙江省市场监督管理局";
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
const C = pres.SchemeColor;
const K = {
  ink: C.text1, jade: C.text2, white: C.background1, tint: C.background2,
  teal: C.accent1, red: C.accent2, blue: C.accent3, gold: C.accent4, muted: C.accent5, goldLight: C.accent6,
};

const W = 13.333, MX = 0.6, CW = W - 2 * MX;
const stepOf = (n, w) => (CW - w) / (n - 1); // 成组卡片等距排列，左右两端与版心对齐

// ---------- 版式 ----------
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: HX.white },
  margin: [0.5, 0.6, 0.6, 0.6],
  objects: [
    { text: { text: "信心与信任 ｜ 从消费看中国经济形势与市场监管的作为", options: { x: MX, y: 7.05, w: 8, h: 0.28, fontSize: 10, color: "8A929C", fontFace: FONT, margin: 0 } } },
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.7, w: CW, h: 0.7, fontSize: 26, bold: true, color: HX.jade, fontFace: FONT, valign: "middle", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.13, y: 7.05, w: 0.6, h: 0.28, fontSize: 10, color: "8A929C", fontFace: FONT, align: "right" },
});
pres.defineSlideMaster({
  title: "DARK",
  background: { color: HX.jade },
  objects: [],
  slideNumber: { x: 12.13, y: 7.05, w: 0.6, h: 0.28, fontSize: 10, color: "7FA3A0", fontFace: FONT, align: "right" },
});

// ---------- 通用工具 ----------
function T(slide, text, opts) {
  const o = Object.assign({ fontFace: FONT, color: K.ink, fontSize: 14, margin: 0, valign: "top", isTextBox: true, lang: "zh-CN" }, opts);
  slide.addText(typeof text === "string" && /\*\*|⏎/.test(text) ? runs(text, { fontSize: o.fontSize, color: o.color, bold: o.bold }) : text, o);
}
// “**重点**”→加粗着色；“⏎”→段内软换行
function runs(str, base, hl) {
  const out = [];
  str.split("⏎").forEach((line, li) => {
    let first = true;
    line.split(/\*\*(.+?)\*\*/).forEach((seg, i) => {
      if (!seg) return;
      const o = Object.assign({}, base, i % 2 ? { bold: true, color: hl || K.red } : {});
      if (li > 0 && first) o.softBreakBefore = true;
      first = false;
      out.push({ text: seg, options: o });
    });
  });
  return out;
}
function paras(items, base, opt = {}) {
  const all = [];
  items.forEach((it, idx) => {
    const r = runs(it, base, opt.hl);
    if (opt.bullet !== false) r[0].options = Object.assign({}, r[0].options, { bullet: { code: "25A0", indent: 14 } });
    r[0].options = Object.assign({}, r[0].options, { paraSpaceAfter: opt.space == null ? 8 : opt.space });
    if (idx < items.length - 1) r[r.length - 1].options = Object.assign({}, r[r.length - 1].options, { breakLine: true });
    all.push(...r);
  });
  return all;
}
function box(slide, x, y, w, h, fill, opts = {}) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0 }, rectRadius: 0.08 }, opts));
}
function pill(slide, x, y, w, h, text, fill, color, size = 12) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0 }, rectRadius: h / 2 });
  T(slide, text, { x, y, w, h, fontSize: size, bold: true, color: color || K.white, align: "center", valign: "middle" });
}
function source(slide, text) {
  T(slide, "资料来源：" + text, { x: MX, y: 6.68, w: CW, h: 0.3, fontSize: 10, color: K.muted });
}
const SEALS = { 1: "观", 2: "析", 3: "鉴", 4: "策", 5: "为", 0: "引", 9: "结" };
const SECNAME = { 0: "开篇", 1: "一 · 观势", 2: "二 · 析因", 3: "三 · 镜鉴", 4: "四 · 谋策", 5: "五 · 作为", 9: "结语" };
const REG = {}; // 页码登记，供讲前更新清单引用
function content(sec, title, notes, key) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: CURRENT_SECTION });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: 0.3, w: 0.32, h: 0.32, fill: { color: K.red }, line: { color: K.red, width: 0 }, rectRadius: 0.05 });
  T(s, SEALS[sec], { x: MX, y: 0.3, w: 0.32, h: 0.32, fontSize: 12, bold: true, color: K.white, align: "center", valign: "middle" });
  T(s, SECNAME[sec], { x: MX + 0.45, y: 0.3, w: 6, h: 0.32, fontSize: 12, bold: true, color: K.red, valign: "middle" });
  s.addText(title, { placeholder: "title", fontFace: FONT, lang: "zh-CN" });
  if (notes) s.addNotes(notes);
  if (key) REG[key] = pres.slides.length;
  return s;
}
let CURRENT_SECTION = "开篇";
function section(title) { CURRENT_SECTION = title; pres.addSection({ title }); }

// 右侧“看点”面板
function insight(slide, items, opt = {}) {
  const x = opt.x || 8.85, y = opt.y || 1.6, w = opt.w || 3.88, h = opt.h || 4.95;
  box(slide, x, y, w, h, K.tint);
  T(slide, opt.head || "看点", { x: x + 0.25, y: y + 0.2, w: w - 0.5, h: 0.35, fontSize: 16, bold: true, color: K.teal });
  const regH = opt.reg ? (opt.regH || 1.3) : 0;
  T(slide, paras(items, { fontSize: opt.size || 14, color: K.ink }, { space: 8 }), { x: x + 0.25, y: y + 0.65, w: w - 0.5, h: h - 0.85 - regH, lineSpacingMultiple: 1.1 });
  if (opt.reg) {
    const ry = y + h - regH - 0.15;
    box(slide, x + 0.15, ry, w - 0.3, regH, HX.redTint);
    T(slide, [{ text: "监管视角  ", options: { bold: true, color: K.red, fontSize: 14 } }, ...runs(opt.reg, { fontSize: 14, color: K.ink })],
      { x: x + 0.3, y: ry + 0.08, w: w - 0.6, h: regH - 0.16, lineSpacingMultiple: 1.08, valign: "middle" });
  }
}
function chartCap(slide, text, x = MX, y = 1.55, w = 7.95) {
  T(slide, text, { x, y, w, h: 0.34, fontSize: 14, bold: true, color: K.ink });
}
function chartOpts(extra) {
  return Object.assign({
    catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT, dataLabelFontFace: FONT, legendFontFace: FONT, titleFontFace: FONT,
    catAxisLabelColor: "374151", valAxisLabelColor: HX.axis, catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
    valGridLine: { color: HX.grid, size: 0.75 }, catGridLine: { style: "none" },
    catAxisLineShow: true, catAxisLineColor: "9CA3AF", valAxisLineShow: false,
    dataLabelFontSize: 12, dataLabelColor: HX.ink,
    showLegend: false, legendPos: "t", legendFontSize: 13, legendColor: "374151",
    showTitle: false, barGapWidthPct: 55,
  }, extra);
}
const ICON_CACHE = {};
async function icon(name, color = "#FFFFFF") {
  const key = name + color;
  if (!ICON_CACHE[key]) {
    const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fa[name], { color, size: 256 }));
    ICON_CACHE[key] = "image/png;base64," + (await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer()).toString("base64");
  }
  return ICON_CACHE[key];
}
async function iconCircle(slide, name, x, y, d, fill) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 } });
  const p = d * 0.27;
  slide.addImage({ data: await icon(name), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
}

// 分篇扉页
const PARTS = ["观势", "析因", "镜鉴", "谋策", "作为"];
function divider(n, title, question, notes) {
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: CURRENT_SECTION });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 2.05, w: 1.6, h: 1.6, fill: { color: K.red }, line: { color: K.red, width: 0 }, rectRadius: 0.12 });
  T(s, SEALS[n], { x: 1.0, y: 2.05, w: 1.6, h: 1.6, fontSize: 66, bold: true, color: K.white, align: "center", valign: "middle" });
  T(s, `第${"一二三四五"[n - 1]}部分 · ${PARTS[n - 1]}`, { x: 3.05, y: 2.0, w: 9, h: 0.45, fontSize: 18, bold: true, color: K.goldLight });
  T(s, title, { x: 3.05, y: 2.5, w: 9.7, h: 0.95, fontSize: 38, bold: true, color: K.white, valign: "middle" });
  T(s, "— " + question, { x: 3.05, y: 3.6, w: 9.5, h: 1.0, fontSize: 20, color: HX.darkText, lineSpacingMultiple: 1.15 });
  PARTS.forEach((p, i) => {
    const active = i === n - 1;
    T(s, `${"一二三四五"[i]} ${p}`, { x: 1.0 + i * 2.35, y: 5.95, w: 2.1, h: 0.4, fontSize: 15, bold: active, color: active ? K.goldLight : "7FA3A0" });
  });
  s.addShape(pres.shapes.LINE, { x: 1.0, y: 5.8, w: 11.33, h: 0, line: { color: "2E5F5F", width: 0.75 } });
  if (notes) s.addNotes(notes);
  return s;
}

// ---------- 数据（讲前更新时只改这里） ----------
const D = {
  gdpYears: ["2012", "2013", "2014", "2015", "2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025", "2026\n上半年"],
  gdp: [7.9, 7.8, 7.4, 7.0, 6.8, 6.9, 6.7, 6.0, 2.2, 8.4, 3.0, 5.2, 5.0, 5.0, 4.7],
  nrYears: ["2019", "2020", "2021", "2022", "2023", "2024", "2025", "2026上半年"],
  nominal: [7.3, 2.7, 13.4, 4.8, 4.6, 4.2, 4.0, 5.3],
  real: [6.0, 2.2, 8.4, 3.0, 5.2, 5.0, 5.0, 4.7],
  engYears: ["2019", "2020", "2021", "2022", "2023", "2024", "2025"],
  engC: [3.5, -0.5, 5.3, 1.0, 4.3, 2.2, 2.6],
  engI: [1.7, 2.2, 1.1, 1.5, 1.5, 1.3, 0.8],
  engX: [0.8, 0.6, 1.7, 0.5, -0.6, 1.5, 1.6],
  rsLabels: ["25年1-2月", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "26年1-2月", "3", "4", "5", "6", "7", "8"],
  rs: [4.0, 5.9, 5.1, 6.4, 4.8, 3.7, 3.4, 3.0, 2.9, 1.3, 0.9, 2.8, 1.7, 0.2, -0.6, 1.0, 0.6, 0.4],
  pLabels: ["25年1月", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "26年1月", "2", "3", "4", "5", "6", "7", "8", "9"],
  cpi: [0.5, -0.7, -0.1, -0.1, -0.1, 0.1, 0.0, -0.4, -0.3, 0.2, 0.7, 0.8, 0.2, 1.3, 1.0, 1.2, 1.2, 1.0, 0.5, 0.8, 0.9],
  ppi: [-2.3, -2.2, -2.5, -2.7, -3.3, -3.6, -3.6, -2.9, -2.3, -2.1, -2.2, -1.9, -1.4, -0.9, 0.5, 2.8, 3.9, 4.1, 3.5, 3.8, 5.0],
  intl: [["美国", 68], ["英国", 62], ["印度", 60], ["日本", 55], ["法国", 53], ["德国", 50], ["韩国", 48], ["中国", 39.9]],
  incYears: ["2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025", "2026上半年"],
  incGDP: [6.8, 6.9, 6.7, 6.0, 2.2, 8.4, 3.0, 5.2, 5.0, 5.0, 4.7],
  incReal: [6.3, 7.3, 6.5, 5.8, 2.1, 8.1, 2.9, 6.1, 5.1, 5.0, 4.2],
  conReal: [6.8, 5.4, 6.2, 5.5, -4.0, 12.6, -0.2, 9.0, 5.1, 4.4, 2.7],
  apcYears: ["2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025"],
  apc: [70.5, 70.3, 70.1, 65.9, 68.6, 66.5, 68.3, 68.3, 68.0],
  depYears: ["2019", "2020", "2021", "2022", "2023", "2024", "2025"],
  dep: [9.7, 11.3, 9.9, 17.8, 16.7, 14.3, 14.6],
  houseYears: ["2019", "2020", "2021", "2022", "2023", "2024", "2025"],
  house: [17.16, 17.61, 17.94, 13.58, 11.17, 9.74, 8.81],
  svcYears: ["2019", "2020", "2021", "2022", "2023", "2024", "2025"],
  svc: [45.9, 42.6, 44.2, 43.2, 45.2, 46.1, 46.1],
  jpYears: Array.from({ length: 26 }, (_, i) => String(1990 + i)),
  jpCPI: [3.1, 3.3, 1.6, 1.3, 0.7, -0.1, 0.1, 1.8, 0.6, -0.3, -0.7, -0.7, -0.9, -0.3, 0.0, -0.3, 0.3, 0.0, 1.4, -1.4, -0.7, -0.3, 0.0, 0.4, 2.7, 0.8],
  aug: [["汽车类", -18.5], ["社零总额", 0.4], ["除汽车外", 2.5], ["通讯器材", 27.3]],
};

// ====================================================================
async function build() {
  // ---------------- 封面 ----------------
  section("开篇");
  {
    const s = pres.addSlide({ masterName: "DARK", sectionTitle: CURRENT_SECTION });
    T(s, "全省市场监管系统干部专题讲座", { x: 1.0, y: 0.9, w: 8, h: 0.4, fontSize: 16, color: K.goldLight, bold: true });
    T(s, "信心与信任", { x: 1.0, y: 1.85, w: 8.5, h: 1.4, fontSize: 66, bold: true, color: K.white, valign: "middle" });
    T(s, "从消费看中国经济形势与市场监管的作为", { x: 1.0, y: 3.3, w: 9, h: 0.6, fontSize: 26, color: HX.darkText });
    s.addShape(pres.shapes.LINE, { x: 1.0, y: 4.3, w: 6.2, h: 0, line: { color: "2E5F5F", width: 0.75 } });
    T(s, "主讲人：〔姓名〕", { x: 1.0, y: 4.5, w: 9, h: 0.45, fontSize: 20, bold: true, color: K.white });
    T(s, "浙江财经大学经济学院 ｜ 挂职浙江省市场监督管理局经济发展处", { x: 1.0, y: 5.0, w: 9.5, h: 0.4, fontSize: 15, color: HX.darkText });
    T(s, "2026年10月20日", { x: 1.0, y: 5.45, w: 6, h: 0.4, fontSize: 15, color: HX.darkText });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.85, y: 1.75, w: 2.3, h: 2.3, fill: { color: K.red }, line: { color: K.red, width: 0 }, rectRadius: 0.15 });
    T(s, "信", { x: 9.85, y: 1.75, w: 2.3, h: 2.3, fontSize: 110, bold: true, color: K.white, align: "center", valign: "middle" });
    T(s, "信心 · 信任", { x: 9.85, y: 4.2, w: 2.3, h: 0.4, fontSize: 14, color: K.goldLight, align: "center" });
    s.addNotes("【开场，约1分钟】各位领导、各位同事，大家好。先做个自我介绍：我是研究居民消费的，今年在省局经济发展处挂职。挂职这段时间最大的感受是：经济学家看消费用的是“望远镜”，各位手里拿的是“显微镜”。今天想把两个镜头对一对。\n题目叫“信心与信任”：宏观政策解决“信心”——老百姓有没有钱花、敢不敢花；市场监管解决“信任”——买不买得到好东西、放不放心买。\n【时间分配，合计约115分钟】开篇5｜观势13｜析因27｜休息10｜镜鉴8｜谋策12｜作为29｜结语与答疑11。\n【讲前】据国家统计局发布日程表，前三季度数据于10月19日（周一）10:00发布（以官网为准），当晚按附录隐藏页清单更新；9月CPI、PPI已于10月14日发布并已更新。");
  }

  // ---------------- 三个数字 ----------------
  {
    const s = content(0, "三个数字：经济在增长，消费在放缓，份额在低位",
      "【约2分钟】（若前三季度数据已发布，开头改为：“昨天上午国家统计局刚发布前三季度数据……”）\n先看三个数字。4.7%——上半年GDP实际增长，名义增长约5.3%；2.5%——今年起统计局发布了包含服务的零售总额，1—8月只增长2.5%，其中商品零售（也就是我们常说的社零）只有1.1%，都明显慢于名义GDP；39.9%——居民消费占GDP的比重，大约四成，主要经济体多在五成上下到近七成。三个数字放在一起，就是今天要讨论的“温差”：有增长、缺消费。\n【举手投票，约1分钟】消费不足，您觉得最主要的原因是？A收入不够 B房子拖累 C不敢花 D不放心。请助教记下大致比例，第22页回扣。\n【贯穿案例】再留一个问题：国家补贴买空调，一台补600元——这600元最后落进了谁的口袋？我们第四部分揭晓。", "s_three");
    const cols = [
      ["4.7%", "2026年上半年GDP实际增速", "名义增长约5.3%；⏎全年目标4.5%—5%", K.teal, K.tint],
      ["2.5%", "1—8月含服务的零售总额增速", "其中商品零售（社零）仅+1.1%；⏎均明显慢于名义GDP", K.red, HX.redTint],
      ["39.9%", "居民消费率（2024年）", "主要经济体多在48%—68%；⏎“十五五”要求明显提高", K.blue, K.tint],
    ];
    const st = stepOf(3, 3.85);
    cols.forEach((c, i) => {
      const x = MX + i * st;
      box(s, x, 1.75, 3.85, 3.1, c[4]);
      T(s, c[0], { x: x + 0.35, y: 1.95, w: 3.3, h: 1.15, fontSize: 60, bold: true, color: c[3], valign: "bottom" });
      T(s, c[1], { x: x + 0.35, y: 3.2, w: 3.3, h: 0.4, fontSize: 16, bold: true });
      T(s, c[2], { x: x + 0.35, y: 3.7, w: 3.3, h: 0.95, fontSize: 14, color: K.muted, lineSpacingMultiple: 1.15 });
    });
    box(s, MX, 5.15, CW, 1.25, K.jade);
    T(s, [
      { text: "今天的核心问题  ", options: { bold: true, color: K.goldLight, fontSize: 17 } },
      { text: "为什么“有增长、缺消费”？哪些堵点是市场监管够得着、管得了、管得好的？", options: { color: K.white, fontSize: 17 } },
    ], { x: MX + 0.4, y: 5.15, w: CW - 0.8, h: 1.25, valign: "middle" });
    source(s, "国家统计局（含服务的零售总额为“社会消费商品和服务零售总额”）；世界银行WDI。名义GDP增速为作者据现价GDP计算。");
  }

  // ---------------- 一条主线 ----------------
  {
    const s = content(0, "一条主线：宏观政策管“信心”，市场监管管“信任”",
      "【约2分钟】这是全课的主线图。信心，回答的是“有没有钱花、敢不敢花”：取决于收入、就业、资产、社保和预期——这是宏观政策和民生政策的主场。信任，回答的是“买不买得到好东西、放不放心买”：取决于质量、价格、维权、平台规则——这是市场监管的主场。\n为什么用乘号？因为两者互相放大：同样一份增收、同一笔补贴，在信任度高的市场里能转化为更多真实消费。但要实事求是：当前首要短板在信心；信任，是我们能直接改变的那一项。\n还要补一句：监管也稳“信心”——对经营主体来说，规则稳定、检查规范，本身就是稳预期。", "s_main");
    const pillar = (x, head, sub, tag, items, fill, col) => {
      box(s, x, 1.65, 5.2, 3.75, fill);
      T(s, head, { x: x + 1.35, y: 1.8, w: 3.6, h: 0.55, fontSize: 26, bold: true, color: col, valign: "middle" });
      T(s, sub, { x: x + 1.35, y: 2.35, w: 3.7, h: 0.35, fontSize: 15, color: K.ink });
      T(s, tag, { x: x + 1.35, y: 2.7, w: 3.7, h: 0.32, fontSize: 13, bold: true, color: col });
      items.forEach((it, i) => {
        T(s, [{ text: it[0], options: { bold: true, fontSize: 15 } }, { text: "  " + it[1], options: { fontSize: 14, color: K.muted } }], { x: x + 0.45, y: 3.2 + i * 0.5, w: 4.6, h: 0.45, valign: "middle" });
      });
    };
    pillar(MX, "信心", "有没有钱花、敢不敢花", "宏观与民生政策的主场", [["收入与就业", "有没有钱、稳不稳"], ["资产与财富", "房子、存款、股票"], ["社会保障", "养老、医疗、教育"], ["未来预期", "对明天的判断"]], K.tint, K.teal);
    pillar(7.53, "信任", "买不买得到好东西、放不放心买", "市场监管的主场", [["质量与安全", "买到的是不是好东西"], ["价格与竞争", "价格公不公道"], ["维权与信用", "出了问题能不能解决"], ["平台与新业态", "规则公不公平"]], HX.redTint, K.red);
    await iconCircle(s, "FaChartLine", MX + 0.35, 1.85, 0.8, K.teal);
    await iconCircle(s, "FaHandshake", 7.88, 1.85, 0.8, K.red);
    s.addShape(pres.shapes.OVAL, { x: 6.12, y: 3.0, w: 1.1, h: 1.1, fill: { color: K.jade }, line: { color: K.white, width: 2 } });
    T(s, "×", { x: 6.12, y: 3.0, w: 1.1, h: 1.1, fontSize: 40, bold: true, color: K.white, align: "center", valign: "middle" });
    T(s, "消费 ≈ 信心 × 信任", { x: MX, y: 5.55, w: CW, h: 0.55, fontSize: 24, bold: true, color: K.jade, align: "center", valign: "middle" });
    T(s, "两者互相放大：同样的增收与补贴，在信任度高的市场里能转化为更多真实消费。⏎当前首要短板在信心；信任，是市场监管能直接改变的那一项。",
      { x: MX, y: 6.1, w: CW, h: 0.6, fontSize: 14, color: K.ink, align: "center" });
  }

  // ---------------- 内容安排 ----------------
  {
    const s = content(0, "内容安排：五个部分，一条逻辑链",
      "【约1分钟】五个部分按照“是什么—为什么—别人怎么做—中央怎么部署—我们怎么干”展开。前两部分偏宏观，后三部分落到市场监管，第五部分时间最长。中间休息10分钟，最后留出答疑时间。");
    const parts = [
      ["观", "观势", "增长的“温差”从何而来", "GDP、名义与实际、三大需求、社零、价格", "约13分钟"],
      ["析", "析因", "消费为什么“提不起来”", "收入、预期、供给、信任四道门槛", "约27分钟"],
      ["鉴", "镜鉴", "他山之石：消费大国的制度底座", "美欧消费者保护、日本教训", "约8分钟"],
      ["策", "谋策", "从“全方位扩内需”到“坚持内需主导”", "政策脉络、以旧换新、补贴归宿", "约12分钟"],
      ["为", "作为", "市场监管如何为消费“增信任”", "职能逻辑、五个抓手、浙江实践、带走清单", "约29分钟"],
    ];
    parts.forEach((p, i) => {
      const y = 1.6 + i * 0.98;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y, w: 0.78, h: 0.78, fill: { color: i < 2 ? K.teal : K.red }, line: { color: K.white, width: 0 }, rectRadius: 0.08 });
      T(s, p[0], { x: MX, y, w: 0.78, h: 0.78, fontSize: 28, bold: true, color: K.white, align: "center", valign: "middle" });
      T(s, [{ text: `${"一二三四五"[i]}、${p[1]} ｜ `, options: { bold: true, color: i < 2 ? K.teal : K.red, fontSize: 20 } }, { text: p[2], options: { bold: true, fontSize: 20 } }], { x: 1.65, y, w: 8.8, h: 0.45, valign: "middle" });
      T(s, p[3], { x: 1.65, y: y + 0.45, w: 8.8, h: 0.33, fontSize: 14, color: K.muted });
      T(s, p[4], { x: 10.9, y, w: 1.83, h: 0.78, fontSize: 15, color: K.muted, align: "right", valign: "middle" });
    });
  }

  // ===================== 第一部分 观势 =====================
  section("一、观势");
  divider(1, "增长的“温差”从何而来", "如何看待宏观数据与微观感受的“温差”？",
    "【过渡，约30秒】第一部分“观势”。我们常听到一种说法：数据挺好，体感不好。这个“温差”是真实存在的，而且有明确的经济学解释。");

  {
    const s = content(1, "增速换挡：中国经济进入“5%左右”的中速增长阶段",
      "【约1.5分钟】一句话结论：从7%—8%换挡到5%左右，是潜在增长率下行的结构性变化——人口、资本回报、技术追赶空间都在变。\n2021年的8.4%是低基数反弹，不是趋势。今年上半年4.7%，一季度5.0%、二季度4.3%，二季度放缓的主要拖累是投资：1—8月固定资产投资下降7.2%。\n【讲前】前三季度数据发布后，把最后一根柱子换成“2026前三季度”。", "s_gdp");
    chartCap(s, "图1  GDP实际增速（%）");
    s.addChart(pres.charts.BAR, [{ name: "GDP增速", labels: D.gdpYears, values: D.gdp }], chartOpts({
      x: MX, y: 1.95, w: 7.95, h: 4.6, barDir: "col", chartColors: D.gdp.map((v, i) => (i === D.gdp.length - 1 ? HX.red : HX.teal)),
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", valAxisMinVal: 0, valAxisMaxVal: 10, valAxisMajorUnit: 2, catAxisLabelFontSize: 11, catAxisLabelRotate: 0,
    }));
    insight(s, [
      "从7%—8%换挡到5%左右：潜在增长率下行是**结构性**的",
      "2026年目标**4.5%—5%**；上半年4.7%，二季度4.3%，**投资下降**是主要拖累",
      "同样5%，靠外需和投资撑起的增长，与靠消费撑起的增长，**质量不同**",
    ], { reg: "换挡期增长更依赖**内需与效率**：降低制度性交易成本，本身就是稳增长政策。" });
    source(s, "国家统计局。2022年及以前、2024年为最终核实数，2023年增速为初步核算数，2025年为初步核算数，2026年为上半年同比。");
  }

  {
    const s = content(1, "体感温差之一：名义增速连续三年低于实际，今年上半年才重新反超",
      "【约2.5分钟】各位有没有听企业说过“货比去年卖得多，钱比去年赚得少”？这就是名义增速低于实际增速。好比“谷贱伤农”：粮食多打了一成，粮价跌了两成，农民反而减收。\n营收、利润、税收、工资都按名义值计算。2023—2025年名义增速约4.6%、4.2%、4.0%，都低于5%左右的实际增速——GDP平减指数连续为负，价格总水平低位运行，这是“体感温差”的第一个来源。\n今年上半年名义增长约5.3%，重新高于实际的4.7%。但要看清动力：主要是国际大宗商品、AI硬件涨价和“反内卷”带来的供给收缩，核心CPI只有1%——这不是需求拉动型的回升。", "s_nominal");
    chartCap(s, "图2  名义与实际GDP增速（%）");
    s.addChart(pres.charts.BAR, [
      { name: "名义增速", labels: D.nrYears, values: D.nominal },
      { name: "实际增速", labels: D.nrYears, values: D.real },
    ], chartOpts({
      x: MX, y: 1.95, w: 7.95, h: 4.6, barDir: "col", barGrouping: "clustered", chartColors: [HX.blue, HX.teal],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", showLegend: true, valAxisMinVal: 0, valAxisMaxVal: 15, valAxisMajorUnit: 3, catAxisLabelFontSize: 11,
    }));
    insight(s, [
      "营收、利润、税收、工资都按**名义值**计价：名义增速才是“体感增速”",
      "2023—2025年名义增速约**4.6%、4.2%、4.0%**，均低于实际：量增价跌",
      "2026年上半年名义约**5.3%**，重回实际之上；但主要来自大宗商品与AI硬件涨价、供给收缩，**核心CPI仅1.0%**",
    ], { reg: "价格低位运行时，重在治理低价无序竞争；回升期防借涨价之名**哄抬、串通**，但不干预正常的成本传导。", regH: 1.5 });
    source(s, "国家统计局；作者据现价GDP计算（2026年上半年69.57万亿元，2025年同期66.05万亿元）。");
  }

  {
    const s = content(1, "增长靠谁拉动：外需“补位”明显，消费拉动仍不够强",
      "【约2分钟】把每年的增长拆成消费、投资、净出口三块。2025年净出口拉动1.6个百分点、贡献率32.7%，是近年高位；投资只拉动0.8个百分点。外需在给内需“补位”，但关税和地缘不确定性在上升，补位不可持续。\n提醒一个常见误读：贡献率是份额。投资塌下去的时候，消费的份额会被动上升，所以看消费要看“拉动了几个百分点”——2025年2.6个点，今年上半年约2.1个点。\n【监管含义】国内大市场是最大底气。公平竞争审查、破除地方保护，是市场监管对全国统一大市场的直接贡献。", "s_eng");
    chartCap(s, "图3  三大需求对GDP增长的拉动（百分点）");
    s.addChart(pres.charts.BAR, [
      { name: "最终消费支出", labels: D.engYears, values: D.engC },
      { name: "资本形成总额", labels: D.engYears, values: D.engI },
      { name: "货物和服务净出口", labels: D.engYears, values: D.engX },
    ], chartOpts({
      x: MX, y: 1.95, w: 7.95, h: 4.6, barDir: "col", barGrouping: "stacked", chartColors: [HX.teal, HX.gold, HX.blue],
      showValue: true, dataLabelPosition: "ctr", dataLabelFormatCode: "0.0", dataLabelColor: HX.white, dataLabelFontBold: true, showLegend: true, valAxisMinVal: -2, valAxisMaxVal: 10, valAxisMajorUnit: 2, barGapWidthPct: 70,
    }));
    insight(s, [
      "2025年：消费拉动**2.6**、投资**0.8**、净出口**1.6**个百分点；净出口贡献率**32.7%**",
      "外需“补位”不可持续：关税与地缘不确定性上升，必须**坚持内需主导**",
      "贡献率是份额：投资走弱时消费份额被动上升，看消费要看**拉动的百分点**（2026年上半年约2.1）",
    ], { reg: "国内大市场是最大底气：用好**公平竞争审查**、破除地方保护，纵深推进**全国统一大市场**建设。", regH: 1.5 });
    source(s, "国家统计局（各年初步核算口径，个别年份与图1最终核实数相差0.1—0.3个百分点）；2026年上半年消费拉动据媒体转述。");
  }

  {
    const s = content(1, "消费“脉冲”在退坡：社零增速从6.4%回落到0.4%",
      "【约2.5分钟】绿色是2025年，红色是2026年。去年5月6.4%是阶段高点，今年5月一度转负，8月0.4%。\n要分清两件事：一是高基数——去年以旧换新和“6·18”抬高了基数；二是需求前置——本来明年后年才换的车和家电，提前到去年买了。用两年复合增速剔除基数，5月约2.8%、8月约1.9%：偏弱，但不是断崖。\n贯穿案例：那台空调，是不是“提前买的”？如果是，今年的补贴就是在“借明年的需求”。\n结构上，8月通讯器材增长27.3%（其中有存储芯片涨价的价格因素），汽车下降18.5%；1—8月服务零售增长4.9%，快于商品零售。\n【监管含义】补贴期防“先提价后打折”，退坡期防清库存中的虚假折扣、翻新冒新。", "s_rs");
    chartCap(s, "图4  社会消费品零售总额当月同比（%）");
    const colors = D.rs.map((v, i) => (i < 11 ? HX.teal : HX.red));
    s.addChart(pres.charts.BAR, [{ name: "社零增速", labels: D.rsLabels, values: D.rs }], chartOpts({
      x: MX, y: 1.95, w: 7.95, h: 4.6, barDir: "col", chartColors: colors, invertedColors: colors, showValue: true, dataLabelPosition: "outEnd",
      dataLabelFormatCode: "0.0", dataLabelFontSize: 11, valAxisMinVal: -2, valAxisMaxVal: 8, valAxisMajorUnit: 2, catAxisLabelFontSize: 11, barGapWidthPct: 40,
    }));
    T(s, "■ 2025年", { x: 5.5, y: 1.57, w: 1.3, h: 0.32, fontSize: 13, color: K.teal, bold: true });
    T(s, "■ 2026年", { x: 6.9, y: 1.57, w: 1.3, h: 0.32, fontSize: 13, color: K.red, bold: true });
    insight(s, [
      "2026年5月一度转负，8月仅**0.4%**；剔除高基数的两年复合增速：5月约**2.8%**、8月约**1.9%**，偏弱而非断崖",
      "8月通讯器材**+27.3%**（含芯片涨价因素），汽车**-18.5%**；1—8月服务零售**+4.9%**",
      "耐用品补贴的**需求前置**叠加去年**高基数**：同比读数被双重压低",
    ], { reg: "补贴期防“**先提价后打折**”；退坡期防清库存中的**虚假折扣、翻新冒新**。" });
    source(s, "国家统计局。2025年起按第五次经济普查修订后的可比口径计算增速；两年复合增速为作者计算。");
  }

  {
    const s = content(1, "价格信号在反转：PPI升至5.0%，核心CPI仍在1%",
      "【约2分钟】红线是PPI，绿线是CPI。PPI在今年3月由负转正，结束了连续41个月的下降；10月14日发布的9月数据升至5.0%，是2022年7月以来最高。统计局解读：主要是国际大宗商品价格上行和人工智能相关需求，煤炭开采、石油天然气开采、电气机械等行业涨幅扩大。\n这些品类在CPI篮子里权重很小。9月CPI上涨0.9%，主要靠汽油（+16.6%）和电子产品；真正说明终端需求的是核心CPI——仍然只有1.0%。夹在中间的中下游企业和个体户最难受：进价涨了，售价不敢涨。\n【监管含义】“两线作战”：上游防囤积哄抬、串通涨价；下游防低价倾销、价格欺诈。", "s_price");
    chartCap(s, "图5  CPI与PPI当月同比（%），2025年1月—2026年9月");
    s.addChart(pres.charts.LINE, [
      { name: "PPI（工业生产者出厂价格）", labels: D.pLabels, values: D.ppi },
      { name: "CPI（居民消费价格）", labels: D.pLabels, values: D.cpi },
    ], chartOpts({
      x: MX, y: 1.95, w: 7.95, h: 4.6, chartColors: [HX.red, HX.teal], lineSize: 2.25, lineDataSymbol: "circle", lineDataSymbolSize: 6,
      showLegend: true, valAxisMinVal: -5, valAxisMaxVal: 5, valAxisMajorUnit: 1, catAxisLabelFontSize: 11, valAxisLabelFormatCode: "0",
    }));
    insight(s, [
      "PPI于3月**由负转正**，9月升至**5.0%**⏎为2022年7月以来最高",
      "9月CPI **0.9%**，**核心CPI仅1.0%**：终端需求偏弱；PPI涨幅集中在能源、原材料与AI相关行业",
      "中下游企业与个体户**两头受挤**：进价涨、售价难涨",
    ], { reg: "价格监管“两线作战”：上游防**囤积哄抬、串通涨价**；下游防**低价倾销、价格欺诈**。", regH: 1.3 });
    source(s, "国家统计局月度价格数据（2026年1月起CPI以2025年为基期）。");
  }

  {
    const s = content(1, "形势判断：供强需弱矛盾突出——短期拖累在投资，长期短板在消费",
      "【约2分钟，承上启下】四句话：总量稳、投资降、消费弱、价格转。\n这和中央的判断一致：2024年中央经济工作会议指出“国内需求不足”；2025年中央经济工作会议指出“国内供强需弱矛盾突出”，并把“坚持内需主导，建设强大国内市场”列为重点任务之首。\n我的判断：潜在增速换挡是结构性的，但价格低位运行说明需求不足真实存在。短期拖累最大的是投资收缩；长期最需要补的短板是居民消费——它是最可持续的需求来源。下面第二部分，专门拆解消费为什么提不起来。", "s_judge");
    const cards = [
      ["总量稳", "GDP +4.7%", "处于目标区间，就业物价总体可控", K.teal, "FaChartLine"],
      ["投资降", "投资 -7.2%", "1—8月；房地产-19.9%⏎民间投资-10.1%", K.red, "FaBalanceScaleLeft"],
      ["消费弱", "社零 +1.1%", "含服务零售+2.5%⏎居民多存少花", K.gold, "FaShoppingBasket"],
      ["价格转", "PPI +5.0%", "9月；核心CPI仅1.0%⏎传导不畅挤压中下游", K.blue, "FaSearchDollar"],
    ];
    const st = stepOf(4, 2.85);
    for (let i = 0; i < 4; i++) {
      const c = cards[i], x = MX + i * st;
      box(s, x, 1.65, 2.85, 3.0, K.tint);
      await iconCircle(s, c[4], x + 0.3, 1.85, 0.72, c[3]);
      T(s, c[0], { x: x + 1.15, y: 1.9, w: 1.6, h: 0.6, fontSize: 24, bold: true, color: c[3], valign: "middle" });
      T(s, c[1], { x: x + 0.3, y: 2.8, w: 2.45, h: 0.5, fontSize: 20, bold: true });
      T(s, runs(c[2], { fontSize: 15, color: K.ink }), { x: x + 0.3, y: 3.4, w: 2.35, h: 1.1, fontSize: 15, color: K.ink, lineSpacingMultiple: 1.15 });
    }
    box(s, MX, 4.9, CW, 1.65, K.jade);
    T(s, [
      { text: "中央判断  ", options: { bold: true, color: K.goldLight, fontSize: 15 } },
      { text: "2024年中央经济工作会议：“国内需求不足”；2025年中央经济工作会议：“国内供强需弱矛盾突出”，把“坚持内需主导，建设强大国内市场”列为重点任务之首。", options: { color: K.white, fontSize: 15, breakLine: true } },
      { text: "本讲判断  ", options: { bold: true, color: K.goldLight, fontSize: 15 } },
      { text: "潜在增速换挡是结构性的，但价格低位运行说明需求不足真实存在：短期拖累最大的是投资收缩，长期最需补的短板是居民消费。", options: { color: K.white, fontSize: 15 } },
    ], { x: MX + 0.35, y: 4.95, w: CW - 0.7, h: 1.55, valign: "middle", lineSpacingMultiple: 1.2 });
    source(s, "国家统计局；中央经济工作会议公开报道。");
  }

  // ===================== 第二部分 析因 =====================
  section("二、析因");
  divider(2, "消费为什么“提不起来”", "中国人天生爱储蓄吗？高储蓄是发展阶段、保障水平与未来预期共同作用的结果，而非“国民性”。",
    "【过渡，约30秒】第二部分是今天的重点之一。“中国人天生爱存钱”这个解释既不科学，也没有政策含义。我们用四道门槛把消费拆开。");

  {
    const s = content(2, "分析框架：消费要跨过“四道门槛”",
      "【约2分钟】一个家庭从“有钱”到“把钱花出去”，要依次回答四个问题：有没有钱？敢不敢花？有没有想买的？买了放不放心？\n经济学上可以写成：C = f(收入Y, 财富W, 预期E, 保障S; 供给质量Q, 信任环境T)。\n要讲清边界：前两道门槛是“信心”问题，主要靠宏观和民生政策；第三道是多部门共管，市场监管管的是标准、准入和质量；第四道才是我们的主责。后面的数据，就按这四道门槛展开。", "s_frame");
    const steps = [
      ["能消费", "收入", "有没有钱", "收入份额 · 增速 · 分配", K.teal, "FaWallet", "收入实际增速 4.2%"],
      ["敢消费", "预期", "敢不敢花", "财富 · 就业 · 保障", K.blue, "FaShieldAlt", "平均消费倾向 68.0%"],
      ["愿消费", "供给", "想不想买", "产品 · 服务 · 场景 · 标准", K.gold, "FaShoppingBasket", "服务消费占比 46.1%"],
      ["放心消费", "信任", "放不放心", "质量 · 价格 · 维权 · 信用", K.red, "FaHandshake", "投诉举报 2646万件"],
    ];
    const st = stepOf(4, 2.85);
    for (let i = 0; i < 4; i++) {
      const sp = steps[i], x = MX + i * st, y = 3.05 - i * 0.3, h = 2.75 + i * 0.3;
      box(s, x, y, 2.85, h, i < 2 ? K.tint : i === 2 ? "FBF4E4" : HX.redTint);
      await iconCircle(s, sp[5], x + 0.25, y + 0.25, 0.68, sp[4]);
      T(s, `${i + 1}  ${sp[0]}`, { x: x + 1.05, y: y + 0.28, w: 1.75, h: 0.6, fontSize: 20, bold: true, color: sp[4], valign: "middle" });
      T(s, `${sp[1]}：${sp[2]}`, { x: x + 0.25, y: y + 1.1, w: 2.45, h: 0.38, fontSize: 16, bold: true });
      T(s, sp[3], { x: x + 0.25, y: y + 1.5, w: 2.45, h: 0.38, fontSize: 14, color: K.muted });
      s.addShape(pres.shapes.LINE, { x: x + 0.25, y: 5.12, w: 2.35, h: 0, line: { color: HX.line, width: 0.75 } });
      T(s, sp[6], { x: x + 0.25, y: 5.2, w: 2.45, h: 0.45, fontSize: 14, bold: true, color: sp[4], valign: "middle" });
    }
    const band = (x, w, text, col, hex) => {
      T(s, text, { x, y: 1.5, w, h: 0.4, fontSize: 15, bold: true, color: col, align: "center" });
      s.addShape(pres.shapes.LINE, { x: x + 0.15, y: 1.95, w: w - 0.3, h: 0, line: { color: hex, width: 1.25 } });
    };
    band(MX, st + 2.85, "①② 信心 —— 宏观与民生政策的主场", K.teal, HX.teal);
    band(MX + 2 * st, 2.85, "③ 供给 —— 多部门共管", K.gold, HX.gold);
    band(MX + 3 * st, 2.85, "④ 信任 —— 市场监管主责", K.red, HX.red);
    T(s, "一个家庭把钱花出去，要依次回答四个问题：有没有钱？敢不敢花？有没有想买的？买了放不放心？", { x: MX, y: 6.0, w: CW, h: 0.45, fontSize: 16, color: K.jade, bold: true, align: "center" });
  }

  {
    const s = content(2, "结构性短板：居民消费率约四成，比日韩低8—15个百分点",
      "【约2.5分钟】中国居民消费率39.9%，韩国约48%，日本约55%，美国约68%。\n讲严谨性：中国的统计口径有低估，比如自有住房的虚拟房租；2023年统计局已改革自有住房服务核算，部分修正了低估。修正之后，和韩国仍差约8个百分点，和美国差近30个百分点——结构性短板是真实的。\n一个换算：居民消费率每提高1个百分点，按140万亿元GDP算，约1.4万亿元消费。“十五五”要求“居民消费率明显提高”，空间在这里。\n【分寸】消费率主要由收入分配、社保和人口结构决定，不是监管能决定的。", "s_intl");
    chartCap(s, "图6  居民消费支出占GDP比重：国际比较（%）");
    const labels = D.intl.map((d) => d[0]).reverse(), vals = D.intl.map((d) => d[1]).reverse();
    s.addChart(pres.charts.BAR, [{ name: "居民消费率", labels, values: vals }], chartOpts({
      x: MX, y: 1.95, w: 7.95, h: 4.6, barDir: "bar", chartColors: vals.map((v, i) => (labels[i] === "中国" ? HX.red : HX.context)),
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "General", valAxisMinVal: 0, valAxisMaxVal: 80, valAxisMajorUnit: 20, catAxisLabelFontSize: 14, barGapWidthPct: 45,
    }));
    insight(s, [
      "中国**39.9%**；韩国约48%、日本约55%、美国约68%",
      "口径有低估（虚拟房租等）；2023年起已改革核算，修正后与韩国**仍差约8个百分点**",
      "居民消费率每提高1个百分点 ≈ **1.4万亿元**消费",
      "“十五五”：**居民消费率明显提高**",
    ], { reg: "消费率主要靠收入分配与社保；监管的贡献在于减少“**不敢买、买不好**”造成的消费流失。", regH: 1.5 });
    source(s, "世界银行WDI（家庭及为家庭服务的非营利机构最终消费支出/GDP，2023年或最近可得年份，取整）；中国为国家统计局2024年数据。");
  }

  {
    const s = content(2, "能不能消费之一：居民收入份额偏低，财产性收入偏薄",
      "【约2.5分钟】一句话结论：差距更大的一块在“不敢花”，而不只是“没钱花”。\n按资金流量表口径，居民可支配收入约占GDP六成，和日韩相近；住户调查口径只有43%左右，因为漏报高收入和实物收入，不宜做国际比较。做个粗略分解：居民消费率≈收入份额×消费倾向≈60%×2/3≈40%。和发达经济体的差距里，消费倾向（也就是储蓄率）占了更大一块。\n结构上，财产净收入只占8%，2025年只增长1.6%——房租、利息、分红都在走弱。城乡收入比2.31，浙江是1.81、连续13年缩小，这是共同富裕示范区“能消费”的底气。\n【与监管的连接】经营净收入占16.7%，和个体工商户的经营环境直接相关——第五部分抓手四会讲。", "s_income");
    chartCap(s, "图7  2025年全国居民人均可支配收入构成（%）", MX, 1.55, 5.4);
    s.addChart(pres.charts.DOUGHNUT, [{ name: "收入构成", labels: ["工资性收入", "经营净收入", "转移净收入", "财产净收入"], values: [56.6, 16.7, 18.6, 8.0] }], chartOpts({
      x: MX, y: 1.95, w: 5.4, h: 4.6, chartColors: [HX.teal, HX.blue, HX.gold, HX.red], holeSize: 55, showPercent: false, showValue: true,
      dataLabelFormatCode: "0.0", dataLabelColor: HX.white, dataLabelFontBold: true, dataLabelFontSize: 13, showLegend: true, legendPos: "b", legendFontSize: 13, dataBorder: { pt: 2, color: "FFFFFF" },
    }));
    const kv = [["8.0%", "财产净收入占比", "2025年仅增长1.6%", K.red], ["2/3", "平均消费倾向", "每100元收入约花68元", K.teal], ["2.31", "城乡收入比", "浙江1.81，连续13年缩小", K.blue]];
    kv.forEach((k, i) => {
      const y = 1.75 + i * 1.6;
      T(s, k[0], { x: 6.2, y, w: 2.5, h: 0.65, fontSize: 30, bold: true, color: k[3], valign: "bottom" });
      T(s, k[1], { x: 6.2, y: y + 0.68, w: 2.5, h: 0.32, fontSize: 15, bold: true });
      T(s, k[2], { x: 6.2, y: y + 1.0, w: 2.5, h: 0.32, fontSize: 13, color: K.muted });
    });
    insight(s, [
      "资金流量表口径：居民可支配收入约占GDP**六成**；住户调查口径约43%，不宜做国际比较",
      "居民消费率≈收入份额×消费倾向≈60%×2/3：差距**更大一块在“不敢花”**",
      "低收入群体**边际消费倾向最高**：提低扩中是“高乘数”政策",
    ], { reg: "收入分配不是监管主责，但**经营净收入（16.7%）**与个体户经营环境直接相关。", regH: 1.3 });
    source(s, "国家统计局《2025年居民收入和消费支出情况》；浙江省统计局；资金流量表口径为作者概算。");
  }

  {
    const s = content(2, "能不能消费之二：今年上半年，收入跑输GDP，消费又跑输收入",
      "【约2分钟】三条线：GDP、人均收入、人均消费的实际增速。疫情三年波动大，2021、2023年消费高增长是低基数反弹。\n关键看最右边：今年上半年GDP 4.7%＞收入4.2%＞消费2.7%。收入名义增长5.2%，消费名义只增长3.7%——多出来的收入变成了储蓄。\n要说公道话：2017、2023—2025年收入并不慢于GDP，今年的“两级落差”是新变化，值得警惕。还要注意口径：人均收入宜和人均GDP比较；人口负增长时，真实落差比图上略大。\n这说明问题不只在“钱”，还在“心”——自然引出第二道门槛。", "s_incline");
    chartCap(s, "图8  GDP、居民收入与消费的实际增速（%）");
    s.addChart(pres.charts.LINE, [
      { name: "GDP", labels: D.incYears, values: D.incGDP },
      { name: "人均可支配收入", labels: D.incYears, values: D.incReal },
      { name: "人均消费支出", labels: D.incYears, values: D.conReal },
    ], chartOpts({
      x: MX, y: 1.95, w: 7.95, h: 4.6, chartColors: [HX.blue, HX.teal, HX.red], lineSize: 2.25, lineDataSymbol: "circle", lineDataSymbolSize: 6,
      showLegend: true, valAxisMinVal: -6, valAxisMaxVal: 14, valAxisMajorUnit: 2, catAxisLabelFontSize: 11, valAxisLabelFormatCode: "0",
    }));
    insight(s, [
      "2026年上半年：GDP **4.7%** ＞ 收入 **4.2%** ＞ 消费 **2.7%**",
      "收入名义+5.2%，消费名义+3.7%：**多出来的收入变成了储蓄**",
      "2017、2023—2025年收入并不慢于GDP：今年的落差是**新变化**",
    ], { reg: "收入转化为消费时，**服务消费、预付式消费**最依赖信任——监管能减少这段“传导损耗”。", regH: 1.5 });
    source(s, "国家统计局。2026年为上半年同比；人均收入宜与人均GDP比较，人口负增长下真实落差略大于图示。");
  }

  {
    const s = content(2, "敢不敢消费之一：消费倾向未回到疫情前，家庭在“加固资产负债表”",
      "【约3分钟】左图是平均消费倾向：每100元可支配收入花掉多少。2019年70.1元，2025年68.0元，没有回到疫情前。差2个百分点看似不多，换算一下：如果回到2019年水平，2025年居民消费可以多出约1.3万亿元，相当于GDP的0.9%。这多出来的两个百分点，就是每户人家抽屉里的一份“防身钱”。\n右图是住户存款每年新增额：2019—2021年平均约10万亿元，2022年以后约15万亿元。但要实事求是——存款多增不等于储蓄多增：少买房，首付款就留在了住户部门；理财赎回、利息滚存也会推高存款。\n方向是清楚的：收入不确定、住房和教育医疗负担、养老保障缺口、人口老龄化，都会推高储蓄率——这些主要是宏观与民生政策的主场。\n【最新】据媒体解读人民银行10月14日数据，今年前三季度住户存款增加约9.39万亿元，比上年同期少增约3.3万亿元，住户贷款减少约1.03万亿元——“存款搬家”与居民去杠杆同时在发生（讲前对照央行原文）。", "s_apc");
    chartCap(s, "图9  平均消费倾向（%）", MX, 1.55, 3.95);
    s.addChart(pres.charts.LINE, [{ name: "平均消费倾向", labels: D.apcYears.map((y, i) => y + "\n" + D.apc[i].toFixed(1)), values: D.apc }], chartOpts({
      x: MX, y: 1.95, w: 3.95, h: 4.6, chartColors: [HX.teal], lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 8,
      showValue: false, valAxisMinVal: 60, valAxisMaxVal: 72, valAxisMajorUnit: 4, catAxisLabelFontSize: 10,
    }));
    chartCap(s, "图10  住户存款年度新增（万亿元）", 4.65, 1.55, 4.0);
    s.addChart(pres.charts.BAR, [{ name: "住户存款新增", labels: D.depYears, values: D.dep }], chartOpts({
      x: 4.65, y: 1.95, w: 3.95, h: 4.6, barDir: "col", chartColors: D.dep.map((v, i) => (i >= 3 ? HX.red : HX.context)),
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", valAxisMinVal: 0, valAxisMaxVal: 20, valAxisMajorUnit: 5, catAxisLabelFontSize: 11,
    }));
    insight(s, [
      "若消费倾向回到2019年的70.1%，2025年可多出约**1.3万亿元**消费（≈GDP的0.9%）",
      "2022—2025年住户存款年均新增约**15.9万亿元**，比2019—2021年均多约5.5万亿",
      "存款多增≠储蓄多增：**少买房**使首付款留在住户部门，理财赎回也会推高存款",
      "收入不确定、住房与教育医疗负担、**人口老龄化**都会推高储蓄率",
    ]);
    source(s, "国家统计局（人均消费支出/人均可支配收入，作者计算；纵轴截断）；中国人民银行金融统计数据。");
  }

  {
    const s = content(2, "敢不敢消费之二：房地产调整从“财富、负债、关联消费”三头压制消费",
      "【约2.5分钟】新房销售面积从2021年的17.9亿平方米降到2025年的8.8亿平方米，几乎腰斩；今年1—8月又降了12.1%。\n房地产通过三个渠道影响消费：一是财富——城镇家庭资产约六成在住房上，房价预期转弱，家庭“感觉变穷”；二是负债——提前还贷、主动去杠杆，挤占当期消费；三是关联消费——装修、家电、家具、建材跟着走弱，还拖累地方财政。\n讲公道话：房价上涨时期“为买房而储蓄”也曾挤出消费，所以净效应因家庭而异。\n【监管含义】调整期中介、装修、物业、预售等领域的消费纠纷容易增多，要提前研判。", "s_house");
    chartCap(s, "图11  全国新建商品房销售面积（亿平方米）");
    s.addChart(pres.charts.BAR, [{ name: "销售面积", labels: D.houseYears, values: D.house }], chartOpts({
      x: MX, y: 1.95, w: 7.95, h: 4.6, barDir: "col", chartColors: D.house.map((v, i) => (i === 2 ? HX.teal : i >= 3 ? HX.red : HX.context)),
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", valAxisMinVal: 0, valAxisMaxVal: 20, valAxisMajorUnit: 5,
    }));
    insight(s, [
      "新房销售面积较2021年峰值下降约**51%**；2026年1—8月再降**12.1%**",
      "2025年开发投资**8.28万亿元**（降17.2%）⏎2026年1—8月再降19.9%",
      "城镇家庭资产约**六成**在住房：财富缩水、提前还贷去杠杆→多存少花",
    ], { reg: "调整期**中介、装修、物业、预售**等领域消费纠纷易增多，需提前研判。" });
    source(s, "国家统计局；家庭资产结构参见中国人民银行《2019年中国城镇居民家庭资产负债情况调查》（住房占59.1%）。");
  }

  {
    const s = content(2, "敢不敢消费之三：就业与保障决定“安全感”",
      "【约2.5分钟】就业：8月是毕业季的季节性高点，16—24岁（不含在校生）失业率18.9%，与上年同期持平；25—29岁7.5%。今年高校毕业生1270万人，再创新高。我们统一说：结构性就业矛盾仍然突出。\n保障：城乡居民基础养老金最低标准今年再提高20元，到每月163元，和职工养老金差距仍然悬殊。养老金偏低，让中国老年人的储蓄率也偏高——这是“敢消费”最深的结构性根源之一，叫人口老龄化下的预防性储蓄。\n政策正在转向“投资于人”：“十五五”规划建议提出实施城乡居民增收计划；2025年起推开失能老年人养老服务消费补贴。\n【与监管的连接】经营主体活跃是吸纳就业的主渠道——小店兴旺，青年和灵活就业才有“饭碗”。", "s_jobs");
    const tiles = [
      ["18.9%", "16—24岁失业率", "2026年8月，不含在校生⏎毕业季季节性高点", K.red],
      ["7.5%", "25—29岁失业率", "2026年8月，不含在校生", K.red],
      ["1270万", "2026届高校毕业生", "规模再创新高⏎同比增加48万", K.blue],
      ["163元", "城乡居民基础养老金", "2026年月最低标准⏎再提高20元", K.gold],
    ];
    const st = stepOf(4, 2.85);
    tiles.forEach((t, i) => {
      const x = MX + i * st;
      box(s, x, 1.65, 2.85, 2.75, K.tint);
      T(s, t[0], { x: x + 0.25, y: 1.8, w: 2.45, h: 0.9, fontSize: 36, bold: true, color: t[3], valign: "bottom" });
      T(s, t[1], { x: x + 0.25, y: 2.8, w: 2.45, h: 0.4, fontSize: 16, bold: true });
      T(s, t[2], { x: x + 0.25, y: 3.3, w: 2.45, h: 0.95, fontSize: 14, color: K.muted, lineSpacingMultiple: 1.1 });
    });
    box(s, MX, 4.65, CW, 1.9, HX.white, { line: { color: HX.line, width: 1 } });
    T(s, paras([
      "**预防性储蓄与老龄化**：未来收入越不确定、保障越薄弱，家庭越倾向“存钱防身”；养老金偏低，老年人储蓄率也偏高",
      "**政策转向“投资于人”**：“十五五”规划建议提出实施城乡居民增收计划；2025年起推开失能老年人养老服务消费补贴",
      "**与监管的连接**：经营主体活跃是吸纳就业的主渠道——小店兴旺，青年和灵活就业才有“饭碗”",
    ], { fontSize: 14 }, { space: 6 }), { x: MX + 0.3, y: 4.75, w: CW - 0.6, h: 1.7, valign: "middle", lineSpacingMultiple: 1.1 });
    source(s, "国家统计局；教育部；人力资源社会保障部；民政部、财政部；“十五五”规划建议。");
  }

  {
    const s = content(2, "愿不愿消费：服务消费占比停在46%，缺的是“对路”又“放心”的供给",
      "【约2.5分钟】收入提高后，消费会从“买东西”转向“买服务”。我国服务性消费占比2013年39.7%，2019年45.9%，2024、2025年都是46.1%，基本停滞。2025年增长最快的是教育文化娱乐，增长9.4%；医疗保健只增长1%。\n为什么停滞？不是简单的“床位不够”——全国养老机构入住率长期不足一半。缺的是“对路”的供给：信不过、付不起、位置不对。\n这里有个概念：餐饮、旅游多是“经验品”，用过才知道好坏；医疗、养老、维修、培训更是“信任品”，用过了也难判断。越是信任品，越需要标准、认证和信用评价——这正是市场监管的用武之地。", "s_svc");
    chartCap(s, "图12  居民人均服务性消费支出占人均消费支出比重（%）");
    s.addChart(pres.charts.LINE, [{ name: "服务消费占比", labels: D.svcYears, values: D.svc }], chartOpts({
      x: MX, y: 1.95, w: 7.95, h: 4.6, chartColors: [HX.teal], lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 9,
      showValue: true, dataLabelPosition: "t", dataLabelFormatCode: "0.0", dataLabelFontSize: 13, valAxisMinVal: 40, valAxisMaxVal: 48, valAxisMajorUnit: 2,
    }));
    insight(s, [
      "2025年人均服务性消费**13,602元**，占比**46.1%**，与上年持平（2013年为39.7%）",
      "增长最快的是**教育文化娱乐（+9.4%）**；医疗保健仅+1.0%",
      "养老机构入住率**不足一半**：缺的不是床位，而是信得过、付得起、对路的服务",
    ], { reg: "医疗、养老、维修、培训是“**信任品**”：标准、认证、信用评价越健全，服务消费越能放开。", regH: 1.5 });
    source(s, "国家统计局（城乡一体化住户调查口径；纵轴截断）；养老机构入住率参见中国社会科学网相关研究。");
  }

  {
    const s = content(2, "放不放心消费：投诉看结构不看总量——网购、售后、合同是信任短板",
      "【约3分钟】2025年全国市场监管部门受理投诉、举报、咨询4386.6万件，其中投诉举报2646万件，同比增长9.8%，为消费者挽回经济损失43.5亿元。\n总量要辩证看：投诉多可能是渠道更畅通了，也会随交易量同步增长。看结构更有意义：网上商品零售大约占社零的四分之一，网购却占了投诉举报的56.9%；售后服务投诉连续三年居首；合同类投诉增长40.3%，集中在退款难、霸王条款。\n讲一个经典理论：阿克洛夫的“柠檬市场”。买方辨不清质量时，好产品卖不出好价。预付卡就是今天最典型的“柠檬”：消费者分不清哪家健身房会跑路，干脆谁家的卡都不办，守规矩的店也跟着吃亏。信任缺失不一定让消费消失，却会让消费“降级”或“外流”。\n贯穿案例：空调买回家，发现能效虚标、安装售后扯皮——下次还敢用补贴买吗？", "s_trust");
    chartCap(s, "图13  2025年12315受理构成（万件）", MX, 1.55, 4.6);
    s.addChart(pres.charts.DOUGHNUT, [{ name: "受理构成", labels: ["投诉", "咨询", "举报"], values: [2036.6, 1740.6, 609.4] }], chartOpts({
      x: MX, y: 1.95, w: 4.6, h: 4.6, chartColors: [HX.red, HX.teal, HX.blue], holeSize: 55, showValue: true, dataLabelFormatCode: "0.0",
      dataLabelColor: HX.white, dataLabelFontBold: true, dataLabelFontSize: 12, showLegend: true, legendPos: "b", legendFontSize: 13, dataBorder: { pt: 2, color: "FFFFFF" },
    }));
    T(s, [{ text: "4386.6", options: { fontSize: 22, bold: true, color: K.jade, breakLine: true } }, { text: "万件", options: { fontSize: 13, color: K.muted } }],
      { x: MX + 1.5, y: 3.4, w: 1.6, h: 0.9, align: "center", valign: "middle" });
    const kv = [["56.9%", "网购占投诉举报⏎（网上零售约占社零1/4）", K.red], ["+40.3%", "合同类投诉⏎（退款难、霸王条款）", K.red], ["2646万件", "投诉举报，同比+9.8%", K.ink], ["43.5亿元", "为消费者挽回经济损失", K.teal]];
    kv.forEach((k, i) => {
      const y = 1.65 + i * 1.22;
      T(s, k[0], { x: 5.45, y, w: 3.2, h: 0.55, fontSize: 26, bold: true, color: k[2], valign: "bottom" });
      T(s, k[1], { x: 5.45, y: y + 0.57, w: 3.2, h: 0.6, fontSize: 13, color: K.muted });
    });
    insight(s, [
      "售后服务投诉**连续三年居首**；智能设备投诉**+26.6%**",
      "**柠檬市场**：买方辨不清质量，好产品卖不出好价——信任缺失会让消费**降级或外流**",
      "信任缺失可以理解为一笔“**隐性税**”",
    ], { head: "看点 · 经济学", reg: "看投诉要看**重复投诉、同类投诉**是否下降，而不只是总量。" });
    source(s, "国家市场监督管理总局2025年全国消费者投诉举报数据（2026年3月15日发布）；国家统计局。");
  }

  {
    const s = content(2, "小结：市场监管不唱主角，却守住政策效能的“最后一公里”",
      "【约2分钟，第二部分总结】开场投票时，多数同志选了收入和房子——这和数据一致，是消费不足的主因，主要靠收入分配、社保、房地产、就业等宏观与民生政策。市场监管不是主角，不能“包打天下”，也不能夸大自己的作用。\n但选“不放心”的同志提醒了我们：这恰恰是我们能直接动手的那一项。补贴发下去了，如果有人先提价后打折，补贴就被截留了；收入增加了，如果买东西不放心、维权太难，钱还是会存进银行。监管守住的是政策效能的最后一公里。\n【过渡】上半场到这里。休息10分钟。", "s_sum");
    const rows = [
      ["门槛", "主要症结", "关键工具", "主责"],
      ["能消费", "居民收入份额偏低、增速放缓、分配不均", "增收计划、提低扩中、财产性收入", "宏观·财政·人社"],
      ["敢消费", "房地产调整、就业与保障不确定、老龄化", "稳楼市、稳就业、提高养老金与医保", "宏观·住建·人社"],
      ["愿消费", "服务供给不对路、标准缺位、不合理限制", "放宽准入、标准引领、清理不合理限制", "多部门共管"],
      ["放心消费", "质量隐患、价格乱象、维权难、内卷式竞争", "质量安全、公平竞争、消费维权、信用", "市场监管主责"],
    ];
    const tbl = rows.map((r, ri) => r.map((c, ci) => ({
      text: c,
      options: {
        bold: ri === 0 || ci === 0 || (ci === 3 && ri === 4), fontSize: 15, fontFace: FONT,
        color: ri === 0 ? HX.white : ci === 3 && ri === 4 ? HX.red : HX.ink,
        fill: { color: ri === 0 ? HX.jade : ri === 4 ? HX.redTint : ri === 3 ? "FBF4E4" : ri % 2 ? HX.tint : HX.white },
        valign: "middle", margin: [0.05, 0.12, 0.05, 0.12],
      },
    })));
    s.addTable(tbl, { x: MX, y: 1.65, w: CW, colW: [1.6, 4.25, 4.13, 2.15], rowH: [0.5, 0.75, 0.75, 0.75, 0.75], border: { type: "solid", pt: 1, color: "FFFFFF" } });
    box(s, MX, 5.45, CW, 1.05, K.jade);
    T(s, [
      { text: "“最后一公里”  ", options: { bold: true, color: K.goldLight, fontSize: 16 } },
      { text: "增收、补贴要转化为真实消费，最后都要经过“买得放心”这一关——这一关的把关人是市场监管。", options: { color: K.white, fontSize: 16 } },
    ], { x: MX + 0.35, y: 5.45, w: CW - 0.7, h: 1.05, valign: "middle" });
  }

  {
    const s = pres.addSlide({ masterName: "DARK", sectionTitle: CURRENT_SECTION });
    T(s, "中场休息 · 10分钟", { x: 1.0, y: 1.9, w: 11.3, h: 1.1, fontSize: 50, bold: true, color: K.white, align: "center", valign: "middle" });
    T(s, "休息时想一想：3·15为什么是3月15日？", { x: 1.0, y: 3.3, w: 11.3, h: 0.55, fontSize: 24, color: K.goldLight, align: "center" });
    T(s, "请把您最想问的一个问题写在便签上，交给工作人员", { x: 1.0, y: 4.1, w: 11.3, h: 0.45, fontSize: 18, color: HX.darkText, align: "center" });
    T(s, "下半场：别国的经验与教训 ｜ 中央怎么部署 ｜ 市场监管怎么干", { x: 1.0, y: 4.8, w: 11.3, h: 0.45, fontSize: 18, color: HX.darkText, align: "center" });
    s.addNotes("【休息10分钟】下半场开头先揭晓3·15的由来，再用1分钟回答便签上的1—2个问题。");
  }

  // ===================== 第三部分 镜鉴 =====================
  section("三、镜鉴");
  divider(3, "他山之石：消费大国的制度底座", "高消费率靠“国民性”吗？别国的经验与教训。",
    "【约30秒】先揭晓：1962年3月15日，肯尼迪向国会提交保护消费者利益的特别咨文，提出安全、知情、选择、被倾听四项权利——3·15就是这么来的。（再用1分钟回答便签上的问题）");

  {
    const s = content(3, "美欧镜鉴：消费者保护决定不了消费率，却决定新消费能否“放得开”",
      "【约3.5分钟】美国家庭消费率高，主要靠收入分配、社保和信贷体系，不能简单归功于消费者保护——印度的消费率就比德国高。但美欧百年消费者保护史说明了另一件事：新产品、新服务、新渠道之所以能被大规模接受，是因为把试错成本从消费者转给了经营者。\n讲个故事：1906年，辛克莱的小说《屠场》写芝加哥肉类加工厂的卫生黑幕，引发舆论，同年美国通过了《纯净食品药品法》。1962年3月15日，肯尼迪提出消费者四项权利。欧盟有14天无理由撤回权，对应我们的网购七日无理由退货。\n中国例子：2008年三聚氰胺事件后，国产婴幼儿奶粉的信任修复用了十多年，大量需求转向进口和代购——信任缺失未必让消费消失，却会让消费外流。\n我们的制度建设起步晚、进步快：1993年消保法，2013年修订，2018年组建总局，2024年消保法实施条例和公平竞争审查条例施行，2025年新修订的反不正当竞争法施行。\n【三条规律】消费率是慢变量；信任决定新消费能否放得开；竞争政策保护竞争而非竞争者。", "s_useu");
    const top = [["1906", "美《纯净食品药品法》"], ["1914", "美联邦贸易委员会成立"], ["1962", "肯尼迪“消费者四权”"], ["1972", "美消费品安全委员会"], ["1985", "欧《产品责任指令》"], ["1993", "欧盟统一大市场"], ["2011", "欧盟14天无理由撤回"], ["2022", "欧盟《数字市场法》"]];
    const bot = [["1993", "《消费者权益保护法》"], ["2013", "消保法修订：网购七日无理由退货"], ["2018", "组建市场监管总局"], ["2024", "消保法实施条例施行"], ["2024", "公平竞争审查条例施行"], ["2025", "新修订反不正当竞争法施行"], ["2025", "五部门优化消费环境三年行动"]];
    const row = (items, ax, col, label) => {
      const x0 = 1.85, span = W - MX - x0, step = span / items.length;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: ax - 0.95, w: 1.05, h: 1.9, fill: { color: col }, line: { color: col, width: 0 }, rectRadius: 0.08 });
      T(s, label, { x: MX, y: ax - 0.95, w: 1.05, h: 1.9, fontSize: 20, bold: true, color: K.white, align: "center", valign: "middle" });
      s.addShape(pres.shapes.LINE, { x: x0, y: ax, w: span, h: 0, line: { color: HX.line, width: 1.5 } });
      items.forEach((t, i) => {
        const cx = x0 + step * i + step / 2;
        s.addShape(pres.shapes.OVAL, { x: cx - 0.09, y: ax - 0.09, w: 0.18, h: 0.18, fill: { color: col }, line: { color: K.white, width: 1.5 } });
        T(s, t[0], { x: cx - step / 2 + 0.03, y: ax - 0.55, w: step - 0.06, h: 0.38, fontSize: 15, bold: true, color: col, align: "center", valign: "bottom" });
        T(s, t[1], { x: cx - step / 2 + 0.03, y: ax + 0.17, w: step - 0.06, h: 0.95, fontSize: 14, align: "center", lineSpacingMultiple: 1.0 });
      });
    };
    row(top, 2.55, K.blue, "美欧");
    row(bot, 4.6, K.red, "中国");
    box(s, MX, 5.9, CW, 0.68, K.tint);
    T(s, [{ text: "核心逻辑  ", options: { bold: true, color: K.teal } }, { text: "消费者保护把“试错成本”从消费者转给经营者——新产品、新服务、新渠道才能被大规模接受。", options: {} }],
      { x: MX + 0.3, y: 5.9, w: CW - 0.6, h: 0.68, fontSize: 15, valign: "middle" });
    source(s, "各国立法与机构公开资料，作者整理。两条时间轴各按时间顺序排列，非同一比例。");
  }

  {
    const s = content(3, "日本教训：价格预期一旦转弱，扭转代价极高",
      "【约3.5分钟】先说分寸：中日在发展阶段、城镇化空间、政策工具和产业升级上差异很大，讲日本是为了“防”，不是说我们会“成为”日本。\n1990年前后资产泡沫破裂，企业部门“还债优先、不借钱”，这是辜朝明说的资产负债表衰退。之后近二十年物价基本零增长，红色柱是负增长年份。注意1997年和2014年的跳升是消费税上调造成的，不是需求回暖；日本物价真正持续回到2%左右要到2022年以后。\n一个故事：2000年代日本牛肉盖饭连锁店打价格战，一碗降到两三百日元，门店越开越多，店员工资越压越低——这就是“低价—低薪—低需求”的样子。\n要讲准确：这一时期日本居民消费占比不降反升，家庭储蓄率降到个位数。日本的教训是“名义增长停滞+通缩预期”，不是“消费率低”。我们的资产负债表压力主要在家庭和地方政府，和日本当年的企业部门不同。", "s_japan");
    chartCap(s, "图14  日本CPI同比涨幅（%），1990—2015年");
    s.addChart(pres.charts.BAR, [{ name: "日本CPI", labels: D.jpYears.map((y) => (+y % 5 === 0 ? y : "")), values: D.jpCPI }], chartOpts({
      x: MX, y: 1.95, w: 7.95, h: 4.6, barDir: "col", chartColors: D.jpCPI.map((v) => (v < 0 ? HX.red : HX.context)), invertedColors: D.jpCPI.map((v) => (v < 0 ? HX.red : HX.context)),
      valAxisMinVal: -2, valAxisMaxVal: 4, valAxisMajorUnit: 1, catAxisLabelFontSize: 12, barGapWidthPct: 35,
    }));
    T(s, "消费税上调", { x: 2.75, y: 2.85, w: 1.2, h: 0.3, fontSize: 11, color: K.muted, align: "center" });
    T(s, "消费税上调", { x: 6.85, y: 2.2, w: 1.2, h: 0.3, fontSize: 11, color: K.muted, align: "center" });
    insight(s, [
      "**先说分寸**：中日发展阶段、城镇化空间、政策工具差异显著——借鉴教训，不作简单类比",
      "泡沫破裂后企业“还债优先”；降价压薪保份额：**低价→低薪→低需求**",
      "这一时期日本居民消费占比**不降反升**：教训是“名义增长停滞+通缩预期”，而非“消费率低”",
    ], { reg: "警惕“内卷式”价格战演变为全行业**降质、降薪、降创新**。" });
    source(s, "日本总务省统计局（全国CPI综合指数，年度同比，取一位小数）。1997、2014年跳升主要源于消费税上调。");
  }

  // ===================== 第四部分 谋策 =====================
  section("四、谋策");
  divider(4, "从“全方位扩内需”到“坚持内需主导”", "中央在怎么部署？市场监管在政策体系中的“坐标”在哪里？",
    "【过渡，约30秒】第四部分，看中央部署，并用经济学的眼光审视以旧换新这项最主要的促消费政策。");

  {
    const s = content(4, "政策脉络：从“全方位扩大国内需求”到“坚持内需主导”",
      "【约3分钟，不逐条念】2024年中央经济工作会议提出“大力提振消费、提高投资效益，全方位扩大国内需求”，2025年中央经济工作会议指出“国内供强需弱矛盾突出”，提出“坚持内需主导”。中间的关键文件依次是：优化消费环境三年行动方案、提振消费专项行动方案、中央财经委治理低价无序竞争、“十五五”规划建议、政府工作报告、扩大消费“十五五”规划。\n红色节点是市场监管牵头或深度参与的：在提振消费这件大事里，市场监管有“专属任务”。\n一个对比：扩大消费“十五五”规划提出2030年社零60万亿元左右，需要年均名义增长约3.7%，而今年1—8月只增长1.1%——这个缺口靠补贴补不上，要靠增收、社保和消费环境的制度建设。", "s_policy");
    const ev = [
      ["2024.12", "中央经济工作会议", "全方位扩大国内需求⏎提振消费列首要任务", 0],
      ["2025.02", "五部门三年行动方案", "优化消费环境⏎总局牵头", 1],
      ["2025.03", "提振消费专项行动方案", "中办、国办印发", 0],
      ["2025.07", "中央财经委第六次会议", "治理企业⏎低价无序竞争", 1],
      ["2025.10", "“十五五”规划建议", "居民消费率明显提高⏎实施城乡居民增收计划", 0],
      ["2025.12", "中央经济工作会议", "供强需弱矛盾突出⏎坚持内需主导", 0],
      ["2026.03", "政府工作报告", "目标4.5%—5%⏎以旧换新2500亿元", 0],
      ["2026.07", "扩大消费“十五五”规划", "2030年社零约60万亿⏎需年均增长约3.7%", 0],
      ["2026.09", "总局放心消费培育通知", "2030年集聚区⏎5000个以上", 1],
    ];
    const ax = 3.75, step = CW / ev.length, bw = 2.1;
    s.addShape(pres.shapes.LINE, { x: MX, y: ax, w: CW, h: 0, line: { color: HX.jade, width: 2 } });
    ev.forEach((e, i) => {
      const cx = MX + step * i + step / 2, up = i % 2 === 0, hl = e[3] === 1;
      const col = hl ? K.red : K.teal;
      s.addShape(pres.shapes.OVAL, { x: cx - 0.11, y: ax - 0.11, w: 0.22, h: 0.22, fill: { color: col }, line: { color: K.white, width: 1.5 } });
      T(s, e[0], { x: cx - 0.8, y: up ? ax - 0.5 : ax + 0.17, w: 1.6, h: 0.33, fontSize: 14, bold: true, color: col, align: "center" });
      const by = up ? 1.6 : 4.4;
      const bx = Math.min(Math.max(cx - bw / 2, MX), W - MX - bw);
      s.addShape(pres.shapes.LINE, { x: cx, y: up ? by + 1.35 : ax + 0.5, w: 0, h: up ? ax - 0.5 - by - 1.35 : by - ax - 0.5, line: { color: HX.line, width: 1 } });
      box(s, bx, by, bw, 1.35, hl ? HX.redTint : K.tint);
      T(s, e[1], { x: bx + 0.08, y: by + 0.1, w: bw - 0.16, h: 0.42, fontSize: 13, bold: true, align: "center", valign: "middle" });
      T(s, runs(e[2], { fontSize: 13, color: K.ink }), { x: bx + 0.08, y: by + 0.55, w: bw - 0.16, h: 0.72, fontSize: 13, align: "center", valign: "top", lineSpacingMultiple: 1.05 });
    });
    box(s, MX, 5.95, CW, 0.62, K.jade);
    T(s, [
      { text: "三个结合  ", options: { bold: true, color: K.goldLight } },
      { text: "投资于物与投资于人紧密结合 ｜ 商品消费与服务消费并重 ｜ 政策刺激与制度建设并举", options: { bold: true, color: K.white } },
    ], { x: MX + 0.3, y: 5.95, w: CW - 0.6, h: 0.62, fontSize: 15, align: "center", valign: "middle" });
    source(s, "中央经济工作会议、政府工作报告及相关文件公开报道，作者整理。红色节点为市场监管牵头或深度参与。");
  }

  {
    const s = content(4, "以旧换新：撬动倍数不低，但耐用品“需求前置”后的回落已经显现",
      "【约3分钟】中央资金2024年1500亿元、2025年3000亿元、2026年2500亿元，另设1000亿元财政金融协同促内需专项资金；今年家电补贴比例从20%降到15%、单件上限从2000元降到1500元——政策设计本身就会影响补贴归宿。2025年带动相关商品销售额约2.6万亿元，大约是中央资金的8—9倍；汽车以旧换新超过1150万辆，家电超过1.29亿件。\n右图是8月的分化：受补贴的数码产品还在高增长，汽车已经明显回落——这就是耐用品需求前置后的“回吐”，叠加去年的高基数。\n要强调一个概念区别：“带动销售额”不等于“新增消费”。2.6万亿元里，有本来就会买的、有提前买的、还有挤出其他消费的。经济学研究补贴，最关心的是净效应——下一页讲。", "s_tradein");
    chartCap(s, "图15  以旧换新中央资金（亿元）", MX, 1.55, 3.95);
    s.addChart(pres.charts.BAR, [{ name: "资金", labels: ["2024", "2025", "2026"], values: [1500, 3000, 2500] }], chartOpts({
      x: MX, y: 1.95, w: 3.95, h: 4.6, barDir: "col", chartColors: [HX.gold], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0",
      valAxisMinVal: 0, valAxisMaxVal: 3500, valAxisMajorUnit: 1000, barGapWidthPct: 60,
    }));
    chartCap(s, "图16  2026年8月零售额同比（%）", 4.65, 1.55, 4.0);
    const augL = D.aug.map((a) => a[0]), augV = D.aug.map((a) => a[1]);
    const augC = augV.map((v, i) => (v < 0 ? HX.red : i === augV.length - 1 ? HX.teal : HX.context));
    s.addChart(pres.charts.BAR, [{ name: "同比", labels: augL, values: augV }], chartOpts({
      x: 4.65, y: 1.95, w: 3.95, h: 4.6, barDir: "bar", chartColors: augC, invertedColors: augC, showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0",
      valAxisMinVal: -30, valAxisMaxVal: 40, valAxisMajorUnit: 10, catAxisLabelFontSize: 13, barGapWidthPct: 50, catAxisLabelPos: "low",
    }));
    insight(s, [
      "2025年带动销售额约**2.6万亿元**，约为中央资金的8—9倍；2026年1—8月1.55万亿元",
      "汽车以旧换新：⏎2025年超**1150万辆**⏎2026年1—8月**535万辆**",
      "“带动销售额”≠“新增消费”：要扣除**本来就会买、提前购买、挤出效应**",
      "2026年另设**1000亿元**财政金融协同促内需专项资金",
    ], { reg: "三项核查贯穿补贴全程：**比价、能效标识、旧机回收**。", regH: 1.0 });
    source(s, "国家发展改革委、商务部、财政部公开数据；国家统计局。");
  }

  {
    const s = content(4, "补贴的“归宿”：一元钱补贴，撬动了多少、落到了谁手里",
      "【约3分钟】这是我自己的研究领域：消费刺激政策的经济学。政府花一元钱补贴消费，要用两把尺子量。\n第一把，撬动了多少——支出响应。带动销售额不等于净新增：本来就会买的那部分，补贴只是转移支付，经济学叫“边际内补贴”，不是浪费但也没有撬动；提前购买是跨期替代；还可能挤出其他消费。\n第二把，落到谁手里——补贴归宿。它和税负归宿是同一套分析：谁的弹性小，谁分得多。家电、汽车产能充裕、竞争充分，补贴大部分应落到消费者；如果观察到明显涨价，先查是否虚构原价、是否串通，而不是把所有涨价都当违法。\n两把尺子之外，还要算社会福利：财政成本、分配公平、节能减排等外部性。\n监管守住的，是第二把尺子里的“违法截留”和“垄断截留”。下一页用一台空调算一笔账。", "s_incid");
    const L = MX, w2 = 3.6;
    box(s, L, 1.65, w2, 2.35, K.blue);
    T(s, "① 撬动了多少：支出响应", { x: L + 0.2, y: 1.78, w: w2 - 0.4, h: 0.4, fontSize: 17, bold: true, color: K.white });
    T(s, "带动销售额≠净新增。要扣除：本来就会买（边际内补贴，属转移支付）、提前购买（跨期替代）、挤出其他消费", { x: L + 0.2, y: 2.25, w: w2 - 0.4, h: 1.65, fontSize: 14, color: K.white, lineSpacingMultiple: 1.15 });
    box(s, L + 3.8, 1.65, w2, 2.35, K.red);
    T(s, "② 落到谁手里：补贴归宿", { x: L + 4.0, y: 1.78, w: w2 - 0.4, h: 0.4, fontSize: 17, bold: true, color: K.white });
    T(s, "由供需弹性与竞争程度决定：供给越有弹性、竞争越充分，越多落到消费者。另有违法截留：虚构原价、虚标能效、虚假回收骗补", { x: L + 4.0, y: 2.25, w: w2 - 0.4, h: 1.65, fontSize: 14, color: K.white, lineSpacingMultiple: 1.15 });
    box(s, L, 4.15, 7.4, 0.85, HX.white, { line: { color: HX.teal, width: 1.5 } });
    T(s, [{ text: "③ 社会福利  ", options: { bold: true, color: K.teal, fontSize: 16 } }, { text: "再计入财政成本、分配公平、节能减排等外部性", options: { fontSize: 15, color: K.ink } }],
      { x: L + 0.2, y: 4.15, w: 7.0, h: 0.85, valign: "middle" });
    T(s, "两把尺子：①主要看政策设计（补贴比例、上限、对象）；②主要看市场弹性与竞争。⏎监管守住的是②中的**违法截留**与**垄断截留**——不把所有涨价都当违法。",
      { x: L, y: 5.2, w: 7.4, h: 1.3, fontSize: 15, color: K.jade, bold: false, lineSpacingMultiple: 1.2 });
    box(s, 8.3, 1.65, 4.43, 4.85, HX.redTint);
    T(s, "监管能做什么", { x: 8.55, y: 1.82, w: 4.0, h: 0.42, fontSize: 17, bold: true, color: K.red });
    T(s, paras([
      "**价格监测**：补贴前后比价，识别虚构原价、先提价后打折",
      "**能效与认证核查**：防虚标等级、套取补贴",
      "**回收真实性核查**：防虚假回收、骗补",
      "**反垄断**：防串通涨价、限定交易",
      "**售后与投诉**：买得放心、修得方便",
    ], { fontSize: 14 }, { space: 9 }), { x: 8.55, y: 2.35, w: 4.0, h: 3.0, lineSpacingMultiple: 1.1 });
    T(s, "供给弹性有限时的价格上涨，是合法的市场结果；要查的是虚构与串通。", { x: 8.55, y: 5.45, w: 4.0, h: 0.9, fontSize: 14, bold: true, color: K.jade, lineSpacingMultiple: 1.15 });
    source(s, "作者研究框架：消费刺激的家庭支出响应、市场归宿与福利。");
  }

  {
    const s = content(4, "案例：一台空调的600元补贴，最后落进了谁的口袋（示意）",
      "【约2.5分钟，现场互动】先请大家心算：原价3000元的空调，补贴600元，商家补贴前先把标价提了300元——消费者实得多少？\n揭晓：实付2700元，比原价只省了300元。600元补贴里，消费者拿到一半——补贴归宿只有50%。要分清两种情况：如果提价是因为补贴拉动需求、供给一时跟不上，这是合法的市场归宿，监管不该管；如果是虚构原价、先提价后打折，就是价格欺诈，必须查。（今年家电补贴比例已从20%降到15%、单件上限从2000元降到1500元，示意数字按去年口径。）如果买家里有一半本来就会买，那么真正“撬动”的新增购买也只有一半。这就是两把尺子同时打折。\n这不是假设：2025年下半年，国家发展改革委会同商务部等部门集中整治了5类问题，其中就有骗补套补和经销商、平台“先涨后补”；2026年起，参与经营主体要履行价格备案、承诺和公示，违规的取消资格。\n法律依据：《明码标价和禁止价格欺诈规定》——不得虚构被比较价格，不得临时提高标价作为折扣基准（讲前核对具体条号）。部门分工：商务部门管资格与资金，可取消资格、追回补贴；市场监管管价格、能效标识、质量和售后，可查处并纳入失信。\n（如有本人研究中的价格传导率数字，在此补充一句。）", "s_aircon");
    // 手绘瀑布图
    const base = 6.05, k = 3.3 / 3300, bw = 1.25;
    const bars = [
      { x: 1.0, from: 0, to: 3000, col: HX.context, lab: "原标价", val: "3000" },
      { x: 2.65, from: 3000, to: 3300, col: HX.red, lab: "补贴前先提价", val: "+300" },
      { x: 4.3, from: 2700, to: 3300, col: HX.teal, lab: "国家补贴", val: "−600" },
      { x: 5.95, from: 0, to: 2700, col: HX.blue, lab: "消费者实付", val: "2700" },
    ];
    s.addShape(pres.shapes.LINE, { x: 0.8, y: base, w: 6.7, h: 0, line: { color: "9CA3AF", width: 1 } });
    bars.forEach((b, i) => {
      const yTop = base - b.to * k, h = (b.to - b.from) * k;
      s.addShape(pres.shapes.RECTANGLE, { x: b.x, y: yTop, w: bw, h, fill: { color: b.col }, line: { color: b.col, width: 0 } });
      T(s, b.val, { x: b.x - 0.1, y: yTop - 0.42, w: bw + 0.2, h: 0.38, fontSize: 16, bold: true, align: "center", valign: "bottom", color: i === 1 ? K.red : K.ink });
      T(s, b.lab, { x: b.x - 0.2, y: base + 0.08, w: bw + 0.4, h: 0.35, fontSize: 13, align: "center", color: K.ink });
      if (i < bars.length - 1) {
        const nextTop = i === 0 ? 3000 : i === 1 ? 3300 : 2700;
        s.addShape(pres.shapes.LINE, { x: b.x + bw, y: base - nextTop * k, w: 1.65 - bw, h: 0, line: { color: "9CA3AF", width: 1, dashType: "dash" } });
      }
    });
    box(s, 0.8, 1.5, 6.7, 0.78, K.tint);
    T(s, "消费者实得 **300元**（归宿50%），另300元归经营者⏎提价源于供需变化＝合法的市场归宿；虚构原价、先提价后打折＝价格欺诈", { x: 0.95, y: 1.5, w: 6.45, h: 0.78, fontSize: 13, valign: "middle", color: K.ink });
    box(s, 7.95, 1.6, 4.78, 4.95, HX.redTint);
    const blocks = [
      ["真实证据", "2025年下半年发改委会同商务部等集中整治5类问题，含骗补套补和经销商、平台“先涨后补”；2026年起参与主体须价格备案、承诺、公示，违规取消资格"],
      ["法律依据", "《明码标价和禁止价格欺诈规定》：不得虚构被比较价格，不得临时提高标价作为折扣基准"],
      ["部门分工", "商务部门：资格与资金（取消资格、追回补贴）⏎市场监管：价格、能效、质量与售后（查处、纳入失信）"],
    ];
    let yy = 1.78;
    blocks.forEach((b) => {
      T(s, b[0], { x: 8.2, y: yy, w: 4.3, h: 0.35, fontSize: 15, bold: true, color: K.red });
      T(s, runs(b[1], { fontSize: 13, color: K.ink }), { x: 8.2, y: yy + 0.38, w: 4.3, h: 1.1, fontSize: 13, lineSpacingMultiple: 1.1 });
      yy += 1.58;
    });
    source(s, "示意数字为作者构造；整治情况见国家发展改革委2025年12月新闻发布会。");
  }

  // ===================== 第五部分 作为 =====================
  section("五、作为");
  divider(5, "市场监管如何为消费“增信任”", "如何把“监管力”转化为“消费力”？",
    "【过渡，约30秒】最后一部分，回到我们自己。先用一张图看懂市场监管为什么管这些事，再讲与消费最相关的五个抓手。");

  {
    const s = content(5, "一张图看懂市场监管：管的是“市场失灵”，不是“市场”",
      "【约3.5分钟】各位每天的工作，从经济学看，是在纠正四类市场失灵。\n第一，信息不对称：消费者看不见质量，所以要有标准、计量、认证认可、检验检测、质量抽查、广告监管、消费维权——这些降低搜寻成本和维权成本。\n第二，安全风险与外部性：食品、特种设备、工业产品一旦出事，代价由社会承担，所以要守底线。\n第三，市场势力与行政性垄断：平台垄断、地方保护、违规补贴、内卷式竞争，靠反垄断、反不正当竞争、公平竞争审查、价格监督检查来“让价格说真话”。\n第四，交易成本与信用缺失：登记注册、信用监管、知识产权保护和综合执法，让进入更容易、守信更划算。\n2018年机构改革，把工商、质监、食药、物价监督检查、反垄断执法等并到一家，就是因为它们解决的是同一件事：让陌生人之间的交易可信——这也是全国统一大市场在执法层面的落点。\n今天只切其中与消费最相关的五刀。", "s_func");
    const rows = [
      ["市场失灵", "典型表现", "对应职能", "对应门槛"],
      ["信息不对称", "柠檬市场；“信任品”质量看不见", "标准、计量、认证认可、检验检测、质量抽查、广告监管、消费维权", "③ 愿消费 ④ 放心"],
      ["安全风险与外部性", "食品、特种设备、工业产品安全事故", "食品安全、特种设备安全、产品质量与缺陷召回", "④ 放心消费"],
      ["市场势力与行政性垄断", "平台垄断、地方保护、违规补贴、内卷式竞争", "反垄断、反不正当竞争、公平竞争审查、价格监督检查、网络交易监管", "① 能消费 ④ 放心"],
      ["交易成本与信用缺失", "进入难、经营成本高、失信成本低", "登记注册、信用监管、知识产权保护（商标、地理标志）、综合执法", "① 能消费 ③ 愿消费"],
    ];
    const tbl = rows.map((r, ri) => r.map((c, ci) => ({
      text: c,
      options: {
        bold: ri === 0 || ci === 0, fontSize: ri === 0 ? 15 : 14, fontFace: FONT,
        color: ri === 0 ? HX.white : ci === 0 ? HX.jade : ci === 3 ? HX.red : HX.ink,
        fill: { color: ri === 0 ? HX.jade : ri % 2 ? HX.tint : HX.white },
        valign: "middle", margin: [0.05, 0.12, 0.05, 0.12],
      },
    })));
    s.addTable(tbl, { x: MX, y: 1.6, w: CW, colW: [2.3, 3.05, 4.53, 2.25], rowH: [0.5, 0.88, 0.88, 0.88, 0.88], border: { type: "solid", pt: 1, color: "FFFFFF" } });
    box(s, MX, 5.8, CW, 0.75, K.jade);
    T(s, [
      { text: "一句话  ", options: { bold: true, color: K.goldLight } },
      { text: "市场失灵在哪里，监管职能就设在哪里——一个部门管“进入—经营—交易—退出”全链条，就是把规则统一起来。", options: { color: K.white } },
    ], { x: MX + 0.3, y: 5.8, w: CW - 0.6, h: 0.75, fontSize: 15, valign: "middle" });
  }

  {
    const s = content(5, "五个抓手：把“监管力”转化为“消费力”",
      "【约1.5分钟】从职能全景里切出与消费最相关的五块。注意右边那张灰卡：第二道门槛“敢消费”不是我们的主场，主要靠宏观和民生政策——知道边界，才有分寸。", "s_levers");
    const levers = [
      ["FaUserShield", "放心消费", "降低“信任成本”", "门槛④", K.red],
      ["FaBalanceScale", "公平竞争", "让价格说真话", "门槛①④", K.red],
      ["FaAward", "质量安全", "好产品出得来", "门槛③④", K.gold],
      ["FaStore", "经营主体", "活力→就业→收入", "门槛①", K.teal],
      ["FaMobileAlt", "新业态", "包容审慎促创新", "门槛③④", K.blue],
    ];
    const w = 1.85, st = stepOf(6, w);
    for (let i = 0; i < 5; i++) {
      const l = levers[i], x = MX + i * st;
      box(s, x, 1.7, w, 4.0, K.tint);
      await iconCircle(s, l[0], x + (w - 0.95) / 2, 1.95, 0.95, l[4]);
      T(s, `抓手${"一二三四五"[i]}`, { x, y: 3.05, w, h: 0.35, fontSize: 13, color: K.muted, align: "center" });
      T(s, l[1], { x, y: 3.4, w, h: 0.55, fontSize: 21, bold: true, align: "center", valign: "middle" });
      T(s, l[2], { x: x + 0.08, y: 4.0, w: w - 0.16, h: 0.65, fontSize: 14, align: "center" });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.25, y: 4.9, w: w - 0.5, h: 0.48, fill: { color: K.white }, line: { color: l[4], width: 1 }, rectRadius: 0.24 });
      T(s, l[3], { x: x + 0.25, y: 4.9, w: w - 0.5, h: 0.48, fontSize: 13, bold: true, color: l[4], align: "center", valign: "middle" });
    }
    const x6 = MX + 5 * st;
    box(s, x6, 1.7, w, 4.0, HX.white, { line: { color: "9CA3AF", width: 1.25, dashType: "dash" } });
    T(s, "门槛②", { x: x6, y: 2.0, w, h: 0.45, fontSize: 18, bold: true, color: K.muted, align: "center" });
    T(s, "敢消费", { x: x6, y: 2.5, w, h: 0.45, fontSize: 20, bold: true, color: K.muted, align: "center" });
    T(s, "主要靠宏观与⏎民生政策⏎⏎监管不越位", { x: x6 + 0.1, y: 3.2, w: w - 0.2, h: 1.6, fontSize: 14, color: K.muted, align: "center" });
    T(s, "五个抓手对应三道门槛；第二道门槛不是我们的主场——知道边界，才有分寸。", { x: MX, y: 6.0, w: CW, h: 0.45, fontSize: 16, bold: true, color: K.jade, align: "center" });
  }

  {
    const s = content(5, "抓手一 · 放心消费：降低“信任成本”，是财政花钱最少的长效政策",
      "【约3分钟】消费者买的不只是商品，还有“出了问题能解决”的保障。无理由退货降低“买错的成本”，预付式消费治理降低“跑路的风险”，在线纠纷解决降低维权的时间成本，信息公示激活“用脚投票”的声誉机制。\n为什么说是财政花钱最少的长效政策？发补贴要真金白银，建立信任主要靠规则和执法。但也要算清另一笔账：退货、存管、冷静期的合规成本落在经营者身上，尤其是小微商户。制度设计宁可规则简明、执法有力，也不要层层加码。\n预付式消费要注意职权边界：格式条款是我们的；资金存管要会同商务、教育、体育等行业主管部门推动。", "s_lever1");
    const st = [["5000个+", "2030年全国放心消费集聚区目标", K.blue], ["超1万家", "浙江2026年新增线下无理由退换货承诺单元", K.red], ["重复投诉率", "建议作为本地核心监测指标", K.teal]];
    st.forEach((t, i) => {
      const y = 1.65 + i * 1.68;
      box(s, MX, y, 4.0, 1.54, K.tint);
      T(s, t[0], { x: MX + 0.3, y: y + 0.15, w: 3.5, h: 0.75, fontSize: 30, bold: true, color: t[2], valign: "middle" });
      T(s, t[1], { x: MX + 0.3, y: y + 0.95, w: 3.5, h: 0.45, fontSize: 14, color: K.ink });
    });
    T(s, "四项重点工作", { x: 5.0, y: 1.65, w: 7.7, h: 0.4, fontSize: 17, bold: true, color: K.jade });
    const acts = [
      ["线下无理由退货", "从商场延伸到商圈、景区、夜市，降低“买错成本”"],
      ["预付式消费治理", "格式条款专项检查（依据最高法预付式消费司法解释）；会同行业主管部门推动资金存管、余额公示"],
      ["在线纠纷解决（ODR）", "引导企业与消费者直接和解，缩短维权链条"],
      ["信息公示与信用约束", "投诉数据公开、放心消费地图，激活“用脚投票”的声誉机制"],
    ];
    acts.forEach((a, i) => {
      const y = 2.15 + i * 0.97;
      T(s, String(i + 1), { x: 5.0, y, w: 0.5, h: 0.8, fontSize: 28, bold: true, color: K.red, valign: "middle" });
      T(s, [{ text: a[0], options: { bold: true, fontSize: 16, breakLine: true } }, { text: a[1], options: { fontSize: 14, color: K.muted } }], { x: 5.55, y, w: 7.15, h: 0.9, valign: "middle" });
    });
    box(s, 5.0, 6.0, 7.73, 0.55, K.jade);
    T(s, "消费者买的不只是商品，还有“出了问题能解决”的保障。", { x: 5.2, y: 6.0, w: 7.4, h: 0.55, fontSize: 15, bold: true, color: K.white, valign: "middle" });
    source(s, "国家市场监督管理总局《关于全面开展放心消费单元和集聚区培育的通知》（2026年9月）；浙江省市场监督管理局。");
  }

  {
    const s = content(5, "浙江实践：放心消费的“浙江样本”",
      "【约3分钟】上半年浙江GDP增长5.7%、社零增长3.0%，都高于全国。要讲公道：差异首先来自收入水平和市场活力，消费环境是加分项——相关不等于因果；而且浙江社零增速也从一季度的4.0%回落到了上半年的3.0%，我们同样面临压力。放心消费能否带来额外增量，要用投诉率、复购率、商圈客流做前后对比来证明，这是我们可以一起做的评估。\n几项有辨识度的工作：2017年起“放心消费在浙江”；“放心贷”把商户的“守信”变成“低息”；7个项目入选总局优化消费环境试点，全国第一；温州依托“浙食链”的全链条数字化监管入选国务院食安办2026年7月发布的典型案例；“放心贷”还为商户减免手续费约1.98亿元、节约利息约10.3亿元；“浙江外卖在线”2021年上线，平台证照核验累计拦截异常入驻4.2万家次。\n【三个追问，留给大家】无理由退货承诺单元超过1万家，消费者知道吗、用过吗？“放心贷”带动的是新增消费，还是替代了商户原有融资？7个试点，哪一个能复制到您所在的地市？\n【互动】请课前约好的衢州局同志用1—2分钟讲一个本地案例。", "s_zj");
    chartCap(s, "图17  2026年上半年：浙江与全国（同比，%）", MX, 1.55, 5.2);
    s.addChart(pres.charts.BAR, [
      { name: "全国", labels: ["GDP", "社零总额"], values: [4.7, 1.3] },
      { name: "浙江", labels: ["GDP", "社零总额"], values: [5.7, 3.0] },
    ], chartOpts({
      x: MX, y: 1.95, w: 5.2, h: 4.15, barDir: "col", barGrouping: "clustered", chartColors: [HX.context, HX.teal], showValue: true,
      dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", showLegend: true, valAxisMinVal: 0, valAxisMaxVal: 7, valAxisMajorUnit: 1, catAxisLabelFontSize: 14, barGapWidthPct: 70,
    }));
    T(s, "差异首先来自收入与市场活力，消费环境是加分项", { x: MX, y: 6.15, w: 5.2, h: 0.4, fontSize: 13, color: K.muted, align: "center" });
    const items = [
      ["起步", "2017年启动“放心消费在浙江”行动⏎首批1.3万家企业承诺无理由退货"],
      ["放心贷", "银行对承诺商户授信与利率优惠，“守信”变“低息”⏎覆盖147个商圈、4.16万户，放贷超1800亿元"],
      ["试点", "7个项目入选总局优化消费环境试点，居全国首位⏎“放心贷”入选全国五大典型案例"],
      ["浙食链", "温州食用农产品全链条数字化监管⏎入选国务院食安办全链条监管创新典型案例"],
      ["外卖在线", "2021年上线；平台证照核验⏎累计拦截异常入驻4.2万家次"],
      ["新目标", "培育“信用促消费”商圈100个以上⏎新增线下无理由退换货承诺单元超1万家"],
    ];
    items.forEach((it, i) => {
      const y = 1.6 + i * 0.84;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.2, y: y + 0.1, w: 1.2, h: 0.58, fill: { color: i === 1 ? K.red : K.teal }, line: { color: K.white, width: 0 }, rectRadius: 0.08 });
      T(s, it[0], { x: 6.2, y: y + 0.1, w: 1.2, h: 0.58, fontSize: 14, bold: true, color: K.white, align: "center", valign: "middle" });
      T(s, runs(it[1], { fontSize: 14, color: K.ink }), { x: 7.55, y, w: 5.18, h: 0.78, fontSize: 14, valign: "middle", lineSpacingMultiple: 1.05 });
    });
    source(s, "浙江省统计局；国家统计局；浙江省市场监督管理局及新华社、第一财经等公开报道。");
  }

  {
    const s = content(5, "现场判断：这四种“低价”，该不该管？",
      "【约2分钟，举手判断】每个情形请大家选：该管、不该管、看情况。\nA 外卖平台长期大额补贴，“0元奶茶”遍地。——看情况。以排除、限制竞争为目的，或低于成本、裹挟商家的，属于治理对象。背景：国务院反垄断反不正当竞争委员会办公室对外卖行业开展了竞争状况调查评估；6月17日总局发布《外卖平台补贴行为规范十条（征求意见稿）》（讲前核实是否已正式出台）。\nB 平台要求商家参加补贴、承担补贴成本，不参加就降权。——该管。新修订《反不正当竞争法》第十四条：平台不得强制或者变相强制平台内经营者按照其定价规则以低于成本的价格销售商品。\nC 行业协会倡议“本地某品类最低售价不低于X元”。——不该这么做。行业协会组织固定价格，涉嫌垄断协议（《反垄断法》关于行业协会的规定，讲前核对条号）。\nD 企业靠技术进步降价30%。——不该管。这是效率型降价，是竞争的本来面目。", "s_quiz");
    const qs = [
      ["A", "外卖平台长期大额补贴，“0元奶茶”遍地"],
      ["B", "平台要求商家参加补贴、承担补贴成本，不参加就降权"],
      ["C", "行业协会倡议“本地某品类最低售价不低于X元”"],
      ["D", "企业靠技术进步，产品降价30%"],
    ];
    const cw = (CW - 0.3) / 2;
    qs.forEach((q, i) => {
      const x = MX + (i % 2) * (cw + 0.3), y = 1.7 + Math.floor(i / 2) * 1.95;
      box(s, x, y, cw, 1.75, i % 3 === 0 ? K.tint : HX.gray);
      s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: y + 0.45, w: 0.85, h: 0.85, fill: { color: K.jade }, line: { color: K.jade, width: 0 } });
      T(s, q[0], { x: x + 0.3, y: y + 0.45, w: 0.85, h: 0.85, fontSize: 30, bold: true, color: K.white, align: "center", valign: "middle" });
      T(s, q[1], { x: x + 1.4, y: y + 0.2, w: cw - 1.65, h: 1.35, fontSize: 19, bold: true, valign: "middle", lineSpacingMultiple: 1.15 });
    });
    const opts = [["该管", K.red], ["不该管", K.teal], ["看情况", K.gold]];
    opts.forEach((o, i) => pill(s, 3.9 + i * 2.0, 5.8, 1.7, 0.6, o[0], o[1], K.white, 17));
  }

  {
    const s = content(5, "抓手二 · 公平竞争：从源头治理“内卷式”竞争，让价格说真话",
      "【约3分钟】三层意思。第一，价格监管“两线作战”：9月PPI上涨5.0%、CPI只涨0.9%，剪刀差从8月的3个百分点扩大到4.1个百分点。上游防囤积哄抬、串通涨价；下游治理低价倾销、价格欺诈。今年上半年总局部署16项专项行动，全国查办不正当竞争案件11465件，并在推动价格法修改，完善低价倾销认定和行业成本调查；上半年还审结经营者集中52起，支持依法并购整合。\n第二，从源头治：内卷的重要根源，是地方以补贴、税收返还、低价供地竞相招商。公平竞争审查由市场监管牵头，是约束这类行为的法定抓手。\n第三，划清边界，并坦率承认：上游涨价里也有“反内卷”带来的供给收缩，要防止从“治乱”滑向“限产保价”。判断标准只有一个：是否损害消费者长期福利和市场动态效率。监管的目标是让价格“说真话”，不是让价格“往上走”。", "s_comp");
    T(s, "价格监管“两线作战”", { x: MX, y: 1.6, w: 5.6, h: 0.4, fontSize: 17, bold: true, color: K.jade });
    const fronts = [
      ["上游", "PPI +5.0%（9月）", "防囤积居奇、哄抬价格、串通涨价", K.gold, HX.tint],
      ["下游", "CPI +0.9%（9月）", "防低价倾销、价格欺诈、虚假折扣、大数据杀熟", K.red, HX.redTint],
    ];
    fronts.forEach((f, i) => {
      const y = 2.05 + i * 1.3;
      box(s, MX, y, 5.6, 1.18, f[4]);
      T(s, f[0], { x: MX + 0.25, y: y + 0.12, w: 1.0, h: 0.45, fontSize: 18, bold: true, color: f[3] });
      T(s, f[1], { x: MX + 1.2, y: y + 0.12, w: 4.2, h: 0.45, fontSize: 18, bold: true });
      T(s, f[2], { x: MX + 0.25, y: y + 0.62, w: 5.1, h: 0.45, fontSize: 14, color: K.ink });
    });
    T(s, paras([
      "2026年上半年：**16项**专项行动，查办不正当竞争案件**11,465件**",
      "推进价格法修改：完善**低价倾销认定**与**行业成本调查**",
      "**政府侧**：公平竞争审查（条例2024年8月施行），清理地方保护与违规补贴",
    ], { fontSize: 14 }, { space: 6 }), { x: MX, y: 4.75, w: 5.6, h: 1.8, lineSpacingMultiple: 1.08 });
    T(s, "综合整治“内卷式”竞争：划清边界", { x: 6.6, y: 1.6, w: 6.1, h: 0.4, fontSize: 17, bold: true, color: K.jade });
    box(s, 6.6, 2.05, 6.13, 2.25, K.tint);
    T(s, "✔ 该治理", { x: 6.85, y: 2.13, w: 3, h: 0.38, fontSize: 16, bold: true, color: K.teal });
    const ok = [
      ["低于成本倾销、排挤竞争对手", "价格法"],
      ["平台强制或变相强制商家低于成本销售", "反不正当竞争法第十四条"],
      ["平台强制“退款不退货”等不合理规则", "平台规则监管办法"],
      ["偷工减料换来的“劣质低价”", "产品质量法"],
    ];
    T(s, ok.map((o, i) => [
      { text: o[0], options: { fontSize: 14, color: K.ink, bullet: { code: "25A0", indent: 14 }, paraSpaceAfter: 3 } },
      { text: "  " + o[1], options: { fontSize: 12, color: K.muted, breakLine: i < ok.length - 1 } },
    ]).flat(), { x: 6.85, y: 2.55, w: 5.75, h: 1.7 });
    box(s, 6.6, 4.4, 6.13, 1.5, HX.redTint);
    T(s, "✘ 不该做", { x: 6.85, y: 4.48, w: 3, h: 0.38, fontSize: 16, bold: true, color: K.red });
    T(s, paras(["干预政府定价目录外的自主定价、设“保护价”", "以“行业自律”名义固定价格、限产保价（涉嫌垄断协议）", "把技术进步带来的效率型降价当“内卷”打压"], { fontSize: 14 }, { space: 3 }), { x: 6.85, y: 4.88, w: 5.75, h: 0.98 });
    T(s, "判断标准：是否损害消费者长期福利与市场动态效率⏎上游涨价中也有供给收缩——防止从“治乱”滑向“限产保价”", { x: 6.6, y: 5.95, w: 6.13, h: 0.62, fontSize: 13, bold: true, color: K.jade, align: "center" });
    source(s, "国家统计局；国家市场监督管理总局2026年上半年新闻发布会；《反不正当竞争法》《网络交易平台规则监督管理办法》《公平竞争审查条例》。");
  }

  {
    const s = content(5, "抓手三 · 质量安全：以标准牵引消费升级，以安全底线托住信心",
      "【约3分钟】标准是“质量的语言”：把看不见的质量变成看得见的标识，大幅降低搜寻成本。链条四环：标准定义好、认证证明好、检测抽查核验好、召回追责兜住底。\n基层工作量最大的三块——食品安全、特种设备、民生计量——是“放心消费”的底线：商场电梯、景区索道和游乐设施、加油机和集贸市场电子秤，哪一样出问题，都会直接打击消费信心。\n三个现实信号：以旧换新补贴聚焦一级能效，能效标识的真实性直接关系财政资金效率；智能设备投诉增长26.6%，说明新产品标准跟不上；养老、托育、家政、文旅这些服务消费，最缺的是“看得见的质量”。\n贯穿案例：那台空调的能效标识，如果是虚标的，补贴就补错了地方。", "s_lever3");
    const chain = [
      ["FaClipboardCheck", "标准", "定义“好”", "产品标准、服务标准、能效等级"],
      ["FaCertificate", "认证", "证明“好”", "绿色产品认证、能效标识核查"],
      ["FaSearchDollar", "检测抽查", "核验“好”", "“双随机”抽查、网络商品抽检"],
      ["FaUndo", "召回追责", "兜住“底”", "缺陷产品召回、质量追溯"],
    ];
    const st = stepOf(4, 2.75);
    for (let i = 0; i < 4; i++) {
      const c = chain[i], x = MX + i * st;
      box(s, x, 1.6, 2.75, 1.95, K.tint);
      await iconCircle(s, c[0], x + 0.22, 1.78, 0.7, i === 3 ? K.red : K.teal);
      T(s, c[1], { x: x + 1.02, y: 1.8, w: 1.65, h: 0.4, fontSize: 18, bold: true });
      T(s, c[2], { x: x + 1.02, y: 2.2, w: 1.65, h: 0.32, fontSize: 14, color: K.teal, bold: true });
      T(s, c[3], { x: x + 0.22, y: 2.68, w: 2.35, h: 0.8, fontSize: 14, color: K.ink, lineSpacingMultiple: 1.1 });
      if (i < 3) T(s, "→", { x: x + 2.75, y: 2.35, w: st - 2.75, h: 0.5, fontSize: 20, bold: true, color: K.muted, align: "center" });
    }
    box(s, MX, 3.7, CW, 0.72, HX.redTint);
    T(s, [{ text: "安全底线  ", options: { bold: true, color: K.red, fontSize: 15 } }, { text: "食品安全全链条监管 ｜ 特种设备：商场电梯、景区索道、游乐设施 ｜ 民生计量：加油机、集贸市场电子秤", options: { fontSize: 14 } }],
      { x: MX + 0.3, y: 3.7, w: CW - 0.6, h: 0.72, valign: "middle" });
    const sig = [
      ["9146万台", "2026年1—8月6类家电以旧换新⏎（另数码和智能产品购新1.08亿件）", "补贴聚焦一级能效：能效标识真实性＝财政资金效率"],
      ["+26.6%", "2025年智能设备投诉举报增速⏎（15.2万件）", "新品标准滞后：标准缺位即监管盲区"],
      ["服务标准", "养老、托育、家政、文旅", "服务消费最缺“看得见的质量”"],
    ];
    const st3 = stepOf(3, 3.85);
    sig.forEach((g, i) => {
      const x = MX + i * st3;
      box(s, x, 4.6, 3.85, 1.95, HX.white, { line: { color: HX.line, width: 1 } });
      T(s, g[0], { x: x + 0.25, y: 4.7, w: 3.4, h: 0.5, fontSize: 22, bold: true, color: i === 1 ? K.red : K.jade, valign: "middle" });
      T(s, runs(g[1], { fontSize: 13, color: K.muted }), { x: x + 0.25, y: 5.22, w: 3.4, h: 0.6, fontSize: 13, lineSpacingMultiple: 1.05 });
      T(s, g[2], { x: x + 0.25, y: 5.85, w: 3.4, h: 0.62, fontSize: 14, bold: true, lineSpacingMultiple: 1.05 });
    });
    source(s, "国家发展改革委；国家市场监督管理总局。");
  }

  {
    const s = content(5, "抓手四 · 经营主体：降低进入与经营成本，守住就业“蓄水池”",
      "【约3分钟】经营主体是市场监管影响“能消费”这道门槛最直接的渠道。2025年全国新设经营主体2574.5万户，其中个体工商户1619.4万户。但新设多不等于活力强，要看净增、存活率和户均就业——更好的主指标是在营活跃企业，增长9.8%。\n个体工商户有两个身份：就业的“蓄水池”，消费的“毛细血管”。小店兴，则烟火旺。也要清醒：个体户大量新设，也可能是工资性岗位不足的被动创业。\n这条“准入—活跃—就业—增收—消费”的链，每一环都会漏损；监管能做的，是不让制度成本成为漏点：落实《民营经济促进法》，涉企检查“无事不扰”、“综合查一次”（源于浙江“大综合一体化”执法改革），个体工商户分型分类帮扶，信用修复“应修尽修”。", "s_lever4");
    chartCap(s, "图18  2025年新设经营主体构成（万户）", MX, 1.55, 4.6);
    s.addChart(pres.charts.DOUGHNUT, [{ name: "新设经营主体", labels: ["个体工商户", "企业"], values: [1619.4, 950.0] }], chartOpts({
      x: MX + 0.1, y: 2.0, w: 4.4, h: 4.4, chartColors: [HX.teal, HX.blue], holeSize: 55, showValue: true, dataLabelFormatCode: "0.0",
      dataLabelColor: HX.white, dataLabelFontBold: true, dataLabelFontSize: 13, showLegend: true, legendPos: "b", legendFontSize: 13, dataBorder: { pt: 2, color: "FFFFFF" },
    }));
    T(s, [{ text: "2574.5", options: { fontSize: 22, bold: true, color: K.jade, breakLine: true } }, { text: "万户", options: { fontSize: 13, color: K.muted } }],
      { x: MX + 1.5, y: 3.5, w: 1.6, h: 0.9, align: "center", valign: "middle" });
    const chain = ["准入便利", "主体活跃", "吸纳就业", "居民增收", "消费扩大"];
    chain.forEach((c, i) => {
      const x = 5.5 + i * 1.47;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.7, w: 1.22, h: 0.75, fill: { color: i === 4 ? K.red : K.teal }, line: { color: K.white, width: 0 }, rectRadius: 0.08 });
      T(s, c, { x, y: 1.7, w: 1.22, h: 0.75, fontSize: 14, bold: true, color: K.white, align: "center", valign: "middle" });
      if (i < 4) T(s, "→", { x: x + 1.22, y: 1.8, w: 0.25, h: 0.5, fontSize: 16, bold: true, color: K.muted, align: "center" });
    });
    box(s, 5.5, 2.7, 7.23, 3.85, K.tint);
    T(s, paras([
      "2025年在营活跃企业**+9.8%**，占全部企业比重提高4.7个百分点",
      "新设多≠活力强：要看**净增、存活率与户均就业**",
      "浙江经营主体**1163.65万户**，平均每5.8名常住人口就有1户",
      "落实《民营经济促进法》：涉企检查“**无事不扰**”、“综合查一次”；个体户**分型分类帮扶**；信用修复“应修尽修”",
    ], { fontSize: 15 }, { space: 12 }), { x: 5.75, y: 2.9, w: 6.8, h: 3.5, lineSpacingMultiple: 1.12 });
    source(s, "国家市场监督管理总局（另有农民专业合作社等5.1万户）；长三角市场监管部门《长三角经营主体发展指数（2025版）》；浙江省统计局。");
  }

  {
    const s = content(5, "抓手五 · 新业态：包容审慎，压实平台主体责任",
      "【约3分钟】两个数：今年1—8月网上商品零售额8.42万亿元，约占社零的25.7%；网购占全部投诉举报的56.9%——四分之一的消费，一半以上的纠纷。\n今年最有分量的案件：4月17日，总局对7家平台的“幽灵外卖”系列案作出处罚，罚没35.97亿元，是《食品安全法》施行以来对外卖平台开出的最大罚单；仅蛋糕类“幽灵店铺”就有6.7万余家。浙江的“外卖在线”在全国率先与平台打通证照核验，累计拦截异常入驻4.2万家次。\n规则也在完善：《网络交易平台规则监督管理办法》2026年2月1日起施行，平台不得强制商家承担“退款不退货”等售后责任。\n【管辖提醒，讲前核对】平台经营者的违法行为一般由其住所地管辖，平台内经营者由实际经营地管辖——属地平台的主体责任，也是浙江市场监管的特殊责任。", "s_lever5");
    const stats = [["25.7%", "网上商品零售额占社零（2026年1—8月）"], ["56.9%", "网购占全部投诉举报（2025年）"]];
    stats.forEach((t, i) => {
      const x = MX + i * 6.18;
      T(s, t[0], { x, y: 1.55, w: 1.9, h: 0.6, fontSize: 30, bold: true, color: i ? K.red : K.teal, valign: "middle" });
      T(s, t[1], { x: x + 2.0, y: 1.55, w: 4.0, h: 0.6, fontSize: 14, color: K.muted, valign: "middle" });
    });
    const grid = [
      ["FaMobileAlt", "直播电商", "虚假宣传、带货责任不清、售后难", "主播与平台责任界定、直播回放留存、重点抽查"],
      ["FaStore", "即时零售与外卖", "补贴大战、低价倾销、食品安全", "入网资质核验、平台规则审查、明厨亮灶"],
      ["FaRegLightbulb", "智能与AI消费", "算法“杀熟”、数据安全、标准缺位", "算法透明、标准先行、个人信息保护"],
      ["FaUsers", "银发与健康消费", "保健品虚假宣传、养老预付费风险", "重点人群专项整治、预付资金监管协同"],
    ];
    const cw = (CW - 0.25) / 2;
    for (let i = 0; i < 4; i++) {
      const g = grid[i], x = MX + (i % 2) * (cw + 0.25), y = 2.25 + Math.floor(i / 2) * 1.6;
      box(s, x, y, cw, 1.48, i % 3 === 0 ? K.tint : HX.gray);
      await iconCircle(s, g[0], x + 0.25, y + 0.22, 0.62, i < 2 ? K.teal : K.blue);
      T(s, g[1], { x: x + 1.05, y: y + 0.18, w: 4.7, h: 0.4, fontSize: 17, bold: true });
      T(s, [{ text: "风险：", options: { bold: true, color: K.red } }, { text: g[2], options: {} }], { x: x + 1.05, y: y + 0.62, w: cw - 1.2, h: 0.38, fontSize: 14 });
      T(s, [{ text: "工具：", options: { bold: true, color: K.teal } }, { text: g[3], options: {} }], { x: x + 1.05, y: y + 1.0, w: cw - 1.2, h: 0.4, fontSize: 14 });
    }
    box(s, MX, 5.5, CW, 0.75, HX.redTint);
    T(s, [{ text: "案例  ", options: { bold: true, color: K.red, fontSize: 15 } }, { text: "2026年4月总局查处“幽灵外卖”：7家平台被罚没35.97亿元，为《食品安全法》施行以来对外卖平台最大罚单；“浙江外卖在线”证照核验累计拦截异常入驻4.2万家次", options: { fontSize: 14 } }],
      { x: MX + 0.25, y: 5.5, w: CW - 0.5, h: 0.75, valign: "middle" });
    T(s, "包容审慎 ≠ 放任不管：规则已出，关键在执行", { x: MX, y: 6.3, w: CW, h: 0.35, fontSize: 15, bold: true, color: K.jade, align: "center" });
    source(s, "国家统计局（占比为作者计算）；国家市场监督管理总局；新华社。");
  }

  {
    const s = content(5, "带走一页纸：三件必做＋七件选做",
      "【约3分钟】这一页建议印成A4发给学员。三件必做放在最上面，每件都配了“怎么算干成了”的指标；七件选做，各地按实际挑。\n【互动，课前预约】请1—2个地市各用1分钟说：我们挑哪一件、为什么。不要现场临时点名。", "s_list");
    const must = [
      ["1", "以旧换新比价与能效检查", "查虚构原价、先提价后打折、虚标能效、虚假回收", "看：补贴品类价格异常率、能效标识抽查合格率"],
      ["2", "预付式消费格式条款治理", "查“概不退款”等无效条款；会同行业部门推动资金存管", "看：预付式投诉占比"],
      ["3", "重复投诉溯源治理", "同一经营者、同类问题反复被投诉的，诉转案、约谈、公示", "看：重复投诉率、和解率"],
    ];
    const st = stepOf(3, 3.85);
    must.forEach((m, i) => {
      const x = MX + i * st;
      box(s, x, 1.6, 3.85, 2.05, HX.redTint);
      pill(s, x + 0.2, 1.75, 0.95, 0.38, "必做" + m[0], K.red, K.white, 13);
      T(s, m[1], { x: x + 1.25, y: 1.7, w: 2.5, h: 0.5, fontSize: 15, bold: true, valign: "middle" });
      T(s, m[2], { x: x + 0.2, y: 2.28, w: 3.45, h: 0.75, fontSize: 14, color: K.ink, lineSpacingMultiple: 1.05 });
      T(s, m[3], { x: x + 0.2, y: 3.05, w: 3.45, h: 0.5, fontSize: 12, color: K.muted, lineSpacingMultiple: 1.0 });
    });
    const opt = [
      ["放心消费", "无理由退货承诺单元扩至商圈、景区、夜市"],
      ["公平竞争", "公平竞争审查：清理违规补贴与地方保护"],
      ["经营主体", "涉企检查“综合查一次”、扫码入企"],
      ["经营主体", "个体户分型帮扶；信用“应修尽修”"],
      ["经营主体", "落实首违不罚、轻微免罚清单"],
      ["新业态", "直播、即时零售、外卖重点监测"],
      ["质量安全", "智能、银发产品抽检；服务质量评价试点"],
    ];
    const tagCol = { "放心消费": K.red, "公平竞争": K.blue, "经营主体": K.teal, "新业态": K.gold, "质量安全": K.jade };
    const cw = (CW - 0.25) / 2;
    opt.forEach((o, i) => {
      const col = i < 4 ? 0 : 1, row = i < 4 ? i : i - 4;
      const x = MX + col * (cw + 0.25), y = 3.9 + row * 0.66;
      box(s, x, y, cw, 0.56, row % 2 ? HX.gray : K.tint);
      pill(s, x + 0.15, y + 0.11, 1.15, 0.34, o[0], tagCol[o[0]], K.white, 12);
      T(s, o[1], { x: x + 1.45, y, w: cw - 1.55, h: 0.56, fontSize: 14, valign: "middle" });
    });
    T(s, "选做七件", { x: MX + cw + 0.25, y: 5.95, w: cw, h: 0.56, fontSize: 14, color: K.muted, align: "right", valign: "middle" });
  }

  // ===================== 结语 =====================
  section("结语");
  {
    const s = content(9, "回到开头：有增长、缺消费，市场监管够得着什么？",
      "【约2分钟】回到开场的问题。够得着的：第四道门槛“信任”的全部，加上第三道门槛里的标准、准入和质量安全。前两道门槛主要靠宏观和民生政策，我们不越位。管得了的：就是刚才那三件必做。管得好不好，一年后看三个数：重复投诉率有没有下降，补贴品类价格异常有没有下降，在营活跃主体占比有没有上升。\n开场我问过，那台空调的600元补贴落进了谁的口袋。希望一年以后，我们能回答：落进了老百姓的口袋——因为有人在守门。", "s_back");
    const cols = [
      ["够得着", "FaHandshake", "门槛④ 信任：全部⏎门槛③：标准、准入、质量安全⏎门槛①②：宏观与民生政策为主⏎——监管不越位", K.teal],
      ["管得了", "FaClipboardCheck", "三件必做：⏎① 以旧换新比价与能效检查⏎② 预付式格式条款治理⏎③ 重复投诉溯源治理", K.red],
      ["管得好，看三个数", "FaChartLine", "① 重复投诉率下降⏎② 补贴品类价格异常下降⏎③ 在营活跃主体占比上升", K.blue],
    ];
    const st = stepOf(3, 3.85);
    for (let i = 0; i < 3; i++) {
      const c = cols[i], x = MX + i * st;
      box(s, x, 1.65, 3.85, 3.6, K.tint);
      await iconCircle(s, c[1], x + 0.3, 1.85, 0.75, c[3]);
      T(s, c[0], { x: x + 1.2, y: 1.9, w: 2.5, h: 0.65, fontSize: 20, bold: true, color: c[3], valign: "middle" });
      T(s, runs(c[2], { fontSize: 16, color: K.ink }), { x: x + 0.3, y: 2.85, w: 3.35, h: 2.3, fontSize: 16, lineSpacingMultiple: 1.45 });
    }
    box(s, MX, 5.5, CW, 1.0, K.jade);
    T(s, "对违法者“利剑高悬”，对守法者“无事不扰”——让消费者时时感受到监管的保护。", { x: MX + 0.35, y: 5.5, w: CW - 0.7, h: 1.0, fontSize: 18, bold: true, color: K.white, valign: "middle", align: "center" });
  }
  {
    const s = pres.addSlide({ masterName: "DARK", sectionTitle: CURRENT_SECTION });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 1.3, w: 1.0, h: 1.0, fill: { color: K.red }, line: { color: K.red, width: 0 }, rectRadius: 0.1 });
    T(s, "信", { x: 1.0, y: 1.3, w: 1.0, h: 1.0, fontSize: 42, bold: true, color: K.white, align: "center", valign: "middle" });
    T(s, "宏观政策提振信心，市场监管守护信任。", { x: 1.0, y: 2.65, w: 11.3, h: 0.9, fontSize: 36, bold: true, color: K.white, valign: "middle" });
    T(s, "信心决定人们敢不敢花，信任决定人们放不放心买。", { x: 1.0, y: 3.6, w: 11.3, h: 0.7, fontSize: 24, color: HX.darkText, valign: "middle" });
    s.addShape(pres.shapes.LINE, { x: 1.0, y: 4.75, w: 6.2, h: 0 , line: { color: "2E5F5F", width: 0.75 } });
    T(s, "谢谢！欢迎批评指正", { x: 1.0, y: 5.0, w: 8, h: 0.6, fontSize: 24, bold: true, color: K.goldLight });
    s.addNotes("【结束语，约1分钟】宏观政策提振信心，市场监管守护信任。信心决定人们敢不敢花，信任决定人们放不放心买。谢谢大家！\n【答疑约8分钟】预案见下一页隐藏页“答疑预案”。");
  }
  {
    const s = content(9, "延伸阅读：本讲引用的主要文件与法规",
      "【不展开，留给学员课后查阅】");
    box(s, MX, 1.6, 5.95, 4.95, K.tint);
    T(s, "政策文件", { x: MX + 0.3, y: 1.75, w: 5.3, h: 0.4, fontSize: 17, bold: true, color: K.teal });
    T(s, paras([
      "2024年、2025年中央经济工作会议",
      "“十五五”规划建议（2025年10月）",
      "2026年《政府工作报告》",
      "《提振消费专项行动方案》（2025年3月）",
      "《扩大消费“十五五”规划》（2026年7月）",
      "《优化消费环境三年行动方案（2025—2027年）》",
      "总局《关于全面开展放心消费单元和集聚区培育的通知》",
    ], { fontSize: 15 }, { space: 12 }), { x: MX + 0.3, y: 2.3, w: 5.4, h: 4.1, lineSpacingMultiple: 1.1 });
    const x2 = MX + 6.18;
    box(s, x2, 1.6, 5.95, 4.95, HX.redTint);
    T(s, "法律法规", { x: x2 + 0.3, y: 1.75, w: 5.3, h: 0.4, fontSize: 17, bold: true, color: K.red });
    T(s, paras([
      "《消费者权益保护法》及实施条例（2024年7月施行）",
      "《反不正当竞争法》（2025年10月15日修订施行）",
      "《反垄断法》《价格法》《产品质量法》《食品安全法》",
      "《公平竞争审查条例》（2024年8月施行）",
      "《民营经济促进法》（2025年5月施行）",
      "《网络交易平台规则监督管理办法》（2026年2月施行）",
      "《明码标价和禁止价格欺诈规定》",
    ], { fontSize: 15 }, { space: 12 }), { x: x2 + 0.3, y: 2.3, w: 5.4, h: 4.1, lineSpacingMultiple: 1.1 });
  }

  // ---------------- 隐藏备用页 ----------------
  {
    const s = content(9, "答疑预案（隐藏页）", "【备用，不投影】每题准备30秒。");
    s.hidden = true;
    const qa = [
      ["反内卷会不会推高物价？", "目标是让价格说真话：治理强制低价、掠夺性低价和劣质低价，效率型降价受保护；防从“治乱”滑向“限产保价”。"],
      ["“先提价后打折”怎么认定？", "依据《明码标价和禁止价格欺诈规定》：虚构被比较价格、临时提高标价作折扣基准属价格欺诈；商务部门取消资格、追回补贴，市场监管查处、纳入失信。"],
      ["投诉量上升，考核怎么办？", "考核看结构：重复投诉率、按期办结率、和解率、承诺单元覆盖率，而不是投诉总量。"],
      ["“无事不扰”和严格执法怎么平衡？", "信用分级分类＋双随机；首违不罚、轻微免罚清单（国市监稽发〔2025〕10号）；食品、特种设备等底线零容忍。"],
      ["外地平台与本地消费者的纠纷归谁管？", "平台经营者一般由住所地管辖，平台内经营者由实际经营地管辖；投诉按网络交易监管有关规定处理（讲前核对条款）。"],
      ["社零这么低，经济会不会出问题？", `回到第${REG.s_judge}页：总量稳、结构在调整；服务零售快于商品零售；补贴回吐与高基数不等于趋势。`],
      ["预付卡跑路，资金监管归谁？", "行业主管部门牵头资金存管；市场监管管合同格式条款、价格和广告。"],
    ];
    qa.forEach((q, i) => {
      const col = i < 4 ? 0 : 1, row = i < 4 ? i : i - 4;
      const cw = (CW - 0.25) / 2, x = MX + col * (cw + 0.25), y = 1.6 + row * 1.25;
      box(s, x, y, cw, 1.15, row % 2 ? HX.gray : K.tint);
      T(s, [{ text: `Q${i + 1} ${q[0]}`, options: { bold: true, fontSize: 14, color: K.red, breakLine: true } }, { text: q[1], options: { fontSize: 12, color: K.ink } }],
        { x: x + 0.2, y: y + 0.08, w: cw - 0.4, h: 1.0, valign: "top", lineSpacingMultiple: 1.05 });
    });
  }
  {
    const s = content(9, "附：数据说明与讲前更新清单（隐藏页）", "【备课用，不投影】");
    s.hidden = true;
    T(s, "数据说明", { x: MX, y: 1.6, w: 5.9, h: 0.4, fontSize: 17, bold: true, color: K.jade });
    T(s, paras([
      "宏观数据来自国家统计局；2025年社零按第五次经济普查修订后可比口径",
      "名义GDP增速、收入占比、平均消费倾向、两年复合增速等为作者据官方数据计算",
      "国际居民消费率为世界银行WDI取整值（2023年或最近年份），引用前再核",
      "日本CPI为总务省年度同比；空调案例为示意数字",
      "市场监管数据来自总局发布会、3·15数据发布及浙江省局公开资料",
    ], { fontSize: 13 }, { space: 6 }), { x: MX, y: 2.1, w: 5.9, h: 4.4, lineSpacingMultiple: 1.08 });
    box(s, 6.85, 1.6, 5.88, 4.95, HX.redTint);
    T(s, "讲前更新清单（前三季度数据10月19日10:00发布）", { x: 7.1, y: 1.72, w: 5.5, h: 0.4, fontSize: 14, bold: true, color: K.red });
    const p = (k) => "第" + REG[k] + "页";
    T(s, paras([
      `前三季度GDP、1—9月社零与投资 → ${p("s_three")}、${p("s_gdp")}、${p("s_nominal")}、${p("s_eng")}、${p("s_rs")}、${p("s_judge")}`,
      `9月CPI、PPI已更新（10月14日发布）；10月数据课后才发布`,
      `前三季度居民收支 → ${p("s_frame")}、${p("s_income")}、${p("s_incline")}`,
      `1—9月房地产 → ${p("s_house")}`,
      `9月青年失业率可能课后才发布，${p("s_jobs")}以8月为准`,
      `以旧换新最新进展 → ${p("s_tradein")}、${p("s_lever3")}`,
      `浙江前三季度数据 → ${p("s_zj")}；网上零售占比 → ${p("s_lever5")}`,
      `外卖补贴十条是否正式出台 → ${p("s_quiz")}`,
      "封面补充主讲人姓名；明码标价规定、反垄断法相关条号讲前核对",
    ], { fontSize: 13 }, { space: 5 }), { x: 7.1, y: 2.15, w: 5.45, h: 4.3, lineSpacingMultiple: 1.05 });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  const zip = await JSZip.loadAsync(fs.readFileSync(OUT));
  const part = "ppt/theme/theme1.xml";
  let xml = await zip.file(part).async("string");
  xml = xml.replace(/<a:ea typeface=""\/>/g, `<a:ea typeface="${FONT}"/>`).replace(/<a:font script="Hans" typeface="[^"]*"\/>/g, `<a:font script="Hans" typeface="${FONT}"/>`);
  zip.file(part, xml);
  // pptxgenjs 会在多片段段落中重复写入 <a:pPr>（不合 OOXML 顺序），只保留段首那一个；项目符号统一着色
  let fixed = 0;
  for (const name of Object.keys(zip.files).filter((n) => /^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(n))) {
    let out = (await zip.file(name).async("string")).replace(/<a:p>([\s\S]*?)<\/a:p>/g, (m, inner) => {
      const cleaned = inner.replace(/<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/g, (pp, off) => {
        if (off === 0) return pp;
        fixed++;
        return "";
      });
      return "<a:p>" + cleaned + "</a:p>";
    });
    out = out.replace(/<a:buSzPct val="100000"\/><a:buChar/g, '<a:buClr><a:srgbClr val="00897B"/></a:buClr><a:buSzPct val="100000"/><a:buChar');
    zip.file(name, out);
  }
  console.log("removed mid-paragraph pPr:", fixed);
  fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  console.log("wrote", OUT, "slides:", pres.slides.length, JSON.stringify(REG));
}

build().catch((e) => { console.error(e); process.exit(1); });
