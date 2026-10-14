// 信心与信任：从消费看中国经济形势与市场监管的作为 —— 讲座课件生成脚本
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));
const fs = require("fs");
const path = require("path");
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
    accent5: "6B7280", accent6: "E0B252", hlink: "4A63A6", folHlink: "6B7280",
  },
};
const HX = {
  ink: "22272E", jade: "0B3B3C", tint: "EEF5F4", white: "FFFFFF",
  teal: "00897B", red: "C8442D", blue: "4A63A6", gold: "B07D0E",
  muted: "6B7280", goldLight: "E0B252", redTint: "FBEFEC", gray: "F3F4F6",
  grid: "E5E7EB", axis: "6B7280", context: "A9B4BE", line: "D1D5DB", darkText: "D7E4E2",
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

const W = 13.333, H = 7.5, MX = 0.6, CW = W - 2 * MX;

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
  slide.addText(text, o);
}
// “**重点**”标记 → 加粗着色的文字片段
function runs(str, base, hl) {
  const out = [];
  str.split(/\*\*(.+?)\*\*/).forEach((seg, i) => {
    if (!seg) return;
    out.push({ text: seg, options: Object.assign({}, base, i % 2 ? { bold: true, color: hl || K.red } : {}) });
  });
  return out;
}
// 多段落（可带项目符号）
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
function source(slide, text) {
  T(slide, "资料来源：" + text, { x: MX, y: 6.68, w: CW, h: 0.3, fontSize: 10, color: K.muted });
}
const SEALS = { 1: "观", 2: "析", 3: "鉴", 4: "策", 5: "为", 0: "引", 9: "结" };
const SECNAME = {
  0: "开篇", 1: "一 · 观势", 2: "二 · 析因", 3: "三 · 镜鉴", 4: "四 · 谋策", 5: "五 · 作为", 9: "结语",
};
function content(sec, title, notes) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: CURRENT_SECTION });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: 0.3, w: 0.32, h: 0.32, fill: { color: K.red }, line: { color: K.red, width: 0 }, rectRadius: 0.05 });
  T(s, SEALS[sec], { x: MX, y: 0.3, w: 0.32, h: 0.32, fontSize: 12, bold: true, color: K.white, align: "center", valign: "middle" });
  T(s, SECNAME[sec], { x: MX + 0.45, y: 0.3, w: 6, h: 0.32, fontSize: 12, bold: true, color: K.red, valign: "middle" });
  s.addText(title, { placeholder: "title", fontFace: FONT, lang: "zh-CN" });
  if (notes) s.addNotes(notes);
  return s;
}
let CURRENT_SECTION = "开篇";
function section(title) { CURRENT_SECTION = title; pres.addSection({ title }); }

// 右侧“看点”面板
function insight(slide, items, opt = {}) {
  const x = opt.x || 8.85, y = opt.y || 1.6, w = opt.w || 3.88, h = opt.h || 4.95;
  box(slide, x, y, w, h, K.tint);
  T(slide, opt.head || "看点", { x: x + 0.25, y: y + 0.2, w: w - 0.5, h: 0.35, fontSize: 16, bold: true, color: K.teal });
  const regH = opt.reg ? (opt.regH || 1.45) : 0;
  T(slide, paras(items, { fontSize: 14, color: K.ink }, { space: 9 }), { x: x + 0.25, y: y + 0.65, w: w - 0.5, h: h - 0.85 - regH, lineSpacingMultiple: 1.12 });
  if (opt.reg) {
    const ry = y + h - regH - 0.15;
    box(slide, x + 0.15, ry, w - 0.3, regH, HX.redTint);
    T(slide, [
      { text: "监管视角  ", options: { bold: true, color: K.red, fontSize: 14 } },
      ...runs(opt.reg, { fontSize: 14, color: K.ink }),
    ], { x: x + 0.3, y: ry + 0.12, w: w - 0.6, h: regH - 0.2, lineSpacingMultiple: 1.1 });
  }
}
function chartCap(slide, text, x = MX, y = 1.55, w = 7.95) {
  T(slide, text, { x, y, w, h: 0.32, fontSize: 13, bold: true, color: K.ink });
}
function chartOpts(extra) {
  return Object.assign({
    catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT, dataLabelFontFace: FONT, legendFontFace: FONT, titleFontFace: FONT,
    catAxisLabelColor: "4B5563", valAxisLabelColor: HX.axis, catAxisLabelFontSize: 11, valAxisLabelFontSize: 11,
    valGridLine: { color: HX.grid, size: 0.75 }, catGridLine: { style: "none" },
    catAxisLineShow: true, catAxisLineColor: "9CA3AF", valAxisLineShow: false,
    dataLabelFontSize: 11, dataLabelColor: HX.ink,
    showLegend: false, legendPos: "t", legendFontSize: 12, legendColor: "374151",
    showTitle: false, barGapWidthPct: 55,
  }, extra);
}
const ICON_CACHE = {};
async function icon(name, color = "#FFFFFF") {
  const key = name + color;
  if (!ICON_CACHE[key]) {
    const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fa[name], { color, size: 256 }));
    const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
    ICON_CACHE[key] = "image/png;base64," + buf.toString("base64");
  }
  return ICON_CACHE[key];
}
async function iconCircle(slide, name, x, y, d, fill) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 } });
  const p = d * 0.27;
  slide.addImage({ data: await icon(name), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
}
function stat(slide, x, y, w, value, label, sub, color) {
  T(slide, value, { x, y, w, h: 0.95, fontSize: 44, bold: true, color: color || K.teal, valign: "bottom" });
  T(slide, label, { x, y: y + 1.0, w, h: 0.38, fontSize: 15, bold: true, color: K.ink });
  if (sub) T(slide, sub, { x, y: y + 1.4, w, h: 0.7, fontSize: 12, color: K.muted, lineSpacingMultiple: 1.1 });
}

// 分篇扉页
const PARTS = ["观势", "析因", "镜鉴", "谋策", "作为"];
function divider(n, title, question, notes) {
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: CURRENT_SECTION });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 2.05, w: 1.6, h: 1.6, fill: { color: K.red }, line: { color: K.red, width: 0 }, rectRadius: 0.12 });
  T(s, SEALS[n], { x: 1.0, y: 2.05, w: 1.6, h: 1.6, fontSize: 66, bold: true, color: K.white, align: "center", valign: "middle" });
  T(s, `第${"一二三四五"[n - 1]}部分 · ${PARTS[n - 1]}`, { x: 3.05, y: 2.0, w: 9, h: 0.45, fontSize: 18, bold: true, color: K.goldLight });
  T(s, title, { x: 3.05, y: 2.5, w: 9.5, h: 0.95, fontSize: 40, bold: true, color: K.white, valign: "middle" });
  T(s, question, { x: 3.05, y: 3.6, w: 9.3, h: 0.9, fontSize: 18, color: HX.darkText, italic: true, lineSpacingMultiple: 1.15 });
  // 进度导航
  PARTS.forEach((p, i) => {
    const active = i === n - 1;
    T(s, `${"一二三四五"[i]} ${p}`, { x: 1.0 + i * 2.35, y: 5.95, w: 2.1, h: 0.4, fontSize: 15, bold: active, color: active ? K.goldLight : "7FA3A0" });
  });
  s.addShape(pres.shapes.LINE, { x: 1.0, y: 5.8, w: 11.33, h: 0, line: { color: "2E5F5F", width: 0.75 } });
  if (notes) s.addNotes(notes);
  return s;
}

// ---------- 数据 ----------
const D = {
  gdpYears: ["2012", "2013", "2014", "2015", "2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025", "2026上半年"],
  gdp: [7.9, 7.8, 7.4, 7.0, 6.8, 6.9, 6.7, 6.0, 2.2, 8.4, 3.1, 5.2, 5.0, 5.0, 4.7],
  nrYears: ["2019", "2020", "2021", "2022", "2023", "2024", "2025"],
  nominal: [7.3, 2.7, 13.4, 4.8, 4.6, 4.2, 4.0],
  real: [6.0, 2.2, 8.4, 3.1, 5.2, 5.0, 5.0],
  engYears: ["2019", "2020", "2021", "2022", "2023", "2024", "2025"],
  engC: [3.5, -0.5, 5.3, 1.0, 4.3, 2.2, 2.6],
  engI: [1.7, 2.2, 1.1, 1.5, 1.5, 1.3, 0.8],
  engX: [0.8, 0.6, 1.7, 0.5, -0.6, 1.5, 1.6],
  rsLabels: ["25年1-2月", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "26年1-2月", "3", "4", "5", "6", "7", "8"],
  rs: [4.0, 5.9, 5.1, 6.4, 4.8, 3.7, 3.4, 3.0, 2.9, 1.3, 0.9, 2.8, 1.7, 0.2, -0.6, 1.0, 0.6, 0.4],
  pLabels: ["25年1月", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "26年1月", "2", "3", "4", "5", "6", "7", "8"],
  cpi: [0.5, -0.7, -0.1, -0.1, -0.1, 0.1, 0.0, -0.4, -0.3, 0.2, 0.7, 0.8, 0.2, 1.3, 1.0, 1.2, 1.2, 1.0, 0.5, 0.8],
  ppi: [-2.3, -2.2, -2.5, -2.7, -3.3, -3.6, -3.6, -2.9, -2.3, -2.1, -2.2, -1.9, -1.4, -0.9, 0.5, 2.8, 3.9, 4.1, 3.5, 3.8],
  intl: [["美国", 68], ["英国", 61], ["印度", 60], ["日本", 55], ["法国", 53], ["德国", 52], ["韩国", 48], ["中国", 40]],
  incYears: ["2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025", "2026上半年"],
  incGDP: [6.8, 6.9, 6.7, 6.0, 2.2, 8.4, 3.1, 5.2, 5.0, 5.0, 4.7],
  incReal: [6.3, 7.3, 6.5, 5.8, 2.1, 8.1, 2.9, 6.1, 5.1, 5.0, 4.2],
  conReal: [6.8, 5.4, 6.2, 5.5, -4.0, 12.6, -0.2, 9.0, 5.1, 4.4, 2.7],
  apcYears: ["2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025"],
  apc: [70.5, 70.3, 70.1, 65.9, 68.6, 66.5, 68.3, 68.3, 68.0],
  depYears: ["2019", "2020", "2021", "2022", "2023", "2024", "2025"],
  dep: [9.7, 11.3, 9.9, 17.8, 16.7, 14.3, 14.6],
  houseYears: ["2019", "2020", "2021", "2022", "2023", "2024", "2025"],
  house: [17.16, 17.61, 17.94, 13.58, 11.17, 9.74, 8.81],
  svcYears: ["2013", "2019", "2020", "2021", "2022", "2023", "2024", "2025"],
  svc: [39.7, 45.9, 42.6, 44.2, 43.2, 45.2, 46.1, 46.1],
  compCats: ["食品烟酒", "居住", "交通通信", "教育文化娱乐", "医疗保健", "生活用品及服务", "衣着", "其他用品及服务"],
  compShare: [29.3, 21.7, 14.6, 11.8, 8.7, 5.7, 5.3, 2.9],
  compGrowth: [2.6, 2.1, 8.3, 9.4, 1.0, 7.7, 2.2, 11.2],
  jpYears: Array.from({ length: 26 }, (_, i) => String(1990 + i)),
  jpCPI: [3.1, 3.3, 1.6, 1.3, 0.7, -0.1, 0.1, 1.8, 0.6, -0.3, -0.7, -0.7, -0.9, -0.3, 0.0, -0.3, 0.2, 0.1, 1.4, -1.4, -0.7, -0.3, 0.0, 0.4, 2.7, 0.8],
};

// ====================================================================
async function build() {
  // ---------------- 封面 ----------------
  section("开篇");
  {
    const s = pres.addSlide({ masterName: "DARK", sectionTitle: CURRENT_SECTION });
    T(s, "全省市场监管系统干部专题讲座", { x: 1.0, y: 0.9, w: 8, h: 0.4, fontSize: 16, color: K.goldLight, bold: true });
    T(s, "信心与信任", { x: 1.0, y: 1.9, w: 8.5, h: 1.4, fontSize: 66, bold: true, color: K.white, valign: "middle" });
    T(s, "从消费看中国经济形势与市场监管的作为", { x: 1.0, y: 3.35, w: 9, h: 0.6, fontSize: 26, color: HX.darkText });
    s.addShape(pres.shapes.LINE, { x: 1.0, y: 4.45, w: 6.2, h: 0, line: { color: "2E5F5F", width: 0.75 } });
    T(s, "浙江财经大学经济学院 ｜ 挂职浙江省市场监督管理局经济发展处", { x: 1.0, y: 4.65, w: 9, h: 0.4, fontSize: 15, color: HX.darkText });
    T(s, "2026年10月", { x: 1.0, y: 5.1, w: 6, h: 0.4, fontSize: 15, color: HX.darkText });
    // 朱印
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.85, y: 1.75, w: 2.3, h: 2.3, fill: { color: K.red }, line: { color: K.red, width: 0 }, rectRadius: 0.15 });
    T(s, "信", { x: 9.85, y: 1.75, w: 2.3, h: 2.3, fontSize: 110, bold: true, color: K.white, align: "center", valign: "middle" });
    T(s, "信心 · 信任", { x: 9.85, y: 4.2, w: 2.3, h: 0.4, fontSize: 14, color: K.goldLight, align: "center" });
    s.addNotes("【开场，约1分钟】各位领导、各位同事，大家好。今天这堂课的题目叫“信心与信任”。为什么取这四个字？因为我想用两个小时回答一个问题：中国经济增长不慢，为什么消费总是提不起来？而在这个问题里，市场监管部门到底处在什么位置。我的基本观点先亮出来：宏观政策解决的是“信心”问题，市场监管解决的是“信任”问题——两者缺一不可。\n【时间分配】开篇5分钟｜观势18分钟｜析因32分钟｜休息10分钟｜镜鉴12分钟｜谋策13分钟｜作为25分钟｜结语5分钟。");
  }

  // ---------------- 三个数字 ----------------
  {
    const s = content(0, "三个数字：经济在增长，消费在放缓，份额在低位",
      "【约2分钟】先看三个数字。4.7%——今年上半年GDP增速，在全球主要经济体里不算低；1.1%——今年1到8月社会消费品零售总额只增长了1.1%，比GDP慢了一大截；39.9%——居民消费占GDP的比重，大约四成，而主要经济体大多在五成上下到近七成之间。注意：社零是名义增速、GDP是实际增速，口径不同，但扣除约1%的物价涨幅后，消费偏弱的结论只会更明显。三个数字放在一起，就是今天要讨论的“温差”：有增长、缺消费。\n【提问互动】大家可以先想一想：如果让你给消费不足找一个原因，你会说什么？（收入？房子？不敢花？没东西买？）——今天我们把这些原因一层层拆开。");
    const cols = [
      ["4.7%", "2026年上半年GDP实际增速", "一季度5.0%、二季度4.3%；\n全年目标4.5%—5%", K.teal],
      ["1.1%", "2026年1—8月社零总额增速", "8月当月仅增长0.4%（名义）；\n扣除物价后实际增长更低", K.red],
      ["39.9%", "居民消费率（2024年）", "主要经济体多在48%—68%；\n“十五五”要求明显提高", K.blue],
    ];
    cols.forEach((c, i) => {
      const x = MX + i * 4.1;
      box(s, x, 1.75, 3.85, 3.1, i === 1 ? HX.redTint : K.tint);
      T(s, c[0], { x: x + 0.35, y: 1.95, w: 3.3, h: 1.15, fontSize: 60, bold: true, color: c[3], valign: "bottom" });
      T(s, c[1], { x: x + 0.35, y: 3.2, w: 3.3, h: 0.4, fontSize: 16, bold: true });
      T(s, c[2], { x: x + 0.35, y: 3.7, w: 3.3, h: 0.9, fontSize: 13, color: K.muted, lineSpacingMultiple: 1.15 });
    });
    box(s, MX, 5.15, CW, 1.25, K.jade);
    T(s, [
      { text: "今天的核心问题  ", options: { bold: true, color: K.goldLight, fontSize: 17 } },
      { text: "为什么“有增长、缺消费”？哪些堵点是市场监管够得着、管得了、管得好的？", options: { color: K.white, fontSize: 17 } },
    ], { x: MX + 0.4, y: 5.15, w: CW - 0.8, h: 1.25, valign: "middle" });
    source(s, "国家统计局；世界银行WDI。居民消费率为支出法口径，2024年。");
  }

  // ---------------- 一条主线 ----------------
  {
    const s = content(0, "一条主线：宏观政策管“信心”，市场监管管“信任”",
      "【约2分钟】这页是全课的“主线图”。消费可以粗略写成“信心×信任”。信心，回答的是“愿不愿意花、敢不敢花”：取决于收入、就业、资产、社保和对未来的预期——这主要是宏观政策和民生政策的主场。信任，回答的是“放不放心在这里花”：取决于产品质量、价格是否公道、出了问题能不能维权、平台规则是否公平——这恰恰是市场监管的主场。\n为什么用乘号不用加号？因为任何一项趋近于零，消费都起不来：老百姓再有钱，买到的东西不放心，也不会花；市场再规范，收入预期不好，也花不动。所以市场监管不是旁观者，而是消费“乘数”里的另一半。");
    // 两根支柱
    const pill = (x, head, sub, items, fill, ic) => {
      box(s, x, 1.65, 5.2, 3.65, fill);
      T(s, head, { x: x + 1.35, y: 1.85, w: 3.6, h: 0.55, fontSize: 26, bold: true, color: fill === K.tint ? K.teal : K.red, valign: "middle" });
      T(s, sub, { x: x + 1.35, y: 2.4, w: 3.7, h: 0.4, fontSize: 14, color: K.muted });
      items.forEach((it, i) => {
        const yy = 3.05 + i * 0.52;
        T(s, [{ text: it[0], options: { bold: true, fontSize: 15 } }, { text: "  " + it[1], options: { fontSize: 14, color: K.muted } }], { x: x + 0.45, y: yy, w: 4.6, h: 0.45, valign: "middle" });
      });
      return ic;
    };
    pill(MX, "信心", "愿不愿意花、敢不敢花", [["收入与就业", "有没有钱、稳不稳"], ["资产与财富", "房子、存款、股票"], ["社会保障", "养老、医疗、教育"], ["未来预期", "对明天的判断"]], K.tint);
    pill(7.53, "信任", "放不放心在这里花", [["质量与标准", "买到的是不是好东西"], ["价格与竞争", "价格公不公道"], ["维权与信用", "出了问题能不能解决"], ["平台与新业态", "规则公不公平"]], HX.redTint);
    await iconCircle(s, "FaChartLine", MX + 0.35, 1.85, 0.8, K.teal);
    await iconCircle(s, "FaHandshake", 7.88, 1.85, 0.8, K.red);
    // 中间乘号
    s.addShape(pres.shapes.OVAL, { x: 6.12, y: 3.0, w: 1.1, h: 1.1, fill: { color: K.jade }, line: { color: K.white, width: 2 } });
    T(s, "×", { x: 6.12, y: 3.0, w: 1.1, h: 1.1, fontSize: 40, bold: true, color: K.white, align: "center", valign: "middle" });
    // 底部结论
    T(s, [
      { text: "消费 ≈ 信心 × 信任", options: { bold: true, fontSize: 22, color: K.jade } },
      { text: "        宏观与民生政策的主场  ←  →  市场监管的主场", options: { fontSize: 15, color: K.muted } },
    ], { x: MX, y: 5.55, w: CW, h: 0.6, align: "center", valign: "middle" });
    T(s, "任何一项趋近于零，消费都起不来——市场监管不是消费问题的旁观者，而是“乘数”的另一半。", { x: MX, y: 6.15, w: CW, h: 0.4, fontSize: 14, color: K.ink, align: "center" });
  }

  // ---------------- 内容安排 ----------------
  {
    const s = content(0, "内容安排：五个部分，一条逻辑链",
      "【约1分钟】五个部分按照“是什么—为什么—别人怎么做—中央怎么部署—我们怎么干”的逻辑展开。前两部分偏宏观，后三部分落到市场监管。中间休息10分钟。");
    const parts = [
      ["观", "观势", "增长的“温差”从何而来", "GDP、价格、三大需求、社零走势", "约18分钟"],
      ["析", "析因", "消费为什么“提不起来”", "收入、财富、预期、供给、信任", "约32分钟"],
      ["鉴", "镜鉴", "消费社会是怎样炼成的", "日本通缩、美欧消费者保护", "约12分钟"],
      ["策", "谋策", "从“补商品”到“投资于人”", "政策脉络、以旧换新、补贴归宿", "约13分钟"],
      ["为", "作为", "市场监管如何为消费“增信任”", "五个抓手、浙江实践、基层清单", "约25分钟"],
    ];
    parts.forEach((p, i) => {
      const y = 1.6 + i * 0.98;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y, w: 0.78, h: 0.78, fill: { color: i < 2 ? K.teal : K.red }, line: { color: K.white, width: 0 }, rectRadius: 0.08 });
      T(s, p[0], { x: MX, y, w: 0.78, h: 0.78, fontSize: 28, bold: true, color: K.white, align: "center", valign: "middle" });
      T(s, [{ text: `${"一二三四五"[i]}、${p[1]} ｜ `, options: { bold: true, color: i < 2 ? K.teal : K.red, fontSize: 20 } }, { text: p[2], options: { bold: true, fontSize: 20 } }], { x: 1.65, y, w: 7.2, h: 0.45, valign: "middle" });
      T(s, p[3], { x: 1.65, y: y + 0.45, w: 7.2, h: 0.33, fontSize: 14, color: K.muted });
      T(s, p[4], { x: 10.9, y, w: 1.83, h: 0.78, fontSize: 15, color: K.muted, align: "right", valign: "middle" });
    });
  }

  // ===================== 第一部分 观势 =====================
  section("一、观势");
  divider(1, "增长的“温差”从何而来", "GDP增长5%左右，为什么不少企业和家庭“体感偏冷”？",
    "【过渡，约30秒】第一部分“观势”。我们经常听到一种说法：数据挺好，体感不好。这个“温差”是真实存在的，而且有明确的经济学解释。我们用五张图把它讲清楚。");

  {
    const s = content(1, "增速换挡：中国经济进入“5%左右”的中速增长阶段",
      "【约3分钟】这张图是2012年以来的GDP实际增速。三个看点：第一，从7%—8%逐步换挡到5%左右，这是人口结构、资本边际回报和技术追赶空间变化的结果，是规律，不是失误。第二，2020—2022年疫情扰动大，2021年的8.4%是低基数反弹，不能看作趋势。第三，今年上半年4.7%，一季度5.0%、二季度4.3%，二季度放缓主要来自内需，尤其是消费和投资。\n【提示】国家统计局通常在10月中下旬发布前三季度数据，讲课前请把最新一期数字补上。");
    chartCap(s, "图1  GDP实际增速（%）");
    const colors = D.gdp.map((v, i) => (i === D.gdp.length - 1 ? HX.red : HX.teal));
    s.addChart(pres.charts.BAR, [{ name: "GDP增速", labels: D.gdpYears, values: D.gdp }], chartOpts({
      x: MX, y: 1.9, w: 7.95, h: 4.65, barDir: "col", chartColors: colors, showValue: true, dataLabelPosition: "outEnd",
      dataLabelFormatCode: "0.0", valAxisMinVal: 0, valAxisMaxVal: 10, valAxisMajorUnit: 2, catAxisLabelFontSize: 10,
    }));
    insight(s, [
      "增速从“8%时代”换挡到“5%左右”，是潜在增长率下行的**规律性变化**",
      "2026年目标为**4.5%—5%**；上半年4.7%，其中二季度4.3%，**内需走弱**是主因",
      "看增速更要看结构：同样5%，靠外需和投资撑起来的增长，与靠消费撑起来的增长，**质量不同**",
    ], { reg: "换挡期增长更依赖**内需与效率**：降低制度性交易成本，本身就是稳增长政策。", regH: 1.25 });
    source(s, "国家统计局。2022年及以前、2024年为最终核实数，2023年为初步核算数，2026年为上半年同比。");
  }

  {
    const s = content(1, "体感温差之一：名义增速连续三年低于实际增速",
      "【约3分钟】这是理解“体感温差”最重要的一张图。蓝色是名义增速，绿色是实际增速。正常年份名义高于实际，因为有温和通胀。但2023—2025年连续三年名义低于实际，也就是GDP平减指数为负，物价总水平在下降。\n为什么这很重要？因为企业的营业收入、利润，政府的税收，职工的工资，全都是按名义值计算的。实际增长5%，名义只有4%，意味着“量在增、价在跌”——企业卖得更多却赚得更少，这就是“体感温差”的第一个来源。\n今年情况开始变化：PPI转正，平减指数降幅收窄。但要注意，这次价格回升更多来自上游成本推动，而不是需求拉动，我们后面会讲到。");
    chartCap(s, "图2  名义与实际GDP增速（%）");
    s.addChart(pres.charts.BAR, [
      { name: "名义增速", labels: D.nrYears, values: D.nominal },
      { name: "实际增速", labels: D.nrYears, values: D.real },
    ], chartOpts({
      x: MX, y: 1.9, w: 7.95, h: 4.65, barDir: "col", barGrouping: "clustered", chartColors: [HX.blue, HX.teal],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", showLegend: true, valAxisMinVal: 0, valAxisMaxVal: 15, valAxisMajorUnit: 3,
    }));
    insight(s, [
      "营收、利润、税收、工资都按**名义值**计价——名义增速才是“体感增速”",
      "2023—2025年名义增速约**4.6%、4.2%、4.0%**，均低于实际增速：量增价跌",
      "2026年PPI转正，名义增速有望回到实际之上，但这轮回升主要是**成本推动**",
    ], { reg: "价格总水平是监管的“温度计”：通缩期重在治理低价无序竞争，回升期还要防成本推动型涨价。", regH: 1.5 });
    source(s, "国家统计局，作者据现价GDP计算名义增速（2024年用最终核实数，2023年基数含第五次经济普查修订）。");
  }

  {
    const s = content(1, "增长靠谁拉动：外需“补位”明显，消费拉动仍不够强",
      "【约3分钟】这张图把每年GDP增长拆成三块：消费、投资、净出口，单位是“拉动了几个百分点”。看2024、2025年：净出口分别拉动1.5和1.6个百分点，2025年对增长的贡献率接近三分之一，是近年高位。也就是说，最近两年外需在给内需“补位”。\n问题是外需补位不可持续：一是外部关税和地缘不确定性在上升；二是长期靠外需，会加剧贸易摩擦。所以中央反复强调“坚持内需主导”。消费2025年拉动2.6个百分点，贡献率52%；“十四五”时期平均贡献率58.8%——但相对于我们的经济体量和发展阶段，消费这台“发动机”还不够强。");
    chartCap(s, "图3  三大需求对GDP增长的拉动（百分点）");
    s.addChart(pres.charts.BAR, [
      { name: "最终消费支出", labels: D.engYears, values: D.engC },
      { name: "资本形成总额", labels: D.engYears, values: D.engI },
      { name: "货物和服务净出口", labels: D.engYears, values: D.engX },
    ], chartOpts({
      x: MX, y: 1.9, w: 7.95, h: 4.65, barDir: "col", barGrouping: "stacked", chartColors: [HX.teal, HX.gold, HX.blue],
      showValue: true, dataLabelPosition: "ctr", dataLabelFormatCode: "0.0", dataLabelColor: HX.white, showLegend: true, valAxisMinVal: -2, valAxisMaxVal: 10, valAxisMajorUnit: 2, barGapWidthPct: 70,
    }));
    insight(s, [
      "2025年：消费拉动**2.6**、投资**0.8**、净出口**1.6**个百分点；净出口贡献率**32.7%**",
      "外需“补位”不可持续：关税与地缘不确定性上升，必须**坚持内需主导**",
      "“十四五”消费平均贡献率58.8%，比“十三五”提高10个百分点，但**消费率仍低**",
    ], { reg: "外需不确定时，**国内大市场**是最大底气：破除地方保护、建设统一大市场是监管的直接贡献。", regH: 1.5 });
    source(s, "国家统计局（各年初步核算口径）。");
  }

  {
    const s = content(1, "消费“脉冲”在退坡：社零增速从6.4%回落到0.4%",
      "【约4分钟】这是社会消费品零售总额的月度增速。绿色是2025年，红色是2026年。2025年5月达到6.4%的阶段高点，此后几乎逐月回落；今年5月一度负增长，8月只有0.4%。\n为什么？最主要的原因是以旧换新等耐用品补贴的“需求前置”：本来明年后年才换的车、换的家电，提前到2025年买了。经济学上叫“跨期替代”。补贴力度边际减弱后，就出现“回吐”，今年8月汽车零售下降18.5%，是最大拖累。\n但也有亮点：通讯器材增长27.3%（数码产品补贴），1—8月服务零售增长4.9%，明显快于商品零售的1.0%。\n【监管提示】补贴退坡期往往是价格乱象高发期，先涨后降、虚假促销要重点盯防。");
    chartCap(s, "图4  社会消费品零售总额当月同比（%）");
    const colors = D.rs.map((v, i) => (i < 11 ? HX.teal : HX.red));
    s.addChart(pres.charts.BAR, [{ name: "社零增速", labels: D.rsLabels, values: D.rs }], chartOpts({
      x: MX, y: 1.9, w: 7.95, h: 4.65, barDir: "col", chartColors: colors, invertedColors: colors, showValue: true, dataLabelPosition: "outEnd",
      dataLabelFormatCode: "0.0", dataLabelFontSize: 10, valAxisMinVal: -2, valAxisMaxVal: 8, valAxisMajorUnit: 2, catAxisLabelFontSize: 10, barGapWidthPct: 40,
    }));
    T(s, "■ 2025年   ", { x: 5.6, y: 1.55, w: 1.4, h: 0.32, fontSize: 12, color: K.teal, bold: true });
    T(s, "■ 2026年", { x: 7.0, y: 1.55, w: 1.4, h: 0.32, fontSize: 12, color: K.red, bold: true });
    insight(s, [
      "2026年5月一度转负，8月仅**0.4%**；阶段高点是2025年5月的**6.4%**",
      "结构分化：8月通讯器材**+27.3%**，汽车**-18.5%**；1—8月服务零售+4.9%，商品零售仅+1.0%",
      "耐用品补贴存在**需求前置**：今天的补贴，部分借用了明天的需求",
    ], { reg: "补贴退坡期是**价格乱象**高发期：先涨后降、虚假折扣、以旧充新需重点盯防。", regH: 1.4 });
    source(s, "国家统计局。2025年起按第五次经济普查修订后的可比口径计算增速。");
  }

  {
    const s = content(1, "价格信号在反转：上游转涨，终端仍弱",
      "【约3分钟】看价格。红线是PPI（工业生产者出厂价格），绿线是CPI（居民消费价格）。PPI在今年3月由负转正，结束了连续41个月的下降，6月涨到4.1%，8月3.8%。原因主要是国际大宗商品、资源品供给偏紧、电子产业链涨价，以及“反内卷”带来的部分行业价格企稳。\n但CPI只有1%左右，核心CPI也是1%左右。上游涨、下游涨不动——这说明终端需求仍然偏弱，成本传导不下去。谁最难受？夹在中间的中下游制造企业、批发零售、个体工商户：进价涨了，售价不敢涨，利润被挤压。\n这对监管意味着“两线作战”：上游要防囤积居奇、哄抬价格、串通涨价；下游仍要防低价倾销和价格欺诈。");
    chartCap(s, "图5  CPI与PPI当月同比（%）");
    s.addChart(pres.charts.LINE, [
      { name: "PPI（工业生产者出厂价格）", labels: D.pLabels, values: D.ppi },
      { name: "CPI（居民消费价格）", labels: D.pLabels, values: D.cpi },
    ], chartOpts({
      x: MX, y: 1.9, w: 7.95, h: 4.65, chartColors: [HX.red, HX.teal], lineSize: 2.25, lineDataSymbol: "circle", lineDataSymbolSize: 6,
      showLegend: true, valAxisMinVal: -5, valAxisMaxVal: 5, valAxisMajorUnit: 1, catAxisLabelFontSize: 10, valAxisLabelFormatCode: "0",
    }));
    insight(s, [
      "PPI于2026年3月**由负转正**，结束连续41个月下降；6月**4.1%**、8月3.8%",
      "CPI 1—8月大多在**0.5%—1.3%**，核心CPI约1%：成本向终端**传导受阻**",
      "中下游企业与个体户**两头受挤**：进价涨、售价难涨",
    ], { reg: "价格监管“两线作战”：上游防**囤积哄抬、串通涨价**；下游防**低价倾销、价格欺诈**。", regH: 1.5 });
    source(s, "国家统计局月度价格数据（2026年1月起CPI以2025年为基期）。");
  }

  {
    const s = content(1, "形势判断：主要矛盾是有效需求不足，最弱一环在居民消费",
      "【约3分钟，承上启下】把前面五张图归纳成四句话：总量稳、结构弱、价格转、预期弱。\n这和中央的判断是一致的：2024年中央经济工作会议明确指出“国内需求不足”；2025年中央经济工作会议把“坚持内需主导，建设强大国内市场”放在重点任务的第一位。\n所以，当前经济的主要矛盾是有效需求不足，而需求里最弱的一环，是居民消费。下面第二部分，我们就专门拆解：消费为什么提不起来。");
    const cards = [
      ["总量稳", "GDP增长4.7%", "增速处于目标区间，就业、物价总体可控", K.teal, "FaChartLine"],
      ["结构弱", "社零增长1.1%", "消费、房地产偏弱，增长较多依靠外需与政策", K.red, "FaBalanceScaleLeft"],
      ["价格转", "PPI由负转正", "上游转涨、终端偏弱，传导不畅挤压中下游", K.gold, "FaSearchDollar"],
      ["预期弱", "储蓄意愿上升", "居民“多存少花”，企业“多看少投”", K.blue, "FaRegLightbulb"],
    ];
    for (let i = 0; i < 4; i++) {
      const c = cards[i], x = MX + i * 3.08;
      box(s, x, 1.65, 2.85, 3.3, K.tint);
      await iconCircle(s, c[4], x + 0.3, 1.9, 0.72, c[3]);
      T(s, c[0], { x: x + 1.15, y: 1.95, w: 1.6, h: 0.6, fontSize: 24, bold: true, color: c[3], valign: "middle" });
      T(s, c[1], { x: x + 0.3, y: 2.85, w: 2.4, h: 0.45, fontSize: 17, bold: true });
      T(s, c[2], { x: x + 0.3, y: 3.4, w: 2.35, h: 1.3, fontSize: 14, color: K.muted, lineSpacingMultiple: 1.15 });
    }
    box(s, MX, 5.2, CW, 1.3, K.jade);
    T(s, [
      { text: "中央判断  ", options: { bold: true, color: K.goldLight, fontSize: 16 } },
      { text: "2024年中央经济工作会议：“国内需求不足”；2025年中央经济工作会议：“坚持内需主导，建设强大国内市场”。", options: { color: K.white, fontSize: 15, breakLine: true } },
      { text: "本讲判断  ", options: { bold: true, color: K.goldLight, fontSize: 16 } },
      { text: "主要矛盾是有效需求不足，最弱一环在居民消费——消费既是“果”（收入与预期），也是“因”（决定增长质量）。", options: { color: K.white, fontSize: 15 } },
    ], { x: MX + 0.35, y: 5.25, w: CW - 0.7, h: 1.2, valign: "middle", lineSpacingMultiple: 1.25 });
  }

  // ===================== 第二部分 析因 =====================
  section("二、析因");
  divider(2, "消费为什么“提不起来”", "中国人天生爱储蓄、不爱花钱吗？——高储蓄是制度与预期的结果，而非“国民性”。",
    "【过渡，约30秒】第二部分是今天的重点。常有人说中国人天生爱存钱，这个解释既不科学也没有政策含义。我们用一个框架把消费拆成四道门槛，逐一看数据。");

  {
    const s = content(2, "分析框架：消费要跨过“四道门槛”",
      "【约3分钟】经济学里的消费函数，最简单的版本只有收入一个变量。但现实中，一个家庭从“有钱”到“把钱花出去”，至少要跨过四道门槛：\n第一道，能消费——有没有钱，取决于收入份额、收入增速和收入分配；\n第二道，敢消费——敢不敢花，取决于资产（尤其是房子）、就业稳定性和社会保障，也就是预防性储蓄动机；\n第三道，愿消费——想不想买，取决于有没有好产品、好服务、好场景，甚至有没有时间；\n第四道，放心消费——放不放心，取决于质量、价格、维权和信用。\n前两道是“信心”问题，主要靠宏观和民生政策；后两道是“供给与信任”问题，恰恰是市场监管的主场。后面的数据，就按这四道门槛展开。");
    const steps = [
      ["能消费", "收入", "有没有钱", "收入份额 · 增速 · 分配", K.teal, "FaWallet", "收入实际增速 4.2%"],
      ["敢消费", "预期", "敢不敢花", "财富 · 就业 · 保障", K.blue, "FaShieldAlt", "平均消费倾向 68.0%"],
      ["愿消费", "供给", "想不想买", "产品 · 服务 · 场景 · 时间", K.gold, "FaShoppingBasket", "服务消费占比 46.1%"],
      ["放心消费", "信任", "放不放心", "质量 · 价格 · 维权 · 信用", K.red, "FaHandshake", "投诉举报 2646万件"],
    ];
    for (let i = 0; i < 4; i++) {
      const st = steps[i], x = MX + i * 3.08, y = 3.05 - i * 0.3, h = 2.75 + i * 0.3;
      box(s, x, y, 2.85, h, i < 2 ? K.tint : HX.redTint);
      await iconCircle(s, st[5], x + 0.25, y + 0.25, 0.68, st[4]);
      T(s, `${i + 1}  ${st[0]}`, { x: x + 1.05, y: y + 0.28, w: 1.75, h: 0.6, fontSize: 20, bold: true, color: st[4], valign: "middle" });
      T(s, `${st[1]}：${st[2]}`, { x: x + 0.25, y: y + 1.1, w: 2.45, h: 0.38, fontSize: 16, bold: true });
      T(s, st[3], { x: x + 0.25, y: y + 1.5, w: 2.45, h: 0.38, fontSize: 14, color: K.muted });
      s.addShape(pres.shapes.LINE, { x: x + 0.25, y: 5.12, w: 2.35, h: 0, line: { color: HX.line, width: 0.75 } });
      T(s, st[6], { x: x + 0.25, y: 5.2, w: 2.45, h: 0.45, fontSize: 13, bold: true, color: st[4], valign: "middle" });
    }
    // 归属标签
    T(s, "信心 —— 宏观与民生政策的主场", { x: MX, y: 1.5, w: 5.9, h: 0.4, fontSize: 15, bold: true, color: K.teal, align: "center" });
    T(s, "信任 —— 市场监管的主场", { x: 6.76, y: 1.5, w: 5.97, h: 0.4, fontSize: 15, bold: true, color: K.red, align: "center" });
    s.addShape(pres.shapes.LINE, { x: MX + 0.2, y: 1.95, w: 5.5, h: 0, line: { color: HX.teal, width: 1.25 } });
    s.addShape(pres.shapes.LINE, { x: 6.96, y: 1.95, w: 5.57, h: 0, line: { color: HX.red, width: 1.25 } });
    T(s, "消费函数：C = f（收入Y，财富W，预期E，保障S；供给质量Q，信任环境T）", { x: MX, y: 6.05, w: CW, h: 0.45, fontSize: 15, color: K.jade, bold: true, align: "center" });
  }

  {
    const s = content(2, "结构性短板：居民消费率比主要经济体低10—28个百分点",
      "【约3分钟】这是居民消费占GDP比重的国际比较。中国大约40%，美国68%，英国、印度六成左右，日本、德国、法国在五成以上，韩国接近五成。\n这里要讲两点严谨性：第一，中国的统计口径存在一定低估，比如自有住房的虚拟房租、政府提供给居民的实物社会转移。2023年国家统计局已经改革了自有住房服务核算方法，部分修正了这个低估。第二，即便学界做了各种调整，中国居民消费率仍比主要经济体低10个百分点以上，这个结构性短板是真实存在的。\n一个直观换算：居民消费率每提高1个百分点，按140万亿元GDP算，就是约1.4万亿元的消费。“十五五”规划要求“居民消费率明显提高”，空间在这里。");
    chartCap(s, "图6  居民消费支出占GDP比重：国际比较（%）");
    const labels = D.intl.map((d) => d[0]).reverse(), vals = D.intl.map((d) => d[1]).reverse();
    s.addChart(pres.charts.BAR, [{ name: "居民消费率", labels, values: vals }], chartOpts({
      x: MX, y: 1.9, w: 7.95, h: 4.65, barDir: "bar", chartColors: vals.map((v, i) => (labels[i] === "中国" ? HX.red : HX.context)),
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0", valAxisMinVal: 0, valAxisMaxVal: 80, valAxisMajorUnit: 20, catAxisLabelFontSize: 13, barGapWidthPct: 45,
    }));
    insight(s, [
      "中国约**40%**，主要经济体多在**48%—68%**",
      "口径有低估（虚拟房租等），2023年起已部分修正，但**差距仍超10个百分点**",
      "居民消费率每提高1个百分点 ≈ **1.4万亿元**消费",
      "“十五五”：**居民消费率明显提高**",
    ], { reg: "消费率主要靠收入分配与社保；监管的贡献在于减少“**不敢买、买不好**”造成的消费流失。", regH: 1.5 });
    source(s, "世界银行WDI（家庭及为家庭服务的非营利机构最终消费支出/GDP，2023年或最近可得年份，取整）；中国为国家统计局2024年数据。");
  }

  {
    const s = content(2, "能不能消费（一）：居民收入“蛋糕”偏小，财产性收入偏薄",
      "【约3分钟】第一道门槛：收入。左边是2025年全国居民人均可支配收入的结构：工资性收入占56.6%，经营净收入16.7%，转移净收入18.6%，财产净收入只有8%。\n三个要点。第一，把人均可支配收入乘以人口，大约61万亿元，只占GDP的43%左右；即便按资金流量表口径也只有六成上下，明显低于发达经济体。国民收入的“蛋糕”分给居民的那一块偏小。第二，财产净收入占比只有8%，而且2025年只增长1.6%，房租、存款利息、分红都在走弱。第三，中位数只有平均数的83.5%，城乡收入比2.31，说明收入分配还不均衡——而低收入群体恰恰是边际消费倾向最高的群体。");
    chartCap(s, "图7  2025年全国居民人均可支配收入构成（%）", MX, 1.55, 5.4);
    s.addChart(pres.charts.DOUGHNUT, [{ name: "收入构成", labels: ["工资性收入", "经营净收入", "转移净收入", "财产净收入"], values: [56.6, 16.7, 18.6, 8.0] }], chartOpts({
      x: MX, y: 1.9, w: 5.4, h: 4.65, chartColors: [HX.teal, HX.blue, HX.gold, HX.red], holeSize: 55, showPercent: false, showValue: true,
      dataLabelFormatCode: "0.0", dataLabelColor: HX.white, dataLabelFontSize: 12, showLegend: true, legendPos: "b", legendFontSize: 12, dataBorder: { pt: 2, color: "FFFFFF" },
    }));
    // 中间三个数
    const kv = [["43,377元", "人均可支配收入", "名义增长5.0%"], ["36,231元", "中位数", "仅为平均数的83.5%"], ["2.31", "城乡收入比", "比上年缩小0.03"]];
    kv.forEach((k, i) => {
      const y = 1.75 + i * 1.6;
      T(s, k[0], { x: 6.2, y, w: 2.5, h: 0.65, fontSize: 28, bold: true, color: i === 1 ? K.red : K.teal, valign: "bottom" });
      T(s, k[1], { x: 6.2, y: y + 0.68, w: 2.5, h: 0.32, fontSize: 14, bold: true });
      T(s, k[2], { x: 6.2, y: y + 1.0, w: 2.5, h: 0.32, fontSize: 12, color: K.muted });
    });
    insight(s, [
      "人均可支配收入×人口≈61万亿元，约占GDP的**43%**（资金流量表口径约六成）",
      "财产净收入仅占**8.0%**，2025年仅增长**1.6%**：房租、利息、分红同步走弱",
      "收入差距压低整体消费倾向：**低收入群体边际消费倾向最高**",
    ]);
    source(s, "国家统计局《2025年居民收入和消费支出情况》；住户调查口径，作者计算占GDP比重。");
  }

  {
    const s = content(2, "能不能消费（二）：收入增长慢于经济增长，消费增长又慢于收入增长",
      "【约3分钟】这张图有三条线：GDP实际增速、居民人均可支配收入实际增速、人均消费支出实际增速。\n疫情三年波动很大，2021年、2023年消费的高增长是低基数反弹，不必过度解读。关键看最右边：今年上半年，GDP增长4.7%，收入增长4.2%，消费只增长2.7%——形成了“两级落差”：收入跑输GDP，消费又跑输收入。\n这告诉我们：问题不只在“钱”，还在“心”。即使收入在增长，如果家庭对未来没有把握，增加的收入也会更多地变成储蓄。这就自然引出第二道门槛——敢不敢消费。");
    chartCap(s, "图8  GDP、居民收入与消费的实际增速（%）");
    s.addChart(pres.charts.LINE, [
      { name: "GDP", labels: D.incYears, values: D.incGDP },
      { name: "人均可支配收入", labels: D.incYears, values: D.incReal },
      { name: "人均消费支出", labels: D.incYears, values: D.conReal },
    ], chartOpts({
      x: MX, y: 1.9, w: 7.95, h: 4.65, chartColors: [HX.blue, HX.teal, HX.red], lineSize: 2.25, lineDataSymbol: "circle", lineDataSymbolSize: 6,
      showLegend: true, valAxisMinVal: -6, valAxisMaxVal: 14, valAxisMajorUnit: 2, catAxisLabelFontSize: 10, valAxisLabelFormatCode: "0",
    }));
    insight(s, [
      "2026年上半年：GDP **4.7%** ＞ 收入 **4.2%** ＞ 消费 **2.7%**，形成“两级落差”",
      "收入名义增长5.2%，消费名义增长仅3.7%：**多出来的收入变成了储蓄**",
      "2021、2023年消费高增长是**低基数反弹**，不代表趋势",
    ], { reg: "收入转化为消费时，**服务消费、预付式消费**最依赖信任——这是监管可减少“传导损耗”的环节。", regH: 1.5 });
    source(s, "国家统计局。2026年为上半年同比；人均指标为全国居民口径。");
  }

  {
    const s = content(2, "敢不敢消费（一）：消费倾向下降，预防性储蓄上升",
      "【约4分钟】左图是平均消费倾向，也就是每100元可支配收入里花掉多少。2019年是70.1元，疫情后一直没有回去，2025年是68.0元。差两个百分点看似不多，但做个换算：如果消费倾向回到2019年水平，2025年居民消费可以多出约1.3万亿元，相当于GDP的0.9%。\n右图是住户存款每年新增额。2019—2021年平均每年新增约10万亿元，2022年以后跳到每年15万亿元左右。这背后有预防性储蓄的增加，也有资产配置从房产、理财转向存款的“存款搬家”——要实事求是，不能把多存的钱全都看作“不敢花”。但方向是清楚的：家庭在“加固资产负债表”。\n经济学上，这叫预防性储蓄：未来越不确定，家庭越倾向存钱防身。");
    chartCap(s, "图9  平均消费倾向（%）", MX, 1.55, 3.85);
    s.addChart(pres.charts.LINE, [{ name: "平均消费倾向", labels: D.apcYears, values: D.apc }], chartOpts({
      x: MX, y: 1.9, w: 3.95, h: 4.65, chartColors: [HX.teal], lineSize: 2.25, lineDataSymbol: "circle", lineDataSymbolSize: 7,
      showValue: true, dataLabelPosition: "t", dataLabelFormatCode: "0.0", dataLabelFontSize: 10, valAxisMinVal: 62, valAxisMaxVal: 72, valAxisMajorUnit: 2, catAxisLabelFontSize: 10,
    }));
    chartCap(s, "图10  住户存款年度新增（万亿元）", 4.65, 1.55, 4.0);
    s.addChart(pres.charts.BAR, [{ name: "住户存款新增", labels: D.depYears, values: D.dep }], chartOpts({
      x: 4.65, y: 1.9, w: 3.95, h: 4.65, barDir: "col", chartColors: D.dep.map((v, i) => (i >= 3 ? HX.red : HX.context)),
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", dataLabelFontSize: 10, valAxisMinVal: 0, valAxisMaxVal: 20, valAxisMajorUnit: 5, catAxisLabelFontSize: 10,
    }));
    insight(s, [
      "若消费倾向回到2019年的70.1%，2025年可多出约**1.3万亿元**消费（≈GDP的0.9%）",
      "2022—2025年住户存款年均新增约**15.9万亿元**，比2019—2021年均多约5.5万亿",
      "多增存款中也有“**存款搬家**”（从房产、理财转入），需实事求是看待",
      "国际经验：储蓄率对**收入不确定性**与**社保覆盖**最敏感——这是宏观政策的主场",
    ]);
    source(s, "国家统计局（人均消费支出/人均可支配收入，作者计算）；中国人民银行金融统计数据。");
  }

  {
    const s = content(2, "敢不敢消费（二）：房地产调整带来“负财富效应”",
      "【约3分钟】第二道门槛里最重要的变量是房子。全国新建商品房销售面积从2021年的17.9亿平方米，降到2025年的8.8亿平方米，几乎腰斩；今年1—8月又下降了12.1%。房地产开发投资2025年下降17.2%。\n为什么房地产会影响消费？两个渠道。一是财富效应：城镇家庭资产约六成在住房上，房价预期转弱，家庭会觉得自己“变穷了”，于是增加储蓄、减少消费。二是关联消费：装修、家电、家具、建材都跟着走弱——今年建材家具类零售持续下降就是例证。\n对监管来说，房地产调整期，中介、装修、物业、预售等领域的消费纠纷也会增多，需要提前关注。");
    chartCap(s, "图11  全国新建商品房销售面积（亿平方米）");
    s.addChart(pres.charts.BAR, [{ name: "销售面积", labels: D.houseYears, values: D.house }], chartOpts({
      x: MX, y: 1.9, w: 7.95, h: 4.65, barDir: "col", chartColors: D.house.map((v, i) => (i === 2 ? HX.teal : i >= 3 ? HX.red : HX.context)),
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", valAxisMinVal: 0, valAxisMaxVal: 20, valAxisMajorUnit: 5,
    }));
    insight(s, [
      "销售面积较2021年峰值下降约**51%**；2026年1—8月再降**12.1%**",
      "2025年房地产开发投资**8.28万亿元**，下降17.2%；2026年1—8月降19.9%",
      "城镇家庭资产约**六成**在住房：房价预期转弱→“感觉变穷”→多存少花",
    ], { reg: "调整期**中介、装修、物业、预售**等领域消费纠纷易增多，需提前研判。", regH: 1.25 });
    source(s, "国家统计局；家庭资产结构参见中国人民银行《2019年中国城镇居民家庭资产负债情况调查》（住房占59.1%）。");
  }

  {
    const s = content(2, "敢不敢消费（三）：就业与保障决定“安全感”",
      "【约3分钟】第二道门槛的另外两个变量：就业和保障。\n青年就业压力比较突出：2025年8月，16—24岁（不含在校生）城镇失业率18.9%，是新口径发布以来的最高值；今年3月25—29岁失业率7.7%，也是该指标发布以来的最高。今年高校毕业生1270万人，再创新高。\n保障方面，城乡居民基础养老金最低标准2025年是每月143元，与职工养老金差距很大。农村老人、灵活就业人员的保障不足，自然要多存钱。\n这些不是市场监管能直接解决的，但要看到：中央已经在转向“投资于人”——今年政府工作报告首次提出制定实施城乡居民增收计划，并新增失能老年人养老服务消费补贴。");
    const tiles = [
      ["18.9%", "16—24岁青年失业率", "2025年8月，不含在校生口径发布以来最高", K.red],
      ["7.7%", "25—29岁失业率", "2026年3月，该指标发布以来最高", K.red],
      ["1270万", "2026届高校毕业生", "规模再创新高，同比增加48万", K.blue],
      ["143元", "城乡居民基础养老金", "每月最低标准（2025年）；2026年继续提高", K.gold],
    ];
    tiles.forEach((t, i) => {
      const x = MX + i * 3.08;
      box(s, x, 1.65, 2.85, 2.75, K.tint);
      T(s, t[0], { x: x + 0.25, y: 1.8, w: 2.45, h: 0.9, fontSize: 36, bold: true, color: t[3], valign: "bottom" });
      T(s, t[1], { x: x + 0.25, y: 2.8, w: 2.45, h: 0.62, fontSize: 15, bold: true, lineSpacingMultiple: 1.05 });
      T(s, t[2], { x: x + 0.25, y: 3.45, w: 2.45, h: 0.85, fontSize: 12, color: K.muted, lineSpacingMultiple: 1.1 });
    });
    box(s, MX, 4.65, CW, 1.85, HX.white, { line: { color: HX.line, width: 1 } });
    T(s, paras([
      "**预防性储蓄**：未来收入越不确定、保障越薄弱，家庭越倾向“存钱防身”——这是理性选择，而非“不爱花钱”",
      "**政策转向“投资于人”**：2026年政府工作报告首提制定实施**城乡居民增收计划**，新增失能老年人养老服务消费补贴",
      "**与监管的连接**：经营主体活跃是吸纳就业的主渠道——个体工商户兴旺，青年和灵活就业才有“饭碗”（见第五部分）",
    ], { fontSize: 14 }, { space: 6 }), { x: MX + 0.3, y: 4.8, w: CW - 0.6, h: 1.6, valign: "middle", lineSpacingMultiple: 1.1 });
    source(s, "国家统计局；教育部；人力资源社会保障部；2026年政府工作报告。");
  }

  {
    const s = content(2, "愿不愿消费（一）：服务消费占比停在46%，“有需求、缺供给”",
      "【约3分钟】第三道门槛：供给。随着收入提高，消费会从“买东西”转向“买服务”——养老、托育、教育、医疗、文旅、体育。这是规律。\n我国居民服务性消费占比从2013年的39.7%升到2019年的45.9%，疫情中回落，2024、2025年都是46.1%，基本停滞。美国个人消费支出里服务占三分之二以上，国内一线城市也已接近六成。\n为什么停滞？很大程度上不是没有需求，而是供给跟不上：好的养老院排不上队，托育机构不放心，景区体验不好。服务消费有一个特点——它是“经验品”甚至“信任品”，买之前看不见质量，买之后也难判断。这就决定了：服务消费越要放开，标准、认证和信用评价就越重要。这正是市场监管的用武之地。");
    chartCap(s, "图12  居民人均服务性消费支出占人均消费支出比重（%）");
    s.addChart(pres.charts.BAR, [{ name: "服务消费占比", labels: D.svcYears, values: D.svc }], chartOpts({
      x: MX, y: 1.9, w: 7.95, h: 4.65, barDir: "col", chartColors: D.svc.map((v, i) => (i >= 6 ? HX.teal : HX.context)),
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", valAxisMinVal: 30, valAxisMaxVal: 50, valAxisMajorUnit: 5,
    }));
    insight(s, [
      "2025年人均服务性消费**13,602元**，占比**46.1%**，与上年持平",
      "美国服务消费占个人消费支出**三分之二以上**；国内一线城市已接近六成",
      "痛点在供给：养老、托育、文旅——**有需求、缺供给、缺放心**",
    ], { reg: "服务多为“**信任品**”，质量事前难观察：标准、认证、信用评价越健全，服务消费越能放开。", regH: 1.5 });
    source(s, "国家统计局（2013年起城乡一体化住户调查口径）；美国经济分析局（BEA）。");
  }

  {
    const s = content(2, "愿不愿消费（二）：钱花在哪里——发展享受型消费增长最快",
      "【约2分钟】左图是2025年人均消费支出的结构：食品烟酒和居住两项合计占51%，仍是“半壁江山”。右图是各项的增速：增长最快的是其他用品及服务、教育文化娱乐、交通通信、生活用品及服务，都在7%以上；而医疗保健只增长1%，居住只增长2.1%。\n这个结构告诉我们，消费升级的方向是“发展型、享受型、服务型”。这些领域恰恰是新业态最多、标准最不完善、投诉增长最快的领域。监管跟不上，升级就会卡壳。");
    chartCap(s, "图13  2025年人均消费支出构成（%）", MX, 1.55, 5.9);
    s.addChart(pres.charts.BAR, [{ name: "占比", labels: D.compCats.slice().reverse(), values: D.compShare.slice().reverse() }], chartOpts({
      x: MX, y: 1.9, w: 5.9, h: 4.65, barDir: "bar", chartColors: [HX.teal], showValue: true, dataLabelPosition: "outEnd",
      dataLabelFormatCode: "0.0", valAxisMinVal: 0, valAxisMaxVal: 35, valAxisMajorUnit: 10, catAxisLabelFontSize: 12, barGapWidthPct: 40,
    }));
    const order = D.compGrowth.map((g, i) => [D.compCats[i], g]).sort((a, b) => a[1] - b[1]);
    chartCap(s, "图14  2025年各类消费支出名义增速（%）", 6.85, 1.55, 5.9);
    s.addChart(pres.charts.BAR, [{ name: "增速", labels: order.map((o) => o[0]), values: order.map((o) => o[1]) }], chartOpts({
      x: 6.85, y: 1.9, w: 5.88, h: 4.65, barDir: "bar", chartColors: order.map((o) => (o[1] >= 7 ? HX.red : HX.context)),
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", valAxisMinVal: 0, valAxisMaxVal: 14, valAxisMajorUnit: 2, catAxisLabelFontSize: 12, barGapWidthPct: 40,
    }));
    source(s, "国家统计局《2025年居民收入和消费支出情况》。人均消费支出29,476元。");
  }

  {
    const s = content(2, "放不放心消费：一年2646万件投诉举报，是“信任赤字”的晴雨表",
      "【约4分钟】第四道门槛：信任。这是我们自己的数据。2025年全国市场监管部门受理投诉、举报、咨询4386.6万件，其中投诉举报2646万件，同比增长9.8%，为消费者挽回经济损失43.5亿元。\n几个结构性特点：网购投诉举报1506.7万件，占56.9%；售后服务投诉连续三年居首；合同类投诉增长40.3%，集中在退款难、霸王条款；智能设备投诉增长26.6%。\n讲一个经济学的经典理论：阿克洛夫的“柠檬市场”。当买方无法辨别质量时，他只愿意按平均质量出价，好产品卖不上价就退出了，最后市场上只剩次品，整个市场萎缩。所以，信任缺失相当于对消费征收了一笔“隐性税”——这笔税，市场监管可以帮老百姓“减免”。\n【强调】投诉量上升不一定是坏事，可能是维权渠道更畅通了；关键看重复投诉和同类投诉是否减少。");
    chartCap(s, "图15  2025年12315受理构成（万件）", MX, 1.55, 4.6);
    s.addChart(pres.charts.DOUGHNUT, [{ name: "受理构成", labels: ["投诉", "咨询", "举报"], values: [2036.6, 1740.6, 609.4] }], chartOpts({
      x: MX, y: 1.9, w: 4.6, h: 4.65, chartColors: [HX.red, HX.teal, HX.blue], holeSize: 55, showValue: true, dataLabelFormatCode: "0.0",
      dataLabelColor: HX.white, dataLabelFontSize: 12, showLegend: true, legendPos: "b", legendFontSize: 12, dataBorder: { pt: 2, color: "FFFFFF" },
    }));
    T(s, [{ text: "4386.6", options: { fontSize: 22, bold: true, color: K.jade, breakLine: true } }, { text: "万件", options: { fontSize: 12, color: K.muted } }],
      { x: MX + 1.5, y: 3.35, w: 1.6, h: 0.9, align: "center", valign: "middle" });
    const kv = [["2646万件", "投诉举报，同比+9.8%"], ["56.9%", "网购投诉举报占比"], ["+40.3%", "合同类投诉（退款难、霸王条款）"], ["43.5亿元", "为消费者挽回经济损失"]];
    kv.forEach((k, i) => {
      const y = 1.7 + i * 1.18;
      T(s, k[0], { x: 5.45, y, w: 3.2, h: 0.6, fontSize: 26, bold: true, color: i === 3 ? K.teal : K.red, valign: "bottom" });
      T(s, k[1], { x: 5.45, y: y + 0.62, w: 3.2, h: 0.4, fontSize: 13, color: K.muted });
    });
    insight(s, [
      "售后服务投诉**连续三年居首**；智能设备投诉**+26.6%**",
      "**柠檬市场**（阿克洛夫）：质量信息不对称会让好产品退出、市场萎缩",
      "信任缺失＝对消费征收的“**隐性税**”；降低维权成本＝给消费“减税”",
    ], { head: "看点 · 经济学", reg: "看投诉不能只看总量，更要看**重复投诉、同类投诉**是否下降。", regH: 1.25 });
    source(s, "国家市场监督管理总局2025年全国消费者投诉举报数据（2026年3月15日发布）。");
  }

  {
    const s = content(2, "小结：市场监管不是提振消费的“主药”，却是让政策见效的“药引子”",
      "【约3分钟，第二部分总结】把四道门槛、主要症结、政策工具和主责部门放在一张表里。前两道门槛——收入和预期——是消费不足的主因，主要靠收入分配、社保、房地产、就业等宏观和民生政策，市场监管不是主角，我们要有这个清醒认识，不能“包打天下”，更不能夸大自己的作用。\n但后两道门槛——供给和信任——市场监管是主责部门。我用中医的“药引子”来比喻：药引子本身不是主药，但它决定药效能不能到达病灶。补贴发下去了，如果商家先涨价后打折，补贴就被吃掉了；收入增加了，如果买东西不放心、维权太难，钱就还是存在银行里。监管守住信任，其他政策才能转化为真实的消费。\n【过渡】上半场到这里。休息10分钟，下半场我们看国际经验，再落到市场监管怎么干。");
    const rows = [
      ["门槛", "主要症结", "关键工具", "主责"],
      ["能消费", "居民收入份额偏低、增速放缓、分配不均", "增收计划、提低扩中、财产性收入", "宏观·财政·人社"],
      ["敢消费", "房地产负财富效应、就业与保障不确定", "稳楼市、稳就业、提高养老金与医保", "宏观·住建·人社"],
      ["愿消费", "服务供给不足、标准缺位、准入限制", "放宽准入、标准引领、清理不合理限制", "发改·商务·市场监管"],
      ["放心消费", "质量隐患、价格乱象、维权难、内卷式竞争", "质量监管、公平竞争、消费维权、信用", "市场监管（主责）"],
    ];
    const tbl = rows.map((r, ri) => r.map((c, ci) => ({
      text: c,
      options: {
        bold: ri === 0 || ci === 0, fontSize: ri === 0 ? 14 : 14, fontFace: FONT,
        color: ri === 0 ? HX.white : ci === 3 && ri >= 3 ? HX.red : HX.ink,
        fill: { color: ri === 0 ? HX.jade : ri >= 3 ? HX.redTint : ri % 2 ? HX.tint : HX.white },
        valign: "middle", margin: [4, 8, 4, 8],
      },
    })));
    s.addTable(tbl, { x: MX, y: 1.65, w: CW, colW: [1.6, 4.1, 4.13, 2.3], rowH: [0.5, 0.72, 0.72, 0.72, 0.72], border: { type: "solid", pt: 1, color: "FFFFFF" } });
    box(s, MX, 5.6, CW, 0.95, K.jade);
    T(s, [
      { text: "“药引子”  ", options: { bold: true, color: K.goldLight, fontSize: 16 } },
      { text: "本身不是主药，却决定药效能否到达病灶——监管守住信任，补贴、增收等政策才能转化为真实消费。", options: { color: K.white, fontSize: 15 } },
    ], { x: MX + 0.35, y: 5.6, w: CW - 0.7, h: 0.95, valign: "middle" });
  }

  // 中场休息
  {
    const s = pres.addSlide({ masterName: "DARK", sectionTitle: CURRENT_SECTION });
    T(s, "中场休息", { x: 1.0, y: 2.1, w: 11.3, h: 1.1, fontSize: 54, bold: true, color: K.white, align: "center", valign: "middle" });
    T(s, "10 分钟", { x: 1.0, y: 3.2, w: 11.3, h: 0.6, fontSize: 26, color: K.goldLight, align: "center" });
    T(s, "下半场：别的国家走过什么路？中央怎么部署？市场监管怎么干？", { x: 1.0, y: 4.4, w: 11.3, h: 0.5, fontSize: 18, color: HX.darkText, align: "center" });
    s.addNotes("【休息10分钟】可在休息时收集问题，下半场开头先回应1—2个。");
  }

  // ===================== 第三部分 镜鉴 =====================
  section("三、镜鉴");
  divider(3, "消费社会是怎样炼成的", "高消费率是“国民性”，还是制度与政策的结果？",
    "【过渡，约30秒】下半场先看国际经验。我选两个参照：一个是反面教材——日本泡沫破裂后的通缩；一个是正面经验——美欧的消费者保护制度。");

  {
    const s = content(3, "日本镜鉴：泡沫破裂后，“通缩心态”持续了二十年",
      "【约5分钟】这是日本1990—2015年的CPI涨幅，红色柱是负增长的年份。1990年前后资产泡沫破裂，股价、地价持续下跌，企业和家庭都转向“还债优先”，这就是辜朝明所说的“资产负债表衰退”。\n之后近二十年，日本物价基本零增长甚至负增长。企业为了保份额，降价、压工资、减少正式雇佣；家庭收入不涨，只能买更便宜的东西；企业又进一步降价——形成“低价—低薪—低需求”的自我强化循环。日本人把这叫“通缩心态”。2013年以后安倍经济学才逐步扭转，代价极高。\n对我们的三点启示：第一，通缩预期一旦形成，扭转成本极高，要早干预；第二，价格战不等于消费者福利，长期看会损害质量、创新和工资；第三，政策最终要落到居民收入上。\n【注意分寸】中日国情差异很大：我们城镇化还有空间、政策工具更多、产业升级更快。讲日本是为了“防”，不是说我们会“成为”日本。");
    chartCap(s, "图16  日本CPI同比涨幅（%），1990—2015年");
    s.addChart(pres.charts.BAR, [{ name: "日本CPI", labels: D.jpYears, values: D.jpCPI }], chartOpts({
      x: MX, y: 1.9, w: 7.95, h: 4.65, barDir: "col", chartColors: D.jpCPI.map((v) => (v < 0 ? HX.red : HX.context)), invertedColors: D.jpCPI.map((v) => (v < 0 ? HX.red : HX.context)),
      valAxisMinVal: -2, valAxisMaxVal: 4, valAxisMajorUnit: 1, catAxisLabelFontSize: 10, barGapWidthPct: 35, catAxisLabelFrequency: 2,
    }));
    insight(s, [
      "1990年泡沫破裂→**资产负债表衰退**：企业与家庭“还债优先”",
      "企业以降价、压薪保份额：**低价→低薪→低需求**自我强化",
      "启示：通缩心态扭转代价极高；**价格战≠消费者福利**；政策要落到居民收入",
    ], { reg: "警惕“内卷式”价格战演变为全行业**降质、降薪、降创新**。", regH: 1.25 });
    source(s, "日本总务省统计局（全国CPI综合指数，年度同比，取一位小数）。");
  }

  {
    const s = content(3, "美欧镜鉴：高消费率背后是一整套“信任基础设施”",
      "【约4分钟】美国家庭消费占GDP近七成，常被解释为“美国人爱花钱”。但如果看制度史，你会发现，高消费率背后是一百多年逐步建立的“信任基础设施”：1906年食品药品立法，1914年成立联邦贸易委员会，1962年肯尼迪提出消费者四项权利——安全、知情、选择、表达，1972年设立消费品安全委员会。欧洲则有1985年产品责任指令、1993年统一大市场、2011年的14天无理由撤回权、2022年针对大平台“守门人”的数字市场法。\n我们的制度建设起步晚、进步快：1993年消保法，2013年修订引入网购七日无理由退货，2018年组建市场监管总局，2025年新修订的反不正当竞争法施行，同年出台优化消费环境三年行动方案。\n核心逻辑：消费者保护把“试错成本”从消费者转移给了经营者，人们才敢于尝试新产品、新服务。");
    // 时间轴
    const top = [["1906", "美国《纯净食品药品法》"], ["1914", "美国联邦贸易委员会成立"], ["1962", "肯尼迪提出消费者“四项权利”"], ["1972", "美国消费品安全委员会设立"], ["1985", "欧共体《产品责任指令》"], ["1993", "欧盟统一大市场启动"], ["2011", "欧盟14天无理由撤回权"], ["2022", "欧盟《数字市场法》：平台“守门人”义务"]];
    const bot = [["1993", "《消费者权益保护法》"], ["2013", "消保法修订：网购七日无理由退货"], ["2018", "组建国家市场监督管理总局"], ["2022", "《反垄断法》修订"], ["2025", "新修订《反不正当竞争法》施行"], ["2025", "《优化消费环境三年行动方案》"]];
    // 两条独立时间轴：上为美欧，下为中国
    const row = (items, ax, col, fill, label) => {
      const x0 = 1.85, span = W - MX - x0, step = span / items.length;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: ax - 0.95, w: 1.05, h: 1.9, fill: { color: fill }, line: { color: fill, width: 0 }, rectRadius: 0.08 });
      T(s, label, { x: MX, y: ax - 0.95, w: 1.05, h: 1.9, fontSize: 20, bold: true, color: K.white, align: "center", valign: "middle" });
      s.addShape(pres.shapes.LINE, { x: x0, y: ax, w: span, h: 0, line: { color: HX.line, width: 1.5 } });
      items.forEach((t, i) => {
        const cx = x0 + step * i + step / 2;
        s.addShape(pres.shapes.OVAL, { x: cx - 0.09, y: ax - 0.09, w: 0.18, h: 0.18, fill: { color: col }, line: { color: K.white, width: 1.5 } });
        T(s, t[0], { x: cx - step / 2 + 0.04, y: ax - 0.55, w: step - 0.08, h: 0.38, fontSize: 15, bold: true, color: col, align: "center", valign: "bottom" });
        T(s, t[1], { x: cx - step / 2 + 0.04, y: ax + 0.18, w: step - 0.08, h: 0.8, fontSize: 12, align: "center", lineSpacingMultiple: 1.05 });
      });
    };
    row(top, 2.65, K.blue, K.blue, "美欧");
    row(bot, 4.75, K.red, K.red, "中国");
    box(s, MX, 5.95, CW, 0.62, K.tint);
    T(s, [{ text: "核心逻辑  ", options: { bold: true, color: K.teal } }, { text: "消费者保护把“试错成本”从消费者转给经营者——人们才敢尝试新产品、新服务。", options: {} }],
      { x: MX + 0.3, y: 5.95, w: CW - 0.6, h: 0.62, fontSize: 15, valign: "middle" });
    source(s, "各国立法与机构公开资料，作者整理。");
  }

  {
    const s = content(3, "镜鉴小结：三条规律",
      "【约3分钟】三条规律。第一，消费率是“慢变量”。它取决于收入分配、社会保障、城镇化这些结构性因素，短期补贴只能熨平波动，改变不了结构。所以对以旧换新这类政策，要有合理预期。\n第二，信任是“基础设施”。成熟消费社会都建立在一整套消费者保护和信用体系之上。它像公路一样，平时感觉不到，缺了就寸步难行。\n第三，竞争政策保护的是“竞争”，不是“竞争者”，更不是“价格”。这句话是反垄断经济学的经典表述。治理内卷，要打的是掠夺性定价、虚假低价，但绝不能滑向“限价保价”“协同涨价”——那就从反内卷变成了卡特尔。这一条，我们在第五部分还会展开。");
    const cards = [
      ["FaSeedling", "消费率是“慢变量”", "取决于收入分配、社会保障和城镇化。短期刺激只能“熨平波动”，不能“改变结构”——对补贴政策要有合理预期。", K.teal],
      ["FaLandmark", "信任是“基础设施”", "成熟消费社会都建在消费者保护与信用体系之上：平时感觉不到，缺了寸步难行。", K.blue],
      ["FaBalanceScale", "保护竞争而非竞争者", "既反垄断高价，也治理掠夺性低价；但绝不能滑向“限价保价”“协同涨价”。", K.red],
    ];
    for (let i = 0; i < 3; i++) {
      const c = cards[i], x = MX + i * 4.1;
      box(s, x, 1.65, 3.85, 3.5, i === 2 ? HX.redTint : K.tint);
      await iconCircle(s, c[0], x + 0.35, 1.9, 0.9, c[3]);
      T(s, `规律${"一二三"[i]}`, { x: x + 1.45, y: 1.9, w: 2.2, h: 0.35, fontSize: 14, bold: true, color: c[3] });
      T(s, c[1], { x: x + 1.45, y: 2.25, w: 2.3, h: 0.75, fontSize: 17, bold: true, valign: "middle" });
      T(s, c[2], { x: x + 0.35, y: 3.15, w: 3.15, h: 1.85, fontSize: 15, color: K.ink, lineSpacingMultiple: 1.3 });
    }
    box(s, MX, 5.45, CW, 1.05, K.jade);
    T(s, [
      { text: "对我们的含义  ", options: { bold: true, color: K.goldLight } },
      { text: "对补贴政策保持合理预期 ｜ 把“信任”当作基础设施来建设 ｜ 治理内卷守住竞争政策底线", options: { color: K.white } },
    ], { x: MX + 0.35, y: 5.45, w: CW - 0.7, h: 1.05, fontSize: 16, valign: "middle" });
  }

  // ===================== 第四部分 谋策 =====================
  section("四、谋策");
  divider(4, "从“补商品”到“投资于人”", "中央在怎么部署？市场监管在政策体系中的“坐标”在哪里？",
    "【过渡，约30秒】第四部分，看中央部署，并用经济学的眼光审视以旧换新这项最主要的促消费政策。");

  {
    const s = content(4, "政策脉络：扩内需从“逆周期调节”上升为“战略基点”",
      "【约4分钟】这是近两年扩内需政策的时间轴。可以看出三个转变：\n第一，从“物”到“人”。早期政策更多是补商品，现在越来越强调“投资于人”：育儿补贴、提高基础养老金、城乡居民增收计划。\n第二，从“商品”到“服务”。服务消费提质惠民行动、清理服务消费领域不合理限制。\n第三，从“补贴刺激”到“制度建设”。优化消费环境三年行动、整治内卷式竞争、修订反不正当竞争法和价格法——这些都是长效制度，也恰恰是市场监管牵头或深度参与的。\n特别提一下：《优化消费环境三年行动方案（2025—2027年）》是市场监管总局牵头，统筹30个部门、149项细化措施。也就是说，在提振消费这件大事里，市场监管是有“专属任务”的。");
    const ev = [
      ["2024.12", "中央经济工作会议", "“大力提振消费”列为首要任务"],
      ["2025.02", "五部门三年行动方案", "优化消费环境（总局牵头）"],
      ["2025.03", "提振消费专项行动方案", "中办、国办印发"],
      ["2025.07", "中央财经委第六次会议", "治理企业低价无序竞争"],
      ["2025.10", "“十五五”规划建议", "居民消费率明显提高"],
      ["2025.12", "中央经济工作会议", "坚持内需主导"],
      ["2026.03", "政府工作报告", "目标4.5%—5%；首提增收计划"],
      ["2026.07", "扩大消费“十五五”规划", "2030年社零60万亿元左右"],
    ];
    const ax = 3.75, step = CW / ev.length, bw = 2.3;
    s.addShape(pres.shapes.LINE, { x: MX, y: ax, w: CW, h: 0, line: { color: HX.jade, width: 2 } });
    ev.forEach((e, i) => {
      const cx = MX + step * i + step / 2, up = i % 2 === 0, hl = i === 1 || i === 3;
      const col = hl ? K.red : K.teal;
      s.addShape(pres.shapes.OVAL, { x: cx - 0.11, y: ax - 0.11, w: 0.22, h: 0.22, fill: { color: col }, line: { color: K.white, width: 1.5 } });
      T(s, e[0], { x: cx - 0.8, y: up ? ax - 0.5 : ax + 0.17, w: 1.6, h: 0.33, fontSize: 14, bold: true, color: col, align: "center" });
      const by = up ? 1.6 : 4.4;
      const bx = Math.min(Math.max(cx - bw / 2, MX), W - MX - bw);
      s.addShape(pres.shapes.LINE, { x: cx, y: up ? by + 1.3 : ax + 0.5, w: 0, h: up ? ax - 0.5 - by - 1.3 : by - ax - 0.5, line: { color: HX.line, width: 1 } });
      box(s, bx, by, bw, 1.3, hl ? HX.redTint : K.tint);
      T(s, e[1], { x: bx + 0.1, y: by + 0.12, w: bw - 0.2, h: 0.45, fontSize: 13, bold: true, align: "center", valign: "middle" });
      T(s, e[2], { x: bx + 0.12, y: by + 0.6, w: bw - 0.24, h: 0.6, fontSize: 13, color: K.muted, align: "center", valign: "top", lineSpacingMultiple: 1.05 });
    });
    box(s, MX, 5.95, CW, 0.62, K.jade);
    T(s, [
      { text: "三个转变  ", options: { bold: true, color: K.goldLight } },
      { text: "从“物”到“人”（投资于人） ｜ 从“商品”到“服务” ｜ 从“补贴刺激”到“制度建设”", options: { bold: true, color: K.white } },
    ], { x: MX + 0.3, y: 5.95, w: CW - 0.6, h: 0.62, fontSize: 15, align: "center", valign: "middle" });
    source(s, "中央经济工作会议、政府工作报告及相关文件公开报道，作者整理。红色节点为市场监管牵头或深度参与。");
  }

  {
    const s = content(4, "以旧换新：政策有效，但边际效应在递减",
      "【约4分钟】以旧换新是这两年力度最大的促消费政策。左图是中央安排的超长期特别国债资金：2024年1500亿元，2025年3000亿元，2026年2500亿元。右图是带动的相关商品销售额：2025年约2.6万亿元，今年1—8月1.55万亿元。\n政策是有效的：2025年汽车以旧换新超过1150万辆，家电超过1.29亿件。但要看到边际效应在递减：今年1—8月汽车以旧换新535万辆，明显少于去年同期；8月汽车零售额下降18.5%。\n这里我要强调一个概念上的区别：“带动销售额”不等于“新增消费”。2.6万亿元里，有一部分是消费者本来就会买的，有一部分是把明年的需求提前到今年，还有一部分可能挤出了其他消费。经济学研究补贴政策，最关心的是“净效应”。这就引出下一页——补贴的归宿。");
    chartCap(s, "图17  以旧换新中央资金（亿元）", MX, 1.55, 3.85);
    s.addChart(pres.charts.BAR, [{ name: "资金", labels: ["2024", "2025", "2026"], values: [1500, 3000, 2500] }], chartOpts({
      x: MX, y: 1.9, w: 3.95, h: 4.65, barDir: "col", chartColors: [HX.gold], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0",
      valAxisMinVal: 0, valAxisMaxVal: 3500, valAxisMajorUnit: 1000, barGapWidthPct: 60,
    }));
    chartCap(s, "图18  带动相关商品销售额（万亿元）", 4.65, 1.55, 4.0);
    s.addChart(pres.charts.BAR, [{ name: "销售额", labels: ["2024", "2025", "2026年1—8月"], values: [1.3, 2.6, 1.55] }], chartOpts({
      x: 4.65, y: 1.9, w: 3.95, h: 4.65, barDir: "col", chartColors: [HX.teal, HX.teal, HX.red], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0#",
      valAxisMinVal: 0, valAxisMaxVal: 3, valAxisMajorUnit: 1, barGapWidthPct: 60,
    }));
    insight(s, [
      "汽车以旧换新：2025年超**1150万辆**；2026年1—8月**535万辆**，8月汽车零售**-18.5%**",
      "“带动销售额”≠“新增消费”：要扣除**本来就会买**、**提前购买**和**挤出效应**",
      "政策在优化：家电聚焦一级能效、汽车向新能源倾斜、纳入智能眼镜与智能家居",
    ]);
    source(s, "国家发展改革委、商务部、财政部公开数据；2024年销售额由累计3.92万亿元减2025年2.6万亿元推算。");
  }

  {
    const s = content(4, "补贴的“归宿”：一元钱补贴，最终落进了谁的口袋",
      "【约5分钟，本部分核心】这是我自己的研究领域——消费刺激政策的经济学。我们问一个问题：政府花一元钱补贴消费，最终买到了什么？要剥四层：\n第一层，带动销售额，这是政策统计的口径；\n第二层，净新增支出——要扣掉三块：本来就会买的（经济学叫“无谓补贴”）、把明年的购买提前到今年的（跨期替代）、因为买了补贴品而少买其他东西的（挤出效应）；\n第三层，消费者实得——如果商家在补贴期先涨价再打折，或者虚标能效等级套取补贴，补贴就被商家拿走了。这在经济学上叫补贴的“市场归宿”，或者“税负归宿”的反面；\n第四层，社会福利——再考虑财政成本、公平和节能减排等外部性。\n关键在第三层：补贴能不能真正落到消费者口袋，市场监管是“守门人”。价格监测、能效标识核查、旧机回收真实性核查、售后投诉处置——每一项都直接决定补贴的效率。");
    const layers = [
      ["带动销售额", "政策统计口径：补贴相关商品的全部销售", 11.0, K.context],
      ["净新增支出", "扣除：本来就会买（无谓补贴）、提前购买（跨期替代）、挤出其他消费", 9.0, K.blue],
      ["消费者实得", "扣除：先涨后降、虚标套补、以次充好——被经营者“吃掉”的补贴", 7.0, K.red],
      ["社会福利", "再计入：财政成本、分配公平、节能减排等外部性", 5.0, K.teal],
    ];
    layers.forEach((l, i) => {
      const y = 1.65 + i * 1.18, w = l[2] * 0.66, x = MX + (7.3 - w) / 2;
      const fillHex = [HX.context, HX.blue, HX.red, HX.teal][i];
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 1.0, fill: { color: fillHex }, line: { color: fillHex, width: 0 }, rectRadius: 0.08 });
      T(s, `${i + 1}  ${l[0]}`, { x, y: y + 0.08, w, h: 0.42, fontSize: 17, bold: true, color: K.white, align: "center", valign: "middle" });
      T(s, l[1], { x: x + 0.15, y: y + 0.5, w: w - 0.3, h: 0.45, fontSize: 11, color: K.white, align: "center", valign: "middle" });
    });
    box(s, 8.25, 1.65, 4.48, 4.72, HX.redTint);
    T(s, "市场监管守护的是第三层", { x: 8.5, y: 1.85, w: 4.0, h: 0.42, fontSize: 17, bold: true, color: K.red });
    T(s, paras([
      "**价格监测**：补贴前后价格比对，防“先涨后降”",
      "**能效与认证核查**：防虚标等级、套取补贴",
      "**“旧”的真实性核查**：防虚假回收、骗补",
      "**售后与投诉处置**：保障“买得放心、修得方便”",
    ], { fontSize: 14 }, { space: 10 }), { x: 8.5, y: 2.4, w: 4.0, h: 2.6, lineSpacingMultiple: 1.1 });
    T(s, "补贴效率 = 净新增 × 归宿到消费者的比例。监管决定后一项。", { x: 8.5, y: 5.15, w: 4.0, h: 1.0, fontSize: 14, bold: true, color: K.jade, lineSpacingMultiple: 1.15 });
    source(s, "作者研究框架：消费刺激的家庭支出响应、市场归宿与福利。");
  }

  // ===================== 第五部分 作为 =====================
  section("五、作为");
  divider(5, "市场监管如何为消费“增信任”", "如何把“监管力”转化为“消费力”？",
    "【过渡，约30秒】最后一部分，回到我们自己。我把市场监管促消费的工作归纳为五个抓手，每个抓手对应消费的一道门槛。");

  {
    const s = content(5, "五个抓手：把“监管力”转化为“消费力”",
      "【约2分钟】五个抓手：放心消费、公平竞争、质量标准、经营主体、新业态。注意每个抓手都对应前面讲的一道门槛：放心消费和公平竞争对应“放心消费”；质量标准和新业态对应“愿消费”；经营主体对应“能消费”——这是市场监管影响居民收入最直接的渠道。下面逐一展开。");
    const levers = [
      ["FaUserShield", "放心消费", "降低“信任成本”", "门槛④ 放心消费", K.red],
      ["FaBalanceScale", "公平竞争", "让价格说真话", "门槛④ 放心消费", K.red],
      ["FaAward", "质量标准", "让好产品“出得来”", "门槛③ 愿消费", K.gold],
      ["FaStore", "经营主体", "活力→就业→收入", "门槛① 能消费", K.teal],
      ["FaMobileAlt", "新业态", "包容审慎促创新", "门槛③ 愿消费", K.gold],
    ];
    for (let i = 0; i < 5; i++) {
      const l = levers[i], x = MX + i * 2.46, w = 2.25;
      box(s, x, 1.7, w, 4.0, K.tint);
      await iconCircle(s, l[0], x + (w - 1.0) / 2, 1.95, 1.0, l[4]);
      T(s, `抓手${"一二三四五"[i]}`, { x, y: 3.15, w, h: 0.35, fontSize: 13, color: K.muted, align: "center" });
      T(s, l[1], { x, y: 3.5, w, h: 0.55, fontSize: 22, bold: true, align: "center", valign: "middle" });
      T(s, l[2], { x: x + 0.1, y: 4.1, w: w - 0.2, h: 0.4, fontSize: 15, align: "center" });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.25, y: 4.85, w: w - 0.5, h: 0.5, fill: { color: K.white }, line: { color: l[4], width: 1 }, rectRadius: 0.25 });
      T(s, l[3], { x: x + 0.25, y: 4.85, w: w - 0.5, h: 0.5, fontSize: 12, bold: true, color: l[4], align: "center", valign: "middle" });
    }
    T(s, "每一个抓手，都对应消费的一道门槛——这是市场监管服务扩内需的“施工图”。", { x: MX, y: 6.0, w: CW, h: 0.45, fontSize: 16, bold: true, color: K.jade, align: "center" });
  }

  {
    const s = content(5, "抓手一 · 放心消费：信任是成本最低的“刺激政策”",
      "【约4分钟】放心消费。左边三个数：2025年受理4386.6万件，挽回损失43.5亿元；到2030年，全国放心消费集聚区要达到5000个以上，打造“放心消费在中国”品牌。\n右边是四项具体工作。我想强调经济学逻辑：消费者买的不只是商品本身，还买了一份“出了问题能解决”的保障。无理由退货降低了“买错的成本”，预付式消费资金监管降低了“跑路的风险”，在线纠纷解决降低了“维权的时间成本”，信息公示激活了“用脚投票”的声誉机制。每降低一分维权成本，就提高一分消费意愿。\n为什么说信任是成本最低的刺激政策？发补贴要真金白银，而建立信任，主要靠规则和执法，财政成本很低，效果却是持久的。");
    const st = [["4386.6万件", "2025年受理投诉举报咨询", K.red], ["43.5亿元", "为消费者挽回经济损失", K.teal], ["5000个+", "2030年全国放心消费集聚区目标", K.blue]];
    st.forEach((t, i) => {
      const y = 1.65 + i * 1.6;
      box(s, MX, y, 4.0, 1.42, K.tint);
      T(s, t[0], { x: MX + 0.3, y: y + 0.12, w: 3.5, h: 0.75, fontSize: 32, bold: true, color: t[2], valign: "middle" });
      T(s, t[1], { x: MX + 0.3, y: y + 0.88, w: 3.5, h: 0.4, fontSize: 14, color: K.ink });
    });
    T(s, "四项重点工作", { x: 5.0, y: 1.65, w: 7.7, h: 0.4, fontSize: 17, bold: true, color: K.jade });
    const acts = [
      ["线下无理由退货", "从商场延伸到商圈、景区、夜市，降低“买错成本”"],
      ["预付式消费治理", "资金存管、余额公示、冷静期，治理“跑路”与“退费难”"],
      ["在线纠纷解决（ODR）", "引导企业与消费者直接和解，缩短维权链条"],
      ["信息公示与信用约束", "投诉数据公开、红黑榜，激活“用脚投票”的声誉机制"],
    ];
    acts.forEach((a, i) => {
      const y = 2.15 + i * 0.95;
      T(s, String(i + 1), { x: 5.0, y, w: 0.5, h: 0.75, fontSize: 28, bold: true, color: K.red, valign: "middle" });
      T(s, [{ text: a[0], options: { bold: true, fontSize: 16, breakLine: true } }, { text: a[1], options: { fontSize: 14, color: K.muted } }], { x: 5.55, y, w: 7.1, h: 0.8, valign: "middle" });
    });
    box(s, 5.0, 5.95, 7.73, 0.6, K.jade);
    T(s, "消费者买的不只是商品，还有“出了问题能解决”的保障。", { x: 5.2, y: 5.95, w: 7.4, h: 0.6, fontSize: 15, bold: true, color: K.white, valign: "middle" });
    source(s, "国家市场监督管理总局；《关于全面开展放心消费单元和集聚区培育的通知》。");
  }

  {
    const s = content(5, "浙江实践：从“放心消费在浙江”到“放心消费在中国”",
      "【约3分钟】讲讲我们浙江。左图是今年上半年浙江和全国的对比：GDP浙江5.7%，全国4.7%；社零浙江3.0%，全国1.3%，是全国的两倍多。消费表现好，背后有市场活力的原因，也有消费环境的原因。\n浙江从2017年开始实施“放心消费在浙江”行动，首批1.3万家企业承诺无理由退货，这一理念后来上升为全国的“放心消费在中国”品牌。近期的亮点包括：“放心贷”覆盖147个放心消费商圈、4.16万户商户，放贷超过1800亿元，入选全国优化消费环境五大典型案例；7个项目入选总局优化消费环境试点，数量全国第一；温州依托“浙食链”的食用农产品全链条数字化监管入选全国首批全链条监管创新案例。\n【建议】这里可以请在座各地市补充本地的鲜活案例，现场互动效果会更好。");
    chartCap(s, "图19  2026年上半年：浙江与全国（同比，%）", MX, 1.55, 5.2);
    s.addChart(pres.charts.BAR, [
      { name: "全国", labels: ["GDP", "社零总额"], values: [4.7, 1.3] },
      { name: "浙江", labels: ["GDP", "社零总额"], values: [5.7, 3.0] },
    ], chartOpts({
      x: MX, y: 1.9, w: 5.2, h: 4.65, barDir: "col", barGrouping: "clustered", chartColors: [HX.context, HX.teal], showValue: true,
      dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", showLegend: true, valAxisMinVal: 0, valAxisMaxVal: 7, valAxisMajorUnit: 1, catAxisLabelFontSize: 13, barGapWidthPct: 70,
    }));
    const items = [
      ["2017", "启动“放心消费在浙江”行动，首批1.3万家企业承诺无理由退货；理念上升为全国“放心消费在中国”品牌"],
      ["放心贷", "覆盖147个放心消费商圈、4.16万户商户，放贷超1800亿元；入选全国优化消费环境五大典型案例"],
      ["试点", "7个项目入选总局优化消费环境试点，数量居全国首位"],
      ["浙食链", "温州食用农产品全链条数字化监管，入选全国首批全链条监管创新案例"],
      ["2026", "培育“信用促消费”商圈100个以上，新增线下无理由退换货承诺单元超1万家"],
    ];
    items.forEach((it, i) => {
      const y = 1.6 + i * 0.98;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.2, y: y + 0.08, w: 1.15, h: 0.62, fill: { color: i === 1 ? K.red : K.teal }, line: { color: K.white, width: 0 }, rectRadius: 0.08 });
      T(s, it[0], { x: 6.2, y: y + 0.08, w: 1.15, h: 0.62, fontSize: 14, bold: true, color: K.white, align: "center", valign: "middle" });
      T(s, it[1], { x: 7.55, y, w: 5.18, h: 0.85, fontSize: 14, valign: "middle", lineSpacingMultiple: 1.1 });
    });
    source(s, "浙江省统计局；国家统计局；浙江省市场监督管理局公开资料。截至2025年底浙江经营主体1163.65万户。");
  }

  {
    const s = content(5, "抓手二 · 公平竞争：治理“内卷式”竞争，让价格说真话",
      "【约5分钟，本部分要讲透】“反内卷”是当前的热点，也最容易讲偏。我讲三层意思。\n第一，价格监管要“两线作战”。上游PPI涨了3.8%，要防囤积居奇、哄抬价格、串通涨价；下游CPI只有0.8%，要继续治理低价倾销、价格欺诈、虚假折扣、大数据杀熟。今年上半年总局部署了16项专项行动，全国查办不正当竞争案件11465件，正在推动价格法修改，完善低价倾销的认定规则和行业成本调查。\n第二，要划清边界。该治理的是：低于成本倾销排挤对手、平台强制商家“最低价”、滥用“仅退款”、靠偷工减料实现的“劣质低价”。不该做的是：行政定价、设“保护价”、为行业协会的“协同涨价”背书——后者本身就涉嫌违反反垄断法的垄断协议条款。也不能把正常的、靠效率实现的降价当作内卷。\n第三，判断标准只有一个：是否损害消费者的长期福利和市场的动态效率。监管的目标是让价格重新“说真话”，而不是让价格“往上走”。");
    // 左：两线作战
    T(s, "价格监管“两线作战”", { x: MX, y: 1.6, w: 5.6, h: 0.4, fontSize: 17, bold: true, color: K.jade });
    const fronts = [
      ["上游", "PPI +3.8%", "防囤积居奇、哄抬价格、串通涨价", K.gold, HX.tint],
      ["下游", "CPI +0.8%", "防低价倾销、价格欺诈、虚假折扣、大数据杀熟", K.red, HX.redTint],
    ];
    fronts.forEach((f, i) => {
      const y = 2.1 + i * 1.45;
      box(s, MX, y, 5.6, 1.3, f[4]);
      T(s, f[0], { x: MX + 0.25, y: y + 0.15, w: 1.0, h: 0.45, fontSize: 18, bold: true, color: f[3] });
      T(s, f[1], { x: MX + 1.2, y: y + 0.15, w: 2.5, h: 0.45, fontSize: 18, bold: true });
      T(s, f[2], { x: MX + 0.25, y: y + 0.68, w: 5.1, h: 0.5, fontSize: 14, color: K.ink });
    });
    T(s, paras([
      "2026年上半年：**16项**专项行动、**39项**预期成果",
      "全国查办不正当竞争案件**11,465件**",
      "推进价格法修改：完善**低价倾销认定**与**行业成本调查**",
    ], { fontSize: 14 }, { space: 6 }), { x: MX, y: 5.05, w: 5.6, h: 1.45, lineSpacingMultiple: 1.1 });
    // 右：边界
    T(s, "反内卷的边界", { x: 6.6, y: 1.6, w: 6.1, h: 0.4, fontSize: 17, bold: true, color: K.jade });
    box(s, 6.6, 2.1, 6.13, 2.0, K.tint);
    T(s, "✔ 该治理", { x: 6.85, y: 2.2, w: 3, h: 0.4, fontSize: 16, bold: true, color: K.teal });
    T(s, paras(["低于成本倾销、排挤竞争对手", "平台强制“最低价”、滥用“仅退款”规则", "偷工减料换来的“劣质低价”"], { fontSize: 14 }, { space: 4 }), { x: 6.85, y: 2.65, w: 5.7, h: 1.4 });
    box(s, 6.6, 4.25, 6.13, 2.0, HX.redTint);
    T(s, "✘ 不该做", { x: 6.85, y: 4.35, w: 3, h: 0.4, fontSize: 16, bold: true, color: K.red });
    T(s, paras(["行政定价、设“保护价”", "为行业“协同涨价”背书（涉嫌垄断协议）", "把效率型降价当作“内卷”打压"], { fontSize: 14 }, { space: 4 }), { x: 6.85, y: 4.8, w: 5.7, h: 1.4 });
    T(s, "判断标准：是否损害消费者长期福利与市场动态效率", { x: 6.6, y: 6.3, w: 6.13, h: 0.35, fontSize: 14, bold: true, color: K.jade, align: "center" });
    source(s, "国家统计局；国家市场监督管理总局2026年上半年新闻发布会。");
  }

  {
    const s = content(5, "抓手三 · 质量标准：以标准升级牵引消费升级",
      "【约4分钟】质量标准。经济学上，标准是“质量的语言”：它把消费者看不见的质量，变成看得见的标识，大幅降低搜寻成本。链条上有四环：标准定义好、认证证明好、检测抽查核验好、召回追责兜住底。\n两个现实信号：第一，以旧换新的规模——今年1—8月6类家电以旧换新9146万台、数码和智能产品购新1.08亿件，补贴聚焦一级能效，能效标识和认证的真实性就直接关系到财政资金的效率；第二，智能设备投诉增长26.6%，说明新产品的标准跟不上市场——智能手表、智能眼镜、AI玩具这些品类，标准缺位就是监管盲区。\n第三点是服务标准：养老、托育、家政、文旅这些服务消费，最缺的是“看得见的质量”。服务标准化和服务质量评价，是释放服务消费的前提。");
    const chain = [
      ["FaClipboardCheck", "标准", "定义“好”", "产品标准、服务标准、能效等级"],
      ["FaCertificate", "认证", "证明“好”", "绿色产品认证、能效标识核查"],
      ["FaSearchDollar", "检测抽查", "核验“好”", "“双随机”抽查、网络商品抽检"],
      ["FaUndo", "召回追责", "兜住“底”", "缺陷产品召回、质量追溯"],
    ];
    for (let i = 0; i < 4; i++) {
      const c = chain[i], x = MX + i * 3.08;
      box(s, x, 1.65, 2.75, 2.65, K.tint);
      await iconCircle(s, c[0], x + 0.25, 1.85, 0.7, i === 3 ? K.red : K.teal);
      T(s, c[1], { x: x + 1.05, y: 1.85, w: 1.6, h: 0.4, fontSize: 18, bold: true });
      T(s, c[2], { x: x + 1.05, y: 2.25, w: 1.6, h: 0.32, fontSize: 14, color: K.teal, bold: true });
      T(s, c[3], { x: x + 0.25, y: 2.8, w: 2.3, h: 1.3, fontSize: 14, color: K.ink, lineSpacingMultiple: 1.2 });
      if (i < 3) T(s, "→", { x: x + 2.75, y: 2.6, w: 0.33, h: 0.5, fontSize: 20, bold: true, color: K.muted, align: "center" });
    }
    const sig = [
      ["9146万台 / 1.08亿件", "2026年1—8月6类家电以旧换新 / 数码和智能产品购新", "补贴聚焦一级能效：能效标识真实性＝财政资金效率"],
      ["+26.6%", "2025年智能设备投诉举报增速（15.2万件）", "新产品标准滞后：标准缺位就是监管盲区"],
      ["服务标准", "养老、托育、家政、文旅", "服务消费最缺“看得见的质量”"],
    ];
    sig.forEach((g, i) => {
      const x = MX + i * 4.1;
      box(s, x, 4.55, 3.85, 1.95, HX.white, { line: { color: HX.line, width: 1 } });
      T(s, g[0], { x: x + 0.25, y: 4.65, w: 3.4, h: 0.55, fontSize: 21, bold: true, color: i === 1 ? K.red : K.jade, valign: "middle" });
      T(s, g[1], { x: x + 0.25, y: 5.2, w: 3.4, h: 0.55, fontSize: 12, color: K.muted, lineSpacingMultiple: 1.05 });
      T(s, g[2], { x: x + 0.25, y: 5.78, w: 3.4, h: 0.62, fontSize: 14, bold: true, lineSpacingMultiple: 1.05 });
    });
    source(s, "国家发展改革委；国家市场监督管理总局。");
  }

  {
    const s = content(5, "抓手四 · 经营主体：“活力—就业—收入—消费”的传导链",
      "【约4分钟】经营主体。这是市场监管影响“能消费”这道门槛最直接的渠道。2025年全国新设经营主体2574.5万户，其中个体工商户1619.4万户，占六成以上；在营活跃企业数量增长9.8%。浙江经营主体1163.65万户，平均不到6个常住人口就有一个经营主体。\n个体工商户有两个身份：它是就业的“蓄水池”——吸纳了大量青年、灵活就业人员和农民工；它也是消费的“毛细血管”——街边小店、夜市、社区服务，构成了消费最基层的场景。小店兴，则烟火旺。\n所以，登记便利化、涉企检查“无事不扰”、信用修复、个体工商户分型分类帮扶，这些看起来是“营商环境”工作，实质上是在打通“活力—就业—收入—消费”这条传导链。");
    chartCap(s, "图20  2025年新设经营主体构成（万户）", MX, 1.55, 4.6);
    s.addChart(pres.charts.DOUGHNUT, [{ name: "新设经营主体", labels: ["个体工商户", "企业"], values: [1619.4, 950.0] }], chartOpts({
      x: MX, y: 1.9, w: 4.6, h: 4.65, chartColors: [HX.teal, HX.blue], holeSize: 55, showValue: true, dataLabelFormatCode: "0.0",
      dataLabelColor: HX.white, dataLabelFontSize: 12, showLegend: true, legendPos: "b", legendFontSize: 12, dataBorder: { pt: 2, color: "FFFFFF" },
    }));
    T(s, [{ text: "2574.5", options: { fontSize: 22, bold: true, color: K.jade, breakLine: true } }, { text: "万户", options: { fontSize: 12, color: K.muted } }],
      { x: MX + 1.5, y: 3.35, w: 1.6, h: 0.9, align: "center", valign: "middle" });
    // 传导链
    const chain = ["准入便利", "主体活跃", "吸纳就业", "居民增收", "消费扩大"];
    chain.forEach((c, i) => {
      const x = 5.5 + i * 1.47;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.75, w: 1.22, h: 0.75, fill: { color: i === 4 ? K.red : K.teal }, line: { color: K.white, width: 0 }, rectRadius: 0.08 });
      T(s, c, { x, y: 1.75, w: 1.22, h: 0.75, fontSize: 14, bold: true, color: K.white, align: "center", valign: "middle" });
      if (i < 4) T(s, "→", { x: x + 1.22, y: 1.85, w: 0.25, h: 0.5, fontSize: 16, bold: true, color: K.muted, align: "center" });
    });
    box(s, 5.5, 2.75, 7.23, 3.75, K.tint);
    T(s, paras([
      "2025年在营活跃企业数量**+9.8%**，占全部企业比重提高4.7个百分点",
      "浙江经营主体**1163.65万户**，约每6名常住人口就有1个经营主体",
      "个体工商户是就业的“**蓄水池**”，也是消费的“**毛细血管**”——小店兴则烟火旺",
      "工作抓手：涉企检查“**无事不扰**”、“综合查一次”；个体工商户**分型分类帮扶**；信用修复“应修尽修”",
    ], { fontSize: 14 }, { space: 10 }), { x: 5.75, y: 2.95, w: 6.8, h: 3.4, lineSpacingMultiple: 1.12 });
    source(s, "国家市场监督管理总局（另有农民专业合作社等5.1万户）；长三角市场监管部门《长三角经营主体发展指数（2025版）》；浙江省统计局。");
  }

  {
    const s = content(5, "抓手五 · 新业态：包容审慎，压实平台“守门人”责任",
      "【约4分钟】新业态。两个数：今年一季度网上商品零售额占社零的24.8%，四分之一的消费在网上；2025年网购投诉举报占全部投诉举报的56.9%，一半以上的纠纷在网上。\n四类重点新业态：直播电商、即时零售和外卖、智能与AI消费、银发与健康消费。每一类都有其典型风险和对应的监管工具。比如即时零售和外卖，补贴大战带来的低价倾销和食品安全风险；银发消费，保健品虚假宣传和养老预付费风险。\n原则是“包容审慎”，但包容审慎不等于放任不管：规则先行、底线监管、科技赋能、平台共治。平台既是市场参与者，也是规则制定者，要压实它作为“守门人”的责任。");
    const stats = [["24.8%", "网上商品零售额占社零比重（2026年一季度）"], ["56.9%", "网购占全部投诉举报比重（2025年）"]];
    stats.forEach((t, i) => {
      const x = MX + i * 6.15;
      T(s, t[0], { x, y: 1.55, w: 1.9, h: 0.6, fontSize: 30, bold: true, color: i ? K.red : K.teal, valign: "middle" });
      T(s, t[1], { x: x + 2.0, y: 1.55, w: 3.9, h: 0.6, fontSize: 14, color: K.muted, valign: "middle" });
    });
    const grid = [
      ["FaMobileAlt", "直播电商", "虚假宣传、带货责任不清、售后难", "主播与平台责任界定、直播回放留存、重点抽查"],
      ["FaStore", "即时零售与外卖", "补贴大战、低价倾销、食品安全", "平台规则审查、约谈指导、“互联网+明厨亮灶”"],
      ["FaRegLightbulb", "智能与AI消费", "算法“杀熟”、数据安全、标准缺位", "算法透明、标准先行、个人信息保护"],
      ["FaUsers", "银发与健康消费", "保健品虚假宣传、养老预付费风险", "重点人群专项整治、预付资金监管"],
    ];
    for (let i = 0; i < 4; i++) {
      const g = grid[i], x = MX + (i % 2) * 6.15, y = 2.35 + Math.floor(i / 2) * 1.85;
      box(s, x, y, 5.95, 1.68, i % 3 === 0 ? K.tint : HX.gray);
      await iconCircle(s, g[0], x + 0.25, y + 0.25, 0.62, i < 2 ? K.teal : K.blue);
      T(s, g[1], { x: x + 1.05, y: y + 0.22, w: 4.7, h: 0.4, fontSize: 17, bold: true });
      T(s, [{ text: "风险：", options: { bold: true, color: K.red } }, { text: g[2], options: {} }], { x: x + 1.05, y: y + 0.68, w: 4.75, h: 0.4, fontSize: 14 });
      T(s, [{ text: "工具：", options: { bold: true, color: K.teal } }, { text: g[3], options: {} }], { x: x + 1.05, y: y + 1.08, w: 4.75, h: 0.45, fontSize: 14 });
    }
    T(s, "包容审慎 ≠ 放任不管：规则先行 · 底线监管 · 科技赋能 · 平台共治", { x: MX, y: 6.12, w: CW, h: 0.45, fontSize: 16, bold: true, color: K.jade, align: "center" });
    source(s, "国家统计局；国家市场监督管理总局。");
  }

  {
    const s = content(5, "落到基层：市县可以马上做的十件实事",
      "【约3分钟】最后给一份基层清单。这十件事有一个共同特点：不需要新增大量财政资金，主要靠规则、执法和服务，而且市县局有权限、能落地。大家可以对照本地情况，挑三五件先干起来，一年后回头看投诉数据、商圈客流和经营主体活跃度的变化。\n【互动建议】可以请1—2位地市局同志分享本地做法。");
    const acts = [
      ["放心消费", "线下无理由退货承诺单元向商圈、景区、夜市扩面"],
      ["放心消费", "预付式消费推行“资金存管+余额公示+冷静期”"],
      ["放心消费", "12315“诉转案”与重复投诉溯源治理"],
      ["公平竞争", "以旧换新领域价格与能效标识专项检查，防“先涨后降”"],
      ["公平竞争", "重点民生商品价格监测与行业成本调查"],
      ["经营主体", "涉企检查“综合查一次”、扫码入企，减少对经营的干扰"],
      ["经营主体", "个体工商户分型分类精准帮扶，信用修复“应修尽修”"],
      ["经营主体", "配合简化外摆经营、促销活动审批，设置“包容期”"],
      ["新业态", "直播带货、即时零售、外卖平台重点监测与抽查"],
      ["质量标准", "智能产品、银发产品质量抽检；服务质量评价试点"],
    ];
    const tagCol = { "放心消费": K.red, "公平竞争": K.blue, "经营主体": K.teal, "新业态": K.gold, "质量标准": K.gold };
    acts.forEach((a, i) => {
      const col = Math.floor(i / 5), row = i % 5, x = MX + col * 6.15, y = 1.65 + row * 0.98;
      box(s, x, y, 5.95, 0.85, row % 2 ? HX.gray : K.tint);
      T(s, String(i + 1).padStart(2, "0"), { x: x + 0.15, y, w: 0.6, h: 0.85, fontSize: 22, bold: true, color: K.jade, valign: "middle" });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.8, y: y + 0.24, w: 1.05, h: 0.38, fill: { color: tagCol[a[0]] }, line: { color: K.white, width: 0 }, rectRadius: 0.19 });
      T(s, a[0], { x: x + 0.8, y: y + 0.24, w: 1.05, h: 0.38, fontSize: 11, bold: true, color: K.white, align: "center", valign: "middle" });
      T(s, a[1], { x: x + 2.0, y, w: 3.85, h: 0.85, fontSize: 14, valign: "middle", lineSpacingMultiple: 1.05 });
    });
  }

  // ===================== 结语 =====================
  section("结语");
  {
    const s = content(9, "结语：市场监管的三个转变",
      "【约3分钟】最后用三个转变收尾。第一，从“守底线”到“促发展”：监管不只是成本，好的监管本身就是生产力，因为它降低了全社会的交易成本。第二，从“事后罚”到“事前防”：信用监管、智慧监管，让守法者无事不扰，让违法者无处可藏。第三，从“保护消费者”到“建设消费社会”：放心消费不只是一项民生工作，它是扩大内需的制度基础设施。");
    const sh = [
      ["从“守底线”", "到“促发展”", "监管不只是成本，也是生产力：好的监管降低全社会交易成本", "FaChartLine"],
      ["从“事后罚”", "到“事前防”", "信用监管、智慧监管：让守法者无事不扰，让违法者无处可藏", "FaShieldAlt"],
      ["从“保护消费者”", "到“建设消费社会”", "放心消费是扩大内需的制度基础设施", "FaLandmark"],
    ];
    for (let i = 0; i < 3; i++) {
      const c = sh[i], x = MX + i * 4.1;
      box(s, x, 1.65, 3.85, 3.55, K.tint);
      await iconCircle(s, c[3], x + 0.35, 1.85, 0.8, i === 2 ? K.red : K.teal);
      T(s, c[0], { x: x + 0.35, y: 2.8, w: 3.2, h: 0.45, fontSize: 19, color: K.muted, bold: true });
      T(s, c[1], { x: x + 0.35, y: 3.25, w: 3.3, h: 0.55, fontSize: 24, bold: true, color: i === 2 ? K.red : K.teal });
      T(s, c[2], { x: x + 0.35, y: 3.95, w: 3.15, h: 1.1, fontSize: 15, lineSpacingMultiple: 1.25 });
    }
    box(s, MX, 5.45, CW, 1.05, K.jade);
    T(s, [
      { text: "监管的最高境界  ", options: { bold: true, color: K.goldLight } },
      { text: "让守法经营者感觉不到监管的存在，让消费者时时感受到监管的保护。", options: { color: K.white } },
    ], { x: MX + 0.35, y: 5.45, w: CW - 0.7, h: 1.05, fontSize: 17, valign: "middle" });
  }
  {
    const s = pres.addSlide({ masterName: "DARK", sectionTitle: CURRENT_SECTION });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 1.3, w: 1.0, h: 1.0, fill: { color: K.red }, line: { color: K.red, width: 0 }, rectRadius: 0.1 });
    T(s, "信", { x: 1.0, y: 1.3, w: 1.0, h: 1.0, fontSize: 42, bold: true, color: K.white, align: "center", valign: "middle" });
    T(s, "宏观政策提振信心，市场监管守护信任。", { x: 1.0, y: 2.65, w: 11.3, h: 0.9, fontSize: 36, bold: true, color: K.white, valign: "middle" });
    T(s, "信心决定人们愿不愿意花钱，信任决定人们敢不敢在这里花钱。", { x: 1.0, y: 3.6, w: 11.3, h: 0.7, fontSize: 22, color: HX.darkText, valign: "middle" });
    s.addShape(pres.shapes.LINE, { x: 1.0, y: 4.75, w: 6.2, h: 0, line: { color: "2E5F5F", width: 0.75 } });
    T(s, "谢谢！欢迎批评指正", { x: 1.0, y: 5.0, w: 8, h: 0.6, fontSize: 24, bold: true, color: K.goldLight });
    s.addNotes("【结束语】宏观政策提振信心，市场监管守护信任。信心决定人们愿不愿意花钱，信任决定人们敢不敢在这里花钱。谢谢大家！\n【答疑预案】①“反内卷会不会推高物价？”——看第37页“反内卷的边界”；②“监管收紧会不会抑制新业态？”——包容审慎不等于放任，规则清楚反而降低合规不确定性；③“以旧换新退坡怎么办？”——从补商品转向补人、补服务，同时守住补贴归宿。");
  }

  // 附录
  {
    const s = content(9, "附：数据说明与讲前更新清单",
      "【不讲，供备课使用】讲课前请按右侧清单更新最新数据。");
    T(s, "数据说明", { x: MX, y: 1.6, w: 5.9, h: 0.4, fontSize: 17, bold: true, color: K.jade });
    T(s, paras([
      "宏观数据来自国家统计局月度/季度/年度发布；2025年社零按第五次经济普查修订后可比口径",
      "名义GDP增速、收入占GDP比重、平均消费倾向、消费倾向缺口等为作者据官方数据计算",
      "国际居民消费率为世界银行WDI取整值（2023年或最近年份），引用前建议再核",
      "日本CPI为总务省统计局年度同比，取一位小数",
      "市场监管数据来自总局新闻发布会、3·15数据发布及浙江省局公开资料",
    ], { fontSize: 14 }, { space: 8 }), { x: MX, y: 2.1, w: 5.9, h: 4.4, lineSpacingMultiple: 1.12 });
    box(s, 6.85, 1.6, 5.88, 4.9, HX.redTint);
    T(s, "讲前更新清单（10月20日前后）", { x: 7.1, y: 1.75, w: 5.4, h: 0.4, fontSize: 17, bold: true, color: K.red });
    T(s, paras([
      "前三季度GDP、9月社零与失业率 → 第6、8、9页",
      "9月CPI、PPI → 第10页",
      "9月青年失业率（不含在校生） → 第19页",
      "浙江前三季度GDP与社零 → 第36页",
      "以旧换新最新进展（发改委/商务部） → 第31页",
      "封面补充主讲人姓名与讲座地点",
    ], { fontSize: 14 }, { space: 10 }), { x: 7.1, y: 2.3, w: 5.4, h: 4.0, lineSpacingMultiple: 1.12 });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  // 主题字体：中文统一为微软雅黑
  const zip = await JSZip.loadAsync(fs.readFileSync(OUT));
  const part = "ppt/theme/theme1.xml";
  let xml = await zip.file(part).async("string");
  xml = xml.replace(/<a:ea typeface=""\/>/g, `<a:ea typeface="${FONT}"/>`).replace(/<a:font script="Hans" typeface="[^"]*"\/>/g, `<a:font script="Hans" typeface="${FONT}"/>`);
  zip.file(part, xml);
  // pptxgenjs 会在多段文字的同一段落中重复写入 <a:pPr>（不合 OOXML 顺序），只保留段首的那一个
  let fixed = 0;
  for (const name of Object.keys(zip.files).filter((n) => /^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(n))) {
    const sx = await zip.file(name).async("string");
    const out = sx.replace(/<a:p>([\s\S]*?)<\/a:p>/g, (m, inner) => {
      const cleaned = inner.replace(/<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/g, (pp, off) => {
        if (off === 0) return pp;
        fixed++;
        return "";
      });
      return "<a:p>" + cleaned + "</a:p>";
    });
    zip.file(name, out);
  }
  console.log("removed mid-paragraph pPr:", fixed);
  fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  console.log("wrote", OUT, "slides:", pres.slides.length);
}

build().catch((e) => { console.error(e); process.exit(1); });
