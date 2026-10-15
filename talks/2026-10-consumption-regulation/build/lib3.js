// 第三稿通用版式与工具：白底、深红标题、黑色正文（沿用提纲约定的沈艳2022课件风格）
const pptxgen = require("pptxgenjs");
const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));
const fs = require("fs");

const FONT = "Microsoft YaHei";
const COL = {
  R: "870C0C", INK: "171717", BLUE: "507B9B", BLUE_L: "A9C0D3", GRAY: "666666", GRAY_L: "BDBDBD",
  RULE: "D9D9D9", TINT: "F7F1F0", TINTB: "EEF3F7", GOLD: "A87B12", WHITE: "FFFFFF", DARK: "2B2B2B", PALE_R: "D9A3A3",
};
const W = 13.333, H = 7.5, MX = 0.6, CW = W - 2 * MX;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "消费提振视角下的中国宏观经济形势分析";
pres.subject = "浙江省市场监管系统干部专题培训";
pres.author = "王泽荣";
pres.company = "浙江财经大学经济学院";
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };

pres.defineSlideMaster({
  title: "PAGE",
  background: { color: COL.WHITE },
  objects: [
    { line: { x: MX, y: 1.4, w: CW, h: 0, line: { color: COL.RULE, width: 0.75 } } },
    { rect: { x: MX, y: 1.375, w: 0.9, h: 0.05, fill: { color: COL.R }, line: { color: COL.R, width: 0 } } },
    { line: { x: MX, y: 6.95, w: CW, h: 0, line: { color: COL.RULE, width: 0.5 } } },
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.62, w: CW, h: 0.72, fontSize: 28, bold: true, color: COL.R, fontFace: FONT, valign: "middle", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.13, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: COL.GRAY, fontFace: FONT, align: "right" },
});
pres.defineSlideMaster({
  title: "BLANK",
  background: { color: COL.WHITE },
  objects: [],
  slideNumber: { x: 12.13, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: COL.GRAY, fontFace: FONT, align: "right" },
});

// ---------- 文本 ----------
// “**重点**”→深红加粗；“⏎”→段内换行
function runs(str, base, hl) {
  const out = [];
  String(str).split("⏎").forEach((line, li) => {
    let first = true;
    line.split(/\*\*(.+?)\*\*/).forEach((seg, i) => {
      if (!seg) return;
      const o = Object.assign({}, base, i % 2 ? { bold: true, color: hl || COL.R } : {});
      if (li > 0 && first) o.softBreakBefore = true;
      first = false;
      out.push({ text: seg, options: o });
    });
  });
  if (!out.length) out.push({ text: " ", options: Object.assign({}, base) });
  return out;
}
function T(slide, text, opts) {
  const o = Object.assign({ fontFace: FONT, color: COL.INK, fontSize: 18, margin: 0, valign: "top", isTextBox: true, lang: "zh-CN" }, opts);
  const content = typeof text === "string" && /\*\*|⏎/.test(text) ? runs(text, { fontSize: o.fontSize, color: o.color, bold: o.bold }, o.hl) : text;
  slide.addText(content, o);
}
// 项目符号段落
function paras(items, base, opt = {}) {
  const all = [];
  items.forEach((it, idx) => {
    const r = runs(it, base, opt.hl);
    if (opt.bullet !== false) r[0].options = Object.assign({}, r[0].options, { bullet: { code: "25A0", indent: opt.indent || 16 } });
    r[0].options = Object.assign({}, r[0].options, { paraSpaceAfter: opt.space == null ? 8 : opt.space });
    if (idx < items.length - 1) r[r.length - 1].options = Object.assign({}, r[r.length - 1].options, { breakLine: true });
    all.push(...r);
  });
  return all;
}
// “要点：解释”式分层段落（要点深红加粗）
function points(slide, items, o) {
  const size = o.size || 18;
  const all = [];
  items.forEach((it, idx) => {
    const head = { text: it[0] + (o.sep == null ? "：" : o.sep), options: { bold: true, color: o.headColor || COL.R, fontSize: size, paraSpaceAfter: o.space == null ? 12 : o.space } };
    if (o.bullet) head.options.bullet = { code: "25A0", indent: 16 };
    const body = runs(it[1], { fontSize: size, color: COL.INK });
    if (idx < items.length - 1) body[body.length - 1].options = Object.assign({}, body[body.length - 1].options, { breakLine: true });
    all.push(head, ...body);
  });
  T(slide, all, { x: o.x, y: o.y, w: o.w, h: o.h, lineSpacingMultiple: o.ls || 1.12, valign: o.valign || "top" });
}

// ---------- 形状 ----------
function rect(slide, x, y, w, h, fill, line) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: line ? { color: line, width: 0.75 } : { color: fill, width: 0 } });
}
function hline(slide, x, y, w, color = COL.RULE, width = 0.75, dash) {
  slide.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: Object.assign({ color, width }, dash ? { dashType: dash } : {}) });
}
function vline(slide, x, y, h, color = COL.RULE, width = 0.75) {
  slide.addShape(pres.shapes.LINE, { x, y, w: 0, h, line: { color, width } });
}
function arrow(slide, x, y, w, h, color = COL.GRAY) {
  slide.addShape(pres.shapes.LINE, { x, y, w, h, line: { color, width: 1.5, endArrowType: "triangle" } });
}
// 资料来源脚注
function src(slide, text) {
  T(slide, text, { x: MX, y: 6.99, w: 11.3, h: 0.32, fontSize: 10, color: COL.GRAY, valign: "top" });
}
// 页面结论条：浅底＋左侧深红竖条
function take(slide, text, o = {}) {
  const x = o.x == null ? MX : o.x, y = o.y == null ? 6.12 : o.y, w = o.w == null ? CW : o.w, h = o.h == null ? 0.66 : o.h;
  rect(slide, x, y, w, h, o.fill || COL.TINT);
  rect(slide, x, y, 0.08, h, COL.R);
  T(slide, text, { x: x + 0.28, y, w: w - 0.45, h, fontSize: o.size || 18, valign: "middle", color: COL.INK, lineSpacingMultiple: 1.05 });
}
// 小标题（带深红短竖条）
function subhead(slide, text, x, y, w, o = {}) {
  rect(slide, x, y + 0.06, 0.07, 0.3, o.bar || COL.R);
  T(slide, text, { x: x + 0.18, y, w: w - 0.18, h: 0.42, fontSize: o.size || 19, bold: true, color: o.color || COL.INK, valign: "middle" });
}
// 关键数字
function kpi(slide, x, y, w, num, label, o = {}) {
  T(slide, num, { x, y, w, h: o.nh || 0.9, fontSize: o.ns || 44, bold: true, color: o.color || COL.R, valign: "bottom" });
  T(slide, label, { x, y: y + (o.nh || 0.9) + 0.05, w, h: o.lh || 0.7, fontSize: o.ls || 15, color: COL.INK, lineSpacingMultiple: 1.05 });
}
// 流程框
function stepBox(slide, x, y, w, h, text, o = {}) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: o.fill || COL.WHITE }, line: { color: o.line || COL.R, width: o.lw || 1.25 }, rectRadius: 0.06 });
  T(slide, text, { x: x + 0.08, y, w: w - 0.16, h, fontSize: o.size || 16, bold: o.bold !== false, color: o.color || COL.INK, align: "center", valign: "middle", lineSpacingMultiple: 1.0 });
}
function chain(slide, steps, o) {
  const n = steps.length, gap = o.gap || 0.42, bw = (o.w - gap * (n - 1)) / n;
  steps.forEach((t, i) => {
    const x = o.x + i * (bw + gap);
    stepBox(slide, x, o.y, bw, o.h, t, o);
    if (i < n - 1) arrow(slide, x + bw + 0.06, o.y + o.h / 2, gap - 0.12, 0, o.arrow || COL.R);
  });
  return bw;
}

// ---------- 表格 ----------
function table(slide, rows, o) {
  const size = o.size || 16;
  const data = rows.map((r, ri) => r.map((c, ci) => {
    const head = ri === 0 && o.header !== false;
    const cell = typeof c === "object" && c !== null && !Array.isArray(c) ? c : { text: String(c) };
    const opt = Object.assign({
      fontFace: FONT, fontSize: head ? (o.headSize || size) : size, color: head ? COL.R : COL.INK, bold: head || (o.boldFirstCol && ci === 0),
      valign: "middle", align: (o.align && o.align[ci]) || "left", margin: [3, 6, 3, 6],
      fill: { color: head ? (o.headFill || "F3E7E6") : (o.zebra && ri % 2 === 0 ? "FAFAFA" : COL.WHITE) },
      border: [{ type: "solid", pt: 0.75, color: COL.RULE }, { type: "none" }, { type: "solid", pt: 0.75, color: COL.RULE }, { type: "none" }],
    }, cell.options || {});
    const text = /\*\*|⏎/.test(cell.text) ? runs(cell.text, { fontSize: opt.fontSize, color: opt.color, bold: opt.bold }) : cell.text;
    return { text, options: opt };
  }));
  slide.addTable(data, { x: o.x, y: o.y, w: o.w, colW: o.colW, rowH: o.rowH, fontFace: FONT, autoPage: false });
}

// ---------- 图表 ----------
function co(extra) {
  return Object.assign({
    catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT, dataLabelFontFace: FONT, legendFontFace: FONT, titleFontFace: FONT,
    catAxisLabelColor: "333333", valAxisLabelColor: COL.GRAY, catAxisLabelFontSize: 14, valAxisLabelFontSize: 13,
    valGridLine: { color: "E6E6E6", size: 0.75 }, catGridLine: { style: "none" },
    catAxisLineShow: true, catAxisLineColor: "999999", valAxisLineShow: false,
    dataLabelFontSize: 13, dataLabelColor: COL.INK,
    showLegend: false, legendPos: "t", legendFontSize: 14, legendColor: "333333",
    showTitle: false, barGapWidthPct: 55,
  }, extra);
}
function chartCap(slide, text, x, y, w) {
  T(slide, text, { x, y, w, h: 0.36, fontSize: 15, bold: true, color: COL.INK, valign: "middle" });
}

// ---------- 页面 ----------
const PARTS = [
  null,
  { n: "一", name: "消费之镜", q: "为什么可以从消费看宏观经济", min: 8, span: "第0—8分钟" },
  { n: "二", name: "外需之变", q: "外部变化怎样传到国内家庭", min: 12, span: "第8—20分钟" },
  { n: "三", name: "内需之困", q: "消费弱在哪里，哪些约束有证据", min: 28, span: "第20—48分钟" },
  { n: "四", name: "三账联读", q: "约束怎样形成，政策怎样起作用", min: 25, span: "第58—83分钟" },
  { n: "五", name: "浙江之察", q: "浙江有哪些特点，潜力怎样转化为交易", min: 20, span: "第83—103分钟" },
  { n: "六", name: "监管之为", q: "哪些消费障碍能在职责内改善", min: 12, span: "第103—115分钟" },
];
const REG = {};
const TIMING = [];
let CURSEC;
function section(t) { CURSEC = t; pres.addSection({ title: t }); }
function curSec() { return CURSEC; }
function titleLen(t) { let n = 0; for (const ch of t) n += /[\x00-\x7F]/.test(ch) ? 0.55 : 1; return n; }
function page(part, title, o = {}) {
  const s = pres.addSlide({ masterName: "PAGE", sectionTitle: CURSEC });
  if (titleLen(title) > 30) console.warn("TITLE TOO LONG (" + titleLen(title) + "): " + title);
  const P = PARTS[part];
  if (P) {
    if (o.opener) {
      rect(s, 0, 0, W, 0.42, COL.R);
      T(s, [{ text: `第${P.n}部分　${P.name}`, options: { bold: true, fontSize: 14, color: COL.WHITE } }, { text: `　｜　${P.q}`, options: { fontSize: 14, color: "F3DADA" } }],
        { x: MX, y: 0, w: 9.6, h: 0.42, valign: "middle" });
      T(s, `${P.min}分钟 · ${P.span}`, { x: 9.2, y: 0, w: 3.53, h: 0.42, fontSize: 13, color: "F3DADA", align: "right", valign: "middle" });
    } else {
      T(s, `${P.n}　${P.name}`, { x: MX, y: 0.2, w: 4, h: 0.32, fontSize: 12, bold: true, color: COL.R, valign: "middle" });
      for (let i = 1; i <= 6; i++) {
        const on = i === part;
        s.addShape(pres.shapes.RECTANGLE, { x: 11.29 + (i - 1) * 0.25, y: 0.29, w: 0.17, h: 0.13, fill: { color: on ? COL.R : (i < part ? COL.PALE_R : "E3E3E3") }, line: { color: on ? COL.R : (i < part ? COL.PALE_R : "E3E3E3"), width: 0 } });
      }
    }
  }
  s.addText(title, { placeholder: "title", fontFace: FONT, lang: "zh-CN" });
  if (o.notes) s.addNotes(o.notes);
  if (o.key) REG[o.key] = pres.slides.length;
  if (o.min != null) TIMING.push([pres.slides.length, o.min, title]);
  return s;
}

// ---------- 输出与后处理 ----------
async function finish(OUT, THEME, applyTheme) {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  const zip = await JSZip.loadAsync(fs.readFileSync(OUT));
  const part = "ppt/theme/theme1.xml";
  let xml = await zip.file(part).async("string");
  xml = xml.replace(/<a:ea typeface=""\/>/g, `<a:ea typeface="${FONT}"/>`).replace(/<a:font script="Hans" typeface="[^"]*"\/>/g, `<a:font script="Hans" typeface="${FONT}"/>`);
  zip.file(part, xml);
  // pptxgenjs 会在多片段段落中重复写入 <a:pPr>，只保留段首那一个；项目符号统一为深红小方块
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
    out = out.replace(/<a:buSzPct val="100000"\/><a:buChar/g, `<a:buClr><a:srgbClr val="${COL.R}"/></a:buClr><a:buSzPct val="70000"/><a:buChar`);
    zip.file(name, out);
  }
  // 待发布数据（null）在图表缓存中不写点，显示为空位
  for (const name of Object.keys(zip.files).filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n))) {
    const x = await zip.file(name).async("string");
    zip.file(name, x.replace(/<c:pt idx="\d+"><c:v><\/c:v><\/c:pt>/g, ""));
  }
  fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  return fixed;
}

module.exports = { section, curSec, pres, FONT, COL, W, H, MX, CW, runs, T, paras, points, rect, hline, vline, arrow, src, take, subhead, kpi, stepBox, chain, table, co, chartCap, page, PARTS, REG, TIMING, finish };
