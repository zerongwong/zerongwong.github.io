// 消费提振视角下的中国宏观经济形势分析 —— 第三稿生成脚本（按局里定稿的六部分提纲）
const fs = require("fs");
const L = require("./lib3.js");
const { pres, COL, W, H, MX, CW, T, paras, points, rect, hline, vline, arrow, src, take, subhead, kpi, stepBox, chain, table, co, chartCap, page, REG, TIMING, finish } = L;
const D = require("./data3.js");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");
const OUT = process.argv[2] || "deck3.pptx";
const THEME = {
  name: "Consumption Lens", headFontFace: L.FONT, bodyFontFace: L.FONT,
  colors: { dk1: "171717", lt1: "FFFFFF", dk2: "870C0C", lt2: "F7F1F0", accent1: "870C0C", accent2: "507B9B", accent3: "A87B12", accent4: "6B7B8C", accent5: "8C8C8C", accent6: "D9A3A3", hlink: "507B9B", folHlink: "666666" },
};
const PARTS = require("./parts3.js");

async function build() {
  for (const p of PARTS) await p({ L, D });
  const fixed = await finish(OUT, THEME, applyTheme);
  const total = TIMING.reduce((a, b) => a + b[1], 0);
  console.log("removed mid-paragraph pPr:", fixed);
  console.log("slides:", pres.slides.length, "timed minutes:", total);
  fs.writeFileSync(OUT.replace(/\.pptx$/, ".timing.json"), JSON.stringify({ REG, TIMING }, null, 1));
}
build().catch((e) => { console.error(e); process.exit(1); });
