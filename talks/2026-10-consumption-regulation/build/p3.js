// 第三部分 内需之困（P14—P29，28分钟）＋ P30 休息
module.exports = async function ({ L, D }) {
  const { pres, COL, W, MX, CW, T, paras, points, rect, hline, vline, arrow, src, take, subhead, kpi, stepBox, chain, table, co, chartCap, page } = L;
  const yl = (a) => a.map((x) => x.replace("2026上半年", "2026\n上半年"));

  L.section("三 内需之困");
  // ---------- P14 消费在增长中的位置 ----------
  {
    const s = page(3, "消费仍是增长第一拉动力，但拉动力明显弱于疫情前", { opener: true, min: 1.5, key: "p_pull", notes:
      "【时长】1.5分钟\n【讲法】回到国内。灰色柱是GDP增速，红色柱是最终消费拉动GDP的百分点。2019年消费拉动3.5个百分点，2025年2.6个，今年上半年2.1个。上半年GDP增长4.7%，其中消费拉动2.1、投资1.7、净出口0.8个百分点。消费贡献率约45%，去年同期是52%。\n消费仍是第一拉动力，但出力明显不如疫情前：上半年消费拉动不到一半，增长更多靠投资和净出口；这两年外需出力偏多。2025年中央经济工作会议的判断是“国内供强需弱矛盾突出”，今年7月统计局发布上半年数据时用的也是这个表述。\n【口径】最终消费包括居民和政府消费；拉动百分点不是贡献率，也不是GDP占比；上半年三项拉动相加4.6，与4.7有舍入尾差。\n【可能的追问】“资本形成拉动1.7个百分点，可1—8月固定资产投资却下降7.2%？”——两者口径不同：资本形成按支出法核算，包含存货变化，也不是固定资产投资那套统计范围，不能直接对比。\n【衔接】GDP回答总体作用。日常的商品和服务消费，还要分开看。\n【进度检查点】本页应在第20分钟左右开始。" });
    chartCap(s, "GDP实际增速与最终消费拉动（百分点）", MX, 1.58, 7.7);
    s.addChart(pres.charts.BAR, [
      { name: "GDP增速", labels: yl(D.gdpYears), values: D.gdpReal },
      { name: "最终消费拉动", labels: yl(D.gdpYears), values: D.pullC },
    ], co({ x: MX, y: 1.95, w: 7.7, h: 4.05, barDir: "col", chartColors: [COL.GRAY_L, COL.R], showValue: true, dataLabelFormatCode: "0.0", showLegend: true, legendPos: "t", valAxisMinVal: -2, valAxisMaxVal: 10, valAxisMajorUnit: 2, valAxisHidden: true, valGridLine: { style: "none" }, barGapWidthPct: 45, catAxisLabelFontSize: 13, catAxisLabelPos: "low" }));
    const px = 8.65, pw = CW - (px - MX);
    T(s, "2026年上半年：GDP增长4.7%", { x: px, y: 1.62, w: pw, h: 0.42, fontSize: 18, bold: true });
    [["消费", "2.1", COL.R], ["投资", "1.7", COL.BLUE], ["净出口", "0.8", COL.GRAY]].forEach((k, i) => {
      const x = px + i * (pw / 3);
      T(s, k[1], { x, y: 2.1, w: pw / 3, h: 0.75, fontSize: 38, bold: true, color: k[2], align: "center", valign: "bottom" });
      T(s, k[0] + "拉动", { x, y: 2.88, w: pw / 3, h: 0.35, fontSize: 14, align: "center", color: COL.GRAY });
    });
    T(s, paras([
      "消费贡献率约45%，去年同期为52%",
      "官方判断：“国内供强需弱矛盾突出”（2025年中央经济工作会议；统计局2026年7月）",
    ], { fontSize: 16, color: COL.INK }, { space: 10 }), { x: px, y: 3.45, w: pw, h: 2.5, lineSpacingMultiple: 1.1 });
    take(s, "消费拉动不到一半，增长更多靠**投资和净出口**——这正是“供强需弱”的含义。", { y: 6.12, h: 0.66, size: 18 });
    src(s, "资料来源：国家统计局。最终消费含居民与政府消费；拉动百分点≠贡献率≠占比；上半年三项拉动相加有舍入尾差。");
  }

  // ---------- P15 社零节奏 ----------
  {
    const s = page(3, "社零增速从6.4%降到0.4%：高基数、补贴减弱与汽车拖累", { min: 2, key: "p_rs", notes:
      "【时长】2分钟\n【讲法】左图是社零月度同比。去年5月6.4%是阶段高点，今年5月一度转负（-0.6%），是2022年12月以来首次负增长；8月0.4%。1—8月累计增长1.1%。\n要分清几件事：一是高基数——去年以旧换新力度大，抬高了基数；二是补贴力度减弱——今年以旧换新中央资金安排2500亿元，少于去年的3000亿元；家电补贴最高比例从20%降到15%，单件上限从2000元降到1500元，只补一级能效产品；三是需求前置——本来明后年才换的车和家电，提前到去年买了，今年自然回吐。汽车1—8月下降13.9%，扣除汽车，社零增长2.7%。\n剔除基数看两年平均增速：一季度约3.5%，7—8月约2%——基数解释了一部分，但动能确实在减弱。\n【口径】1—2月合并发布，不是单月；社零为名义值，不覆盖全部服务消费；同比受基数影响，不据单月推断全年。\n【衔接】金额的变化，还要结合价格看。" });
    chartCap(s, "社会消费品零售总额月度同比（%）", MX, 1.58, 7.9);
    const rsLab = D.rsLabels.map((x) => (x.includes("1-2") ? "1-2" : x));
    s.addChart(pres.charts.LINE, [{ name: "社零同比", labels: rsLab, values: D.rs }], co({
      x: MX, y: 1.95, w: 7.9, h: 4.0, chartColors: [COL.R], lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 7, showValue: true, dataLabelFormatCode: "0.0", dataLabelPosition: "t", dataLabelFontSize: 11,
      valAxisMinVal: -2, valAxisMaxVal: 8, valAxisMajorUnit: 2, catAxisLabelFontSize: 12, catAxisLabelPos: "low",
    }));
    T(s, "2025年", { x: MX + 0.6, y: 5.9, w: 4.0, h: 0.28, fontSize: 12, color: COL.GRAY, align: "center" });
    T(s, "2026年", { x: MX + 4.95, y: 5.9, w: 2.8, h: 0.28, fontSize: 12, color: COL.GRAY, align: "center" });
    const px = 8.85, pw = CW - (px - MX);
    T(s, "年度增速（%）", { x: px, y: 1.58, w: pw, h: 0.36, fontSize: 15, bold: true });
    s.addChart(pres.charts.BAR, [{ name: "年度", labels: D.rsAnnualYears, values: D.rsAnnual }], co({
      x: px, y: 1.92, w: pw, h: 2.2, barDir: "col", chartColors: D.rsAnnual.map((v, i) => (i === D.rsAnnual.length - 1 ? COL.R : COL.BLUE_L)), showValue: true, dataLabelFormatCode: "0.0", dataLabelFontSize: 11,
      valAxisHidden: true, valGridLine: { style: "none" }, catAxisLabelFontSize: 11, valAxisMinVal: -6, valAxisMaxVal: 15, catAxisLabelPos: "low", barGapWidthPct: 35,
    }));
    T(s, paras([
      "1—8月社零 **+1.1%**；扣除汽车 **+2.7%**",
      "汽车 **-13.9%**",
      "两年平均：一季度约**3.5%**，7—8月约**2%**",
    ], { fontSize: 17, color: COL.INK }, { space: 6 }), { x: px, y: 4.3, w: pw, h: 1.6, lineSpacingMultiple: 1.1 });
    take(s, "有基数和补贴节奏的因素，也有**耐用品“需求前置”后的回吐**：不全是需求塌陷，但动能确实在减弱。", { y: 6.12, h: 0.66, size: 18 });
    src(s, "资料来源：国家统计局。1—2月合并发布；社零为名义增速。");
  }

  // ---------- P16 价格 ----------
  {
    const s = page(3, "物价由负转正，但核心CPI仅1.0%：需求端并不热", { min: 1.5, key: "p_price", notes:
      "【时长】1.5分钟\n【讲法】价格是今年最大的变化之一。PPI连续下降41个月后，今年3月转正，9月已经涨到5.0%；CPI从去年的零附近回到今年9月的0.9%。\n但要看清是谁在涨：PPI上涨主要由石油、煤炭、有色、化工链带动，带有明显的输入型、成本型特征；核心CPI只有1.0%，需求端并不热。9月食品价格下降1.7%，猪肉降10.0%、鲜菜降8.1%。家庭的感受是“吃的便宜了，开车、用的东西贵了”。\n还有一个变化：据机构测算，GDP平减指数今年二季度转正，此前连续十多个季度为负。名义增速重新跑赢实际增速，对企业营收和财政收入是好事，但要防止成本上涨挤压中下游利润（第38页）。\n【口径】不能用总体CPI机械平减社零；二者统计对象和权重不同。2026年CPI以2025年为基期。\n【衔接】总量背后，不同消费类别的方向并不一样。" });
    chartCap(s, "CPI与PPI同比（%）", MX, 1.58, 7.9);
    const pLab = D.pLabels.map((x) => x.replace(/^2[56]年/, "").replace("月", ""));
    s.addChart(pres.charts.LINE, [
      { name: "CPI", labels: pLab, values: D.cpi },
      { name: "PPI", labels: pLab, values: D.ppi },
    ], co({ x: MX, y: 1.95, w: 7.9, h: 3.95, chartColors: [COL.R, COL.BLUE], lineSize: 2.5, lineDataSymbol: "none", showLegend: true, legendPos: "t", valAxisMinVal: -4, valAxisMaxVal: 6, valAxisMajorUnit: 2, catAxisLabelFontSize: 12, catAxisLabelPos: "low" }));
    T(s, "2025年（月）", { x: MX + 0.6, y: 5.88, w: 4.2, h: 0.28, fontSize: 12, color: COL.GRAY, align: "center" });
    T(s, "2026年（月）", { x: MX + 4.9, y: 5.88, w: 2.9, h: 0.28, fontSize: 12, color: COL.GRAY, align: "center" });
    const px = 8.85, pw = CW - (px - MX);
    T(s, "2026年9月", { x: px, y: 1.6, w: pw, h: 0.4, fontSize: 18, bold: true });
    table(s, [
      ["指标", "同比"],
      ["CPI", "+0.9%"],
      ["核心CPI", "+1.0%"],
      ["食品", "-1.7%"],
      ["　猪肉", "-10.0%"],
      ["　鲜菜", "-8.1%"],
      ["PPI", "+5.0%"],
    ], { x: px, y: 2.05, w: pw, colW: [pw * 0.6, pw * 0.4], rowH: 0.42, size: 16, align: ["left", "right"] });
    T(s, "PPI上涨主要由石油、煤炭、有色、化工链带动", { x: px, y: 5.15, w: pw, h: 0.75, fontSize: 14, color: COL.GRAY, lineSpacingMultiple: 1.05 });
    take(s, "**上游涨得快、终端涨得慢**：家庭感受到“吃的便宜、开的用的贵”；企业感受到成本上升、售价难提。", { y: 6.12, h: 0.66, size: 18 });
    src(s, "资料来源：国家统计局（2026年10月14日发布9月数据）。");
  }

  // ---------- P17 结构 ----------
  {
    const s = page(3, "商品消费分化：通讯器材增长，汽车和住房相关品类下降", { min: 1.5, key: "p_struct", notes:
      "【时长】1.5分钟\n【讲法】8月限额以上单位的分类数据：通讯器材增长27.3%，可能有新机型和存储芯片涨价的因素；烟酒、化妆品、药品也在增长。另一头，汽车下降18.5%，金银珠宝下降17.5%（可能与金价高企有关），建材下降11.8%、家具下降7.9%——这些和住房调整直接相关。\n1—8月，服务零售增长4.9%，快于商品零售的1.0%；餐饮收入增长2.4%，但限额以上餐饮8月下降0.7%。\n分化有三个来源，不能只用一个原因解释：一是补贴节奏，汽车是去年补贴集中释放后的回吐，数码产品仍在补贴范围内；二是住房调整，拖累家具、建材；三是收入预期，大额支出更谨慎。还要看到，消费偏弱是普遍的：上半年居民教育文化娱乐支出增速从去年同期的11.8%降到4.9%，服务也在放缓。\n【口径】分类为限额以上单位；服务零售与餐饮有交叉，不能相加；增速高不等于贡献大。\n【衔接】这是卖方的数据。家庭自己的账本是什么样？" });
    chartCap(s, "2026年8月限额以上单位商品零售额同比（%）", MX, 1.58, 7.6);
    const lab = D.augCats.map((r) => r[0]), val = D.augCats.map((r) => r[1]);
    s.addChart(pres.charts.BAR, [{ name: "8月同比", labels: lab, values: val }], co({
      x: MX, y: 1.95, w: 7.6, h: 4.05, barDir: "bar", chartColors: val.map((v) => (v >= 0 ? COL.BLUE : COL.R)), showValue: true, dataLabelFormatCode: "0.0", dataLabelFontSize: 12,
      valAxisMinVal: -25, valAxisMaxVal: 35, valAxisHidden: true, valGridLine: { style: "none" }, catAxisOrientation: "maxMin", catAxisLabelPos: "low", catAxisLabelFontSize: 13, barGapWidthPct: 35,
    }));
    const px = 8.55, pw = CW - (px - MX);
    T(s, "1—8月累计同比", { x: px, y: 1.6, w: pw, h: 0.4, fontSize: 18, bold: true });
    table(s, [
      ["类别", "同比"],
      ["服务零售额", "+4.9%"],
      ["餐饮收入", "+2.4%"],
      ["商品零售", "+1.0%"],
      ["网上商品零售", "+4.3%"],
    ], { x: px, y: 2.05, w: pw, colW: [pw * 0.62, pw * 0.38], rowH: 0.44, size: 16, align: ["left", "right"] });
    T(s, "网上商品零售额占社零约25.7%", { x: px, y: 4.4, w: pw, h: 0.4, fontSize: 14, color: COL.GRAY });
    take(s, "分化有三个来源：**补贴节奏**（汽车回吐、数码仍受补贴）、**住房调整**（家具、建材）、**收入预期**（大额更谨慎）。", { y: 6.12, h: 0.66, size: 17 });
    src(s, "资料来源：国家统计局、商务部。分类为限额以上单位；服务零售与餐饮口径有交叉，不能相加。");
  }

  // ---------- P18 消费倾向 ----------
  {
    const s = page(3, "居民收入在增长，花钱的比例却还没回到疫情前", { min: 1.5, key: "p_apc", notes:
      "【时长】1.5分钟\n【讲法】看家庭自己的账本——住户调查。左图是平均消费倾向，也就是每100元可支配收入里花掉多少：2019年70.1元，2020年掉到65.9元，之后在66.5—68.6元之间波动，没有回到疫情前。今年上半年收入名义增长5.2%，消费增长3.7%，消费慢于收入；上半年消费倾向64.6%，低于去年同期的65.5%。\n右图是服务性消费占比，已回到46%左右，说明支出结构在向服务转移。但上半年教育文化娱乐支出增速从去年同期的11.8%降到4.9%，交通通信从8.4%降到4.7%。\n粗算一下这个差距的分量：按2025年人均可支配收入43377元、约14亿人计算，消费倾向每回升1个百分点，大约对应6000亿元居民消费。\n【口径】平均消费倾向不是边际消费倾向；上半年数据有季节性，只与上年同期比较。\n【衔接】支出从哪里来？首先要看收入，以及收入在不同人之间的分布。" });
    const cw = (CW - 0.5) / 2;
    chartCap(s, "平均消费倾向（人均消费支出÷人均可支配收入，%）", MX, 1.58, cw);
    s.addChart(pres.charts.LINE, [{ name: "平均消费倾向", labels: D.apcYears, values: D.apc }], co({
      x: MX, y: 1.95, w: cw, h: 3.4, chartColors: [COL.R], lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 8, showValue: true, dataLabelFormatCode: "0.0", dataLabelPosition: "t",
      valAxisMinVal: 62, valAxisMaxVal: 72, valAxisMajorUnit: 2, valAxisHidden: true, valGridLine: { style: "none" },
    }));
    const x2 = MX + cw + 0.5;
    chartCap(s, "服务性消费支出占居民消费支出比重（%）", x2, 1.58, cw);
    s.addChart(pres.charts.BAR, [{ name: "服务性消费占比", labels: D.apcYears, values: D.svcShare }], co({
      x: x2, y: 1.95, w: cw, h: 3.4, barDir: "col", chartColors: [COL.BLUE], showValue: true, dataLabelFormatCode: "0.0",
      valAxisMinVal: 30, valAxisMaxVal: 50, valAxisHidden: true, valGridLine: { style: "none" }, barGapWidthPct: 45,
    }));
    T(s, "2026年上半年：收入名义增长5.2%，消费增长3.7%；消费倾向64.6%，低于上年同期的65.5%", { x: MX, y: 5.42, w: CW, h: 0.45, fontSize: 16, color: COL.INK, valign: "middle" });
    take(s, "**挣得慢了，花得也少了**：按2025年数据粗算，消费倾向每回升1个百分点，约对应**6000亿元**居民消费。", { y: 6.12, h: 0.66, size: 18 });
    src(s, "资料来源：国家统计局住户调查；平均消费倾向与6000亿元为作者计算。");
  }

  // ---------- P19 收入分布 ----------
  {
    const s = page(3, "收入中位数只有平均数的83%，上半年收入增速低于GDP", { min: 2, key: "p_income", notes:
      "【时长】2分钟\n【讲法】平均数会“被拉高”。上半年全国居民人均可支配收入22981元，中位数19036元，只有平均数的83%——一半以上的人达不到平均数。中位数名义增长4.7%，也低于平均数的5.2%。\n右图对比收入和GDP的实际增速。2019年以来大体同步，但今年上半年居民收入实际增长4.2%，低于GDP的4.7%。上半年财产净收入只增长1.1%，资产端没有给家庭“加分”。农村居民收入增长6.4%，快于城镇的4.4%。\n所以“数据在增长、感觉没增长”有它的道理：多数家庭的感受，取决于中位数，和自己那份收入稳不稳。\n【口径】中位数增速为名义值；均值与中位数之差不能直接换算不平等程度；住户调查季度数据不含自产自用部分。\n【衔接】收入稳不稳，首先取决于工作稳不稳。" });
    chartCap(s, "2026年上半年人均可支配收入（元）", MX, 1.58, 4.4);
    s.addChart(pres.charts.BAR, [{ name: "收入", labels: ["平均数", "中位数"], values: [D.incH1, D.medH1] }], co({
      x: MX, y: 1.95, w: 4.4, h: 3.5, barDir: "col", chartColors: [COL.BLUE_L, COL.R], showValue: true, dataLabelFormatCode: "#,##0", dataLabelFontSize: 15,
      valAxisMinVal: 0, valAxisMaxVal: 26000, valAxisHidden: true, valGridLine: { style: "none" }, barGapWidthPct: 60, catAxisLabelFontSize: 15,
    }));
    T(s, "名义增长：平均数5.2%，中位数4.7%", { x: MX, y: 5.45, w: 4.4, h: 0.4, fontSize: 15, color: COL.GRAY, align: "center" });
    const x2 = MX + 4.9, w2 = CW - 4.9;
    chartCap(s, "居民人均可支配收入与GDP实际增速（%）", x2, 1.58, w2);
    s.addChart(pres.charts.LINE, [
      { name: "GDP", labels: yl(D.gdpYears), values: D.gdpReal },
      { name: "居民收入", labels: yl(D.gdpYears), values: D.incGrowthReal },
    ], co({ x: x2, y: 1.95, w: w2, h: 3.9, chartColors: [COL.GRAY_L, COL.R], lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 7, showLegend: true, legendPos: "t", showValue: false,
      valAxisMinVal: 0, valAxisMaxVal: 10, valAxisMajorUnit: 2, catAxisLabelFontSize: 12 }));
    take(s, "多数家庭的感受取决于**中位数**和**收入稳不稳**；上半年财产净收入仅增长1.1%，资产端也没有“加分”。", { y: 6.12, h: 0.66, size: 18 });
    src(s, "资料来源：国家统计局住户调查与国民经济核算。2026年上半年居民收入实际增长4.2%，GDP增长4.7%。");
  }

  // ---------- P20 就业 ----------
  {
    const s = page(3, "总体就业基本稳定，压力集中在青年和新进入者", { min: 2, key: "p_jobs", notes:
      "【时长】2分钟\n【讲法】全国城镇调查失业率今年在5.0%—5.4%之间，8月5.3%，基本平稳。但分年龄看，16—24岁（不含在校生）8月是18.9%，毕业季季节性冲高，和去年同期持平；25—29岁7.5%，3月一度到7.7%，是这个口径发布以来的最高；30—59岁主力人群3.9%，比较稳定。\n今年高校毕业生1270万人，比去年多48万。2025年农民工月均收入5075元，只增长2.3%。\n失业率看的是“有没有工作”，消费更看“收入稳不稳”：工时、奖金、订单型收入的波动，会先压缩大额、长期的支出。\n【口径】青年失业率为不含在校生口径；8月含毕业季季节因素，与上年同月比较；统计口径调整前后不直接拼接。\n【衔接】对未来的担心，能不能在家庭的花钱意愿里看到？" });
    chartCap(s, "城镇调查失业率（%）", MX, 1.58, 7.9);
    const idx = D.uLabels.map((_, i) => i).filter((i) => i >= 5);
    const labs = idx.map((i) => D.uLabels[i].replace(/^2[56]年/, "").replace("月", ""));
    s.addChart(pres.charts.LINE, [
      { name: "16—24岁（不含在校生）", labels: labs, values: idx.map((i) => D.youth[i]) },
      { name: "25—29岁", labels: labs, values: idx.map((i) => D.y2529[i]) },
      { name: "全国城镇", labels: labs, values: idx.map((i) => D.urate[i]) },
    ], co({ x: MX, y: 1.95, w: 7.9, h: 4.0, chartColors: [COL.R, COL.BLUE, COL.GRAY], lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 6, showLegend: true, legendPos: "t", legendFontSize: 13,
      valAxisMinVal: 0, valAxisMaxVal: 20, valAxisMajorUnit: 5, catAxisLabelFontSize: 12, catAxisLabelPos: "low" }));
    T(s, "2025年（月）", { x: MX + 0.6, y: 5.88, w: 3.0, h: 0.28, fontSize: 12, color: COL.GRAY, align: "center" });
    T(s, "2026年（月）", { x: MX + 3.8, y: 5.88, w: 4.0, h: 0.28, fontSize: 12, color: COL.GRAY, align: "center" });
    const px = 8.85, pw = CW - (px - MX);
    [["1270万", "2026届高校毕业生⏎比上届多48万"], ["18.9%", "8月16—24岁失业率⏎与上年同期持平"], ["+2.3%", "2025年农民工⏎月均收入增速"]].forEach((k, i) => {
      const y = 1.65 + i * 1.42;
      T(s, k[0], { x: px, y, w: pw, h: 0.65, fontSize: 34, bold: true, color: COL.R, valign: "bottom" });
      T(s, k[1], { x: px, y: y + 0.66, w: pw, h: 0.68, fontSize: 14, lineSpacingMultiple: 1.05 });
    });
    take(s, "失业率看“有没有工作”，消费更看“收入稳不稳”：**青年、灵活就业和订单型收入**的波动，会先压缩大额支出。", { y: 6.12, h: 0.66, size: 18 });
    src(s, "资料来源：国家统计局；教育部。青年失业率为不含在校生口径，8月含毕业季因素。");
  }

  // ---------- P21 预期 ----------
  {
    const s = page(3, "敢不敢花，取决于对长期收入和未来负担的判断", { min: 1.5, key: "p_expect", notes:
      "【时长】1.5分钟\n【讲法】一个很应景的例子：10月12日，2026年诺贝尔经济学奖授予斯坦福大学的罗伯特·霍尔和哥伦比亚大学的迈克尔·伍德福德，表彰他们为宏观经济稳定奠定了微观基础。霍尔1978年的经典发现是：如果家庭按长期收入安排消费，消费就只会对“关于未来收入的新消息”作出反应，可预期的、临时的变化大多被平滑掉。所以，一次性的奖金、消费券，宽裕的家庭多半存起来；但手头紧的家庭会花掉较多——发给谁很关键。真正改变换房、换车这类大额决策的，是稳定的工作和可预期的保障。伍德福德研究的则是预期怎样被政策引导。\n用到今天：消费倾向没有回到疫情前（第18页），贷款净减少（第22页），这些都是“行为”上的谨慎。\n提醒一句：“信心不足”不能当作解释一切的万能答案，它需要证据。官方连续的消费者信心调查可得性有限，这里主要用行为数据来判断。\n【衔接】存款和贷款这两个信号，具体看一看。\n【超时可压缩】前半场超时可简讲本页至30秒：只讲诺奖一句和“长期收入”。" });
    rect(s, MX, 1.62, CW, 1.2, COL.TINT);
    rect(s, MX, 1.62, 0.08, 1.2, COL.R);
    T(s, [
      { text: "2026年诺贝尔经济学奖（10月12日）", options: { bold: true, color: COL.R, fontSize: 18, breakLine: true } },
      { text: "授予霍尔（R. Hall）与伍德福德（M. Woodford），表彰其“为宏观经济稳定奠定微观基础”。霍尔的经典发现：消费只对", options: { fontSize: 17, color: COL.INK } },
      { text: "关于长期收入的新消息", options: { fontSize: 17, bold: true, color: COL.R } },
      { text: "作出反应，临时变化大多被平滑掉。", options: { fontSize: 17, color: COL.INK } },
    ], { x: MX + 0.3, y: 1.62, w: CW - 0.45, h: 1.2, valign: "middle", lineSpacingMultiple: 1.1 });
    const cw = (CW - 0.5) / 2, y = 3.05;
    [[MX, "一次性收入", "奖金、消费券、临时补贴", "宽裕家庭多存，手头紧的家庭多花：发给谁很关键", COL.GRAY], [MX + cw + 0.5, "持续性收入与保障预期", "稳定的工作、可预期的养老医疗保障", "才会改变换房、换车等大额、长期的决策", COL.R]].forEach((c) => {
      T(s, c[1], { x: c[0], y, w: cw, h: 0.45, fontSize: 20, bold: true, color: c[4] });
      T(s, c[2], { x: c[0], y: y + 0.48, w: cw, h: 0.4, fontSize: 16, color: COL.GRAY });
      T(s, "→ " + c[3], { x: c[0], y: y + 0.92, w: cw, h: 0.5, fontSize: 17 });
    });
    vline(s, MX + cw + 0.25, y + 0.05, 1.35, COL.RULE);
    T(s, "行为证据：消费倾向没有回到疫情前（第18页）；住户贷款净减少（第22页）。", { x: MX, y: 4.7, w: CW, h: 0.6, fontSize: 16, color: COL.INK, lineSpacingMultiple: 1.1 });
    take(s, "提振消费，既要让家庭对**未来收入和负担**有底，也要把钱**发给最需要的人**。", { y: 6.12, h: 0.66, size: 18 });
    src(s, "资料来源：诺贝尔奖官网公告（2026年10月12日）；中国人民银行；国家统计局。");
  }

  // ---------- P22 存贷款 ----------
  {
    const s = page(3, "居民贷款净减少、存款少增：借钱意愿明显偏弱", { min: 2, key: "p_deposit", notes:
      "【时长】2分钟\n【讲法】左图是住户存款每年的增加额：2019年不到10万亿元，2022年跳到17.8万亿元，之后每年都在14万亿元以上。今年前三季度增加9.39万亿元，比上年同期少增3.34万亿元；同期非银行业金融机构存款增加4.41万亿元——一部分钱可能从存款搬到了理财、股市。\n更值得注意的是贷款：今年1—8月住户贷款净减少1.03万亿元——借的比还的少，而去年1—8月还是增加7110亿元；到9月末，前三季度累计仍是减少1.03万亿元。\n所以今年不能简单说“居民在拼命存钱”：存款增量比去年少了。真正干净的信号是贷款——家庭借钱的意愿明显偏弱，不愿意为未来的消费和购房借钱。\n【口径】存款增量不是储蓄率；存贷款增量相减不能称为“居民净储蓄”；贷款减少不能单独证明提前还贷。\n【衔接】家庭最大的一笔资产和负债，都和房子有关。" });
    chartCap(s, "住户存款年度增加额（万亿元）", MX, 1.58, 7.4);
    s.addChart(pres.charts.BAR, [{ name: "住户存款增加额", labels: D.depYears, values: D.dep }], co({
      x: MX, y: 1.95, w: 7.4, h: 4.0, barDir: "col", chartColors: D.dep.map((v, i) => (i >= 3 ? COL.R : COL.BLUE_L)), showValue: true, dataLabelFormatCode: "0.0",
      valAxisMinVal: 0, valAxisMaxVal: 20, valAxisHidden: true, valGridLine: { style: "none" }, barGapWidthPct: 45, catAxisLabelFontSize: 14,
    }));
    const px = 8.4, pw = CW - (px - MX);
    T(s, "2026年住户存贷款", { x: px, y: 1.6, w: pw, h: 0.4, fontSize: 18, bold: true });
    [["+9.39万亿元", "住户存款增加（1—9月）⏎比上年同期少增3.34万亿元", COL.BLUE], ["-1.03万亿元", "住户贷款净减少（1—8月）⏎上年同期为增加0.71万亿元", COL.R]].forEach((k, i) => {
      const y = 2.1 + i * 1.6;
      T(s, k[0], { x: px, y, w: pw, h: 0.7, fontSize: 32, bold: true, color: k[2], valign: "bottom" });
      T(s, k[1], { x: px, y: y + 0.72, w: pw, h: 0.75, fontSize: 15, lineSpacingMultiple: 1.05 });
    });
    take(s, "**贷款净减少**是更清楚的信号：家庭借钱的意愿明显偏弱；存款少增，部分可能是资产“搬家”。", { y: 6.12, h: 0.66, size: 18 });
    src(s, "资料来源：中国人民银行金融统计数据（2026年10月14日发布）。增量不是余额，也不是储蓄率。");
  }

  // ---------- P23 房地产 ----------
  {
    const s = page(3, "房地产仍在调整：新房销售面积不到2021年的一半", { min: 2, key: "p_house", notes:
      "【时长】2分钟\n【讲法】新建商品房销售面积2021年接近18亿平方米，2025年降到8.8亿平方米，不到高点的一半；今年1—8月又下降12.1%。二手房成交在增长，1—8月网签面积增长10.6%——交易在从新房向二手房转移。\n价格方面，看浙江的三个70城样本城市：8月杭州新房同比上涨2.2%，二手房下降3.5%；宁波新房降0.8%、二手房降2.7%；温州新房降3.8%、二手房降5.3%。同一城市，新房和二手房走势也不一样。\n房子是中国家庭最大的资产。房价预期不稳，影响的不只是开发商：家庭的“财富感”、换房链条、装修家居家电，都会受影响（机制见第34页）。\n【口径】销售面积为商品房；房价为70城住宅价格指数，上年同月=100；三市不代表全省；不能用销售额除以面积冒充同质房价。\n【补充】1—8月房地产开发投资按可比口径下降19.9%，全国固定资产投资下降7.2%，其中民间投资下降10.1%。\n【衔接】另一类资产——股票——变化如何？" });
    chartCap(s, "全国新建商品房销售面积（亿平方米）", MX, 1.58, 6.9);
    s.addChart(pres.charts.BAR, [{ name: "销售面积", labels: D.houseYears, values: D.houseArea }], co({
      x: MX, y: 1.95, w: 6.9, h: 4.0, barDir: "col", chartColors: D.houseArea.map((v, i) => (i === 2 ? COL.BLUE : (i === D.houseArea.length - 1 ? COL.R : COL.BLUE_L))), showValue: true, dataLabelFormatCode: "0.0",
      valAxisMinVal: 0, valAxisMaxVal: 20, valAxisHidden: true, valGridLine: { style: "none" }, barGapWidthPct: 45,
    }));
    const px = 7.85, pw = CW - (px - MX);
    T(s, "2026年1—8月：新房销售面积 -12.1%，二手房网签面积 +10.6%", { x: px, y: 1.6, w: pw, h: 0.75, fontSize: 16, bold: true, lineSpacingMultiple: 1.05 });
    T(s, "8月住宅价格同比（%）", { x: px, y: 2.5, w: pw, h: 0.36, fontSize: 15, bold: true, color: COL.GRAY });
    table(s, [
      ["城市", "新房", "二手房"],
      ...D.city70.map((r) => [r[0], (r[1] > 0 ? "+" : "") + r[1].toFixed(1), r[2].toFixed(1)]),
    ], { x: px, y: 2.9, w: pw, colW: [pw * 0.34, pw * 0.33, pw * 0.33], rowH: 0.45, size: 16, align: ["left", "right", "right"] });
    T(s, "三市为70城样本，不代表全省", { x: px, y: 4.85, w: pw, h: 0.35, fontSize: 13, color: COL.GRAY });
    take(s, "房子是中国家庭最大的资产：价格预期不稳，家庭的**“财富感”**和**换房链条**都受影响（机制见第34页）。", { y: 6.12, h: 0.66, size: 18 });
    src(s, "资料来源：国家统计局；住房城乡建设部。房价为70个大中城市住宅销售价格指数（上年同月=100）。");
  }

  // ---------- P24 股市 ----------
  {
    const s = page(3, "股市两年回升后今年回调：财富效应有“门槛”", { min: 2, key: "p_stock", notes:
      "【时长】2分钟\n【讲法】沪深300指数：2021—2023年连跌三年，2022年跌了21.6%；2024、2025年分别上涨14.7%和17.7%；今年前三季度回调5.9%。\n股市涨了，消费会不会跟着涨？先看两个事实：一是波动大，二是覆盖窄——持股集中在部分家庭，多数家庭的资产以住房和存款为主。股市怎样影响消费，第35页展开。\n【口径】价格指数，不含分红再投资；2026年为1—9月，非全年；指数涨跌不等于居民实际收益，本页不预测行情、不作投资建议。\n【衔接】家庭能花多少钱，还要看债务和融资成本。" });
    chartCap(s, "沪深300指数年度涨跌幅（%）", MX, 1.58, 7.4);
    s.addChart(pres.charts.BAR, [{ name: "沪深300", labels: D.csiYears.map((x) => x.replace("⏎", "\n")), values: D.csi }], co({
      x: MX, y: 1.95, w: 7.4, h: 4.0, barDir: "col", chartColors: D.csi.map((v) => (v >= 0 ? COL.R : COL.BLUE)), showValue: true, dataLabelFormatCode: "0.0",
      valAxisMinVal: -25, valAxisMaxVal: 20, valAxisHidden: true, valGridLine: { style: "none" }, barGapWidthPct: 45, catAxisLabelPos: "low",
    }));
    const px = 8.4, pw = CW - (px - MX);
    T(s, "股市与家庭财富", { x: px, y: 1.6, w: pw, h: 0.42, fontSize: 18, bold: true });
    points(s, [
      ["波动大", "2022年下跌21.6%，2024、2025年回升，今年前三季度回调5.9%"],
      ["覆盖窄", "持股集中在部分家庭，多数家庭的资产以住房和存款为主"],
      ["怎样影响消费", "通过财富感和资产收入，机制见第35页"],
    ], { x: px, y: 2.15, w: pw, h: 3.8, size: 16, space: 12 });
    take(s, "股市的财富效应集中在持股家庭，要“**赚得到、拿得出、靠得住**”才会变成消费（第35页）。", { y: 6.12, h: 0.66, size: 18 });
    src(s, "资料来源：中证指数有限公司（价格指数，不含分红再投资）。2026年为1—9月。本页不构成投资建议。");
  }

  // ---------- P25 债务 ----------
  {
    const s = page(3, "利率处在低位，居民贷款仍在收缩：约束主要在需求侧", { min: 2, key: "p_loan", notes:
      "【时长】2分钟\n【讲法】按用途看二季度末的住户贷款：个人住房贷款36.29万亿元，同比下降3.8%，上半年减少7163亿元；不含住房的消费贷款20.83万亿元，下降1.7%；只有住户经营性贷款在增长，2.8%。住户贷款合计82.91万亿元，下降1.3%。\n与此同时，新发放的房贷利率约3.1%，处在历史低位。利率这么低，贷款还在收缩，说明约束主要不在“借不到”，而在“不想借、不敢借”——背后是收入预期和房价预期。\n住户经营性贷款在增长，和浙江关系很大：大量个体户、小微经营者的经营资金，是以家庭名义借的。\n【口径】个人住房贷款为人民币口径，其余两项为本外币口径；期限分类不等于用途分类；余额下降不能单独证明提前还贷。\n【衔接】同样的收入和债务变化，对不同家庭影响不同——人口结构也在改变需求。" });
    table(s, [
      ["贷款用途（2026年二季度末）", "余额", "同比"],
      ...D.loansQ2.map((r) => [r[0], r[1], { text: r[2], options: { color: r[2].startsWith("-") ? COL.R : COL.BLUE, bold: true } }]),
      ["住户贷款合计", "82.91万亿元", { text: "-1.3%", options: { color: COL.R, bold: true } }],
    ], { x: MX, y: 1.7, w: 7.6, colW: [4.0, 2.0, 1.6], rowH: [0.5, 0.62, 0.62, 0.62, 0.62], size: 16, align: ["left", "right", "right"] });
    T(s, "上半年个人住房贷款减少7163亿元", { x: MX, y: 4.85, w: 7.6, h: 0.4, fontSize: 15, color: COL.GRAY });
    const px = 8.6, pw = CW - (px - MX);
    T(s, "约3.1%", { x: px, y: 1.7, w: pw, h: 0.8, fontSize: 40, bold: true, color: COL.R, valign: "bottom" });
    T(s, "2026年6月新发放个人住房贷款利率，处于历史低位", { x: px, y: 2.55, w: pw, h: 0.75, fontSize: 15, lineSpacingMultiple: 1.05 });
    T(s, "利率低、贷款仍收缩⏎→ 约束主要不在“借不到”⏎→ 更多是“不想借、不敢借”", { x: px, y: 3.5, w: pw, h: 1.5, fontSize: 17, bold: true, lineSpacingMultiple: 1.2 });
    take(s, "约束主要在需求侧：**收入预期**和**房价预期**。住户经营贷仍在增长，与浙江的个体户、小微经营密切相关。", { y: 6.12, h: 0.66, size: 18 });
    src(s, "资料来源：中国人民银行金融机构贷款投向统计报告、货币政策执行报告。住房贷款为人民币口径，其余为本外币。");
  }

  // ---------- P26 人口 ----------
  {
    const s = page(3, "老龄化改变的是“买什么”：服务需求在升，供给要跟上", { min: 1.5, key: "p_pop", notes:
      "【时长】1.5分钟\n【讲法】60岁及以上人口占比，从2020年的18.7%升到2025年末的23.0%，年均上升近0.9个百分点。\n人口结构改变的，首先是“买什么”：老年家庭需要照护、康复、助餐、适老化改造；有孩子的家庭需要托育、教育；小家庭、独居增多，家政、维修、外卖这类“替代家务”的服务需求上升。\n老年人的消费倾向并不一定低，但他们对“风险”更敏感：看病、养老的钱要先留着。保障越不确定，越要多存。\n【口径】人口规模只能提示需求方向，不能直接乘一个价格就当成市场规模；年龄结构不是当期消费变化的全部原因。\n【衔接】有需求、有支付能力，交易为什么还是没发生？\n【超时可压缩】前半场超时可简讲本页至30秒：只讲“买什么在变”。" });
    chartCap(s, "60岁及以上人口占总人口比重（%）", MX, 1.58, 6.4);
    s.addChart(pres.charts.BAR, [{ name: "60岁及以上占比", labels: ["2020", "2021", "2022", "2023", "2024", "2025"], values: [18.7, 18.9, 19.8, 21.1, 22.0, 23.0] }], co({
      x: MX, y: 1.95, w: 6.4, h: 4.0, barDir: "col", chartColors: [COL.BLUE, COL.BLUE, COL.BLUE, COL.BLUE, COL.BLUE, COL.R], showValue: true, dataLabelFormatCode: "0.0",
      valAxisMinVal: 0, valAxisMaxVal: 25, valAxisHidden: true, valGridLine: { style: "none" }, barGapWidthPct: 45,
    }));
    const px = 7.4, pw = CW - (px - MX);
    points(s, [
      ["老年家庭", "照护、康复、助餐、适老化改造；对看病养老的“风险”更敏感，保障越不确定越要多存"],
      ["有孩家庭", "托育、教育、亲子服务；时间约束和费用约束并存"],
      ["小家庭、独居", "家政、维修、外卖等“替代家务”的服务需求上升"],
    ], { x: px, y: 1.62, w: pw, h: 4.3, size: 16, space: 14 });
    take(s, "人口结构提示的是**需求方向**，不是现成的市场规模：关键在服务供给能不能跟上。", { y: 6.12, h: 0.66, size: 18 });
    src(s, "资料来源：国家统计局第七次全国人口普查、历年国民经济和社会发展统计公报。");
  }

  // ---------- P27 交易障碍 ----------
  {
    const s = page(3, "需求已经存在，交易却卡在“找不到、信不过、退不了”", { min: 1.5, key: "p_barrier", notes:
      "【时长】1.5分钟\n【讲法】设想一个家庭，想为行动不便的老人买上门照护服务，也付得起。他们会问：附近有没有靠谱的机构？临时缺人怎么办？收费包括哪些项目，会不会临时加价？服务质量怎么核实？办了预付卡，机构关门了钱能不能退？任何一个问题答不上来，交易就可能不发生。\n右边是今年3月公布的衢州2025年度消费维权十大典型案例中的四个：口腔种植用的基台品牌和约定不符；美容祛痘套餐，调解中认定商家承诺的效果属于合同内容，退还1.15万元；保健品“会议营销”，为老年人追回10余万元；平台下单的家政服务纠纷，全额退还600元。\n这些都是“信得过”这道门槛上的问题——恰恰是市场监管的着力点。\n【口径】左侧为虚拟教学场景；投诉只提示问题，不能据此推算全体消费者的顾虑比例。\n【衔接】把第三部分的证据归纳一下。" });
    const cw = (CW - 0.5) / 2;
    subhead(s, "教学场景：为老人购买上门照护", MX, 1.6, cw);
    table(s, [
      ["家庭的需要", "付款前的顾虑"],
      ["每周固定上门，时间衔接", "附近有没有供给？临时缺人怎么办？"],
      ["稳定的护理人员", "收费含哪些项目？会不会临时加价？"],
      ["可验证的专业能力", "质量怎么核实？预付款能否退？"],
    ], { x: MX, y: 2.12, w: cw, colW: [cw * 0.45, cw * 0.55], rowH: [0.45, 0.72, 0.72, 0.72], size: 15 });
    const x2 = MX + cw + 0.5;
    subhead(s, "衢州2025年度消费维权十大典型案例（节选）", x2, 1.6, cw);
    T(s, paras([
      "**口腔医疗**：种植牙基台品牌与约定不符（开化）",
      "**美容服务**：商家承诺的祛痘效果被认定为合同内容，退还1.15万元（柯城）",
      "**老年消费**：保健品“会议营销”，为老年人追回10余万元（江山）",
      "**家政服务**：平台下单的家政服务纠纷，经调解全额退还600元",
    ], { fontSize: 16, color: COL.INK }, { space: 10 }), { x: x2, y: 2.15, w: cw, h: 3.8, lineSpacingMultiple: 1.1 });
    take(s, "有钱、有需要之后，**可得性与可信度**仍可能挡住交易——这正是市场监管的着力点（第六部分）。", { y: 6.12, h: 0.66, size: 18 });
    src(s, "资料来源：《2025年度衢州市消费维权十大典型案例》（衢州日报，2026年3月17日）。左侧为虚拟教学场景。");
  }

  // ---------- P28 证据归纳 ----------
  {
    const s = page(3, "四道门槛的证据：前两道管总量，后两道管交易", { min: 2, key: "p_evidence", notes:
      "【时长】2分钟\n【讲法】把第三部分归纳成一张表。\n“付得起”：收入增速放缓，中位数低于平均数，青年就业有压力——证据较强。\n“敢花钱”：消费倾向没回到疫情前，贷款净减少，房价还在调整——几组行为数据方向一致，证据较强。\n“买得到”：服务消费快于商品，照护、托育等供给有缺口——只有结构数据和案例，属于间接证据。\n“信得过”：预付式、价格、质量投诉高发——投诉能提示问题，但不能代表全体，也是间接证据，需要本地调查验证。\n我的判断：宏观上，收入与预期是首要约束；在我们的职责上，交易信任是能直接改善的那一块。两者不矛盾：信任改善，会让同样的收入更愿意花出去。\n【口径】仅归纳本部分证据；“证据强度”指证据的直接程度，不是各因素的贡献权重。\n【衔接】有房、有存款的家庭，为什么还推迟支出？" });
    table(s, [
      ["门槛", "本部分的证据", "证据强度", "仍需回答的问题"],
      ["① 付得起", "收入增速放缓；中位数仅为平均数的83%；青年失业率高", "较强", "哪些群体的收入稳定性承压最大？"],
      ["② 敢花钱", "消费倾向未回到疫情前；贷款净减少；房价调整", "较强", "存款少增，是资产搬家还是收入放缓？"],
      ["③ 买得到", "服务快于商品；照护、托育等供给缺口", "较弱（间接）", "本地哪些服务“有需求、无供给”？"],
      ["④ 信得过", "预付式、价格、质量等投诉高发", "较弱（间接）", "本地最高频的交易障碍是什么？"],
    ], { x: MX, y: 1.68, w: CW, colW: [1.7, 5.0, 1.8, 3.63], rowH: [0.5, 0.78, 0.78, 0.78, 0.78], size: 16, align: ["left", "left", "center", "left"], boldFirstCol: true });
    take(s, "宏观上，**收入与预期**是首要约束；在职责上，**交易信任**是市场监管能直接改善的那一块——信任改善，同样的收入更愿意花出去。", { y: 5.6, h: 0.95, size: 18 });
    src(s, "说明：仅归纳P14—P27的证据；“证据强度”指证据的直接程度，不代表各因素的贡献权重。");
  }

  // ---------- P29 家底之问 ----------
  {
    const s = page(3, "有房、有存款，为什么还推迟一笔支出？", { min: 1.5, key: "p_puzzle", notes:
      "【时长】1.5分钟\n【讲法】最后留一个问题。一个家庭：房子市值300万元，存款30万元，房贷还欠105万元，算下来净资产225万元。看起来家底不薄。这个家庭想换一辆20万元的车，却一直在犹豫。为什么？\n请大家默想：要判断这个家庭“现在敢花多少钱”，还需要知道什么？——每个月收入多少、稳不稳；每个月要还多少；未来几年有哪些必须花的大钱。\n休息之后，我们就把这几样放进同一本账里。\n【口径】虚拟教学假设，不采集学员个人财产信息，不邀请披露真实家庭账。\n【衔接】下半场：三本账。" });
    table(s, [
      ["家庭资产与负债（教学假设）", "金额"],
      ["住房市值", "300万元"],
      ["银行存款", "30万元"],
      ["房贷余额", "105万元"],
      [{ text: "净资产", options: { bold: true } }, { text: "225万元", options: { bold: true, color: COL.R } }],
    ], { x: MX, y: 1.75, w: 5.6, colW: [3.6, 2.0], rowH: 0.62, size: 18, align: ["left", "right"] });
    const px = 6.8, pw = CW - (px - MX);
    T(s, "想换一辆20万元的车，却一直犹豫。⏎仅凭“家底”，能判断这个家庭现在敢花多少钱吗？", { x: px, y: 1.75, w: pw, h: 1.3, fontSize: 20, bold: true, color: COL.R, lineSpacingMultiple: 1.15 });
    T(s, "还需要知道：", { x: px, y: 3.25, w: pw, h: 0.4, fontSize: 18, bold: true });
    T(s, paras(["每月收入多少、稳不稳", "每月要还多少钱", "未来几年有哪些必须花的大钱"], { fontSize: 18, color: COL.INK }, { space: 8 }), { x: px, y: 3.7, w: pw, h: 1.6, lineSpacingMultiple: 1.1 });
    take(s, "下半场：把**收入、现金和家底**放进同一本账里看。", { y: 6.12, h: 0.66, size: 19 });
    src(s, "说明：虚拟家庭教学假设，不采集学员个人财产信息。");
  }

  // ---------- P30 休息 ----------
  {
    L.section("休息");
    const s = pres.addSlide({ masterName: "BLANK", sectionTitle: L.curSec() });
    rect(s, 0, 0, 0.35, 7.5, COL.R);
    T(s, "休息10分钟", { x: 1.1, y: 2.0, w: 11, h: 1.2, fontSize: 54, bold: true, color: COL.R, valign: "middle" });
    T(s, "第58分钟继续", { x: 1.1, y: 3.3, w: 11, h: 0.6, fontSize: 24, color: COL.INK });
    hline(s, 1.1, 4.25, 6.5, COL.RULE, 1);
    T(s, "下半场：三本账　·　浙江观察　·　市场监管的着力点", { x: 1.1, y: 4.5, w: 11, h: 0.55, fontSize: 20, color: COL.GRAY });
    s.addNotes("【时长】10分钟（第48—58分钟）\n【安排】休息时间不放视频、不加新内容。返回后先用一句话回顾：“上半场我们看到，消费弱在大额、长期的支出上，约束集中在收入、预期和交易信任。”\n【可选】休息时可以随口问一句：“3·15为什么是3月15日？”——1962年3月15日，美国总统肯尼迪提出消费者的四项权利。");
    L.TIMING.push([pres.slides.length, 10, "休息"]);
  }
};
