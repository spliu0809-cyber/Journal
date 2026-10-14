const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/f25551c0-5fe6-4bc8-860a-eb8e37c686b3_f01c8f19-07ca-4f39-af74-65e6fa8d4fa6/pptx/scripts/apply_theme.js");

const THEME = {
  name: "NeuroVesicle",
  headFontFace: "Microsoft JhengHei",
  bodyFontFace: "Microsoft JhengHei",
  colors: {
    dk1: "16223D", // midnight navy
    lt1: "FFFFFF",
    dk2: "3A4A63",
    lt2: "EEF3F6",
    accent1: "00A896", // teal (stroke)
    accent2: "F2A14B", // amber (PD)
    accent3: "6C5CE7", // violet
    accent4: "C84B5A", // crimson (gap)
    accent5: "2E86AB",
    accent6: "8D99AE",
    hlink: "00A896",
    folHlink: "6C5CE7",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "Journal review";
pres.title = "iPS-Exosome 治療中風與帕金森氏症 文獻盤點";
const C = pres.SchemeColor;

const FOOT = "iPS-Exosome 文獻盤點｜檢索日 2026-10-14";

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "COVER",
  background: { color: THEME.colors.dk1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.9, y: 2.1, w: 9.3, h: 2.0, fontSize: 36, bold: true, color: C.background1, align: "left", valign: "bottom" }, text: " " } },
    { placeholder: { options: { name: "body", type: "body", x: 0.9, y: 4.4, w: 9.0, h: 1.5, fontSize: 16, color: THEME.colors.accent6, align: "left", valign: "top" }, text: " " } },
  ],
});

pres.defineSlideMaster({
  title: "DIVIDER",
  background: { color: THEME.colors.dk1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 1.0, y: 2.7, w: 9.0, h: 1.2, fontSize: 36, bold: true, color: C.background1, align: "left", valign: "middle" }, text: " " } },
    { placeholder: { options: { name: "body", type: "body", x: 1.0, y: 4.0, w: 8.6, h: 1.1, fontSize: 15, color: THEME.colors.accent6, align: "left", valign: "top" }, text: " " } },
  ],
});

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.34, w: 12.1, h: 0.72, fontSize: 28, bold: true, color: C.text1, align: "left", valign: "middle" }, text: " " } },
    { text: { text: FOOT, options: { x: 0.6, y: 6.95, w: 8.0, h: 0.3, fontSize: 10, color: THEME.colors.accent6, isTextBox: true, margin: 0 } } },
  ],
  slideNumber: { x: 12.5, y: 6.95, w: 0.5, h: 0.3, fontSize: 10, color: THEME.colors.accent6, align: "right" },
});

pres.defineSlideMaster({
  title: "TABLE",
  background: { color: C.background1 },
  margin: [0.5, 0.5, 0.9, 0.5],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.34, w: 12.1, h: 0.72, fontSize: 28, bold: true, color: C.text1, align: "left", valign: "middle" }, text: " " } },
    { placeholder: { options: { name: "table", type: "body", x: 0.6, y: 1.35, w: 12.1, h: 5.2, color: C.text1 }, text: " " } },
    { text: { text: FOOT, options: { x: 0.6, y: 6.95, w: 8.0, h: 0.3, fontSize: 10, color: THEME.colors.accent6, isTextBox: true, margin: 0 } } },
  ],
  slideNumber: { x: 12.5, y: 6.95, w: 0.5, h: 0.3, fontSize: 10, color: THEME.colors.accent6, align: "right" },
});

pres.defineSlideMaster({
  title: "CHART",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.34, w: 12.1, h: 0.72, fontSize: 28, bold: true, color: C.text1, align: "left", valign: "middle" }, text: " " } },
    { placeholder: { options: { name: "chart", type: "chart", x: 0.6, y: 1.4, w: 7.5, h: 5.0, color: C.text1 }, text: " " } },
    { placeholder: { options: { name: "body", type: "body", x: 8.4, y: 1.4, w: 4.3, h: 5.0, fontSize: 14, color: C.text1, valign: "top" }, text: " " } },
    { text: { text: FOOT, options: { x: 0.6, y: 6.95, w: 8.0, h: 0.3, fontSize: 10, color: THEME.colors.accent6, isTextBox: true, margin: 0 } } },
  ],
  slideNumber: { x: 12.5, y: 6.95, w: 0.5, h: 0.3, fontSize: 10, color: THEME.colors.accent6, align: "right" },
});

pres.defineSlideMaster({
  title: "CLOSING",
  background: { color: THEME.colors.dk1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.75, y: 0.5, w: 11.8, h: 0.8, fontSize: 30, bold: true, color: C.background1, align: "left", valign: "middle" }, text: " " } },
  ],
});

// ---------- helpers ----------
// Repeated motif: a numbered disc + label
function disc(slide, x, y, label, fill, size) {
  const d = size || 0.52;
  slide.addShape(pres.ShapeType.ellipse, {
    x, y, w: d, h: d, fill: { color: fill },
    objectName: "disc-" + label,
  });
  slide.addText(label, {
    x, y, w: d, h: d, align: "center", valign: "middle",
    fontSize: d > 0.6 ? 16 : 13, bold: true, color: "FFFFFF", isTextBox: true, margin: 0,
  });
}

function card(slide, o) {
  slide.addShape(pres.ShapeType.roundRect, {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.08,
    fill: { color: o.fill || THEME.colors.lt2 },
    line: { color: o.fill || THEME.colors.lt2, width: 0.5 },
    objectName: o.name,
  });
}

const TBL_HEAD = (txt) => ({ text: txt, options: { bold: true, color: "FFFFFF", fill: { color: THEME.colors.dk1 }, fontSize: 12, valign: "middle", align: "center" } });
const cell = (txt, opts) => ({ text: txt, options: Object.assign({ fontSize: 11, color: THEME.colors.dk1, valign: "top" }, opts || {}) });

// ============ 1. Cover ============
pres.addSection({ title: "封面" });
{
  const s = pres.addSlide({ masterName: "COVER", sectionTitle: "封面" });
  // motif: cluster of vesicle discs
  const vesicles = [
    [10.95, 1.5, 1.5], [12.25, 2.75, 0.85], [10.35, 3.15, 1.05],
    [11.75, 4.3, 1.25], [10.1, 5.1, 0.7], [12.4, 5.55, 0.95],
  ];
  const vcol = [THEME.colors.accent1, THEME.colors.accent2, THEME.colors.accent3, THEME.colors.accent1, THEME.colors.accent5, THEME.colors.accent2];
  vesicles.forEach((v, i) => {
    s.addShape(pres.ShapeType.ellipse, {
      x: v[0], y: v[1], w: v[2], h: v[2],
      fill: { color: vcol[i], transparency: 25 },
      line: { color: vcol[i], width: 1.25 },
      objectName: "vesicle-" + i,
    });
  });

  s.addText("PubMed 文獻盤點", {
    x: 0.9, y: 1.65, w: 6.0, h: 0.45, fontSize: 15, bold: true,
    color: THEME.colors.accent1, charSpacing: 2, isTextBox: true, margin: 0,
  });
  s.addText("iPS-Exosome\n治療中風與帕金森氏症", { placeholder: "title", lineSpacingMultiple: 1.15 });
  s.addText(
    [
      { text: "逐篇文獻整理、機制歸納與證據缺口分析", options: { breakLine: true, fontSize: 17, color: "FFFFFF" } },
      { text: "檢索日：2026-10-14　｜　中風 6 筆核心 + 3 筆旁證　｜　帕金森 3 筆核心 + 4 筆旁證", options: { fontSize: 13 } },
    ],
    { placeholder: "body" }
  );
  s.addNotes("本次環境無法直連 PubMed E-utilities（組織網路政策封鎖 NCBI／EuropePMC／OpenAlex），改以網路檢索逐篇核對 PubMed/PMC 書目資料。標註「摘要層級」者僅取得摘要資訊。");
}

// ============ 2. 執行摘要 ============
pres.addSection({ title: "執行摘要" });
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "執行摘要" });
  s.addText("兩個疾病方向的證據成熟度差距極大", { placeholder: "title" });

  const stats = [
    { n: "4", l: "中風：同儕審查\n活體前臨床研究", c: THEME.colors.accent1 },
    { n: "2", l: "中風：人體臨床\n試驗／註冊案", c: THEME.colors.accent5 },
    { n: "1", l: "帕金森：同儕審查\n研究（僅活體外）", c: THEME.colors.accent2 },
    { n: "0", l: "帕金森：同儕審查\n活體療效研究", c: THEME.colors.accent4 },
  ];
  stats.forEach((st, i) => {
    const x = 0.6 + i * 3.1;
    card(s, { x, y: 1.35, w: 2.85, h: 1.95, fill: THEME.colors.lt2, name: "stat-" + i });
    s.addText(st.n, { x, y: 1.45, w: 2.85, h: 1.0, fontSize: 54, bold: true, color: st.c, align: "center", valign: "middle", isTextBox: true, margin: 0 });
    s.addText(st.l, { x: x + 0.15, y: 2.45, w: 2.55, h: 0.75, fontSize: 12, color: THEME.colors.dk2, align: "center", valign: "top", isTextBox: true, margin: 0 });
  });

  const points = [
    { t: "中風端已完成轉譯接續", d: "大鼠／小鼠 MCAO 模型療效一致，並已進入人體安全性試驗（ESOC 2026 摘要、NCT07145294 招募中）。" },
    { t: "主力來源是 iMSC 而非 iPSC 本體", d: "兩篇機制最強的中風研究皆使用 iPSC→MSC（iMSC）衍生 sEV，共同終點收斂於血管新生。" },
    { t: "免疫與腦腸軸為第二機制主線", d: "Treg／Th17 調節與腸道菌相調節，但兩篇皆為電針併用設計，iPSC-sEV 單獨貢獻未被切分。" },
    { t: "帕金森是明確空白區", d: "尚無已同儕審查、以純化 iPSC 外泌體在 MPTP／6-OHDA／A53T 活體模型證實療效的論文。" },
  ];
  points.forEach((p, i) => {
    const y = 3.62 + i * 0.84;
    disc(s, 0.6, y, String(i + 1), i === 3 ? THEME.colors.accent4 : THEME.colors.accent1, 0.44);
    s.addText(p.t, { x: 1.2, y: y - 0.04, w: 3.6, h: 0.3, fontSize: 14, bold: true, color: THEME.colors.dk1, isTextBox: true, margin: 0 });
    s.addText(p.d, { x: 4.9, y: y - 0.06, w: 7.8, h: 0.68, fontSize: 12, color: THEME.colors.dk2, isTextBox: true, margin: 0 });
  });
  s.addNotes("一句話結論：中風方向可直接談臨床轉譯，帕金森方向目前是研究機會而非既有證據。");
}

// ============ 3. 檢索策略 ============
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "執行摘要" });
  s.addText("檢索策略與納入／排除原則", { placeholder: "title" });

  const cols = [
    {
      h: "檢索範圍", c: THEME.colors.accent1,
      rows: ["PubMed／PMC 索引書目", "ClinicalTrials.gov 註冊試驗", "ESOC 2026 會議摘要", "bioRxiv 預印本（標註）"],
    },
    {
      h: "關鍵詞組合", c: THEME.colors.accent5,
      rows: ["iPSC / induced pluripotent stem cell", "exosome / small extracellular vesicle", "stroke / MCAO / cerebral ischemia", "Parkinson / MPTP / 6-OHDA / α-synuclein"],
    },
    {
      h: "納入條件", c: THEME.colors.accent3,
      rows: ["外泌體／sEV 來源為 iPSC 或 iPSC 衍生細胞", "以治療為目的（非僅生物標記）", "具疾病模型或人體受試者", "可取得書目與摘要層級結果"],
    },
    {
      h: "列為旁證", c: THEME.colors.accent6,
      rows: ["非 iPSC 來源（BM-MSC、UC-MSC、F3 NSC）", "非腦部缺血（如缺血肢體）", "致病機制方向而非治療", "回顧文獻與統合分析"],
    },
  ];
  cols.forEach((col, i) => {
    const x = 0.6 + i * 3.1;
    card(s, { x, y: 1.3, w: 2.85, h: 3.6, fill: THEME.colors.lt2, name: "col-" + i });
    disc(s, x + 0.2, 1.52, String(i + 1), col.c, 0.44);
    s.addText(col.h, { x: x + 0.78, y: 1.5, w: 1.9, h: 0.45, fontSize: 15, bold: true, color: col.c, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(
      col.rows.map((r, j) => ({ text: r, options: { bullet: true, breakLine: j !== col.rows.length - 1, paraSpaceAfter: 8 } })),
      { x: x + 0.22, y: 2.2, w: 2.45, h: 2.6, fontSize: 12, color: THEME.colors.dk2, valign: "top", isTextBox: true, margin: 0 }
    );
  });

  card(s, { x: 0.6, y: 5.25, w: 12.1, h: 1.4, fill: "FBEAEA", name: "limit-note" });
  s.addText(
    [
      { text: "執行限制：", options: { bold: true, color: THEME.colors.accent4 } },
      { text: "本環境之組織網路政策封鎖 NCBI E-utilities／EuropePMC／OpenAlex 直連，書目資料以網路檢索逐篇核對；標註「摘要層級」者未取得全文統計數據。", options: { color: THEME.colors.dk2 } },
    ],
    { x: 0.95, y: 5.45, w: 11.4, h: 1.0, fontSize: 13, valign: "middle", isTextBox: true, margin: 0 }
  );
  s.addNotes("若需補齊全文層級的效應量（梗塞體積百分比、行為學分數），建議在可連 PubMed 的環境重跑檢索並抓取全文。");
}

// ============ 4. Divider 中風 ============
pres.addSection({ title: "中風" });
{
  const s = pres.addSlide({ masterName: "DIVIDER", sectionTitle: "中風" });
  disc(s, 1.0, 1.75, "01", THEME.colors.accent1, 0.75);
  s.addText("缺血性腦中風", { placeholder: "title" });
  s.addText("4 篇同儕審查前臨床研究、2 筆人體試驗、3 筆高度相關旁證", { placeholder: "body" });
  [[11.0, 2.0, 1.6], [12.2, 3.6, 0.9], [10.6, 4.4, 1.2]].forEach((v, i) => {
    s.addShape(pres.ShapeType.ellipse, {
      x: v[0], y: v[1], w: v[2], h: v[2],
      fill: { color: THEME.colors.accent1, transparency: 30 },
      line: { color: THEME.colors.accent1, width: 1.25 },
      objectName: "div1-vesicle-" + i,
    });
  });
}

// ============ 5. 中風 表格 ============
{
  const s = pres.addSlide({ masterName: "TABLE", sectionTitle: "中風" });
  s.addText("中風：逐篇前臨床研究", { placeholder: "title" });
  const rows = [
    [TBL_HEAD("#"), TBL_HEAD("文獻／期刊・年"), TBL_HEAD("外泌體來源"), TBL_HEAD("模型"), TBL_HEAD("主要結果"), TBL_HEAD("機制")],
    [
      cell("S1", { bold: true, color: THEME.colors.accent1, align: "center" }),
      cell("Xia Y, et al.\nStem Cell Res Ther 2020;11:313\nPMC7374834"),
      cell("hiPSC→MSC\n（iMSC）small EV"),
      cell("大鼠 MCAO\n＋ 內皮細胞活體外"),
      cell("梗塞體積顯著縮小、血管新生增加、長期神經功能缺損改善；內皮細胞遷移與管腔形成增強"),
      cell("活化 STAT3 → 抑制中風誘發自噬（LC3-II/I、Beclin-1、P62）；STAT3 抑制劑 stattic 可部分逆轉"),
    ],
    [
      cell("S2", { bold: true, color: THEME.colors.accent1, align: "center" }),
      cell("Lu G, Su X, … Chan WY.\nBiomedicines 2023;11(9):2550\nPMC10525838"),
      cell("hiPS-MSC-EV\n（超過濾分離）"),
      cell("小鼠 MCAO，靜脈\n100 µg × 3 時間點\n＋ HT22 OGD/R"),
      cell("梗塞體積縮小、自發活動能力改善；活體外促進 HT22 增生、降低凋亡"),
      cell("梗塞側腦組織 VEGF、CXCR4 表現上升 → 血管新生與神經修復"),
    ],
    [
      cell("S3", { bold: true, color: THEME.colors.accent1, align: "center" }),
      cell("Deng P, et al.\nCells 2022;11(5):820\n（2024 Correction）"),
      cell("人類 iPSC\nsmall EV"),
      cell("小鼠 MCAO\n電針（EA）併用"),
      cell("神經學評分改善、梗塞範圍縮小、神經元凋亡降低"),
      cell("抑制 Th1／Th17、促進 Treg；抑制微膠細胞與星狀細胞活化；涉及 IL-33/ST2 軸"),
    ],
    [
      cell("S4", { bold: true, color: THEME.colors.accent1, align: "center" }),
      cell("Zhang Q, et al.\nFront Immunol 2023;14:1107559\nPMC9895941"),
      cell("人類 iPSC\nsmall EV"),
      cell("小鼠 MCAO\n電針（EA）併用"),
      cell("神經功能顯著改善、腦與腸道損傷同步減輕"),
      cell("經腦—腸軸調節腸道菌相組成，連動腦／結腸發炎指標"),
    ],
  ];
  s.addTable(rows, {
    placeholder: "table", x: 0.6, y: 1.35, w: 12.1,
    colW: [0.5, 2.5, 1.5, 1.8, 2.9, 2.9],
    border: { type: "solid", color: "D5DEE4", pt: 0.75 },
    fill: { color: "FFFFFF" },
    rowH: [0.42, 1.2, 1.2, 1.05, 0.95],
    margin: 0.07,
    autoPage: false,
  });
  s.addNotes("S1 與 S2 是兩篇機制最完整的研究，都指向血管新生。S3／S4 來自同一研究團隊（孫華、張建民），為電針併用設計。");
}

// ============ 6. 中風 機制整合 ============
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "中風" });
  s.addText("中風療效的三條機制主線", { placeholder: "title" });

  const mechs = [
    {
      h: "血管新生", sub: "S1 · S2", c: THEME.colors.accent1,
      items: ["STAT3 活化 → 自噬抑制", "VEGF ／ CXCR4 表現上升", "內皮細胞遷移與管腔形成", "終點：梗塞體積縮小"],
    },
    {
      h: "免疫調節", sub: "S3", c: THEME.colors.accent3,
      items: ["Th1 ／ Th17 反應抑制", "調節型 T 細胞（Treg）增加", "微膠細胞／星狀細胞活化下降", "IL-33 ／ ST2 軸參與"],
    },
    {
      h: "腦—腸軸", sub: "S4", c: THEME.colors.accent2,
      items: ["腸道菌相組成改變", "腦與結腸發炎指標連動", "腸道損傷同步減輕", "機制細節仍待釐清"],
    },
  ];
  mechs.forEach((m, i) => {
    const x = 0.6 + i * 4.15;
    card(s, { x, y: 1.3, w: 3.85, h: 3.0, fill: THEME.colors.lt2, name: "mech-" + i });
    disc(s, x + 0.25, 1.55, String(i + 1), m.c, 0.5);
    s.addText(m.h, { x: x + 0.88, y: 1.5, w: 2.2, h: 0.4, fontSize: 18, bold: true, color: m.c, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(m.sub, { x: x + 0.88, y: 1.9, w: 2.2, h: 0.28, fontSize: 11, italic: true, color: THEME.colors.accent6, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(
      m.items.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j !== m.items.length - 1, paraSpaceAfter: 10 } })),
      { x: x + 0.28, y: 2.35, w: 3.3, h: 1.9, fontSize: 12.5, color: THEME.colors.dk2, valign: "top", isTextBox: true, margin: 0 }
    );
    if (i < 2) {
      s.addShape(pres.ShapeType.rightArrow, {
        x: x + 3.93, y: 2.6, w: 0.28, h: 0.4,
        fill: { color: THEME.colors.accent6 }, line: { color: THEME.colors.accent6, width: 0 },
        objectName: "arrow-" + i,
      });
    }
  });

  card(s, { x: 0.6, y: 4.6, w: 12.1, h: 1.75, fill: "E7F6F3", name: "mech-takeaway" });
  disc(s, 0.95, 4.95, "★", THEME.colors.accent1, 0.5);
  s.addText("方法學限制必須明講", { x: 1.6, y: 4.92, w: 4.0, h: 0.38, fontSize: 15, bold: true, color: THEME.colors.dk1, isTextBox: true, margin: 0 });
  s.addText(
    "S3／S4 皆為「電針 ＋ iPSC-sEV」併用設計，未設 iPSC-sEV 單獨治療組，因此免疫與腦腸軸這兩條機制中 sEV 的獨立貢獻尚未被切分；血管新生這條線則有 S1／S2 兩個獨立團隊、兩種物種（大鼠、小鼠）支持，可信度最高。",
    { x: 1.6, y: 5.35, w: 10.8, h: 0.9, fontSize: 13, color: THEME.colors.dk2, valign: "top", isTextBox: true, margin: 0 }
  );
}

// ============ 7. 臨床進展 ============
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "中風" });
  s.addText("中風：已進入人體階段", { placeholder: "title" });

  const trials = [
    {
      tag: "S5", name: "ESOC 2026 Abstract OT115", c: THEME.colors.accent5,
      rows: [
        ["類型", "同種異體 hiPSC-exosome 臨床安全性／有效性試驗"],
        ["收案條件", "皮質型缺血性中風、基礎 NIHSS 6–20、發病 1–7 天內"],
        ["主要終點", "90 天內嚴重不良事件（SAE）"],
        ["現況", "摘要層級，尚無結果數據（PMC13144839）"],
      ],
    },
    {
      tag: "S6", name: "NCT07145294", c: THEME.colors.accent3,
      rows: [
        ["類型", "hiPSC 衍生 small EV 治療急性缺血性中風"],
        ["主辦", "上海市第六人民醫院"],
        ["評估重點", "神經功能促進與認知功能障礙改善"],
        ["現況", "招募中（2025-08 最後確認）"],
      ],
    },
  ];
  trials.forEach((t, i) => {
    const x = 0.6 + i * 6.3;
    card(s, { x, y: 1.4, w: 5.8, h: 3.3, fill: THEME.colors.lt2, name: "trial-" + i });
    disc(s, x + 0.28, 1.65, t.tag, t.c, 0.5);
    s.addText(t.name, { x: x + 0.92, y: 1.62, w: 4.6, h: 0.42, fontSize: 17, bold: true, color: t.c, valign: "middle", isTextBox: true, margin: 0 });
    t.rows.forEach((r, j) => {
      const y = 2.3 + j * 0.56;
      s.addText(r[0], { x: x + 0.3, y, w: 1.0, h: 0.5, fontSize: 11.5, bold: true, color: THEME.colors.dk1, valign: "top", isTextBox: true, margin: 0 });
      s.addText(r[1], { x: x + 1.35, y, w: 4.2, h: 0.5, fontSize: 11.5, color: THEME.colors.dk2, valign: "top", isTextBox: true, margin: 0 });
    });
  });

  s.addText("高度相關旁證（非腦部 iPSC 外泌體本體，但支撐同一技術路線）", {
    x: 0.6, y: 4.95, w: 12.1, h: 0.35, fontSize: 15, bold: true, color: THEME.colors.dk1, isTextBox: true, margin: 0,
  });
  const side = [
    { tag: "S7", t: "iPSC→內皮細胞外泌體促進缺血肢體血管新生（PMC6329927）", d: "部位為肢體而非腦，但直接支持血管新生機制" },
    { tag: "S8", t: "誘導型 NSC／NPC 外泌體促進中風後恢復（PMC9217134）", d: "小鼠 MCAO 感覺動作恢復與 NPC-EV 相當" },
    { tag: "S9", t: "幹細胞外泌體治療中風之網絡統合分析", d: "動物模型中 iPSC 外泌體顯著降低梗塞體積與神經缺損評分；靜脈給藥為最佳途徑" },
  ];
  side.forEach((v, i) => {
    const y = 5.42 + i * 0.5;
    disc(s, 0.6, y, v.tag, THEME.colors.accent6, 0.4);
    s.addText(v.t, { x: 1.12, y: y - 0.02, w: 5.6, h: 0.42, fontSize: 11.5, bold: true, color: THEME.colors.dk1, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(v.d, { x: 6.85, y: y - 0.02, w: 5.85, h: 0.42, fontSize: 11.5, color: THEME.colors.dk2, valign: "middle", isTextBox: true, margin: 0 });
  });
}

// ============ 8. Divider PD ============
pres.addSection({ title: "帕金森氏症" });
{
  const s = pres.addSlide({ masterName: "DIVIDER", sectionTitle: "帕金森氏症" });
  disc(s, 1.0, 1.75, "02", THEME.colors.accent2, 0.75);
  s.addText("帕金森氏症", { placeholder: "title" });
  s.addText("3 筆核心文獻：1 篇活體外對比、1 篇預印本、1 篇會議摘要", { placeholder: "body" });
  [[11.0, 2.0, 1.6], [12.2, 3.6, 0.9], [10.6, 4.4, 1.2]].forEach((v, i) => {
    s.addShape(pres.ShapeType.ellipse, {
      x: v[0], y: v[1], w: v[2], h: v[2],
      fill: { color: THEME.colors.accent2, transparency: 30 },
      line: { color: THEME.colors.accent2, width: 1.25 },
      objectName: "div2-vesicle-" + i,
    });
  });
}

// ============ 9. PD 表格 ============
{
  const s = pres.addSlide({ masterName: "TABLE", sectionTitle: "帕金森氏症" });
  s.addText("帕金森氏症：逐篇文獻", { placeholder: "title" });
  const rows = [
    [TBL_HEAD("#"), TBL_HEAD("文獻／來源・年"), TBL_HEAD("外泌體來源"), TBL_HEAD("模型"), TBL_HEAD("主要結果"), TBL_HEAD("證據等級")],
    [
      cell("P1", { bold: true, color: THEME.colors.accent2, align: "center" }),
      cell("A Comparative Analysis of iPSC, MSC, and 293T-Derived Exosomes for Parkinson's Disease Therapy\nInt J Mol Sci 2026;27(10):4212\nPMID 42196195"),
      cell("iPSC vs MSC\nvs 293T 三方對比"),
      cell("MPP⁺ 誘導\nSH-SY5Y 細胞"),
      cell("iPSC 衍生外泌體在所評估各項指標上均優於 MSC 與 293T；可減輕 MPP⁺ 誘導之細胞凋亡。母細胞來源決定囊泡貨物特性"),
      cell("同儕審查\n但僅活體外；\n單一研究未重複", { color: THEME.colors.accent2 }),
    ],
    [
      cell("P2", { bold: true, color: THEME.colors.accent2, align: "center" }),
      cell("iPSC-derived MSC secretome activates distinct neuroprotective pathways in a preclinical model of PD\nbioRxiv 2025\n10.1101/2025.11.07.687231"),
      cell("hiPSC→MSC（iMSC）\n分泌體（含 EV），\n對照骨髓 MSC"),
      cell("PD 活體\n動物模型"),
      cell("iMSC 組運動與非運動功能改善較明顯、黑質緻密部多巴胺神經元保存較佳；兩種分泌體蛋白質體圖譜不同"),
      cell("預印本\n未經同儕審查；\n為分泌體非純化 EV", { color: THEME.colors.accent4 }),
    ],
    [
      cell("P3", { bold: true, color: THEME.colors.accent2, align: "center" }),
      cell("iPSC-derived extracellular vesicles rescue Parkinson's-associated deficits\nFENS Forum 2026 poster\nUniversity of Haifa"),
      cell("iPSC 衍生 EV"),
      cell("A53T SNCA 病人\niPSC 多巴胺神經元\n＋ 小鼠模型"),
      cell("報告可回復 PD 相關缺陷；完整數據與效應量尚未發表"),
      cell("會議摘要\n無同儕審查全文", { color: THEME.colors.accent4 }),
    ],
  ];
  s.addTable(rows, {
    placeholder: "table", x: 0.6, y: 1.35, w: 12.1,
    colW: [0.5, 3.1, 1.85, 1.75, 3.1, 1.8],
    border: { type: "solid", color: "D5DEE4", pt: 0.75 },
    fill: { color: "FFFFFF" },
    rowH: [0.42, 1.55, 1.55, 1.3],
    margin: 0.07,
    autoPage: false,
  });
  s.addNotes("P1 是目前唯一同儕審查、且明確比較 iPSC 外泌體與其他來源的 PD 研究，但僅止於 SH-SY5Y 細胞層級。");
}

// ============ 10. PD 旁證與缺口 ============
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "帕金森氏症" });
  s.addText("帕金森氏症：周邊文獻與證據缺口", { placeholder: "title" });

  card(s, { x: 0.6, y: 1.4, w: 6.0, h: 4.05, fill: THEME.colors.lt2, name: "pd-side" });
  s.addText("相關但不屬於「iPSC 外泌體治療」", { x: 0.85, y: 1.6, w: 5.5, h: 0.35, fontSize: 15, bold: true, color: THEME.colors.dk1, isTextBox: true, margin: 0 });
  const pdSide = [
    { tag: "P4", t: "de Rus Jacquet A, et al. eLife 2021", d: "LRRK2 G2019S 突變 iPSC 星狀細胞 EV 神經營養支持不足 → 神經元萎縮。屬致病方向，非治療" },
    { tag: "P5", t: "Lee EJ, et al. J Nanobiotechnol 2022", d: "F3 人類 NSC 細胞株（非 iPSC）EV；6-OHDA 模型保護多巴胺神經元、降低 ROS 與發炎細胞激素" },
    { tag: "P6", t: "Glial EV 雙面角色回顧 Int J Mol Sci 2026", d: "彙整 iPSC 膠細胞模型中 EV 可為保護亦可為有害，PD 突變會改變 EV 貨物組成" },
    { tag: "P7", t: "多篇 MSC 外泌體 PD 研究 2020–2025", d: "UC-MSC、脂肪 MSC miR-188-3p、MSC miR-100a-5p 等，可作為效價比較的非 iPSC 基準線" },
  ];
  pdSide.forEach((v, i) => {
    const y = 2.1 + i * 0.84;
    disc(s, 0.85, y, v.tag, THEME.colors.accent6, 0.42);
    s.addText(v.t, { x: 1.4, y: y - 0.02, w: 4.95, h: 0.3, fontSize: 12.5, bold: true, color: THEME.colors.dk1, isTextBox: true, margin: 0 });
    s.addText(v.d, { x: 1.4, y: y + 0.26, w: 4.95, h: 0.56, fontSize: 11, color: THEME.colors.dk2, valign: "top", isTextBox: true, margin: 0 });
  });

  card(s, { x: 6.9, y: 1.4, w: 5.8, h: 4.05, fill: "FBEAEA", name: "pd-gap" });
  disc(s, 7.15, 1.6, "!", THEME.colors.accent4, 0.5);
  s.addText("四個具體缺口", { x: 7.78, y: 1.58, w: 4.6, h: 0.42, fontSize: 17, bold: true, color: THEME.colors.accent4, valign: "middle", isTextBox: true, margin: 0 });
  const gaps = [
    "缺活體療效：沒有以純化 iPSC 外泌體在 MPTP／6-OHDA／A53T 動物模型完成行為學與組織學驗證的同儕審查論文",
    "缺劑量與給藥途徑：中風端已確立靜脈給藥，PD 端的靜脈 vs 鼻腔 vs 腦內定位注射尚無比較",
    "缺機制歸因：P1 僅報告「優於 MSC／293T」，未指認具體 miRNA／蛋白質貨物與下游路徑",
    "缺複製驗證：三筆核心文獻分屬三個團隊、三種證據型態，彼此無法互相印證",
  ];
  gaps.forEach((g, i) => {
    const y = 2.22 + i * 0.78;
    disc(s, 7.15, y, String(i + 1), THEME.colors.accent4, 0.4);
    s.addText(g, { x: 7.7, y: y - 0.04, w: 4.75, h: 0.72, fontSize: 11.5, color: THEME.colors.dk2, valign: "top", isTextBox: true, margin: 0 });
  });

  card(s, { x: 0.6, y: 5.7, w: 12.1, h: 0.95, fill: "E7F6F3", name: "pd-opportunity" });
  s.addText(
    [
      { text: "轉譯機會：", options: { bold: true, color: THEME.colors.accent1 } },
      { text: "中風端已驗證的 iMSC-sEV 製程（超過濾分離、靜脈 100 µg 多次給藥）與血管新生／免疫調節讀值，可直接移植到 PD 活體模型，補上目前最大的空白。", options: { color: THEME.colors.dk2 } },
    ],
    { x: 0.9, y: 5.85, w: 11.5, h: 0.7, fontSize: 13, valign: "middle", isTextBox: true, margin: 0 }
  );
}

// ============ 11. Chart 比較 ============
pres.addSection({ title: "綜合比較" });
{
  const s = pres.addSlide({ masterName: "CHART", sectionTitle: "綜合比較" });
  s.addText("兩個方向的證據量對比", { placeholder: "title" });

  s.addChart(
    [
      {
        type: pres.ChartType.bar,
        data: [
          { name: "中風", labels: ["同儕審查活體研究", "活體外研究", "人體試驗／註冊", "預印本／會議摘要", "高度相關旁證"], values: [4, 2, 2, 1, 3] },
          { name: "帕金森氏症", labels: ["同儕審查活體研究", "活體外研究", "人體試驗／註冊", "預印本／會議摘要", "高度相關旁證"], values: [0, 1, 0, 2, 4] },
        ],
      },
    ],
    {
      placeholder: "chart",
      barDir: "bar", barGrouping: "clustered", barGapWidthPct: 55,
      chartColors: [THEME.colors.accent1, THEME.colors.accent2],
      showTitle: true, title: "依證據型態的文獻筆數", titleColor: THEME.colors.dk1,
      titleFontSize: 14, titleFontFace: "+mn-lt",
      showValue: true, dataLabelPosition: "outEnd", dataLabelColor: THEME.colors.dk2,
      dataLabelFontSize: 11, dataLabelFontFace: "+mn-lt",
      showLegend: true, legendPos: "b", legendColor: THEME.colors.dk2, legendFontSize: 11, legendFontFace: "+mn-lt",
      catAxisLabelColor: THEME.colors.dk2, catAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt",
      valAxisLabelColor: THEME.colors.dk2, valAxisLabelFontSize: 11, valAxisLabelFontFace: "+mn-lt",
      valAxisMaxVal: 5, valAxisMajorUnit: 1,
      valGridLine: { color: "E3E9ED", size: 0.75 },
      catGridLine: { style: "none" },
      catAxisLineShow: false, valAxisLineShow: false,
    }
  );

  s.addText(
    [
      { text: "讀圖重點", options: { bold: true, fontSize: 16, color: THEME.colors.dk1, breakLine: true, paraSpaceAfter: 10 } },
      { text: "中風方向在「活體療效」與「人體試驗」兩個關鍵層級都有實體文獻，是可以對外談轉譯進度的方向。", options: { bullet: true, breakLine: true, paraSpaceAfter: 12 } },
      { text: "帕金森方向的筆數集中在旁證與未審查文獻，核心格（活體療效、人體試驗）為零。", options: { bullet: true, breakLine: true, paraSpaceAfter: 12 } },
      { text: "兩者唯一重疊的細胞來源是 iPSC→MSC（iMSC）：中風端 S1／S2 已證實療效，PD 端僅有 P2 預印本。", options: { bullet: true, breakLine: true, paraSpaceAfter: 12 } },
      { text: "結論：PD 的 iPSC 外泌體不是「證據不足」，而是「尚未被做」。", options: { bullet: true, bold: true, color: THEME.colors.accent4 } },
    ],
    { placeholder: "body" }
  );
  s.addNotes("筆數依本次盤點之納入標準計算；旁證包含非 iPSC 來源與非腦部缺血研究。");
}

// ============ 12. 結論 ============
{
  const s = pres.addSlide({ masterName: "CLOSING", sectionTitle: "綜合比較" });
  s.addText("結論與建議下一步", { placeholder: "title" });

  const concl = [
    { h: "可直接引用", c: THEME.colors.accent1, items: ["iMSC-sEV 於大鼠／小鼠 MCAO 縮小梗塞體積（S1、S2）", "機制收斂於血管新生：STAT3-自噬軸、VEGF／CXCR4", "已有人體安全性試驗進行中（S5、S6）"] },
    { h: "引用需加註", c: THEME.colors.accent2, items: ["S3／S4 為電針併用，sEV 單獨貢獻未切分", "P1 僅 SH-SY5Y 活體外，且為單一研究", "P2 為預印本、P3 為會議摘要"] },
    { h: "建議下一步", c: THEME.colors.accent3, items: ["以中風端 iMSC-sEV 製程切入 6-OHDA／MPTP 活體模型", "設 iPSC-sEV 單獨治療組以切分效果", "補做 miRNA／蛋白質貨物分析指認機制", "在可連 PubMed 環境重跑檢索補齊全文效應量"] },
  ];
  concl.forEach((col, i) => {
    const x = 0.75 + i * 4.1;
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 1.5, w: 3.8, h: 3.6, rectRadius: 0.08,
      fill: { color: "FFFFFF", transparency: 92 },
      line: { color: col.c, width: 1.25 },
      objectName: "concl-card-" + i,
    });
    disc(s, x + 0.28, 1.78, String(i + 1), col.c, 0.5);
    s.addText(col.h, { x: x + 0.9, y: 1.75, w: 2.6, h: 0.42, fontSize: 18, bold: true, color: col.c, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(
      col.items.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j !== col.items.length - 1, paraSpaceAfter: 12 } })),
      { x: x + 0.3, y: 2.45, w: 3.25, h: 2.5, fontSize: 12.5, color: "DCE4EA", valign: "top", isTextBox: true, margin: 0 }
    );
  });

  s.addText("完整逐篇表格（含 DOI／PMCID 全欄位）：papers.md", {
    x: 0.75, y: 5.45, w: 11.8, h: 0.4, fontSize: 12, italic: true, color: THEME.colors.accent6, isTextBox: true, margin: 0,
  });
}

(async () => {
  const out = "/home/user/Journal/ips_exosome_stroke_pd/iPS-Exosome_中風與帕金森_文獻盤點.pptx";
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("written:", out);
})();
