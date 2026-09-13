// Builds slides/1st_stage/D_operations_management.pptx from lib.js components.
// Run: node scripts/pptx_slidegen/build_D_operations_management.js
//
// Mirrors slides/1st_stage/D_operations_management.html. See docs/13_pptx_slide_template_spec.md.

const path = require("path");
const {
  newPres, addCoverSlide, addDividerSlide, addHeader, addFreqBar, mkYears,
  addTermRows, addRowList, addExamQuestion, drawBarChart,
  INK, INK_SOFT, RED, LINE, GHOST, F_HEAD, F_BODY, F_MONO,
} = require("./lib");

const pres = newPres();
pres.author = "SMEC";
pres.title = "運営管理";

// ---------- Slide 01: Cover ----------
addCoverSlide(pres, {
  eyebrow: "中小企業診断士 第1次試験",
  subjectNo: "SUBJECT 04 / 07",
  title: "運営管理",
  subtitle: "生産管理（工場でどう作るか）と店舗・販売管理（店でどう売るか）の2分野、35論点。暗記事項と計算問題（タクトタイム、EOQ、ライリーの法則等）の両方が出題される。",
  stats: [["35", "収録論点数"], ["2", "分野"], ["32/35", "頻出ランクA"]],
  tag: "docs/textbook/D_operations_management_textbook.md",
  notes: "表紙スライド。運営管理の全体像を示す。",
});

// ---------- Slide 02: 生産管理概論／生産の管理 区切り ----------
addDividerSlide(pres, {
  ghostNo: "01",
  partNo: "PART 01",
  partLabel: "運営管理 ／ D-1〜D-9",
  title: "生産管理概論／生産計画",
  desc: "工場管理の目標（PQCDSME）から、生産方式・ライン編成・需要予測・需給計画・スケジューリングまで。町工場をイメージしながら読み進める。",
  chips: ["D-2 JIT・かんばん", "D-6 ラインバランシング", "D-9 ジョンソンの規則"],
  notes: "生産管理概論／生産計画パートの区切りスライド。",
});

// ---------- Slide 03: D-1 生産管理概論（PQCDSME） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-1 ／ 生産管理概論（PQCDSME）",
    title: "7つの目標はトレードオフの関係",
    overview: "工場管理者が同時に追う7つの目標。一方を追求すると他方が犠牲になりやすい。",
    tag: "運営管理",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "P：生産性", v: "少ない人・設備・時間でどれだけ多く作れるか" },
    { k: "Q：品質", v: "不良品を出さず要求品質を満たすか" },
    { k: "C：コスト", v: "材料費・労務費・経費をどれだけ抑えられるか" },
    { k: "D：納期", v: "決められた期日までに決められた量を届けられるか" },
    { k: "S・M・E", v: "Safety（安全）・Morale（士気）・Environment（環境）" },
  ], { fontSize: 10.5, labelW: 1.7, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "SはSales（販売）ではなくSafety（安全）、EはEfficiency（効率）ではなくEnvironment（環境）。並び順より互いにトレードオフになりうる理解が問われる。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：D-10 生産統制（PQCDSMEを実際に管理する活動）。",
    years: mkYears(new Set(["'17", "'18", "'19", "'20", "'21", "'22", "'23", "'25"])),
  });
}

// ---------- Slide 04: D-2 生産システム ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-2 ／ 生産システム（JIT、かんばん方式、ERP、SCM）",
    title: "かんばんはJITを実現する道具",
    overview: "必要な物を必要な時に必要な量だけ作るJIT。後工程引取りの仕組みがかんばん。",
    tag: "運営管理",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "JIT", v: "必要な物を必要な時に必要な量だけ作る・運ぶ（トヨタ生産方式）" },
    { k: "かんばん方式", v: "後工程が使った分だけ前工程に指示する情報伝達の道具（プル型）" },
    { k: "ERP", v: "会計・人事・生産・販売を1システムで統合管理" },
    { k: "SCM", v: "複数の会社をまたぐ調達〜販売の供給連鎖全体を最適化" },
  ], { fontSize: 11, labelW: 1.8, gap: 0.42 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "かんばん方式はJITを実現する道具であり、JITそのものではない。JIT（プル型）とMRP（D-8、プッシュ型）は対照的な考え方で混同しない。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：D-8 需給計画（MRP、プッシュ型との対比）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'20", "'21", "'22", "'25"])),
  });
}

// ---------- Slide 05: D-3 環境配慮型生産 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-3 ／ 環境配慮型生産（3R、LCA）",
    title: "優先順位はリデュース＞リユース＞リサイクル",
    overview: "3Rの優先順位と、製品の一生涯を評価するLCA。",
    tag: "運営管理",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "①リデュース", desc: "そもそも廃棄物・使用資源の発生量を減らす（最優先）" },
    { name: "②リユース", desc: "使えるものはそのまま繰り返し使う" },
    { name: "③リサイクル", desc: "使い終わったものを資源として再生利用する（優先順位は最後）" },
    { name: "LCA", desc: "原材料採取→製造→輸送→使用→廃棄まで一生涯の環境負荷を定量評価" },
  ], { nameW: 1.9, rowH: 0.9 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：LCAは製造工程だけでなく調達〜廃棄までの全体を対象とする。",
    years: mkYears(new Set(["'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 06: D-4 製品開発・製品設計 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-4 ／ 生産計画：製品開発・製品設計",
    title: "価値＝機能÷コスト",
    overview: "VA/VEは機能を維持・向上しながら価値を高める活動。コスト削減だけが目的ではない。",
    tag: "運営管理",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "VA/VE", v: "価値（V）＝機能（F）÷コスト（C）を高める体系的活動" },
    { k: "コンカレントエンジニアリング", v: "設計・生産技術・調達・販売が同時並行で検討し開発期間を短縮" },
  ], { fontSize: 11, labelW: 2.9, gap: 0.42 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "家電メーカーが企画段階から設計者・生産技術者・調達担当者が一緒に検討し、「金型が高額」「量産時に調達しにくい」等の懸念を早期に反映するのがコンカレントエンジニアリング。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.78, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.86;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「機能を下げてコストを下げれば必ず価値が上がる」は誤り。コンカレントエンジニアリングは「同時並行で情報共有」であり「独立して同時に別設計」ではない。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：D-16 工場計画（SLP、設計後のレイアウト検討にもつながる）。",
    years: mkYears(new Set(["'16", "'17", "'22", "'23", "'25"])),
  });
}

// ---------- Slide 07: D-5 生産方式 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-5 ／ 生産計画：生産方式",
    title: "3つの分類軸で生産方式を整理する",
    overview: "「いつ作るか」「どれだけまとめて作るか」「どう配置するか」の3軸。",
    tag: "運営管理",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "見込／受注生産", desc: "需要予測で先に作る（欠品・過剰在庫リスク）／注文後に作る（リードタイム長い）" },
    { name: "個別／ロット／連続生産", desc: "1個ずつ専用計画／まとまった数量で段取替え／同一製品を継続生産" },
    { name: "ライン／セル生産", desc: "大量生産向き・柔軟性低い／多能工が一連工程を担当し多品種少量に強い" },
  ], { nameW: 2.6, rowH: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：セル生産は多能工化（D-11）が前提。「少人化」と「省人化」は別概念。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'22", "'24", "'25"])),
  });
}

// ---------- Slide 08: D-6 ライン編成（ラインバランシング） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-6 ／ 生産計画：ライン編成（ラインバランシング）",
    title: "ネック工程がサイクルタイムを決める",
    overview: "タクトタイムは目標値、サイクルタイムは実績値（最も遅い工程の時間）。",
    tag: "運営管理",
  });
  const proseX = 0.55, proseW = 7.6;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "タクトタイム", v: "実稼働時間 ÷ 必要生産数（目標値）" },
    { k: "ラインバランシング効率", v: "各工程作業時間合計 ÷ (工程数×サイクルタイム) × 100" },
  ], { fontSize: 10, labelW: 2.5, gap: 0.55 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "分母のサイクルタイムには「平均」ではなく「最大（ネック工程）」の時間を使う。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.55, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  const rx = 8.55, rw = 4.05;
  drawBarChart(s, rx, 1.95, rw, 3.5, {
    bars: [
      { label: "第1工程", value: 200 },
      { label: "第2工程", value: 252, on: true },
      { label: "第3工程", value: 180 },
      { label: "第4工程", value: 220 },
    ],
  });
  s.addText("第2工程（赤）がネック＝サイクルタイム252秒", { x: rx - 0.3, y: 5.9, w: rw + 0.6, h: 0.3, align: "center", fontFace: F_BODY, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "具体例：852秒÷(4×252秒)×100＝約84.5％、バランスロス率約15.5％。",
    years: mkYears(new Set(["'16", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 09: D-6 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-6 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第21問）",
    overview: "計算問題。前のスライドの公式で自分の手で計算してみる。",
    tag: "運営管理",
  });
  const cy = addExamQuestion(s, {
    stem: "ある工場では、1種類の製品をライン生産方式で生産することを計画している。1カ月の製品の必要生産数は2,000個、稼働日は20日である。1日の実稼働時間は7時間である。この生産ラインのタクトタイムが含まれる範囲として、最も適切なものはどれか。",
    stemH: 0.7,
    choices: [
      { badge: "ア", text: "100秒未満" },
      { badge: "イ", text: "100秒以上、200秒未満" },
      { badge: "ウ", text: "200秒以上、300秒未満" },
      { badge: "エ", text: "300秒以上、400秒未満" },
      { badge: "オ", text: "400秒以上" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/D1JI2025.pdf（令和7年度第1次試験）第21問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 10: D-6 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-6 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第21問）",
    overview: "正解はウ。タクトタイム＝25,200秒÷100個＝252秒/個。",
    tag: "運営管理",
  });
  const choices = [
    { badge: "ア", text: "100秒未満" },
    { badge: "イ", text: "100秒以上、200秒未満" },
    { badge: "ウ", text: "200秒以上、300秒未満" },
    { badge: "エ", text: "300秒以上、400秒未満" },
    { badge: "オ", text: "400秒以上" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 2 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：ウ", options: { bold: true, color: RED } },
    { text: "。①1日あたりの必要生産数＝2,000個÷20日＝100個/日。②1日の実稼働時間＝7時間＝25,200秒。③タクトタイム＝25,200秒÷100個＝", options: {} },
    { text: "252秒/個", options: { bold: true } },
    { text: "。252秒は「200秒以上、300秒未満」に含まれる。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.78;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "タクトタイムは「必要な生産量を稼働時間内に満たすには1個あたり何秒で生産すればよいか」を表す目標値で、ラインバランシングの前提となる基準時間として使われる。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.58;
  s.addText("関連知識：docs/07_key_formulas_and_frameworks.mdのラインバランシング効率の分母にこのタクトタイムを使う。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.38;
  s.addText("出典：past_exams/1st_stage/1ji2025/D1JI2025.pdf（令和7年度第1次試験）第21問／正解：past_exams/1st_stage_answers/r07/d_v2_20250902.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 11: D-7 需要予測 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-7 ／ 生産計画：需要予測",
    title: "αが大きいほど直近実績を重視",
    overview: "移動平均法は均等ウェイト、指数平滑法は直近ほど大きなウェイト。",
    tag: "運営管理",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "移動平均法", v: "直近n期間の実績値の単純平均。各期に均等なウェイト" },
    { k: "指数平滑法", v: "次期予測値＝α×当期実績値＋(1－α)×当期予測値" },
  ], { fontSize: 11, labelW: 2.0, gap: 0.42 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "先月実績100個・予測90個。α＝0.3→次期予測93個。α＝0.6→次期予測96個。αが大きいほど直近実績（100個）に近づく。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.55, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.63;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「αが大きいほど予測が安定する」は誤り。αが大きいほど直近実績への反応が敏感になり予測は振れやすくなる。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：D-31 商品在庫管理・需要予測（店舗版でも同じ考え方を使う）。",
    years: mkYears(new Set(["'18", "'19", "'20", "'21"])),
  });
}

// ---------- Slide 12: D-8 需給計画（MRP） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-8 ／ 生産計画：需給計画（MRP、MPS、BOM）",
    title: "総所要量から在庫・入庫予定を引く",
    overview: "MPS（生産計画）とBOM（部品表）から必要部品を展開し、正味所要量を計算する。",
    tag: "運営管理",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "MPS", v: "最終製品を「いつ・いくつ作るか」定めた基準生産計画" },
    { k: "BOM", v: "製品1個に必要な部品・員数を示す部品表" },
    { k: "正味所要量", v: "総所要量－（現在庫＋入庫予定）" },
  ], { fontSize: 11, labelW: 1.7, gap: 0.42 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "自転車1台＝車輪2個。来月50台生産→総所要量100個。在庫15個・入荷予定10個→正味所要量＝100－25＝75個をリードタイム分前倒しで発注。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.55, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.63;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "員数（自転車1台＝車輪2個）の掛け忘れに注意。MRPは計画に基づき部品を押し出すプッシュ型で、かんばん（プル型、D-2）とは対照的。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：D-2 生産システム（JIT・かんばん＝プル型との対比）。",
    years: mkYears(new Set(["'16", "'18", "'19", "'20", "'21", "'22", "'23", "'24"])),
  });
}

// ---------- Slide 13: D-9 スケジューリング ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-9 ／ 生産計画：スケジューリング",
    title: "短い工程に着目して投入順序を決める",
    overview: "PERTのクリティカルパスと、2機械フローショップのジョンソンの規則。",
    tag: "運営管理",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "クリティカルパス", v: "PERT図で最も時間のかかる経路。余裕時間ゼロで遅れが全体に直結" },
    { k: "ジョンソンの規則", v: "第1工程が短いジョブは前方へ昇順、第2工程が短いジョブは後方へ降順" },
  ], { fontSize: 11, labelW: 2.2, gap: 0.42 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "4ジョブa,b,c,dの投入順序＝a→c→d→b（前方は第1工程時間昇順、後方は第2工程時間降順）。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.58, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.66;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「第1工程が長いジョブを先に」ではなく「短い工程に着目」。ジョンソンの規則は2台機械・同一工程順序のフローショップ限定。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：D-10 生産統制（進捗管理でスケジュール通りかを日々チェック）。",
    years: mkYears(new Set(["'16", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 14: D-9 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-9 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第6問）",
    overview: "計算問題。前のスライドの公式で自分の手で計算してみる。",
    tag: "運営管理",
  });
  const cy = addExamQuestion(s, {
    stem: "2台の機械を用いて、第1工程、第2工程の順で作業するフローショップにおいて、各ジョブの作業時間が下表に与えられている（ジョブa：第1工程3・第2工程4／ジョブb：第1工程4・第2工程1／ジョブc：第1工程5・第2工程7／ジョブd：第1工程6・第2工程2）。メイクスパンが最小になるように各ジョブがフローショップに投入された場合、メイクスパンにおける2台の機械の非稼働時間の合計として、最も適切なものを下記の解答群から選べ。",
    stemH: 1.1,
    choices: [
      { badge: "ア", text: "3" },
      { badge: "イ", text: "4" },
      { badge: "ウ", text: "5" },
      { badge: "エ", text: "6" },
      { badge: "オ", text: "7" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/D1JI2025.pdf（令和7年度第1次試験）第6問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 15: D-9 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-9 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第6問）",
    overview: "正解はエ（6）。投入順序a→c→d→bでメイクスパン19、非稼働時間合計6。",
    tag: "運営管理",
  });
  const choices = [
    { badge: "ア", text: "3" },
    { badge: "イ", text: "4" },
    { badge: "ウ", text: "5" },
    { badge: "エ", text: "6" },
    { badge: "オ", text: "7" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 3 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：エ（6）", options: { bold: true, color: RED } },
    { text: "。ジョンソンの規則より投入順序＝a→c→d→b。第1工程：0-3-8-14-18（合計18稼働、非稼働1）。第2工程：a(3-7)→c(8-15)→d(15-17)→b(18-19)。メイクスパン＝19。第2工程の非稼働＝開始前3＋a後の待ち1＋d後の待ち1＝5。2台合計＝1＋5＝", options: {} },
    { text: "6", options: { bold: true } },
    { text: "。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.0, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 1.08;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "a：min(3,4)=3→前方、b：min(4,1)=1→後方、c：min(5,7)=5→前方、d：min(6,2)=2→後方。前方は第1工程時間昇順（a→c）、後方は第2工程時間降順（d→b）。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.58;
  s.addText("関連知識：ジョンソンの規則は2機械n個ジョブのフローショップでメイクスパンを最小化する代表的手法。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.38;
  s.addText("出典：past_exams/1st_stage/1ji2025/D1JI2025.pdf（令和7年度第1次試験）第6問／正解：past_exams/1st_stage_answers/r07/d_v2_20250902.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 16: 作業の管理／設備の管理／物の管理 区切り ----------
addDividerSlide(pres, {
  ghostNo: "02",
  partNo: "PART 02",
  partLabel: "運営管理 ／ D-10〜D-19",
  title: "生産統制と作業・設備管理",
  desc: "計画通りに実行できているかをチェックする生産統制から、標準時間・作業改善、5S・TPMによる設備管理、工場レイアウト、資材・外注管理まで。",
  chips: ["D-13 ECRSの原則", "D-14 5S・TPM", "D-16 SLP"],
  notes: "生産統制／作業・設備・物の管理パートの区切りスライド。",
});

// ---------- Slide 17: D-10 生産統制 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-10 ／ 生産統制",
    title: "「モノ」「能力」「スケジュール」を管理する",
    overview: "計画通りに実行できているかを日々チェック・調整する3つの管理。",
    tag: "運営管理",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "現品管理", tag: "モノ", desc: "仕掛品・材料・製品が今どこに・どれだけあるかを把握し紛失・取り違えを防ぐ" },
    { name: "余力管理", tag: "能力", desc: "職場・設備の保有能力と実際の負荷を比較し過負荷・手待ちを調整する" },
    { name: "進捗管理", tag: "スケジュール", desc: "計画に対する実際の進み具合を把握し、遅れていれば挽回策を打つ" },
  ], { nameW: 2.0, tagW: 1.6, rowH: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：3つの管理対象を入れ替えた誤答が定番。名称と対象を正確に対応させる。",
    years: mkYears(new Set(["'16", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 18: D-11 作業管理 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-11 ／ 作業管理",
    title: "標準時間＝正味時間＋余裕時間",
    overview: "熟練者が正常なペースで作業するのに必要な時間。外掛け法と内掛け法で余裕率の定義が違う。",
    tag: "運営管理",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "正味時間", v: "観測時間の代表値×レイティング係数" },
    { k: "外掛け法", v: "標準時間＝正味時間×（1＋余裕率）" },
    { k: "内掛け法", v: "標準時間＝正味時間÷（1－余裕率）" },
    { k: "標準作業", v: "サイクルタイム・作業順序・標準手持ちの3要素（トヨタ生産方式）" },
  ], { fontSize: 10.5, labelW: 1.9, gap: 0.38 });
  cy += 0.08;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "観測時間48秒、レイティング110％→正味時間52.8秒。外掛け余裕率20％→標準時間＝52.8×1.2＝63.36秒。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.58, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.66;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "外掛け法・内掛け法は割る対象が違うため余裕率の数値自体が変わる。標準作業の3要素と標準時間の構成要素を混同しない。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：D-5 セル生産方式（多能工化が前提条件）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24"])),
  });
}

// ---------- Slide 19: D-12 作業の分析手法 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-12 ／ 作業の分析手法",
    title: "工程図記号で「見える化」する",
    overview: "貯蔵（計画的）と滞留（計画外）の違いが頻出のひっかけ。",
    tag: "運営管理",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "○ 加工", desc: "形状・性質を変える作業" },
    { name: "⇒ 運搬", desc: "位置を移動させる" },
    { name: "□／◇ 検査", desc: "□数量検査（個数・重量）／◇品質検査（規格適合）" },
    { name: "▽／D 貯蔵・滞留", desc: "▽計画的にためる／D計画外のやむを得ない足止め" },
  ], { nameW: 2.5, rowH: 0.9 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "動作分析：ギルブレス夫妻のサーブリッグ分析（つかむ・運ぶ・探す等に分解）。D-13につながる。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24"])),
  });
}

// ---------- Slide 20: D-13 作業の改善 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-13 ／ 作業の改善",
    title: "まず「なくせないか」から考える",
    overview: "ECRSはE→C→R→Sの順で検討する。順番自体が頻出の出題ポイント。",
    tag: "運営管理",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "E：排除", v: "その作業自体をなくせないか（最初に検討）" },
    { k: "C：結合", v: "複数の作業を1つにまとめられないか" },
    { k: "R：交換", v: "作業の順序・場所を入れ替えられないか" },
    { k: "S：簡素化", v: "より簡単な方法にできないか" },
  ], { fontSize: 11, labelW: 1.6, gap: 0.38 });
  cy += 0.08;
  s.addText([
    { text: "動作経済の原則：", options: { bold: true, color: INK } },
    { text: "身体の使用／作業場の配置／工具・設備の設計、の3分類でムダのない動き方・環境を作る。", options: { color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 10.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.48;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「まず簡素化を検討し、それでも難しければ排除を検討する」のように順序を入れ替えた記述は誤り。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：D-12 動作分析（サーブリッグ分析の結果がここでの改善につながる）。",
    years: mkYears(new Set(["'17", "'19", "'20", "'22", "'23", "'25"])),
  });
}

// ---------- Slide 21: D-14 設備管理 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-14 ／ 設備管理（5S、TPM）",
    title: "自主保全が現場のオペレーターの手で行われる",
    overview: "5Sは職場環境の基本活動。TPMは全員参加の設備保全活動。",
    tag: "運営管理",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "整理／整頓", v: "不要品を捨てる／必要な物をすぐ取り出せる状態に置く" },
    { k: "清掃・清潔・躾", v: "掃除しながら点検／維持する／ルールを守る習慣" },
    { k: "TPM・自主保全", v: "現場オペレーター自身が日常点検・簡単な保全を行う" },
    { k: "設備総合効率（OEE）", v: "時間稼働率×性能稼働率×良品率" },
  ], { fontSize: 10.5, labelW: 2.3, gap: 0.38 });
  cy += 0.08;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "時間稼働率90％×性能稼働率約92.6％×良品率95％＝設備総合効率約79.2％。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.58, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.66;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「整理」と「整頓」の意味を入れ替えた選択肢が頻出。自主保全は保全担当者だけでなく現場オペレーターも行う。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：D-17 設備投資の経済性分析（設備の更新判断につながる）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 22: D-15 設備の評価と更新 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-15 ／ 設備の評価と更新",
    title: "耐用年数は税法上の目安であって物理的限界ではない",
    overview: "稼働率・減価償却・耐用年数を踏まえて設備更新を判断する。",
    tag: "運営管理",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "稼働率", v: "実際生産量 ÷ 標準（基準）生産量 × 100" },
    { k: "減価償却", v: "取得原価を耐用年数にわたり費用配分（定額法／定率法）" },
    { k: "耐用年数", v: "税法上定められた使用可能期間の目安（物理的限界とは別）" },
  ], { fontSize: 11, labelW: 1.6, gap: 0.42 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「耐用年数＝実際に使用できる物理的限界」は誤り。老朽化による稼働率・良品率低下、修理費増加も総合的に考慮して更新判断する。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  s.addShape("line", { x: 0.55, y: 6.55, w: 12.25, h: 0, line: { color: INK, width: 0.75 } });
  s.addShape("ellipse", { x: 0.55, y: 6.63, w: 0.36, h: 0.36, fill: { color: INK_SOFT }, line: { type: "none" } });
  s.addText("－", { x: 0.55, y: 6.63, w: 0.36, h: 0.36, align: "center", valign: "middle", fontFace: F_MONO, fontSize: 13, bold: true, color: "FFFFFF", isTextBox: true, margin: 0 });
  s.addText("集計データなし　／　関連：D-17 設備投資の経済性分析（更新すべきかの投資判断）。", {
    x: 1.0, y: 6.63, w: 11.5, h: 0.36, valign: "middle", fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 23: D-16 工場計画 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-16 ／ 工場計画（SLP、DI分析、プラントレイアウト）",
    title: "勘や経験ではなく体系的にレイアウトを決める",
    overview: "SLPは物の流れ分析→近接性評価→レイアウト案作成の手順で進める。",
    tag: "運営管理",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "SLP", desc: "P-Q分析で物の流れを把握→アクティビティ相互関係図（A・E・I・O・U・X）→レイアウト案" },
    { name: "DI分析", desc: "方向（Direction）と関係の強さ（Intensity）から工程間の結びつきを検討" },
    { name: "機能別／製品別／固定型", desc: "工程ごとにまとめる／加工順に並べる／大型製品の周りに配置" },
  ], { nameW: 2.6, rowH: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：SLPは体系的・定量的な手順。DI分析は「距離の近さだけ」ではなく方向＋強さの2視点。",
    years: mkYears(new Set(["'17", "'18", "'19", "'20", "'21", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 24: D-17 設備投資の経済性分析 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-17 ／ 設備投資の経済性分析",
    title: "回収期間法は時間的価値を考慮しない",
    overview: "設備更新の判断は「変化する部分（差額）」に着目する差額原価収益分析の考え方。",
    tag: "運営管理",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "回収期間法", v: "投資額 ÷ 年間キャッシュフロー（簡便だが時間的価値を考慮しない）" },
    { k: "NPV法", v: "将来CFを割り引いた合計－投資額。プラスなら投資価値あり" },
  ], { fontSize: 11, labelW: 1.7, gap: 0.42 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "投資額600万円、年間CF200万円→回収期間＝600÷200＝3年。4年目以降の利益や割引率は考慮されない。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.55, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.63;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「回収期間が短い投資ほど必ず優れている」とは限らない。より厳密にはNPV法（B-26参照）を使う。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：財務・会計B-24〜B-26（NPV法・IRR法の詳細）。",
    years: mkYears(new Set(["'16", "'19", "'23", "'24"])),
  });
}

// ---------- Slide 25: D-18 資材管理 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-18 ／ 資材管理（資材計画、常備品）",
    title: "常備品は在庫管理、非常備品は都度手配",
    overview: "生産に必要な資材を、必要な時に必要な量だけ調達・管理する活動。",
    tag: "運営管理",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "常備品", v: "多くの製品に共通し継続的に使う資材（ネジ等）。発注点方式で管理" },
    { k: "非常備品", v: "特定注文にのみ使う資材（特注部品）。都度手配" },
  ], { fontSize: 11, labelW: 1.7, gap: 0.42 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "常備品は在庫を持って継続管理、非常備品は注文ごとに個別手配、という対応を正確に覚える。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "C", rankLabel: "直近10年で1回出題",
    related: "関連：D-20 在庫管理（常備品の発注方式・EOQへつながる）。",
    years: mkYears(new Set(["'24"])),
  });
}

// ---------- Slide 26: D-19 調達・外注管理 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-19 ／ 調達・外注管理",
    title: "外注は単なるコスト削減策ではない",
    overview: "自社に無い技術の活用、生産能力の変動吸収など多様な目的で行われる。",
    tag: "運営管理",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "外注のメリット", v: "技術・設備の活用、生産能力の変動吸収、固定費の変動費化" },
    { k: "外注のデメリット", v: "技術・ノウハウ流出リスク、品質管理の難しさ、依存度上昇リスク" },
  ], { fontSize: 11, labelW: 1.9, gap: 0.42 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "繁忙期にだけ生産量が急増する部品加工を近隣の協力会社に外注し、自社で新規設備投資・人員採用をせずに需要変動へ対応する。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.55, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.63;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "外注比率を高めすぎると自社の技術力・ノウハウの空洞化リスクにつながる点も出題対象。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：D-18 資材管理（非常備品の都度手配・外部調達との接点）。",
    years: mkYears(new Set(["'16", "'17", "'24"])),
  });
}

// ---------- Slide 27: 在庫・品質管理／店舗・商業集積 区切り ----------
addDividerSlide(pres, {
  ghostNo: "03",
  partNo: "PART 03",
  partLabel: "運営管理 ／ D-20〜D-28",
  title: "在庫・品質管理／店舗集積",
  desc: "在庫管理（ABC分析、EOQ）・品質管理（QC7つ道具）・廃棄物管理で第1部を締めくくり、視点を店舗に移して出店法規・商圏分析・業態・商品計画・仕入・陳列まで。",
  chips: ["D-20 EOQ・発注方式", "D-21 QC7つ道具", "D-24 ライリーの法則"],
  notes: "在庫・品質管理／店舗・商業集積パートの区切りスライド。",
});

// ---------- Slide 28: D-20 在庫管理 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-20 ／ 在庫管理（ABC分析、発注方式、EOQ）",
    title: "定量発注は発注量固定、定期発注は発注時期固定",
    overview: "ABC分析で重要度にメリハリをつけ、EOQで最適な発注量を求める。",
    tag: "運営管理",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "ABC分析", desc: "金額順に並べ累積構成比でA・B・Cにランク分け（パレートの法則）" },
    { name: "定量発注方式", tag: "発注量固定", desc: "発注点で発注、発注間隔は需要により変動。Cランク向き" },
    { name: "定期発注方式", tag: "発注時期固定", desc: "一定間隔で需要予測に基づき発注量を変動。Aランク向き" },
    { name: "EOQ", tag: "経済的発注量", desc: "√(2×年間需要量×発注コスト÷単位あたり保管コスト)" },
  ], { nameW: 2.3, tagW: 2.0, rowH: 0.9 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：安全係数は高くするほど欠品リスクが減る。ダブルビン方式の発注量は補充点相当量。",
    years: mkYears(new Set(["'17", "'19", "'20", "'21", "'22", "'23", "'24"])),
  });
}

// ---------- Slide 29: D-20 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-20 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第32問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "運営管理",
  });
  const cy = addExamQuestion(s, {
    stem: "小売店舗における在庫管理に関する記述として、最も適切なものはどれか。",
    stemH: 0.5,
    choices: [
      { badge: "ア", text: "安全在庫を設定する場合、欠品のリスクを小さくするためには、安全係数を低く設定する必要がある。", h: 0.6 },
      { badge: "イ", text: "ダブルビン方式を採用している場合、発注量は補充点の2倍に相当する量である。", h: 0.6 },
      { badge: "ウ", text: "定期発注方式を採用している場合、発注から納品までの調達期間を長くすると、発注のための需要予測量は多くなる。", h: 0.6 },
      { badge: "エ", text: "定期発注方式を採用している場合、発注間隔を長くすると、発注1回当たりの発注量は少なくなる。", h: 0.6 },
      { badge: "オ", text: "定量発注方式を採用している場合、発注点に基づいて発注すれば、販売量の増減にかかわらず、発注間隔は一定になる。", h: 0.6 },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/D1JI2025.pdf（令和7年度第1次試験）第32問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 30: D-20 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-20 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第32問）",
    overview: "正解はウ。調達期間が長いほど見積もるべき需要期間が延びる。",
    tag: "運営管理",
  });
  const choices = [
    { badge: "ア", text: "安全在庫を設定する場合、欠品のリスクを小さくするためには、安全係数を低く設定する必要がある。", h: 0.6 },
    { badge: "イ", text: "ダブルビン方式を採用している場合、発注量は補充点の2倍に相当する量である。", h: 0.6 },
    { badge: "ウ", text: "定期発注方式を採用している場合、発注から納品までの調達期間を長くすると、発注のための需要予測量は多くなる。", h: 0.6 },
    { badge: "エ", text: "定期発注方式を採用している場合、発注間隔を長くすると、発注1回当たりの発注量は少なくなる。", h: 0.6 },
    { badge: "オ", text: "定量発注方式を採用している場合、発注点に基づいて発注すれば、販売量の増減にかかわらず、発注間隔は一定になる。", h: 0.6 },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 2 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：ウ", options: { bold: true, color: RED } },
    { text: "。定期発注方式では発注のたびに「調達期間＋発注サイクル」分の需要を予測して発注量を決める。調達期間を長くすると需要予測量は", options: {} },
    { text: "多くなる", options: { bold: true } },
    { text: "。ア：欠品リスクを下げるには安全係数を", options: {} },
    { text: "高く", options: { bold: true } },
    { text: "設定する必要がある。イ：ダブルビン方式の発注量はビン1つ分＝", options: {} },
    { text: "補充点相当量", options: { bold: true } },
    { text: "。エ：発注間隔を長くすると発注量はむしろ多くなる。オ：定量発注方式は発注量が一定で発注間隔は", options: {} },
    { text: "変動", options: { bold: true } },
    { text: "する。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.0, fontFace: F_BODY, fontSize: 10, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.18 });
  cy += 1.08;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "安全在庫＝安全係数×需要のばらつき（標準偏差）×√調達期間、で計算される。欠品率を下げたい場合は安全係数を高く設定する。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.58;
  s.addText("関連知識：定量発注方式は「発注量固定・発注時期変動」、定期発注方式は「発注時期固定・発注量変動」で整理する。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.38;
  s.addText("出典：past_exams/1st_stage/1ji2025/D1JI2025.pdf（令和7年度第1次試験）第32問／正解：past_exams/1st_stage_answers/r07/d_v2_20250902.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 31: D-21 品質管理 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-21 ／ 品質管理",
    title: "数値データはQC7つ道具、言語データは新QC7つ道具",
    overview: "扱うデータの種類で2つの道具群を使い分ける。",
    tag: "運営管理",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "QC7つ道具", tag: "数値データ", desc: "パレート図・特性要因図・ヒストグラム・チェックシート・散布図・層別・管理図" },
    { name: "新QC7つ道具", tag: "言語データ", desc: "連関図法・親和図法・系統図法・マトリックス図法・PDPC法・アロー図法" },
  ], { nameW: 2.3, tagW: 1.9, rowH: 1.3 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：名称と特徴の組み合わせ入替が頻出。TQM＝全社的品質管理、ISO9000＝品質国際規格。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 32: D-21 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-21 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第5問）",
    overview: "組み合わせ問題。前のスライドの内容で答えられるか考えてみる。",
    tag: "運営管理",
  });
  const cy = addExamQuestion(s, {
    stem: "以下に示すQC七つ道具、新QC七つ道具の名称と、その特徴に関する記述の組み合わせとして、最も適切なものを下記の解答群から選べ。\n【名称】ａ：マトリックス図法／ｂ：ヒストグラム／ｃ：連関図法／ｄ：チェックシート\n【特徴】①不良原因の中から対になる要素を見つけ解決手段の優先順位付けの重みを求める／②結果と要因の関係を分析し因果関係を明らかにする／③品質管理データを簡単に記録し現状を数値化する／④計量値の「ずれ」や「ばらつき」を視覚的に表し工程の実力をつかむ",
    stemH: 1.3,
    choices: [
      { badge: "ア", text: "a：① b：② c：③ d：④" },
      { badge: "イ", text: "a：① b：④ c：② d：③" },
      { badge: "ウ", text: "a：② b：① c：④ d：③" },
      { badge: "エ", text: "a：③ b：① c：④ d：②" },
      { badge: "オ", text: "a：③ b：④ c：② d：①" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/D1JI2025.pdf（令和7年度第1次試験）第5問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 33: D-21 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-21 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第5問）",
    overview: "正解はイ。a：①、b：④、c：②、d：③が正しい対応。",
    tag: "運営管理",
  });
  const choices = [
    { badge: "ア", text: "a：① b：② c：③ d：④" },
    { badge: "イ", text: "a：① b：④ c：② d：③" },
    { badge: "ウ", text: "a：② b：① c：④ d：③" },
    { badge: "エ", text: "a：③ b：① c：④ d：②" },
    { badge: "オ", text: "a：③ b：④ c：② d：①" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 1 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：イ", options: { bold: true, color: RED } },
    { text: "。ａ：マトリックス図法＝①（行と列に要素を配置し関連の強さから優先順位の重みを求める）。ｂ：ヒストグラム＝④（計量値データのばらつきを柱状図で表す）。ｃ：連関図法＝②（結果と要因を矢印で結び因果関係を整理）。ｄ：チェックシート＝③（点検項目にデータを記録し現状を数値化）。他の選択肢はいずれかの組み合わせで手法名と特徴がずれている。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.0, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 1.08;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "「数値データ→QC7つ道具」「言語データ・構造の整理→新QC7つ道具」という大枠で覚えたうえで、各手法の定義を1つずつ正確に区別する。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.58;
  s.addText("関連知識：マトリックス・データ解析法だけが新QC7つ道具の中で唯一数値データを扱う点は要注意。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.38;
  s.addText("出典：past_exams/1st_stage/1ji2025/D1JI2025.pdf（令和7年度第1次試験）第5問／正解：past_exams/1st_stage_answers/r07/d_v2_20250902.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 34: D-22 廃棄物等の管理 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-22 ／ 廃棄物等の管理（環境保全法規、ISO14000）",
    title: "マニフェストで最終処分までを追跡する",
    overview: "委託しても排出事業者の責任は完全には終わらない。",
    tag: "運営管理",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "廃棄物処理法", v: "適正処理を義務付け。産廃委託時はマニフェスト（管理票）交付・保管が必須" },
    { k: "資源有効利用促進法", v: "3Rの考え方に基づき資源の有効利用・リサイクルを促す" },
    { k: "ISO14000シリーズ", v: "環境マネジメントシステムの国際規格（ISO9000＝品質とは分野が異なる）" },
  ], { fontSize: 11, labelW: 2.4, gap: 0.42 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「処理業者に委託すれば排出事業者の責任は完全に終わる」は誤り。最終処分までの適正処理を確認する責任は排出事業者に残る。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：D-3 環境配慮型生産（3R・LCAと同じ環境配慮の文脈）。",
    years: mkYears(new Set(["'19", "'20", "'21", "'23", "'24"])),
  });
}

// ---------- Slide 35: D-23 店舗施設に関する法律知識 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-23 ／ 店舗施設に関する法律知識",
    title: "大店立地法の目的は「生活環境の保持」",
    overview: "まちづくり三法。大店立地法は中小小売業の保護が目的ではない。",
    tag: "運営管理",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "大店立地法", desc: "1,000㎡超の店舗新設時、周辺の生活環境（交通・騒音・廃棄物）への配慮を義務付け" },
    { name: "都市計画法", desc: "用途地域ごとに建てられる建物の用途を規制" },
    { name: "中心市街地活性化法", desc: "郊外大型店の出店等で衰退した中心市街地を活性化する支援策" },
  ], { nameW: 2.6, rowH: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：大店立地法の目的を「中小小売業者の保護」と誤解しない（旧大店法との違い）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 36: D-24 店舗立地と出店 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-24 ／ 店舗立地と出店（立地条件、商圏分析）",
    title: "人口に比例、距離の2乗に反比例",
    overview: "ライリーの法則とハフモデル。どちらも重力モデル的な考え方。",
    tag: "運営管理",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "ライリーの法則", v: "吸引力比(A:B)＝(人口A÷人口B)×(距離B÷距離A)²" },
    { k: "ハフモデル", v: "来店確率＝(売場面積÷距離ᵈ)÷Σ(各店の売場面積÷距離ᵈ)" },
  ], { fontSize: 11, labelW: 2.1, gap: 0.42 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "人口比A:B=3:8で吸引力比3:2→距離比A:B＝1:2（人口で劣るA市の方がX町に近い位置にある）。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.58, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.66;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "距離の比を「2乗する」のを忘れる計算ミスに注意。公式の丸暗記だけでなく結果の意味も確認する。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：D-25 商業集積（出店先の業態選択にもつながる）。",
    years: mkYears(new Set(["'18", "'20", "'21", "'22", "'25"])),
  });
}

// ---------- Slide 37: D-24 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-24 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第24問）",
    overview: "計算問題。前のスライドの公式で自分の手で計算してみる。",
    tag: "運営管理",
  });
  const cy = addExamQuestion(s, {
    stem: "A市とB市の人口比がA：B＝3：8である場合、それぞれの市がその中間にあるX町からどの程度の購買力を吸引できるか、ライリー・モデルを用いて計算した。計算した結果、A市とB市の吸引力の比がA：B＝3：2のとき、X町からA市までの距離XAとX町からB市までの距離XBの比として、最も適切なものはどれか。",
    stemH: 0.75,
    choices: [
      { badge: "ア", text: "XA：XB＝1：4" },
      { badge: "イ", text: "XA：XB＝1：2" },
      { badge: "ウ", text: "XA：XB＝1：1" },
      { badge: "エ", text: "XA：XB＝2：1" },
      { badge: "オ", text: "XA：XB＝4：1" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/D1JI2025.pdf（令和7年度第1次試験）第24問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 38: D-24 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-24 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第24問）",
    overview: "正解はイ（1：2）。人口で劣るA市の方がX町に近い。",
    tag: "運営管理",
  });
  const choices = [
    { badge: "ア", text: "XA：XB＝1：4" },
    { badge: "イ", text: "XA：XB＝1：2" },
    { badge: "ウ", text: "XA：XB＝1：1" },
    { badge: "エ", text: "XA：XB＝2：1" },
    { badge: "オ", text: "XA：XB＝4：1" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 1 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：イ（1：2）", options: { bold: true, color: RED } },
    { text: "。吸引力の比＝(人口A÷人口B)×(距離XB÷距離XA)²。3／2＝(3／8)×(XB／XA)²より、(XB／XA)²＝(3/2)÷(3/8)＝4→XB／XA＝2。よってXA：XB＝", options: {} },
    { text: "1：2", options: { bold: true } },
    { text: "。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.85, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.93;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "人口で劣るA市がB市より吸引力で善戦する（3：2）には、A市の方がX町に近い位置（XA＜XB）である必要があり、計算結果と整合する。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.58;
  s.addText("関連知識：2都市の吸引力が等しくなる商圏分岐点を求める「コンバースの法則」もあわせて押さえるとよい。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.38;
  s.addText("出典：past_exams/1st_stage/1ji2025/D1JI2025.pdf（令和7年度第1次試験）第24問／正解：past_exams/1st_stage_answers/r07/d_v2_20250902.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 39: D-25 商業集積と業種・業態 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-25 ／ 商業集積と業種・業態",
    title: "業種は「何を売るか」、業態は「どう売るか」",
    overview: "近年の小売業の変化は多くが「業態」の多様化として説明される。",
    tag: "運営管理",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "ショッピングセンター（SC）", v: "デベロッパーが計画的に開発・運営する複数テナント施設" },
    { k: "商店街", v: "特定地域に自然発生的に形成された独立店舗の集まり" },
    { k: "業種／業態", v: "何を売るか（酒屋・米屋）／どう売るか（コンビニ・専門店）" },
  ], { fontSize: 11, labelW: 2.8, gap: 0.42 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「業種」と「業態」を逆に覚えない。コンビニ・ドラッグストア・ネットスーパーの台頭は「業態」の多様化。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：D-26 商品販売計画（業態ごとに品揃えの幅・奥行きが変わる）。",
    years: mkYears(new Set(["'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 40: D-26 商品販売計画 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-26 ／ 商品販売計画",
    title: "「幅」と「奥行き」は別の軸",
    overview: "マーチャンダイジング＝何を・どれだけ・いつ・いくらで仕入れて売るかの計画。",
    tag: "運営管理",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "幅", v: "取り扱う商品カテゴリーの多さ（例：コンビニは幅広い）" },
    { k: "奥行き（深さ）", v: "1カテゴリー内のバリエーションの豊富さ（例：専門店は奥行きが深い）" },
  ], { fontSize: 11, labelW: 2.0, gap: 0.42 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "コンビニは食品・日用品・雑誌など幅広く（幅が広い）、各カテゴリー内の品目数は絞る。スポーツ用品専門店は幅は狭いが、1競技用品のブランド・モデル数（奥行き）は豊富。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.78, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.86;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「品揃えが多いか少ないか」の1軸で捉えない。幅と奥行きは別の軸で店舗の戦略的性格を表す。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：D-28 売場構成・陳列（VMDによる商品の見せ方）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 41: D-27 商品調達・取引条件 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-27 ／ 商品調達・取引条件",
    title: "消化仕入は「売れた時点」で仕入が成立",
    overview: "在庫リスクを仕入先側が負う特殊な仕入形態。",
    tag: "運営管理",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "随時仕入", desc: "必要な都度、その時の状況に応じて仕入れる" },
    { name: "一括大量仕入", desc: "まとめて大量仕入で単価を下げる（在庫リスク増）" },
    { name: "消化（売上）仕入", desc: "店頭商品が実際に売れた時点で仕入が成立（百貨店等）" },
  ], { nameW: 2.6, rowH: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "B", rankLabel: "直近10年で2回出題",
    related: "ひっかけ：「店頭に並べた時点で仕入成立」は誤り。実際に販売された時点で成立する。",
    years: mkYears(new Set(["'16", "'19"])),
  });
}

// ---------- Slide 42: D-28 売場構成・陳列 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "D-28 ／ 売場構成・陳列（VMD、棚割り）",
    title: "VP→PP→IPで大きさが変わる",
    overview: "店舗全体・コーナー・個々の商品という3階層で視覚的に演出する。",
    tag: "運営管理",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "VP", tag: "Visual Presentation", desc: "ストア全体・ウィンドウ等での提案演出" },
    { name: "PP", tag: "Point of Presentation", desc: "コーナーの目立つ場所での訴求" },
    { name: "IP", tag: "Item Presentation", desc: "個々の商品の陳列（棚割り）" },
  ], { nameW: 1.2, tagW: 2.9, rowH: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "具体例：売れ筋商品を目線の高さ（ゴールデンライン）に配置し、関連商品を近くに並べる（クロスMD）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

pres.writeFile({ fileName: path.join(__dirname, "../../slides/1st_stage/D_operations_management.pptx") })
  .then(() => console.log("wrote", "slides/1st_stage/D_operations_management.pptx"));
