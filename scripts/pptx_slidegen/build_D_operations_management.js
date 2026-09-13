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

pres.writeFile({ fileName: path.join(__dirname, "../../slides/1st_stage/D_operations_management.pptx") })
  .then(() => console.log("wrote", "slides/1st_stage/D_operations_management.pptx"));
