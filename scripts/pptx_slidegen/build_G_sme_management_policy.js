// Builds slides/1st_stage/G_sme_management_policy.pptx from lib.js components.
// Run: node scripts/pptx_slidegen/build_G_sme_management_policy.js
//
// Mirrors slides/1st_stage/G_sme_management_policy.html. See docs/13_pptx_slide_template_spec.md.

const path = require("path");
const {
  newPres, addCoverSlide, addDividerSlide, addHeader, addFreqBar, mkYears,
  addTermRows, addRowList, addExamQuestion,
  INK, INK_SOFT, RED, LINE, GHOST, F_HEAD, F_BODY, F_MONO,
} = require("./lib");

const pres = newPres();
pres.author = "SMEC";
pres.title = "中小企業経営・政策";

// ---------- Slide 01: Cover ----------
addCoverSlide(pres, {
  eyebrow: "中小企業診断士 第1次試験",
  subjectNo: "SUBJECT 07 / 07",
  title: "中小企業経営・政策",
  subtitle: "中小企業経営（統計・経営実態）と中小企業政策（支援制度）の2分野、28論点。当年度の中小企業白書がそのまま出題範囲になる、唯一の時事色の強い科目。",
  stats: [["28", "収録論点数"], ["2", "分野"], ["24/28", "頻出ランクA"]],
  tag: "docs/textbook/G_sme_management_policy_textbook.md",
  notes: "表紙スライド。中小企業経営・政策の全体像を示す。",
});

// ---------- Slide 02: G-1〜G-17 区切り ----------
addDividerSlide(pres, {
  ghostNo: "01",
  partNo: "PART 01",
  partLabel: "中小企業経営・政策 ／ G-1〜G-17",
  title: "中小企業経営",
  desc: "「そもそも中小企業とは何か」という中小企業基本法の定義から、統計データ・業種特性・経営課題、事業承継やDXなど最近の動向まで。",
  chips: ["G-1 中小企業基本法の定義", "G-5 ローカルベンチマーク", "G-9 交際費特例"],
  notes: "中小企業経営パートの区切りスライド。",
});

// ---------- Slide 03: G-1 中小企業基本法の定義 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-1 ／ 中小企業基本法の定義",
    title: "資本金基準と従業員基準はどちらか一方でよい",
    overview: "飲食店業は「小売業」に分類される、業種区分の落とし穴が頻出。",
    tag: "中小企業経営・政策",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "製造業・建設業等", tag: "資本金3億円以下", desc: "または従業員300人以下" },
    { name: "卸売業", tag: "資本金1億円以下", desc: "または従業員100人以下" },
    { name: "サービス業", tag: "資本金5,000万円以下", desc: "または従業員100人以下" },
    { name: "小売業（飲食店業含む）", tag: "資本金5,000万円以下", desc: "または従業員50人以下" },
  ], { nameW: 2.7, tagW: 2.1, rowH: 0.9 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：基準はAND（両方満たす）ではなくOR（どちらか一方）。旅館業等は従業員基準の特例あり（例：旅館業は200人以下）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 04: G-1 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-1 ／ 過去問で確認する",
    title: "こう出題される（令和6年度 第17問設問1）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "中小企業経営・政策",
  });
  const cy = addExamQuestion(s, {
    stem: "中小企業基本法第2条に基づく「中小企業者」に含まれる企業に関する正誤の組み合わせとして、最も適切なものを下記の解答群から選べ。ａ：常時使用する従業員数が60人の日本料理店（資本金３千万円）。ｂ：常時使用する従業員数が80人の旅館（資本金６千万円）。ｃ：常時使用する従業員数が120人の生活関連サービス業（資本金８千万円）。",
    stemH: 1.1,
    choices: [
      { badge: "ア", text: "ａ：正　ｂ：正　ｃ：誤", h: 0.42 },
      { badge: "イ", text: "ａ：正　ｂ：誤　ｃ：誤", h: 0.42 },
      { badge: "ウ", text: "ａ：誤　ｂ：正　ｃ：正", h: 0.42 },
      { badge: "エ", text: "ａ：誤　ｂ：誤　ｃ：正", h: 0.42 },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2024/G1JI2024.pdf（令和6年度第1次試験）第17問設問1", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 05: G-1 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-1 ／ 過去問で確認する",
    title: "解答＆解説（令和6年度 第17問設問1）",
    overview: "正解はア。ａ・ｂはOR条件・特例業種で中小企業者に該当。",
    tag: "中小企業経営・政策",
  });
  const choices = [
    { badge: "ア", text: "ａ：正　ｂ：正　ｃ：誤", h: 0.42 },
    { badge: "イ", text: "ａ：正　ｂ：誤　ｃ：誤", h: 0.42 },
    { badge: "ウ", text: "ａ：誤　ｂ：正　ｃ：正", h: 0.42 },
    { badge: "エ", text: "ａ：誤　ｂ：誤　ｃ：正", h: 0.42 },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 0 });
  cy += 0.06;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "正解：ア", options: { bold: true, color: RED } },
    { text: "。ａ（正）：日本料理店（飲食店）は「小売業」に区分。従業員60人は基準（50人以下）超だが、資本金3,000万円は5,000万円以下を満たしOR条件で該当。ｂ（正）：旅館業は施行令の特例で従業員基準が200人以下に緩和。資本金6,000万円は超過だが従業員80人は特例基準内でOR条件により該当。ｃ（誤）：生活関連サービス業は特例なしの通常サービス業で基準は資本金5,000万円以下", options: {} },
    { text: "または", options: { bold: true } },
    { text: "従業員100人以下。資本金8,000万円・従業員120人はいずれも超過し非該当。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.85, fontFace: F_BODY, fontSize: 9, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.12 });
  cy += 0.9;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "資本金と従業員数の「どちらか一方」を満たせば足りるというOR条件が、業種区分の落とし穴と組み合わさって出題される。", options: { fontFace: F_BODY, fontSize: 9.5, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.55, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.6;
  s.addText("関連知識：飲食店業が小売業に含まれる点、旅館業等一部業種に従業員基準の特例がある点は毎年出題されうる定番のひっかけ。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 9.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.35;
  s.addText("出典：past_exams/1st_stage/1ji2024/G1JI2024.pdf（令和6年度第1次試験）第17問設問1／正解：past_exams/1st_stage_answers/r06/2024g.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 06: G-2 各種統計にみる中小企業の動向 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-2 ／ 各種統計にみる中小企業の動向",
    title: "数は多いが、1社あたりの稼ぐ力は大企業に及ばない",
    overview: "具体的な数値は年度更新のため、この科目では「方向性」の理解が要になる。",
    tag: "中小企業経営・政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "経済センサス", v: "国内すべての事業所・企業を対象にした国の基幹統計" },
    { k: "中小企業実態基本調査", v: "中小企業庁が実施。付加価値額等の財務データを把握" },
    { k: "法人企業統計調査", v: "財務省が実施。資本金階級別の財務諸表データを集計" },
  ], { fontSize: 11.5, labelW: 2.6, gap: 0.45 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "中小企業は企業数ベースで大部分を占めるが、付加価値額シェアは企業数割合ほど高くない（数は多いが稼ぐ力は小さい）。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：G-5 中小企業の経営指標（個社レベルの分析ツール）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 07: G-3 産業構造と中小企業性業種 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-3 ／ 産業構造と中小企業、中小企業性業種",
    title: "中小企業性業種＝斜陽産業ではない",
    overview: "大企業の規模の経済が働きにくい業種に中小企業が生き残りやすい。",
    tag: "中小企業経営・政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "中小企業性業種", v: "大企業の参入が少なく中小企業が主な担い手（印刷業、繊維工業、伝統工芸品等）" },
    { k: "装置産業", v: "巨額の設備投資が必要で大企業中心（鉄鋼、石油化学、電力等）" },
  ], { fontSize: 11.5, labelW: 2.4, gap: 0.5 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「中小企業性業種＝衰退産業」は誤り。ブランド・地域性等で高収益を実現する企業も存在する。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：G-4 地域社会・地域産業（産地・企業城下町の形成）。",
    years: mkYears(new Set(["'16", "'18", "'20", "'22", "'25"])),
  });
}

// ---------- Slide 08: G-4 地域社会・地域産業と中小企業 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-4 ／ 地域社会・地域産業と中小企業",
    title: "産地と企業城下町の違い",
    overview: "地域資源活用・農商工連携も近年重視される取り組み。",
    tag: "中小企業経営・政策",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "産地", desc: "同業種の中小企業が特定地域に集積し、調達〜製造〜販売の分業体制を構築" },
    { name: "企業城下町", desc: "特定の大企業（親工場）を中心に下請企業群が地域に集積" },
  ], { nameW: 2.2, tagW: 0, rowH: 1.3 });
  addFreqBar(s, {
    y: 6.55, rank: "B", rankLabel: "頻出論点",
    related: "ひっかけ：「産地」と「企業城下町」の構造の違い（同業種集積か、親工場中心の下請構造か）を混同しない。",
    years: mkYears(new Set(["'16", "'24"])),
  });
}

// ---------- Slide 09: G-5 中小企業の経営指標 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-5 ／ 中小企業の経営指標（各種統計）",
    title: "ローカルベンチマークは「企業の健康診断」",
    overview: "財務6指標＋非財務4視点＋商流・業務フローで総合的に把握する。",
    tag: "中小企業経営・政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "財務6指標", v: "売上高増加率・営業利益率・労働生産性・EBITDA有利子負債倍率・営業運転資本回転期間・自己資本比率" },
    { k: "非財務4視点", v: "経営者・事業・関係者・内部管理体制への着目" },
  ], { fontSize: 11, labelW: 2.1, gap: 0.55 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「売上高人件費比率」「総資本回転率」等それらしい指標を紛れ込ませる出題に注意。6指標・4視点の正確な名称を覚える。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：G-2 各種統計にみる中小企業の動向（マクロ統計との対比）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 10: G-5 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-5 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第11問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "中小企業経営・政策",
  });
  const cy = addExamQuestion(s, {
    stem: "経済産業省が提供するローカルベンチマーク（「企業の健康診断」）に関する記述の正誤の組み合わせとして、最も適切なものを下記の解答群から選べ。ａ：財務分析における「6つの指標」には「売上高人件費比率」「総資本回転率」が含まれる。ｂ：非財務情報における「業務フロー」では自社業務の流れを把握し工夫や差別化ポイントを理解することが重要である。ｃ：非財務情報における「4つの視点」には「経営者への着目」「内部管理体制への着目」が含まれる。",
    stemH: 1.15,
    choices: [
      { badge: "ア", text: "ａ：正　ｂ：正　ｃ：正", h: 0.42 },
      { badge: "イ", text: "ａ：正　ｂ：誤　ｃ：正", h: 0.42 },
      { badge: "ウ", text: "ａ：誤　ｂ：正　ｃ：正", h: 0.42 },
      { badge: "エ", text: "ａ：誤　ｂ：正　ｃ：誤", h: 0.42 },
      { badge: "オ", text: "ａ：誤　ｂ：誤　ｃ：誤", h: 0.42 },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/G1JI2025.pdf（令和7年度第1次試験）第11問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 11: G-5 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-5 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第11問）",
    overview: "正解はウ。ａの2指標は6指標に含まれない。",
    tag: "中小企業経営・政策",
  });
  const choices = [
    { badge: "ア", text: "ａ：正　ｂ：正　ｃ：正", h: 0.42 },
    { badge: "イ", text: "ａ：正　ｂ：誤　ｃ：正", h: 0.42 },
    { badge: "ウ", text: "ａ：誤　ｂ：正　ｃ：正", h: 0.42 },
    { badge: "エ", text: "ａ：誤　ｂ：正　ｃ：誤", h: 0.42 },
    { badge: "オ", text: "ａ：誤　ｂ：誤　ｃ：誤", h: 0.42 },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 2 });
  cy += 0.06;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "正解：ウ", options: { bold: true, color: RED } },
    { text: "。ａ（誤）：財務6指標は①売上高増加率②営業利益率③労働生産性④EBITDA有利子負債倍率⑤営業運転資本回転期間⑥自己資本比率であり、「売上高人件費比率」「総資本回転率」は", options: {} },
    { text: "含まれない", options: { bold: true } },
    { text: "。ｂ（正）：業務フローの説明どおり正しい。ｃ（正）：非財務4視点（経営者・事業・関係者・内部管理体制）の説明どおり正しい。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.72, fontFace: F_BODY, fontSize: 9, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.12 });
  cy += 0.78;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "決算書の数字だけでなく、経営者の資質・後継者の有無、主要取引先との関係性、社内の情報共有体制まで含めて総合的に評価するのがローカルベンチマークの考え方。", options: { fontFace: F_BODY, fontSize: 9, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.45, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  cy += 0.5;
  s.addText("関連知識：財務の「6つの指標」と非財務の「4つの視点」＋「商流・業務フロー」の構成をセットで覚えること。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.35;
  s.addText("出典：past_exams/1st_stage/1ji2025/G1JI2025.pdf（令和7年度第1次試験）第11問／正解：past_exams/1st_stage_answers/r07/2025g.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 12: G-6 中小企業経営の特質と課題 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-6 ／ 中小企業経営の特質と課題（経営基盤、経営の多様性）",
    title: "「中小企業＝弱者」と単純化しない",
    overview: "積極的に成長投資・DXに取り組む企業は大企業と遜色ない例もある。",
    tag: "中小企業経営・政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "資金調達力の弱さ", v: "信用力が低く間接金融に依存。直接金融は限定的" },
    { k: "人材確保の難しさ", v: "知名度・待遇面で大企業に見劣りし採用競争で不利" },
    { k: "経営者への依存度", v: "オーナー経営者の能力・健康状態が経営全体を左右" },
  ], { fontSize: 11, labelW: 2.3, gap: 0.42 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "白書は「稼ぐ力を高めれば大企業と伍していける」と強調。「中小企業＝一律に弱い」という固定観念で解答すると矛盾する。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：G-12 経営革新・事業創造（「稼ぐ力」向上の取り組み）。",
    years: mkYears(new Set(["'16", "'21", "'22", "'23", "'25"])),
  });
}

// ---------- Slide 13: G-7 業種・業態別経営特質と課題 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-7 ／ 業種・業態別経営特質と課題",
    title: "建設業は重層下請、卸売業は「中抜き」の脅威",
    overview: "業種別の課題を取り違えないよう白書の業種別分析を押さえる。",
    tag: "中小企業経営・政策",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "建設業", desc: "元請・下請・孫請の重層構造。価格転嫁・工期管理が難しい。担い手不足" },
    { name: "製造業", desc: "系列・下請依存が多い。下請からの脱却（自社製品開発）が課題" },
    { name: "卸売業", desc: "メーカー直接取引拡大・EC化による「中抜き」の脅威" },
    { name: "サービス業", desc: "労働集約的な業種が多く、労働生産性向上が大きな政策課題" },
  ], { nameW: 1.8, tagW: 0, rowH: 0.9 });
  addFreqBar(s, {
    y: 6.55, rank: "－", rankLabel: "直近10年単独出題なし",
    related: "関連：G-23 経営サポート（取引・官公需支援は下請問題への対応）。",
    years: mkYears(new Set([])),
  });
}

// ---------- Slide 14: G-8 中小企業の経営環境と経営革新 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-8 ／ 中小企業の経営環境と経営革新",
    title: "グローバル化＝海外進出が常に望ましい、ではない",
    overview: "供給網のリスク管理（分散・強靱化）という文脈で語られる。",
    tag: "中小企業経営・政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "サプライチェーンの強靱化", v: "調達先の分散、国内回帰（リショアリング）、在庫適正化" },
    { k: "価格転嫁の推進", v: "原材料費・エネルギー・人件費上昇分を取引価格に反映" },
  ], { fontSize: 11.5, labelW: 2.8, gap: 0.5 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「グローバル化＝海外進出推進が常に望ましい」ではなく、リスク分散・強靱化の文脈で語られる点に注意。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "B", rankLabel: "頻出論点",
    related: "関連：G-15 海外展開（企業レベルの海外展開の進め方）。",
    years: mkYears(new Set(["'23", "'24"])),
  });
}

// ---------- Slide 15: G-9 金融・信用保証、税制 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-9 ／ 金融・信用保証、税制",
    title: "「中小企業者」と「中小法人」は別概念",
    overview: "税制優遇の多くは業種問わず資本金1億円以下で一律判定される。",
    tag: "中小企業経営・政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "間接金融依存", v: "中小企業は信用力が乏しく銀行借入への依存度が高い" },
    { k: "中小企業基本法の基準", v: "業種ごとに資本金または従業員数のいずれかで判定" },
    { k: "租税特別措置法の基準", v: "業種問わず資本金1億円以下で一律判定（中小法人）" },
  ], { fontSize: 10.5, labelW: 2.7, gap: 0.4 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「中小企業基本法上の中小企業者」と「租税特別措置法上の中小法人」は判定基準が異なる別概念。税制の話は必ずどちらの基準か確認する。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：G-25 財務サポート（税制、事業承継税制の詳細）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 16: G-9 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-9 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第27問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "中小企業経営・政策",
  });
  const cy = addExamQuestion(s, {
    stem: "交際費等は原則として損金の額に算入しないこととされている。ただし特例として、【Ａ】は800万円までの交際費等の【Ｂ】の損金算入、または接待飲食費の50％の損金算入の選択適用が認められている。空欄ＡとＢに入る語句の組み合わせとして、最も適切なものを下記の解答群から選べ。",
    stemH: 0.85,
    choices: [
      { badge: "ア", text: "Ａ：資本金または出資金の額が１億円以下の法人　Ｂ：50％", h: 0.42 },
      { badge: "イ", text: "Ａ：資本金または出資金の額が１億円以下の法人　Ｂ：全額", h: 0.42 },
      { badge: "ウ", text: "Ａ：中小企業基本法に定める中小企業に該当する法人　Ｂ：50％", h: 0.42 },
      { badge: "エ", text: "Ａ：中小企業基本法に定める中小企業に該当する法人　Ｂ：全額", h: 0.42 },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/G1JI2025.pdf（令和7年度第1次試験）第27問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 17: G-9 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-9 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第27問）",
    overview: "正解はイ。判定基準は資本金1億円以下、控除枠は全額損金算入。",
    tag: "中小企業経営・政策",
  });
  const choices = [
    { badge: "ア", text: "Ａ：資本金または出資金の額が１億円以下の法人　Ｂ：50％", h: 0.42 },
    { badge: "イ", text: "Ａ：資本金または出資金の額が１億円以下の法人　Ｂ：全額", h: 0.42 },
    { badge: "ウ", text: "Ａ：中小企業基本法に定める中小企業に該当する法人　Ｂ：50％", h: 0.42 },
    { badge: "エ", text: "Ａ：中小企業基本法に定める中小企業に該当する法人　Ｂ：全額", h: 0.42 },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 1 });
  cy += 0.06;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "正解：イ", options: { bold: true, color: RED } },
    { text: "。租税特別措置法により、資本金（出資金）", options: {} },
    { text: "1億円以下", options: { bold: true } },
    { text: "の法人（大法人の100％子会社等除く）は、年800万円までの交際費等の", options: {} },
    { text: "全額", options: { bold: true } },
    { text: "損金算入、または接待飲食費50％相当額の損金算入のいずれか有利な方を選択適用できる。ア：Ａは正しいがＢの「50％」は誤り。ウ・エ：Ａの「中小企業基本法に定める中小企業」が誤り（判定基準は資本金1億円以下の一律基準）。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.9, fontFace: F_BODY, fontSize: 9, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.12 });
  cy += 0.95;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "資本金3,000万円の製造業も、資本金1億円ぴったりの卸売業も、いずれも資本金1億円以下という税制上の基準で判定される。", options: { fontFace: F_BODY, fontSize: 9, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.4, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  cy += 0.45;
  s.addText("関連知識：税制上の中小企業向け特例（軽減税率、交際費特例、少額減価償却資産の特例等）の多くは「資本金1億円以下」を基準にしている。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.35;
  s.addText("出典：past_exams/1st_stage/1ji2025/G1JI2025.pdf（令和7年度第1次試験）第27問／正解：past_exams/1st_stage_answers/r07/2025g.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 18: G-10 労働・雇用・人材育成 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-10 ／ 労働・雇用・人材育成",
    title: "人手不足対応は「採用を増やす」だけではない",
    overview: "多様な人材の活用と省力化・生産性向上の両輪で対応する。",
    tag: "中小企業経営・政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "多様な人材の活用", v: "女性・高齢者・外国人材・副業兼業人材の労働参加を促す" },
    { k: "省力化・生産性向上", v: "DX・設備投資で少ない人手でも成果を上げる" },
    { k: "賃上げのジレンマ", v: "定着には賃上げが必要だが価格転嫁力の弱さが原資確保を阻む" },
  ], { fontSize: 11, labelW: 2.3, gap: 0.42 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「人手不足対応＝採用強化のみ」は不十分。省力化・多様な人材活用・賃上げの組合せが白書の問題意識。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：G-22 経営サポート（雇用人材支援策）。",
    years: mkYears(new Set(["'16", "'19", "'21", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 19: G-11 環境対応・エネルギー対応 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-11 ／ 環境対応・エネルギー対応",
    title: "環境対応はコスト増だけの話ではない",
    overview: "取引先要請への対応リスクとコスト削減機会の両面がある。",
    tag: "中小企業経営・政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "カーボンニュートラル", v: "2050年温室効果ガス排出実質ゼロを目指す政府目標" },
    { k: "サプライチェーン波及", v: "大企業の環境対応要請が取引先の中小企業にも波及" },
    { k: "GX関連補助金", v: "省エネ設備投資・再エネ導入の支援策" },
  ], { fontSize: 11, labelW: 2.6, gap: 0.42 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「負担が増えるだけ」という一面的理解は誤り。取引喪失リスクとコスト削減機会の両面から出題される。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "－", rankLabel: "直近10年単独出題なし",
    related: "関連：G-8 中小企業の経営環境（サプライチェーン強靱化と同時に語られる）。",
    years: mkYears(new Set([])),
  });
}

// ---------- Slide 20: G-12 経営革新・事業創造、生産性向上 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-12 ／ 経営革新・事業創造、生産性向上の取り組み",
    title: "生産性向上＝人員削減ではない",
    overview: "「現状維持は最大のリスク」。付加価値額の増加と労働投入量の最適化の両輪。",
    tag: "中小企業経営・政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "付加価値額の増加", v: "価格転嫁、成長投資、新事業展開、事業承継・M&Aによる再編" },
    { k: "労働投入量の最適化", v: "省力化投資、DX推進による生産性向上" },
  ], { fontSize: 11.5, labelW: 2.6, gap: 0.5 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「生産性向上＝人員削減」は誤り。限られた人手を高付加価値業務に振り向けることが本質。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：G-20 経営サポート（経営革新計画の認定制度）。",
    years: mkYears(new Set(["'16", "'18", "'19", "'21"])),
  });
}

// ---------- Slide 21: G-13 情報技術の活用・DX ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-13 ／ 情報技術の活用・DX",
    title: "ITツール導入＝DXではない",
    overview: "デジタイゼーション→デジタライゼーション→DXという段階の違いを意識する。",
    tag: "中小企業経営・政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "DX", v: "デジタル技術で業務プロセス・ビジネスモデルを変革し競争優位を確立" },
    { k: "DX認定制度", v: "企業のDXへの取り組み方針を国が認定する制度" },
  ], { fontSize: 11.5, labelW: 2.3, gap: 0.5 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "ツール導入自体はDXの手段の一つ。業務プロセス・事業モデルの変革を伴って初めてDXと呼べる。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：F-13 経営革新と情報システム（F科目のDX論点と同一テーマ）。",
    years: mkYears(new Set(["'17", "'20", "'22", "'25"])),
  });
}

// ---------- Slide 22: G-14 事業承継・M&A ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-14 ／ 事業承継・M&A",
    title: "後継者不在は経営権・資産・税負担の3面から考える",
    overview: "第三者承継（M&A）が政策的に強く後押しされている。",
    tag: "中小企業経営・政策",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "親族内承継", desc: "経営者の子など親族に引き継ぐ" },
    { name: "親族外承継", tag: "従業員承継", desc: "役員・従業員に引き継ぐ" },
    { name: "社外への引継ぎ", tag: "M&A・第三者承継", desc: "外部の企業・個人に譲渡。近年最も政策的に後押しされている" },
  ], { nameW: 2.0, tagW: 2.3, rowH: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：G-25 財務サポート（事業承継税制、税負担の軽減策）。",
    years: mkYears(new Set(["'17", "'18", "'20", "'21", "'22", "'23"])),
  });
}

// ---------- Slide 23: G-15 海外展開 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-15 ／ 海外展開",
    title: "まず輸出から始める段階的アプローチが重視される",
    overview: "JETRO等の公的支援を活用しながら海外展開を進めるケースが多い。",
    tag: "中小企業経営・政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "海外展開の段階", v: "輸出→ライセンス供与→海外直接投資（現地法人・工場設立）" },
    { k: "JETRO", v: "日本貿易振興機構。中小企業の海外展開を支援する公的機関" },
  ], { fontSize: 11.5, labelW: 2.3, gap: 0.5 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "いきなり現地生産ではなく、輸出から始め市場の反応を見る手堅いアプローチが重視される。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：G-22 経営サポート（海外展開支援の具体的施策）。",
    years: mkYears(new Set(["'16", "'17", "'20", "'22", "'24"])),
  });
}

// ---------- Slide 24: G-16 創業・アントレプレナーシップ ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-16 ／ 創業・アントレプレナーシップ",
    title: "「創業」と「第二創業」は異なる概念",
    overview: "日本は諸外国比で開業率が低く、創業の裾野拡大が政策課題。",
    tag: "中小企業経営・政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "開業率", v: "日本は諸外国と比較して低水準にあると指摘される" },
    { k: "アントレプレナーシップ", v: "リスクを取って新しい事業に挑戦する起業家精神そのもの" },
    { k: "第二創業", v: "既存企業が新分野に進出すること（創業とは異なる概念）" },
  ], { fontSize: 10.5, labelW: 2.6, gap: 0.4 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "創業支援策の対象が「新規開業者」か「既存企業の新事業展開（第二創業）」かを区別する。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：G-19 経営サポート（創業・ベンチャー支援の具体施策）。",
    years: mkYears(new Set(["'18", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 25: G-17 知的財産権・無形資産の活用 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-17 ／ 知的財産権・無形資産の活用",
    title: "G科目で問われるのは「経営における活用のあり方」",
    overview: "知財の制度内容そのものはE科目（経営法務）の範囲。",
    tag: "中小企業経営・政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "無形資産", v: "ブランド・ノウハウ・顧客基盤・データ・人的資本等、B/Sに現れにくい価値" },
    { k: "知財経営", v: "独自の技術・ブランドを経営戦略に明確に位置づけ活用する考え方" },
    { k: "地域団体商標", v: "地域名＋商品名の商標を組合等が登録できる制度" },
  ], { fontSize: 10.5, labelW: 2.3, gap: 0.4 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "特許権・商標権の制度要件はE科目の範囲。G科目は「経営における活用のあり方」というより経営寄りの視点で問われる。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：G-21 経営サポート（知的財産支援策）。",
    years: mkYears(new Set(["'20", "'21", "'23"])),
  });
}

pres.writeFile({ fileName: path.join(__dirname, "../../slides/1st_stage/G_sme_management_policy.pptx") })
  .then(() => console.log("wrote", "slides/1st_stage/G_sme_management_policy.pptx"));
