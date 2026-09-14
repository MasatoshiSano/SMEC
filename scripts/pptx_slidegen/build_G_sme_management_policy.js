// Builds slides/1st_stage/G_sme_management_policy.pptx from lib.js components.
// Run: node scripts/pptx_slidegen/build_G_sme_management_policy.js
//
// Mirrors slides/1st_stage/G_sme_management_policy.html. See docs/13_pptx_slide_template_spec.md.

const path = require("path");
const {
  newPres, addCoverSlide, addDividerSlide, addHeader, addFreqBar, mkYears,
  addTermRows, addRowList, addExamQuestion, drawProcessStepsVertical,
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
  addRowList(s, 0.55, 1.9, 12.25, [
    { name: "製造業・建設業・運輸業等", tag: "資本金3億円以下", desc: "または常時使用する従業員300人以下のいずれかを満たせば中小企業者に該当" },
    { name: "卸売業", tag: "資本金1億円以下", desc: "または従業員100人以下。メーカーと小売の中間流通を担う業種区分" },
    { name: "サービス業", tag: "資本金5,000万円以下", desc: "または従業員100人以下。旅館業等は施行令で従業員基準が200人以下に緩和される特例あり" },
    { name: "小売業（飲食店業含む）", tag: "資本金5,000万円以下", desc: "または従業員50人以下。飲食店業もこの区分に含まれる" },
    { name: "小規模企業者", tag: "製造業等20人以下", desc: "商業・サービス業は従業員5人以下。中小企業の中でもさらに小さい区分" },
    { name: "特例業種の例", tag: "ソフトウェア業等", desc: "資本金3億円以下・従業員300人以下と、通常のサービス業基準より緩和" },
  ], { nameW: 2.7, tagW: 2.1, rowH: 0.75 });
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
    { k: "経済センサス", v: "国内すべての事業所・企業を対象にした国の基幹統計。企業数・従業者数・売上高を網羅的に把握" },
    { k: "中小企業実態基本調査", v: "中小企業庁が実施。法人・個人事業者を対象に付加価値額等の財務データを把握" },
    { k: "法人企業統計調査", v: "財務省が実施。資本金階級別の財務諸表データを集計" },
    { k: "出題パターン", v: "具体的な数値そのものを問うタイプと、増加/減少・大企業比の高低という方向性を問うタイプがある" },
  ], { fontSize: 10.5, labelW: 2.4, gap: 0.42 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "中小企業は企業数ベースで大部分を占めるが、付加価値額シェアは企業数割合ほど高くない（数は多いが稼ぐ力は小さい）。数値を覚えきれなくても白書のメインメッセージから方向性を掴んでおくこと。", options: { color: RED } },
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
    { k: "中小企業性業種", v: "大企業の参入が少なく中小企業が主な担い手（印刷業、木材・木製品製造業、繊維工業、家具製造業、伝統的工芸品産業等）" },
    { k: "生き残る理由", v: "多品種少量生産に向く、または地域の資源・技術に根ざし、大企業の規模の経済が働きにくい" },
    { k: "装置産業", v: "巨額の設備投資が必要で大企業中心（鉄鋼、石油化学、電力等）" },
  ], { fontSize: 10.5, labelW: 2.2, gap: 0.42 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「中小企業性業種＝衰退産業」は誤り。酒蔵・陶磁器産地のようにブランド・地域性・匠の技術で高収益を実現する企業も存在する。", options: { color: RED } },
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
    { name: "産地", desc: "同業種の中小企業が特定地域に集積し、調達〜製造〜販売の分業体制を構築（例：眼鏡フレーム、刃物の産地）" },
    { name: "企業城下町", desc: "特定の大企業（親工場）を中心に下請企業群が地域に集積する型の産業集積" },
    { name: "地域資源活用・農商工連携", desc: "地域外の企業・大学・金融機関と連携し、農産品・観光資源・伝統技術を活かした新事業を創出する近年の取り組み" },
  ], { nameW: 3.4, tagW: 0, rowH: 0.9 });
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
    { k: "財務6指標", v: "①売上高増加率②営業利益率③労働生産性④EBITDA有利子負債倍率⑤営業運転資本回転期間⑥自己資本比率" },
    { k: "非財務4視点", v: "①経営者への着目②事業への着目③関係者への着目④内部管理体制への着目、の4観点で分析" },
    { k: "商流・業務フロー", v: "自社の業務の流れを把握し、工夫や差別化ポイントを理解すること。非財務分析の一部" },
    { k: "ツールの目的", v: "決算書の数字だけでなく将来性や潜在的リスクを含め総合的に企業を評価すること" },
  ], { fontSize: 10, labelW: 1.9, gap: 0.4 });
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
    { k: "資金調達力の弱さ", v: "信用力が大企業より低く間接金融（借入）に依存しがち。直接金融は限定的" },
    { k: "人材確保の難しさ", v: "知名度・待遇面で大企業に見劣りし、採用競争で不利になりやすい" },
    { k: "情報収集力の限界", v: "市場動向・法改正・技術動向を収集・分析する専任部署を持てないことが多い" },
    { k: "経営者への依存度の高さ", v: "オーナー経営者の能力・健康状態が経営全体を大きく左右する" },
    { k: "経営の多様性", v: "業種・規模・経営者の考え方によって経営のあり方が極めて多様な点も特徴" },
  ], { fontSize: 9.5, labelW: 2.4, gap: 0.36 });
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
    { k: "グローバル化の進展と反動", v: "貿易・投資拡大が進む一方、地政学リスクや感染症拡大で供給網の海外依存リスクが再認識" },
    { k: "サプライチェーンの強靱化", v: "調達先の分散、国内回帰（リショアリング）、在庫の適正化" },
    { k: "価格転嫁の推進", v: "原材料費・エネルギー価格・人件費上昇分を取引価格に適切に反映させること" },
  ], { fontSize: 10.5, labelW: 2.8, gap: 0.42 });
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
  const proseX = 0.55, proseW = 7.4;
  const diagX = 8.25, diagW = 4.05;
  let cy = 1.85;
  s.addText(
    "中小企業は信用力が乏しく間接金融（借入）に依存しがちなため、信用保証協会が保証人になる信用保証制度（詳細はG-24）が整備されている。税制面では、右図のとおり「どの法律の基準で判定するか」を混同しないことが最大のポイント。",
    { x: proseX, y: cy, w: proseW, h: 0.9, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 1.0;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "間接金融依存", v: "担保力・情報開示力が乏しく、直接金融より銀行借入への依存度が高い" },
    { k: "信用保証制度", v: "信用保証協会が公的な保証人になり融資を後押しする仕組み" },
  ], { fontSize: 9.5, labelW: 2.1, gap: 0.42 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「中小企業基本法上の中小企業者」と「租税特別措置法上の中小法人」は判定基準が異なる別概念。税制の話は必ずどちらの基準か確認する。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, fontFace: F_BODY, fontSize: 9, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  const boxY = 1.9, boxH = 3.9, boxW = (diagW - 0.15) / 2;
  s.addShape("rect", { x: diagX, y: boxY, w: boxW, h: boxH, fill: { color: GHOST }, line: { color: INK, width: 1 } });
  s.addText([
    { text: "中小企業基本法\n\n", options: { bold: true, fontSize: 11, color: INK, breakLine: true } },
    { text: "業種ごとに\n資本金 or 従業員数\n\n", options: { fontSize: 8.5, color: INK_SOFT, breakLine: true } },
    { text: "「中小企業者」\n", options: { bold: true, fontSize: 9.5, color: INK, breakLine: true } },
    { text: "（業種別・G-1）", options: { fontSize: 7.5, color: INK_SOFT } },
  ], { x: diagX + 0.05, y: boxY, w: boxW - 0.1, h: boxH, align: "center", valign: "middle", fontFace: F_BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  const box2X = diagX + boxW + 0.15;
  s.addShape("rect", { x: box2X, y: boxY, w: boxW, h: boxH, fill: { color: RED }, line: { type: "none" } });
  s.addText([
    { text: "租税特別措置法\n\n", options: { bold: true, fontSize: 11, color: "FFFFFF", breakLine: true } },
    { text: "業種問わず\n資本金1億円以下\n\n", options: { fontSize: 8.5, color: "FFFFFF", breakLine: true } },
    { text: "「中小法人」\n", options: { bold: true, fontSize: 9.5, color: "FFFFFF", breakLine: true } },
    { text: "（一律・税制優遇）", options: { fontSize: 7.5, color: "FFFFFF" } },
  ], { x: box2X + 0.05, y: boxY, w: boxW - 0.1, h: boxH, align: "center", valign: "middle", fontFace: F_BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  s.addText("判定基準が違う「別の中小企業」の話", { x: diagX, y: boxY + boxH + 0.1, w: diagW, h: 0.3, align: "center", fontFace: F_BODY, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0 });

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
    { k: "労働供給制約社会", v: "働き手そのものが減っていく社会の到来で、今後さらに人手不足が深刻化する懸念" },
    { k: "多様な人材の活用", v: "女性・高齢者・外国人材・副業兼業人材など活用されてこなかった層の労働参加を促す" },
    { k: "省力化・生産性向上", v: "人手を増やすのではなくDX・設備投資で少ない人手でも成果を上げる" },
    { k: "賃上げのジレンマ", v: "定着には賃上げが不可欠だが価格転嫁力の弱さが原資（利益）確保を阻む" },
  ], { fontSize: 10, labelW: 2.4, gap: 0.4 });
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
    { k: "稼ぐ力", v: "中小企業白書が一貫して掲げる重要テーマ（付加価値を生み出す力）の向上" },
    { k: "付加価値額の増加", v: "価格転嫁、成長投資、新事業展開、事業承継・M&Aによる事業再編" },
    { k: "労働投入量の最適化", v: "省力化投資、DX推進による生産性向上" },
    { k: "現状維持は最大のリスク", v: "短期損益にとらわれず長期視点で事業・組織構造を見直す「戦略を持った経営」への転換" },
  ], { fontSize: 10, labelW: 2.5, gap: 0.4 });
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
  const proseX = 0.55, proseW = 7.4;
  const diagX = 8.25, diagW = 4.05;
  let cy = 1.85;
  s.addText(
    "DXは「ITツールを導入すること」自体ではなく、デジタル技術を土台に業務プロセスやビジネスモデルそのものを変革し、競争優位を確立する取り組みを指す。右図の3段階のうち、どこまで進んでいるかが問われる。",
    { x: proseX, y: cy, w: proseW, h: 0.85, fontFace: F_BODY, fontSize: 10.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.95;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "DX認定制度", v: "企業のDXへの取り組み方針を経済産業省が認定する制度" },
    { k: "中小企業のDXの遅れ", v: "大企業に比べ人材・予算の制約から取り組みが遅れがち" },
  ], { fontSize: 10, labelW: 2.3, gap: 0.4 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "ツール導入自体はDXの手段の一つ。業務プロセス・事業モデルの変革を伴って初めてDXと呼べる。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  drawProcessStepsVertical(s, diagX, 1.9, diagW, 4.0, [
    { num: "①", label: "デジタイゼーション", desc: "紙を単純にデータ化するだけ" },
    { num: "②", label: "デジタライゼーション", desc: "個々の業務プロセスを効率化" },
    { num: "③", label: "DX（事業変革）", desc: "組織・ビジネスモデルを変革" },
  ]);

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
    { k: "海外展開の段階", v: "輸出→ライセンス供与→海外直接投資（現地法人・工場設立）の順に段階を踏む" },
    { k: "中小企業特有の制約", v: "大企業に比べノウハウ・人材・資金が限られるため公的機関の支援を活用しながら進める" },
    { k: "JETRO", v: "日本貿易振興機構。海外市場調査・商談会や見本市への出展支援等を行う中核機関" },
  ], { fontSize: 10.5, labelW: 2.5, gap: 0.42 });
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

// ---------- Slide 26: G-18〜G-28 区切り ----------
addDividerSlide(pres, {
  ghostNo: "02",
  partNo: "PART 02",
  partLabel: "中小企業経営・政策 ／ G-18〜G-28",
  title: "中小企業政策",
  desc: "国や自治体が中小企業をどう支援しているか。支援策の名前は年度で変わるが、「どんな課題に、どんな支援の枠組みがあるか」という制度の骨格を押さえる。",
  chips: ["G-20 経営革新計画", "G-24 信用保証制度", "G-25 事業承継税制"],
  notes: "中小企業政策パートの区切りスライド。",
});

// ---------- Slide 27: G-18 中小企業関連法規 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-18 ／ 中小企業関連法規（中小企業基本法、中小企業支援法等）",
    title: "「憲法」「支援体制」「計画認定」で役割が異なる",
    overview: "名称が似た3つの法律の役割を区別する。",
    tag: "中小企業経営・政策",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "中小企業基本法", desc: "中小企業政策の理念・基本方針を定める「憲法」。中小企業者の定義もここに規定" },
    { name: "中小企業支援法", desc: "都道府県等中小企業支援センター等、支援体制の整備を定める" },
    { name: "中小企業等経営強化法", desc: "経営革新計画・経営力向上計画等、個々の企業の取り組みを認定・支援する枠組み" },
  ], { nameW: 2.6, tagW: 0, rowH: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：3つの法律名は似ているが役割が異なる別の法律。「理念・定義」「支援体制」「個別計画認定」で区別する。",
    years: mkYears(new Set(["'16", "'17", "'18", "'20", "'22", "'24"])),
  });
}

// ---------- Slide 28: G-19 創業・ベンチャー支援 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-19 ／ 経営サポート：創業・ベンチャー支援",
    title: "年齢要件の数値を微妙にずらすひっかけが定番",
    overview: "日本公庫の創業融資は女性・若年者・高齢者に特別利率を適用。",
    tag: "中小企業経営・政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "新規開業・スタートアップ支援資金", v: "日本政策金融公庫国民生活事業の創業融資制度。貸付利率などに特例を設けて創業を支援" },
    { k: "対象者", v: "新規開業しようとする者、または新規開業して概ね7年以内の者" },
    { k: "特別利率の対象", v: "女性、または若年者・高齢者。運転資金・設備資金（土地取得資金を除く）に適用" },
    { k: "創業支援等事業計画", v: "市区町村が民間の創業支援機関と連携して策定。認定計画に基づく創業者には登録免許税の軽減等の優遇" },
  ], { fontSize: 9.5, labelW: 3.3, gap: 0.4 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "若年者・高齢者の年齢基準は選択肢で数値をずらして出題される定番パターン。制度は改定されうる前提も忘れずに。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：G-16 創業・アントレプレナーシップ（経営面の一般的な考え方）。",
    years: mkYears(new Set(["'18", "'19", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 29: G-19 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-19 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第20問設問1）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "中小企業経営・政策",
  });
  const cy = addExamQuestion(s, {
    stem: "日本政策金融公庫国民生活事業が実施している新規開業・スタートアップ支援資金は、女性、若年者または高齢者の場合には、運転資金と設備資金（土地取得資金を除く）に特別利率が適用される。「若年者または高齢者」の要件の定めとして、最も適切なものはどれか。",
    stemH: 0.85,
    choices: [
      { badge: "ア", text: "若年者は30歳未満、高齢者は55歳以上。", h: 0.42 },
      { badge: "イ", text: "若年者は30歳未満、高齢者は65歳以上。", h: 0.42 },
      { badge: "ウ", text: "若年者は35歳未満、高齢者は55歳以上。", h: 0.42 },
      { badge: "エ", text: "若年者は35歳未満、高齢者は60歳以上。", h: 0.42 },
      { badge: "オ", text: "若年者は35歳未満、高齢者は65歳以上。", h: 0.42 },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/G1JI2025.pdf（令和7年度第1次試験）第20問設問1", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 30: G-19 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-19 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第20問設問1）",
    overview: "正解はウ。若年者35歳未満、高齢者55歳以上（令和7年度時点）。",
    tag: "中小企業経営・政策",
  });
  const choices = [
    { badge: "ア", text: "若年者は30歳未満、高齢者は55歳以上。", h: 0.42 },
    { badge: "イ", text: "若年者は30歳未満、高齢者は65歳以上。", h: 0.42 },
    { badge: "ウ", text: "若年者は35歳未満、高齢者は55歳以上。", h: 0.42 },
    { badge: "エ", text: "若年者は35歳未満、高齢者は60歳以上。", h: 0.42 },
    { badge: "オ", text: "若年者は35歳未満、高齢者は65歳以上。", h: 0.42 },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 2 });
  cy += 0.06;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "正解：ウ", options: { bold: true, color: RED } },
    { text: "。令和7年度時点、日本政策金融公庫国民生活事業の新規開業・スタートアップ支援資金では、若年者は", options: {} },
    { text: "35歳未満", options: { bold: true } },
    { text: "、高齢者は", options: {} },
    { text: "55歳以上", options: { bold: true } },
    { text: "が特別利率適用の年齢要件。ア・イ・エ・オはいずれかの年齢基準が実際の制度と異なる。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, fontFace: F_BODY, fontSize: 9, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.12 });
  cy += 0.76;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "「女性、若者/シニア起業家」向けに優遇金利を設ける枠組みは日本公庫の他の融資制度でもたびたび登場する。年齢基準は制度改定で変わりうるため最新情報を確認すること。", options: { fontFace: F_BODY, fontSize: 9, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.45, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  cy += 0.5;
  s.addText("関連知識：年齢要件の数値を微妙にずらす出題は定番パターン。正確な数値を覚えつつ「制度は改定されうる」前提も忘れないこと。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.35;
  s.addText("出典：past_exams/1st_stage/1ji2025/G1JI2025.pdf（令和7年度第1次試験）第20問設問1／正解：past_exams/1st_stage_answers/r07/2025g.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 31: G-20 経営革新支援、新事業展開支援 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-20 ／ 経営サポート：経営革新支援、新事業展開支援",
    title: "新規性は「全国初」ではなく自社基準でよい",
    overview: "新事業活動は5類型のいずれかに該当すれば対象になる。",
    tag: "中小企業経営・政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "経営革新計画", v: "新事業活動に取り組む計画。都道府県知事等の承認で特別貸付・保証特例" },
    { k: "新事業活動5類型", v: "新商品開発／新役務開発／新生産販売方式／新提供方式／新経営管理方法" },
    { k: "経営力向上計画", v: "生産性向上に着目。国の認定で経営強化税制等の税制優遇と結びつく" },
  ], { fontSize: 10.5, labelW: 2.4, gap: 0.4 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「新規性＝全国初でなければならない」は誤り。既に他社が採用済みでも自社にとって新しければ対象。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：G-12 経営革新・事業創造（「稼ぐ力」向上との連続性）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 32: G-20 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-20 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第29問設問1）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "中小企業経営・政策",
  });
  const cy = addExamQuestion(s, {
    stem: "経営革新計画の「該当する事業内容」に関する記述の正誤の組み合わせとして、最も適切なものを下記の解答群から選べ。ａ：商品の新たな販売方式の導入は対象となる。ｂ：自社にとって新しいものであれば、他社で採用されているものも対象となる。",
    stemH: 0.75,
    choices: [
      { badge: "ア", text: "ａ：正　ｂ：正", h: 0.42 },
      { badge: "イ", text: "ａ：正　ｂ：誤", h: 0.42 },
      { badge: "ウ", text: "ａ：誤　ｂ：正", h: 0.42 },
      { badge: "エ", text: "ａ：誤　ｂ：誤", h: 0.42 },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/G1JI2025.pdf（令和7年度第1次試験）第29問設問1", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 33: G-20 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-20 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第29問設問1）",
    overview: "正解はア。ａ・ｂともに正しい記述。",
    tag: "中小企業経営・政策",
  });
  const choices = [
    { badge: "ア", text: "ａ：正　ｂ：正", h: 0.42 },
    { badge: "イ", text: "ａ：正　ｂ：誤", h: 0.42 },
    { badge: "ウ", text: "ａ：誤　ｂ：正", h: 0.42 },
    { badge: "エ", text: "ａ：誤　ｂ：誤", h: 0.42 },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 0 });
  cy += 0.06;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "正解：ア", options: { bold: true, color: RED } },
    { text: "。ａ（正）：「商品の新たな販売方式の導入」は新事業活動5類型の③にそのまま該当する。ｂ（正）：新事業活動の新規性は業界全体・全国で見て新しいことまでは要求されず、", options: {} },
    { text: "その中小企業者にとって新しい", options: { bold: true } },
    { text: "取り組みであれば足りる（既に他社が採用済みでも対象になりうる）。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.72, fontFace: F_BODY, fontSize: 9.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  cy += 0.78;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "ある町工場が他の工場で既に広く使われている加工方式でも、自社にとって新しい取り組みであれば経営革新計画の「新事業活動」に該当しうる。", options: { fontFace: F_BODY, fontSize: 9.5, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.55;
  s.addText("関連知識：経営革新計画の経営目標は「付加価値額（または一人当たり付加価値額）」と「給与支給総額」の両方について事業期間に応じた伸び率を満たす必要がある。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.35;
  s.addText("出典：past_exams/1st_stage/1ji2025/G1JI2025.pdf（令和7年度第1次試験）第29問設問1／正解：past_exams/1st_stage_answers/r07/2025g.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 34: G-21 知的財産支援、再生支援 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-21 ／ 経営サポート：知的財産支援、再生支援",
    title: "再生支援は中小企業活性化協議会が中心的役割",
    overview: "M&A支援は事業承継・引継ぎ支援センターと役割の重心が異なる。",
    tag: "中小企業経営・政策",
  });
  addRowList(s, 0.55, 1.9, 12.25, [
    { name: "知的財産支援", desc: "中小企業が特許出願する際の審査請求料・特許料の減免制度等、コスト面での支援" },
    { name: "中小企業活性化協議会", desc: "各都道府県に設置。資金繰りに窮した企業の相談を受け、専門家とともに再生計画策定・金融機関との調整を支援（旧：中小企業再生支援協議会）" },
    { name: "事業承継・引継ぎ支援センター", desc: "各都道府県に設置。事業承継の第三者承継（M&A）を支援する専門機関" },
  ], { nameW: 3.2, tagW: 0, rowH: 1.5 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：「再生支援＝活性化協議会」「M&Aマッチング＝引継ぎ支援センター」を対応させて覚える。連携はするが役割の重心が違う。",
    years: mkYears(new Set(["'16", "'18", "'19"])),
  });
}

// ---------- Slide 35: G-22 雇用人材支援、海外展開支援 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-22 ／ 経営サポート：雇用人材支援、海外展開支援",
    title: "JETROは海外展開支援の中核機関",
    overview: "当年度の施策名・数値目標は白書・施策総覧で都度確認する。",
    tag: "中小企業経営・政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "雇用人材支援", v: "雇用調整助成金、人材確保等支援助成金等、人材の確保・育成を後押しする各種助成金" },
    { k: "人材育成機関の活用", v: "生産性向上に資する人材育成を行う機関の活用支援" },
    { k: "JETRO", v: "日本貿易振興機構。海外市場調査、商談会・見本市出展支援、現地情報提供を行う中核機関" },
    { k: "当年度の代表施策", v: "「新規輸出1万者支援プログラム」等、輸出初挑戦企業の裾野拡大施策（施策名は年度で変わる）" },
  ], { fontSize: 9.5, labelW: 2.4, gap: 0.4 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "海外展開の一般的な考え方（G-15）と、当年度の具体的な支援施策名（G-22）は分けて整理する。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：G-15 海外展開（企業レベルの海外展開の進め方）。",
    years: mkYears(new Set(["'16", "'17", "'19", "'22"])),
  });
}

// ---------- Slide 36: G-23 取引・官公需支援、小規模企業支援 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-23 ／ 経営サポート：取引・官公需支援、小規模企業支援",
    title: "下請法は民間取引、官公需法は国・自治体の調達",
    overview: "相手方が民間か国・自治体かで対象法律が異なる。",
    tag: "中小企業経営・政策",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "下請代金支払遅延等防止法", desc: "親事業者の支払遅延・不当減額・買いたたき等を規制（民間の取引）" },
    { name: "官公需法", desc: "国・地方公共団体の調達で中小企業者の受注機会確保に努力義務" },
    { name: "小規模企業振興基本法", desc: "2014年制定。小規模企業の持続的発展を支える政策の柱" },
  ], { nameW: 3.0, tagW: 0, rowH: 0.9 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：価格交渉促進月間・パートナーシップ構築宣言も取引適正化の一環。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 37: G-24 金融サポート ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-24 ／ 金融サポート（信用保証制度、政府系金融機関）",
    title: "普通保証2億円、無担保保証8,000万円は別枠",
    overview: "一般保証とセーフティネット保証・小口零細企業保証は別の仕組み。",
    tag: "中小企業経営・政策",
  });
  const proseX = 0.55, proseW = 7.4;
  const diagX = 8.25, diagW = 4.05;
  let cy = 1.85;
  s.addText(
    "一般保証は「担保の有無」で2つの枠に分かれ、さらに小規模事業者向けにもう一段小さい枠がある。3種類の限度額の大小関係を右図のイメージで押さえておくと、名称の入れ替えに惑わされにくい。",
    { x: proseX, y: cy, w: proseW, h: 0.95, fontFace: F_BODY, fontSize: 10.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 }
  );
  cy += 1.05;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "日本政策金融公庫", v: "国民生活事業（個人・小規模）と中小企業事業の2部門" },
  ], { fontSize: 10.5, labelW: 2.1, gap: 0.42 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "普通保証・無担保保証・無担保無保証人保証（小口零細企業保証）の限度額と名称を正確に区別する。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.7, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  // 3-bar comparison diagram: 普通保証(2億円) > 無担保保証(8,000万円) > 小口零細(具体的な限度額の記載なし)
  {
    const baseY = 1.9, baseH = 4.0;
    const axisY = baseY + baseH - 0.35;
    s.addShape("line", { x: diagX, y: axisY, w: diagW, h: 0, line: { color: INK, width: 1.25 } });
    const bars = [
      { label: "普通保証", amount: "2億円", h: 2.9, fill: RED, textColor: "FFFFFF", amountColor: RED },
      { label: "無担保保証", amount: "8,000万円", h: 1.55, fill: GHOST, textColor: INK, amountColor: INK },
      { label: "無担保無保証人\n（小口零細）", amount: "", h: 0.85, fill: GHOST, textColor: INK, amountColor: INK },
    ];
    const barW = 0.95, gap = (diagW - bars.length * barW) / (bars.length + 1);
    let bx = diagX + gap;
    bars.forEach((b) => {
      const by = axisY - b.h;
      s.addShape("rect", {
        x: bx, y: by, w: barW, h: b.h,
        fill: { color: b.fill }, line: b.fill === GHOST ? { color: INK, width: 1 } : { type: "none" },
      });
      if (b.amount) {
        s.addText(b.amount, {
          x: bx - 0.15, y: by - 0.32, w: barW + 0.3, h: 0.28,
          fontFace: F_BODY, fontSize: 9.5, bold: true, color: b.amountColor,
          align: "center", isTextBox: true, margin: 0,
        });
      }
      s.addText(b.label, {
        x: bx - 0.1, y: axisY + 0.06, w: barW + 0.2, h: 0.5,
        fontFace: F_BODY, fontSize: 7.5, color: INK,
        align: "center", isTextBox: true, margin: 0, lineSpacingMultiple: 1.1,
      });
      bx += barW + gap;
    });
    s.addText("一般保証の限度額イメージ（担保不要な枠ほど小さい）", {
      x: diagX, y: axisY + 0.62, w: diagW, h: 0.3,
      fontFace: F_BODY, fontSize: 8, color: INK_SOFT, align: "center", isTextBox: true, margin: 0,
    });
  }

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：G-9 金融・信用保証、税制（信用保証制度の基礎）。",
    years: mkYears(new Set(["'16", "'18", "'20", "'21", "'23", "'24"])),
  });
}

// ---------- Slide 38: G-24 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-24 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第26問設問1）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "中小企業経営・政策",
  });
  const cy = addExamQuestion(s, {
    stem: "信用保証制度において、一般保証の限度額は、普通保証が【Ａ】、【Ｂ】が8,000万円以内となっている。空欄ＡとＢに入る語句の組み合わせとして、最も適切なものはどれか。",
    stemH: 0.65,
    choices: [
      { badge: "ア", text: "Ａ：１億円以内　Ｂ：無担保保証", h: 0.42 },
      { badge: "イ", text: "Ａ：１億円以内　Ｂ：無担保無保証人保証", h: 0.42 },
      { badge: "ウ", text: "Ａ：２億円以内　Ｂ：無担保保証", h: 0.42 },
      { badge: "エ", text: "Ａ：２億円以内　Ｂ：無担保無保証人保証", h: 0.42 },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/G1JI2025.pdf（令和7年度第1次試験）第26問設問1", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 39: G-24 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-24 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第26問設問1）",
    overview: "正解はウ。普通保証2億円以内、無担保保証の名称も正しい。",
    tag: "中小企業経営・政策",
  });
  const choices = [
    { badge: "ア", text: "Ａ：１億円以内　Ｂ：無担保保証", h: 0.42 },
    { badge: "イ", text: "Ａ：１億円以内　Ｂ：無担保無保証人保証", h: 0.42 },
    { badge: "ウ", text: "Ａ：２億円以内　Ｂ：無担保保証", h: 0.42 },
    { badge: "エ", text: "Ａ：２億円以内　Ｂ：無担保無保証人保証", h: 0.42 },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 2 });
  cy += 0.06;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "正解：ウ", options: { bold: true, color: RED } },
    { text: "。信用保証協会の一般保証は「普通保証」（限度額", options: {} },
    { text: "2億円以内", options: { bold: true } },
    { text: "）と「無担保保証」（限度額8,000万円以内、担保不要）の2種類で、両者は別枠でカウントされる。ア・イ：Ａの限度額「1億円以内」が誤り（正しくは2億円以内）。イ・エ：Ｂの「無担保無保証人保証」は、より小規模向けの別の保証類型（小口零細企業保証制度）を指し、本問で対比すべき「無担保保証」とは異なる。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.9, fontFace: F_BODY, fontSize: 9, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.12 });
  cy += 0.95;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "取引先の倒産や自然災害等で経営が悪化した場合は、一般保証とは別枠のセーフティネット保証を利用できる。", options: { fontFace: F_BODY, fontSize: 9, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.4, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  cy += 0.45;
  s.addText("関連知識：普通保証（2億円）・無担保保証（8,000万円）・無担保無保証人保証（小口零細企業保証、より小規模向け）の3類型を限度額とセットで区別すること。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.35;
  s.addText("出典：past_exams/1st_stage/1ji2025/G1JI2025.pdf（令和7年度第1次試験）第26問設問1／正解：past_exams/1st_stage_answers/r07/2025g.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 40: G-25 財務サポート ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-25 ／ 財務サポート（税制、事業承継税制）",
    title: "特例措置は対象株式数の上限なし、猶予割合100％",
    overview: "一般措置との違い（上限有無・猶予割合）が頻出の比較ポイント。",
    tag: "中小企業経営・政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "事業承継税制", v: "非上場株式の贈与・相続税を都道府県知事の認定で猶予・最終免除" },
    { k: "一般措置", v: "対象株式は総株式数の3分の2まで。猶予割合80％" },
    { k: "特例措置", v: "対象株式の上限撤廃（全株式）。猶予割合100％。特例承継計画の事前提出が条件" },
  ], { fontSize: 10, labelW: 1.8, gap: 0.4 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "特例措置は「特例承継計画」の事前提出という手続き要件がある点を見落としやすい。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：G-14 事業承継・M&A（承継の3方式との関連）。",
    years: mkYears(new Set(["'17", "'18", "'20", "'21", "'22", "'23", "'25"])),
  });
}

// ---------- Slide 41: G-26 商業・地域サポート ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-26 ／ 商業・地域サポート",
    title: "まちづくり三法は目的の異なる3つの法律",
    overview: "郊外大型店の出店増加で衰退した中心市街地の再生を図る。",
    tag: "中小企業経営・政策",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "大規模小売店舗立地法", desc: "大型店の立地環境（周辺の生活環境）を整備" },
    { name: "中心市街地活性化法", desc: "郊外化で衰退した中心市街地の再生を図る" },
    { name: "改正都市計画法", desc: "広域的な土地利用規制" },
  ], { nameW: 2.8, tagW: 0, rowH: 0.9 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：3法の目的（立地環境整備／中心市街地再生／広域土地利用規制）を取り違えない。",
    years: mkYears(new Set(["'16", "'18", "'19", "'20", "'23", "'24"])),
  });
}

// ---------- Slide 42: G-27 中小企業支援事業の実施体制 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-27 ／ 中小企業支援事業の実施体制（中小機構、支援センター等）",
    title: "商工会と商工会議所は根拠法・所管が異なる別組織",
    overview: "よろず支援拠点と中小企業支援センターも設置根拠が異なる。",
    tag: "中小企業経営・政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "中小企業基盤整備機構", v: "中小企業政策実施の中核的独立行政法人。ファンド出資・共済制度運営等" },
    { k: "よろず支援拠点", v: "中小機構が委託。あらゆる経営相談に無料対応するワンストップ窓口" },
    { k: "商工会／商工会議所", v: "商工会＝町村部・中小企業庁所管／商工会議所＝市部が中心" },
  ], { fontSize: 10, labelW: 2.4, gap: 0.4 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「商工会」と「商工会議所」は名前が似ているが根拠法・管轄・所管が異なる別組織。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：G-18 中小企業関連法規（中小企業支援法との関係）。",
    years: mkYears(new Set(["'21", "'22", "'23", "'25"])),
  });
}

// ---------- Slide 43: G-28 中小企業政策の役割と変遷 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "G-28 ／ 中小企業政策の役割と変遷（沿革）",
    title: "「格差是正型」から「自立支援型」へ転換した",
    overview: "1999年の抜本改正で中小企業の位置づけが大きく変わった。",
    tag: "中小企業経営・政策",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "1963年 制定", tag: "格差是正型", desc: "大企業と中小企業の「二重構造」是正が目的。弱者保護の発想" },
    { name: "1999年 抜本改正", tag: "自立支援型", desc: "「多様な事業活動を行い経済の基盤を形成する存在」へ位置づけ直し" },
  ], { nameW: 2.2, tagW: 1.8, rowH: 1.4 });
  addFreqBar(s, {
    y: 6.55, rank: "C", rankLabel: "出題実績あり",
    related: "ひっかけ：年号そのものより「制定時＝弱者保護、改正後＝自立支援」という理念転換の方向性を理解することが重要。",
    years: mkYears(new Set(["'21"])),
  });
}

pres.writeFile({ fileName: path.join(__dirname, "../../slides/1st_stage/G_sme_management_policy.pptx") })
  .then(() => console.log("wrote", "slides/1st_stage/G_sme_management_policy.pptx"));
