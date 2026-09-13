// Builds slides/1st_stage/E_business_law.pptx from lib.js components.
// Run: node scripts/pptx_slidegen/build_E_business_law.js
//
// Mirrors slides/1st_stage/E_business_law.html. See docs/13_pptx_slide_template_spec.md.

const path = require("path");
const {
  newPres, addCoverSlide, addDividerSlide, addHeader, addFreqBar, mkYears,
  addTermRows, addRowList, addExamQuestion, drawBarChart,
  INK, INK_SOFT, RED, LINE, GHOST, F_HEAD, F_BODY, F_MONO,
} = require("./lib");

const pres = newPres();
pres.author = "SMEC";
pres.title = "経営法務";

// ---------- Slide 01: Cover ----------
addCoverSlide(pres, {
  eyebrow: "中小企業診断士 第1次試験",
  subjectNo: "SUBJECT 05 / 07",
  title: "経営法務",
  subtitle: "会社の設立・倒産、知的財産権、取引・民法・会社法、独禁法、資本市場まで27論点。丸暗記ではなく「似た制度どうしの比較」がカギになる科目。",
  stats: [["27", "収録論点数"], ["5", "分野"], ["18/27", "頻出ランクA"]],
  tag: "docs/textbook/E_business_law_textbook.md",
  notes: "表紙スライド。経営法務の全体像を示す。",
});

// ---------- Slide 02: 事業開始・倒産／知的財産権 区切り ----------
addDividerSlide(pres, {
  ghostNo: "01",
  partNo: "PART 01",
  partLabel: "経営法務 ／ E-1〜E-9",
  title: "事業開始・倒産／知的財産権",
  desc: "会社を始めるとき・たたむときの基本ルールから、特許・商標・意匠・実用新案・著作権という知的財産権の最頻出分野まで。",
  chips: ["E-4 事業譲渡と会社分割", "E-6 特許権", "E-7 商標権"],
  notes: "事業開始・倒産／知的財産権パートの区切りスライド。",
});

// ---------- Slide 03: E-1 個人の事業開始 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-1 ／ 個人の事業開始（個人事業の特徴、開業準備）",
    title: "無限責任か有限責任かが最大の違い",
    overview: "個人事業は開業が簡単な代わりに、失敗時のリスクが大きい。",
    tag: "経営法務",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "設立手続", desc: "個人事業：開業届のみ／法人：定款作成・認証・設立登記が必要" },
    { name: "責任の範囲", desc: "個人事業：無限責任（個人財産で返済義務）／法人：出資額の範囲の有限責任" },
    { name: "税金", desc: "個人事業：所得税（累進課税）／法人：法人税" },
  ], { nameW: 2.0, tagW: 0, rowH: 1.1 });
  addFreqBar(s, {
    y: 6.55, rank: "－", rankLabel: "直近10年単独出題なし",
    related: "ひっかけ：「登記が不要で簡単」と「無限責任か有限責任か」は別の話。簡単に始められる代わりにリスクが大きい表裏の関係。",
    years: mkYears(new Set([])),
  });
}

// ---------- Slide 04: E-2 法人の事業開始 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-2 ／ 法人の事業開始（会社の設立と登記）",
    title: "発行可能株式総数は絶対的記載事項ではない",
    overview: "定款作成→認証→出資履行→設立時取締役選任→設立登記の順で会社が成立する。",
    tag: "経営法務",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "絶対的記載事項", v: "目的・商号・本店所在地・出資財産の価額（最低額）・発起人の氏名住所" },
    { k: "定款認証", v: "株式会社は公証人の認証が必要（合同会社は不要）" },
    { k: "設立登記", v: "登記により会社が法的に成立する（成立要件）" },
  ], { fontSize: 11.5, labelW: 2.1, gap: 0.4 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "発行可能株式総数は絶対的記載事項に含まれない。定款に記載がなくても定款は無効にならないが、成立時までに別途定める必要がある。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：E-18 会社法（機関設計・取締役会の運営ルール）。",
    years: mkYears(new Set(["'20", "'22", "'23", "'24"])),
  });
}

// ---------- Slide 05: E-3 届出・手続 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-3 ／ 届出・手続（許認可、労働保険・社会保険、税務上の届出）",
    title: "業種によって必要な許認可が異なる",
    overview: "会社設立だけでは事業を始められない業種もある。",
    tag: "経営法務",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "許認可", v: "飲食業（保健所）、建設業（建設業許可）、古物商（公安委員会）等" },
    { k: "労働・社会保険", v: "従業員雇用時は労災・雇用保険、健康保険・厚生年金への加入手続" },
  ], { fontSize: 11.5, labelW: 2.3, gap: 0.45 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「許可」「認可」「届出」「登録」で行政手続の性質が異なる。個々の業種の詳細より「事前に確認すべき事項が複数ある」という全体像を理解しておけば十分。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "－", rankLabel: "直近10年単独出題なし",
    related: "関連：E-2 法人の事業開始（会社設立から営業開始までの流れ）。",
    years: mkYears(new Set([])),
  });
}

// ---------- Slide 06: E-4 合併等の手続 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-4 ／ 合併等の手続（合併・営業譲渡、組織変更）",
    title: "「バラ売り」か「まるごと引っ越し」か",
    overview: "経営法務で最頻出。包括承継か特定承継かで同意・保護手続の要否が変わる。",
    tag: "経営法務",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "事業譲渡", tag: "特定承継", desc: "個々の契約を個別移転。契約相手方の個別同意が必要。債権者保護手続の規定なし" },
    { name: "会社分割", tag: "包括承継", desc: "契約上の地位も一括移転。相手方の同意は不要だが原則債権者保護手続が必要" },
    { name: "合併", tag: "包括承継", desc: "消滅会社の権利義務全部が移転。同意不要、原則債権者保護手続が必要" },
  ], { nameW: 1.8, tagW: 1.6, rowH: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：「事業譲渡も包括承継」「会社分割にも保護手続不要」は誤り。3点（承継の性質・同意要否・保護手続要否）をセットで覚える。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 07: E-4 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-4 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第6問設問1）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "経営法務",
  });
  const cy = addExamQuestion(s, {
    stem: "X社代表取締役甲氏と中小企業診断士の会話：「A事業部門をY社に譲渡する際、事業譲渡と会社分割とで契約関係の移転の効力に違いはあるか（空欄A）」「債権者保護手続に違いはあるか。会社分割で免責的債務引受をしてもらう予定（空欄B）」。空欄ＡとＢに入る記述の組み合わせとして、最も適切なものはどれか。",
    stemH: 1.1,
    choices: [
      { badge: "ア", text: "A：会社分割は包括承継で契約相手方の同意不要、事業譲渡は個別に同意が必要／B：会社分割は債権者保護手続が必要、事業譲渡は不要（ただし免責的債務引受には個別に債権者の同意が必要）", h: 0.72 },
      { badge: "イ", text: "A：（アと同じ）／B：会社分割も事業譲渡も債権者保護手続が必要", h: 0.5 },
      { badge: "ウ", text: "A：会社分割も事業譲渡も包括承継で同意不要／B：会社分割も事業譲渡も債権者保護手続が必要", h: 0.5 },
      { badge: "エ", text: "A：（ウと同じ）／B：会社分割も事業譲渡も債権者保護手続は不要", h: 0.5 },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/E1JI2025.pdf（令和7年度第1次試験）第6問設問1", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 08: E-4 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-4 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第6問設問1）",
    overview: "正解はア。事業譲渡は特定承継で同意必要、保護手続の規定はなし。",
    tag: "経営法務",
  });
  const choices = [
    { badge: "ア", text: "A：会社分割は包括承継で契約相手方の同意不要、事業譲渡は個別に同意が必要／B：会社分割は債権者保護手続が必要、事業譲渡は不要（ただし免責的債務引受には個別に債権者の同意が必要）", h: 0.72 },
    { badge: "イ", text: "A：（アと同じ）／B：会社分割も事業譲渡も債権者保護手続が必要", h: 0.5 },
    { badge: "ウ", text: "A：会社分割も事業譲渡も包括承継で同意不要／B：会社分割も事業譲渡も債権者保護手続が必要", h: 0.5 },
    { badge: "エ", text: "A：（ウと同じ）／B：会社分割も事業譲渡も債権者保護手続は不要", h: 0.5 },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 0 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：ア", options: { bold: true, color: RED } },
    { text: "。会社分割は分割契約・計画に定めた権利義務が包括的に承継され契約相手方の個別同意は不要（会社法759条等）。事業譲渡は個々の財産・契約ごとの特定承継で契約相手方の同意が原則必要。会社分割は原則債権者保護手続（官報公告・催告）が必要、事業譲渡には会社法上の規定はないが、免責的債務引受には民法上、債権者の個別承諾が必要（民法472条3項）。イ：事業譲渡にも保護手続が必要としている点が誤り。ウ・エ：事業譲渡を包括承継としている点が誤り。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.85, fontFace: F_BODY, fontSize: 9.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  cy += 0.92;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "事業譲渡は個々の契約を一つひとつ相手方の同意を得て切り替える必要があり手間がかかるが、譲渡対象を個別に選別できる自由度がある。会社分割は包括的に移転する代わり、債権者保護のため保護手続を経る必要がある。", options: { fontFace: F_BODY, fontSize: 9.5, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.75;
  s.addText("関連知識：事業譲渡と会社分割の比較は経営法務の最頻出論点。「包括承継か特定承継か」「契約相手方の同意の要否」「債権者保護手続の要否」の3点をセットで押さえること。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 9.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.35;
  s.addText("出典：past_exams/1st_stage/1ji2025/E1JI2025.pdf（令和7年度第1次試験）第6問設問1／正解：past_exams/1st_stage_answers/r07/2025e.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 09: E-5 倒産等の手続 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-5 ／ 倒産等の手続",
    title: "会社更生法は株式会社のみ、民事再生法は誰でも",
    overview: "再建型（民事再生・会社更生）と清算型（破産・解散清算）に大別される。",
    tag: "経営法務",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "民事再生法", tag: "再建型", desc: "株式会社に限らず利用可。現経営陣が続投し再建（DIP型）" },
    { name: "会社更生法", tag: "再建型・株式会社限定", desc: "更生管財人が経営権を握り現経営陣は原則退任。大企業向け" },
    { name: "破産法", tag: "清算型", desc: "破産管財人が財産を換価し配当後、会社を消滅させる" },
  ], { nameW: 1.8, tagW: 2.3, rowH: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：対象（株式会社限定か否か）と経営権の帰趨（続投か退任か）の違いが頻出。",
    years: mkYears(new Set(["'16", "'19", "'21", "'23"])),
  });
}

// ---------- Slide 10: E-6 特許権 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-6 ／ 産業財産権の内容と取得方法（特許権）",
    title: "職務発明の通常実施権は登録不要",
    overview: "出願から20年。職務発明では使用者等が無償の通常実施権を当然取得する。",
    tag: "経営法務",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "存続期間", v: "出願日から20年（延長制度あり）" },
    { k: "新規性喪失の例外", v: "公開日から1年以内に手続を伴い出願すれば新規性を失わない扱い" },
    { k: "職務発明", v: "特許を受ける権利は原則発明者本人。使用者等は無償の通常実施権を当然取得" },
  ], { fontSize: 11.5, labelW: 2.1, gap: 0.4 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "使用者等の通常実施権は登録不要で法律上当然に発生し第三者にも対抗できる。「過去の職務」に属する発明も職務発明に含まれる。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：E-11 知的財産権に関する契約（専用実施権・通常実施権の違い）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 11: E-6 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-6 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第9問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "経営法務",
  });
  const cy = addExamQuestion(s, {
    stem: "特許法第35条に規定する職務発明に関する記述として、最も適切なものはどれか。",
    stemH: 0.5,
    choices: [
      { badge: "ア", text: "従業者等がした発明については、その発明が職務発明である場合を除き、あらかじめ使用者等に特許を受ける権利を取得させることを定めた契約・勤務規則等の条項は、無効である。", h: 0.65 },
      { badge: "イ", text: "使用者等が有する通常実施権は、その発生後に特許権を取得した者に対しては、その効力を有しない。", h: 0.55 },
      { badge: "ウ", text: "使用者等が有する通常実施権は、登録しなければ、その効力を生じない。", h: 0.5 },
      { badge: "エ", text: "職務発明は現在の職務に属する発明を指すので、同一企業内であっても従業者等の過去の職務に属する発明は、職務発明と認められる場合はない。", h: 0.6 },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/E1JI2025.pdf（令和7年度第1次試験）第9問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 12: E-6 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-6 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第9問）",
    overview: "正解はア。条文どおり、自由発明への権利取得条項は無効。",
    tag: "経営法務",
  });
  const choices = [
    { badge: "ア", text: "従業者等がした発明については、その発明が職務発明である場合を除き、あらかじめ使用者等に特許を受ける権利を取得させることを定めた契約・勤務規則等の条項は、無効である。", h: 0.65 },
    { badge: "イ", text: "使用者等が有する通常実施権は、その発生後に特許権を取得した者に対しては、その効力を有しない。", h: 0.55 },
    { badge: "ウ", text: "使用者等が有する通常実施権は、登録しなければ、その効力を生じない。", h: 0.5 },
    { badge: "エ", text: "職務発明は現在の職務に属する発明を指すので、同一企業内であっても従業者等の過去の職務に属する発明は、職務発明と認められる場合はない。", h: 0.6 },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 0 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：ア", options: { bold: true, color: RED } },
    { text: "。特許法35条2項どおりの記述で正しい（職務発明でない発明＝自由発明について会社に権利取得させる取り決めは無効）。イ・ウ：使用者等の法定通常実施権（35条1項）は", options: {} },
    { text: "登録不要", options: { bold: true } },
    { text: "で法律上当然に発生し、後に特許権を取得した第三者にも対抗できる。エ：条文は「現在", options: {} },
    { text: "又は過去", options: { bold: true } },
    { text: "の職務」と規定しており、過去の職務に属する発明も職務発明に含まれる。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.75, fontFace: F_BODY, fontSize: 9.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.18 });
  cy += 0.82;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "職務発明制度は①特許を受ける権利は原則発明者本人、②使用者等は無償の法定通常実施権を当然取得、③契約・勤務規則等であらかじめ使用者等に権利取得させることも可能、という3点構造。", options: { fontFace: F_BODY, fontSize: 9.5, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.55;
  s.addText("関連知識：存続期間等の基礎知識はdocs/07_key_formulas_and_frameworks.md参照。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 9.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.35;
  s.addText("出典：past_exams/1st_stage/1ji2025/E1JI2025.pdf（令和7年度第1次試験）第9問／正解：past_exams/1st_stage_answers/r07/2025e.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 13: E-7 商標権 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-7 ／ 産業財産権の内容と取得方法（商標権）",
    title: "更新すれば半永久的に権利を維持できる",
    overview: "登録日から10年、更新可。不使用が3年続くと誰でも取消審判を請求できる。",
    tag: "経営法務",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "存続期間", v: "登録日から10年（更新登録で何度でも延長可能）" },
    { k: "先使用権", v: "登録前から使用し周知になっていれば使い続けられる" },
    { k: "不使用取消審判", v: "継続3年以上不使用なら「何人も」取消しを請求できる" },
  ], { fontSize: 11.5, labelW: 2.1, gap: 0.4 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "不使用取消審判の期間は「3年」（2年ではない）。請求権者は「利害関係人のみ」ではなく「何人も」。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：E-8 意匠権・実用新案権（存続期間の起算点・年数の横断比較）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 14: E-7 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-7 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第16問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "経営法務",
  });
  const cy = addExamQuestion(s, {
    stem: "X社代表取締役甲氏と中小企業診断士の会話。Y社が同一文字商標を先に登録しており、甲氏は自社の商標登録が難しいか相談。Y社はその商標を使用していない。「商標法第50条第1項の不使用取消審判は、継続して国内で【Ａ】年以上、登録商標の使用をしていない場合に請求できる」「不使用取消審判は【Ｂ】」。空欄ＡとＢに入る数値と記述の組み合わせとして、最も適切なものはどれか。",
    stemH: 1.0,
    choices: [
      { badge: "ア", text: "Ａ：2　Ｂ：誰でも請求することができます", h: 0.5 },
      { badge: "イ", text: "Ａ：2　Ｂ：利害関係人のみが請求することができます", h: 0.5 },
      { badge: "ウ", text: "Ａ：3　Ｂ：誰でも請求することができます", h: 0.5 },
      { badge: "エ", text: "Ａ：3　Ｂ：利害関係人のみが請求することができます", h: 0.5 },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/E1JI2025.pdf（令和7年度第1次試験）第16問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 15: E-7 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-7 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第16問）",
    overview: "正解はウ（3年、誰でも請求可能）。",
    tag: "経営法務",
  });
  const choices = [
    { badge: "ア", text: "Ａ：2　Ｂ：誰でも請求することができます", h: 0.5 },
    { badge: "イ", text: "Ａ：2　Ｂ：利害関係人のみが請求することができます", h: 0.5 },
    { badge: "ウ", text: "Ａ：3　Ｂ：誰でも請求することができます", h: 0.5 },
    { badge: "エ", text: "Ａ：3　Ｂ：利害関係人のみが請求することができます", h: 0.5 },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 2 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：ウ", options: { bold: true, color: RED } },
    { text: "。商標法50条1項は「継続して", options: {} },
    { text: "3年", options: { bold: true } },
    { text: "以上使用していないときは、", options: {} },
    { text: "何人も", options: { bold: true } },
    { text: "取消審判を請求できる」と規定する。ア・イ：期間を「2年」としている点が誤り。イ・エ：請求権者を「利害関係人のみ」としている点が誤り（無効審判など利害関係人限定の審判類型と対比される）。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.8, fontFace: F_BODY, fontSize: 9.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.18 });
  cy += 0.85;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "商標権の存続期間は登録から10年（更新可・半永久的）だが、不使用のまま3年が経過すると第三者から取消審判を請求されうる。商標は「使用してこそ価値がある」という思想（信用の化体）に基づく制度。", options: { fontFace: F_BODY, fontSize: 9.5, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.75;
  s.addText("関連知識：特許の新規性喪失例外の「1年」や他の年数規定と混同しないこと。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 9.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.35;
  s.addText("出典：past_exams/1st_stage/1ji2025/E1JI2025.pdf（令和7年度第1次試験）第16問／正解：past_exams/1st_stage_answers/r07/2025e.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 16: E-8 意匠権・実用新案権（存続期間の横断比較, 右カラム図解） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-8 ／ 産業財産権の内容と取得方法（意匠権・実用新案権）",
    title: "存続期間の起算点と年数を横断整理する",
    overview: "特許・実用新案・意匠・商標で起算点も年数も異なる。混同が最頻出のひっかけ。",
    tag: "経営法務",
  });
  const proseX = 0.55, proseW = 7.4;
  const diagX = 8.25, diagW = 4.05;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "意匠権", v: "出願日から25年（2020年4月以降の出願）", gap: 0.42 },
    { k: "実用新案権", v: "出願日から10年。無審査主義（実体審査なし）", gap: 0.42 },
    { k: "秘密意匠", v: "請求により登録日から最長3年間、公開を遅らせられる", gap: 0.42 },
  ], { fontSize: 11, labelW: 1.7 });
  cy += 0.1;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「意匠権は登録から20年」は2020年改正前の知識で誤り。実用新案は無審査のため権利行使時に技術評価書の提示が必要。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.7, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  drawBarChart(s, diagX, 1.95, diagW, 3.5, {
    bars: [
      { label: "特許", value: 20, on: false },
      { label: "実用新案", value: 10, on: false },
      { label: "意匠", value: 25, on: true },
      { label: "商標", value: 10, on: false },
    ],
    maxValue: 25,
  });
  s.addText("存続期間の比較（いずれも出願/登録日起算）", {
    x: diagX, y: 5.88, w: diagW, h: 0.28, align: "center",
    fontFace: F_BODY, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "具体例：秘密意匠は「先に権利は確保しつつ公開時期を遅らせたい」ニーズに対応。関連意匠制度でバリエーションデザインをまとめて登録可能。",
    years: mkYears(new Set(["'16", "'17", "'18", "'20", "'21", "'22", "'23", "'24"])),
  });
}

// ---------- Slide 17: E-9 著作権の内容 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-9 ／ 著作権の内容",
    title: "職務著作では会社自身が著作者になる",
    overview: "創作と同時に自動発生（無方式主義）。特許の職務発明とは扱いが対照的。",
    tag: "経営法務",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "著作権（財産権）", v: "複製権・上演権等。他人に譲渡できる" },
    { k: "著作者人格権", v: "氏名表示権・同一性保持権等。一身専属で譲渡できない" },
    { k: "職務著作", v: "会社の発意・従業員が職務上作成・会社名義公表なら会社が著作者" },
  ], { fontSize: 11.5, labelW: 2.3, gap: 0.4 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "特許の職務発明は発明者本人が原則権利者だが、著作権の職務著作は要件を満たせば会社自身が著作者になる点が対照的。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：E-10 著作権の成立と保護（保護期間の詳細）。",
    years: mkYears(new Set(["'17", "'18", "'20", "'22", "'25"])),
  });
}

pres.writeFile({ fileName: path.join(__dirname, "../../slides/1st_stage/E_business_law.pptx") })
  .then(() => console.log("wrote", "slides/1st_stage/E_business_law.pptx"));
