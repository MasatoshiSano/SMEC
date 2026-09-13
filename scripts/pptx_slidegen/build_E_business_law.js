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

// ---------- Slide 18: E-10〜E-19 区切り ----------
addDividerSlide(pres, {
  ghostNo: "02",
  partNo: "PART 02",
  partLabel: "経営法務 ／ E-10〜E-19",
  title: "取引法務・民法・会社法",
  desc: "知財の契約・保護期間から、英文契約や紛争解決の実務知識、そして経営法務で最頻出の民法・会社法まで。",
  chips: ["E-15 英文契約", "E-17 民法（遺留分）", "E-18 会社法（取締役会）"],
  notes: "知的財産権（契約）／取引関係法務／民法・会社法パートの区切りスライド。",
});

// ---------- Slide 19: E-10 著作権の成立と保護 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-10 ／ 著作権の成立と保護（保護期間、侵害への措置）",
    title: "起算点は「死後」か「公表後」かで変わる",
    overview: "2018年改正で死後50年から70年に延長。無名・法人名義・映画は公表後70年。",
    tag: "経営法務",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "原則", v: "著作者の死後70年" },
    { k: "無名・変名・法人名義", v: "公表後70年" },
    { k: "映画の著作物", v: "公表後70年" },
    { k: "権利侵害への措置", v: "差止請求・損害賠償請求・名誉回復措置の請求ができる" },
  ], { fontSize: 11, labelW: 2.3, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「すべて死後70年」と単純に覚えると誤る。無名・変名・法人名義・映画は公表後70年が基準。2018年改正前（死後50年時代）の過去問は当時の年数で判断。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "B", rankLabel: "頻出論点",
    related: "関連：E-9 著作権の内容（著作者人格権・職務著作の基礎）。",
    years: mkYears(new Set(["'17", "'19"])),
  });
}

// ---------- Slide 20: E-11 知的財産権に関する契約 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-11 ／ 知的財産権に関する契約（移転契約、ライセンス契約）",
    title: "専用実施権は登録必須、通常実施権は登録不要",
    overview: "著作権譲渡は「翻訳権等」が留保推定される点に注意。",
    tag: "経営法務",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "著作権譲渡契約", desc: "「著作権を譲渡する」とのみ記載すると翻訳権・翻案権等は譲渡人に留保されたと推定される" },
    { name: "専用実施権", tag: "登録が効力発生要件", desc: "独占的な実施権。設定範囲では特許権者自身も実施不可。自ら差止請求も可能" },
    { name: "通常実施権", tag: "登録不要", desc: "非独占的で複数の相手に重ねて許諾可。当事者間の契約のみで効力発生" },
  ], { nameW: 2.0, tagW: 2.2, rowH: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：E-6 特許権（職務発明の法定通常実施権も「登録不要」の原則に沿う）。",
    years: mkYears(new Set(["'17", "'19", "'21", "'25"])),
  });
}

// ---------- Slide 21: E-12 トレードシークレット・国際条約 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-12 ／ トレードシークレット・国際条約",
    title: "営業秘密は3要件をすべて満たして初めて保護される",
    overview: "秘密管理性・有用性・非公知性。1つでも欠けると保護されない。",
    tag: "経営法務",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "秘密管理性", v: "秘密として管理する意思が従業員等にも客観的に認識できる状態" },
    { k: "有用性", v: "事業活動に役立つ技術上・営業上の情報であること" },
    { k: "非公知性", v: "一般に知られておらず容易に知ることができない状態" },
    { k: "パリ条約", v: "内国民待遇・優先権制度（特許/実用新案12か月、意匠/商標6か月）" },
  ], { fontSize: 11, labelW: 2.0, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「業界内で広く知れ渡っているノウハウ」は非公知性を欠き、「管理されていない社内資料」は秘密管理性を欠くため、いずれも営業秘密に該当しない。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：PCT（特許協力条約）＝1国際出願で複数加盟国に同時出願と同様の効果。",
    years: mkYears(new Set(["'16", "'20", "'21", "'24", "'25"])),
  });
}

// ---------- Slide 22: E-13 契約に関する基礎知識 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-13 ／ 契約に関する基礎知識（成立要件、有効要件）",
    title: "契約は口頭の合意だけでも成立する",
    overview: "申込みと承諾の意思表示が合致した時点で成立。書面は原則不要。",
    tag: "経営法務",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "契約の成立", v: "「申込み」と「承諾」の意思表示が合致した時点。書面は原則不要" },
    { k: "例外", v: "保証契約は書面（電磁的記録を含む）を要件とする" },
    { k: "定型約款", v: "画一的な取引条件。表示していれば個別の条項未認識でも合意とみなす。不当条項は合意しなかったものとみなす" },
  ], { fontSize: 11, labelW: 1.9, gap: 0.42 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「契約は書面がなければ成立しない」は誤り。口頭でも成立するが、保証契約など例外的に書面が必要な契約類型もある。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "B", rankLabel: "頻出論点",
    related: "具体例：ネット通販の利用規約は定型約款に該当することが多く、個々の条項を読んでいなくても原則拘束力を持つ。",
    years: mkYears(new Set(["'16", "'20"])),
  });
}

// ---------- Slide 23: E-14 契約の類型 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-14 ／ 契約の類型（秘密保持、共同開発、売買、フランチャイズ）",
    title: "システム開発契約は「請負」か「準委任」かが焦点",
    overview: "上流工程は準委任、仕様確定後の開発・実装は請負が採用されやすい。",
    tag: "経営法務",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "秘密保持契約", tag: "NDA", desc: "M&A検討・共同開発の前段階でほぼ必ず締結。目的外使用・第三者開示を禁止" },
    { name: "請負契約", desc: "成果物の完成を約束。完成義務を負う（システム開発の仕様確定後の工程）" },
    { name: "準委任契約", desc: "業務の遂行を約束。完成義務までは負わない（要件定義等の上流工程）" },
    { name: "フランチャイズ契約", desc: "商標・ノウハウ使用の対価にロイヤリティ。独禁法（優越的地位の濫用）とも関連" },
  ], { nameW: 2.2, tagW: 1.2, rowH: 0.9 });
  addFreqBar(s, {
    y: 6.55, rank: "B", rankLabel: "頻出論点",
    related: "ひっかけ：「準委任契約なのに完成義務を負う」といった誤った組み合わせに注意。",
    years: mkYears(new Set(["'16", "'18"])),
  });
}

// ---------- Slide 24: E-15 外国企業との取引・英文契約 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-15 ／ 外国企業との取引・英文契約",
    title: "仲裁は非公開・一審制、訴訟は公開・審級あり",
    overview: "英文契約特有の条項の趣旨を問う出題が最頻出。",
    tag: "経営法務",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "準拠法条項", v: "契約の解釈・履行にどの国の法律を適用するか" },
    { k: "仲裁条項", v: "非公開・原則一審制。仲裁判断は確定判決と同様の効力" },
    { k: "完全合意条項", v: "契約書の内容が完全な合意。それ以前の合意は効力を失う" },
    { k: "不可抗力条項", v: "天災地変等、当事者の責によらない事由での免責" },
  ], { fontSize: 11, labelW: 2.0, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "仲裁＝非公開・一審制、訴訟＝原則公開・控訴上告ありという対比を押さえておくこと。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：E-16 紛争解決方法（仲裁・訴訟・調停の比較）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 25: E-16 紛争解決方法 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-16 ／ 紛争解決方法（訴訟、仲裁、調停）",
    title: "少額訴訟は60万円以下、即日判決が原則",
    overview: "「公開・非公開」「審級の有無」「合意か判断か」で4手続を整理する。",
    tag: "経営法務",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "訴訟", desc: "裁判所が判決で強制的に解決。原則公開、控訴・上告の審級制度あり" },
    { name: "少額訴訟", tag: "60万円以下", desc: "簡易裁判所。原則1回の期日で審理、即日判決。同一裁判所で年10回まで" },
    { name: "民事調停", desc: "話し合いによる合意で解決。調停調書は確定判決と同様の効力" },
    { name: "仲裁", desc: "仲裁人の判断で解決するADR。非公開・原則一審制" },
  ], { nameW: 1.7, tagW: 1.6, rowH: 0.9 });
  addFreqBar(s, {
    y: 6.55, rank: "C", rankLabel: "出題実績あり",
    related: "ひっかけ：少額訴訟の対象額「60万円以下」を正確に覚える。仲裁・調停・訴訟の性質の違いに注意。",
    years: mkYears(new Set(["'25"])),
  });
}

// ---------- Slide 26: E-17 民法 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-17 ／ 民法（物権、債権、相続）",
    title: "兄弟姉妹には遺留分が認められない",
    overview: "経営法務で最頻出の分野。消滅時効・保証・契約不適合責任・共有・遺留分。",
    tag: "経営法務",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "消滅時効", v: "主観的起算点から5年、客観的起算点から10年のいずれか早い方" },
    { k: "個人根保証", v: "極度額を書面で定めないと無効。事業性融資は公正証書による保証意思確認が原則必要" },
    { k: "共有", v: "保存行為＝単独可／管理行為＝持分の過半数／変更行為＝全員の同意" },
    { k: "遺留分", v: "配偶者・子・直系尊属に認められる。兄弟姉妹には認められない" },
  ], { fontSize: 10.5, labelW: 1.9, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "遺留分侵害額請求権の時効は「相続開始を知った時から1年」（相続開始の時からではない）。相続開始前の放棄のみ家庭裁判所の許可が必要。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：E-23 事業承継（遺留分に関する民法特例＝除外合意・固定合意）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 27: E-17 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-17 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第22問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "経営法務",
  });
  const cy = addExamQuestion(s, {
    stem: "遺留分に関する記述として、最も適切なものはどれか。",
    stemH: 0.5,
    choices: [
      { badge: "ア", text: "遺留分侵害額請求権は、裁判外で行使することも可能である。", h: 0.5 },
      { badge: "イ", text: "遺留分侵害額請求権は、相続開始の時から1年以内に行使しなければ、時効によって消滅する。", h: 0.55 },
      { badge: "ウ", text: "相続の開始後における遺留分の放棄は、家庭裁判所の許可を受けたときに限り、その効力を生ずる。", h: 0.55 },
      { badge: "エ", text: "被相続人の兄弟姉妹は、被相続人の相続について遺留分を有する。", h: 0.5 },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/E1JI2025.pdf（令和7年度第1次試験）第22問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 28: E-17 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-17 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第22問）",
    overview: "正解はア。遺留分侵害額請求権は裁判外の意思表示でも行使できる。",
    tag: "経営法務",
  });
  const choices = [
    { badge: "ア", text: "遺留分侵害額請求権は、裁判外で行使することも可能である。", h: 0.5 },
    { badge: "イ", text: "遺留分侵害額請求権は、相続開始の時から1年以内に行使しなければ、時効によって消滅する。", h: 0.55 },
    { badge: "ウ", text: "相続の開始後における遺留分の放棄は、家庭裁判所の許可を受けたときに限り、その効力を生ずる。", h: 0.55 },
    { badge: "エ", text: "被相続人の兄弟姉妹は、被相続人の相続について遺留分を有する。", h: 0.5 },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 0 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：ア", options: { bold: true, color: RED } },
    { text: "。遺留分侵害額請求権（民法1046条）は権利者の一方的な意思表示で効果が生じる形成権であり、裁判外の意思表示（内容証明郵便等）でも行使できる。イ：1年の起算点は「相続開始の時」ではなく「相続開始", options: {} },
    { text: "及び侵害する贈与・遺贈があったことを知った時", options: { bold: true } },
    { text: "」（民法1048条）。ウ：家庭裁判所の許可が必要なのは「相続開始", options: {} },
    { text: "前", options: { bold: true } },
    { text: "」の放棄であり、開始", options: {} },
    { text: "後", options: { bold: true } },
    { text: "の放棄には許可は不要（許可を要する時期が逆）。エ：兄弟姉妹には遺留分が認められない（民法1042条1項）。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.8, fontFace: F_BODY, fontSize: 9.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  cy += 0.85;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "先代経営者が長男に自社株式のほとんどを生前贈与していた場合、長女は遺留分を侵害されたとして長男に遺留分侵害額請求権を行使できる。内容証明郵便で意思表示をするだけで足り、裁判を起こす必要はない。", options: { fontFace: F_BODY, fontSize: 9.5, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.65;
  s.addText("関連知識：遺留分は配偶者・子（代襲相続人を含む）・直系尊属にのみ認められ、兄弟姉妹には認められない点が頻出。事業承継では経営承継円滑化法上の遺留分に関する民法特例（除外合意・固定合意）とあわせて出題される。", {
    x: 0.55, y: cy, w: 12.25, h: 0.35, fontFace: F_BODY, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15,
  });
  cy += 0.4;
  s.addText("出典：past_exams/1st_stage/1ji2025/E1JI2025.pdf（令和7年度第1次試験）第22問／正解：past_exams/1st_stage_answers/r07/2025e.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 29: E-18 会社法（機関設計） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-18 ／ 会社法（会社の種類、株式、機関設計）",
    title: "取締役会には委任状による代理出席が認められない",
    overview: "株主総会の規律と混同しやすい取締役会の運営ルールが頻出。",
    tag: "経営法務",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "招集通知", v: "会日の1週間前までに発する。定款の定めにより短縮可能" },
    { k: "決議の定足数", v: "過半数の出席・過半数の賛成。特別利害関係取締役は定足数からも除外" },
    { k: "職務執行状況の報告", v: "3か月に1回以上。書面決議による省略の対象外（省略不可）" },
    { k: "議決権の代理行使", v: "取締役会では認められない（株主総会では認められる）" },
  ], { fontSize: 10.5, labelW: 2.2, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "招集通知期間は定款で短縮できる（できないは誤り）。報告義務は省略できない。代理出席は認められない。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：E-2 法人の事業開始（設立時の機関設計の基礎）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 30: E-18 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-18 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第2問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "経営法務",
  });
  const cy = addExamQuestion(s, {
    stem: "取締役会に関する記述として、最も適切なものはどれか。なお、特別取締役の定めはないものとする。",
    stemH: 0.55,
    choices: [
      { badge: "ア", text: "代表取締役および業務執行取締役は、3カ月に1回以上、職務執行の状況を取締役会に報告しなければならないが、定款の定めがあれば、取締役全員への書面による通知により、その報告を省略することができる。", h: 0.65 },
      { badge: "イ", text: "取締役会の決議の定足数は、当該決議について特別の利害関係があり議決に加わることができない取締役の人数を除いて計算することができる。", h: 0.6 },
      { badge: "ウ", text: "取締役会を招集する者は、取締役会の日の1週間前までに、その通知を発しなければならず、この通知期間は、定款の定めによっても短縮することはできない。", h: 0.6 },
      { badge: "エ", text: "取締役が現実に開催される予定の取締役会に出席できない場合、その取締役は、事前に他の取締役に対して委任状を発行することにより、当該取締役会において議決権を行使することができる。", h: 0.65 },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/E1JI2025.pdf（令和7年度第1次試験）第2問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 31: E-18 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-18 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第2問）",
    overview: "正解はイ。特別利害関係取締役は定足数からも除外される。",
    tag: "経営法務",
  });
  const choices = [
    { badge: "ア", text: "代表取締役および業務執行取締役は、3カ月に1回以上、職務執行の状況を取締役会に報告しなければならないが、定款の定めがあれば、取締役全員への書面による通知により、その報告を省略することができる。", h: 0.65 },
    { badge: "イ", text: "取締役会の決議の定足数は、当該決議について特別の利害関係があり議決に加わることができない取締役の人数を除いて計算することができる。", h: 0.6 },
    { badge: "ウ", text: "取締役会を招集する者は、取締役会の日の1週間前までに、その通知を発しなければならず、この通知期間は、定款の定めによっても短縮することはできない。", h: 0.6 },
    { badge: "エ", text: "取締役が現実に開催される予定の取締役会に出席できない場合、その取締役は、事前に他の取締役に対して委任状を発行することにより、当該取締役会において議決権を行使することができる。", h: 0.65 },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 1 });
  cy += 0.06;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "正解：イ", options: { bold: true, color: RED } },
    { text: "。取締役会の決議は議決に加わることができる取締役の過半数の出席と、その過半数の賛成で成立（会社法369条1項）。特別の利害関係を有する取締役は議決に加わることができず（同条2項）、定足数の算定対象からも除外される。ア：職務執行状況の報告義務（363条2項）は決議の省略（370条）とは異なり、372条2項により明文で省略", options: {} },
    { text: "できない", options: { bold: true } },
    { text: "とされる。ウ：招集通知期間（1週間）は定款で", options: {} },
    { text: "短縮できる", options: { bold: true } },
    { text: "（368条1項かっこ書き）。エ：取締役会は個人的信頼関係に基づく合議体で、株主総会と異なり委任状による代理行使は", options: {} },
    { text: "認められない", options: { bold: true } },
    { text: "。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.72, fontFace: F_BODY, fontSize: 9, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.12 });
  cy += 0.78;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "資本金1,000万円の非大会社が取締役会を設置する場合、監査役1名で足りる。ある議案について特別利害関係を持つ取締役が1名いれば、その者を除いた残りの取締役の過半数が出席していれば定足数を満たす。", options: { fontFace: F_BODY, fontSize: 9, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.55, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  cy += 0.6;
  s.addText("関連知識：取締役会の規律（代理行使不可、決議省略の対象範囲）は株主総会の規律（代理行使可）と対比されやすい。両者の違いを意識して整理すること。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.35;
  s.addText("出典：past_exams/1st_stage/1ji2025/E1JI2025.pdf（令和7年度第1次試験）第2問／正解：past_exams/1st_stage_answers/r07/2025e.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 32: E-19 会社法（計算・合併） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "E-19 ／ 会社法（会社の計算、合併、取締役の忠実義務）",
    title: "資本準備金にできるのは払込額の2分の1まで",
    overview: "配当は分配可能額の範囲内。取締役は善管注意義務・忠実義務を負う。",
    tag: "経営法務",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "剰余金の配当", v: "原則株主総会決議。分配可能額を超える配当は不可（財源規制）" },
    { k: "資本金組入れ", v: "払込額の2分の1を超えない額は資本準備金にできる" },
    { k: "取締役の義務", v: "善管注意義務・忠実義務。利益相反取引・競業取引は取締役会の承認が必要" },
    { k: "株価算定", v: "純資産価額方式・類似業種比準方式・DCF方式（インカムアプローチ）" },
  ], { fontSize: 10.5, labelW: 1.9, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「払込金額の全額を資本準備金にできる」は誤り。上限は2分の1を超えない額。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：E-23 事業承継（非上場株式の評価手法は本論点の株価算定と共通）。",
    years: mkYears(new Set(["'17", "'18", "'24"])),
  });
}

pres.writeFile({ fileName: path.join(__dirname, "../../slides/1st_stage/E_business_law.pptx") })
  .then(() => console.log("wrote", "slides/1st_stage/E_business_law.pptx"));
