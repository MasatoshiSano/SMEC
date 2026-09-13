// Builds slides/1st_stage/C_business_administration.pptx from lib.js components.
// Run: node scripts/pptx_slidegen/build_C_business_administration.js
//
// This mirrors slides/1st_stage/C_business_administration.html (表紙〜C-3) as a
// PowerPoint deck. When extending to C-4 onward, add new blocks below following the
// same pattern — see docs/13_pptx_slide_template_spec.md for the full guide.

const path = require("path");
const {
  newPres, addCoverSlide, addDividerSlide, addHeader, addFreqBar, mkYears,
  addTermRows, addRowList, addExamQuestion, drawScopeBars, drawVennOverlap,
  INK, INK_SOFT, RED, LINE, GHOST, F_BODY, F_MONO,
} = require("./lib");

const pres = newPres();
pres.author = "SMEC";
pres.title = "企業経営理論";

// ---------- Slide 01: Cover ----------
addCoverSlide(pres, {
  eyebrow: "中小企業診断士 第1次試験",
  subjectNo: "SUBJECT 03 / 07",
  title: "企業経営理論",
  subtitle: "経営戦略論・組織論・マーケティング論の3分野、42論点。うち40論点が頻出ランクA（直近10年で3回以上出題）という、1次試験随一の頻出科目。",
  stats: [["42", "収録論点数"], ["3", "分野"], ["40/42", "頻出ランクA"]],
  tag: "docs/textbook/C_business_administration_textbook.md",
  notes: "表紙スライド。企業経営理論の全体像を示す。",
});

// ---------- Slide 02: 経営戦略論 区切り ----------
addDividerSlide(pres, {
  ghostNo: "01",
  partNo: "PART 01",
  partLabel: "企業経営理論 ／ C-1〜C-14",
  title: "経営戦略論",
  desc: "企業がどの事業を持ち、どのように経営資源を配分して競争に勝つかを考える分野。VRIO分析やPPM、5フォース分析など、教科書レベルでも重要フレームワークが集中する。",
  chips: ["C-1 VRIO分析", "C-4 PPM", "C-6 5フォース分析"],
  notes: "経営戦略論パートの区切りスライド。",
});

// ---------- Slide 03: C-1 VRIO ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-1 ／ 戦略の考え方・階層別戦略・VRIO",
    title: "VRIO分析で競争優位の源泉を見抜く",
    overview: "自社の経営資源が本当に競争力の源泉になっているかを、4つの視点でチェックする。",
    tag: "企業経営理論",
  });
  // VRIOの4基準はすでに用語リストで十分に視覚化できており、右側に同じ内容を
  // 円で再掲するだけの図解は情報を追加しないと判断し、全幅テキストにする。
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "経営戦略には企業戦略・事業戦略・機能別戦略の3階層があり、まず自社を取り巻く状況を分析する。代表的な枠組みがVRIO分析：経営資源が競争力の源泉かを4つの視点でチェックする。",
    { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 }
  );
  cy += 0.56;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "V：価値", v: "外部の機会をつかむ・脅威を防ぐのに役立つか" },
    { k: "R：希少性", v: "他社があまり持っていない珍しい資源か" },
    { k: "I：模倣困難性", v: "他社が真似しようとしてもすぐには真似できないか" },
    { k: "O：組織", v: "その資源を活かせる組織体制になっているか" },
  ], { fontSize: 11.5, labelW: 1.55, gap: 0.3 });
  cy += 0.06;
  s.addShape("rect", { x: proseX, y: cy, w: proseW, h: 0.4, fill: { color: GHOST }, line: { type: "none" } });
  s.addShape("line", { x: proseX, y: cy, w: 0, h: 0.4, line: { color: INK, width: 2.5 } });
  s.addText("この4つがすべて揃って初めて「持続的な競争優位」が実現する。1つでも欠けると優位性は一時的なもので終わる。", {
    x: proseX + 0.15, y: cy, w: proseW - 0.3, h: 0.4, valign: "middle",
    fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0,
  });
  cy += 0.48;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "三代続く和菓子屋の「職人の技」：おいしい和菓子が作れる（価値）、他店にはない技術（希少性）、何十年もかけないと習得できない（模倣困難性）、技を商品化・販売できる体制がある（組織）→ 4つすべて揃っている。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.93;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "模倣困難性の3大源泉は「経路依存性（長い歴史の積み重ね）」「因果関係の曖昧性」「社会的複雑性（組織文化・人間関係等）」。逆に「価値の源泉が業界に知れ渡っている」「代替資源が外部調達できる」は模倣困難性を", options: { color: RED } },
    { text: "弱める", options: { bold: true, color: RED } },
    { text: "要因。「資源の価値が特定市場に限定される」はVの範囲の話で、模倣困難性(I)とは別軸。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-6 業界の構造分析、C-9〜C-11 技術経営など、外部・内部環境分析の土台になる論点全般と結びつく。",
    years: mkYears(new Set(["'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 04: C-1 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-1 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第3問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "企業経営理論",
  });
  const cy = addExamQuestion(s, {
    stem: "VRIOフレームワークに基づけば、自社の経営資源の模倣困難性は企業の持続的な競争優位性を左右する。競合企業が経営資源を模倣する際のコスト上の不利をもたらし、自社の経営資源の模倣困難性を高める要因として、最も適切なものはどれか。",
    choices: [
      { badge: "ア", text: "競合企業には他に将来性が高く注力するべき経営資源があること。" },
      { badge: "イ", text: "自社が歴史的な経緯で長い時間をかけて、独自の経営資源を獲得してきたこと。" },
      { badge: "ウ", text: "自社の経営資源がどのように競争優位性につながっているのかが既に明確になっており、業界に知れ渡っていること。" },
      { badge: "エ", text: "自社の経営資源の価値が特定の市場だけに限定されており、他市場では活用できないこと。" },
      { badge: "オ", text: "自社の経営資源を代替できる別の経営資源が外部の企業から調達可能であること。" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第3問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 05: C-1 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-1 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第3問）",
    overview: "正解はイ。ウ・オを弱める要因と取り違えていないか確認する。",
    tag: "企業経営理論",
  });
  const choices = [
    { badge: "ア", text: "競合企業には他に将来性が高く注力するべき経営資源があること。" },
    { badge: "イ", text: "自社が歴史的な経緯で長い時間をかけて、独自の経営資源を獲得してきたこと。" },
    { badge: "ウ", text: "自社の経営資源がどのように競争優位性につながっているのかが既に明確になっており、業界に知れ渡っていること。" },
    { badge: "エ", text: "自社の経営資源の価値が特定の市場だけに限定されており、他市場では活用できないこと。" },
    { badge: "オ", text: "自社の経営資源を代替できる別の経営資源が外部の企業から調達可能であること。" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 1 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：イ", options: { bold: true, color: RED } },
    { text: "（経路依存性）。ウ・オはむしろ模倣困難性を", options: {} },
    { text: "弱める", options: { bold: true } },
    { text: "要因（業界に知れ渡っている＝因果関係の曖昧性の喪失／代替資源が外部調達可能）。エは資源価値の適用範囲(V)の話で模倣困難性(I)とは別軸。アは自社の模倣困難性と無関係。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.75, fontFace: F_BODY, fontSize: 11, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });
  cy += 0.85;
  s.addText("関連知識：模倣困難性の3大源泉は「経路依存性」「因果関係の曖昧性」「社会的複雑性」（この設問はV・Iのみが対象でR・Oは問われていない）。", {
    x: 0.55, y: cy, w: 12.25, h: 0.4, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.5;
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第3問／正解：past_exams/1st_stage_answers/r07/2025c.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 06: C-2 ドメイン ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-2 ／ 企業戦略：ドメイン（企業・事業ドメイン）",
    title: "ドメインの2つの定義方法",
    overview: "「自社は何屋さんか」の定義の仕方一つで、気づける事業機会の広さが変わる。",
    tag: "企業経営理論",
  });
  // 「狭い/広い」という視野の広さの比較は文章だけでは伝わりにくく、幅の違う
  // バーで見せると一目で分かるため、右側に図解を置く。
  const proseX = 0.55, proseW = 7.4;
  const diagX = 8.25, diagW = 4.05;
  let cy = 1.85;
  s.addText([
    { text: "ドメイン", options: { bold: true, color: INK } },
    { text: "＝「自社が事業を営む領域」。「うちの会社は何屋さんか」を定義すること。", options: { color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.3, fontFace: F_BODY, fontSize: 11.5, isTextBox: true, margin: 0 });
  cy += 0.38;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "物理的定義", v: "作っている「モノ」で定義する（例：「うちは鉄道車両を作る会社だ」）", gap: 0.32 },
    { k: "機能的定義", v: "提供している「価値・機能」で定義する（例：「うちは人や物を運ぶサービスを提供する会社だ」）", gap: 0.5 },
  ], { fontSize: 11.5, labelW: 1.55 });
  cy += 0.1;
  s.addText(
    "機能的定義の方が視野が広がり新しい事業機会に気づきやすい。「鉄道会社が自分たちを『鉄道業』と狭く定義したために自動車・航空機の台頭に対応できなかった」というレビットの「マーケティング近視眼」は有名な失敗例。",
    { x: proseX, y: cy, w: proseW, h: 0.65, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 }
  );
  cy += 0.72;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "コンビニチェーンが自社を「小売業」ではなく「生活インフラ業」と定義し直すと、公共料金の収納代行、ATM、宅配便の受付など、単なる「モノを売る」以上のサービス展開に踏み出しやすくなる。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.93;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "ドメインの決め方（物理的 vs 機能的）と、C-3の「多角化」の判断を混同しないこと。ドメインは「自社の立ち位置の定義」、多角化は「新しい事業領域への進出」という別の概念。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.45, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  drawScopeBars(s, diagX, 2.3, diagW, 3.0, {
    narrowLabel: "物理的定義（モノ）",
    wideLabel: "機能的定義（価値・機能）",
    resultLabel: "→ 気づける事業機会が広がる",
  });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-3 多角化（アンゾフの成長マトリクスの発想の土台になる）。",
    years: mkYears(new Set(["'16", "'17", "'19", "'23"])),
  });
}

// ---------- Slide 07: アンゾフの成長マトリクス ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-2 ／ 企業戦略：アンゾフの成長マトリクス",
    title: "誰に・何を売るかで成長戦略を4つに分ける",
    overview: "製品と市場がそれぞれ「既存」か「新規」かの組み合わせで、成長戦略は4パターンに整理できる。",
    tag: "企業経営理論",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "市場浸透戦略", tag: "既存製品×既存市場", desc: "既存顧客のリピート購入促進やシェア拡大など、今ある製品・市場のままで成長を図る" },
    { name: "新製品開発戦略", tag: "新製品×既存市場", desc: "既存顧客層向けに新しい製品を投入する（例：既存客向けの新フレーバー・新機能）" },
    { name: "新市場開拓戦略", tag: "既存製品×新市場", desc: "今の製品のまま、新しい顧客層・地域・海外市場に展開する" },
    { name: "多角化戦略", tag: "新製品×新市場", desc: "新しい製品で、新しい市場に進出する。既存の強みが活かせるかで関連型・非関連型に分かれる（C-3参照）" },
  ]);

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "解き方のコツ：選択肢の具体的施策が実際にはどの象限（製品×市場の新旧）に当たるかを先に判定し、選択肢が主張する戦略名と一致するかを照合する。",
    years: mkYears(new Set(["'16", "'17", "'19", "'23"])),
  });
}

// ---------- Slide 08: C-2 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-2 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第2問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "企業経営理論",
  });
  const cy = addExamQuestion(s, {
    stem: "食品メーカーA社は、日本国内の一般消費者向けのシリアルの開発・製造・販売を行う専業企業である。同社は、現在、今後の成長戦略のための施策を検討している。アンゾフの成長マトリクスの考え方に基づく記述として、最も適切なものはどれか。",
    choices: [
      { badge: "ア", text: "自社ECサイトを立ち上げ、既存の一般消費者向けシリアル市場でのシェア拡大を狙うことで、「多角化戦略」を行う。" },
      { badge: "イ", text: "飲食事業者向けのSNSマーケティング支援事業を新しく開始し、次の収益の柱とすることで、「市場浸透戦略」を行う。" },
      { badge: "ウ", text: "一部の既存販売地域向けに、地域名産品を用いた限定フレーバーを新しく開発することで、「多角化戦略」を行う。" },
      { badge: "エ", text: "既存顧客のリピート購入を促進するキャンペーンを実施し、既存市場でのシェア拡大を狙うことで、「新市場開拓戦略」を行う。" },
      { badge: "オ", text: "コア・ユーザーである中高年層向けに、新しく低糖質シリアルを開発することで、「新製品開発戦略」を行う。" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第2問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 09: C-2 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-2 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第2問）",
    overview: "正解はオ。施策の実態と選択肢が主張する戦略名が一致するかで見分ける。",
    tag: "企業経営理論",
  });
  const choices = [
    { badge: "ア", text: "自社ECサイトを立ち上げ、既存の一般消費者向けシリアル市場でのシェア拡大を狙うことで、「多角化戦略」を行う。" },
    { badge: "イ", text: "飲食事業者向けのSNSマーケティング支援事業を新しく開始し、次の収益の柱とすることで、「市場浸透戦略」を行う。" },
    { badge: "ウ", text: "一部の既存販売地域向けに、地域名産品を用いた限定フレーバーを新しく開発することで、「多角化戦略」を行う。" },
    { badge: "エ", text: "既存顧客のリピート購入を促進するキャンペーンを実施し、既存市場でのシェア拡大を狙うことで、「新市場開拓戦略」を行う。" },
    { badge: "オ", text: "コア・ユーザーである中高年層向けに、新しく低糖質シリアルを開発することで、「新製品開発戦略」を行う。" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 4 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：オ", options: { bold: true, color: RED } },
    { text: "。既存顧客層（既存市場）向けに新しい低糖質シリアル（新製品）＝新製品開発戦略で、施策の実態と戦略名が一致。ア〜エはすべて、施策の実態（製品×市場の組み合わせ）と選択肢が主張する戦略名が食い違っている。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.6, fontFace: F_BODY, fontSize: 11, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });
  cy += 0.7;
  s.addText("関連知識：まず各選択肢の施策が実際にはどの象限かを判定し、その上で主張されている戦略名と一致するか照合するのが速い解き方。", {
    x: 0.55, y: cy, w: 12.25, h: 0.4, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.5;
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第2問／正解：past_exams/1st_stage_answers/r07/2025c.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 10: C-3 多角化 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-3 ／ 企業戦略：多角化（シナジー、関連型・非関連型）",
    title: "多角化とシナジーの関係",
    overview: "既存事業との関連性が高いほど、多角化はシナジーを生みやすい。",
    tag: "企業経営理論",
  });
  // 「既存事業との重なり＝シナジー」は円の重なりというまさに図解向きの構造
  // なので、右側にベン図を置く。
  const proseX = 0.55, proseW = 7.4;
  const diagX = 8.25, diagW = 4.05;
  let cy = 1.85;
  s.addText([
    { text: "多角化", options: { bold: true, color: INK } },
    { text: "＝既存事業とは別の新しい事業に進出すること。2種類ある。", options: { color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.3, fontFace: F_BODY, fontSize: 11.5, isTextBox: true, margin: 0 });
  cy += 0.38;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "関連型多角化", v: "既存事業と技術・顧客・販売チャネル等で関連がある分野に進出（例：カメラメーカーが医療用内視鏡事業に進出。レンズ技術という共通の強みを活かせる）", gap: 0.75 },
    { k: "非関連型多角化", v: "既存事業とほとんど関連がない分野に進出（例：繊維会社が化粧品事業に進出）", gap: 0.5 },
  ], { fontSize: 11.5, labelW: 1.55 });
  cy += 0.1;
  s.addShape("rect", { x: proseX, y: cy, w: proseW, h: 0.48, fill: { color: GHOST }, line: { type: "none" } });
  s.addShape("line", { x: proseX, y: cy, w: 0, h: 0.48, line: { color: INK, width: 2.5 } });
  s.addText([
    { text: "シナジー（相乗効果）：", options: { bold: true } },
    { text: "「1+1が2ではなく3以上になる効果」。技術・ブランド・販売網を複数事業で共有すると単独でやるより効率が良くなる。関連型多角化の方がシナジーを生みやすい。", options: {} },
  ], { x: proseX + 0.15, y: cy, w: proseW - 0.3, h: 0.48, valign: "middle", fontFace: F_BODY, fontSize: 10, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  cy += 0.58;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "富士フイルムは写真フィルム事業で培った化学・材料技術（酸化を防ぐ技術等）を応用して化粧品事業（アンチエイジング）に進出した。一見畑違いだが技術的関連性が高い「関連型多角化」の代表例。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.93;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「新市場に新技術で進出する」ことが必ずしも「多角化」とは限らない。アンゾフの成長マトリクスの「新製品×新市場」が多角化に該当。既存製品のままの海外展開（新市場開拓）や既存市場向けの新商品（新製品開発）と混同しないこと。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  drawVennOverlap(s, diagX, 2.2, diagW, 2.6, {
    leftLabel: "既存事業", rightLabel: "新規事業", overlapLabel: "シナジー",
  });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-2 ドメイン（自社の立ち位置の定義が多角化判断の前提になる）。",
    years: mkYears(new Set(["'18", "'21", "'22", "'24", "'25"])),
  });
}

const outPath = path.join(__dirname, "..", "..", "slides", "1st_stage", "C_business_administration.pptx");
pres.writeFile({ fileName: outPath }).then(() => {
  console.log("wrote", outPath);
});
