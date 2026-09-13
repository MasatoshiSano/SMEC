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
  drawQuadrant, draw5Forces,
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

// ---------- Slide 11: C-4 PPM ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-4 ／ 企業戦略：PPM（プロダクト・ポートフォリオ・マネジメント）",
    title: "PPMで事業ポートフォリオの資金配分を判断する",
    overview: "市場成長率×市場シェアの2軸で事業を4分類し、どこに投資しどこから資金を引き上げるかを判断する。",
    tag: "企業経営理論",
  });
  drawQuadrant(s, 0.55, 1.9, 12.25, 3.9, {
    cells: [
      { pos: "tl", label: "花形（スター）", sublabel: "稼ぐが投資も必要。将来の主力候補" },
      { pos: "tr", label: "問題児", sublabel: "将来性はあるが今は稼げない。育てるか撤退か判断が必要" },
      { pos: "bl", label: "金のなる木", sublabel: "投資が要らず安定して稼ぐ。他事業の資金源" },
      { pos: "br", label: "負け犬", sublabel: "稼げず将来性も薄い。撤退候補" },
    ],
    axisCaption: "縦軸：市場成長率（上＝高い）／横軸：市場シェア（左＝高い）。「金のなる木」の資金を「問題児」に投資して「花形」に育てるのが基本セオリー。",
  });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：PPMは「市場の魅力度」と「自社の競争力」の2軸だけの単純化モデルで、事業間のシナジーや社会的意義のある事業は考慮できないという限界がある。",
    years: mkYears(new Set(["'16", "'17", "'19", "'20", "'21", "'22", "'23", "'24"])),
  });
}

// ---------- Slide 12: C-5 垂直統合・M&A・戦略的提携 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-5 ／ 企業戦略：垂直統合、M&A、戦略的提携",
    title: "事業を拡大・強化する3つの方法",
    overview: "資本を統合するか（M&A）、契約ベースで組むか（提携）、自社に取り込むか（垂直統合）で選択肢が変わる。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "会社が事業を拡大・強化する主な方法は3つ。資本を統合する度合いが異なる。",
    { x: proseX, y: cy, w: proseW, h: 0.3, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0 }
  );
  cy += 0.38;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "垂直統合", v: "自社のバリューチェーンの川上（原材料側）・川下（消費者側）を自社に取り込む" },
    { k: "M&A", v: "他社を買収・合併し、技術・顧客・ブランドを一気に手に入れる（TOB・MBOはその手法）" },
    { k: "戦略的提携", v: "資本を統合せず契約ベースで協力する（ジョイントベンチャーもこの一種）" },
  ], { fontSize: 11.5, labelW: 1.4, gap: 0.42 });
  cy += 0.08;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "自動車部品メーカーが鋼材メーカーを買収すれば「川上への垂直統合」、完成車メーカーを買収すれば「川下への垂直統合」。2社が共同出資でJVを設立するのは「戦略的提携」。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.68;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "ジョイントベンチャーは独立性を保つため、経営統合するM&Aより共有できる情報の範囲は一般に", options: { color: RED } },
    { text: "狭い", options: { bold: true, color: RED } },
    { text: "（広いは誤り）。MBOは上場企業の非公開化だけでなく、非上場企業の事業承継の手段としても行われる。CVC（事業会社がベンチャー企業へ出資しシナジーを狙う手法）も周辺用語として頻出。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-11 新規事業開発（CVCによるベンチャー投資）、C-14 事業承継（MBOによる従業員承継）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 13: C-5 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-5 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第6問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "企業経営理論",
  });
  const cy = addExamQuestion(s, {
    stem: "企業間の連携戦略に関する記述として、最も適切なものはどれか。",
    choices: [
      { badge: "ア", text: "TOBとは、買収者が対象企業の株式を公開市場で株主から買い付ける手法のことを指す。" },
      { badge: "イ", text: "経営陣が中核となって既存株主から株式を買い取ることをMBOと呼ぶが、MBOは上場企業では起こるが、非上場企業では起こらない。" },
      { badge: "ウ", text: "仕入先の事業を買収し、事業のバリューチェーンの変革を目指す買収を水平的M&Aと呼ぶ。" },
      { badge: "エ", text: "ジョイントベンチャーは、M&Aよりも当事者同士で共有できる情報の範囲が広く範囲の経済を享受できるので、より大きな相乗効果が期待される。" },
      { badge: "オ", text: "ベンチャー企業が開発した革新的な技術やビジネスモデルを取り込み、自社の既存事業との間でシナジーを発現させることなどを目的に、事業会社がベンチャー企業に投資をすることをCVC（コーポレート・ベンチャー・キャピタル）と呼ぶ。" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第6問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 14: C-5 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-5 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第6問）",
    overview: "正解はオ。TOB・MBO・垂直/水平M&A・JVの定義のズレを見抜く。",
    tag: "企業経営理論",
  });
  const choices = [
    { badge: "ア", text: "TOBとは、買収者が対象企業の株式を公開市場で株主から買い付ける手法のことを指す。" },
    { badge: "イ", text: "経営陣が中核となって既存株主から株式を買い取ることをMBOと呼ぶが、MBOは上場企業では起こるが、非上場企業では起こらない。" },
    { badge: "ウ", text: "仕入先の事業を買収し、事業のバリューチェーンの変革を目指す買収を水平的M&Aと呼ぶ。" },
    { badge: "エ", text: "ジョイントベンチャーは、M&Aよりも当事者同士で共有できる情報の範囲が広く範囲の経済を享受できるので、より大きな相乗効果が期待される。" },
    { badge: "オ", text: "ベンチャー企業が開発した革新的な技術やビジネスモデルを取り込み、自社の既存事業との間でシナジーを発現させることなどを目的に、事業会社がベンチャー企業に投資をすることをCVC（コーポレート・ベンチャー・キャピタル）と呼ぶ。" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 4 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：オ", options: { bold: true, color: RED } },
    { text: "。CVCの定義そのもの。ア：TOBは通常「取引所を介さず」株主から直接買い集める手法で「公開市場で買い付け」は不正確。イ：MBOは非上場企業の事業承継手段としても行われる。ウ：仕入先（川上）の買収は", options: {} },
    { text: "垂直的", options: { bold: true } },
    { text: "M&A（水平的は誤り）。エ：JVは独立性を保つため共有情報の範囲はM&Aより", options: {} },
    { text: "狭い", options: { bold: true } },
    { text: "（広いは誤り）。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.85, fontFace: F_BODY, fontSize: 11, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });
  cy += 0.93;
  s.addText("関連知識：M&Aの類型（水平・垂直・多角化型）とTOB・MBO・CVC等の周辺用語はセットで頻出。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.4;
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第6問／正解：past_exams/1st_stage_answers/r07/2025c.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 15: C-6 5フォース・価値連鎖 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-6 ／ 競争戦略：業界の構造分析（5フォース、価値連鎖）",
    title: "5つの力で「儲かりやすい業界か」を判断する",
    overview: "業界の競争構造を5つの力（ポーター）に分解し、力が弱いほど儲けやすい業界と判断する。",
    tag: "企業経営理論",
  });
  draw5Forces(s, 1.4, 1.85, 10.5, 3.5, {
    center: "業界内の\n既存競合",
    top: "新規参入の脅威\n（参入障壁が低いほど大）",
    bottom: "代替品の脅威（例：新聞にとってのニュースアプリ）",
    left: "売り手の\n交渉力",
    right: "買い手の\n交渉力",
  });
  s.addText([
    { text: "もう1つの分析軸が", options: {} },
    { text: "バリューチェーン（価値連鎖）分析", options: { bold: true } },
    { text: "：活動を「主活動」（購買物流→製造→出荷物流→販売→サービス）と「支援活動」（全般管理・人事・技術開発・調達）に分け、付加価値の源泉を探る。", options: {} },
  ], { x: 0.55, y: 5.5, w: 12.25, h: 0.5, fontFace: F_BODY, fontSize: 9.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：「参入障壁が高く撤退障壁が低い」のように複数条件を組み合わせた業界特性は、単純発想では判断できない。設問の前提条件を個別に確認すること。",
    years: mkYears(new Set(["'16", "'18", "'19", "'20", "'21", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 16: C-7 競争優位の3つの基本戦略 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-7 ／ 競争戦略：競争優位の戦略（コスト・差別化・集中）",
    title: "競争に勝つための3つの基本戦略",
    overview: "「競争優位の種類」×「狙う市場の広さ」の組み合わせで、とるべき戦略が変わる（ポーター）。",
    tag: "企業経営理論",
  });
  drawQuadrant(s, 0.55, 1.9, 12.25, 3.9, {
    cells: [
      { pos: "tl", label: "コストリーダーシップ", sublabel: "業界最低コストで価格競争に勝つ（例：格安航空会社）" },
      { pos: "tr", label: "差別化", sublabel: "品質・ブランド・デザインで独自の価値を出す（例：高級ホテル）" },
      { pos: "bl", label: "コスト集中", sublabel: "特定の狭い顧客層に絞ってコストで勝つ" },
      { pos: "br", label: "差別化集中", sublabel: "特定の狭い顧客層に絞って差別化で勝つ" },
    ],
    axisCaption: "縦軸：狙う市場の広さ（上＝広い市場全体、下＝特定の狭い層＝集中戦略）／横軸：競争優位の源泉（左＝コスト、右＝差別化）。両方を中途半端に狙うと収益性が落ちる（スタック・イン・ザ・ミドル）。",
  });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：集中戦略は「誰に売るか」の軸で、コストリーダーシップ・差別化そのものとは別軸。ユニクロのSPAによる効率化は単純な安売りでなく差別化的コストリーダーシップと解釈される。",
    years: mkYears(new Set(["'16", "'17", "'19", "'21", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 17: C-8 競争地位別戦略 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-8 ／ 競争戦略：競争地位別戦略",
    title: "業界内シェアの順位で戦い方を変える（コトラー）",
    overview: "市場シェアの順位に応じて、企業が取るべき戦略は4パターンに分かれる。",
    tag: "企業経営理論",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "リーダー", tag: "業界1位", desc: "市場全体を拡大させ、他社の攻撃を防ぐ「全方位戦略」" },
    { name: "チャレンジャー", tag: "2位以下・体力あり", desc: "リーダーとの「差別化」戦略。同じ土俵で戦うと体力負けするため" },
    { name: "ニッチャー", tag: "特定分野に特化", desc: "特定のセグメントに経営資源を集中する「集中」戦略" },
    { name: "フォロワー", tag: "シェア・体力とも劣る", desc: "リーダーの成功パターンを模倣し、低コストで生き残る" },
  ], { nameW: 2.1, tagW: 2.3 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：チャレンジャーが取るべきは「同質化（真似）」ではなく「差別化」。同質化はむしろフォロワーの戦略に近い考え方であり、混同しないこと。",
    years: mkYears(new Set(["'16", "'22", "'23"])),
  });
}

// ---------- Slide 18: C-9 MOT・コア技術戦略 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-9 ／ 技術経営（MOT）：技術戦略、知財戦略",
    title: "コア技術を軸に複数事業へ展開する",
    overview: "技術力を経済的価値に変えるマネジメント。コア技術戦略と知財戦略が中心。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText([
    { text: "MOT（技術経営）", options: { bold: true, color: INK } },
    { text: "は技術力を利益に変えるマネジメント。代表例が", options: { color: INK } },
    { text: "コア技術戦略", options: { bold: true, color: INK } },
    { text: "：特定の技術分野に集中し、複数の製品・事業に展開する戦略。", options: { color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 11.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });
  cy += 0.56;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "コア技術戦略", v: "1つの技術を土台に複数製品を展開し、得た知見を技術自体の強化に還元する好循環" },
    { k: "知財戦略", v: "特許として公開・独占するか、ノウハウとして秘匿するかの選択も含む" },
  ], { fontSize: 11.5, labelW: 1.55, gap: 0.4 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "「精密加工技術」をコア技術に持つ工作機械メーカーが、医療機器・航空宇宙部品にも展開し、各分野で得たノウハウを工作機械の技術力向上にも還元する。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.68;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "コア技術戦略は「既存技術に固執し続けること」ではない。陳腐化・模倣時は技術自体を進化させるのが前提。「既存市場でのシェア拡大が最優先」も誤りで、本質は", options: { color: RED } },
    { text: "複数事業への横展開", options: { bold: true, color: RED } },
    { text: "。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-10 イノベーションマネジメント、C-11 新規事業開発と同じ技術経営（MOT）分野。",
    years: mkYears(new Set(["'16", "'19", "'20", "'21", "'24", "'25"])),
  });
}

// ---------- Slide 19: C-9 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-9 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第9問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "企業経営理論",
  });
  const cy = addExamQuestion(s, {
    stem: "特定の技術分野に集中し、その技術をベースとした製品を次々と開発・導入する戦略として「コア技術戦略」がある。コア技術戦略の特徴に関する記述として、最も適切なものはどれか。",
    choices: [
      { badge: "ア", text: "コア技術戦略では、コア技術が陳腐化したり、競合企業に模倣されたりしても、コア技術の入れ替えは考えず、既存のコア技術にこだわった技術開発や製品開発を追求する。" },
      { badge: "イ", text: "コア技術戦略では、コア技術の活用に注力し、既存市場でのシェア拡大を最優先の目標とする。" },
      { badge: "ウ", text: "コア技術戦略では、コア技術を活用して顧客ニーズに合致した製品を開発することが望ましいが、開発の初期段階ではコア技術よりも既存顧客の要求を優先する。" },
      { badge: "エ", text: "コア技術戦略では、コア技術を基盤に、多様な製品を開発し、その学習成果をコア技術の強化や発展につなげる。" },
      { badge: "オ", text: "コア技術戦略では、特定の技術に経営資源を集中させるため、事業の多角化が難しく、コア技術に依存するリスクを分散しにくい。" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第9問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 20: C-9 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-9 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第9問）",
    overview: "正解はエ。「複数製品への横展開」がコア技術戦略の本質。",
    tag: "企業経営理論",
  });
  const choices = [
    { badge: "ア", text: "コア技術戦略では、コア技術が陳腐化したり、競合企業に模倣されたりしても、コア技術の入れ替えは考えず、既存のコア技術にこだわった技術開発や製品開発を追求する。" },
    { badge: "イ", text: "コア技術戦略では、コア技術の活用に注力し、既存市場でのシェア拡大を最優先の目標とする。" },
    { badge: "ウ", text: "コア技術戦略では、コア技術を活用して顧客ニーズに合致した製品を開発することが望ましいが、開発の初期段階ではコア技術よりも既存顧客の要求を優先する。" },
    { badge: "エ", text: "コア技術戦略では、コア技術を基盤に、多様な製品を開発し、その学習成果をコア技術の強化や発展につなげる。" },
    { badge: "オ", text: "コア技術戦略では、特定の技術に経営資源を集中させるため、事業の多角化が難しく、コア技術に依存するリスクを分散しにくい。" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 3 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：エ", options: { bold: true, color: RED } },
    { text: "。単一のコア技術を軸に複数製品・事業へ展開し、得た学習成果を技術自体の強化に還元する好循環がコア技術戦略の核心。ア：陳腐化・模倣時も固執するのは「コンピタンス・トラップ」に近く誤り。イ：既存市場シェア拡大最優先は市場浸透戦略の説明で誤り。オ：コア技術を軸にした展開はむしろ", options: {} },
    { text: "多角化しやすい", options: { bold: true } },
    { text: "。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.85, fontFace: F_BODY, fontSize: 11, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });
  cy += 0.93;
  s.addText("関連知識：1つの技術基盤を複数事業に展開する考え方は「技術のプラットフォーム化」「技術シナジー」とも関連する。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.4;
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第9問／正解：past_exams/1st_stage_answers/r07/2025c.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 21: C-10 イノベーションマネジメント ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-10 ／ 技術経営（MOT）：イノベーションマネジメント",
    title: "3つのイノベーション理論を区別する",
    overview: "「誰が」「どこから」イノベーションを起こすかで異なる3つの代表理論がある。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "イノベーションに関する3つの代表理論。「誰が」「どこから」起こすかがそれぞれ異なる。",
    { x: proseX, y: cy, w: proseW, h: 0.3, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0 }
  );
  cy += 0.38;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "イノベーションのジレンマ", v: "既存顧客への対応に注力するあまり、新興企業の破壊的イノベーションに対応できず地位を奪われる（クリステンセン）", gap: 0.55 },
    { k: "オープンイノベーション", v: "社外（大学・他社・スタートアップ）のアイデア・技術を積極的に取り込む考え方", gap: 0.4 },
    { k: "ユーザー・イノベーション", v: "企業でなくユーザー自身が改良・開発を行う。先端的ニーズを持つリード・ユーザーが主な担い手（フォン・ヒッペル）", gap: 0.55 },
  ], { fontSize: 11.5, labelW: 2.3 });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "ユーザー・イノベーションは「企業が顧客ニーズをアンケート調査すること」ではなく、", options: { color: RED } },
    { text: "ユーザー自身", options: { bold: true, color: RED } },
    { text: "が開発主体。個人単独よりコミュニティを通じた方が普及しやすい。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "具体例：フィルムカメラの高性能化に注力する大手を尻目に、性能は劣るが手軽なデジタルカメラで新興メーカーが台頭した例が「イノベーションのジレンマ」の典型。",
    years: mkYears(new Set(["'16", "'17", "'18", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 22: C-10 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-10 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第10問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "企業経営理論",
  });
  const cy = addExamQuestion(s, {
    stem: "ユーザー・イノベーションに関する記述として、最も適切なものはどれか。",
    choices: [
      { badge: "ア", text: "個人によるユーザー・イノベーションは、コミュニティを通じたユーザー・イノベーションよりも普及しやすい傾向にある。" },
      { badge: "イ", text: "ユーザー・イノベーションとは、企業の開発者が主体となり、アンケートやインタビューを活用して、顧客のニーズを調査し、その結果を製品開発に反映するプロセスである。" },
      { badge: "ウ", text: "ユーザー・イノベーションは、イノベーションのジレンマを引き起こし、既存企業が破壊的イノベーションに対応できなくなる原因となる。" },
      { badge: "エ", text: "ユーザー・イノベーションは、ユーザーが持つニーズ情報の粘着性が高く、技術情報の粘着性が低い場合に起こりやすい。" },
      { badge: "オ", text: "ユーザー・イノベーションは、リード・ユーザーと呼ばれる、特定の企業への忠誠度が高いユーザーによって引き起こされる傾向が強い。" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第10問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 23: C-10 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-10 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第10問）",
    overview: "正解はオ。主体の取り違え・普及しやすさの逆転に注意する。",
    tag: "企業経営理論",
  });
  const choices = [
    { badge: "ア", text: "個人によるユーザー・イノベーションは、コミュニティを通じたユーザー・イノベーションよりも普及しやすい傾向にある。" },
    { badge: "イ", text: "ユーザー・イノベーションとは、企業の開発者が主体となり、アンケートやインタビューを活用して、顧客のニーズを調査し、その結果を製品開発に反映するプロセスである。" },
    { badge: "ウ", text: "ユーザー・イノベーションは、イノベーションのジレンマを引き起こし、既存企業が破壊的イノベーションに対応できなくなる原因となる。" },
    { badge: "エ", text: "ユーザー・イノベーションは、ユーザーが持つニーズ情報の粘着性が高く、技術情報の粘着性が低い場合に起こりやすい。" },
    { badge: "オ", text: "ユーザー・イノベーションは、リード・ユーザーと呼ばれる、特定の企業への忠誠度が高いユーザーによって引き起こされる傾向が強い。" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 4 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：オ", options: { bold: true, color: RED } },
    { text: "。フォン・ヒッペルの提唱した「リード・ユーザー」（市場ニーズを先取りする強いニーズを持つユーザー）が主な担い手。ア：個人単独よりコミュニティを通じた方が普及しやすく記述は逆。イ：主体は企業でなく", options: {} },
    { text: "ユーザー自身", options: { bold: true } },
    { text: "で誤り。ウ：ジレンマの原因はユーザー・イノベーションではなく持続的イノベーションへの偏重。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.85, fontFace: F_BODY, fontSize: 11, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });
  cy += 0.93;
  s.addText("関連知識：リード・ユーザー法は新製品開発のマーケティングリサーチ手法として2次試験（事例II）でも問われることがある。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.4;
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第10問／正解：past_exams/1st_stage_answers/r07/2025c.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 24: C-11 新規事業開発・リーンスタートアップ ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-11 ／ 技術経営（MOT）：新規事業開発",
    title: "リーン・スタートアップで新規事業の不確実性に対処する",
    overview: "完成品をいきなり作らず、最小限の製品で顧客の反応を見ながら素早く軌道修正する。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText([
    { text: "リーン・スタートアップ", options: { bold: true, color: INK } },
    { text: "（エリック・リース）：新規性が高く顧客の存在も未知な事業で、完成品を作る前に検証を繰り返す方法論。", options: { color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.35, fontFace: F_BODY, fontSize: 11.5, isTextBox: true, margin: 0 });
  cy += 0.42;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "MVP", v: "実用最小限の製品。コストをかけず最低限の機能だけを持つ試作版" },
    { k: "構築－計測－学習", v: "MVPを顧客に使ってもらい反応を見て学習する、というサイクルを素早く回す" },
    { k: "ピボット", v: "学習結果を踏まえ、必要なら事業の方向転換をすること" },
  ], { fontSize: 11.5, labelW: 2.0, gap: 0.4 });
  cy += 0.1;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "リーン・スタートアップは「トヨタ生産方式」（ムダを省き素早く改善サイクルを回す）の考え方から影響を受けている。流行に敏感で自ら情報収集する「アーリー・アダプター」を巻き込むことも推奨される。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "具体例：新アプリで、いきなり全機能を実装せず最低限の試作版をリリースし、実際の反応を見て本格開発の方向性を決めるのがMVPの考え方。",
    years: mkYears(new Set(["'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 25: C-12 国際経営戦略：海外進出形態 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-12 ／ 国際経営戦略：海外進出形態",
    title: "投資リスクの違いで3つの海外進出形態を選ぶ",
    overview: "輸出・ライセンス供与・海外直接投資の順にリスクとコントロールの強さが増す。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "企業が海外に進出する形態には段階があり、リスクとコントロールの強さがトレードオフになる。",
    { x: proseX, y: cy, w: proseW, h: 0.3, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0 }
  );
  cy += 0.38;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "輸出", v: "自国で作った製品を海外に売る。最も投資リスクが低い" },
    { k: "ライセンス供与", v: "技術・ブランドの使用権を海外企業に許諾しロイヤリティを得る" },
    { k: "海外直接投資（FDI）", v: "海外に子会社・工場を設立。投資額・リスクは大きいが現地のコントロールが強い" },
  ], { fontSize: 11.5, labelW: 2.1, gap: 0.4 });
  cy += 0.08;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "ダニングの折衷理論（OLIパラダイム）：①所有優位性（O：自社の技術・ブランド等の強み）、②立地優位性（L：進出先市場の魅力）、③内部化優位性（I：自社管理が有利か）の3条件が揃うときFDIが選ばれる。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.93;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "3条件のうち一部しか揃わない場合の進出形態を問う問題が頻出。例えば「所有優位性はあるが内部化優位性がない」場合は", options: { color: RED } },
    { text: "ライセンス供与", options: { bold: true, color: RED } },
    { text: "が適切、というように組み合わせごとに最適形態が変わる。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-5 M&A・戦略的提携（海外企業とのJVもOLIの内部化優位性の判断と関わる）。",
    years: mkYears(new Set(["'17", "'18", "'20", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 26: C-13 CSR・CSV・ガバナンス ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-13 ／ 企業の社会的責任：CSR・CSV・コーポレートガバナンス",
    title: "「コストか投資か」でCSRとCSVを区別する",
    overview: "社会的責任・共有価値の創造・企業統治という3つの関連概念を整理する。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "企業が利益追求だけでなく社会と向き合うための3つの関連概念。",
    { x: proseX, y: cy, w: proseW, h: 0.3, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0 }
  );
  cy += 0.38;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "CSR", v: "企業の社会的責任。環境保護・地域貢献など、利益とは別に果たすべき責任（コスト）" },
    { k: "CSV", v: "共有価値の創造（ポーター）。社会課題の解決そのものを事業機会・利益創出につなげる発想（投資）" },
    { k: "コーポレートガバナンス", v: "株主・社会の利益を守るため経営者の暴走を防ぐ統治の仕組み（社外取締役・監査役制度等）" },
  ], { fontSize: 11.5, labelW: 2.3, gap: 0.42 });
  cy += 0.08;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "農業支援で質の高い原料を安定調達し、自社製品の品質向上と現地農家の所得向上の両方につながる取り組みはCSVの典型例。単なる寄付活動（CSR）とは異なり、事業戦略そのものに組み込まれている。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.93;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "CSRとCSVは似ているが、「コストか投資か」「本業と一体か切り離されているか」という観点で区別される。混同しないこと。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：経営法務科目の会社法（社外取締役・監査役会等の機関設計）とも関連する。",
    years: mkYears(new Set(["'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 27: C-14 ファミリービジネス・事業承継 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-14 ／ ファミリービジネスの戦略：創業・事業承継",
    title: "事業承継の3つのパターン",
    overview: "日本の中小企業の多くは家族経営。経営権・所有権の引き継ぎ方は大きく3つに分かれる。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText([
    { text: "日本の中小企業の多くは家族経営（ファミリービジネス）。経営権・所有権を後継者にどう引き継ぐかが特有の課題（", options: { color: INK } },
    { text: "事業承継", options: { bold: true, color: INK } },
    { text: "）。", options: { color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 11.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.46;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "親族内承継", v: "子どもなど親族に引き継ぐ" },
    { k: "親族外承継", v: "役員・従業員に引き継ぐ（従業員承継。MBOもこの一種になりうる）" },
    { k: "M&A（第三者承継）", v: "外部の企業や個人に売却する" },
  ], { fontSize: 11.5, labelW: 2.0, gap: 0.35 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "後継者不在の町工場が長年勤めた工場長に経営を引き継がせる（従業員承継）ケースや、同業の大手企業に事業を売却する（M&A）ケースが近年増加している。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.68;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "ファミリービジネスの強み（長期的視点・迅速な意思決定）と弱み（後継者育成の難しさ・身内経営によるガバナンスの脆弱化）を対比で理解しておく。事業承継の法制度（経営承継円滑化法）は経営法務科目とも関連。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-5 M&A（第三者承継としてのM&A、MBOによる従業員承継）。",
    years: mkYears(new Set(["'17", "'18", "'20", "'21", "'22"])),
  });
}

const outPath = path.join(__dirname, "..", "..", "slides", "1st_stage", "C_business_administration.pptx");
pres.writeFile({ fileName: outPath }).then(() => {
  console.log("wrote", outPath);
});
