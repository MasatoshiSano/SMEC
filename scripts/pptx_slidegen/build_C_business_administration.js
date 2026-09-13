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
  drawQuadrant, draw5Forces, drawProcessSteps, drawProcessStepsVertical,
  INK, INK_SOFT, RED, LINE, GHOST, F_HEAD, F_BODY, F_MONO,
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
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "三代続く和菓子屋が長年かけて築いた「職人の技」は、他社が一朝一夕に習得できない（経路依存性＝イ）。逆にその技が広く知られていたり（ウ）、代替の資源を外部調達できたり（オ）すれば、模倣は容易になり優位性は弱まる。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.93;
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
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "コンビニが自社を「小売業」ではなく「生活インフラ業」と定義し直すと、公共料金の収納代行・ATM・宅配受付など、単なる「モノを売る」以上の展開がしやすくなる。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.93;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "ドメイン（自社の立ち位置の定義）とC-3「多角化」（新領域への進出）を混同しないこと。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.45, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  drawScopeBars(s, diagX, 2.3, diagW, 3.0, {
    narrowLabel: "物理的定義（モノ）",
    wideLabel: "機能的定義（価値・機能）",
    resultLabel: "→ 事業機会が広がる",
  });
  s.addText(
    "レビットの「マーケティング近視眼」：鉄道会社が自らを「鉄道業」と狭く定義したため自動車・航空機の台頭に対応できなかった、という有名な失敗例。",
    { x: diagX, y: 5.55, w: diagW, h: 0.85, fontFace: F_BODY, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 }
  );

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
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "例えばアの「自社ECサイトを立ち上げ、既存シリアル市場でのシェア拡大を狙う」は、製品・市場とも既存のままの施策＝実態は「市場浸透戦略」なのに、選択肢は「多角化戦略」と誤って名付けている。まず施策の実態を製品×市場で判定してから戦略名が一致するか照合するのがコツ。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.93;
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
    { k: "関連型多角化", v: "既存事業と技術・顧客・販売チャネル等で関連がある分野に進出（例：カメラメーカーが医療用内視鏡事業に進出）", gap: 0.6 },
    { k: "非関連型多角化", v: "既存事業とほとんど関連がない分野に進出（例：繊維会社が化粧品事業に進出）", gap: 0.5 },
  ], { fontSize: 11.5, labelW: 1.55 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "富士フイルムは写真フィルム事業で培った化学・材料技術を応用して化粧品事業（アンチエイジング）に進出した。技術的関連性が高い「関連型多角化」の代表例。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.93;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「新市場に新技術で進出する」＝多角化とは限らない。アンゾフの成長マトリクスの「新製品×新市場」が該当。新市場開拓・新製品開発と混同しないこと。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  drawVennOverlap(s, diagX, 2.2, diagW, 2.6, {
    leftLabel: "既存事業", rightLabel: "新規事業", overlapLabel: "シナジー",
  });
  s.addText(
    "「1+1が2でなく3以上になる効果」。技術・ブランド・販売網を複数事業で共有すると効率が良くなる。関連型多角化ほどシナジーを生みやすい。",
    { x: diagX, y: 5.55, w: diagW, h: 0.85, fontFace: F_BODY, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 }
  );

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
  {
    const proseX = 0.55, proseW = 7.4;
    const diagX = 8.25, diagW = 4.05;
    let cy = 1.85;
    s.addText(
      "市場成長率×市場シェアの2軸で事業を4分類し、資金配分を判断するフレームワーク。",
      { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 }
    );
    cy += 0.48;
    cy = addTermRows(s, proseX, cy, proseW, [
      { k: "花形", v: "稼ぐが投資も必要。将来の主力候補", gap: 0.4 },
      { k: "問題児", v: "将来性はあるが今は稼げない。育てるか撤退か判断", gap: 0.4 },
      { k: "金のなる木", v: "投資が要らず安定して稼ぐ。他事業の資金源", gap: 0.4 },
      { k: "負け犬", v: "稼げず将来性も薄い。撤退候補", gap: 0.4 },
    ], { fontSize: 11.5, labelW: 1.5 });
    cy += 0.1;
    s.addText([
      { text: "ひっかけ：", options: { bold: true, color: RED } },
      { text: "PPMは「市場の魅力度」と「自社の競争力」の2軸だけの単純化モデル。事業間のシナジーや社会的意義のある事業は考慮できない限界がある。", options: { color: RED } },
    ], { x: proseX, y: cy, w: proseW, h: 0.7, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

    drawQuadrant(s, diagX, 1.9, diagW, 3.9, {
      cells: [
        { pos: "tl", label: "花形" },
        { pos: "tr", label: "問題児" },
        { pos: "bl", label: "金のなる木" },
        { pos: "br", label: "負け犬" },
      ],
      axisCaption: "縦軸：市場成長率（上＝高い）／横軸：市場シェア（左＝高い）",
    });
  }
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-7 競争優位の戦略（各象限で有効な戦略の方向性が異なる）。",
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
  ], { x: proseX, y: cy, w: proseW, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.93;
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
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "大手飲料メーカーが自社にない発酵技術を持つスタートアップに出資し、将来の自社事業とのシナジーを狙うのがCVCの典型例。買収して経営統合するM&Aとは異なり、出資先の独立性を保ったまま関係を築く点がポイント。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
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
  {
    const proseX = 0.55, proseW = 7.4;
    const diagX = 8.25, diagW = 4.05;
    let cy = 1.85;
    s.addText(
      "業界の競争構造を5つの力（ポーター）に分解し、力が弱いほど儲けやすい業界と判断する。",
      { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 }
    );
    cy += 0.48;
    cy = addTermRows(s, proseX, cy, proseW, [
      { k: "新規参入の脅威", v: "参入障壁が低いほど大きい", gap: 0.4 },
      { k: "売り手の交渉力", v: "仕入先が強いと買い叩かれる", gap: 0.4 },
      { k: "買い手の交渉力", v: "顧客が強いと値下げ圧力を受ける", gap: 0.4 },
      { k: "代替品の脅威", v: "例：新聞紙にとってのニュースアプリ", gap: 0.4 },
    ], { fontSize: 11.5, labelW: 2.1 });
    cy += 0.1;
    s.addText([
      { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
      { text: "もう1つの分析軸がバリューチェーン（価値連鎖）分析：活動を「主活動」（購買物流→製造→出荷物流→販売→サービス）と「支援活動」に分け、付加価値の源泉を探る。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
    ], { x: proseX, y: cy, w: proseW, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

    draw5Forces(s, diagX, 1.9, diagW, 3.9, {
      center: "既存競合",
      top: "新規参入",
      bottom: "代替品",
      left: "売り手",
      right: "買い手",
    });
    s.addText(
      "5フォース分析（ポーター）。中心の「既存競合」が4方向から受ける圧力を分析する。",
      { x: diagX, y: 5.85, w: diagW, h: 0.55, fontFace: F_BODY, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 }
    );
  }
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
  {
    const proseX = 0.55, proseW = 7.4;
    const diagX = 8.25, diagW = 4.05;
    let cy = 1.85;
    s.addText(
      "「競争優位の種類」×「狙う市場の広さ」の組み合わせで、とるべき戦略が変わる（ポーター）。",
      { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 }
    );
    cy += 0.48;
    cy = addTermRows(s, proseX, cy, proseW, [
      { k: "コストリーダーシップ", v: "業界最低コストで価格競争に勝つ（例：格安航空会社）", gap: 0.55 },
      { k: "差別化", v: "品質・ブランド・デザインで独自の価値を出す（例：高級ホテル）", gap: 0.55 },
      { k: "集中戦略", v: "特定の狭い顧客層に絞ってコスト or 差別化で勝つ", gap: 0.4 },
    ], { fontSize: 11.5, labelW: 2.4 });
    cy += 0.1;
    s.addText([
      { text: "ひっかけ：", options: { bold: true, color: RED } },
      { text: "集中戦略は「誰に売るか」という別軸の話。両方を中途半端に狙うと収益性が落ちる（スタック・イン・ザ・ミドル）。", options: { color: RED } },
    ], { x: proseX, y: cy, w: proseW, h: 0.6, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

    drawQuadrant(s, diagX, 1.9, diagW, 3.9, {
      cells: [
        { pos: "tl", label: "コストリーダーシップ" },
        { pos: "tr", label: "差別化" },
        { pos: "bl", label: "コスト集中" },
        { pos: "br", label: "差別化集中" },
      ],
      axisCaption: "縦軸：狙う市場の広さ（上＝広い、下＝狭い）／横軸：優位の源泉（左＝コスト、右＝差別化）",
    });
  }
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-4 PPM（各象限で有効な戦略の方向性が異なる）。ユニクロのSPAは差別化的コストリーダーシップと解釈される。",
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
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "「精密加工技術」をコア技術に持つ工作機械メーカーが、医療機器・航空宇宙部品にも技術を横展開し、各分野で得た知見を工作機械自体の技術力向上にも還元するのがエの記述そのものの実践例。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
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
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "マウンテンバイクは、既存の自転車に飽き足らない愛好家（リード・ユーザー）自身が山道走行用に独自改造・開発したことから発展した、ユーザー・イノベーションの代表的な歴史的事例としてよく引用される。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
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

// ---------- Slide 28: 組織論 区切り ----------
addDividerSlide(pres, {
  ghostNo: "02",
  partNo: "PART 02",
  partLabel: "企業経営理論 ／ C-15〜C-29",
  title: "組織論",
  desc: "人と人とをどう組み合わせれば、会社としての力を最大化できるかを扱う分野。モチベーション理論・リーダーシップ理論など心理学的な要素を含む論点が多い。",
  chips: ["C-19 モチベーション理論", "C-24 組織変革", "C-25 労働関連法規"],
  notes: "組織論パートの区切りスライド。",
});

// ---------- Slide 29: C-15 組織の考え方 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-15 ／ 組織の考え方（組織目標、分業と調整、権限と責任）",
    title: "組織が機能するための3つの要素",
    overview: "共通目標のために複数人が協力する仕組みが組織。分業・調整・権限と責任の3つが土台になる。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "組織とは「共通の目標のために複数の人が協力する仕組み」。機能するには3つの要素が必要。",
    { x: proseX, y: cy, w: proseW, h: 0.3, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0 }
  );
  cy += 0.38;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "分業", v: "仕事を細分化し、担当者が専門的に担当する（効率化につながる）" },
    { k: "調整", v: "分業した仕事がバラバラにならないよう全体をまとめる（会議・報告ライン・ルール等）" },
    { k: "権限と責任", v: "決定できる範囲（権限）とその結果への責任は原則セットでなければならない" },
  ], { fontSize: 11.5, labelW: 1.7, gap: 0.4 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "飲食店チェーンでは各店舗の店長に「仕入れ量を決める権限」と同時に「食材ロスを一定以下に抑える責任」を負わせる。権限だけ与えて責任は本部が負う、では店長の判断が甘くなりがち。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-16 組織構造の形態（分業・調整の具体的な型）。",
    years: mkYears(new Set(["'16", "'17", "'19", "'20", "'21", "'22", "'24", "'25"])),
  });
}


// ---------- Slide 30: C-15 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-15 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第14問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "企業経営理論",
  });
  const cy = addExamQuestion(s, {
    stem: "組織における分業と調整に関する記述として、最も適切なものはどれか。",
    stemH: 1.0,
    choices: [
      { badge: "ア", text: "作業現場において、仕事を分業し、個々の作業範囲を特定の領域に狭く限定すると、作業者のスキルが均質化し、業務の幅が広がるため、キャリア形成の選択肢も増える。" },
      { badge: "イ", text: "作業手順を標準化し、作業内容を確定させることは、計画目標の達成率を高め、生産性向上につながる。このような標準化と計画目標の関係は、「計画のグレシャムの法則」として知られている。" },
      { badge: "ウ", text: "仕事の分業が過度に進むと、組織メンバーは自分の仕事が組織全体にどのような意味を持っているか実感できず、仕事への意欲が低下することがある。こうした現象は、「アンダーマイニング効果」として知られている。" },
      { badge: "エ", text: "仕事の分業を進めると、個々の作業が単純化され、機械化が容易となるため、各工程間の調整が不要となり、業務は効率化されやすい。" },
      { badge: "オ", text: "定型的な作業は標準化によってあらかじめ調整し、想定外の事態には上位層が事後的に対応することで、仕事は効率的に行われる。" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第14問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 31: C-15 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-15 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第14問）",
    overview: "正解はオ。定型業務は標準化、例外は上位層の判断に委ねる分業が効率的。",
    tag: "企業経営理論",
  });
  const choices = [
      { badge: "ア", text: "作業現場において、仕事を分業し、個々の作業範囲を特定の領域に狭く限定すると、作業者のスキルが均質化し、業務の幅が広がるため、キャリア形成の選択肢も増える。" },
      { badge: "イ", text: "作業手順を標準化し、作業内容を確定させることは、計画目標の達成率を高め、生産性向上につながる。このような標準化と計画目標の関係は、「計画のグレシャムの法則」として知られている。" },
      { badge: "ウ", text: "仕事の分業が過度に進むと、組織メンバーは自分の仕事が組織全体にどのような意味を持っているか実感できず、仕事への意欲が低下することがある。こうした現象は、「アンダーマイニング効果」として知られている。" },
      { badge: "エ", text: "仕事の分業を進めると、個々の作業が単純化され、機械化が容易となるため、各工程間の調整が不要となり、業務は効率化されやすい。" },
      { badge: "オ", text: "定型的な作業は標準化によってあらかじめ調整し、想定外の事態には上位層が事後的に対応することで、仕事は効率的に行われる。" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 4 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：オ", options: { bold: true, color: RED } },
    { text: "。定型業務は事前にルール・手順を標準化して調整し、例外・想定外の事態だけ上位層が個別に判断する役割分担が、分業と調整の基本的な考え方。ア：作業範囲を狭めると業務の幅はむしろ狭まり、キャリア選択肢が増えるとは言えない。イ：「計画のグレシャムの法則」は定型業務が非定型の計画業務を後回しにする現象を指し、標準化と生産性向上の説明とは無関係。ウ：「アンダーマイニング効果」は外的報酬が内発的動機づけを損なう現象で、過度な分業による疎外感とは別概念。エ：分業が進むほど工程間の調整の必要性はむしろ", options: {  } },
    { text: "高まる", options: { bold: true } },
    { text: "。", options: {  } },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.05, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.22 });
  cy += 1.05;
  s.addText([
    { text: "具体例\\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "組織の調整メカニズムには、標準化（ルール・手順）、階層による調整（上司の判断）、相互調整（現場同士の連絡）がある。定型業務は標準化、例外は階層に委ねる役割分担が効率的とされる。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.72;
  s.addText("関連知識：「計画のグレシャムの法則」（サイモン）と「アンダーマイニング効果」（過剰正当化効果）はこの設問での使われ方が誤りで、用語の正しい意味を知っていれば消去できる。", {
    x: 0.55, y: cy, w: 12.25, h: 0.5, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.50;
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第14問／正解：past_exams/1st_stage_answers/r07/2025c.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}
// ---------- Slide 32: C-16 組織構造の形態 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-16 ／ 組織構造の形態",
    title: "4つの組織形態とメリット・デメリット",
    overview: "職能別・事業部制・マトリックス・ネットワークの4形態で、専門性と連携のトレードオフが変わる。",
    tag: "企業経営理論",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "職能別組織", tag: "機能ごとに部門化", desc: "専門性が高まり効率的。ただし部門間連携が悪化しがちで全社視点の人材が育ちにくい" },
    { name: "事業部制組織", tag: "製品・地域ごと", desc: "事業ごとの意思決定が速く経営者人材が育ちやすい。ただし機能重複でコストがかさむ" },
    { name: "マトリックス組織", tag: "職能×事業の二重報告", desc: "柔軟に人材活用できる。ただし指揮命令系統が二重化し混乱を招きやすい" },
    { name: "ネットワーク組織", tag: "独立組織が緩やかに連携", desc: "柔軟性が高い。ただし統制が効きにくい" },
  ], { nameW: 2.3, tagW: 2.6 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：「組織は戦略に従う」（チャンドラー）が有名だが、逆に「戦略は組織に従う」という視点もある。どちらか一方が絶対ではない。",
    years: mkYears(new Set(["'16", "'17", "'20", "'21", "'22", "'23", "'25"])),
  });
}


// ---------- Slide 33: C-16 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-16 ／ 過去問で確認する",
    title: "こう出題される（令和5年度 第14問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "企業経営理論",
  });
  const cy = addExamQuestion(s, {
    stem: "主要な組織形態に関する記述として、最も適切なものはどれか。",
    stemH: 0.8,
    choices: [
      { badge: "ア", text: "機能別組織では、機能別部門の管理をそれぞれの部門の長に任せることから、事業部制組織よりも次世代経営者の育成を行いやすい。" },
      { badge: "イ", text: "機能別組織では、知識の蓄積が容易であるため、事業の内容や範囲にかかわらず経営者は意思決定を迅速に行いやすい。" },
      { badge: "ウ", text: "事業部制組織では、各事業部が自律的に判断できるために、事業部間で重複する投資が生じやすい。" },
      { badge: "エ", text: "事業部制組織では、各事業部が素早く有機的に連携できるため、機能別組織よりも事業横断的なシナジーを創出しやすい。" },
      { badge: "オ", text: "マトリックス組織は、複数の命令系統があることで組織運営が難しいため、不確実性が低い環境において採用されやすい。" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2023/C1JI2023.pdf（令和5年度第1次試験）第14問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 34: C-16 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-16 ／ 過去問で確認する",
    title: "解答＆解説（令和5年度 第14問）",
    overview: "正解はウ。事業部制は自律性が高い分、事業部間で機能・投資が重複しやすい。",
    tag: "企業経営理論",
  });
  const choices = [
      { badge: "ア", text: "機能別組織では、機能別部門の管理をそれぞれの部門の長に任せることから、事業部制組織よりも次世代経営者の育成を行いやすい。" },
      { badge: "イ", text: "機能別組織では、知識の蓄積が容易であるため、事業の内容や範囲にかかわらず経営者は意思決定を迅速に行いやすい。" },
      { badge: "ウ", text: "事業部制組織では、各事業部が自律的に判断できるために、事業部間で重複する投資が生じやすい。" },
      { badge: "エ", text: "事業部制組織では、各事業部が素早く有機的に連携できるため、機能別組織よりも事業横断的なシナジーを創出しやすい。" },
      { badge: "オ", text: "マトリックス組織は、複数の命令系統があることで組織運営が難しいため、不確実性が低い環境において採用されやすい。" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 2 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：ウ", options: { bold: true, color: RED } },
    { text: "。事業部制組織は各事業部が損益責任を持ち自律的に意思決定できる反面、営業・経理などの機能を事業部ごとに重複して抱えやすく、投資の重複が生じる。ア：機能別組織は部門長が専門機能しか管理せず、全社視点を持つ次世代経営者は", options: {  } },
    { text: "事業部制の方が", options: { bold: true } },
    { text: "育ちやすい。イ：機能別組織は事業の幅が広がるほど部門間調整の手間が増え、意思決定はむしろ遅くなる。エ：事業部制は独立採算で動くため、事業部間の連携（横断シナジー）はむしろ機能別組織より生まれにくい。オ：マトリックス組織は", options: {  } },
    { text: "不確実性の高い環境", options: { bold: true } },
    { text: "でこそ採用される。", options: {  } },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.05, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.22 });
  cy += 1.05;
  s.addText([
    { text: "具体例\\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "機能別組織と事業部制組織は、専門性の高さ vs 事業ごとの機動力、効率性 vs 重複コストという表裏の関係にある。どちらの型かでメリット・デメリットが逆になる点を対比で押さえておく。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.72;
  s.addText("関連知識：マトリックス組織は「ワン・マン・ツー・ボス問題」（指揮命令系統の二重化）を抱えるが、複雑で変化の速い事業環境に対応するために採用される。", {
    x: 0.55, y: cy, w: 12.25, h: 0.5, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.50;
  s.addText("出典：past_exams/1st_stage/1ji2023/C1JI2023.pdf（令和5年度第1次試験）第14問／正解：past_exams/1st_stage_answers/r05/2023c.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}
// ---------- Slide 35: C-17 組織文化 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-17 ／ 組織文化（レベル、経営戦略との関係、変革）",
    title: "「うちの会社らしいやり方」が組織文化",
    overview: "メンバーに共有された価値観・行動様式。統一の力にも変革の足かせにもなる。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText([
    { text: "組織文化", options: { bold: true, color: INK } },
    { text: "とは、組織のメンバーに共有されている価値観・行動様式。「うちの会社らしいやり方」のこと。目に見える「制度・慣習」から目に見えにくい「価値観」「基本的な前提」まで階層があるとされる（シャインの組織文化論）。", options: { color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.65, fontFace: F_BODY, fontSize: 11.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 });
  cy += 0.72;
  s.addShape("rect", { x: proseX, y: cy, w: proseW, h: 0.5, fill: { color: GHOST }, line: { type: "none" } });
  s.addShape("line", { x: proseX, y: cy, w: 0, h: 0.5, line: { color: INK, width: 2.5 } });
  s.addText("強い組織文化はメンバーの行動を統一し効率を高める一方、環境変化時に「今までのやり方」に固執させ、変革の足かせになることもある。", {
    x: proseX + 0.15, y: cy, w: proseW - 0.3, h: 0.5, valign: "middle",
    fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15,
  });
  cy += 0.6;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "「顧客第一主義」を掲げる会社で、社員全員がそれを当たり前として行動している状態は強い組織文化が根付いている例。ただし文化が強すぎるとデジタル化のような新しい変化への対応が遅れることもある。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-24 組織変革（強い組織文化は変革の解凍段階を難しくする要因になりうる）。",
    years: mkYears(new Set(["'17", "'24", "'25"])),
  });
}

// ---------- Slide 36: C-18 意思決定 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-18 ／ 組織における個人：意思決定",
    title: "人は「最適解」でなく「満足解」で意思決定する",
    overview: "人間の合理性には限界がある（サイモンの限定合理性）。認知バイアスの影響も受ける。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText([
    { text: "人間は経済学が想定するような「完全に合理的な意思決定」はできない（サイモンの", options: { color: INK } },
    { text: "限定合理性", options: { bold: true, color: INK } },
    { text: "）。情報・時間・認知能力に限界があるため、「最適な答え」でなく「まあ満足できる答え」で打ち切る（", options: { color: INK } },
    { text: "満足化原理", options: { bold: true, color: INK } },
    { text: "）。", options: { color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.75, fontFace: F_BODY, fontSize: 11.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 });
  cy += 0.82;
  s.addShape("rect", { x: proseX, y: cy, w: proseW, h: 0.5, fill: { color: GHOST }, line: { type: "none" } });
  s.addShape("line", { x: proseX, y: cy, w: 0, h: 0.5, line: { color: INK, width: 2.5 } });
  s.addText([
    { text: "人間の判断は", options: {} },
    { text: "認知バイアス", options: { bold: true } },
    { text: "（思考の偏り）の影響も受ける。例：一度決めたことに固執する、都合の良い情報ばかり集める、など。", options: {} },
  ], { x: proseX + 0.15, y: cy, w: proseW - 0.3, h: 0.5, valign: "middle", fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  cy += 0.6;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "新商品の企画会議で、本来は市場調査をもっと徹底すべきなのに、時間的制約から「これくらいで十分だろう」と判断してしまうのは限定合理性・満足化原理の表れ。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "B", rankLabel: "頻出論点",
    related: "関連：C-22 集団の心理（グループシンクも集団レベルの非合理な意思決定の一種）。",
    years: mkYears(new Set(["'19", "'24"])),
  });
}

// ---------- Slide 37: C-19 モチベーション理論 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-19 ／ 組織における個人：モチベーション理論",
    title: "社員のやる気を説明する4つの理論",
    overview: "提唱者名と理論の核心をセットで覚える。10年間ほぼ毎年出題される最頻出論点。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "社員のやる気を高める要因を説明する代表理論4つ。",
    { x: proseX, y: cy, w: proseW, h: 0.3, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0 }
  );
  cy += 0.38;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "マズロー", v: "欲求段階説：生理的→安全→社会的→承認→自己実現の5段階。低次が満たされると次の欲求が現れる", gap: 0.4 },
    { k: "ハーズバーグ", v: "動機づけ・衛生理論：満足に関わる要因（達成感等）と不満足に関わる要因（給与等）は別物", gap: 0.4 },
    { k: "期待理論", v: "（ブルーム）モチベーション＝期待×手段性×誘意性の掛け算で決まる", gap: 0.35 },
    { k: "目標設定理論", v: "（ロック）具体的でやや高難度な目標が、曖昧・簡単すぎる目標よりやる気を高める", gap: 0.4 },
  ], { fontSize: 11, labelW: 1.7 });
  cy += 0.05;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "ハーズバーグの動機づけ要因・衛生要因の区別が特に頻出。給与を上げても不満は減るが積極的な「やる気」は上がらない（給与は衛生要因）。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.45, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "具体例：給与を大幅に上げても意欲が変わらないのは、給与が「衛生要因」で不満防止にとどまるため。やる気には達成感等「動機づけ要因」への働きかけが必要。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}


// ---------- Slide 38: C-19 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-19 ／ 過去問で確認する",
    title: "こう出題される（令和5年度 第17問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "企業経営理論",
  });
  const cy = addExamQuestion(s, {
    stem: "E. A. ロックとG. P. レイサムらが提唱した目標設定理論に則した管理者の判断と行動に関する記述として、最も適切なものはどれか。",
    stemH: 1.1,
    choices: [
      { badge: "ア", text: "自分には目標を達成できる能力があるという信念を持つ人ほど、達成が困難な状況になると目標を断念する傾向があるため、自分の能力を過信しないように部下に伝えた。" },
      { badge: "イ", text: "達成に多くの努力を要する目標は、達成できる見込みが立てづらく部下からの反発や抵抗が予想されるため、容易に達成できる業績目標を設定した。" },
      { badge: "ウ", text: "達成の難易度が高い目標を設定するにあたっては、部下にその目標を受容させることが重要であるため、その目標が公正で妥当であることを強調して部下に伝えた。" },
      { badge: "エ", text: "一人ひとりの目標の内容が職場で公表されると、目標に対するコミットメントが阻害されるため、各自の目標が互いに知られることのないように配慮した。" },
      { badge: "オ", text: "明確な数値目標を設定すると、目標達成に対する心理的プレッシャーが高まり、部下の達成意欲が低下するため、自由に解釈できる定性的な目標を設定した。" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2023/C1JI2023.pdf（令和5年度第1次試験）第17問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 39: C-19 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-19 ／ 過去問で確認する",
    title: "解答＆解説（令和5年度 第17問）",
    overview: "正解はウ。困難な目標も「公正・妥当」と伝え受容させることがモチベーションを高める。",
    tag: "企業経営理論",
  });
  const choices = [
      { badge: "ア", text: "自分には目標を達成できる能力があるという信念を持つ人ほど、達成が困難な状況になると目標を断念する傾向があるため、自分の能力を過信しないように部下に伝えた。" },
      { badge: "イ", text: "達成に多くの努力を要する目標は、達成できる見込みが立てづらく部下からの反発や抵抗が予想されるため、容易に達成できる業績目標を設定した。" },
      { badge: "ウ", text: "達成の難易度が高い目標を設定するにあたっては、部下にその目標を受容させることが重要であるため、その目標が公正で妥当であることを強調して部下に伝えた。" },
      { badge: "エ", text: "一人ひとりの目標の内容が職場で公表されると、目標に対するコミットメントが阻害されるため、各自の目標が互いに知られることのないように配慮した。" },
      { badge: "オ", text: "明確な数値目標を設定すると、目標達成に対する心理的プレッシャーが高まり、部下の達成意欲が低下するため、自由に解釈できる定性的な目標を設定した。" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 2 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：ウ", options: { bold: true, color: RED } },
    { text: "。目標設定理論では、具体的で困難な目標ほどモチベーションを高めるが、それには部下がその目標を自分のものとして受け入れる「目標受容」が前提となる。目標の公正さ・妥当性を伝えて納得感を高めることは目標受容を促す適切な行動。ア：自己効力感が高い人ほど困難な状況でも粘り強く取り組む傾向があり、逆の説明。イ：具体的で困難な目標の方が簡単な目標より高い成果を引き出すのが理論の核心。エ：目標の公表は一般にコミットメントを", options: {  } },
    { text: "高める", options: { bold: true } },
    { text: "（社会的公約の効果）。オ：明確な数値目標の方が努力の方向づけが明確になりモチベーションを", options: {  } },
    { text: "高める", options: { bold: true } },
    { text: "。", options: {  } },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.05, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.22 });
  cy += 1.05;
  s.addText([
    { text: "具体例\\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "C-19の各理論は「提唱者名」と「核心的な主張」をセットで押さえ、混同しないことが頻出対策のポイント（マズロー、ハーズバーグ、期待理論、目標設定理論）。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.72;
  s.addText("関連知識：目標設定理論のポイントは「①具体性」「②適度な困難さ」に加え、「③目標受容」と「④フィードバック」が成果に影響する点。", {
    x: 0.55, y: cy, w: 12.25, h: 0.5, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.50;
  s.addText("出典：past_exams/1st_stage/1ji2023/C1JI2023.pdf（令和5年度第1次試験）第17問／正解：past_exams/1st_stage_answers/r05/2023c.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}
// ---------- Slide 40: C-20 リーダーシップ理論 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-20 ／ リーダーシップ理論",
    title: "「行動の型」と「状況対応」を区別する",
    overview: "PM理論はリーダーの行動の型、SL理論は部下の状況に応じたスタイル変化を説明する。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "リーダーシップに関する代表理論。「行動の型」を分類する理論と「状況で変える」理論を区別する。",
    { x: proseX, y: cy, w: proseW, h: 0.3, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0 }
  );
  cy += 0.38;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "PM理論", v: "（三隅二不二）P機能（目標達成）とM機能（集団維持）の2軸。両方高い「PM型」が理想" },
    { k: "SL理論", v: "（ハーシー＆ブランチャード）部下の成熟度に応じ指示型→説得型→参加型→委任型へスタイルを変える", gap: 0.5 },
    { k: "変革型リーダーシップ", v: "内発的動機づけで大きな変革を推進（対義語：アメとムチの交換型リーダーシップ）", gap: 0.45 },
  ], { fontSize: 11.5, labelW: 2.1 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "新入社員には具体的に指示する「指示型」、ベテランには仕事を任せる「委任型」が適切、というのがSL理論の考え方そのもの。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：PM理論は「行動の型」の分類、SL理論は「部下の状況に応じたスタイル変化」という違いを混同しないこと。",
    years: mkYears(new Set(["'18", "'19", "'21", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 41: C-21 パワーとコンフリクト ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-21 ／ パワーとコンフリクト",
    title: "パワーの源泉は役職だけではない",
    overview: "公式の権限以外にも複数のパワーの源泉があり、適度なコンフリクトは組織の活性化にもなる。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText([
    { text: "組織内では上下関係・部門間で「パワー（影響力）」の行使や意見対立（", options: { color: INK } },
    { text: "コンフリクト", options: { bold: true, color: INK } },
    { text: "）が生じる。", options: { color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.3, fontFace: F_BODY, fontSize: 11.5, isTextBox: true, margin: 0 });
  cy += 0.38;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "公式の権限", v: "地位に基づくパワー" },
    { k: "専門性パワー", v: "専門知識に基づくパワー" },
    { k: "準拠パワー", v: "人望に基づくパワー（以上、フレンチ＆レイヴンの分類）" },
  ], { fontSize: 11.5, labelW: 1.7, gap: 0.35 });
  cy += 0.05;
  s.addShape("rect", { x: proseX, y: cy, w: proseW, h: 0.45, fill: { color: GHOST }, line: { type: "none" } });
  s.addShape("line", { x: proseX, y: cy, w: 0, h: 0.45, line: { color: INK, width: 2.5 } });
  s.addText("コンフリクトは必ずしも悪ではない。適度なコンフリクトは新しい視点をもたらし組織を活性化するが、過度だと機能不全、皆無だと停滞した組織になりがち。", {
    x: proseX + 0.15, y: cy, w: proseW - 0.3, h: 0.45, valign: "middle",
    fontFace: F_BODY, fontSize: 10, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.1,
  });
  cy += 0.53;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "新人でもその分野で誰より詳しい専門知識があれば、周囲から意見を求められる実質的な影響力（専門性パワー）を持つことがある。役職に基づく公式パワーとは別の源泉。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-20 リーダーシップ理論（パワーの使い方とリーダーシップスタイルは密接に関連）。",
    years: mkYears(new Set(["'18", "'19", "'21", "'22", "'24"])),
  });
}


// ---------- Slide 42: C-21 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-21 ／ 過去問で確認する",
    title: "こう出題される（令和6年度 第20問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "企業経営理論",
  });
  const cy = addExamQuestion(s, {
    stem: "組織や集団においては、意見の相違や利害の不一致から、個人間でコンフリクトが発生することが一般的である。コンフリクトへの対処は、自己の利益を追求する度合いと、相手の利益追求を許容し協力する度合いとの組み合わせに応じて、「回避」、「競争」、「協調」、「妥協」、「適応」の5類型に分類される。コンフリクトへの対処に関する記述として、最も適切なものはどれか。",
    stemH: 1.85,
    choices: [
      { badge: "ア", text: "「回避」とは、自己の利益を強く主張しない一方で相手の利益もあまり許容できない場合に、問題解決を延期して様子を見るという対処である。互いの対立点が表立つのを避けたい場合にとられやすい。" },
      { badge: "イ", text: "「競争」とは、相手の利益を最大限に許容しつつ、相手に命令したり相手を説得したりすることで自己の利益も追求するという対処である。権力志向的で高い職位の人間から、順応的な低い職位の人間に対してとられやすい。" },
      { badge: "ウ", text: "「協調」とは、双方がある程度の利益を獲得しつつ互いに犠牲も払うという対処である。互いの対立点を曖昧にすることでコンフリクトを自然に解消しようとする場合にとられやすい。" },
      { badge: "エ", text: "「妥協」とは、当事者の一方のみが自己の利益を犠牲にして相手の利益を最大限に許容するという対処である。互いにある程度の利益をとりつつ犠牲も払うという折り合いがつけられない場合にとられやすい。" },
      { badge: "オ", text: "「適応」とは、相手の利益を犠牲にして自己の利益を追求するという対処である。自己の利益を一方的に追求することで、相手との長期的な関係が損なわれても問題ないと判断される場合にとられやすい。" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2024/C1JI2024.pdf（令和6年度第1次試験）第20問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 43: C-21 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-21 ／ 過去問で確認する",
    title: "解答＆解説（令和6年度 第20問）",
    overview: "正解はア。自己主張も協力度もどちらも低い組み合わせが「回避」。",
    tag: "企業経営理論",
  });
  const choices = [
      { badge: "ア", text: "「回避」とは、自己の利益を強く主張しない一方で相手の利益もあまり許容できない場合に、問題解決を延期して様子を見るという対処である。互いの対立点が表立つのを避けたい場合にとられやすい。" },
      { badge: "イ", text: "「競争」とは、相手の利益を最大限に許容しつつ、相手に命令したり相手を説得したりすることで自己の利益も追求するという対処である。権力志向的で高い職位の人間から、順応的な低い職位の人間に対してとられやすい。" },
      { badge: "ウ", text: "「協調」とは、双方がある程度の利益を獲得しつつ互いに犠牲も払うという対処である。互いの対立点を曖昧にすることでコンフリクトを自然に解消しようとする場合にとられやすい。" },
      { badge: "エ", text: "「妥協」とは、当事者の一方のみが自己の利益を犠牲にして相手の利益を最大限に許容するという対処である。互いにある程度の利益をとりつつ犠牲も払うという折り合いがつけられない場合にとられやすい。" },
      { badge: "オ", text: "「適応」とは、相手の利益を犠牲にして自己の利益を追求するという対処である。自己の利益を一方的に追求することで、相手との長期的な関係が損なわれても問題ないと判断される場合にとられやすい。" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 0 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：ア", options: { bold: true, color: RED } },
    { text: "。5類型は「自己の利益追求の強さ」×「相手への協力度」の2軸で整理される（トーマス＝キルマン・モデル）。自己主張・協力度がともに低いのが「回避」で、対立を先送りし様子見する対処という説明は定義どおり。イ：「相手の利益を最大限に許容」は協力度が高いことを意味し、協力度が低い「競争」の定義と矛盾。ウ・エ・オ：説明文と用語の対応が入れ替わっている（ウの説明は「妥協」、エの説明は「適応」、オの説明は「競争」の定義に近い）。", options: {  } },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.05, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.22 });
  cy += 1.05;
  s.addText([
    { text: "具体例\\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "ウ・エ・オで「用語」と「説明文」が入れ替わって対応づけられている点に気づけるかがポイント。5類型を「自己主張×協力度」の2軸マトリクスで整理して覚えておくと素早く見抜ける。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.72;
  s.addText("関連知識：フレンチ＆レイヴンのパワーの分類（公式の権限・専門性パワー・準拠パワーなど）とあわせて、組織内の対人関係の理論として整理しておくこと。", {
    x: 0.55, y: cy, w: 12.25, h: 0.5, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.50;
  s.addText("出典：past_exams/1st_stage/1ji2024/C1JI2024.pdf（令和6年度第1次試験）第20問／正解：past_exams/1st_stage_answers/r06/2024c.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}
// ---------- Slide 44: C-22 集団の心理 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-22 ／ 集団の心理",
    title: "集団は個人とは違う心理現象を起こす",
    overview: "グループシンク・グループシフト・心理的安全性という3つの集団心理現象を区別する。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "集団で意思決定すると、個人では起きない特有の心理現象が生じることがある。",
    { x: proseX, y: cy, w: proseW, h: 0.3, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0 }
  );
  cy += 0.38;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "グループシンク", v: "団結・和を重視するあまり少数意見が出せず、不合理な結論に至る現象（集団浅慮）", gap: 0.4 },
    { k: "グループシフト", v: "議論の結果、個人で考えるより極端な結論（リスク志向 or 慎重）に振れる現象", gap: 0.4 },
    { k: "心理的安全性", v: "意見や懸念を安心して発言できる状態。生産性の高いチームの共通要因として注目", gap: 0.4 },
  ], { fontSize: 11.5, labelW: 1.9 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "会議で上司の意見に誰も異を唱えず危険な決定が承認されるのはグループシンクの典型。逆に率直に懸念を言い合える雰囲気（心理的安全性）があれば、こうした失敗を防ぎやすい。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-18 意思決定（限定合理性・認知バイアスは個人レベル、本論点は集団レベルの非合理性）。",
    years: mkYears(new Set(["'18", "'20", "'21", "'23"])),
  });
}

// ---------- Slide 45: C-23 組織間関係 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-23 ／ 組織間関係：資源依存、取引コスト、クラスター",
    title: "会社は他組織との関係の中で活動する",
    overview: "資源依存・取引コスト・産業クラスターという3つの理論で組織間関係を説明する。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "会社は単独で存在せず、他の組織（取引先・競合・行政等）との関係の中で活動する。",
    { x: proseX, y: cy, w: proseW, h: 0.3, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0 }
  );
  cy += 0.38;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "資源依存理論", v: "不足する経営資源を外部に依存し、その依存関係がパワー関係を生む（強く依存する相手には強く出られない）", gap: 0.5 },
    { k: "取引コスト理論", v: "市場で取引するか自社内で行うかは、交渉・契約・監視等の「取引コスト」を比較して決まる", gap: 0.4 },
    { k: "産業クラスター", v: "特定地域に関連企業・大学・研究機関が集積し情報交換・イノベーションが起きやすくなる（例：シリコンバレー）", gap: 0.5 },
  ], { fontSize: 11.5, labelW: 1.9 });
  cy += 0.08;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "特定部品の他社との取引コスト（品質管理・納期調整の手間）が高すぎる場合、自社で内製化（垂直統合）した方が効率的、と判断されることがある（取引コスト理論）。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-5 垂直統合・M&A・戦略的提携（取引コスト理論は垂直統合の是非を判断する理論的裏付け）。",
    years: mkYears(new Set(["'19", "'21", "'22", "'23", "'24", "'25"])),
  });
}


// ---------- Slide 46: C-23 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-23 ／ 過去問で確認する",
    title: "こう出題される（令和6年度 第21問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "企業経営理論",
  });
  const cy = addExamQuestion(s, {
    stem: "組織間関係や組織間ネットワークに関する記述として、最も適切なものはどれか。",
    stemH: 0.7,
    choices: [
      { badge: "ア", text: "「埋め込まれた紐帯」では機会主義的行動が生じやすいため、組織間ネットワークにおける「埋め込まれた紐帯」の比率を減らすことが望ましい。" },
      { badge: "イ", text: "「埋め込まれた紐帯」で結ばれた組織間ネットワークでは、暗黙的な知識の移転が促進されやすい。" },
      { badge: "ウ", text: "「弱い紐帯の強み」を最大限享受しようとすれば、関係を取り結ぶ組織を絞り込み、弱い紐帯を強い紐帯に転換することが不可欠である。" },
      { badge: "エ", text: "組織にとって新奇性の高い知識をより獲得するためには、これまでに築いてきた組織との紐帯をいっそう強めることが望ましい。" },
      { badge: "オ", text: "他の組織からの影響を極力排除するためには、「埋め込まれた紐帯」のみによって構成される組織間ネットワークを構築することが望ましい。" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2024/C1JI2024.pdf（令和6年度第1次試験）第21問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 47: C-23 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-23 ／ 過去問で確認する",
    title: "解答＆解説（令和6年度 第21問）",
    overview: "正解はイ。信頼関係に基づく「埋め込まれた紐帯」は暗黙知の移転を促す。",
    tag: "企業経営理論",
  });
  const choices = [
      { badge: "ア", text: "「埋め込まれた紐帯」では機会主義的行動が生じやすいため、組織間ネットワークにおける「埋め込まれた紐帯」の比率を減らすことが望ましい。" },
      { badge: "イ", text: "「埋め込まれた紐帯」で結ばれた組織間ネットワークでは、暗黙的な知識の移転が促進されやすい。" },
      { badge: "ウ", text: "「弱い紐帯の強み」を最大限享受しようとすれば、関係を取り結ぶ組織を絞り込み、弱い紐帯を強い紐帯に転換することが不可欠である。" },
      { badge: "エ", text: "組織にとって新奇性の高い知識をより獲得するためには、これまでに築いてきた組織との紐帯をいっそう強めることが望ましい。" },
      { badge: "オ", text: "他の組織からの影響を極力排除するためには、「埋め込まれた紐帯」のみによって構成される組織間ネットワークを構築することが望ましい。" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 1 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：イ", options: { bold: true, color: RED } },
    { text: "。「埋め込まれた紐帯」（信頼関係に基づく緊密な取引関係）で結ばれた組織間では、頻繁で深いコミュニケーションが行われるため、言語化しにくい暗黙知の共有・移転が進みやすい。ア：埋め込まれた紐帯は信頼関係に基づくため、むしろ機会主義的行動を", options: {  } },
    { text: "抑制する", options: { bold: true } },
    { text: "。ウ：「弱い紐帯の強み」は疎遠で多様な相手とのつながりが新しい情報をもたらす考え方で、関係先を絞ることはこの強みを失わせる。エ：新奇性の高い知識は", options: {  } },
    { text: "弱い紐帯", options: { bold: true } },
    { text: "を通じてもたらされやすい。オ：埋め込まれた紐帯のみのネットワークは特定の関係先への依存が偏り、新しい情報を取りこぼすリスクがある。", options: {  } },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.05, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.22 });
  cy += 1.05;
  s.addText([
    { text: "具体例\\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "資源依存理論・取引コスト理論とあわせて、組織間関係論の代表的な視点を整理しておく。強い紐帯＝信頼・暗黙知の共有、弱い紐帯＝新しい情報の獲得、という対比が頻出。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.72;
  s.addText("関連知識：グラノヴェッターの「弱い紐帯の強み」と埋め込み理論は対比で出題されやすい。取引コスト理論（内製化 vs 市場取引の選択）ともセットで押さえておくこと。", {
    x: 0.55, y: cy, w: 12.25, h: 0.5, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.50;
  s.addText("出典：past_exams/1st_stage/1ji2024/C1JI2024.pdf（令和6年度第1次試験）第21問／正解：past_exams/1st_stage_answers/r06/2024c.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}
// ---------- Slide 48: C-24 組織変革と組織成長 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-24 ／ 組織変革と組織成長",
    title: "レヴィンの3段階モデルで変革を進める",
    overview: "組織は創業期→成長期→成熟期とライフサイクルをたどる。変革は3段階で進める。",
    tag: "企業経営理論",
  });
  {
    const proseX = 0.55, proseW = 7.4;
    const diagX = 8.25, diagW = 3.4;
    let cy = 1.85;
    s.addText(
      "組織変革はレヴィンの3段階モデルで進める。組織はライフサイクル（創業期→成長期→成熟期→再生・衰退期）もたどる。",
      { x: proseX, y: cy, w: proseW, h: 0.55, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
    );
    cy += 0.63;
    cy = addTermRows(s, proseX, cy, proseW, [
      { k: "解凍", v: "現状維持の意識を崩し、変化の必要性を認識させる", gap: 0.4 },
      { k: "変化", v: "新しいやり方を実際に導入する", gap: 0.4 },
      { k: "再凍結", v: "新しいやり方を定着させ、元に戻らないようにする", gap: 0.4 },
    ], { fontSize: 11.5, labelW: 1.3 });
    cy += 0.1;
    s.addText([
      { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
      { text: "新システム導入時、いきなり入れ替えず「今のままではまずい」と危機感を共有し（解凍）、試験導入し（変化）、新業務フローとして定着させる（再凍結）と成功しやすい。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
    ], { x: proseX, y: cy, w: proseW, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

    drawProcessStepsVertical(s, diagX, 1.9, diagW, 4.5, [
      { num: "STEP 1", label: "解凍" },
      { num: "STEP 2", label: "変化" },
      { num: "STEP 3", label: "再凍結" },
    ]);
  }
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：組織のライフサイクルでは各段階の経営課題（資金調達→仕組み化→硬直化対応）が異なる。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}


// ---------- Slide 49: C-24 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-24 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第23問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "企業経営理論",
  });
  const cy = addExamQuestion(s, {
    stem: "組織における個人が組織変革に抵抗を示す理由に関する記述として、最も不適切なものはどれか。",
    stemH: 0.75,
    choices: [
      { badge: "ア", text: "外部環境が変化しているのに、組織内の慣習が従来の安定的かつ効率的な仕事の進め方を維持し強化していることに対して、危機感を抱くから。" },
      { badge: "イ", text: "組織変革によって、従来よりも仕事の成果がうまく出せなくなることを心配するから。" },
      { badge: "ウ", text: "組織変革によって、新しい仕事のやり方が少なくとも短期的には従来よりも非効率になると感じるから。" },
      { badge: "エ", text: "組織変革によって、自らの職務にとって未知で不確実な状態がもたらされることに不安を感じるから。" },
      { badge: "オ", text: "たとえ客観的に外部環境が変化したとしても、偏った情報を収集したり、情報を偏って解釈したりすることで、従来と同じ環境が現在も継続していると考えるから。" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第23問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 50: C-24 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-24 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第23問）",
    overview: "正解はア。「危機感を抱く」ことは変革を後押しする心理で、抵抗理由にはならない（最も不適切）。",
    tag: "企業経営理論",
  });
  const choices = [
      { badge: "ア", text: "外部環境が変化しているのに、組織内の慣習が従来の安定的かつ効率的な仕事の進め方を維持し強化していることに対して、危機感を抱くから。" },
      { badge: "イ", text: "組織変革によって、従来よりも仕事の成果がうまく出せなくなることを心配するから。" },
      { badge: "ウ", text: "組織変革によって、新しい仕事のやり方が少なくとも短期的には従来よりも非効率になると感じるから。" },
      { badge: "エ", text: "組織変革によって、自らの職務にとって未知で不確実な状態がもたらされることに不安を感じるから。" },
      { badge: "オ", text: "たとえ客観的に外部環境が変化したとしても、偏った情報を収集したり、情報を偏って解釈したりすることで、従来と同じ環境が現在も継続していると考えるから。" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 0 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：ア", options: { bold: true, color: RED } },
    { text: "（最も不適切）。外部環境の変化と従来のやり方とのズレに「危機感を抱く」ことは、レヴィンの3段階モデルの「解凍」――変化の必要性を認識する段階――そのものであり、変革を", options: {  } },
    { text: "後押しする", options: { bold: true } },
    { text: "心理であって抵抗の理由にはならない。イ〜オはいずれも、成果への不安・短期的非効率さ・未知への不安・認知バイアスという、変革抵抗の典型的な理由として適切。", options: {  } },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.05, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.22 });
  cy += 1.05;
  s.addText([
    { text: "具体例\\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "レヴィンの3段階モデル（解凍→変化→再凍結）の「解凍」は変革に<b>抵抗しない</b>方向の心理（危機感の共有）。イ〜オはすべて「なぜ人は変わりたくないと感じるか」という抵抗側の心理を説明しており、アだけ論理の向きが逆。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.72;
  s.addText("関連知識：「最も不適切なものはどれか」型の設問では、他の選択肢群と論理の向きが逆になっている1つを探すのが定石。", {
    x: 0.55, y: cy, w: 12.25, h: 0.5, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.50;
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第23問／正解：past_exams/1st_stage_answers/r07/2025c.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}
// ---------- Slide 51: C-25 労働関連法規 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-25 ／ 人的資源管理：労働関連法規",
    title: "押さえるべき7つの労働関連法規",
    overview: "法律名と規制内容を正確に対応させる。適用対象・義務か努力義務かの違いに注意。",
    tag: "企業経営理論",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "労働基準法", desc: "労働時間・休憩・休日・賃金支払いなど労働条件の最低基準" },
    { name: "労働契約法", desc: "労働契約の締結・変更・終了に関するルール" },
    { name: "男女雇用機会均等法", desc: "性別を理由とする差別的取扱いの禁止" },
    { name: "育児・介護休業法", desc: "育児・介護のための休業制度" },
    { name: "パート・有期雇用労働法", desc: "正社員と非正規社員との不合理な待遇差の禁止（同一労働同一賃金）" },
    { name: "労働安全衛生法", desc: "労働者の安全と健康の確保" },
    { name: "労働者派遣法", desc: "派遣労働者の保護、派遣期間の制限など" },
  ], { rowH: 0.6, nameW: 3.3 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "具体例：正社員と契約社員とで基本給・賞与に不合理な差をつけるのは、パート・有期雇用労働法（同一労働同一賃金）に抵触するおそれがある。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 52: C-26 雇用管理 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-26 ／ 人的資源管理：雇用管理",
    title: "ジョブローテーションでゼネラリストを育てる",
    overview: "日本企業に特徴的な定期的な人事異動の仕組みと、その狙いを理解する。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText([
    { text: "人材をどう採用し配置するかという管理分野。日本企業に特徴的な仕組みが", options: { color: INK } },
    { text: "ジョブローテーション", options: { bold: true, color: INK } },
    { text: "（定期的な人事異動）。", options: { color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 11.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.48;
  s.addShape("rect", { x: proseX, y: cy, w: proseW, h: 0.65, fill: { color: GHOST }, line: { type: "none" } });
  s.addShape("line", { x: proseX, y: cy, w: 0, h: 0.65, line: { color: INK, width: 2.5 } });
  s.addText([
    { text: "特定の職務を限定せず様々な部署を経験させ、会社全体を理解した", options: {} },
    { text: "ゼネラリスト", options: { bold: true } },
    { text: "を育成する狙いがある。対比されるのが欧米型の「ジョブ型雇用」（特定の職務・スキルを明確に定義して採用）。", options: {} },
  ], { x: proseX + 0.15, y: cy, w: proseW - 0.3, h: 0.65, valign: "middle", fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.73;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "入社した社員を営業部・経理部・製造部と数年おきに異動させ、幅広い業務知識を身につけさせるのは、日本型雇用に特徴的なジョブローテーションの一例。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-29 戦略的人的資源管理（採用・配置方針も経営戦略と整合させるべきという発想）。",
    years: mkYears(new Set(["'16", "'17", "'25"])),
  });
}

// ---------- Slide 53: C-27 評価・処遇 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-27 ／ 人的資源管理：評価・処遇",
    title: "3つの賃金体系とハロー効果に注意する",
    overview: "年功給・職務給・職能給の違いと、人事評価をゆがめるハロー効果を理解する。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "社員の働きをどう評価し賃金に反映するか。代表的な賃金体系は3つ。",
    { x: proseX, y: cy, w: proseW, h: 0.3, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0 }
  );
  cy += 0.38;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "年功給", v: "勤続年数・年齢に応じて賃金が上がる" },
    { k: "職務給", v: "職務の難易度・責任に応じて賃金を決める" },
    { k: "職能給", v: "個人の持つ能力（職能）に応じて賃金を決める" },
  ], { fontSize: 11.5, labelW: 1.4, gap: 0.35 });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "日本企業は伝統的に年功給・職能給中心だが、近年は成果や職務を重視する動きも進む。人事評価では評価者の主観による歪み（", options: { color: RED } },
    { text: "ハロー効果", options: { bold: true, color: RED } },
    { text: "：1つの優れた点に引きずられ全体を過大評価）にも注意。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.55, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.63;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "「彼は有名大学出身だから仕事もできるはずだ」という先入観で高評価をつけるのはハロー効果の典型例で、公正な人事評価を妨げる。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.55, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-29 戦略的人的資源管理（評価制度も経営戦略のタイプと整合させるべき）。",
    years: mkYears(new Set(["'16", "'18", "'19", "'20", "'22", "'24"])),
  });
}

// ---------- Slide 54: C-28 人材育成 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-28 ／ 人的資源管理：人材育成",
    title: "OJT・Off-JT・メンタリングの3手法",
    overview: "実務を通じて学ぶか、業務を離れて学ぶか、先輩が継続支援するかで手法が変わる。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "人材育成の代表的な3手法。",
    { x: proseX, y: cy, w: proseW, h: 0.3, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0 }
  );
  cy += 0.38;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "OJT", v: "On the Job Training。実際の仕事を通じて学ぶ" },
    { k: "Off-JT", v: "Off the Job Training。業務を離れて研修などで学ぶ" },
    { k: "メンタリング", v: "経験豊富な先輩社員が若手を継続的に支援する" },
  ], { fontSize: 11.5, labelW: 1.6, gap: 0.35 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "新人が先輩社員に同行して営業の実務を学ぶのはOJT、外部の研修会社が実施する集合研修に参加するのはOff-JT。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-26 雇用管理（ジョブローテーションもOJTの一形態と捉えられる）。",
    years: mkYears(new Set(["'17", "'18", "'20"])),
  });
}

// ---------- Slide 55: C-29 戦略的人的資源管理 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-29 ／ 人的資源管理：戦略的人的資源管理",
    title: "人事施策を経営戦略と一貫させる（SHRM）",
    overview: "採用・評価・育成・報酬を場当たり的でなく、経営戦略と整合させて設計する考え方。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText([
    { text: "戦略的人的資源管理（SHRM）", options: { bold: true, color: INK } },
    { text: "：人事施策（採用・評価・育成・報酬等）を、会社全体の経営戦略と一貫性を持たせて設計する考え方。「どんな戦略を取るかで、必要な人材像・評価基準・育成方針も変わるべき」という発想。", options: { color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.75, fontFace: F_BODY, fontSize: 11.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 });
  cy += 0.85;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "コストリーダーシップ戦略（C-7参照）の会社では効率性・生産性重視の評価制度が合理的だが、差別化戦略の会社では創造性・イノベーションを促す評価制度の方が戦略と整合的。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "C", rankLabel: "出題実績あり",
    related: "関連：C-7 競争優位の戦略（採用すべき人事施策は戦略のタイプによって変わる）。出題頻度は低いが基本概念として押さえておく。",
    years: mkYears(new Set(["'16"])),
  });
}

// ---------- Slide 56: マーケティング論 区切り ----------
addDividerSlide(pres, {
  ghostNo: "03",
  partNo: "PART 03",
  partLabel: "企業経営理論 ／ C-30〜C-42",
  title: "マーケティング論",
  desc: "「誰に」「何を」「どう」届けるかを設計する分野。STP・4Pという基本フレームワークに、CRM・デジタル・ブランディングなど現代的なテーマが加わる、最頻出（A）論点が並ぶ分野。",
  chips: ["C-31 STP・4P", "C-37 ブランディング", "C-38 消費者行動"],
  notes: "マーケティング論パートの区切りスライド。",
});

// ---------- Slide 57: C-30 マーケティングの基礎概念 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-30 ／ マーケティングの基礎概念",
    title: "マーケティングとは「顧客視点」で考えること",
    overview: "「作ったものを売る」販売志向から「求められるものを作る」顧客志向への転換が出発点。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "マーケティングとは、単に「モノを売る」ことではなく、顧客が本当に求めている価値を理解し、それを届ける仕組みを作ること。時代とともに考え方が転換してきた。",
    { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.58;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "販売志向", v: "「作ったものをいかに売るか」が起点。工場をフル稼働させ、広告と営業力で売りさばく発想" },
    { k: "顧客志向", v: "「顧客が何を求めているか」が起点。ニーズを調査し、そのニーズに合った商品を企画・開発する" },
  ], { fontSize: 11.5, labelW: 1.7, gap: 0.5 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "「工場をフル稼働させて作った商品を、広告と営業力で売りさばく」のが販売志向、「まず顧客のニーズを調査し、そのニーズに合った商品を企画・開発する」のが顧客志向（マーケティング・コンセプト）。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：現在の主流は顧客志向だが、販売志向が「常に誤り」というわけではない。両者の違いは「発想の起点がどちらか」という点で理解する。",
    years: mkYears(new Set(["'16", "'17", "'20", "'23", "'25"])),
  });
}


// ---------- Slide 58: C-30 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-30 ／ 過去問で確認する",
    title: "こう出題される（令和5年度 第28問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "企業経営理論",
  });
  const cy = addExamQuestion(s, {
    stem: "顧客価値に関する記述として、最も適切なものはどれか。",
    stemH: 0.6,
    choices: [
      { badge: "ア", text: "顧客が製品やサービスに期待する最も基本的な機能によってもたらされる価値を、機能的価値という。最も基本的な価値であるため、機能的価値が不十分であったり不明確であったりする製品やサービスが顧客に受け入れられることはない。" },
      { badge: "イ", text: "実際に製品やサービスを購入し、使用感などを経験してみなければ分からない価値を経験価値という。経験価値によって製品やサービスを訴求するためには、すでに利用した顧客によるクチコミなどをなるべく発生させないようにしつつ、プロモーションを行うべきである。" },
      { badge: "ウ", text: "製品やサービスの機能的価値は、「あって当たり前」の本質機能と、付加的な付随機能に分けることができる。このうち本質機能による機能的価値は、1つでも欠ければ競合する製品やサービスに比べて大幅に魅力が劣るため、自社の製品やサービスを差別化するために最も力を入れなければならない価値である。", h: 0.72 },
      { badge: "エ", text: "どのような基本的な機能を期待するかは顧客ごとに異なるのに対して、多くの顧客が製品やサービスに期待する感覚的価値は一般的に似通っているため、感覚的価値を訴求する製品やサービスは差別化が難しく、価格競争に陥りやすい。" },
      { badge: "オ", text: "文脈価値とは、顧客が製品やサービスを利用した際の状況に依存する。すなわち、顧客による特定の製品やサービスの利用と、その際の周辺環境や情景あるいは誰と一緒に利用したか、などの状況とともに創り出される価値である。" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2023/C1JI2023.pdf（令和5年度第1次試験）第28問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 59: C-30 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-30 ／ 過去問で確認する",
    title: "解答＆解説（令和5年度 第28問）",
    overview: "正解はオ。文脈価値は「いつ・どこで・誰と」利用したかという状況とともに生まれる価値。",
    tag: "企業経営理論",
  });
  const choices = [
      { badge: "ア", text: "顧客が製品やサービスに期待する最も基本的な機能によってもたらされる価値を、機能的価値という。最も基本的な価値であるため、機能的価値が不十分であったり不明確であったりする製品やサービスが顧客に受け入れられることはない。" },
      { badge: "イ", text: "実際に製品やサービスを購入し、使用感などを経験してみなければ分からない価値を経験価値という。経験価値によって製品やサービスを訴求するためには、すでに利用した顧客によるクチコミなどをなるべく発生させないようにしつつ、プロモーションを行うべきである。" },
      { badge: "ウ", text: "製品やサービスの機能的価値は、「あって当たり前」の本質機能と、付加的な付随機能に分けることができる。このうち本質機能による機能的価値は、1つでも欠ければ競合する製品やサービスに比べて大幅に魅力が劣るため、自社の製品やサービスを差別化するために最も力を入れなければならない価値である。", h: 0.72 },
      { badge: "エ", text: "どのような基本的な機能を期待するかは顧客ごとに異なるのに対して、多くの顧客が製品やサービスに期待する感覚的価値は一般的に似通っているため、感覚的価値を訴求する製品やサービスは差別化が難しく、価格競争に陥りやすい。" },
      { badge: "オ", text: "文脈価値とは、顧客が製品やサービスを利用した際の状況に依存する。すなわち、顧客による特定の製品やサービスの利用と、その際の周辺環境や情景あるいは誰と一緒に利用したか、などの状況とともに創り出される価値である。" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 4 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：オ", options: { bold: true, color: RED } },
    { text: "。「文脈価値」は、製品・サービスそのものに固定的に備わる価値ではなく、利用状況に応じて生まれる価値、という定義どおりの説明。ア：機能面で多少劣っても他の価値で顧客に選ばれる製品は実際に多く、言い切りすぎ。イ：経験価値を伝えるには、むしろ既に利用した顧客の口コミを積極的に活用する方が効果的で逆の説明。ウ：本質機能は欠ければ不満につながるが他社も当然備えており、差別化の源泉にはなりにくい。エ：感覚的価値はむしろ個人差が大きく、独自性を打ち出しやすい差別化の源泉になりやすい。", options: {  } },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.05, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.22 });
  cy += 1.05;
  s.addText([
    { text: "具体例\\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "顧客価値は「機能的価値（基本機能）」「経験価値（使ってみて分かる価値）」「感覚的価値（デザイン・好み）」「文脈価値（利用状況に応じた価値）」に整理されることが多い。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.72;
  s.addText("関連知識：基本機能は差別化の土台にはなるが決め手にはなりにくく、感覚的価値・文脈価値の方が差別化に直結しやすい、という向きを押さえておくこと。", {
    x: 0.55, y: cy, w: 12.25, h: 0.5, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.50;
  s.addText("出典：past_exams/1st_stage/1ji2023/C1JI2023.pdf（令和5年度第1次試験）第28問／正解：past_exams/1st_stage_answers/r05/2023c.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}
// ---------- Slide 60: C-31 マーケティング計画と戦略（STP、4P） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-31 ／ マーケティング計画と戦略（STP、4P）",
    title: "「誰に」を決めるSTP、「どう売るか」を決める4P",
    overview: "STPで狙う市場と立ち位置を決め、4Pで具体的な打ち手を組み立てる、マーケティング戦略の2大フレームワーク。",
    tag: "企業経営理論",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "セグメンテーション", tag: "STP①", desc: "市場を年齢・地域・ライフスタイルなどの基準で細分化する" },
    { name: "ターゲティング", tag: "STP②", desc: "細分化した市場（セグメント）の中から、自社が狙う市場を選ぶ" },
    { name: "ポジショニング", tag: "STP③", desc: "選んだ市場の中で、競合と比べた自社商品の独自の位置づけを明確にする" },
    { name: "Product", tag: "4P①", desc: "何を売るか（製品そのものの設計）" },
    { name: "Price", tag: "4P②", desc: "いくらで売るか（価格設定）" },
    { name: "Place", tag: "4P③", desc: "どこで売るか（流通・チャネル）" },
    { name: "Promotion", tag: "4P④", desc: "どう知らせ、購買を促すか（広告・販促）" },
  ], { rowH: 0.6, nameW: 2.6, tagW: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-2 アンゾフの成長マトリクス（「誰に」「何を」の整理の仕方が共通する）。",
    years: mkYears(new Set(["'17", "'18", "'19", "'20", "'24"])),
  });
}

// ---------- Slide 61: C-32 マーケティング・リサーチ ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-32 ／ マーケティング・リサーチ",
    title: "施策の前に「調べる」：定性調査と定量調査",
    overview: "市場や顧客を調査する活動がマーケティング・リサーチ。目的に応じて調査手法を使い分ける。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "マーケティング施策を考える前に、市場や顧客を調査する活動がマーケティング・リサーチ。調査方法は大きく2つに分かれる。",
    { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.58;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "定性調査", v: "インタビューやアンケートの自由記述で「なぜ」「どう思うか」を深掘りする。少人数向け・初期段階向き" },
    { k: "定量調査", v: "多数のデータを集めて数値で傾向を把握する。大規模なアンケート等で需要規模の検証に向く" },
  ], { fontSize: 11.5, labelW: 1.7, gap: 0.5 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "新商品のコンセプトを固める初期段階では少人数へのインタビュー（定性調査）でアイデアを練り、コンセプトが固まった段階で大規模なアンケート（定量調査）を行い需要規模を検証する、という使い分けが一般的。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-30 マーケティングの基礎概念（顧客志向を実現するための情報収集手段）。",
    years: mkYears(new Set(["'16", "'18", "'19", "'21", "'24", "'25"])),
  });
}


// ---------- Slide 62: C-32 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-32 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第28問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "企業経営理論",
  });
  const cy = addExamQuestion(s, {
    stem: "マーケティング・リサーチに関する記述として、最も適切なものはどれか。",
    stemH: 0.6,
    choices: [
      { badge: "ア", text: "1950年代から60年代にかけて、アメリカでは精神分析や臨床心理学を応用して消費者の深層心理を調べるニューロ・マーケティング・リサーチが盛んに行われたが、分析者による恣意的な解釈などが入り込む傾向があり、次第に行われなくなった。" },
      { badge: "イ", text: "自社の製品に対する顧客の満足度（変数X：7点尺度で回答）と顧客の年収（変数Y）との間の相関分析を行った結果、相関係数がゼロであったため、XとYは相互に無関係であると結論づけた。" },
      { badge: "ウ", text: "自社の製品に対する顧客の満足度（変数X：7点尺度で回答）を連続尺度とみて顧客の居住地（変数S：都道府県で回答）との間に関係があるかどうかを調べるために、カイ2乗分析を行った。" },
      { badge: "エ", text: "自社の製品を購入した顧客からサービスセンターに寄せられる手紙とハガキの内容を分析した結果、製品Aより製品Bに寄せられる不満の方が多いことが分かった。このため、直ちに製品Bの販売を中止することにした。" },
      { badge: "オ", text: "自社の製品を購入した全顧客を対象とする全数調査は、得られる回答データの正確性が高い一方で時間とコストの観点から現実的ではないため、単純無作為抽出法、層化抽出法などの標本抽出方法によるサンプリング調査が行われることが多い。" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第28問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 63: C-32 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-32 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第28問）",
    overview: "正解はオ。全数調査は非現実的なため、標本抽出によるサンプリング調査が一般的。",
    tag: "企業経営理論",
  });
  const choices = [
      { badge: "ア", text: "1950年代から60年代にかけて、アメリカでは精神分析や臨床心理学を応用して消費者の深層心理を調べるニューロ・マーケティング・リサーチが盛んに行われたが、分析者による恣意的な解釈などが入り込む傾向があり、次第に行われなくなった。" },
      { badge: "イ", text: "自社の製品に対する顧客の満足度（変数X：7点尺度で回答）と顧客の年収（変数Y）との間の相関分析を行った結果、相関係数がゼロであったため、XとYは相互に無関係であると結論づけた。" },
      { badge: "ウ", text: "自社の製品に対する顧客の満足度（変数X：7点尺度で回答）を連続尺度とみて顧客の居住地（変数S：都道府県で回答）との間に関係があるかどうかを調べるために、カイ2乗分析を行った。" },
      { badge: "エ", text: "自社の製品を購入した顧客からサービスセンターに寄せられる手紙とハガキの内容を分析した結果、製品Aより製品Bに寄せられる不満の方が多いことが分かった。このため、直ちに製品Bの販売を中止することにした。" },
      { badge: "オ", text: "自社の製品を購入した全顧客を対象とする全数調査は、得られる回答データの正確性が高い一方で時間とコストの観点から現実的ではないため、単純無作為抽出法、層化抽出法などの標本抽出方法によるサンプリング調査が行われることが多い。" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 4 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：オ", options: { bold: true, color: RED } },
    { text: "。全数調査はデータの精度は高いが対象が多いほど時間・費用がかさみ非現実的になるため、母集団から一部を抽出する標本調査が実務では広く用いられる。ア：1950〜60年代に流行したのは「モチベーション・リサーチ」であり「ニューロ・マーケティング・リサーチ」は近年発展した別の手法で名称が誤り。イ：相関係数ゼロは「直線的な関係がない」ことを示すにすぎず、曲線的な関係の可能性を否定できない。ウ：カイ2乗分析は名義尺度同士のクロス集計向けで、連続尺度をそのまま扱うのには不向き。エ：手紙・ハガキの苦情は自発的な一部の声にすぎず（自己選択バイアス）、これだけで販売中止を判断するのは早計。", options: {  } },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.05, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.22 });
  cy += 1.05;
  s.addText([
    { text: "具体例\\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "マーケティング・リサーチでは「定性調査／定量調査」の使い分けに加え、データの尺度（名義・順序・間隔・比率）に応じた分析手法の選択が問われる。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.72;
  s.addText("関連知識：全数調査（センサス）と標本調査（サンプリング）のコスト・精度のトレードオフは統計調査全般の基本知識として押さえておくこと。", {
    x: 0.55, y: cy, w: 12.25, h: 0.5, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.50;
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第28問／正解：past_exams/1st_stage_answers/r07/2025c.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}
// ---------- Slide 64: C-33 リレーションシップ・CRM ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-33 ／ リレーションシップ・CRM",
    title: "新規獲得より既存顧客との関係強化（CRM）",
    overview: "新規顧客の獲得コストは既存顧客の維持コストより一般に高く、既存顧客との関係強化が収益性に直結する。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText([
    { text: "新規顧客の獲得だけでなく、既存顧客との長期的な関係を維持・強化する考え方が", options: { color: INK } },
    { text: "リレーションシップ・マーケティング", options: { bold: true, color: INK } },
    { text: "。これを支えるシステムが", options: { color: INK } },
    { text: "CRM（顧客関係管理）", options: { bold: true, color: INK } },
    { text: "。", options: { color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, fontFace: F_BODY, fontSize: 11.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 });
  cy += 0.68;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "リレーションシップ", v: "新規獲得だけでなく既存顧客との長期的な関係を維持・強化する考え方" },
    { k: "CRM", v: "顧客の購買履歴や属性データを一元管理し、個別対応や継続的な関係構築に活用するシステム" },
  ], { fontSize: 11.5, labelW: 2.1, gap: 0.4 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "ある会員制ネットショップが、購買履歴に基づいて一人ひとりに合ったクーポンやレコメンドを送るのはCRMの活用例。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：新規顧客の獲得コストは既存顧客の維持コストより高いのが一般的。CRMは既存顧客との関係強化を通じて収益性を高める手段と位置づける。",
    years: mkYears(new Set(["'16", "'18", "'19", "'21", "'22"])),
  });
}

// ---------- Slide 65: C-34 サービス・マーケティング ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-34 ／ サービス・マーケティング",
    title: "モノと違う、サービスならではの4つの特性",
    overview: "無形性・同時性・非均質性・消滅性という4特性が、サービス・マーケティング特有の課題を生む。",
    tag: "企業経営理論",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "無形性", tag: "形がない", desc: "事前に確認・試すことができない" },
    { name: "同時性", tag: "生産＝消費", desc: "提供と消費が同時に起こる（例：美容院の施術）" },
    { name: "非均質性", tag: "ばらつく", desc: "提供者や状況によって品質にばらつきが出る" },
    { name: "消滅性", tag: "在庫できない", desc: "売れ残った提供機会は後から売ることができない（例：空席のまま出発した飛行機の座席）" },
  ], { nameW: 2.0, tagW: 2.1 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：近年は「製造業のサービス化」（機械を売るだけでなく稼働監視サービスまで提供する等）も重要テーマ。",
    years: mkYears(new Set(["'17", "'18", "'19", "'20", "'22", "'25"])),
  });
}

// ---------- Slide 66: C-35 デジタル・マーケティング ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-35 ／ デジタル・マーケティング",
    title: "ネット時代の購買行動モデル：AIDMAからAISASへ",
    overview: "インターネット普及で、消費者は購入前に「検索」し、購入後に「共有」するようになった。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText([
    { text: "従来のマス広告時代の消費者行動モデルが", options: { color: INK } },
    { text: "AIDMA", options: { bold: true, color: INK } },
    { text: "、インターネット時代のモデルとして提唱されたのが", options: { color: INK } },
    { text: "AISAS", options: { bold: true, color: INK } },
    { text: "。", options: { color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 11.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 });
  cy += 0.58;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "AIDMA", v: "Attention注意→Interest関心→Desire欲求→Memory記憶→Action購買（マス広告時代）" },
    { k: "AISAS", v: "Attention注意→Interest関心→Search検索→Action購買→Share共有（インターネット時代）" },
  ], { fontSize: 11, labelW: 1.3, gap: 0.5 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "テレビCMを見て興味を持った商品について、購入前にスマートフォンでレビューを検索し（Search）、購入後にSNSに写真を投稿する（Share）という一連の行動は、AISASモデルそのもの。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.75, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：AIDMAの「Desire・Memory」がAISASでは「Search」に、「Action」の後に「Share」が加わる点が変化のポイント。",
    years: mkYears(new Set(["'19", "'20", "'23", "'24"])),
  });
}

// ---------- Slide 67: C-36 プロダクト・マネジメント：PLC、新製品開発 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-36 ／ プロダクト・マネジメント：PLC、新製品開発",
    title: "プロダクト・ライフサイクル（PLC）で戦略を切り替える",
    overview: "製品は導入期→成長期→成熟期→衰退期の4段階をたどり、各段階で有効な戦略が異なる。",
    tag: "企業経営理論",
  });
  {
    const proseX = 0.55, proseW = 7.4;
    const diagX = 8.25, diagW = 3.4;
    let cy = 1.85;
    s.addText(
      "製品は導入期→成長期→成熟期→衰退期の4段階をたどり、各段階で有効な戦略が異なる。",
      { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 }
    );
    cy += 0.48;
    cy = addTermRows(s, proseX, cy, proseW, [
      { k: "導入期", v: "販売量少、認知度が低い。認知拡大に注力", gap: 0.4 },
      { k: "成長期", v: "販売量が急拡大、競合が参入。シェア・販売網の拡大", gap: 0.4 },
      { k: "成熟期", v: "販売量ピーク・競争激化。差別化、ブランド強化", gap: 0.4 },
      { k: "衰退期", v: "販売量が減少。撤退判断、または縮小継続", gap: 0.4 },
    ], { fontSize: 11.5, labelW: 1.1 });
    cy += 0.1;
    s.addText([
      { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
      { text: "新しいガジェットが発売直後（導入期）はアーリー・アダプター層に的を絞り、市場拡大につれ（成長期）大量生産・販路拡大を進め、競合増加で差別化が難しい時期（成熟期）にはブランドの独自性を強調する。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
    ], { x: proseX, y: cy, w: proseW, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

    drawProcessStepsVertical(s, diagX, 1.9, diagW, 4.5, [
      { label: "導入期", desc: "認知拡大" },
      { label: "成長期", desc: "シェア拡大" },
      { label: "成熟期", desc: "差別化" },
      { label: "衰退期", desc: "撤退判断" },
    ]);
  }
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-37 ブランディング（成熟期の差別化戦略の柱となる）。",
    years: mkYears(new Set(["'17", "'18", "'19", "'20", "'22", "'23", "'24"])),
  });
}

// ---------- Slide 68: C-37 ブランディング ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-37 ／ ブランディング",
    title: "ブランド・エクイティという「見えない資産」",
    overview: "ブランド力があれば、同じ機能の商品でも高価格で選ばれ、リピート購入される（アーカーのブランド・エクイティ論）。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText([
    { text: "ブランド・エクイティ", options: { bold: true, color: INK } },
    { text: "とは、ブランドが持つ資産的な価値のこと。同じ機能の商品でも、ブランド力があれば高い価格でも選ばれたり、リピート購入されたりする。", options: { color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, fontFace: F_BODY, fontSize: 11.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 });
  cy += 0.68;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "ブランド認知", v: "そのブランドが知られているか" },
    { k: "知覚品質", v: "品質が高いと思われているか（消費者の主観的な評価）" },
    { k: "ブランド・ロイヤルティ", v: "そのブランドへの愛着・継続購買意向" },
  ], { fontSize: 11, labelW: 2.1, gap: 0.4 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "同じような機能のスニーカーでも、有名スポーツブランドのロゴが付いているだけで高い価格で売れるのは、そのブランドが持つブランド・エクイティ（ブランド価値）によるもの。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：ブランド・エクイティは「知覚品質」であり、客観的な品質の高さそのものではない。消費者にどう「思われているか」が本質。",
    years: mkYears(new Set(["'16", "'18", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}


// ---------- Slide 69: C-37 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-37 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第37問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "企業経営理論",
  });
  const cy = addExamQuestion(s, {
    stem: "ブランディングに関する記述として、最も適切なものはどれか。",
    stemH: 0.6,
    choices: [
      { badge: "ア", text: "COO（Country of Origin）イメージは、製品に対する消費者の選好に影響を及ぼし、時間を経ても変化することはない。" },
      { badge: "イ", text: "コ・ブランディングは、コミュニケーション・コストを削減できるが、ブランド・エクイティが希釈化するリスクがあるため、新製品には適用できない。" },
      { badge: "ウ", text: "第三者ソース（専門誌、専門機関、知名度が高い評論家、ステータスがあるユーザーなど）による高評価は、ブランドに対する高い信頼性を付与し、消費者の態度を向上させるが、ソーシャルネットワーク上での影響力は極めて小さい。" },
      { badge: "エ", text: "著名人による推奨は、当該著名人のイメージと結びつく知覚は形成されるが、ブランドが思い出されないリスクがあり、ファン以外の注意はひきつけられない。" },
      { badge: "オ", text: "ライセンス供与は、在庫費用や製造費用をかけずにブランド認知を増やすことができるが、過剰露出になった場合、消費者の飽きが生じるリスクがある。" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第37問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 70: C-37 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-37 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第37問）",
    overview: "正解はオ。ブランド・ライセンシングは在庫・製造費を負わずに認知を広げられるが、過剰露出のリスクがある。",
    tag: "企業経営理論",
  });
  const choices = [
      { badge: "ア", text: "COO（Country of Origin）イメージは、製品に対する消費者の選好に影響を及ぼし、時間を経ても変化することはない。" },
      { badge: "イ", text: "コ・ブランディングは、コミュニケーション・コストを削減できるが、ブランド・エクイティが希釈化するリスクがあるため、新製品には適用できない。" },
      { badge: "ウ", text: "第三者ソース（専門誌、専門機関、知名度が高い評論家、ステータスがあるユーザーなど）による高評価は、ブランドに対する高い信頼性を付与し、消費者の態度を向上させるが、ソーシャルネットワーク上での影響力は極めて小さい。" },
      { badge: "エ", text: "著名人による推奨は、当該著名人のイメージと結びつく知覚は形成されるが、ブランドが思い出されないリスクがあり、ファン以外の注意はひきつけられない。" },
      { badge: "オ", text: "ライセンス供与は、在庫費用や製造費用をかけずにブランド認知を増やすことができるが、過剰露出になった場合、消費者の飽きが生じるリスクがある。" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 4 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：オ", options: { bold: true, color: RED } },
    { text: "。ブランド・ライセンシング（他社に自社ブランド名の使用権を供与する仕組み）は、供与側が製造・在庫コストを負わずにブランド露出を増やせる一方、露出しすぎるとブランドの特別感が薄れ消費者が飽きるリスクがある。ア：COOイメージはその国の技術力・イメージの変化とともに", options: {  } },
    { text: "時間の経過で変化しうる", options: { bold: true } },
    { text: "。イ：コ・ブランディングはむしろ新製品の認知獲得のために活用されることが多い。ウ：近年はSNS上の口コミ・インフルエンサーの影響力はむしろ", options: {  } },
    { text: "非常に大きい", options: { bold: true } },
    { text: "。エ：著名人推奨は、そのファン層にはむしろ強く注意を引きつけられる。", options: {  } },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.05, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.22 });
  cy += 1.05;
  s.addText([
    { text: "具体例\\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "ブランド要素の活用手法として、コ・ブランディング（共同ブランド展開）、ライセンス供与（使用権の供与）、セレブリティ・エンドースメント（著名人推奨）、第三者ソースの活用が代表的。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.72;
  s.addText("関連知識：それぞれの手法の「メリットと固有のリスク」をセットで覚えておくこと。", {
    x: 0.55, y: cy, w: 12.25, h: 0.5, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.50;
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第37問／正解：past_exams/1st_stage_answers/r07/2025c.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}
// ---------- Slide 71: C-38 消費者行動 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-38 ／ 消費者行動",
    title: "「高関与」か「低関与」かで購買行動が変わる",
    overview: "10年間ほぼ毎年出題される最頻出論点。関与度による購買行動の違いを理解する。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "消費者が商品を購入するまでの心理プロセスを理解することもマーケティングの重要な要素。購買への関与度によって行動パターンが異なる。",
    { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.58;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "高関与型", v: "じっくり検討する購買（住宅、自動車など）。多くの情報を集めて比較検討する" },
    { k: "低関与型", v: "あまり深く考えない購買（日用品など）。習慣的に選ばれることが多い" },
  ], { fontSize: 11.5, labelW: 1.5, gap: 0.5 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "住宅の購入では、多くの情報を集めて比較検討する「高関与型」の意思決定が行われるが、コンビニでのお菓子の購入は、ほとんど検討せず習慣的に選ばれる「低関与型」の意思決定であることが多い。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-39 プライシング（高関与型では価格以外の情報も比較検討材料になる）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}


// ---------- Slide 72: C-38 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-38 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第35問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "企業経営理論",
  });
  const cy = addExamQuestion(s, {
    stem: "消費行動におけるさまざまな現象に関する記述として、最も適切なものはどれか。",
    stemH: 0.55,
    choices: [
      { badge: "ア", text: "顕示的消費には、SNS映えする料理の写真を投稿したり、動画で自分の趣味のよさを公開したりする行為や、裕福な家庭における親が子供に高級品を身に着けさせる行為が含まれる。" },
      { badge: "イ", text: "現代の消費スタイルを包括的に記述する概念であるリキッド消費は、永続的で、所有ベースで、物質主義的な消費スタイルである。" },
      { badge: "ウ", text: "消費者情報処理の枠組みにおける快楽消費の評価基準には、感覚的な満足や空想、美的な楽しみ、感情的反応だけでなく、合理的判断や思考などの功利も含まれる。" },
      { badge: "エ", text: "消費の文化的意味や消費経験のダイナミックな側面を解明する消費文化理論の研究では、参与観察、デプス・インタビューなどを用いた定性的アプローチがとられており、定量的アプローチはとられていない。" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第35問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 73: C-38 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-38 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第35問）",
    overview: "正解はア。SNS映え投稿や高級品の誇示はいずれも顕示的消費の現代的な現れ。",
    tag: "企業経営理論",
  });
  const choices = [
      { badge: "ア", text: "顕示的消費には、SNS映えする料理の写真を投稿したり、動画で自分の趣味のよさを公開したりする行為や、裕福な家庭における親が子供に高級品を身に着けさせる行為が含まれる。" },
      { badge: "イ", text: "現代の消費スタイルを包括的に記述する概念であるリキッド消費は、永続的で、所有ベースで、物質主義的な消費スタイルである。" },
      { badge: "ウ", text: "消費者情報処理の枠組みにおける快楽消費の評価基準には、感覚的な満足や空想、美的な楽しみ、感情的反応だけでなく、合理的判断や思考などの功利も含まれる。" },
      { badge: "エ", text: "消費の文化的意味や消費経験のダイナミックな側面を解明する消費文化理論の研究では、参与観察、デプス・インタビューなどを用いた定性的アプローチがとられており、定量的アプローチはとられていない。" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 0 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：ア", options: { bold: true, color: RED } },
    { text: "。顕示的消費（ヴェブレンの誇示的消費）とは、自分の富や地位を他者に見せつけるための消費行動。SNSでの映える投稿や高級品を身につけさせる行為はいずれも該当する。イ：「リキッド消費」はむしろ「一時的」「アクセスベース」「脱物質主義的」であることが特徴で、説明は対比概念の「ソリッド消費」に近く逆。ウ：快楽消費の評価基準は「情緒的」側面が中心で、合理的判断などの「功利的」側面とは対比される概念。エ：消費文化理論は定性的アプローチが中心だが、定量的アプローチも併用されることがあり言い切りは誤り。", options: {  } },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.05, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.22 });
  cy += 1.05;
  s.addText([
    { text: "具体例\\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "顕示的消費（ヴェブレン）、リキッド消費（バーディ＆エッカート）、快楽消費（ハーシュマン＆ホルブルック）は消費者行動論の代表的な概念。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.72;
  s.addText("関連知識：似た名前の対概念（ソリッド消費 vs リキッド消費、功利的消費 vs 快楽的消費）とセットで整理しておくこと。", {
    x: 0.55, y: cy, w: 12.25, h: 0.5, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.50;
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第35問／正解：past_exams/1st_stage_answers/r07/2025c.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}
// ---------- Slide 74: C-39 プライシング ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-39 ／ プライシング",
    title: "価格はどう決まるか：3つのアプローチと心理的価格設定",
    overview: "コストベース・需要ベース・競争ベースという3つの価格決定アプローチに加え、心理的な価格設定の工夫も出題される。",
    tag: "企業経営理論",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "コストベース", tag: "原価起点", desc: "原価に一定の利益を上乗せして価格を決める" },
    { name: "需要ベース", tag: "顧客起点", desc: "顧客が「これくらいなら払ってもよい」と思う金額を基準に価格を決める" },
    { name: "競争ベース", tag: "競合起点", desc: "競合の価格を基準に価格を決める" },
    { name: "端数価格", tag: "心理効果", desc: "「1,000円」でなく「980円」と表示し、実際以上に安く感じさせる" },
  ], { nameW: 2.0, tagW: 1.6 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "具体例：高級ブランド品は原価でなく「顧客がブランド価値にどれだけ払う意思があるか」（需要ベース）であえて高価格を維持しブランドイメージを守る。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'21", "'22", "'24", "'25"])),
  });
}


// ---------- Slide 75: C-39 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-39 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第30問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "企業経営理論",
  });
  const cy = addExamQuestion(s, {
    stem: "プライシングに関する記述として、最も適切なものはどれか。",
    stemH: 0.6,
    choices: [
      { badge: "ア", text: "価格を消費者による「支出の痛み」として捉えれば、価格は安いほど売れ行きは伸びることになるが、他方で価格を「品質のバロメーター」や「プレステージ」として捉えれば、必ずしも安い価格の方が売れ行きが伸びるとは言えない。" },
      { badge: "イ", text: "顧客によって異なる価格を提示するプライシングは、鉄道サービスにおける学割やホテル業界におけるメンバーシップ制度などに見られるように古くから行われてきたが、時期によって価格を変えるプライシングが行われるようになったのは、デジタル・マーケティングによるダイナミック・プライシングが浸透した近年になってからである。", h: 0.72 },
      { badge: "ウ", text: "サブスクリプションは、製品やサービスの所有権の移転を行わずに使用する権利だけを販売するものであり、音楽や映像などがデジタル化され所有権の移転を行わずに転送できるようになったことなどに示されるように、デジタル時代になって誕生した新しい契約形態である。" },
      { badge: "エ", text: "自社製品を有利に扱ってくれる流通業者に対して、メーカーは割引価格などの金銭的見返りを提供することが多い。そのうち短期的な金銭的見返りは一般的にリベートと呼ばれ、長期的に提供されるアローワンスと区別される。" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第30問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 76: C-39 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-39 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第30問）",
    overview: "正解はア。価格には「安いほど得」と「高いほど高品質」という相反する2つの心理的な働きがある。",
    tag: "企業経営理論",
  });
  const choices = [
      { badge: "ア", text: "価格を消費者による「支出の痛み」として捉えれば、価格は安いほど売れ行きは伸びることになるが、他方で価格を「品質のバロメーター」や「プレステージ」として捉えれば、必ずしも安い価格の方が売れ行きが伸びるとは言えない。" },
      { badge: "イ", text: "顧客によって異なる価格を提示するプライシングは、鉄道サービスにおける学割やホテル業界におけるメンバーシップ制度などに見られるように古くから行われてきたが、時期によって価格を変えるプライシングが行われるようになったのは、デジタル・マーケティングによるダイナミック・プライシングが浸透した近年になってからである。", h: 0.72 },
      { badge: "ウ", text: "サブスクリプションは、製品やサービスの所有権の移転を行わずに使用する権利だけを販売するものであり、音楽や映像などがデジタル化され所有権の移転を行わずに転送できるようになったことなどに示されるように、デジタル時代になって誕生した新しい契約形態である。" },
      { badge: "エ", text: "自社製品を有利に扱ってくれる流通業者に対して、メーカーは割引価格などの金銭的見返りを提供することが多い。そのうち短期的な金銭的見返りは一般的にリベートと呼ばれ、長期的に提供されるアローワンスと区別される。" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 0 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：ア", options: { bold: true, color: RED } },
    { text: "。価格には「安いほど得（支出の痛みが小さい）」という側面と、「高い価格＝高品質・ステータス」という品質・プレステージの手がかりとしての側面があり、後者が強い高級品などでは必ずしも安いほど売れるとは限らない。イ：時期による価格変動は航空業界のイールドマネジメントなどデジタル化以前から行われてきた慣行。ウ：所有権を移転せず利用権のみを販売する契約形態自体は新聞・雑誌の定期購読など", options: {  } },
    { text: "デジタル化以前から存在", options: { bold: true } },
    { text: "している。エ：短期的な見返りは一般に「アローワンス」、継続的なものは「リベート」と呼ばれ、呼び方が入れ替わっている。", options: {  } },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.05, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.22 });
  cy += 1.05;
  s.addText([
    { text: "具体例\\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "高級ブランド品は「顧客がブランド価値にどれだけ払う意思があるか」（需要ベース）で価格が決められ、あえて高価格を維持することでブランドイメージを守る例が、アの考え方と対応する。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.72;
  s.addText("関連知識：価格の心理的効果として「支出の痛み」と「品質・ステータスのシグナル」という相反する2つの働きがある点を押さえておくこと。", {
    x: 0.55, y: cy, w: 12.25, h: 0.5, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.50;
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第30問／正解：past_exams/1st_stage_answers/r07/2025c.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}
// ---------- Slide 77: C-40 コミュニケーション ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-40 ／ コミュニケーション",
    title: "マスメディアとSPメディアを組み合わせる",
    overview: "企業が消費者に情報を伝える手段には、マスメディア広告とSP（セールス・プロモーション）メディアがある。",
    tag: "企業経営理論",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText([
    { text: "企業が消費者に情報を伝える「コミュニケーション」手段には、テレビCMなどの", options: { color: INK } },
    { text: "マスメディア", options: { bold: true, color: INK } },
    { text: "を使った広告と、チラシ・ノベルティなどの", options: { color: INK } },
    { text: "SPメディア", options: { bold: true, color: INK } },
    { text: "がある。", options: { color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, fontFace: F_BODY, fontSize: 11.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 });
  cy += 0.68;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "マスメディア広告", v: "テレビ・新聞・雑誌等、不特定多数に一斉に届ける広告" },
    { k: "SPメディア", v: "チラシ・ノベルティ・店頭POP等、購買行動に近い場面で使う販促媒体" },
  ], { fontSize: 11, labelW: 2.1, gap: 0.4 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "新商品のテレビCM（マスメディア広告）で認知を広げつつ、店頭でのポップやサンプリング（SPメディア）で最後の後押しをする、という組み合わせが一般的。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：C-41 プロモーション（コミュニケーション手段の使い分けとセットで問われる）。",
    years: mkYears(new Set(["'17", "'20", "'21", "'23", "'24", "'25"])),
  });
}


// ---------- Slide 78: C-40 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-40 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第31問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "企業経営理論",
  });
  const cy = addExamQuestion(s, {
    stem: "マーケティング・コミュニケーションに関する記述として、最も適切なものはどれか。",
    stemH: 0.6,
    choices: [
      { badge: "ア", text: "インフィード広告は、SNSのコンテキストやデザインとマッチさせて表示され、あたかも投稿の1つであるかのようにタイムラインに溶け込み、ユーザーの利用体験を妨げないことを目指す広告である。" },
      { badge: "イ", text: "コミュニケーションに対する消費者の反応を表したモデルであるAIDMAモデルやAISASモデル、FCBグリッドなどの階層モデルでは、消費者の反応が認知段階から感情段階を経て行動段階に進むと考える点で共通している。" },
      { badge: "ウ", text: "サードパーティ・クッキーの利用が法律などにより制限されると、インターネット広告の配信精度や広告効果が低下するリスクがあるが、リ・ターゲティング広告には影響はない。" },
      { badge: "エ", text: "日本では2022年に改正された個人情報保護法において、クッキーは個人情報であり、個人情報保護の対象に含まれると規定された。" },
      { badge: "オ", text: "ネイティブ広告においては、広告のデザインとフォーマットを広告が配信される媒体の記事やコンテンツの形式や機能と一体化させることが必要であるが、広告の内容やクリック後に表示される遷移先のコンテンツは自由に設定することができる。" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第31問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 79: C-40 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-40 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第31問）",
    overview: "正解はア。インフィード広告はSNSのフィードに自然に溶け込み利用体験を妨げない広告。",
    tag: "企業経営理論",
  });
  const choices = [
      { badge: "ア", text: "インフィード広告は、SNSのコンテキストやデザインとマッチさせて表示され、あたかも投稿の1つであるかのようにタイムラインに溶け込み、ユーザーの利用体験を妨げないことを目指す広告である。" },
      { badge: "イ", text: "コミュニケーションに対する消費者の反応を表したモデルであるAIDMAモデルやAISASモデル、FCBグリッドなどの階層モデルでは、消費者の反応が認知段階から感情段階を経て行動段階に進むと考える点で共通している。" },
      { badge: "ウ", text: "サードパーティ・クッキーの利用が法律などにより制限されると、インターネット広告の配信精度や広告効果が低下するリスクがあるが、リ・ターゲティング広告には影響はない。" },
      { badge: "エ", text: "日本では2022年に改正された個人情報保護法において、クッキーは個人情報であり、個人情報保護の対象に含まれると規定された。" },
      { badge: "オ", text: "ネイティブ広告においては、広告のデザインとフォーマットを広告が配信される媒体の記事やコンテンツの形式や機能と一体化させることが必要であるが、広告の内容やクリック後に表示される遷移先のコンテンツは自由に設定することができる。" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 0 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：ア", options: { bold: true, color: RED } },
    { text: "。インフィード広告は、SNSやニュースアプリのフィードの中に周囲の投稿とデザインを合わせて自然に溶け込む形で配信され、利用体験を妨げにくいことを狙った広告形式。イ：FCBグリッドは「関与度×思考・感情のどちらが優位か」で反応を4パターンに分類する枠組みで、必ずしも一方向の順序を前提としない。ウ：リ・ターゲティング広告はまさにサードパーティ・クッキーでユーザーを識別する手法のため、制限の直接的な影響を", options: {  } },
    { text: "受ける", options: { bold: true } },
    { text: "。エ：クッキー単体は原則個人情報そのものには当たらず、一律の規定ではない。オ：ネイティブ広告は遷移先のコンテンツも整合性が重視され「自由に設定できる」わけではない。", options: {  } },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.05, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.22 });
  cy += 1.05;
  s.addText([
    { text: "具体例\\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "インフィード広告はネイティブ広告の一形態（フィード型ネイティブ広告）と位置づけられることが多い。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.72;
  s.addText("関連知識：AIDMA/AISASなどの「階層モデル」とFCBグリッドのような「関与度モデル」は性質が異なる分類の仕方である点を区別しておくこと。", {
    x: 0.55, y: cy, w: 12.25, h: 0.5, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.50;
  s.addText("出典：past_exams/1st_stage/1ji2025/C1JI2025.pdf（令和7年度第1次試験）第31問／正解：past_exams/1st_stage_answers/r07/2025c.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}
// ---------- Slide 80: C-41 プロモーション ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-41 ／ プロモーション",
    title: "3つのプロモーション手法と、プル／プッシュの使い分け",
    overview: "人的販売・販売促進・PRという3手法、そして消費者に直接働きかけるか流通業者に働きかけるかの戦略選択が問われる。",
    tag: "企業経営理論",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "人的販売", tag: "対面", desc: "営業担当者が直接顧客に働きかける（高額商品・法人営業で重要）" },
    { name: "販売促進", tag: "短期", desc: "値引き、クーポン、サンプリングなど、短期的に購買を促す施策" },
    { name: "PR", tag: "間接", desc: "報道機関への情報提供等で評判を高める。掲載可否や内容は自社でコントロールできない" },
    { name: "プル戦略", tag: "対消費者", desc: "広告等で消費者の需要を喚起し、小売店に「置いてほしい」と言わせる" },
    { name: "プッシュ戦略", tag: "対流通", desc: "卸・小売店に営業をかけて棚を確保する" },
  ], { rowH: 0.75, nameW: 2.0, tagW: 1.6 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "具体例：テレビCMで消費者の関心を高め店頭で探させるのがプル戦略、営業担当が小売店を回り陳列棚を確保するのがプッシュ戦略。",
    years: mkYears(new Set(["'18", "'24", "'25"])),
  });
}

// ---------- Slide 81: C-42 流通チャネル ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "C-42 ／ 流通チャネル",
    title: "チャネルの開放度と、オムニチャネル時代の3つの考え方",
    overview: "開放的・選択的・専属的という開放度の違いに加え、実店舗とネットの垣根をなくすオムニチャネル系の概念も頻出。",
    tag: "企業経営理論",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "開放的チャネル", tag: "広く", desc: "多くの店舗に幅広く置いてもらう（例：コンビニで売る日用品）" },
    { name: "選択的チャネル", tag: "絞る", desc: "一定の基準を満たす店舗にのみ置く（例：百貨店・ブランド専門店限定の高級化粧品）" },
    { name: "専属的チャネル", tag: "独占", desc: "特定の1社にのみ独占的に販売権を与える" },
    { name: "オムニチャネル", tag: "融合", desc: "実店舗とネット通販の垣根をなくし、どちらでもスムーズに買い物できるようにする" },
    { name: "O2O", tag: "Online→Offline", desc: "オンラインでの行動をオフライン（実店舗）の購買に結びつける" },
    { name: "OMO", tag: "意識させない", desc: "オンラインとオフラインの垣根をそもそも意識させない" },
  ], { rowH: 0.6, nameW: 2.3, tagW: 1.9 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "具体例：ネットで注文して店舗で受け取れるサービスはO2O・オムニチャネルの代表例。",
    years: mkYears(new Set(["'16", "'17", "'18", "'21", "'22", "'23"])),
  });
}

const outPath = path.join(__dirname, "..", "..", "slides", "1st_stage", "C_business_administration.pptx");
pres.writeFile({ fileName: outPath }).then(() => {
  console.log("wrote", outPath);
});
