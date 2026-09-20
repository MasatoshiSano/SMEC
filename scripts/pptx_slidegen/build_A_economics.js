// Builds slides/1st_stage/A_economics.pptx from lib.js components.
// Run: node scripts/pptx_slidegen/build_A_economics.js
//
// Mirrors slides/1st_stage/A_economics.html (HTML is the source of truth for
// content). See docs/13_pptx_slide_template_spec.md for the component guide.

const path = require("path");
const {
  newPres, addCoverSlide, addDividerSlide, addHeader, addFreqBar, mkYears,
  addTermRows, addRowList, addExamQuestion,
  drawQuadrant, drawLineChart,
  INK, INK_SOFT, RED, LINE, GHOST, F_HEAD, F_BODY, F_MONO,
} = require("./lib");

const pres = newPres();

// ---------- Slide 01: 表紙 ----------
addCoverSlide(pres, {
  eyebrow: "中小企業診断士 第1次試験",
  subjectNo: "SUBJECT 01 / 07",
  title: "経済学・経済政策",
  subtitle: "マクロ経済学とミクロ経済学の2分野、36論点。「用語の定義」と「グラフの動く方向」を正確に押さえれば得点源にできる、数式・グラフ問題中心の科目。",
  stats: [["36", "収録論点数"], ["4", "分野"], ["28/36", "頻出ランクA"]],
  tag: "docs/textbook/A_economics_textbook.md",
  notes: "表紙スライド。経済学・経済政策の全体像を示す。",
});

// ---------- Slide 02: PART1 経済指標の見方や読み方 区切り ----------
addDividerSlide(pres, {
  ghostNo: "01",
  partNo: "PART 01",
  partLabel: "経済学・経済政策 ／ A-1〜A-6",
  title: "経済指標の見方や読み方",
  desc: "日本経済は今どんな状態にあるのか。それを客観的に把握するための「ものさし」となる経済指標を扱う。ニュースで見る数字の意味を正確に理解することがゴール。",
  chips: ["A-1 GDP", "A-3 物価指数", "A-6 国際収支"],
  notes: "PART1区切りスライド。",
});

// ---------- Slide 03: A-1 国民経済計算の概念・国民所得統計 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-1 ／ 国民経済計算の概念・国民所得統計",
    title: "GDPは「国内」の付加価値の合計",
    overview: "同じ経済活動を生産・分配・支出の3側面から見ても合計額は必ず一致する（三面等価の原則）。",
    tag: "経済学・経済政策",
  });
  s.addText(
    "GDP・GNI（GNP）・NNP・NIは無関係な4つの用語ではなく、GDPを起点に「集計範囲を国民ベースに変える」「減価償却を引く」「間接税・補助金を調整する」という一連の計算でつながる所得概念の連鎖。",
    { x: 0.55, y: 1.72, w: 12.25, h: 0.5, valign: "top", fontFace: F_BODY, fontSize: 10.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  let cy03 = addRowList(s, 0.55, 2.4, 12.25, [
    { name: "GDP", tag: "国内総生産", desc: "国内で生産された付加価値の合計（国籍を問わない）" },
    { name: "GNI／GNP", tag: "国民総所得", desc: "自国民（居住者）が生み出した所得の合計。GNI＝GDP＋海外からの純所得受取" },
    { name: "NNP", tag: "国民純生産", desc: "GNP（GNI）から固定資本減耗（減価償却相当）を差し引いたもの" },
    { name: "NI", tag: "国民所得", desc: "NNPから間接税を差し引き、補助金を加えたもの" },
  ], { nameW: 1.7, tagW: 2.0, rowH: 0.75 });
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "海外に生産拠点を持つ製造業A社の場合、国内工場の生産額はGDPに算入されるが海外子会社の生産額は算入されない。一方、海外子会社からの配当・利子収入（純所得受取）はGNIには算入される、という違いが決算上も生じる。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy03 + 0.12, w: 12.25, h: 0.75, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：GDPは一定期間の「フロー」概念。ある時点の資産総額を示す「ストック」（国富等）と混同しないこと。中古品・株式の売買はGDPに計上されない。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 04: A-2 雇用統計 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-2 ／ 雇用統計（失業率、有効求人倍率等）",
    title: "完全失業率と有効求人倍率で労働市場を見る",
    overview: "15歳以上人口は労働力人口と非労働力人口に分かれ、労働力人口はさらに就業者と完全失業者に分かれる。",
    tag: "経済学・経済政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "15歳以上人口は「労働力人口」（働く意思のある人）と「非労働力人口」に分かれ、労働力人口はさらに「就業者」と「完全失業者」に分かれる。",
    { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.58;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "完全失業率", v: "完全失業者数 ÷ 労働力人口 × 100。仕事がなく、求職活動をしている人が対象", gap: 0.5 },
    { k: "有効求人倍率", v: "有効求人数 ÷ 有効求職者数。1.0超は「人手不足・売り手市場」を示す", gap: 0.5 },
  ], { fontSize: 11.5, labelW: 1.7 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "好景気で企業が採用を増やすと有効求人倍率は上昇し、完全失業率は低下する傾向。不況期は逆の動きになる。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.68;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "完全失業率の低下＝景気回復とは限らない。求職をあきらめた人（求職意欲喪失効果）が非労働力人口に移ると、分母が縮み見かけ上失業率が下がることがある。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：A-4 景気動向指数（完全失業率は遅行指数の代表例）。",
    years: mkYears(new Set(["'17", "'20", "'21", "'25"])),
  });
}

// ---------- Slide 05: A-3 物価指数 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-3 ／ 物価指数（CPI、GDPデフレーター等）",
    title: "3つの物価指数と2つの計算方式",
    overview: "物価の測り方には対象範囲の異なる3指標と、ウェイトの取り方が異なる2つの計算方式がある。",
    tag: "経済学・経済政策",
  });
  s.addText(
    "3つの物価指数の違いは「何を測るか（対象範囲）」だけでなく、CPIとGDPデフレーターでは「どうウェイトを付けるか（計算方式）」も異なる点が試験で問われる、2つの軸の組み合わせ。",
    { x: 0.55, y: 1.72, w: 12.25, h: 0.5, valign: "top", fontFace: F_BODY, fontSize: 10.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  let cy05 = addRowList(s, 0.55, 2.4, 12.25, [
    { name: "CPI", tag: "消費者物価指数", desc: "総務省公表。消費者が購入する財・サービスの価格変動（ラスパイレス方式）" },
    { name: "CGPI", tag: "企業物価指数", desc: "日本銀行公表。企業間で取引される商品（原材料・中間財）の価格変動" },
    { name: "GDPデフレーター", tag: "名目GDP÷実質GDP×100", desc: "経済全体（消費・投資・政府支出・輸出入）の物価動向（パーシェ方式に近い）" },
  ], { rowH: 1.0, nameW: 2.6, tagW: 2.6 });
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "小売チェーンのB社は販売価格の値上げ幅を検討する際に消費者が体感する消費者物価指数（CPI）を参考にする一方、製造業A社は仕入れる原材料・部品のコスト上昇を把握するために企業物価指数（CGPI）を確認する、というように立場によって参照すべき指数が異なる。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy05 + 0.1, w: 12.25, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：CPI（ラスパイレス方式＝基準年数量固定）は代替効果を反映できず物価上昇率を実態より高めに示す傾向。CPIとGDPデフレーターは算出方式が異なる点が頻出。",
    years: mkYears(new Set(["'16", "'18", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 06: A-4 景気動向指数 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-4 ／ 景気動向指数（DI・CI）",
    title: "DIは波及度合い、CIは大きさを見る",
    overview: "内閣府が毎月公表する景気動向指数には、着目点の異なるDIとCIの2種類がある。",
    tag: "経済学・経済政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "景気動向指数は複数の経済指標（約30種）を統合したもの。着目点の違う2種類がある。",
    { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.48;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "DI（拡散指数）", v: "改善している指標の割合。景気の「波及の度合い（範囲）」を見る。50%が拡張／後退の境目", gap: 0.5 },
    { k: "CI（合成指数）", v: "各指標の変化率を合成。景気変動の「大きさ・テンポ（量）」を見る", gap: 0.5 },
  ], { fontSize: 11.5, labelW: 1.7 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "景気回復期は、まず株価・新規求人数（先行指数）が上向き、生産・有効求人倍率（一致指数）が続き、最後に失業率（遅行指数）が下がる、という順序で波及する。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.65, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.73;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「DIは量、CIは範囲」という説明は逆。DI＝波及度合い（範囲）、CI＝大きさ・テンポ（量）が正しい対応。景気の「山」「谷」は一致指数をもとに事後決定される。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：A-2 雇用統計（完全失業率は遅行指数の代表例）。",
    years: mkYears(new Set(["'17", "'18", "'23"])),
  });
}

// ---------- Slide 07: A-5 マネーストック統計 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-5 ／ マネーストック統計",
    title: "M1〜M3の違いは「対象金融機関の広さ」",
    overview: "金融機関・中央政府を除く経済主体が保有する通貨量の残高。範囲の狭いものから広いものへ4段階に分類される。",
    tag: "経済学・経済政策",
  });
  s.addText(
    "M1〜M3の違いは「通貨の種類」ではなく「対象とする金融機関の範囲の広さ」。ゆうちょ銀行・信用組合などを含むかどうかで段階的に範囲が広がっていく。",
    { x: 0.55, y: 1.72, w: 12.25, h: 0.4, valign: "top", fontFace: F_BODY, fontSize: 10.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  let cy07 = addRowList(s, 0.55, 2.4, 12.25, [
    { name: "M1", tag: "現金＋預金通貨", desc: "いつでも引き出せる現金通貨と普通・当座預金" },
    { name: "M2", tag: "M1＋準通貨＋CD", desc: "対象は国内銀行・信用金庫等に限定（ゆうちょ銀行等は含まない）" },
    { name: "M3", tag: "範囲を拡大", desc: "ゆうちょ銀行・信用組合・農協なども含む全預金取扱機関が対象" },
    { name: "広義流動性", tag: "最も広い範囲", desc: "M3に金銭信託・投資信託・国債・外債等の金融商品を加えたもの" },
  ], { rowH: 0.75, nameW: 2.0, tagW: 2.2 });
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "製造業A社の経理部が全社の余資を複数の金融機関に分散して預けている場合、都市銀行・地方銀行の預金はM2に含まれるが、系列の信用組合や農協の口座はM2には含まれずM3で初めて捕捉される。全社の資金繰りを把握するにはより広い範囲の統計を見る必要が生じる。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy07 + 0.1, w: 12.25, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：M1〜M3の違いは通貨の種類でなく対象金融機関の広さ。マネタリーベース（日銀が直接供給する通貨）とマネーストックは別概念で、両者は信用乗数を介した間接的な関係。",
    years: mkYears(new Set(["'17", "'19", "'20", "'21", "'23"])),
  });
}

// ---------- Slide 08: A-6 国際収支・為替レート ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-6 ／ 国際収支・為替レート",
    title: "経常収支＋資本移転等収支－金融収支＋誤差脱漏＝0",
    overview: "国際収支統計は一国が一定期間に外国と行った取引の記録。複式簿記の考え方に基づく恒等式が成り立つ。",
    tag: "経済学・経済政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "国際収支統計は一国が一定期間に外国と行った取引の記録。3つの大項目で構成される。",
    { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.48;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "経常収支", v: "貿易・サービス収支、第一次所得収支（配当・利子）、第二次所得収支（無償の資金移動）の合計", gap: 0.55 },
    { k: "資本移転等収支", v: "対価を伴わない資本の移転（インフラ無償援助等）", gap: 0.4 },
    { k: "金融収支", v: "対外的な金融資産・負債の増減。海外への投資（資産の増加）はプラスとして計上", gap: 0.5 },
  ], { fontSize: 11.5, labelW: 2.1 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "近年の日本は貿易収支が赤字の月・年もあるが、過去の対外投資から得られる配当・利子収入（第一次所得収支）が大幅黒字のため、経常収支全体は黒字を維持している。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：金融収支は「対外資産の増加はプラス」と符号が直感に反する。「金融収支の黒字＝海外への資金流出（対外資産の増加）超過」という意味を正しく理解すること。",
    years: mkYears(new Set(["'16", "'17", "'19", "'20", "'23", "'25"])),
  });
}

// ---------- Slide 09: PART2 マクロ経済理論・国際経済 区切り ----------
addDividerSlide(pres, {
  ghostNo: "02",
  partNo: "PART 02",
  partLabel: "経済学・経済政策 ／ A-7〜A-17",
  title: "マクロ経済理論・国際経済",
  desc: "一国全体の経済がどう動くかを理論的に説明するマクロ経済学の中核。45度線分析・IS-LM分析という2大フレームワークに加え、開放経済（貿易・為替）まで扱う、A科目で最も出題数が多い山場。",
  chips: ["A-9 IS-LM分析", "A-14 比較生産費説", "A-17 マンデル=フレミング"],
  notes: "PART2区切りスライド。",
});

// ---------- Slide 10: A-7 生産物市場とGDP決定理論 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-7 ／ 生産物市場とGDP決定理論",
    title: "45度線分析：総支出線と45度線の交点で均衡GDPが決まる",
    overview: "有効需要の原理：GDPの水準は需要（Y=C+I+G）の大きさで決まる。政府支出の増加はそれ以上にGDPを押し上げる（乗数効果）。",
    tag: "経済学・経済政策",
  });
  drawLineChart(s, 0.55, 1.9, 7.4, 4.2, {
    xLabel: "国民所得 Y", yLabel: "総支出 E",
    series: [
      { x1: 0, y1: 0, x2: 0.8, y2: 0.8, color: INK_SOFT, width: 1.3, dashType: "dash", label: "45°", labelDx: -0.55, labelDy: 0.05 },
      { x1: 0.05, y1: 0.15, x2: 0.95, y2: 0.72, color: RED, width: 2.2, label: "C+I+G" },
    ],
    point: { nx: 0.435, ny: 0.435, label: "均衡" },
  });
  s.addText([
    { text: "総支出線（C+I+G）", options: { bold: true, color: INK } },
    { text: "の傾き＝限界消費性向（1より小さい）のため45度線より緩やか。", options: {} },
  ], { x: 8.25, y: 2.1, w: 4.05, h: 1.0, fontFace: F_BODY, fontSize: 11, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.35 });
  s.addText([
    { text: "政府支出Gを1増やすとGDPは1/(1-c)倍（乗数）増加する。均衡予算（同額増税とセット）でもGDPは支出増加分だけ増える（均衡予算乗数＝1）。", options: {} },
  ], { x: 8.25, y: 3.3, w: 4.05, h: 1.6, fontFace: F_BODY, fontSize: 10.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.4 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：A-9 IS-LM分析（財市場の均衡を利子率も含めて拡張したもの）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 11: A-8 貨幣市場と利子率 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-8 ／ 貨幣市場と利子率",
    title: "貨幣需要の3つの動機（流動性選好理論）",
    overview: "ケインズは貨幣を持ちたがる動機を3つに整理した。投機的動機だけが利子率に依存する点がポイント。",
    tag: "経済学・経済政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "貨幣需要は「所得が増えるほど増加し、利子率が上がるほど減少する」性質を持つ。3つの動機に整理される。",
    { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.48;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "取引動機", v: "日々の決済のための需要。所得に依存する", gap: 0.4 },
    { k: "予備的動機", v: "不測の事態に備える需要。これも所得に依存する", gap: 0.4 },
    { k: "投機的動機", v: "投資タイミングを計る需要。利子率が低いほど大きくなる", gap: 0.4 },
  ], { fontSize: 11.5, labelW: 1.7 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "好況で所得が増えると取引に必要なお金が増え利子率に上昇圧力。金融緩和で貨幣供給を増やすと利子率は低下圧力を受ける。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.68;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "投機的動機は「利子率」に依存し「所得」には依存しない。取引・予備的動機は逆に「所得」に依存し「利子率」にはあまり依存しない。流動性の罠（利子率が下限に達し金融政策が効かなくなる状態）もあわせて押さえる。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "B", rankLabel: "頻出論点",
    related: "関連：A-9 IS-LM分析（貨幣市場の均衡がLM曲線として組み込まれる）。",
    years: mkYears(new Set(["'24", "'25"])),
  });
}

// ---------- Slide 12: A-9 IS-LM曲線・分析 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-9 ／ IS-LM曲線・分析",
    title: "財市場と貨幣市場を同時に満たすIS-LM分析",
    overview: "IS曲線（財市場の均衡）は右下がり、LM曲線（貨幣市場の均衡）は右上がり。交点で均衡GDPと均衡利子率が決まる。",
    tag: "経済学・経済政策",
  });
  drawLineChart(s, 0.55, 1.9, 7.4, 4.2, {
    xLabel: "Y", yLabel: "r",
    series: [
      { x1: 0.05, y1: 0.85, x2: 0.95, y2: 0.1, color: INK, width: 2.2, label: "IS", labelDy: -0.4 },
      { x1: 0.05, y1: 0.1, x2: 0.95, y2: 0.85, color: RED, width: 2.2, label: "LM" },
    ],
    point: { nx: 0.5, ny: 0.475, label: "E" },
  });
  s.addText([
    { text: "IS曲線", options: { bold: true } }, { text: "：財市場の均衡。右下がり（利子率↓→投資↑→GDP↑）\n\n", options: {} },
    { text: "LM曲線", options: { bold: true } }, { text: "：貨幣市場の均衡。右上がり（GDP↑→取引需要↑→利子率↑）\n\n", options: {} },
    { text: "財政政策", options: { bold: true, color: INK } }, { text: "→ISが右へシフト（クラウディングアウト発生）\n\n", options: {} },
    { text: "金融政策", options: { bold: true, color: RED } }, { text: "→LMが右へシフト", options: {} },
  ], { x: 8.25, y: 2.0, w: 4.05, h: 4.0, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：IS・LM曲線が緩やかなほど「もう一方」の政策が効きやすい（IS緩やか→金融政策が効く、LM緩やか→財政政策が効く）という対応関係が頻出。",
    years: mkYears(new Set(["'16", "'17", "'20", "'21", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 13: A-9 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-9 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第10問設問2）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "経済学・経済政策",
  });
  const cy = addExamQuestion(s, {
    stem: "IS曲線とLM曲線に関する記述の正誤の組み合わせとして、最も適切なものを下記の解答群から選べ。\n\nａ　限界消費性向が大きいほどIS曲線の傾きはより緩やかになり、貨幣供給の増加によるGDPの拡大効果は大きくなる。\nｂ　投資の利子感応度が小さいほど、IS曲線の傾きはより急になり、貨幣供給の増加によるGDPの拡大効果は小さくなる。\nｃ　貨幣需要の所得感応度が大きいほど、LM曲線の傾きはより緩やかになり、政府支出の増加によるGDPの拡大効果は小さくなる。\nｄ　貨幣需要の利子感応度が大きいほど、LM曲線の傾きはより緩やかになり、政府支出の増加によるGDPの拡大効果は大きくなる。",
    choices: [
      { badge: "ア", text: "ａ：正　ｂ：正　ｃ：誤　ｄ：正" },
      { badge: "イ", text: "ａ：正　ｂ：誤　ｃ：正　ｄ：誤" },
      { badge: "ウ", text: "ａ：正　ｂ：誤　ｃ：誤　ｄ：正" },
      { badge: "エ", text: "ａ：誤　ｂ：正　ｃ：正　ｄ：誤" },
      { badge: "オ", text: "ａ：誤　ｂ：正　ｃ：誤　ｄ：正" },
    ],
    stemH: 1.15,
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/A1JI2025.pdf（令和7年度第1次試験）第10問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 14: A-9 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-9 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第10問設問2）",
    overview: "正解はア。cの前半（LM曲線の傾きの説明）が誤りである点に注意。",
    tag: "経済学・経済政策",
  });
  const choices = [
    { badge: "ア", text: "ａ：正　ｂ：正　ｃ：誤　ｄ：正" },
    { badge: "イ", text: "ａ：正　ｂ：誤　ｃ：正　ｄ：誤" },
    { badge: "ウ", text: "ａ：正　ｂ：誤　ｃ：誤　ｄ：正" },
    { badge: "エ", text: "ａ：誤　ｂ：正　ｃ：正　ｄ：誤" },
    { badge: "オ", text: "ａ：誤　ｂ：正　ｃ：誤　ｄ：正" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 0 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：ア", options: { bold: true, color: RED } },
    { text: "（a正・b正・c誤・d正）。ａ：限界消費性向cが大きいほど乗数1/(1-c)が大きくIS曲線は緩やかになり、金融政策の効果は大きくなる＝正。ｂ：投資の利子感応度が小さいとIS曲線は急になり、金融政策の効果は小さくなる＝正。ｃ：貨幣需要の所得感応度が大きいとLM曲線は急になる（「緩やかになる」が誤り）。ｄ：貨幣需要の利子感応度が大きいとLM曲線は緩やかになり、財政政策の効果は大きくなる＝正。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.1, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });
  cy += 1.18;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "投資が利子率にほとんど反応しない経済（IS曲線が急）では、金融緩和で利子率を下げても設備投資は増えにくく金融政策の効果は限定的。逆にこの経済では財政政策の方がGDP押し上げ効果が大きい。", options: { fontFace: F_BODY, fontSize: 9.5, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.78;
  s.addText("関連知識：IS・LM曲線の傾きの4決定要因と政策効果の対応は「その曲線自体が緩やかであるほど、もう一方の政策が効きやすい」という覚え方が有効。", {
    x: 0.55, y: cy, w: 12.25, h: 0.35, fontFace: F_BODY, fontSize: 9.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.42;
  s.addText("出典：past_exams/1st_stage/1ji2025/A1JI2025.pdf（令和7年度第1次試験）第10問／正解：past_exams/1st_stage_answers/r07/2025a.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 15: A-10 政府支出・租税と財政政策 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-10 ／ 政府支出・租税と財政政策",
    title: "裁量的財政政策とビルトイン・スタビライザー",
    overview: "財政政策は歳出・歳入を通じて景気を調整する。政府が発動する裁量的政策と、制度に組み込まれた自動調整機能がある。",
    tag: "経済学・経済政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "財政政策は歳出（政府支出）・歳入（租税）を通じて景気を調整する。不況期は拡張的、好況期は緊縮的に運用する。",
    { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.48;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "裁量的財政政策", v: "政府が意図的・都度判断で発動する政策", gap: 0.4 },
    { k: "ビルトイン・スタビライザー", v: "累進課税・失業保険給付など、制度に組み込まれた自動安定化機能", gap: 0.4 },
  ], { fontSize: 11.5, labelW: 2.6 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "不況期に所得税を減税しても、消費者が「将来の増税」を予想して減税分を貯蓄に回すと、期待ほど消費が増えない（リカードの中立命題）。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.93;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "ビルトイン・スタビライザーは「政府が都度判断して発動する政策」ではなく、税制・社会保障制度に組み込まれた自動的な機能。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.45, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：A-9 IS-LM分析（クラウディングアウトは財政政策の効果を弱める代表的な要因）。",
    years: mkYears(new Set(["'16", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 16: A-11 貨幣理論と金融政策 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-11 ／ 貨幣理論と金融政策",
    title: "中央銀行が持つ6つの金融政策手段",
    overview: "日本銀行が金利・通貨量を調整して物価安定・景気調整を図る手段。買いオペ＝資金供給（緩和）の方向を正確に覚える。",
    tag: "経済学・経済政策",
  });
  s.addText(
    "6つの手段は歴史的に発展した順に並ぶ。政策金利操作・公開市場操作・預金準備率操作という伝統的な手段だけでは金利がゼロに近づくと打つ手がなくなるため、量的緩和・マイナス金利・YCCという非伝統的な手段が追加されてきた。",
    { x: 0.55, y: 1.72, w: 12.25, h: 0.5, valign: "top", fontFace: F_BODY, fontSize: 10.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  let cy16 = addRowList(s, 0.55, 2.4, 12.25, [
    { name: "公開市場操作", tag: "オペレーション", desc: "国債等の売買で資金量を調整。買いオペ＝資金供給（緩和）、売りオペ＝資金吸収（引締）" },
    { name: "政策金利操作", tag: "無担保コールレート", desc: "中央銀行が誘導目標とする短期金利を操作" },
    { name: "預金準備率操作", tag: "準備率", desc: "金融機関が中央銀行に預け入れる比率を上下させ貸出余力を調整" },
    { name: "量的緩和政策", tag: "QE", desc: "政策金利がほぼゼロになった後、資金供給「量」そのものを拡大" },
    { name: "マイナス金利政策", tag: "", desc: "日銀当座預金の一部にマイナス金利を適用し貸出・投資を促す" },
    { name: "イールドカーブ・コントロール", tag: "YCC", desc: "長期金利（長期国債利回り）にも誘導目標を設定" },
  ], { rowH: 0.52, nameW: 3.3, tagW: 1.9 });
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "設備投資を計画する製造業A社にとって、日銀が政策金利を引き上げれば借入金利が上昇し投資判断は慎重になりやすい一方、量的緩和や利下げの局面では資金調達コストが下がり、投資に踏み切りやすくなる。金融政策の方向転換は企業の資金計画に直接影響する。", options: { fontFace: F_BODY, fontSize: 9.5, color: INK } },
  ], { x: 0.55, y: cy16 + 0.08, w: 12.25, h: 0.65, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：「買いオペは資金を吸収する」は誤り。買いオペは資金を「供給する」緩和的な政策。売りオペとの方向を逆に覚えないこと。",
    years: mkYears(new Set(["'19", "'20", "'21"])),
  });
}

// ---------- Slide 17: A-12 雇用と物価水準（フィリップス曲線等） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-12 ／ 雇用と物価水準（フィリップス曲線等）",
    title: "短期は右下がり、長期は垂直（自然失業率仮説）",
    overview: "フィリップス曲線はインフレ率と失業率のトレードオフを示す。ただし長期的にはこのトレードオフは成立しない。",
    tag: "経済学・経済政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "フィリップス曲線：縦軸インフレ率、横軸失業率の右下がりの関係。「失業率を下げようとするとインフレ率が高まる」という経験則。",
    { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.58;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "短期フィリップス曲線", v: "右下がり。インフレ率と失業率のトレードオフが観察される", gap: 0.4 },
    { k: "自然失業率仮説", v: "予想インフレ率が織り込まれると曲線が上方シフト。長期は自然失業率の水準で垂直になる", gap: 0.5 },
  ], { fontSize: 11.5, labelW: 2.1 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "1970年代の石油危機後、原油高でインフレが進む一方、不況で失業率も高止まりする「スタグフレーション」が発生し、単純なフィリップス曲線の想定が崩れた。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.75, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：「トレードオフは長期でも常に成立」は誤り。長期の曲線は自然失業率の水準で垂直になる。",
    years: mkYears(new Set(["'16", "'18", "'19", "'20", "'22", "'24"])),
  });
}

// ---------- Slide 18: A-13 景気変動と景気循環 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-13 ／ 景気変動と景気循環",
    title: "4つの景気循環：周期の短い順に原因とペアで覚える",
    overview: "キチン（在庫）＜ジュグラー（設備）＜クズネッツ（建築）＜コンドラチェフ（技術革新）の順に周期が長くなる。",
    tag: "経済学・経済政策",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "キチン・サイクル", tag: "約40か月", desc: "企業の在庫投資（在庫の積み増し・取り崩し）の変動による短期循環" },
    { name: "ジュグラー・サイクル", tag: "約7〜10年", desc: "企業の設備投資（生産設備の更新投資）の波による中期循環" },
    { name: "クズネッツの波", tag: "約15〜25年", desc: "住宅・商工業建築物の建て替え需要の変動（建築循環）" },
    { name: "コンドラチェフ・サイクル", tag: "約50年", desc: "鉄道・電力・情報通信技術のような大規模な技術革新の普及" },
  ], { rowH: 0.85, nameW: 2.9, tagW: 1.7 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：A-4 景気動向指数（DI・CI）とあわせて景気の波を把握する視点。周期と原因のペアを入れ替えた誤答が頻出。",
    years: mkYears(new Set(["'17", "'22", "'25"])),
  });
}

// ---------- Slide 19: A-13 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-13 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第6問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "経済学・経済政策",
  });
  const cy = addExamQuestion(s, {
    stem: "景気循環の周期性に関する記述の正誤の組み合わせとして、最も適切なものを下記の解答群から選べ。\n\nａ　コンドラチェフ・サイクルは、約50年の周期をもつ景気循環で、大規模な技術革新などに起因して生じると考えられている。\nｂ　キチン・サイクルは、約20年の周期をもつ景気循環で、住宅や商工業建築の建て替えなどに起因して生じると考えられている。\nｃ　ジュグラー・サイクルは、約7〜10年の周期をもつ景気循環で、生産設備の更新投資などに起因して生じると考えられている。",
    choices: [
      { badge: "ア", text: "ａ：正　ｂ：正　ｃ：正" },
      { badge: "イ", text: "ａ：正　ｂ：正　ｃ：誤" },
      { badge: "ウ", text: "ａ：正　ｂ：誤　ｃ：正" },
      { badge: "エ", text: "ａ：誤　ｂ：正　ｃ：正" },
      { badge: "オ", text: "ａ：誤　ｂ：正　ｃ：誤" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/A1JI2025.pdf（令和7年度第1次試験）第6問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 20: A-13 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-13 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第6問）",
    overview: "正解はウ。ｂはキチン・サイクルとクズネッツの波の説明が入れ替わっている。",
    tag: "経済学・経済政策",
  });
  const choices = [
    { badge: "ア", text: "ａ：正　ｂ：正　ｃ：正" },
    { badge: "イ", text: "ａ：正　ｂ：正　ｃ：誤" },
    { badge: "ウ", text: "ａ：正　ｂ：誤　ｃ：正" },
    { badge: "エ", text: "ａ：誤　ｂ：正　ｃ：正" },
    { badge: "オ", text: "ａ：誤　ｂ：正　ｃ：誤" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 2 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：ウ", options: { bold: true, color: RED } },
    { text: "（a正・b誤・c正）。ａ：コンドラチェフ・サイクルは約50年周期・技術革新が原動力＝正。ｂ：キチン・サイクルは本来「約40か月（3〜4年）」の在庫投資による短期循環。「約20年・住宅や商工業建築の建て替え」はクズネッツの波の説明であり、名称と内容が入れ替わっている＝誤。ｃ：ジュグラー・サイクルは約7〜10年・設備投資が原因＝正。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.0, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });
  cy += 1.08;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "企業が在庫を積みすぎ、その後在庫調整で生産を絞る数年単位のサイクルはキチン・サイクルの典型例。鉄道・電力の普及のような技術革新が経済を長期に押し上げる動きはコンドラチェフ・サイクルとして説明される。", options: { fontFace: F_BODY, fontSize: 9.5, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.78;
  s.addText("関連知識：4大景気循環（キチン・ジュグラー・クズネッツ・コンドラチェフ）は「周期の長さ×原因（在庫・設備・建築・技術革新）」のペアで暗記する。", {
    x: 0.55, y: cy, w: 12.25, h: 0.35, fontFace: F_BODY, fontSize: 9.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.42;
  s.addText("出典：past_exams/1st_stage/1ji2025/A1JI2025.pdf（令和7年度第1次試験）第6問／正解：past_exams/1st_stage_answers/r07/2025a.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 21: A-14 比較生産費説と貿易理論 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-14 ／ 比較生産費説と貿易理論",
    title: "「相対的な得意」で特化すると両国とも得をする",
    overview: "リカードの比較生産費説：絶対優位でなく比較優位（機会費用の小ささ）に基づいて特化・貿易すれば両国の生産量が増える。",
    tag: "経済学・経済政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "機会費用（ある財1単位の生産に必要な労働量 ÷ もう一方の財1単位に必要な労働量）が相手国より低い財に特化すると、貿易の利益が生まれる。",
    { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.58;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "絶対優位", v: "ある国が2財のどちらも他国より効率よく作れること（貿易の必要条件ではない）", gap: 0.4 },
    { k: "比較優位", v: "相対的に機会費用が低い財を持つこと。これに基づき特化すれば両国とも得をする", gap: 0.4 },
  ], { fontSize: 11.5, labelW: 1.5 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "カカオ1単位の労働量：A国5、B国4。大豆1単位：A国10、B国2。A国のカカオの機会費用＝5÷10＝0.5、B国＝4÷2＝2。A国はカカオに、B国は大豆に比較優位を持ち、それぞれ完全特化すると特化前より多く生産・消費できる。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.93;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "機会費用の式（自国の労働量÷もう一方の財の労働量）をどちらの財について聞かれているか取り違えないこと。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：A-23 経済余剰（貿易の利益を余剰の増加として捉える視点）。",
    years: mkYears(new Set(["'16", "'17", "'22", "'24"])),
  });
}

// ---------- Slide 22: A-14 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-14 ／ 過去問で確認する",
    title: "こう出題される（令和6年度 第22問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "経済学・経済政策",
  });
  const cy = addExamQuestion(s, {
    stem: "下表に従って、比較生産費説に基づく国際分業を考える。カカオ1単位を生産するのに必要な労働量は、Ａ国では5、Ｂ国では4である。同様に、大豆1単位を生産するのに必要な労働量は、Ａ国では10、Ｂ国では2である。労働は両国で同質で、当初はどちらの国もカカオと大豆をそれぞれ40単位ずつ生産していたものとする。このような状況に関する記述の正誤の組み合わせとして、最も適切なものを下記の解答群から選べ。\n\nａ　Ａ国におけるカカオ1単位の機会費用は、大豆2単位である。\nｂ　Ｂ国における大豆のカカオに対する相対価格は、Ａ国のそれよりも高い。\nｃ　Ｂ国で2つの財の生産に必要となる労働量の合計は240である。\nｄ　当初の労働量を維持しながら、Ａ国がカカオの生産に、Ｂ国が大豆の生産にそれぞれ完全特化したとき、各国におけるカカオと大豆の生産量はどちらも120となる。",
    choices: [
      { badge: "ア", text: "ａ：正　ｂ：正　ｃ：誤　ｄ：誤" },
      { badge: "イ", text: "ａ：正　ｂ：誤　ｃ：正　ｄ：誤" },
      { badge: "ウ", text: "ａ：誤　ｂ：正　ｃ：正　ｄ：誤" },
      { badge: "エ", text: "ａ：誤　ｂ：誤　ｃ：正　ｄ：正" },
      { badge: "オ", text: "ａ：誤　ｂ：誤　ｃ：誤　ｄ：正" },
    ],
    stemH: 1.85,
  });
  s.addText("出典：past_exams/1st_stage/1ji2024/A1JI2024.pdf（令和6年度第1次試験）第22問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 23: A-14 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-14 ／ 過去問で確認する",
    title: "解答＆解説（令和6年度 第22問）",
    overview: "正解はエ。機会費用の式（自国の労働量÷もう一方の財の労働量）を正確に適用できるかがカギ。",
    tag: "経済学・経済政策",
  });
  const choices = [
    { badge: "ア", text: "ａ：正　ｂ：正　ｃ：誤　ｄ：誤" },
    { badge: "イ", text: "ａ：正　ｂ：誤　ｃ：正　ｄ：誤" },
    { badge: "ウ", text: "ａ：誤　ｂ：正　ｃ：正　ｄ：誤" },
    { badge: "エ", text: "ａ：誤　ｂ：誤　ｃ：正　ｄ：正" },
    { badge: "オ", text: "ａ：誤　ｂ：誤　ｃ：誤　ｄ：正" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 3 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：エ", options: { bold: true, color: RED } },
    { text: "（a誤・b誤・c正・d正）。ａ：Ａ国のカカオ1単位の機会費用＝5÷10＝大豆0.5単位（「2単位」はＢ国の機会費用）＝誤。ｂ：大豆の相対価格（大豆の労働量÷カカオの労働量）はＡ国＝10÷5＝2、Ｂ国＝2÷4＝0.5。Ｂ国の方が低いため「Ｂ国の方が高い」は誤。ｃ：Ｂ国の総労働量＝4×40＋2×40＝240＝正。ｄ：Ａ国はカカオに600÷5＝120単位、Ｂ国は大豆に240÷2＝120単位、完全特化。両国とも120＝正。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.0, fontFace: F_BODY, fontSize: 10, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 1.08;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "特化前は両国合計でカカオ80単位・大豆80単位だったのに対し、特化後は各国120単位ずつを生産でき、貿易を通じてより多くの量を消費できる可能性が生まれる（貿易の利益）。", options: { fontFace: F_BODY, fontSize: 9.5, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.68;
  s.addText("関連知識：「各国の総労働量（一定）÷特化財の労働係数＝特化後の生産量」という計算パターンは頻出。", {
    x: 0.55, y: cy, w: 12.25, h: 0.35, fontFace: F_BODY, fontSize: 9.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.42;
  s.addText("出典：past_exams/1st_stage/1ji2024/A1JI2024.pdf（令和6年度第1次試験）第22問／正解：past_exams/1st_stage_answers/r06/2024a.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 24: A-15 貿易政策（関税・非関税障壁） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-15 ／ 貿易政策（関税・非関税障壁）",
    title: "保護貿易は社会全体の総余剰を減らす",
    overview: "関税・非関税障壁いずれも国内生産者を保護する一方、自由貿易と比べ社会全体の総余剰は小さくなる（死荷重の発生）。",
    tag: "経済学・経済政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "自由貿易は理論上、両国全体の余剰を最大化するが、各国政府は国内産業保護のため保護貿易政策をとることがある。",
    { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.48;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "関税", v: "輸入品への課税。国内価格↑で消費者余剰↓・生産者余剰↑・政府に関税収入。総余剰は死荷重の分だけ減少", gap: 0.6 },
    { k: "非関税障壁", v: "輸入数量割当（クオータ）、輸出自主規制、国内基準を使った実質的な輸入制限など", gap: 0.5 },
  ], { fontSize: 11.5, labelW: 1.6 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "輸入農産物に高関税をかけると、消費者は割高な価格を負担（消費者余剰↓）するが、国内農家は保護され（生産者余剰↑）、政府に関税収入が入る。社会全体の総余剰は自由貿易時より小さくなる。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.75, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：A-23 経済余剰（死荷重の考え方の土台）。",
    years: mkYears(new Set(["'17", "'18", "'19", "'21", "'23", "'25"])),
  });
}

// ---------- Slide 25: A-16 国際収支と為替レート決定理論 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-16 ／ 国際収支と為替レート決定理論",
    title: "為替レートの決定理論とマーシャル＝ラーナー条件",
    overview: "為替レートの変動要因を説明する2理論と、円安が貿易収支を改善させる条件（弾力性の和）を押さえる。",
    tag: "経済学・経済政策",
  });
  s.addText(
    "為替レートが「なぜ動くか」を説明する理論（購買力平価説・金利平価説）と、「動いた結果、貿易収支がどうなるか」を判定する条件（マーシャル＝ラーナー条件）は別の話。後者はさらに短期と長期で結果が違う（Jカーブ効果）点まで押さえる。",
    { x: 0.55, y: 1.72, w: 12.25, h: 0.5, valign: "top", fontFace: F_BODY, fontSize: 10.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  let cy25 = addRowList(s, 0.55, 2.4, 12.25, [
    { name: "購買力平価説", tag: "PPP", desc: "二国間の物価上昇率の差で為替レートが決まる。自国の物価上昇率が高いほど自国通貨は減価" },
    { name: "金利平価説", tag: "", desc: "内外の金利差が為替レートの予想変化率と等しくなるよう調整される" },
    { name: "マーシャル＝ラーナー条件", tag: "弾力性の和＞1", desc: "輸出入の価格弾力性の和が1超なら円安は貿易収支を改善、＝1なら不変、＜1なら悪化" },
    { name: "Jカーブ効果", tag: "", desc: "円安直後は数量が反応せず貿易収支が一時悪化し、時間経過とともに改善する現象" },
  ], { rowH: 0.85, nameW: 2.9, tagW: 2.1 });
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "輸出型の製造業A社は急な円安局面でも、契約済みの受注分の輸出数量をすぐには増やせないため円換算の売上高は当面あまり伸びず、数量ベースでの受注増加が本格化するのは半年〜1年後、というJカーブ効果を実務でも経験する。", options: { fontFace: F_BODY, fontSize: 9.5, color: INK } },
  ], { x: 0.55, y: cy25 + 0.05, w: 12.25, h: 0.65, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：弾力性の和＝1のケースを「悪化する」とする誤りが典型。正しくは「変化しない（中立）」。",
    years: mkYears(new Set(["'19", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 26: A-16 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-16 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第11問設問2）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "経済学・経済政策",
  });
  const cy = addExamQuestion(s, {
    stem: "円建ての貿易収支NXが、NX＝PX(e)－eP＊M(e)と表されるとする（X(e)は輸出量、M(e)は輸入量、eは円建て為替レート、Pは円建て輸出財価格（一定）、P＊はドル建て輸入財価格（一定））。為替レートの変化が貿易収支に及ぼす影響に関する記述の正誤の組み合わせとして、最も適切なものを下記の解答群から選べ。\n\nａ　為替レートが変化しても輸出量と輸入量は変化しないとき、為替レートの円安・ドル高への変化は、貿易収支を悪化させる。\nｂ　輸入の価格弾力性と輸出の価格弾力性がいずれも1より大きいとき、為替レートの円安・ドル高への変化は、貿易収支を改善させる。\nｃ　輸入の価格弾力性と輸出の価格弾力性の合計が1に等しいとき、為替レートの円安・ドル高への変化は、貿易収支を悪化させる。",
    choices: [
      { badge: "ア", text: "ａ：正　ｂ：正　ｃ：正" },
      { badge: "イ", text: "ａ：正　ｂ：正　ｃ：誤" },
      { badge: "ウ", text: "ａ：正　ｂ：誤　ｃ：正" },
      { badge: "エ", text: "ａ：誤　ｂ：正　ｃ：正" },
      { badge: "オ", text: "ａ：誤　ｂ：正　ｃ：誤" },
    ],
    stemH: 1.4,
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/A1JI2025.pdf（令和7年度第1次試験）第11問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 27: A-16 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-16 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第11問設問2）",
    overview: "正解はイ。ｃは「悪化」でなく「変化しない（中立）」が正しい。",
    tag: "経済学・経済政策",
  });
  const choices = [
    { badge: "ア", text: "ａ：正　ｂ：正　ｃ：正" },
    { badge: "イ", text: "ａ：正　ｂ：正　ｃ：誤" },
    { badge: "ウ", text: "ａ：正　ｂ：誤　ｃ：正" },
    { badge: "エ", text: "ａ：誤　ｂ：正　ｃ：正" },
    { badge: "オ", text: "ａ：誤　ｂ：正　ｃ：誤" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 1 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：イ", options: { bold: true, color: RED } },
    { text: "（a正・b正・c誤）。ａ：輸出入数量が一定なら円安は輸入代金の円換算額だけを機械的に押し上げ貿易収支は悪化＝正（Jカーブ効果の谷の局面と同じ理屈）。ｂ：弾力性の和が1超なら数量効果が価格効果を上回り貿易収支は改善＝正。ｃ：弾力性の和がちょうど1のときは数量効果と価格効果が相殺し貿易収支は変化しない（「悪化する」は誤り）。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.95, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });
  cy += 1.03;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "急激な円安直後は輸入企業が契約済み数量のまま購入を続け輸入代金（円換算）が膨らみ貿易収支が一時悪化する（Jカーブの谷）。時間をかけて数量調整が進むと貿易収支は改善に向かう。", options: { fontFace: F_BODY, fontSize: 9.5, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.78;
  s.addText("関連知識：マーシャル＝ラーナー条件は、短期は弾力性の和が小さいため悪化し、時間とともに改善するJカーブ効果とセットで理解する。", {
    x: 0.55, y: cy, w: 12.25, h: 0.35, fontFace: F_BODY, fontSize: 9.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.42;
  s.addText("出典：past_exams/1st_stage/1ji2025/A1JI2025.pdf（令和7年度第1次試験）第11問／正解：past_exams/1st_stage_answers/r07/2025a.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 28: A-17 国際マクロ経済の理論と政策 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-17 ／ 国際マクロ経済の理論と政策",
    title: "為替相場制度で財政・金融政策の効き方が正反対になる",
    overview: "マンデル＝フレミング・モデル：資本移動自由な小国では、変動相場制と固定相場制で政策効果が逆転する。",
    tag: "経済学・経済政策",
  });
  drawQuadrant(s, 0.55, 1.9, 12.25, 3.9, {
    cells: [
      { pos: "tl", label: "変動相場制×財政政策", sublabel: "効果はほぼ無効。利子率↑→自国通貨高→純輸出減で相殺" },
      { pos: "tr", label: "変動相場制×金融政策", sublabel: "効果は有効。利子率↓→自国通貨安→輸出増でGDP拡大" },
      { pos: "bl", label: "固定相場制×財政政策", sublabel: "効果は増幅。為替介入が金融緩和と同様の効果を持つ" },
      { pos: "br", label: "固定相場制×金融政策", sublabel: "効果は無効。為替介入で緩和効果が打ち消される" },
    ],
    axisCaption: "資本移動が自由な小国が前提。「変動相場制→財政無効・金融有効」「固定相場制→財政有効・金融無効」という対応関係が最重要。",
  });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：A-9 IS-LM分析（開放経済への拡張モデル）。",
    years: mkYears(new Set(["'18", "'19", "'20", "'21", "'22", "'23", "'24"])),
  });
}

// ---------- Slide 29: PART3 主要経済理論・市場メカニズム 区切り ----------
addDividerSlide(pres, {
  ghostNo: "03",
  partNo: "PART 03",
  partLabel: "経済学・経済政策 ／ A-18〜A-24",
  title: "主要経済理論とミクロ経済学",
  desc: "ケインズ派・古典派・マネタリズムという学派の違いを整理したのち、ミクロ経済学に入り、価格がどう決まり、市場がどれだけの豊かさ（余剰）を生むかを学ぶ。",
  chips: ["A-21 市場均衡", "A-23 経済余剰", "A-24 市場の失敗"],
  notes: "PART3区切りスライド。",
});

// ---------- Slide 30: A-18 ケインズ理論 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-18 ／ ケインズ理論",
    title: "有効需要の原理と政府の積極的介入",
    overview: "1930年代の世界恐慌を背景にケインズが打ち立てた理論。現代マクロ経済学（A-7〜A-13）の土台になっている。",
    tag: "経済学・経済政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "ケインズ経済学の中心的な考え方は2点。",
    { x: proseX, y: cy, w: proseW, h: 0.35, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0 }
  );
  cy += 0.43;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "有効需要の原理", v: "GDPの水準は供給側でなく需要（有効需要）の大きさで決まる", gap: 0.4 },
    { k: "市場調整の限界", v: "賃金・価格は下方硬直的で、非自発的失業が発生しうる。政府の財政・金融政策による介入が必要", gap: 0.55 },
  ], { fontSize: 11.5, labelW: 2.0 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "世界恐慌下でアメリカが実施したニューディール政策（公共事業による雇用創出）は「政府による有効需要の創出」を体現した代表例。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.68;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「市場に任せれば賃金・価格の調整で自動的に完全雇用が実現する」（A-19の古典派の立場）とは対照的。「非自発的失業は存在しない」はケインズ理論と正反対の主張。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "－", rankLabel: "集計データなし",
    related: "関連：A-7 生産物市場とGDP決定理論（有効需要の原理を45度線分析で具体化）。基礎理論として他の計算問題に組み込まれる形で出題されることが多い。",
    years: mkYears(new Set([])),
  });
}

// ---------- Slide 31: A-19 古典派・新古典派の理論 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-19 ／ 古典派・新古典派の理論",
    title: "ケインズと対照的な、市場メカニズムを信頼する立場",
    overview: "古典派経済学（ケインズ以前の主流派）は価格の伸縮性を信頼し、需要不足による不況は本来生じないと考える。",
    tag: "経済学・経済政策",
  });
  s.addText(
    "古典派の3つの考え方は別々の主張ではなく、いずれも「市場メカニズム（価格の伸縮性）を信頼すれば自動的に完全雇用が実現する」という一つの立場から出てくる帰結。ケインズ理論（A-18）とは正反対の前提に立つ点が対比の軸になる。",
    { x: 0.55, y: 1.72, w: 12.25, h: 0.5, valign: "top", fontFace: F_BODY, fontSize: 10.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  let cy31 = addRowList(s, 0.55, 2.4, 12.25, [
    { name: "セイの法則", tag: "供給が需要を創造", desc: "生産物は生産者の所得となりそのまま支出に回るため、需要不足の不況は本来生じない" },
    { name: "価格の伸縮性", tag: "自動調整", desc: "失業が発生しても賃金が下がり労働需要が回復、市場メカニズムで自動的に完全雇用が実現する" },
    { name: "古典派の二分法", tag: "貨幣の中立性", desc: "実物部門と貨幣部門は独立。貨幣供給量の変化は物価水準のみに影響し実物変数には影響しない" },
  ], { rowH: 1.0, nameW: 2.6, tagW: 2.3 });
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "人件費を抑えたい小売チェーンのB社が「不況なら賃金を下げれば人員整理をせずに済む」と考えるのは、賃金の伸縮性で自動的に均衡が回復するという古典派に近い発想。しかし実際には賃下げへの抵抗が大きく、市場任せでは失業が長期化しうる（ケインズ派の立場）という対立が人事戦略の前提にも現れる。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy31 + 0.1, w: 12.25, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "C", rankLabel: "出題実績あり",
    related: "ひっかけ：「貨幣供給量の変化は実物変数にも影響する」は誤り。古典派の二分法（貨幣の中立性）では物価水準のみに影響する。",
    years: mkYears(new Set(["'22"])),
  });
}

// ---------- Slide 32: A-20 マネタリズム ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-20 ／ マネタリズム",
    title: "「裁量」より「k%ルール」を重視するフリードマンの立場",
    overview: "マネタリズムは貨幣供給量（マネーストック）のコントロールを最重要視し、裁量的な政策運営に否定的な立場をとる。",
    tag: "経済学・経済政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "貨幣数量説（MV＝PT。M：貨幣供給量、V：流通速度、P：物価水準、T：取引量）が理論的土台。VとTが安定的なら、Mの増加はほぼそのままPの上昇（インフレ）につながる。",
    { x: proseX, y: cy, w: proseW, h: 0.65, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.73;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "k%ルール", v: "裁量的な政策運営でなく、貨幣供給量を一定率で機械的に増加させるべきという主張", gap: 0.4 },
    { k: "裁量的政策への批判", v: "政策効果が現れるまでのタイムラグの不確実性が、かえって景気を不安定化させるとする", gap: 0.4 },
  ], { fontSize: 11.5, labelW: 2.3 });
  cy += 0.1;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "マネタリズムは「政府による裁量的な財政・金融政策を積極活用すべき」というケインズ的立場とは正反対。「貨幣供給量より財政政策を重視する」という記述は誤り。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "－", rankLabel: "集計データなし",
    related: "関連：A-12 フィリップス曲線（自然失業率仮説はマネタリズムの立場からケインズ的総需要管理政策を批判する論拠）。",
    years: mkYears(new Set([])),
  });
}

// ---------- Slide 33: A-21 市場均衡・不均衡 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-21 ／ 市場均衡・不均衡",
    title: "需要曲線と供給曲線の交点で価格・数量が決まる",
    overview: "価格が自動調整されて需要と供給を一致させる働きが価格メカニズム。政府が上限・下限規制を課すと不均衡が解消されなくなる。",
    tag: "経済学・経済政策",
  });
  drawLineChart(s, 0.55, 1.9, 7.4, 4.2, {
    xLabel: "数量 Q", yLabel: "価格 P",
    series: [
      { x1: 0.05, y1: 0.85, x2: 0.95, y2: 0.1, color: INK, width: 2.2, label: "D", labelDy: -0.4 },
      { x1: 0.05, y1: 0.1, x2: 0.95, y2: 0.85, color: RED, width: 2.2, label: "S" },
    ],
    point: { nx: 0.5, ny: 0.475, label: "E" },
  });
  s.addText([
    { text: "超過需要", options: { bold: true } }, { text: "：価格が均衡より低く需要量＞供給量（品不足）。通常は価格上昇で解消\n\n", options: {} },
    { text: "超過供給", options: { bold: true } }, { text: "：価格が均衡より高く供給量＞需要量（売れ残り）。通常は価格下落で解消\n\n", options: {} },
    { text: "上限規制", options: { bold: true, color: RED } }, { text: "（家賃統制等）→超過需要が解消されない\n\n", options: {} },
    { text: "下限規制", options: { bold: true, color: RED } }, { text: "（最低賃金等）→超過供給（失業）が解消されない", options: {} },
  ], { x: 8.25, y: 2.0, w: 4.05, h: 4.0, fontFace: F_BODY, fontSize: 10, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：「超過需要が発生すると価格は下落する」は逆。超過需要は価格の上昇によって解消される。",
    years: mkYears(new Set(["'16", "'17", "'18", "'21", "'22", "'25"])),
  });
}

// ---------- Slide 34: A-22 弾力性の概念（価格弾力性等） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-22 ／ 弾力性の概念（価格弾力性等）",
    title: "弾力性の大きさで売上の増減が決まる",
    overview: "需要の価格弾力性＝｜需要量の変化率÷価格の変化率｜。1を境に、値下げが売上を増やすか減らすかが変わる。",
    tag: "経済学・経済政策",
  });
  s.addText(
    "弾力性の大きさ（1より大きいか小さいか）によって、値下げが売上を増やすか減らすかが逆転する。極端なケース（0・∞）は需要曲線の形（垂直・水平）とそのまま対応している。",
    { x: 0.55, y: 1.72, w: 12.25, h: 0.5, valign: "top", fontFace: F_BODY, fontSize: 10.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  addRowList(s, 0.55, 2.4, 12.25, [
    { name: "弾力性＞1", tag: "弾力的", desc: "価格を下げると支出総額（売上）は増加する" },
    { name: "弾力性＝1", tag: "単位弾力的", desc: "価格を変えても支出総額は変化しない" },
    { name: "弾力性＜1", tag: "非弾力的", desc: "価格を下げると支出総額は減少する" },
    { name: "弾力性＝0／∞", tag: "完全非弾力的／完全弾力的", desc: "需要曲線が垂直＝0、水平＝∞に対応" },
  ], { rowH: 0.85, nameW: 2.0, tagW: 3.0 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：「弾力性ゼロなら価格上昇でも支出総額は変化しない」は誤り。「長期の弾力性は短期より低い」も誤り（長期の方が高くなりやすい）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 35: A-22 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-22 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第13問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "経済学・経済政策",
  });
  const cy = addExamQuestion(s, {
    stem: "需要の価格弾力性（絶対値）に関する記述の正誤の組み合わせとして、最も適切なものを下記の解答群から選べ。\n\nａ　需要の価格弾力性がゼロであれば、価格の上昇によっても消費者の支出総額は変化しない。\nｂ　需要の価格弾力性が1より大きければ、価格の下落によって消費者の支出総額は増加する。\nｃ　同一の財について長期間で計った需要の価格弾力性は、短期間で計った場合よりも低くなりやすい。\nｄ　需要曲線が横軸に水平な直線である場合、需要の価格弾力性は無限大である。",
    choices: [
      { badge: "ア", text: "ａ：正　ｂ：正　ｃ：正　ｄ：誤" },
      { badge: "イ", text: "ａ：正　ｂ：誤　ｃ：正　ｄ：誤" },
      { badge: "ウ", text: "ａ：誤　ｂ：正　ｃ：誤　ｄ：正" },
      { badge: "エ", text: "ａ：誤　ｂ：誤　ｃ：正　ｄ：正" },
      { badge: "オ", text: "ａ：誤　ｂ：誤　ｃ：誤　ｄ：誤" },
    ],
    stemH: 1.15,
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/A1JI2025.pdf（令和7年度第1次試験）第13問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 36: A-22 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-22 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第13問）",
    overview: "正解はウ。ａ・ｃの「変化しない／低くなりやすい」という向きの誤りに注意。",
    tag: "経済学・経済政策",
  });
  const choices = [
    { badge: "ア", text: "ａ：正　ｂ：正　ｃ：正　ｄ：誤" },
    { badge: "イ", text: "ａ：正　ｂ：誤　ｃ：正　ｄ：誤" },
    { badge: "ウ", text: "ａ：誤　ｂ：正　ｃ：誤　ｄ：正" },
    { badge: "エ", text: "ａ：誤　ｂ：誤　ｃ：正　ｄ：正" },
    { badge: "オ", text: "ａ：誤　ｂ：誤　ｃ：誤　ｄ：誤" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 2 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：ウ", options: { bold: true, color: RED } },
    { text: "（a誤・b正・c誤・d正）。ａ：弾力性ゼロ＝数量が一切変化しない。数量一定のまま価格が上昇すれば支出総額は増加する（「変化しない」が誤り）。ｂ：弾力性が1超なら価格下落で需要量がそれ以上の割合で増え支出総額は増加＝正。ｃ：長期の方が代替財を探す余裕があり弾力性は高くなりやすい（「低くなりやすい」が誤り）。ｄ：需要曲線が水平＝価格のわずかな変化に需要量が無限に反応＝弾力性は無限大＝正。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.0, fontFace: F_BODY, fontSize: 10, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 1.08;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "生活必需品（米・電気）は代替が利きにくく非弾力的なため値上げで売上が増えやすいが、嗜好品・ブランド品は弾力的なため値上げするとかえって売上が減ることがある。", options: { fontFace: F_BODY, fontSize: 9.5, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.68;
  s.addText("関連知識：「弾力性＞1：弾力的（値下げで総収入増）／弾力性＜1：非弾力的（値下げで総収入減）」に加え、需要曲線の形状と弾力性の対応、長期は短期より弾力的という対応は繰り返し出題される定番論点。", {
    x: 0.55, y: cy, w: 12.25, h: 0.4, fontFace: F_BODY, fontSize: 9.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.47;
  s.addText("出典：past_exams/1st_stage/1ji2025/A1JI2025.pdf（令和7年度第1次試験）第13問／正解：past_exams/1st_stage_answers/r07/2025a.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 37: A-23 経済余剰（消費者余剰・生産者余剰） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-23 ／ 経済余剰（消費者余剰・生産者余剰）",
    title: "総余剰は均衡取引量で最大化される",
    overview: "消費者余剰＋生産者余剰＝総余剰。数量規制などで取引量が均衡水準より減ると、総余剰の一部が死荷重として失われる。",
    tag: "経済学・経済政策",
  });
  drawLineChart(s, 0.55, 1.9, 7.4, 4.2, {
    xLabel: "数量 Q", yLabel: "価格 P",
    series: [
      { x1: 0.05, y1: 0.85, x2: 0.95, y2: 0.1, color: INK, width: 2.2, label: "D", labelDy: -0.4 },
      { x1: 0.05, y1: 0.1, x2: 0.95, y2: 0.85, color: RED, width: 2.2, label: "S" },
    ],
    point: { nx: 0.5, ny: 0.475, label: "E" },
  });
  s.addText([
    { text: "■ 消費者余剰", options: { bold: true, color: "2E5495" } }, { text: "：買い手が得する部分（払ってもよい額－実際の支払額）\n\n", options: {} },
    { text: "■ 生産者余剰", options: { bold: true, color: "6B6B6B" } }, { text: "：売り手が得する部分（受取額－最低売りたい額）\n\n", options: {} },
    { text: "■ 死荷重", options: { bold: true, color: RED } }, { text: "：数量規制等で取引量が減ったとき失われる余剰", options: {} },
  ], { x: 8.25, y: 2.3, w: 4.05, h: 3.0, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.4 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：死荷重は「消費者余剰・生産者余剰の減少分の合計」ではなく、その減少分から税収等で社会に残った部分を差し引いた「純粋に失われた部分」。",
    years: mkYears(new Set(["'17", "'18", "'19", "'20", "'22", "'23", "'24"])),
  });
}

// ---------- Slide 38: A-24 競争的市場の資源配分機能・市場の失敗 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-24 ／ 競争的市場の資源配分機能・市場の失敗",
    title: "完全競争の理想が崩れる4つの「市場の失敗」",
    overview: "完全競争市場では総余剰が最大化される（パレート効率性）が、現実にはこの理想が実現しない市場の失敗がある。",
    tag: "経済学・経済政策",
  });
  s.addText(
    "完全競争市場ならA-23の総余剰最大化（パレート効率性）が実現するはずだが、現実にはそれを崩す要因が4つある。うち独占・寡占と情報の非対称性は、A-31・A-29でそれぞれ詳しく扱う内容の入口にあたる。",
    { x: 0.55, y: 1.72, w: 12.25, h: 0.5, valign: "top", fontFace: F_BODY, fontSize: 10.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  let cy38 = addRowList(s, 0.55, 2.4, 12.25, [
    { name: "外部性", tag: "外部不経済／外部経済", desc: "対価のやり取りなしに第三者に影響。悪影響（公害）は生産過大、良い影響（受粉）は過少になりがち" },
    { name: "公共財", tag: "非競合性・非排除性", desc: "フリーライダー問題により民間市場だけでは十分な量が供給されにくい" },
    { name: "独占・寡占", tag: "A-31参照", desc: "少数の供給者が価格支配力を持つと生産量過少・価格過大になる" },
    { name: "情報の非対称性", tag: "A-29参照", desc: "取引当事者間の情報格差" },
  ], { rowH: 0.85, nameW: 2.3, tagW: 2.7 });
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "化学メーカーのA社が排水処理コストを十分に負担しないまま生産すると、周辺住民への健康被害という外部不経済が発生し、社会的に望ましい水準より生産が過大になる。政府はこうした外部性を是正するため、排出量に応じた環境税（ピグー税）を課すことがある。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy38 + 0.08, w: 12.25, h: 0.75, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：公共財の「非競合性」（消費の奪い合いにならない）と「非排除性」（対価を払わない人を締め出せない）は別の性質。ピグー税・コースの定理も頻出。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 39: PART4 消費と生産の理論・組織と戦略の経済学・所得分配 区切り ----------
addDividerSlide(pres, {
  ghostNo: "04",
  partNo: "PART 04",
  partLabel: "経済学・経済政策 ／ A-25〜A-36",
  title: "消費と生産の理論・所得分配",
  desc: "消費者・企業それぞれの内部の意思決定メカニズムから、不完全競争市場の分析（ゲーム理論・独占・寡占）、最後は所得がどう分配されるかまで、ミクロ経済学の応用範囲を一気に扱う。",
  chips: ["A-27 生産関数・費用関数", "A-30 ゲーム理論", "A-31 独占・寡占"],
  notes: "PART4区切りスライド。",
});

// ---------- Slide 40: A-25 効用理論・予算制約と消費者行動 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-25 ／ 効用理論・予算制約と消費者行動",
    title: "無差別曲線と予算制約線が接する点で効用最大化",
    overview: "消費者は限られた予算の中で、効用（満足度）が最大になるように購入量を決める。",
    tag: "経済学・経済政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "同じ満足度が得られるX財・Y財の組み合わせを結んだ線が無差別曲線。予算内でできるだけ高い効用に到達しようとする。",
    { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.48;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "無差別曲線", v: "原点に対して凸型の右下がり。原点から遠いほど効用水準が高い", gap: 0.4 },
    { k: "限界代替率（MRS）", v: "X財1単位増の代わりに手放してよいY財の量。逓減する（限界代替率逓減の法則）", gap: 0.4 },
    { k: "消費者均衡", v: "限界代替率＝価格比（Px/Py）が成り立つ点。無差別曲線と予算制約線が接する", gap: 0.4 },
  ], { fontSize: 11, labelW: 2.3 });
  cy += 0.1;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「限界代替率逓減の法則」（他財との交換比率が逓減）と「限界効用逓減の法則」（その財自体の追加的満足度が逓減）は別概念。混同しないこと。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：A-26 需要曲線の導出（消費者均衡から需要曲線を導く）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'23", "'24"])),
  });
}

// ---------- Slide 41: A-26 需要曲線の導出 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-26 ／ 需要曲線の導出",
    title: "代替効果と所得効果、そして例外のギッフェン財",
    overview: "価格変化時の需要量の変化は代替効果と所得効果に分解できる。財の分類によって所得効果の向きが変わる。",
    tag: "経済学・経済政策",
  });
  s.addText(
    "価格が下がったとき需要量が増えるか減るかは、代替効果と所得効果という2つの力の合成で決まる。財の分類（上級財・下級財・ギッフェン財）は、この2つの力がどちらの向きに、どれだけ強く働くかの違いを表している。",
    { x: 0.55, y: 1.72, w: 12.25, h: 0.5, valign: "top", fontFace: F_BODY, fontSize: 10.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  let cy41 = addRowList(s, 0.55, 2.4, 12.25, [
    { name: "上級財（正常財）", tag: "所得↑→需要↑", desc: "代替効果・所得効果とも需要量を増やす方向。需要曲線は右下がり" },
    { name: "下級財（劣等財）", tag: "所得↑→需要↓", desc: "所得効果は需要を減らす方向だが通常は代替効果が上回り、やはり右下がり" },
    { name: "ギッフェン財", tag: "下級財の例外", desc: "所得効果が代替効果を上回る特殊な例外。価格が下がると需要量が減る（右上がり）" },
  ], { rowH: 1.0, nameW: 2.4, tagW: 2.1 });
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "食品メーカーのA社が主力商品を値下げすると、通常は代替効果・所得効果とも需要量を押し上げる方向に働き販売数量は増える。値下げで浮いた予算がより高級な代替品に回り、逆に自社商品の販売数量が減ることがある特殊なケースがギッフェン財として議論される。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy41 + 0.1, w: 12.25, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：「下級財は必ず右上がり」は誤り。下級財でも通常は右下がりで、代替効果を所得効果が上回る特殊例のみギッフェン財。",
    years: mkYears(new Set(["'16", "'21", "'25"])),
  });
}

// ---------- Slide 42: A-27 生産関数・費用関数 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-27 ／ 生産関数・費用関数",
    title: "限界生産力逓減の法則とU字型の費用曲線",
    overview: "資本を固定したまま労働投入を増やすと、生産量の増加分はやがて逓減する（限界生産力逓減の法則）。",
    tag: "経済学・経済政策",
  });
  s.addText(
    "労働を増やすほど追加的な生産量がだんだん小さくなる「限界生産力逓減の法則」が、短期の費用曲線がU字型を描く理由の土台。TC・AC・AVC・MCは互いに計算式でつながっている。",
    { x: 0.55, y: 1.72, w: 12.25, h: 0.5, valign: "top", fontFace: F_BODY, fontSize: 10.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  let cy42 = addRowList(s, 0.55, 2.4, 12.25, [
    { name: "総費用（TC）", tag: "固定費用＋可変費用", desc: "固定費用は生産量ゼロでも発生し、生産量に応じて変わらない" },
    { name: "平均費用（AC）", tag: "TC÷生産量", desc: "1単位あたりの総費用。限界生産力逓減の結果、短期はU字型になりやすい" },
    { name: "平均可変費用（AVC）", tag: "可変費用÷生産量", desc: "1単位あたりの可変費用" },
    { name: "限界費用（MC）", tag: "生産量+1の追加費用", desc: "平均費用曲線の最低点を通過する（MC＜ACならAC低下中、MC＞ACならAC上昇中）" },
  ], { rowH: 0.75, nameW: 2.3, tagW: 2.3 });
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "食品加工工場を持つ製造業A社が同じ設備のまま作業員だけを増やしていくと、当初は増員分だけ生産量が伸びるが、設備の処理能力に近づくにつれ1人追加あたりの生産量の伸びはだんだん小さくなる（限界生産力逓減）。その結果、増員を続けるほど1個あたりの平均費用はある時点から上昇に転じる。", options: { fontFace: F_BODY, fontSize: 9.5, color: INK } },
  ], { x: 0.55, y: cy42 + 0.08, w: 12.25, h: 0.75, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：「固定費用も生産量に応じて増減する」は誤り。固定費用は生産量ゼロでも発生し金額は変わらない。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 43: A-28 利潤最大化と供給曲線 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-28 ／ 利潤最大化と供給曲線",
    title: "利潤最大化条件はMR＝MC（完全競争ではP＝MC）",
    overview: "企業は限界収入（MR）＝限界費用（MC）となる生産量で利潤を最大化する。",
    tag: "経済学・経済政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "MR（限界収入）がMC（限界費用）を上回る間は生産を増やすほど利潤が増え、下回れば減る。MR＝MCとなる点で利潤が最大化される。",
    { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.58;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "完全競争企業", v: "価格受容者（プライス・テイカー）。MR＝市場価格Pのため、利潤最大化条件は「P＝MC」", gap: 0.4 },
    { k: "供給曲線", v: "完全競争企業の限界費用曲線そのもの（平均可変費用を上回る部分に限る）", gap: 0.4 },
  ], { fontSize: 11.5, labelW: 2.0 });
  cy += 0.1;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「企業は平均費用が最小になる生産量で利潤を最大化する」は誤り。利潤最大化条件はあくまで「限界収入＝限界費用」であり、平均費用最小の生産量とは一般に一致しない。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "B", rankLabel: "頻出論点",
    related: "関連：A-27 生産関数・費用関数（限界費用曲線の形状が供給曲線の土台）。",
    years: mkYears(new Set(["'16", "'22"])),
  });
}

// ---------- Slide 44: A-29 情報の不完全性（情報の非対称性） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-29 ／ 情報の不完全性（情報の非対称性）",
    title: "逆選択（契約前）とモラルハザード（契約後）",
    overview: "取引の一方だけが多くの情報を持つ状態が情報の非対称性。発生タイミングの違いで2つの問題に分かれる。",
    tag: "経済学・経済政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "情報の非対称性が引き起こす2つの問題は、発生タイミング（契約前か後か）で区別する。",
    { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.48;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "逆選択", v: "契約前に生じる。質の良い売り手ほど市場から撤退し、質の悪い商品ばかり残る（中古車の「レモン市場」）", gap: 0.55 },
    { k: "モラルハザード", v: "契約後に生じる。行動を観察されにくいことを利用し注意義務を怠る（火災保険加入後の油断など）", gap: 0.55 },
  ], { fontSize: 11.5, labelW: 1.7 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "中古車市場では買い手が車の状態を判断できず「平均的な質」の価格しか払わないため、質の良い車の売り手が市場から撤退し、質の悪い車ばかりが出回る（逆選択）。対策としてシグナリング・スクリーニングがある。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.75, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：「健康に不安のある人ほど保険に入りたがる」は逆選択、「保険加入後に不摂生になる」はモラルハザード。段階を確認する。",
    years: mkYears(new Set(["'19", "'22", "'23"])),
  });
}

// ---------- Slide 45: A-30 ゲーム理論（ナッシュ均衡、囚人のジレンマ等） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-30 ／ ゲーム理論（ナッシュ均衡、囚人のジレンマ等）",
    title: "利得表から「動く誘因がない」組み合わせを探す",
    overview: "ナッシュ均衡＝相手の戦略を所与としたとき、自分だけ変えても利得が改善しない組み合わせ。両者について最適反応を機械的に洗い出す。",
    tag: "経済学・経済政策",
  });
  drawQuadrant(s, 0.55, 1.9, 12.25, 3.9, {
    cells: [
      { pos: "tl", label: "守る・守る", sublabel: "（50、40）" },
      { pos: "tr", label: "守る・破る", sublabel: "（－20、60）" },
      { pos: "bl", label: "破る・守る", sublabel: "（60、－20）" },
      { pos: "br", label: "破る・破る", sublabel: "（0、0）★ナッシュ均衡" },
    ],
    axisCaption: "企業X（行）×企業Y（列）のカルテル利得表（カッコ内左＝X、右＝Y）。両者とも「破る」から動く誘因がなく、ここがナッシュ均衡（囚人のジレンマ構造）。",
  });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：「利得の合計が最大の組み合わせが必ずナッシュ均衡」は誤り。社会的に望ましい組み合わせとは限らない。",
    years: mkYears(new Set(["'17", "'18", "'20", "'22", "'23"])),
  });
}

// ---------- Slide 46: A-30 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-30 ／ 過去問で確認する",
    title: "こう出題される（令和5年度 第22問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "経済学・経済政策",
  });
  const cy = addExamQuestion(s, {
    stem: "特定の財の市場において競合関係にある企業同士が、同一価格での販売を約束するカルテルを結ぶことは、互いの企業にとって有利となる場合がある。企業Ｘと企業Ｙは、それぞれ一定の販売価格で合意したカルテルを守るか、あるいはそれを破ってより低い価格で販売するかを選択するものとする。下表は両企業の利得表であり、カッコ内の左側が企業Ｘの利得、右側が企業Ｙの利得を表している（企業X＼企業Y：カルテルを守る＝(50,40)／カルテルを破る＝(-20,60)。企業Xがカルテルを破る場合：カルテルを守る＝(60,-20)／カルテルを破る＝(0,0)）。このゲームに関する記述として、最も適切な組み合わせを下記の解答群から選べ。\n\nａ　企業Ｘが「カルテルを守る」場合において、企業Ｙの最適反応は「カルテルを破る」である。\nｂ　企業Ｙが「カルテルを守る」場合において、企業Ｘの最適反応は「カルテルを守る」である。\nｃ　このゲームにおけるナッシュ均衡は、企業Ｘ、企業Ｙともに「カルテルを守る」ケースである。\nｄ　このゲームにおけるナッシュ均衡は、企業Ｘ、企業Ｙともに「カルテルを破る」ケースである。",
    choices: [
      { badge: "ア", text: "ａとｃ" },
      { badge: "イ", text: "ａとｄ" },
      { badge: "ウ", text: "ｂとｃ" },
      { badge: "エ", text: "ｂとｄ" },
    ],
    stemH: 2.0,
  });
  s.addText("出典：past_exams/1st_stage/1ji2023/A1JI2023.pdf（令和5年度第1次試験）第22問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 47: A-30 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-30 ／ 過去問で確認する",
    title: "解答＆解説（令和5年度 第22問）",
    overview: "正解はイ（ａとｄ）。ｂの最適反応の向きの取り違えに注意。",
    tag: "経済学・経済政策",
  });
  const choices = [
    { badge: "ア", text: "ａとｃ" },
    { badge: "イ", text: "ａとｄ" },
    { badge: "ウ", text: "ｂとｃ" },
    { badge: "エ", text: "ｂとｄ" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 1 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：イ（ａとｄ）", options: { bold: true, color: RED } },
    { text: "。ａ：Xが「守る」のときYの利得は守る40・破る60→Yの最適反応は「破る」＝正。ｂ：Yが「守る」のときXの利得は守る50・破る60→Xの最適反応は「破る」（「守る」としている点が誤り）。ｃ：（守る,守る）はXに「破る」へ逸脱する誘因があるためナッシュ均衡でない＝誤。ｄ：（破る,破る）はX・Yともに0→守るに変えると－20になり動く誘因がない＝ナッシュ均衡＝正。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.2, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });
  cy += 1.28;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "両者がカルテルを守れば高い利得（50,40）を得られるにもかかわらず、各企業には個別に裏切る誘因が働くため、社会的に望ましくない（破る,破る）がナッシュ均衡として実現する。典型的な囚人のジレンマ構造。", options: { fontFace: F_BODY, fontSize: 9.5, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.68;
  s.addText("関連知識：ナッシュ均衡を求める手順は「各プレイヤーについて相手の各戦略に対する自分の最適反応を洗い出し、双方の最適反応が一致するマスを探す」という機械的な作業として身につける。", {
    x: 0.55, y: cy, w: 12.25, h: 0.35, fontFace: F_BODY, fontSize: 9.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.42;
  s.addText("出典：past_exams/1st_stage/1ji2023/A1JI2023.pdf（令和5年度第1次試験）第22問／正解：past_exams/1st_stage_answers/r05/2023a.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 48: A-31 独占の弊害と寡占化の協調行動 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-31 ／ 独占の弊害と寡占化の協調行動",
    title: "独占は「価格支配者」、生産量過少・価格過大が弊害",
    overview: "独占企業もMR＝MCで生産量を決めるが、右下がりの需要曲線に直面するため完全競争より生産量は過少・価格は過大になる。",
    tag: "経済学・経済政策",
  });
  s.addText(
    "独占企業もMR＝MCで生産量を決める点は完全競争と同じだが、右下がりの需要曲線に直面するため生産量過少・価格過大になる（独占の弊害）。企業が複数いる寡占市場では、これに企業同士の駆け引き（数量か価格か）と協調行動が加わる。",
    { x: 0.55, y: 1.72, w: 12.25, h: 0.5, valign: "top", fontFace: F_BODY, fontSize: 10.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  let cy48 = addRowList(s, 0.55, 2.4, 12.25, [
    { name: "独占の弊害", tag: "価格支配者", desc: "完全競争と比べ生産量過少・価格過大。総余剰が減少し死荷重が発生" },
    { name: "クールノー・モデル", tag: "数量競争", desc: "各企業が相手企業の生産量を所与として自社の最適生産量を決定" },
    { name: "ベルトラン・モデル", tag: "価格競争", desc: "各企業が相手企業の価格を所与として自社の最適価格を決定" },
    { name: "価格の下方硬直性", tag: "屈折需要曲線", desc: "値上げは追随されず顧客減、値下げは追随され値下げ競争になるため現状維持しがち" },
  ], { rowH: 0.75, nameW: 2.5, tagW: 1.9 });
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "家電量販店のように大手数社が価格を意識し合う寡占市場では、A社が値下げに踏み切ると他社も追随し値下げ競争になりやすい一方、A社が値上げしても他社が追随しなければA社だけ顧客を失うため、各社とも値上げには踏み切りにくい（価格の下方硬直性）。", options: { fontFace: F_BODY, fontSize: 9.5, color: INK } },
  ], { x: 0.55, y: cy48 + 0.08, w: 12.25, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：独占企業の価格はMRでなく需要曲線の高さで決まる。「独占企業はP＝MCで生産」は完全競争の条件。",
    years: mkYears(new Set(["'16", "'18", "'19", "'20", "'21", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 49: A-32 製品差別化と独占的競争 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-32 ／ 製品差別化と独占的競争",
    title: "完全競争と独占の中間、長期的には超過利潤ゼロに収束",
    overview: "独占的競争市場：多数の企業＋差別化された製品＋自由な参入退出。短期の超過利潤は新規参入で長期的に消える。",
    tag: "経済学・経済政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "独占的競争市場は完全競争と独占の中間的性質を持つ。飲食店やアパレルなど身の回りの多くの業種がこれにあたる。",
    { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.48;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "完全競争と同じ点", v: "多数の企業が存在し、参入・退出が自由", gap: 0.4 },
    { k: "独占と同じ点", v: "製品が差別化されているため各社がある程度の価格支配力を持つ", gap: 0.4 },
    { k: "長期均衡", v: "新規参入で需要曲線が左シフトし続け、最終的に利潤ゼロの水準に落ち着く", gap: 0.4 },
  ], { fontSize: 11, labelW: 2.1 });
  cy += 0.1;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "独占的競争は「独占と同じく長期的にも高い超過利潤を得続けられる」市場ではない。参入自由のため長期的には利潤ゼロに収束する。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "B", rankLabel: "頻出論点",
    related: "関連：A-31 独占の弊害（独占との違いを対比して理解する）。",
    years: mkYears(new Set(["'20", "'22"])),
  });
}

// ---------- Slide 50: A-33 規模の経済性・範囲の経済性 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-33 ／ 規模の経済性・範囲の経済性",
    title: "「同じ製品を多く」か「複数の製品を一緒に」か",
    overview: "規模の経済性は生産量、範囲の経済性は複数事業の組み合わせに着目する、別の着眼点の概念。",
    tag: "経済学・経済政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "規模の経済性・範囲の経済性はどちらも「大企業ほど有利」に見えるが、着眼点が異なる別概念。",
    { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.48;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "規模の経済性", v: "生産量（事業規模）が大きいほど1単位あたりの平均費用が下がる（固定費用の分散が主因）", gap: 0.5 },
    { k: "範囲の経済性", v: "複数の異なる製品・事業を1社で手がける方が総費用が安く済む（設備・ノウハウの共有）", gap: 0.5 },
  ], { fontSize: 11.5, labelW: 1.9 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "自動車メーカーが生産台数を増やすほど1台あたりの設備費負担が薄まるのが規模の経済性。同じ工場・エンジン技術で乗用車とトラックも作る方が別会社より安く済むのが範囲の経済性。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.75, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "C", rankLabel: "出題実績あり",
    related: "関連：C-3 多角化・シナジー（範囲の経済性の考え方と関連が深い概念）。",
    years: mkYears(new Set(["'16"])),
  });
}

// ---------- Slide 51: A-34 公正性・公平性の概念 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-34 ／ 公正性・公平性の概念",
    title: "ジニ係数は0（完全平等）〜1（完全不平等）",
    overview: "所得分配の公平性を測る道具がローレンツ曲線とジニ係数。値の大小の向きを取り違えないこと。",
    tag: "経済学・経済政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "ローレンツ曲線：横軸に所得の低い世帯からの累積比率、縦軸にその累積所得比率をとった曲線。",
    { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.48;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "ローレンツ曲線", v: "完全平等なら45度の均等分配線に一致。不平等なほど均等分配線より下に膨らむ", gap: 0.4 },
    { k: "ジニ係数", v: "均等分配線とローレンツ曲線に囲まれた面積から算出。0（完全平等）〜1（完全不平等）", gap: 0.4 },
    { k: "垂直的公平／水平的公平", v: "異なる状況の人には異なる扱い（累進課税）／同じ状況の人には同じ扱い", gap: 0.4 },
  ], { fontSize: 11, labelW: 2.3 });
  cy += 0.1;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「ジニ係数が1に近いほど平等」は逆。正しくは0に近いほど平等、1に近いほど不平等。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "B", rankLabel: "頻出論点",
    related: "関連：A-36 所得再分配と税制（ジニ係数は再分配政策の効果測定にも使われる）。",
    years: mkYears(new Set(["'20", "'25"])),
  });
}

// ---------- Slide 52: A-35 生産要素市場と生産要素報酬 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-35 ／ 生産要素市場と生産要素報酬",
    title: "賃金は労働の限界生産力価値で決まる（限界生産力説）",
    overview: "生産要素（労働・資本）そのものが取引される市場。賃金は生産要素が生み出す追加的な価値によって決まる。",
    tag: "経済学・経済政策",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "企業が利潤最大化するように労働者の雇用量を決める条件は、追加で得られる収入と支払う賃金が等しい点。",
    { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  cy += 0.48;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "労働の限界生産力価値", v: "労働を1単位追加で雇うことで得られる収入の増加分（限界生産物×製品価格）", gap: 0.45 },
    { k: "限界生産力説", v: "賃金率＝労働の限界生産力価値。生産要素への報酬はその要素が生む追加的価値で決まる", gap: 0.45 },
  ], { fontSize: 11.5, labelW: 2.3 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "もう1人雇うことで得られる生産物の増加分の市場価値が月30万円なら、企業が支払ってよい賃金の上限はおおよそ月30万円になる。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.68;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「賃金は労働者の生活費を基準に決まる」という考え方と、経済学の「賃金は労働の限界生産力で決まる」（限界生産力説）を混同しないこと。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：A-28 利潤最大化（MR＝MCの考え方を生産要素市場に応用したもの）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'20", "'21", "'23", "'24"])),
  });
}

// ---------- Slide 53: A-36 所得再分配と税制・補助金 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "A-36 ／ 所得再分配と税制・補助金",
    title: "累進課税・社会保障・負の所得税で格差を是正する",
    overview: "市場メカニズムに任せた結果の所得格差を是正するため、政府は税制・社会保障で所得の再分配を行う。",
    tag: "経済学・経済政策",
  });
  s.addText(
    "市場に任せた結果生じる所得格差（A-34）を政府が是正する手段は、大きく「税でならす」（累進課税）と「給付でならす」（社会保障・負の所得税）の2系統に整理できる。",
    { x: 0.55, y: 1.72, w: 12.25, h: 0.4, valign: "top", fontFace: F_BODY, fontSize: 10.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 }
  );
  let cy53 = addRowList(s, 0.55, 2.4, 12.25, [
    { name: "累進課税制度", tag: "垂直的公平", desc: "所得が高いほど税率も高くなる仕組み。所得税・相続税が代表例" },
    { name: "社会保障制度", tag: "", desc: "年金・医療保険・生活保護・失業給付などを通じた所得移転" },
    { name: "負の所得税", tag: "", desc: "一定所得を下回る世帯に給付を行う（マイナスの税を課す）という制度案" },
  ], { rowH: 1.0, nameW: 2.5, tagW: 1.7 });
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "製造業A社の給与計算担当者は、累進課税により高所得の管理職ほど源泉徴収の税率区分（限界税率）が上がる点を踏まえて手取り額を説明する必要がある。企業が負担する厚生年金・健康保険料も、社会保障制度を通じた所得再分配の一部を企業側が担っている側面がある。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy53 + 0.1, w: 12.25, h: 0.85, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：「消費税は高所得者ほど負担割合が高い累進的な税」は誤り。消費税は低所得者ほど所得に対する負担割合が高い逆進的な税。",
    years: mkYears(new Set(["'17", "'22", "'24", "'25"])),
  });
}

const outPath = path.join(__dirname, "..", "..", "slides", "1st_stage", "A_economics.pptx");
pres.writeFile({ fileName: outPath }).then(() => {
  console.log("wrote", outPath);
});
