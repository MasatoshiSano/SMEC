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
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "GDP", tag: "国内総生産", desc: "国内で生産された付加価値の合計（国籍を問わない）" },
    { name: "GNI／GNP", tag: "国民総所得", desc: "自国民（居住者）が生み出した所得の合計。GNI＝GDP＋海外からの純所得受取" },
    { name: "NNP", tag: "国民純生産", desc: "GNP（GNI）から固定資本減耗（減価償却相当）を差し引いたもの" },
    { name: "NI", tag: "国民所得", desc: "NNPから間接税を差し引き、補助金を加えたもの" },
  ], { nameW: 1.7, tagW: 2.0 });
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
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "CPI", tag: "消費者物価指数", desc: "総務省公表。消費者が購入する財・サービスの価格変動（ラスパイレス方式）" },
    { name: "CGPI", tag: "企業物価指数", desc: "日本銀行公表。企業間で取引される商品（原材料・中間財）の価格変動" },
    { name: "GDPデフレーター", tag: "名目GDP÷実質GDP×100", desc: "経済全体（消費・投資・政府支出・輸出入）の物価動向（パーシェ方式に近い）" },
  ], { rowH: 1.0, nameW: 2.6, tagW: 2.6 });
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
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "M1", tag: "現金＋預金通貨", desc: "いつでも引き出せる現金通貨と普通・当座預金" },
    { name: "M2", tag: "M1＋準通貨＋CD", desc: "対象は国内銀行・信用金庫等に限定（ゆうちょ銀行等は含まない）" },
    { name: "M3", tag: "範囲を拡大", desc: "ゆうちょ銀行・信用組合・農協なども含む全預金取扱機関が対象" },
    { name: "広義流動性", tag: "最も広い範囲", desc: "M3に金銭信託・投資信託・国債・外債等の金融商品を加えたもの" },
  ], { rowH: 0.75, nameW: 2.0, tagW: 2.2 });
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

const outPath = path.join(__dirname, "..", "..", "slides", "1st_stage", "A_economics.pptx");
pres.writeFile({ fileName: outPath }).then(() => {
  console.log("wrote", outPath);
});
