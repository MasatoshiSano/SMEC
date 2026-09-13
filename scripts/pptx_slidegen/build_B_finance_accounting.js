// Builds slides/1st_stage/B_finance_accounting.pptx from lib.js components.
// Run: node scripts/pptx_slidegen/build_B_finance_accounting.js
//
// Mirrors slides/1st_stage/B_finance_accounting.html. See docs/13_pptx_slide_template_spec.md.

const path = require("path");
const {
  newPres, addCoverSlide, addDividerSlide, addHeader, addFreqBar, mkYears,
  addTermRows, addRowList, addExamQuestion, drawLineChart,
  INK, INK_SOFT, RED, LINE, GHOST, F_HEAD, F_BODY, F_MONO,
} = require("./lib");

const pres = newPres();
pres.author = "SMEC";
pres.title = "財務・会計";

// ---------- Slide 01: Cover ----------
addCoverSlide(pres, {
  eyebrow: "中小企業診断士 第1次試験",
  subjectNo: "SUBJECT 02 / 07",
  title: "財務・会計",
  subtitle: "簿記・会計の基礎から原価計算、経営分析、投資評価、企業価値まで34論点。計算問題を自分の手で解けるようになることが最重要な、2次試験事例IVにも直結する科目。",
  stats: [["34", "収録論点数"], ["6", "分野"], ["27/34", "頻出ランクA"]],
  tag: "docs/textbook/B_finance_accounting_textbook.md",
  notes: "表紙スライド。財務・会計の全体像を示す。",
});

// ---------- Slide 02: 簿記・会計の基礎 区切り ----------
addDividerSlide(pres, {
  ghostNo: "01",
  partNo: "PART 01",
  partLabel: "財務・会計 ／ B-1〜B-11",
  title: "簿記・会計の基礎／原価計算",
  desc: "会社のお金の動きを記録し、財務諸表にまとめる技術（簿記）と、財務諸表そのもの（B/S・P/L・C/F）、そして社内向けのコスト情報を作る原価計算。以降すべての土台になる分野。",
  chips: ["B-1 仕訳の基本", "B-5 貸借対照表", "B-9 税効果会計"],
  notes: "簿記・会計の基礎パートの区切りスライド。",
});

// ---------- Slide 03: B-1 簿記原理・会計帳簿 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-1 ／ 簿記の基礎：簿記原理・会計帳簿",
    title: "取引を「借方」「貸方」に分けて記録する",
    overview: "複式簿記では、1つの取引を必ず「原因」と「結果」の2側面から記録する。",
    tag: "財務・会計",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "簿記は会社の取引を一定ルールで帳簿に記録する技術。複式簿記は、1つの取引を必ず借方（左）・貸方（右）の2面で記録する。",
    { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 }
  );
  cy += 0.46;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "資産・費用", v: "増えたら借方（左）に記録" },
    { k: "負債・純資産・収益", v: "増えたら貸方（右）に記録" },
  ], { fontSize: 11.5, labelW: 2.1, gap: 0.4 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "現金100万円で開業＝（借）現金100万／（貸）資本金100万。商品120万円を掛け売り＝（借）売掛金120万／（貸）売上120万。借方合計＝貸方合計は、どんな取引でも必ず一致する（貸借平均の原理）。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.75, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.83;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "借方・貸方が逆になっている誤答肢が定番。「資産・費用の増加＝借方」「負債・純資産・収益の増加＝貸方」の対応を体に覚え込ませる。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.45, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：B-2 決算処理一巡（仕訳の集大成）、B-3〜B-5 財務諸表の作成。",
    years: mkYears(new Set(["'16", "'18", "'19", "'20", "'21", "'22", "'23", "'24"])),
  });
}

// ---------- Slide 04: B-2 決算処理一巡 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-2 ／ 簿記の基礎：決算処理一巡",
    title: "試算表→決算整理→財務諸表の一連の流れ",
    overview: "1年に1回、それまでの取引記録をまとめて財務諸表を作る一連のプロセス。",
    tag: "財務・会計",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "①試算表", desc: "各勘定科目の残高を一覧化し、借方合計＝貸方合計になっているか検算する" },
    { name: "②決算整理", desc: "減価償却など、期中の記録だけでは反映されない項目を修正する仕訳を行う" },
    { name: "③精算表", desc: "試算表→決算整理→P/L・B/Sへの流れを1枚にまとめる（試験で頻出）" },
    { name: "④財務諸表", desc: "損益計算書・貸借対照表を作成する" },
    { name: "⑤帳簿の締切", desc: "収益・費用の勘定をゼロに戻し、儲けを純資産に振り替える" },
  ], { nameW: 2.0, rowH: 0.78 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "具体例：建物1,000万円（耐用20年、定額法）→減価償却費＝1,000万÷20年＝50万円／年。",
    years: mkYears(new Set(["'17", "'19", "'20", "'22", "'25"])),
  });
}

// ---------- Slide 05: B-3 貸借対照表・損益計算書の作成 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-3 ／ 簿記の基礎：貸借対照表・損益計算書の作成",
    title: "仕訳の集大成としてのB/S・P/L",
    overview: "決算整理後の残高試算表を、資産・負債・純資産＝B/S、収益・費用＝P/Lに振り分ける。",
    tag: "財務・会計",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "B-1・B-2で記録した仕訳の集大成が貸借対照表（B/S）と損益計算書（P/L）。決算整理後の残高試算表の勘定科目を、次のルールで2つの表に振り分ける。",
    { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 }
  );
  cy += 0.56;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "資産・負債・純資産", v: "→ 貸借対照表（B/S）へ" },
    { k: "収益・費用", v: "→ 損益計算書（P/L）へ（収益－費用＝当期純利益）" },
  ], { fontSize: 11.5, labelW: 2.1, gap: 0.4 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "P/Lで計算した「当期純利益」は、最終的にB/Sの純資産（利益剰余金）に加算される。2つの表は別々のものではなく、1つの決算処理から生まれる表裏一体の関係。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.68;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「当期純利益はB/Sのどこに反映されるか」（→純資産の利益剰余金に加算）というP/L・B/Sのつながりを問う問題が頻出。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  s.addShape("line", { x: 0.55, y: 6.55, w: 12.25, h: 0, line: { color: INK, width: 0.75 } });
  s.addShape("ellipse", { x: 0.55, y: 6.63, w: 0.36, h: 0.36, fill: { color: INK_SOFT }, line: { type: "none" } });
  s.addText("－", { x: 0.55, y: 6.63, w: 0.36, h: 0.36, align: "center", valign: "middle", fontFace: F_MONO, fontSize: 13, bold: true, color: "FFFFFF", isTextBox: true, margin: 0 });
  s.addText("集計データなし　／　関連：B-4 損益計算書、B-5 貸借対照表（それぞれの表の中身を詳しく見る）。", {
    x: 1.0, y: 6.63, w: 11.5, h: 0.36, valign: "middle", fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 06: B-4 損益計算書 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-4 ／ 企業会計の基礎：損益計算書",
    title: "5段階で儲けを積み上げる",
    overview: "1年間の経営成績を、儲けの種類ごとに5段階の利益で表す。",
    tag: "財務・会計",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "①売上総利益", tag: "売上高－原価", desc: "いわゆる「粗利」" },
    { name: "②営業利益", tag: "①－販管費", desc: "本業でどれだけ儲けたか" },
    { name: "③経常利益", tag: "②＋営業外損益", desc: "本業＋財務活動を含めた通常の儲け" },
    { name: "④税引前当期純利益", tag: "③＋特別損益", desc: "臨時的な損益も含めた最終的な儲け" },
    { name: "⑤当期純利益", tag: "④－法人税等", desc: "株主に帰属する最終利益" },
  ], { nameW: 2.7, tagW: 1.8, rowH: 0.78 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：「支払利息増で経常利益は落ちたが営業利益は好調」→本業は問題なく財務体質に課題、と読み解けるかがポイント。",
    years: mkYears(new Set(["'16", "'17", "'18", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 07: B-5 貸借対照表 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-5 ／ 企業会計の基礎：貸借対照表",
    title: "資産 ＝ 負債 ＋ 純資産",
    overview: "決算日時点での財政状態（何を持ち、何を返す必要があるか）を表す。",
    tag: "財務・会計",
  });
  const proseX = 0.55, proseW = 7.6;
  let cy = 1.85;
  s.addText(
    "左側（借方）に資産、右側（貸方）に負債・純資産を並べ、左右の合計は必ず一致する。資産・負債はさらに「流動」（1年以内）と「固定」（1年超）に分ける（ワン・イヤー・ルール）。",
    { x: proseX, y: cy, w: proseW, h: 0.7, fontFace: F_BODY, fontSize: 11, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 }
  );
  cy += 0.78;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "流動資産", v: "現金、売掛金、棚卸資産" },
    { k: "固定資産", v: "建物、機械、土地" },
    { k: "流動負債", v: "買掛金、短期借入金" },
    { k: "固定負債", v: "長期借入金、社債" },
  ], { fontSize: 11, labelW: 1.5, gap: 0.36 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「利益剰余金が多い＝現金が多い」は誤り。現金は資産の部に別途「現金預金」として計上される。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  // right column: B/S balance diagram
  const rx = 8.55, rw = 4.05;
  s.addShape("rect", { x: rx, y: 2.0, w: 1.5, h: 3.6, fill: { type: "none" }, line: { color: INK, width: 1.5 } });
  s.addText("資産", { x: rx, y: 1.68, w: 1.5, h: 0.3, align: "center", fontFace: F_BODY, fontSize: 11, bold: true, color: INK, isTextBox: true, margin: 0 });
  s.addShape("rect", { x: rx + 1.9, y: 2.0, w: 1.9, h: 1.7, fill: { type: "none" }, line: { color: INK, width: 1.25 } });
  s.addText("負債", { x: rx + 1.9, y: 2.65, w: 1.9, h: 0.4, align: "center", valign: "middle", fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0 });
  s.addShape("rect", { x: rx + 1.9, y: 3.9, w: 1.9, h: 1.7, fill: { type: "none" }, line: { color: RED, width: 1.25 } });
  s.addText("純資産", { x: rx + 1.9, y: 4.55, w: 1.9, h: 0.4, align: "center", valign: "middle", fontFace: F_BODY, fontSize: 10.5, color: RED, isTextBox: true, margin: 0 });
  s.addText("資産合計＝負債合計＋純資産合計", { x: rx - 0.3, y: 5.75, w: 4.6, h: 0.3, align: "center", fontFace: F_BODY, fontSize: 9.5, color: INK_SOFT, isTextBox: true, margin: 0 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：B-13 安全性分析（流動比率・自己資本比率等はこのB/Sから計算する）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 08: B-6 キャッシュ・フロー計算書 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-6 ／ 企業会計の基礎：キャッシュ・フロー計算書",
    title: "利益と現金の動きは一致しない",
    overview: "「現金の動き」だけを営業・投資・財務の3区分で表す書類。間接法の調整が頻出。",
    tag: "財務・会計",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "間接法：当期純利益からスタートし、現金の動きを伴わない項目を調整して営業CFを逆算する（試験で頻出）。",
    { x: proseX, y: cy, w: proseW, h: 0.35, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, isTextBox: true, margin: 0 }
  );
  cy += 0.42;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "減価償却費", v: "現金支出を伴わない（非資金費用）→ 加算" },
    { k: "売上債権の増加", v: "まだ現金回収していない → 減算" },
    { k: "棚卸資産の増加", v: "現金を払って在庫を増やした → 減算" },
    { k: "仕入債務の増加", v: "支払いを先延ばしできている → 加算" },
  ], { fontSize: 11.5, labelW: 2.3, gap: 0.4 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「資産の増加＝お金が寝ている＝マイナス」「負債の増加＝支払いを待ってもらえる＝プラス」とイメージで覚える。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：B-19 キャッシュ・フロー管理（FCFはこの営業CFの計算がベース）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'20", "'21", "'23", "'24"])),
  });
}

// ---------- Slide 09: B-7 企業結合 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-7 ／ 企業会計の基礎：企業結合",
    title: "買収額と時価純資産の差額が「のれん」",
    overview: "合併・会社分割・連結決算の会計処理。パーチェス法とのれんの扱いが頻出。",
    tag: "財務・会計",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "会社を買収する場合、原則パーチェス法で処理する。買収先の資産・負債を時価評価し、取得原価が時価純資産を上回った差額を「のれん」として資産計上、20年以内で規則的に償却する。",
    { x: proseX, y: cy, w: proseW, h: 0.6, fontFace: F_BODY, fontSize: 11, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 }
  );
  cy += 0.68;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "合併", v: "複数社が1社になる（吸収合併／新設合併）" },
    { k: "会社分割", v: "事業の一部を切り離し承継させる（吸収分割／新設分割）" },
    { k: "連結決算", v: "親会社が子会社を含めグループ全体を1社として財務諸表を作成" },
  ], { fontSize: 11, labelW: 1.7, gap: 0.36 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "買収額1,200万円、時価純資産1,000万円→のれん＝200万円。20年均等償却なら毎期10万円を費用計上。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.58, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.66;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「負ののれん」（取得原価＜時価純資産）は資産計上せず、発生した期の利益（特別利益）に一括計上する。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：B-32 M&Aにおける企業評価（買収額そのものの算定方法）。",
    years: mkYears(new Set(["'16", "'18", "'19", "'20", "'21", "'23", "'25"])),
  });
}

// ---------- Slide 10: B-8 会計ディスクロージャー ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-8 ／ 企業会計の基礎：会計ディスクロージャー",
    title: "速報性と詳細さのトレードオフ",
    overview: "株主・投資家・債権者への財務状況の開示。根拠法令とタイミングの対応が頻出。",
    tag: "財務・会計",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "決算短信", tag: "取引所規則", desc: "決算発表時に投資家へ最速で開示する速報資料" },
    { name: "有価証券報告書", tag: "金商法", desc: "事業年度ごとの最も詳細で公的な開示書類（決算短信より遅い）" },
    { name: "半期報告書", tag: "金商法", desc: "2024年4月以後、四半期報告書は廃止され第2四半期はこれに統合" },
    { name: "計算書類", tag: "会社法", desc: "株主総会に提出するB/S・P/L等（会社法上の開示義務）" },
  ], { nameW: 2.4, tagW: 1.6, rowH: 0.78 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：「根拠法令」と「開示タイミング」の対応関係が問われやすい。2024年度の四半期報告書廃止という制度変更も要注意。",
    years: mkYears(new Set(["'19", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 11: B-9 税務会計の基礎 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-9 ／ 税務会計の基礎",
    title: "会計上の利益と税務上の課税所得はズレる",
    overview: "益金・損金の申告調整と、そのズレのうち一時差異を調整する税効果会計。",
    tag: "財務・会計",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText([
    { text: "課税所得＝益金－損金。", options: { bold: true, color: INK } },
    { text: "会計上の当期純利益を出発点に、会計と税法のズレを加減算する（申告調整）。将来解消するズレ（一時差異）は税効果会計で調整する。", options: { color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.6, fontFace: F_BODY, fontSize: 11, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });
  cy += 0.68;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "将来減算一時差異", v: "将来、税金が減る効果 → 繰延税金資産を計上" },
    { k: "将来加算一時差異", v: "将来、税金が増える効果 → 繰延税金負債を計上" },
  ], { fontSize: 11.5, labelW: 2.3, gap: 0.4 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "実効税率30％。賞与引当金3,000千円（前期末）→繰延税金資産900千円。当期末3,300千円に増額→繰延税金資産990千円（増加額90千円）。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.58;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「将来、税金の支払いが減る権利」だから資産、「将来、税金の支払いが増える義務」だから負債、とイメージで結びつける。交際費の損金不算入超過等の「永久差異」は税効果会計の対象外。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：B-2 決算処理一巡（税効果会計も決算整理仕訳の一種）。",
    years: mkYears(new Set(["'17", "'19", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 12: B-9 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-9 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第9問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "財務・会計",
  });
  const cy = addExamQuestion(s, {
    stem: "税効果会計に関する記述として、最も適切なものはどれか。なお、法人税等の実効税率を30％として計算していることを前提とする。また、繰延税金資産の回収可能性は考慮しなくてよい。",
    stemH: 0.85,
    choices: [
      { badge: "ア", text: "期首に備品を購入し、減価償却方法を定額法とすると、取得原価10,000千円、残存価額1,000千円、会計上の耐用年数が5年、税務上の耐用年数が4年の場合、1年目の終わりには135千円の繰延税金負債が計上される。", h: 0.6 },
      { badge: "イ", text: "期首に備品を購入し、減価償却方法を定額法とすると、取得原価10,000千円、残存価額1,000千円、会計上の耐用年数が5年、税務上の耐用年数が6年の場合、1年目の終わりには90千円の繰延税金負債が計上される。", h: 0.6 },
      { badge: "ウ", text: "前期に計上した賞与引当金3,000千円が全額損金不算入となり、繰延税金資産が900千円計上されていたとする。当期末に賞与引当金3,300千円を設定し、同額が損金不算入になった場合、繰延税金資産は90千円だけ増加することになる。", h: 0.6 },
      { badge: "エ", text: "当期に積立金方式による圧縮記帳を行ったことにより将来減算一時差異が10,000千円生じた場合、当期末には3,000千円の繰延税金資産が計上される。", h: 0.6 },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/B1JI2025.pdf（令和7年度第1次試験）第9問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 13: B-9 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-9 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第9問）",
    overview: "正解はウ。前期末900千円→当期末990千円で増加額90千円。",
    tag: "財務・会計",
  });
  const choices = [
    { badge: "ア", text: "期首に備品を購入し、減価償却方法を定額法とすると、取得原価10,000千円、残存価額1,000千円、会計上の耐用年数が5年、税務上の耐用年数が4年の場合、1年目の終わりには135千円の繰延税金負債が計上される。", h: 0.6 },
    { badge: "イ", text: "期首に備品を購入し、減価償却方法を定額法とすると、取得原価10,000千円、残存価額1,000千円、会計上の耐用年数が5年、税務上の耐用年数が6年の場合、1年目の終わりには90千円の繰延税金負債が計上される。", h: 0.6 },
    { badge: "ウ", text: "前期に計上した賞与引当金3,000千円が全額損金不算入となり、繰延税金資産が900千円計上されていたとする。当期末に賞与引当金3,300千円を設定し、同額が損金不算入になった場合、繰延税金資産は90千円だけ増加することになる。", h: 0.6 },
    { badge: "エ", text: "当期に積立金方式による圧縮記帳を行ったことにより将来減算一時差異が10,000千円生じた場合、当期末には3,000千円の繰延税金資産が計上される。", h: 0.6 },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 2 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：ウ", options: { bold: true, color: RED } },
    { text: "。前期末の繰延税金資産＝3,000千円×30％＝900千円。当期末＝3,300千円×30％＝990千円。増加額＝990－900＝", options: {} },
    { text: "90千円", options: { bold: true } },
    { text: "で記述と一致。エ：積立金方式の圧縮記帳は会計上の帳簿価額が税務上より高くなるため将来加算一時差異が生じ、計上されるのは繰延税金", options: {} },
    { text: "負債", options: { bold: true } },
    { text: "（3,000千円）。イ：会計上の減価償却費1,800千円、税務上1,500千円。会計上の方が多く償却するため将来減算一時差異（繰延税金", options: {} },
    { text: "資産", options: { bold: true } },
    { text: "）が生じ、「繰延税金負債」としている点が誤り。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.0, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 1.08;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "税効果会計は「会計上の資産・負債」と「税務上の資産・負債」の差異（一時差異）に着目する。将来減算一時差異→繰延税金資産、将来加算一時差異→繰延税金負債、という対応をまず固定してから各選択肢の数値を検算する。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.78;
  s.addText("関連知識：アは数値としては会計上の計算と整合するが、他の選択肢との比較で最も明確に正しいのはウ。「向き」（資産か負債か）で先に消去できる選択肢を削るのが速い。", {
    x: 0.55, y: cy, w: 12.25, h: 0.4, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  cy += 0.48;
  s.addText("出典：past_exams/1st_stage/1ji2025/B1JI2025.pdf（令和7年度第1次試験）第9問／正解：past_exams/1st_stage_answers/r07/2025b.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 14: B-10 原価計算の種類 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-10 ／ 原価計算：原価概念・原価計算の種類と方法",
    title: "原価の3要素と3つの分類軸",
    overview: "材料費・労務費・経費の3要素。生産形態・算定タイミング・固定費の扱いで方法が分かれる。",
    tag: "財務・会計",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "個別 vs 総合", desc: "受注生産（造船・建設）は個別原価計算、大量見込生産（食品等）は総合原価計算" },
    { name: "実際 vs 標準", desc: "実際にかかったコストを事後集計するのが実際原価計算、あらかじめ標準原価を決め差異分析するのが標準原価計算" },
    { name: "全部 vs 直接", desc: "固定費を製品原価に含めるのが全部原価計算（制度会計）、含めないのが直接原価計算（社内判断・CVP分析用）" },
  ], { nameW: 2.3, rowH: 1.2 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：全部＝固定費を含める、直接＝含めない、の一点に集約される違い。財務諸表作成には全部原価計算のみが認められる。",
    years: mkYears(new Set(["'16", "'17", "'18", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 15: B-11 原価情報の利用 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-11 ／ 原価計算：原価情報の利用",
    title: "意思決定に使うべきコストだけに注目する",
    overview: "差額原価収益分析：選択肢によって変化するコスト・収益だけに注目する考え方。",
    tag: "財務・会計",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText([
    { text: "差額原価収益分析：", options: { bold: true, color: INK } },
    { text: "複数の選択肢を比較する際、選択肢によって変わる（差が出る）コスト・収益だけに注目して判断する。", options: { color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 11.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });
  cy += 0.48;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "埋没原価", v: "既に支払い回収不可能な過去の支出。意思決定に影響させてはいけない" },
    { k: "機会原価", v: "ある選択肢を選んだことで失った、他の選択肢を選んでいたら得られた利益" },
  ], { fontSize: 11.5, labelW: 1.7, gap: 0.4 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "特別注文（1,000個、1個300円）を、追加変動費250円・遊休設備利用で受注できるなら、固定費（埋没原価に近い）は判断に関係なく、1個50円の追加の儲け（限界利益）が生まれ受注すべきと判断できる。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.78, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.86;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "既に支払った過去のコストを「もったいないから」と将来の判断に引きずる「サンクコストの誤謬」が典型的なひっかけ。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：B-16 利益計画（プロダクト・ミックス）でも同じ「差額で考える」発想を使う。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 16: 経営分析／利益と資金の管理 区切り ----------
addDividerSlide(pres, {
  ghostNo: "02",
  partNo: "PART 02",
  partLabel: "財務・会計 ／ B-12〜B-19",
  title: "経営分析／利益と資金の管理",
  desc: "B/S・P/Lの数値から会社の実力（儲ける力・安全性・生産性）を読み解く経営分析と、利益を計画的に生み出し資金繰りを管理する分野。「儲け」と「手元の現金」は別物、という視点が核心。",
  chips: ["B-12 ROA・ROE", "B-15 CVP分析", "B-19 FCF"],
  notes: "経営分析／利益と資金の管理パートの区切りスライド。",
});

// ---------- Slide 17: B-12 収益性分析 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-12 ／ 経営分析：収益性分析",
    title: "ROEは3つの要素に分解できる",
    overview: "会社がどれだけ効率よく儲けているかを見る分析。デュポン分解でROEの中身を切り分ける。",
    tag: "財務・会計",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "ROA", v: "当期純利益 ÷ 総資産 × 100（総資産全体でどれだけ効率よく利益を生んだか）" },
    { k: "ROE", v: "当期純利益 ÷ 自己資本 × 100（株主のお金でどれだけ効率よく利益を生んだか）" },
    { k: "デュポン分解", v: "ROE ＝ 売上高純利益率 × 総資産回転率 × 財務レバレッジ" },
  ], { fontSize: 11.5, labelW: 2.0, gap: 0.42 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "売上高1億円・当期純利益500万円・総資産5,000万円・自己資本2,500万円→ROA＝10％、ROE＝20％。総資産回転率2回転×財務レバレッジ2倍×純利益率5％＝20％で検算一致。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.55, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.63;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "負債を多く使う（財務レバレッジが高い）会社ほど、ROEはROAより大きく上振れしやすい。ROAの分子は当期純利益／経常利益のどちらか、問題文の指示に従うこと。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：B-13 安全性分析、B-23 最適資本構成（財務レバレッジは共通概念）。",
    years: mkYears(new Set(["'18", "'19", "'23"])),
  });
}

// ---------- Slide 18: B-13 安全性分析 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-13 ／ 経営分析：安全性分析",
    title: "倒産せずに返済できる体力を測る",
    overview: "短期・長期それぞれの支払い能力を見る代表指標。分母・分子の組み合わせを正確に。",
    tag: "財務・会計",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "流動比率", tag: "流動資産÷流動負債", desc: "短期の支払い能力。200％以上が望ましいとされる" },
    { name: "当座比率", tag: "当座資産÷流動負債", desc: "棚卸資産を除いた厳しめの短期支払い能力" },
    { name: "自己資本比率", tag: "自己資本÷総資産", desc: "高いほど財務的に安定" },
    { name: "固定長期適合率", tag: "固定資産÷(自己資本+固定負債)", desc: "100％超は短期資金で固定資産を賄う危険な状態" },
  ], { nameW: 2.1, tagW: 3.1, rowH: 0.78 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：固定比率が100％超でも固定長期適合率が100％以下なら直ちに危険ではない。",
    years: mkYears(new Set(["'16", "'17", "'19", "'20", "'21", "'23", "'24"])),
  });
}

// ---------- Slide 19: B-14 生産性・成長性分析 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-14 ／ 経営分析：生産性・成長性分析",
    title: "1人あたりどれだけ付加価値を生んだか",
    overview: "限られた経営資源から生み出した付加価値を測る労働生産性と、規模の伸びを測る成長性分析。",
    tag: "財務・会計",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "労働生産性", v: "付加価値額 ÷ 従業員数" },
    { k: "控除法", v: "付加価値＝売上高－外部購入価値（材料費・外注費）" },
    { k: "加算法", v: "付加価値＝経常利益＋人件費＋金融費用＋賃借料＋租税公課＋減価償却費" },
    { k: "売上高増加率", v: "(当期売上高－前期売上高) ÷ 前期売上高 × 100" },
  ], { fontSize: 11, labelW: 2.0, gap: 0.4 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "売上高3億円、外部購入価値1億8,000万円、従業員20人→控除法の付加価値＝1億2,000万円、労働生産性＝600万円／人。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.58, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.66;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "労働生産性が高い＝従業員の給料が高いとは限らない（利益として会社に残る場合もある）。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：B-12 収益性分析（同じ財務諸表から異なる切り口で会社を読み解く）。",
    years: mkYears(new Set(["'18", "'22", "'23"])),
  });
}

// ---------- Slide 20: B-15 CVP分析 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-15 ／ 経営分析：CVP分析（損益分岐点分析）",
    title: "売上高がいくらで利益ゼロになるか",
    overview: "原価を変動費・固定費に分解し、損益分岐点売上高・安全余裕率・営業レバレッジを求める。",
    tag: "財務・会計",
  });
  const proseX = 0.55, proseW = 7.6;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "限界利益", v: "売上高－変動費" },
    { k: "損益分岐点売上高", v: "固定費 ÷ 限界利益率" },
    { k: "安全余裕率", v: "1－損益分岐点比率" },
    { k: "営業レバレッジ度", v: "限界利益 ÷ 営業利益" },
  ], { fontSize: 10.5, labelW: 2.1, gap: 0.42 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "固定費を削減し変動費を増やすと営業レバレッジは低下する。製造業は小売業より営業レバレッジが高い傾向。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.55, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  const rx = 8.55, rw = 4.05;
  drawLineChart(s, rx, 1.95, rw, 3.7, {
    xLabel: "", yLabel: "",
    series: [
      { x1: 0, y1: 0, x2: 1, y2: 1, color: INK, width: 2, label: "売上高", labelDx: -0.55, labelDy: -0.15 },
      { x1: 0, y1: 0.3, x2: 1, y2: 0.65, color: RED, width: 2, label: "総原価", labelDx: -0.55, labelDy: -0.15 },
    ],
    point: { nx: 0.46, ny: 0.46, label: "BEP" },
  });
  s.addText("損益分岐点＝売上高線と総原価線の交点", { x: rx - 0.3, y: 5.75, w: rw + 0.6, h: 0.3, align: "center", fontFace: F_BODY, fontSize: 9.5, color: INK_SOFT, isTextBox: true, margin: 0 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "具体例：売上500万・変動費200万・固定費240万→限界利益300万、損益分岐点400万、安全余裕率20％、営業レバレッジ5倍。",
    years: mkYears(new Set(["'16", "'18", "'20", "'21", "'22", "'24", "'25"])),
  });
}

// ---------- Slide 21: B-15 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-15 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第13問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "財務・会計",
  });
  const cy = addExamQuestion(s, {
    stem: "営業レバレッジ（オペレーティング・レバレッジ）に関する記述として、最も適切なものはどれか。",
    stemH: 0.55,
    choices: [
      { badge: "ア", text: "営業レバレッジが低い企業は、営業レバレッジが高い企業に比べて、売上高が減少しても利益が減少しにくい状態であるといえる。" },
      { badge: "イ", text: "営業レバレッジの状況は、営業利益と当期純利益から把握できる。" },
      { badge: "ウ", text: "営業レバレッジは、一般的に、製造業の企業よりも小売業の企業の方が高くなる傾向にある。" },
      { badge: "エ", text: "営業レバレッジは、固定費を削減して変動費を増やすことによって高めることができる。" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/B1JI2025.pdf（令和7年度第1次試験）第13問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 22: B-15 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-15 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第13問）",
    overview: "正解はア。固定費比率が低い企業ほど、売上変動の利益への影響は小さい。",
    tag: "財務・会計",
  });
  const choices = [
    { badge: "ア", text: "営業レバレッジが低い企業は、営業レバレッジが高い企業に比べて、売上高が減少しても利益が減少しにくい状態であるといえる。" },
    { badge: "イ", text: "営業レバレッジの状況は、営業利益と当期純利益から把握できる。" },
    { badge: "ウ", text: "営業レバレッジは、一般的に、製造業の企業よりも小売業の企業の方が高くなる傾向にある。" },
    { badge: "エ", text: "営業レバレッジは、固定費を削減して変動費を増やすことによって高めることができる。" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 0 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：ア", options: { bold: true, color: RED } },
    { text: "。営業レバレッジ度＝限界利益÷営業利益で、固定費比率が高いほど値は大きくなる。固定費比率が低い企業ほど、売上高が変動しても利益への影響は相対的に小さい。イ：営業レバレッジ度は「限界利益÷営業利益」で計算するもので、営業利益と当期純利益からは把握できない。ウ：製造業の方が小売業よりも営業レバレッジは", options: {} },
    { text: "高くなる", options: { bold: true } },
    { text: "傾向にあり記述は逆。エ：固定費を削減し変動費を増やすと、営業レバレッジはむしろ", options: {} },
    { text: "低下", options: { bold: true } },
    { text: "する。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.0, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 1.08;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "固定費比率が高い＝ハイリスク・ハイリターンな費用構造、という理解が土台。CVP分析（損益分岐点）の考え方と表裏一体。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.58;
  s.addText("関連知識：docs/07_key_formulas_and_frameworks.mdのCVP分析参照。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.38;
  s.addText("出典：past_exams/1st_stage/1ji2025/B1JI2025.pdf（令和7年度第1次試験）第13問／正解：past_exams/1st_stage_answers/r07/2025b.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 23: B-16 利益計画 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-16 ／ 利益と資金の管理：利益計画",
    title: "ボトルネック1単位あたりの限界利益で優先順位を決める",
    overview: "複数製品を作る際、制約資源1単位あたりの限界利益が高い製品を優先する。",
    tag: "財務・会計",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText(
    "限られた生産能力（ボトルネック）をどの製品にどれだけ振り向けるか（プロダクト・ミックス）は、「1個あたりの限界利益」ではなく「制約資源1単位あたりの限界利益」で優先順位を決める。",
    { x: proseX, y: cy, w: proseW, h: 0.55, fontFace: F_BODY, fontSize: 11.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 }
  );
  cy += 0.63;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "加工機械480分／日がボトルネック。製品X＝限界利益1,000円・機械5分（200円/分）、製品Y＝限界利益1,500円・機械10分（150円/分）。1個あたりはYが高いが、機械1分あたりではXの方が効率的→Xを優先。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.7, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.78;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「1個あたりの限界利益が高い方を優先」という単純な発想は、制約条件がある場合は誤り。必ず制約資源1単位あたりで比較する。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "C", rankLabel: "直近10年で1回出題",
    related: "関連：B-11 原価情報の利用（差額原価収益分析と同じ「差で考える」発想）。",
    years: mkYears(new Set(["'16"])),
  });
}

// ---------- Slide 24: B-17 予算・実績差異分析 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-17 ／ 利益と資金の管理：予算・実績差異分析",
    title: "価格差異と数量差異に分けて原因を切り分ける",
    overview: "標準原価計算における差異分析。掛け合わせる数量（実際か標準か）の組み合わせがポイント。",
    tag: "財務・会計",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "価格差異", v: "(実際価格－標準価格) × 実際消費量" },
    { k: "数量差異", v: "(実際消費量－標準消費量) × 標準価格" },
  ], { fontSize: 11.5, labelW: 1.7, gap: 0.42 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "標準：1kgあたり500円、1個2kg。実際：520円、2.2kg消費。価格差異＝(520－500)×2.2kg＝44円（不利）。数量差異＝(2.2－2)×500円＝100円（不利）。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.55, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.63;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「価格のズレは実際消費量を掛ける」「数量のズレは標準価格で固定して掛ける」という掛け合わせる数量の組み合わせを逆にしやすい。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "C", rankLabel: "直近10年で1回出題",
    related: "関連：B-10 原価計算（標準原価計算の考え方がベース）。",
    years: mkYears(new Set(["'21"])),
  });
}

// ---------- Slide 25: B-18 資金繰りと資金計画 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-18 ／ 利益と資金の管理：資金繰りと資金計画",
    title: "利益が出ていても現金が足りなくなる",
    overview: "運転資金＝売上債権＋棚卸資産－仕入債務。急拡大局面ほど資金がショートしやすい。",
    tag: "財務・会計",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  s.addText([
    { text: "資金繰り：", options: { bold: true, color: INK } },
    { text: "事業継続に必要な現金がいつ・いくら必要かを予測し、ショートしないよう管理すること。中心は「運転資金」の考え方。", options: { color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.45, fontFace: F_BODY, fontSize: 11.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });
  cy += 0.53;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "運転資金", v: "売上債権＋棚卸資産－仕入債務" },
  ], { fontSize: 11.5, labelW: 1.7, gap: 0.4 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "売上債権2,000万・棚卸資産1,500万・仕入債務1,200万→運転資金＝2,300万円。売上急拡大で売掛金・在庫が膨らむと運転資金の必要額が増加する（黒字倒産の典型パターン）。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.78, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.86;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「利益が出ている＝現金に余裕がある」とは限らない。「勘定合って銭足らず」の状況を理解しておく。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "B", rankLabel: "直近10年で2回出題",
    related: "関連：B-19 キャッシュ・フロー管理（CCCは運転資金を日数で見る指標）。",
    years: mkYears(new Set(["'21", "'22"])),
  });
}

// ---------- Slide 26: B-19 キャッシュ・フロー管理 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-19 ／ 利益と資金の管理：キャッシュ・フロー管理",
    title: "FCF＝営業CF－投資CF",
    overview: "自由に使える現金がFCF。CCCは仕入から回収までの日数を測る資金効率の指標。",
    tag: "財務・会計",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "FCF", v: "営業キャッシュ・フロー－投資キャッシュ・フロー（設備投資額）" },
    { k: "CCC", v: "売上債権回転日数＋棚卸資産回転日数－仕入債務回転日数（短いほど資金効率が良い）" },
  ], { fontSize: 11.5, labelW: 1.3, gap: 0.42 });
  cy += 0.1;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "当期純利益1,000万・減価償却費200万・売上債権150万増加・仕入債務80万増加・設備投資300万→営業CF1,130万、FCF＝1,130－300＝830万円。", options: { fontFace: F_BODY, fontSize: 10.5, color: INK } },
  ], { x: proseX, y: cy, w: proseW, h: 0.55, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.63;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「売上債権の増加額は加算」「設備投資額は加算」はB-6の向きを逆にした典型的な誤り。売上債権増加→減算、設備投資額→減算が正しい向き。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.4, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "B", rankLabel: "直近10年で2回出題",
    related: "関連：B-6 キャッシュ・フロー計算書（間接法の調整がこのFCF計算のベース）。",
    years: mkYears(new Set(["'23", "'25"])),
  });
}

// ---------- Slide 27: B-19 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-19 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第21問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "財務・会計",
  });
  const cy = addExamQuestion(s, {
    stem: "当期純利益からフリー・キャッシュ・フローを計算する場合の記述として、最も適切なものはどれか。なお、税金は存在しないものとする。",
    stemH: 0.6,
    choices: [
      { badge: "ア", text: "売上債権の増加額は、当期純利益に加算される。" },
      { badge: "イ", text: "減価償却費は、当期純利益から減算される。" },
      { badge: "ウ", text: "仕入債務の増加額は、当期純利益に加算される。" },
      { badge: "エ", text: "設備投資額は、当期純利益に加算される。" },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/B1JI2025.pdf（令和7年度第1次試験）第21問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 28: B-19 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "B-19 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第21問）",
    overview: "正解はウ。仕入債務の増加＝支払いの先延ばし＝現金流出が抑えられる＝加算。",
    tag: "財務・会計",
  });
  const choices = [
    { badge: "ア", text: "売上債権の増加額は、当期純利益に加算される。" },
    { badge: "イ", text: "減価償却費は、当期純利益から減算される。" },
    { badge: "ウ", text: "仕入債務の増加額は、当期純利益に加算される。" },
    { badge: "エ", text: "設備投資額は、当期純利益に加算される。" },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 2 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：ウ", options: { bold: true, color: RED } },
    { text: "。仕入債務が増加する＝支払いを先延ばしできている＝現金流出が抑えられている、ということなので当期純利益に加算される。ア：売上債権が増加する＝まだ現金を回収できていない、ということなのでキャッシュベースでは", options: {} },
    { text: "減算", options: { bold: true } },
    { text: "しなければならない。イ：減価償却費は実際の現金支出を伴わない（非資金費用）ため、当期純利益に", options: {} },
    { text: "加算", options: { bold: true } },
    { text: "して戻す。エ：設備投資額は実際に現金が出ていく支出なので、当期純利益から", options: {} },
    { text: "減算", options: { bold: true } },
    { text: "する。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 1.0, fontFace: F_BODY, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 1.08;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "FCF＝営業CF－投資CF、間接法での営業CFは「当期純利益＋非資金費用－運転資本の増加＋運転資本の減少」という構造。", options: { fontFace: F_BODY, fontSize: 10, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.58;
  s.addText("関連知識：B-18の運転資金の定義（売上債権＋棚卸資産－仕入債務）と対応させて覚えること。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.38;
  s.addText("出典：past_exams/1st_stage/1ji2025/B1JI2025.pdf（令和7年度第1次試験）第21問／正解：past_exams/1st_stage_answers/r07/2025b.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

pres.writeFile({ fileName: path.join(__dirname, "../../slides/1st_stage/B_finance_accounting.pptx") })
  .then(() => console.log("wrote", "slides/1st_stage/B_finance_accounting.pptx"));
