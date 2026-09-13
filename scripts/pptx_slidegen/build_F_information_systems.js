// Builds slides/1st_stage/F_information_systems.pptx from lib.js components.
// Run: node scripts/pptx_slidegen/build_F_information_systems.js
//
// Mirrors slides/1st_stage/F_information_systems.html. See docs/13_pptx_slide_template_spec.md.

const path = require("path");
const {
  newPres, addCoverSlide, addDividerSlide, addHeader, addFreqBar, mkYears,
  addTermRows, addRowList, addExamQuestion,
  INK, INK_SOFT, RED, LINE, GHOST, F_HEAD, F_BODY, F_MONO,
} = require("./lib");

const pres = newPres();
pres.author = "SMEC";
pres.title = "経営情報システム";

// ---------- Slide 01: Cover ----------
addCoverSlide(pres, {
  eyebrow: "中小企業診断士 第1次試験",
  subjectNo: "SUBJECT 06 / 07",
  title: "経営情報システム",
  subtitle: "ハードウェア・ソフトウェアからネットワーク・データベース、経営とITの結びつきまで24論点。「文系だから苦手」を克服し、得点源にできる科目。",
  stats: [["24", "収録論点数"], ["2", "分野"], ["21/24", "頻出ランクA"]],
  tag: "docs/textbook/F_information_systems_textbook.md",
  notes: "表紙スライド。経営情報システムの全体像を示す。",
});

// ---------- Slide 02: F-1〜F-11 区切り ----------
addDividerSlide(pres, {
  ghostNo: "01",
  partNo: "PART 01",
  partLabel: "経営情報システム ／ F-1〜F-11",
  title: "情報通信技術の基礎",
  desc: "コンピュータの仕組み（ハード・ソフト）から、プログラムの作り方、システム構成、データベース・ネットワークの基礎まで、ITの土台となる知識。",
  chips: ["F-9 正規化", "F-10 プロトコル", "F-11 信頼性設計"],
  notes: "情報通信技術に関する基礎的知識パートの区切りスライド。",
});

// ---------- Slide 03: F-1 ハードウェア ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-1 ／ ハードウェア（構成要素、機能と処理）",
    title: "主記憶は揮発性、補助記憶は不揮発性",
    overview: "RAIDレベルごとの最低必要台数・耐えられる故障台数が頻出。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "主記憶装置", v: "CPUが直接読み書きするRAM。電源を切るとデータが消える（揮発性）" },
    { k: "補助記憶装置", v: "HDD・SSD等。電源を切ってもデータは消えない（不揮発性）" },
    { k: "RAID1", v: "ミラーリング。1台故障しても継続稼働。実質容量は半分" },
    { k: "RAID5／RAID6", v: "最低3台／4台必要。1台／2台までの故障に耐える" },
  ], { fontSize: 11, labelW: 2.1, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「主記憶装置は不揮発性」は誤り。RAIDの最低必要台数・耐障害台数を数値で入れ替える出題に注意。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：F-11 システム評価（信頼性設計の考え方全般）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'23", "'25"])),
  });
}

// ---------- Slide 04: F-2 ソフトウェア ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-2 ／ ソフトウェア（OS、ミドルウェア、パッケージ）",
    title: "ミドルウェアはOSの一部ではない",
    overview: "OS→ミドルウェア→アプリケーションソフトという3層構造で理解する。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "OS", v: "ハードウェアを管理し複数アプリが動く土台。メモリ管理・タスク管理を担う" },
    { k: "ミドルウェア", v: "OSとアプリの中間層。DBMS・Webサーバソフト等、共通機能を提供" },
    { k: "パッケージソフト", v: "既製品。オーダーメイド（個別開発）ソフトと対比される" },
  ], { fontSize: 11.5, labelW: 2.1, gap: 0.42 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「ミドルウェアはOSの一部」は誤り。OSとは独立した中間層のソフトウェア。パッケージは既製機能の範囲内でのみカスタマイズ可能。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：F-8 外部情報資源の活用（クラウドで提供される範囲との対比）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'24"])),
  });
}

// ---------- Slide 05: F-3 プログラム設計 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-3 ／ プログラム設計（アルゴリズム、データ構造）",
    title: "スタックはLIFO、キューはFIFO",
    overview: "二分探索は事前にソート済みであることが前提。",
    tag: "経営情報システム",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "スタック", tag: "LIFO", desc: "後から入れたものを先に取り出す。積み重ねたお皿のイメージ" },
    { name: "キュー", tag: "FIFO", desc: "先に入れたものを先に取り出す。レジの待ち行列のイメージ" },
    { name: "二分探索", desc: "データがソート済みである前提で範囲を半分ずつ絞り込む（線形探索より高速）" },
    { name: "計算量", tag: "O(logn)<O(n)<O(n²)", desc: "データ量が増えたときの処理時間の増え方。左ほど効率が良い" },
  ], { nameW: 1.7, tagW: 2.3, rowH: 0.9 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：LIFO/FIFOの入出力順序、計算量の大小関係を逆に覚える誤りに注意。",
    years: mkYears(new Set(["'17", "'21", "'22", "'23"])),
  });
}

// ---------- Slide 06: F-4 ソフトウェア開発 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-4 ／ ソフトウェア開発（構造化、オブジェクト指向）",
    title: "カプセル化は複数クラスの統合ではない",
    overview: "構造化の3構造と、オブジェクト指向の3性質を区別する。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "構造化の3構造", v: "順次・選択（分岐）・反復（繰り返し）の組み合わせで表現" },
    { k: "カプセル化", v: "データと処理を1つのクラスにまとめ、外部から実装を隠す" },
    { k: "継承", v: "親クラスの性質を子クラスが引き継ぐこと" },
    { k: "ポリモーフィズム", v: "同じ名前の命令でもオブジェクトの種類により処理内容が異なる性質" },
  ], { fontSize: 10.5, labelW: 2.2, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「カプセル化＝複数クラスの統合」は誤り。「継承＝機能の呼び出し」も不正確（親の性質を引き継ぐ関係）。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：F-16／F-17 情報システム開発（開発プロセスとの関係）。",
    years: mkYears(new Set(["'18", "'20", "'21", "'22", "'23", "'24"])),
  });
}

// ---------- Slide 07: F-5 システム構成 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-5 ／ システム構成（クライアントサーバ、仮想化）",
    title: "P2Pはサーバを介さず対等に通信する方式",
    overview: "仮想化は独立したOS環境を複数動かす技術で、マルチタスクとは異なる。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "3層クライアントサーバ", v: "クライアント・アプリケーションサーバ・DBサーバで役割分担" },
    { k: "P2P", v: "サーバを介さず対等なコンピュータ同士が直接通信" },
    { k: "仮想化（ハイパーバイザー）", v: "1台の物理機に複数の独立した仮想環境を作る" },
    { k: "コンテナ技術", v: "仮想化より軽量にアプリ実行環境を分離（例：Docker）" },
  ], { fontSize: 10.5, labelW: 2.5, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「P2Pはクライアントサーバ方式の一種」は誤り。仮想化は「マルチタスク」とは異なる独立OS環境の技術。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：F-8 外部情報資源の活用（クラウドは仮想化技術を土台にする）。",
    years: mkYears(new Set(["'19", "'20", "'21", "'22", "'23", "'25"])),
  });
}

// ---------- Slide 08: F-6 情報処理の形態 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-6 ／ 情報処理の形態（集中/分散/並列、リアルタイム/バッチ）",
    title: "ATMはリアルタイム、給与計算はバッチ",
    overview: "「即時性」か「まとめて効率化」かで使い分ける。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "集中処理", v: "1台に処理を集約。管理は容易だが故障すると全体停止" },
    { k: "分散処理", v: "複数台に処理を分散。柔軟だが管理は複雑" },
    { k: "リアルタイム処理", v: "発生の都度即座に処理（例：ATM）" },
    { k: "バッチ処理", v: "ためて一括処理（例：夜間の給与計算）" },
  ], { fontSize: 11, labelW: 2.1, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「バッチ処理は結果がすぐ得られ緊急業務向き」は誤り。即時性が要るのはリアルタイム処理。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "C", rankLabel: "出題実績あり",
    related: "関連：F-9 データベース（トランザクション管理もリアルタイム性と関連）。",
    years: mkYears(new Set(["'19"])),
  });
}

// ---------- Slide 09: F-7 ヒューマンインタフェース ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-7 ／ ヒューマンインタフェース（Webデザイン、マルチメディア）",
    title: "JPEGは非可逆圧縮、元の画質には戻せない",
    overview: "ユーザビリティとアクセシビリティは別の観点。両方を意図的に設計する。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "ユーザビリティ", v: "目的の操作をどれだけ効率よく・迷わず行えるか（使いやすさ）" },
    { k: "アクセシビリティ", v: "誰もが情報・機能に到達し利用できるか（利用の公平性）" },
    { k: "可逆圧縮", v: "完全に元通りに復元できる（例：ZIP、PNG）" },
    { k: "非可逆圧縮", v: "情報を間引き高圧縮率を実現、元に戻せない（例：JPEG、MP3）" },
  ], { fontSize: 10.5, labelW: 2.1, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「JPEGは可逆圧縮」は誤り。「ユーザビリティが高ければアクセシビリティも自動的に満たされる」も誤り。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：F-15 情報システムの適用領域（Webシステムのユーザ体験設計）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 10: F-8 外部情報資源の活用 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-8 ／ 外部情報資源の活用（クラウド、OSS）",
    title: "SaaSは開発ではなく既製アプリの利用",
    overview: "SaaS/PaaS/IaaSは「どこまで自分で管理するか」の範囲で区別する。",
    tag: "経営情報システム",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "SaaS", desc: "アプリケーションソフトそのものを提供。利用者はデータの入力・利用のみ" },
    { name: "PaaS", desc: "開発・実行の基盤（OS・DB等）を提供。利用者は自社アプリを用意" },
    { name: "IaaS", desc: "サーバ・ストレージ等インフラのみ提供。利用者がOS以上すべて管理" },
    { name: "OSS", desc: "ソースコード公開。ライセンス条件下で改変・再配布が可能なことが多い" },
  ], { nameW: 1.4, tagW: 0, rowH: 0.9 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：「SaaSは利用者が自らアプリを開発して稼働させる」はPaaSの説明で誤り。OSSは改変・再配布が一律禁止ではない。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'21", "'22", "'24"])),
  });
}

// ---------- Slide 11: F-9 データベース ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-9 ／ データベースの種類と構成、DBMS（正規化）",
    title: "主キーが単一属性なら第2正規形は自動的にクリア",
    overview: "正規化は1NF→2NF→3NFの順に、重複・矛盾のリスクを段階的に排除する。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "第1正規形", v: "各セルに単一の値のみ。繰り返し項目がない" },
    { k: "第2正規形", v: "非キー属性が主キーの一部にのみ依存する状態（部分関数従属）がない" },
    { k: "第3正規形", v: "非キー属性同士が間接的に決まる関係（推移的関数従属）がない" },
    { k: "ACID特性", v: "原子性・一貫性・独立性・永続性。トランザクションが満たすべき4性質" },
  ], { fontSize: 10.5, labelW: 1.9, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "主キーが単一属性の表は第1正規形なら自動的に第2正規形も満たす（部分関数従属は複合キーでのみ発生）。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：F-3 プログラム設計（データ構造の基礎知識）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 12: F-9 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-9 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第11問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "経営情報システム",
  });
  const cy = addExamQuestion(s, {
    stem: "ある学校における開講講座の一覧表（主キー：開講コード。講座コード・講座名・講師コード・講師が列にあり、同一講座を複数の開講コードで開講し、講師は開講コードの異なる同一名の講座を担当することがある）に関する正規化の観点からの記述の正誤の組み合わせとして、最も適切なものを下記の解答群から選べ。ａ：第1正規形である。ｂ：第2正規形である。ｃ：第3正規形である。",
    stemH: 1.1,
    choices: [
      { badge: "ア", text: "ａ：正　ｂ：正　ｃ：正", h: 0.42 },
      { badge: "イ", text: "ａ：正　ｂ：正　ｃ：誤", h: 0.42 },
      { badge: "ウ", text: "ａ：正　ｂ：誤　ｃ：正", h: 0.42 },
      { badge: "エ", text: "ａ：正　ｂ：誤　ｃ：誤", h: 0.42 },
      { badge: "オ", text: "ａ：誤　ｂ：誤　ｃ：誤", h: 0.42 },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/F1JI2025.pdf（令和7年度第1次試験）第11問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 13: F-9 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-9 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第11問）",
    overview: "正解はイ。単一属性の主キーなら第2正規形は自動的にクリアする。",
    tag: "経営情報システム",
  });
  const choices = [
    { badge: "ア", text: "ａ：正　ｂ：正　ｃ：正", h: 0.42 },
    { badge: "イ", text: "ａ：正　ｂ：正　ｃ：誤", h: 0.42 },
    { badge: "ウ", text: "ａ：正　ｂ：誤　ｃ：正", h: 0.42 },
    { badge: "エ", text: "ａ：正　ｂ：誤　ｃ：誤", h: 0.42 },
    { badge: "オ", text: "ａ：誤　ｂ：誤　ｃ：誤", h: 0.42 },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 1 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：イ", options: { bold: true, color: RED } },
    { text: "。ａ（正）：各セルに単一の値のみで繰り返し項目がなく第1正規形を満たす。ｂ（正）：主キーは「開講コード」という単一属性で複合キーではないため部分関数従属が発生しようがなく、第1正規形なら自動的に第2正規形も満たす。ｃ（誤）：「開講コード→講座コード→講座名」「開講コード→講師コード→講師」という", options: {} },
    { text: "推移的関数従属", options: { bold: true } },
    { text: "が存在するため第3正規形の条件を満たさない。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.85, fontFace: F_BODY, fontSize: 9.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  cy += 0.92;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "この表を「開講講座マスタ（講座コード・講座名）」等に分割すれば、講座名の重複入力・更新漏れを防げる（第3正規形への分解）。", options: { fontFace: F_BODY, fontSize: 9.5, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.55, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.6;
  s.addText("関連知識：主キーが複合キー（例：受注番号＋商品コード）の場合にのみ部分関数従属の検討が必要になる点も対比で押さえておくこと。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 9.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.35;
  s.addText("出典：past_exams/1st_stage/1ji2025/F1JI2025.pdf（令和7年度第1次試験）第11問／正解：past_exams/1st_stage_answers/r07/f_v2_20250902.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 14: F-10 通信ネットワーク ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-10 ／ 通信ネットワークの種類と構成、プロトコル",
    title: "信頼性のTCP、速度のUDP",
    overview: "似た役割のプロトコルを入れ替えるひっかけが頻出。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "TCP／UDP", v: "TCP＝信頼性重視（再送制御あり）／UDP＝速度重視（再送制御なし）" },
    { k: "IP", v: "宛先IPアドレスに基づく経路選択とアドレス指定" },
    { k: "DHCP", v: "機器にIPアドレス等の設定情報を自動的に割り当てる" },
    { k: "ARP／DNS", v: "ARP＝IPからMACアドレスを解決／DNS＝ドメイン名をIPアドレスに変換" },
  ], { fontSize: 10.5, labelW: 2.2, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "経路選択のIPとアドレス解決のARP、設定自動割当のDHCPと監視のSNMPを混同しない。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：F-20 情報セキュリティ管理・対策（暗号化通信のHTTPS等）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 15: F-10 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-10 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第5問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "経営情報システム",
  });
  const cy = addExamQuestion(s, {
    stem: "情報ネットワークで用いられる通信プロトコルに関する記述とその用語の組み合わせとして、最も適切なものを下記の解答群から選べ。ａ：信頼性よりも速度を重視し、確認や再送制御を行わないプロトコル。ｂ：機器にIPアドレス等の設定情報を自動的に割り当てるプロトコル。ｃ：正しい宛先に経路選択・アドレス指定を行うプロトコル。",
    stemH: 0.85,
    choices: [
      { badge: "ア", text: "ａ：TCP　ｂ：SNMP　ｃ：ARP", h: 0.42 },
      { badge: "イ", text: "ａ：TCP　ｂ：SNMP　ｃ：IP", h: 0.42 },
      { badge: "ウ", text: "ａ：UDP　ｂ：DHCP　ｃ：ARP", h: 0.42 },
      { badge: "エ", text: "ａ：UDP　ｂ：DHCP　ｃ：IP", h: 0.42 },
      { badge: "オ", text: "ａ：UDP　ｂ：SNMP　ｃ：ARP", h: 0.42 },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/F1JI2025.pdf（令和7年度第1次試験）第5問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 16: F-10 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-10 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第5問）",
    overview: "正解はエ。ａ＝UDP、ｂ＝DHCP、ｃ＝IPが正しい対応。",
    tag: "経営情報システム",
  });
  const choices = [
    { badge: "ア", text: "ａ：TCP　ｂ：SNMP　ｃ：ARP", h: 0.42 },
    { badge: "イ", text: "ａ：TCP　ｂ：SNMP　ｃ：IP", h: 0.42 },
    { badge: "ウ", text: "ａ：UDP　ｂ：DHCP　ｃ：ARP", h: 0.42 },
    { badge: "エ", text: "ａ：UDP　ｂ：DHCP　ｃ：IP", h: 0.42 },
    { badge: "オ", text: "ａ：UDP　ｂ：SNMP　ｃ：ARP", h: 0.42 },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 3 });
  cy += 0.08;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.1;
  s.addText([
    { text: "正解：エ", options: { bold: true, color: RED } },
    { text: "。ａ＝", options: {} },
    { text: "UDP", options: { bold: true } },
    { text: "（接続確立・再送制御を行わず速度優先。TCPは逆に信頼性重視）。ｂ＝", options: {} },
    { text: "DHCP", options: { bold: true } },
    { text: "（IPアドレス等の設定を自動割当。SNMPは機器の稼働監視用でこの機能はない）。ｃ＝", options: {} },
    { text: "IP", options: { bold: true } },
    { text: "（経路選択とアドレス指定。ARPはIPからMACアドレスを解決するのみで経路選択は行わない）。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.85, fontFace: F_BODY, fontSize: 9.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  cy += 0.92;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "TCP/IPの階層で、IP＝ネットワーク層、TCP/UDP＝トランスポート層、DHCP・SNMP＝アプリケーション層に位置づけられる。", options: { fontFace: F_BODY, fontSize: 9.5, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.55, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  cy += 0.6;
  s.addText("関連知識：「信頼性のTCP・速度のUDP」「自動割当のDHCP」「経路選択のIP」「アドレス解決のARP」「監視のSNMP」という役割の対応をセットで暗記しておくこと。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 9.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.35;
  s.addText("出典：past_exams/1st_stage/1ji2025/F1JI2025.pdf（令和7年度第1次試験）第5問／正解：past_exams/1st_stage_answers/r07/f_v2_20250902.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 17: F-11 システム評価 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-11 ／ システム評価（性能、信頼性、経済性）",
    title: "「事前に防ぐ」か「起きたらどうするか」で整理する",
    overview: "似た名前の信頼性設計用語を入れ替える出題が最頻出。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "フォールトアボイダンス", v: "部品の品質を高め故障自体を未然に防ぐ" },
    { k: "フォールトトレランス", v: "多重化・冗長化で一部故障しても全体は動き続ける" },
    { k: "フェイルセーフ／フェイルソフト", v: "セーフ＝安全側に導く／ソフト＝機能を縮退させ継続" },
    { k: "稼働率", v: "MTBF ÷（MTBF＋MTTR）で算出" },
  ], { fontSize: 10.5, labelW: 2.7, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "フェイルオーバ（待機系へ切替）とフェイルソフト（機能縮退）、フォールトアボイダンスとトレランスを入れ替える出題に注意。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：F-1 ハードウェア（RAIDも信頼性設計の一種）。",
    years: mkYears(new Set(["'17", "'18", "'19", "'21", "'24", "'25"])),
  });
}

// ---------- Slide 18: F-11 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-11 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第10問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "経営情報システム",
  });
  const cy = addExamQuestion(s, {
    stem: "情報システムの信頼性設計に関する記述として、最も適切なものはどれか。",
    stemH: 0.5,
    choices: [
      { badge: "ア", text: "フェイルオーバとは、故障や障害が発生したときに、一部の機能を低下させても、限定的ながら重要な機能だけでも稼働し続けるように設計することである。", h: 0.6 },
      { badge: "イ", text: "フォールトアボイダンスとは、部品一つ一つの信頼性を高めることで、故障や障害が発生しないように設計することである。", h: 0.55 },
      { badge: "ウ", text: "フォールトトレランスとは、人為的な操作ミスがあっても危険が生じず、システムに異常が起こらないように設計することである。", h: 0.55 },
      { badge: "エ", text: "フォールトマスキングとは、故障や障害が発生したときに、システムの被害を最小限にとどめるように設計することである。", h: 0.55 },
      { badge: "オ", text: "フォールバックとは、故障や障害が発生したときに、待機系システムが処理を継続するように設計することである。", h: 0.55 },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/F1JI2025.pdf（令和7年度第1次試験）第10問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 19: F-11 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-11 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第10問）",
    overview: "正解はイ。フォールトアボイダンスは故障の未然防止を指す。",
    tag: "経営情報システム",
  });
  const choices = [
    { badge: "ア", text: "フェイルオーバとは、故障や障害が発生したときに、一部の機能を低下させても、限定的ながら重要な機能だけでも稼働し続けるように設計することである。", h: 0.6 },
    { badge: "イ", text: "フォールトアボイダンスとは、部品一つ一つの信頼性を高めることで、故障や障害が発生しないように設計することである。", h: 0.55 },
    { badge: "ウ", text: "フォールトトレランスとは、人為的な操作ミスがあっても危険が生じず、システムに異常が起こらないように設計することである。", h: 0.55 },
    { badge: "エ", text: "フォールトマスキングとは、故障や障害が発生したときに、システムの被害を最小限にとどめるように設計することである。", h: 0.55 },
    { badge: "オ", text: "フォールバックとは、故障や障害が発生したときに、待機系システムが処理を継続するように設計することである。", h: 0.55 },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 1 });
  cy += 0.06;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "正解：イ", options: { bold: true, color: RED } },
    { text: "。フォールトアボイダンスは部品一つひとつの品質・信頼性を高め故障自体を未然に防ぐ設計思想で記述どおり正しい。ア：これは", options: {} },
    { text: "フェイルソフト", options: { bold: true } },
    { text: "の説明（フェイルオーバは待機系への自動切替）。ウ：これは", options: {} },
    { text: "フールプルーフ", options: { bold: true } },
    { text: "の説明（フォールトトレランスは多重化で故障を許容）。エ：これは", options: {} },
    { text: "フェイルセーフ", options: { bold: true } },
    { text: "の説明（フォールトマスキングは多重化結果の照合で影響を隠蔽）。オ：これは", options: {} },
    { text: "フェイルオーバ", options: { bold: true } },
    { text: "の説明（フォールバックは機能を縮退させ継続運用）。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.7, fontFace: F_BODY, fontSize: 9, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.1 });
  cy += 0.76;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "MTBFが200時間、MTTRが10時間なら稼働率は200÷(200+10)≒95.2％。MTBFを増やす（フォールトアボイダンスの発想）かMTTRを減らせば稼働率は上がる。", options: { fontFace: F_BODY, fontSize: 9, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.45, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  cy += 0.5;
  s.addText("関連知識：「事前に防ぐ（アボイダンス）」「起きても耐える（トレランス）」「起きたら安全側（セーフ）」「起きたら縮退（ソフト）」「起きたら切替（オーバ）」「誤操作防止（フールプルーフ）」で整理すること。", {
    x: 0.55, y: cy, w: 12.25, h: 0.35, fontFace: F_BODY, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15,
  });
  cy += 0.4;
  s.addText("出典：past_exams/1st_stage/1ji2025/F1JI2025.pdf（令和7年度第1次試験）第10問／正解：past_exams/1st_stage_answers/r07/f_v2_20250902.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

pres.writeFile({ fileName: path.join(__dirname, "../../slides/1st_stage/F_information_systems.pptx") })
  .then(() => console.log("wrote", "slides/1st_stage/F_information_systems.pptx"));
