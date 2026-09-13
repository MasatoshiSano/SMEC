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

// ---------- Slide 20: F-12〜F-24 区切り ----------
addDividerSlide(pres, {
  ghostNo: "02",
  partNo: "PART 02",
  partLabel: "経営情報システム ／ F-12〜F-24",
  title: "経営情報管理",
  desc: "ITそのものの技術知識から一歩進み、「経営にITをどう活かすか」「開発・セキュリティ・データをどうマネジメントするか」という視点を扱う。",
  chips: ["F-17 アジャイル開発", "F-18 WBS", "F-20 ランサムウェア対策"],
  notes: "経営情報管理パートの区切りスライド。",
});

// ---------- Slide 21: F-12 情報システム戦略 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-12 ／ 情報システム戦略（経営戦略との連携、システム化計画）",
    title: "全体最適化計画を経ずに着手すると「サイロ化」する",
    overview: "情報システム戦略は経営戦略との整合性が本質。現場レベルの開発手法とは別の話。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "全体最適化計画", v: "会社全体の設計図。どの業務にどのシステムを配置するか" },
    { k: "個別システム化計画", v: "全体最適化計画をもとに個別システムの開発を具体化" },
    { k: "CIO", v: "最高情報責任者。経営戦略とIT戦略を橋渡しする役職" },
  ], { fontSize: 11.5, labelW: 2.4, gap: 0.42 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "情報システム戦略は「経営戦略との整合性」の話で、F-16／F-17の「個別開発の進め方」とは異なる次元の論点。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：F-13 経営革新と情報システム（DXとの連続性）。",
    years: mkYears(new Set(["'17", "'18", "'24"])),
  });
}

// ---------- Slide 22: F-13 経営革新と情報システム ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-13 ／ 経営革新と情報システム（DX、業務プロセス改善）",
    title: "DXは単なるデジタル化ではない",
    overview: "デジタイゼーション→デジタライゼーション→DX（ビジネスモデル変革）の順に広がる。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "DX", v: "デジタル技術で業務プロセス・組織・ビジネスモデルそのものを変革" },
    { k: "BPR", v: "既存業務を前提とせずゼロベースで業務プロセスを再設計" },
    { k: "RPA", v: "定型的な事務作業をソフトウェアロボットに自動化させる" },
    { k: "2025年の崖", v: "レガシーシステムを放置すると生じうる多額の経済損失という経産省の問題提起" },
  ], { fontSize: 10.5, labelW: 2.0, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「DX＝紙をデータ化すること」は不十分。デジタル化を土台にしつつビジネスモデルまで変革する広い概念。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：F-24 データサイエンス（AI活用もDXの重要な手段）。",
    years: mkYears(new Set(["'17", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 23: F-14 ITガバナンス ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-14 ／ ITガバナンス（内部統制、システム監査）",
    title: "システム監査は開発部門が自ら行うものではない",
    overview: "内部統制＝自己点検の仕組み、システム監査＝独立した第三者の検証。",
    tag: "経営情報システム",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "内部統制", desc: "経営者自身の責任で整備・運用する社内のルール・仕組み（COSOフレームワーク）" },
    { name: "システム監査", desc: "開発・運用から独立した第三者（監査人）が適切性を客観的に検証する活動" },
  ], { nameW: 2.3, tagW: 0, rowH: 1.3 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：「システム監査はシステム部門が自ら行う」は独立性の観点から誤り。",
    years: mkYears(new Set(["'16", "'20", "'25"])),
  });
}

// ---------- Slide 24: F-15 情報システムの適用領域 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-15 ／ 情報システムの適用領域（SCM、CRM）",
    title: "SCMは供給側、CRMは顧客側の最適化",
    overview: "ERPは特定1業務ではなく複数の基幹業務を横断的に統合する点が特徴。",
    tag: "経営情報システム",
  });
  addRowList(s, 0.55, 1.95, 12.25, [
    { name: "ERP", desc: "会計・人事・生産・販売等の基幹業務データを一元管理・連携させる" },
    { name: "SCM", desc: "調達〜生産〜物流〜販売を企業の垣根を越えて最適化（供給側）" },
    { name: "CRM", desc: "顧客の属性・購買履歴を一元管理し関係を強化・維持（顧客側）" },
    { name: "SFA／POS", desc: "SFA＝営業支援／POS＝販売時点での実績記録" },
  ], { nameW: 1.6, tagW: 0, rowH: 0.9 });
  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "ひっかけ：「SCMとCRMは同じ概念」は誤り。対象範囲・向いている方向が異なる。",
    years: mkYears(new Set(["'16", "'19", "'24", "'25"])),
  });
}

// ---------- Slide 25: F-16 情報システム開発（ウォーターフォール型） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-16 ／ 情報システム開発（ウォーターフォール型）",
    title: "要件変更に弱く、手戻りコストが大きい",
    overview: "V字モデルで各設計工程と対応するテスト工程が結びつく。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "工程の流れ", v: "要件定義→外部設計→内部設計→プログラミング→テスト→運用保守" },
    { k: "テストの順序", v: "単体テスト→結合テスト→システムテスト→運用テスト" },
    { k: "V字モデル", v: "要件定義⇔運用テスト、外部設計⇔システムテスト等の対応関係" },
  ], { fontSize: 11, labelW: 2.0, gap: 0.42 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「要件変更に柔軟」は誤り。手戻りコストが大きく変更には弱い（柔軟なのはF-17のアジャイル型）。テスト順序の入れ替えにも注意。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：F-17 情報システム開発（アジャイル型との対比）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'25"])),
  });
}

// ---------- Slide 26: F-17 情報システム開発（アジャイル型） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-17 ／ 情報システム開発（アジャイル型、インクリメンタル型）",
    title: "デイリースクラムは成果物のお披露目の場ではない",
    overview: "似た用語の定義を入れ替えるひっかけが非常に多い分野。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "デイリースクラム", v: "毎日の進捗共有（15分程度）。成果物提示はスプリントレビューの役割" },
    { k: "スプリントレビュー", v: "完成機能をステークホルダーに見せフィードバックを得る場" },
    { k: "ペアプログラミング", v: "XPのプラクティス。2人が1台の端末でレビューしながら開発" },
    { k: "DevOps", v: "開発と運用が密接に連携し、CI/CDでリリースを高速化する文化" },
  ], { fontSize: 10.5, labelW: 2.3, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「ローコード開発＝反復的な機能単位開発」は誤り（それはインクリメンタル型）。「DevOps＝開発運用の分離」も誤り。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：F-16 情報システム開発（ウォーターフォール型との対比）。",
    years: mkYears(new Set(["'19", "'21", "'22", "'24", "'25"])),
  });
}

// ---------- Slide 27: F-17 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-17 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第13問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "経営情報システム",
  });
  const cy = addExamQuestion(s, {
    stem: "システム開発手法に関する記述の正誤の組み合わせとして、最も適切なものを下記の解答群から選べ。ａ：デイリースクラムでは、スプリントの成果をステークホルダーに提示し、フィードバックを得る。ｂ：ローコード開発では、システムの全体像をモデル化し、優先度を付けた機能単位で計画、設計、構築を反復的に行う。ｃ：DevOpsでは、開発と運用のフェーズを明確に分離して、システムの導入や更新を柔軟かつ迅速に行う。ｄ：XPにおけるペアプログラミングでは、2人のプログラマがペアとなり、相談やレビューを行いながら、協力してプログラムの開発を行う。",
    stemH: 1.3,
    choices: [
      { badge: "ア", text: "ａ：正　ｂ：正　ｃ：誤　ｄ：誤", h: 0.42 },
      { badge: "イ", text: "ａ：正　ｂ：誤　ｃ：正　ｄ：正", h: 0.42 },
      { badge: "ウ", text: "ａ：誤　ｂ：正　ｃ：誤　ｄ：正", h: 0.42 },
      { badge: "エ", text: "ａ：誤　ｂ：誤　ｃ：正　ｄ：誤", h: 0.42 },
      { badge: "オ", text: "ａ：誤　ｂ：誤　ｃ：誤　ｄ：正", h: 0.42 },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/F1JI2025.pdf（令和7年度第1次試験）第13問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 28: F-17 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-17 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第13問）",
    overview: "正解はオ。a〜cはいずれも別の用語の説明で誤り。",
    tag: "経営情報システム",
  });
  const choices = [
    { badge: "ア", text: "ａ：正　ｂ：正　ｃ：誤　ｄ：誤", h: 0.42 },
    { badge: "イ", text: "ａ：正　ｂ：誤　ｃ：正　ｄ：正", h: 0.42 },
    { badge: "ウ", text: "ａ：誤　ｂ：正　ｃ：誤　ｄ：正", h: 0.42 },
    { badge: "エ", text: "ａ：誤　ｂ：誤　ｃ：正　ｄ：誤", h: 0.42 },
    { badge: "オ", text: "ａ：誤　ｂ：誤　ｃ：誤　ｄ：正", h: 0.42 },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 4 });
  cy += 0.06;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "正解：オ", options: { bold: true, color: RED } },
    { text: "。ａ（誤）：ステークホルダーへの成果提示は", options: {} },
    { text: "スプリントレビュー", options: { bold: true } },
    { text: "の説明（デイリースクラムは毎日の進捗共有）。ｂ（誤）：反復的な計画・設計・構築の進め方は", options: {} },
    { text: "インクリメンタル型", options: { bold: true } },
    { text: "の説明（ローコード開発はコード記述量を減らす手段）。ｃ（誤）：DevOpsは開発と運用を", options: {} },
    { text: "密接に連携", options: { bold: true } },
    { text: "させる文化で「分離」は正反対。ｄ（正）：ペアプログラミングの説明どおり正しい。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.72, fontFace: F_BODY, fontSize: 9, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.12 });
  cy += 0.78;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "スクラムの主要イベントは、デイリースクラム（進捗共有）、スプリントレビュー（成果物確認）、スプリントレトロスペクティブ（進め方の振り返り）と目的が異なる。", options: { fontFace: F_BODY, fontSize: 9, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  cy += 0.55;
  s.addText("関連知識：正しい定義文に別の用語名を当てはめる出題パターンが頻出。用語と定義を1対1で正確に結びつけて覚えること。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.35;
  s.addText("出典：past_exams/1st_stage/1ji2025/F1JI2025.pdf（令和7年度第1次試験）第13問／正解：past_exams/1st_stage_answers/r07/f_v2_20250902.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 29: F-18 組織と人材 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-18 ／ 組織と人材（プロジェクトマネジメント、サービスマネジメント）",
    title: "WBS辞書はスコープ記述書とは別物",
    overview: "インシデント管理（応急対応）と問題管理（根本対策）を区別する。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "ワークパッケージ", v: "WBS最下位の作業単位。進捗・コスト管理の最小単位" },
    { k: "100％ルール", v: "上位作業を抜け漏れ・重複なく下位作業に分解する原則" },
    { k: "クリティカルパス", v: "PERT図で最も時間のかかる作業経路（遅れると全体が遅れる）" },
    { k: "インシデント管理／問題管理", v: "インシデント＝迅速な応急対応／問題＝根本原因の恒久対策" },
  ], { fontSize: 10, labelW: 2.5, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「WBS辞書＝プロジェクト全体の範囲を記述」は誤り（それはスコープ記述書）。「イテレーション＝早期作業を詳細計画」も誤り（それはローリングウェーブ計画法）。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：F-17 情報システム開発（イテレーションはアジャイル開発の反復単位）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'20", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 30: F-18 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-18 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第18問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "経営情報システム",
  });
  const cy = addExamQuestion(s, {
    stem: "WBS（Work Breakdown Structure）に関する記述の正誤の組み合わせとして、最も適切なものを下記の解答群から選べ。ａ：WBS辞書とは、プロジェクト全体の範囲、成果物、前提条件や制約条件などを記述した文書のことである。ｂ：ワークパッケージとは、WBSの最下位レベルの作業群のことで、進捗状況などをコントロールする際の最小単位である。ｃ：100％ルールとは、WBSの階層構造において、上位の作業を過不足なく下位の複数の作業に展開するルールのことである。ｄ：イテレーションとは、WBSを作成する際、早期に完了しなければならない作業は詳細に計画し、将来の作業は概略にとどめておいて、時期がきたら詳細化を繰り返す反復計画技法のことである。",
    stemH: 1.4,
    choices: [
      { badge: "ア", text: "ａ：正　ｂ：正　ｃ：正　ｄ：誤", h: 0.42 },
      { badge: "イ", text: "ａ：正　ｂ：正　ｃ：誤　ｄ：正", h: 0.42 },
      { badge: "ウ", text: "ａ：正　ｂ：誤　ｃ：誤　ｄ：誤", h: 0.42 },
      { badge: "エ", text: "ａ：誤　ｂ：正　ｃ：正　ｄ：誤", h: 0.42 },
      { badge: "オ", text: "ａ：誤　ｂ：誤　ｃ：正　ｄ：正", h: 0.42 },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/F1JI2025.pdf（令和7年度第1次試験）第18問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 31: F-18 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-18 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第18問）",
    overview: "正解はエ。ａ・ｄは正しい定義に別の用語名を当てている。",
    tag: "経営情報システム",
  });
  const choices = [
    { badge: "ア", text: "ａ：正　ｂ：正　ｃ：正　ｄ：誤", h: 0.42 },
    { badge: "イ", text: "ａ：正　ｂ：正　ｃ：誤　ｄ：正", h: 0.42 },
    { badge: "ウ", text: "ａ：正　ｂ：誤　ｃ：誤　ｄ：誤", h: 0.42 },
    { badge: "エ", text: "ａ：誤　ｂ：正　ｃ：正　ｄ：誤", h: 0.42 },
    { badge: "オ", text: "ａ：誤　ｂ：誤　ｃ：正　ｄ：正", h: 0.42 },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 3 });
  cy += 0.06;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "正解：エ", options: { bold: true, color: RED } },
    { text: "。ａ（誤）：プロジェクト全体の範囲・成果物・制約条件を記述するのは", options: {} },
    { text: "スコープ記述書", options: { bold: true } },
    { text: "の説明。WBS辞書は各ワークパッケージごとの作業内容・担当・完了基準を記述する補助文書。ｂ（正）：ワークパッケージの説明どおり正しい。ｃ（正）：100％ルールの説明どおり正しい。ｄ（誤）：早期作業を詳細計画し将来は概略にとどめる技法は", options: {} },
    { text: "ローリングウェーブ計画法", options: { bold: true } },
    { text: "の説明（イテレーションはアジャイル開発の反復単位）。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.72, fontFace: F_BODY, fontSize: 9, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.12 });
  cy += 0.78;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "「要件定義」という上位作業を「業務ヒアリング」「要件一覧の作成」「レビュー」に抜け漏れ・重複なく分解するのが100％ルールの適用例。", options: { fontFace: F_BODY, fontSize: 9, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.4, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  cy += 0.45;
  s.addText("関連知識：WBS辞書とスコープ記述書、イテレーションとローリングウェーブ計画法のように「定義は正しいが用語名が違う」ひっかけが頻出。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.35;
  s.addText("出典：past_exams/1st_stage/1ji2025/F1JI2025.pdf（令和7年度第1次試験）第18問／正解：past_exams/1st_stage_answers/r07/f_v2_20250902.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 32: F-19 情報セキュリティの概念 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-19 ／ 情報セキュリティの概念（機密性・完全性・可用性）",
    title: "「漏れない」「正しい」「使える」の3本柱",
    overview: "1つの対策が複数の要素にまたがることもある点に注意。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "機密性", v: "許可された人だけが情報にアクセスできる（例：権限設定）" },
    { k: "完全性", v: "情報が正確で不正に改ざん・破壊されていない（例：改ざん検知）" },
    { k: "可用性", v: "必要なときに情報・システムを使える（例：二重化、バックアップ）" },
    { k: "リスク対応", v: "回避・低減・移転・受容の4方針" },
  ], { fontSize: 10.5, labelW: 1.8, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「漏れない＝機密性」「正しい＝完全性」「使える＝可用性」で判断。バックアップは主に可用性だが完全性にも関わりうる。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "－", rankLabel: "直近10年単独出題なし",
    related: "関連：F-20 情報セキュリティ管理・対策（CIAを守る具体的な技術・対策）。",
    years: mkYears(new Set([])),
  });
}

// ---------- Slide 33: F-20 情報セキュリティ管理・対策 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-20 ／ 情報セキュリティ管理・対策",
    title: "EPPは事前防止、EDRは事後の検知・対応",
    overview: "ランサムウェア対策では初動でネットワークからの切り離しを優先する。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "EPP／EDR", v: "EPP＝侵入の事前防止／EDR＝侵入後の検知・対応" },
    { k: "共通鍵／公開鍵暗号", v: "共通鍵＝高速だが鍵配送に工夫要／公開鍵＝安全だが低速" },
    { k: "多要素認証", v: "知識・所持・生体の複数要素を組み合わせる" },
    { k: "ランサムウェア対策", v: "バックアップは平常時ネットワークから切り離す。感染時はまず切り離す" },
  ], { fontSize: 10, labelW: 2.3, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "EPPとEDRの役割の入れ替え、「バックアップ機器は常時接続すべき」の誤りに注意。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：F-10 通信ネットワーク（HTTPS等の暗号化通信）。",
    years: mkYears(new Set(["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

// ---------- Slide 34: F-20 過去問チェック（設問） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-20 ／ 過去問で確認する",
    title: "こう出題される（令和7年度 第19問）",
    overview: "前のスライドの内容で答えられるか、解答を見る前に考えてみる。",
    tag: "経営情報システム",
  });
  const cy = addExamQuestion(s, {
    stem: "近年、中小企業においてもランサムウェアによる被害が増加している。ランサムウェアに関する記述として、最も適切なものはどれか。",
    stemH: 0.55,
    choices: [
      { badge: "ア", text: "EDRは、マルウェア感染防止や外部からの攻撃通信ブロックなど、ランサムウェア攻撃による侵入の事前防止を担う。", h: 0.55 },
      { badge: "イ", text: "EPPは、PCやサーバに侵入してしまったランサムウェアを検知し、異常や不審な挙動があればシステム担当者に通知するなど侵入後の事後対処を担う。", h: 0.6 },
      { badge: "ウ", text: "ランサムウェアに感染した際に早期復旧できるように、バックアップデータを保存した機器は、常にネットワークに接続しておく。", h: 0.55 },
      { badge: "エ", text: "ランサムウェアに感染した場合は、速やかに感染した端末の電源を切り、システム担当者やセキュリティベンダに報告する。", h: 0.55 },
      { badge: "オ", text: "ランサムウェアの主要な侵入経路は、VPN機器、リモートデスクトップ、不審メールやその添付ファイルである。", h: 0.5 },
    ],
  });
  s.addText("出典：past_exams/1st_stage/1ji2025/F1JI2025.pdf（令和7年度第1次試験）第19問", {
    x: 0.55, y: cy + 0.15, w: 12.25, h: 0.3,
    fontFace: F_MONO, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 35: F-20 過去問チェック（解答＆解説） ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-20 ／ 過去問で確認する",
    title: "解答＆解説（令和7年度 第19問）",
    overview: "正解はオ。IPAの調査に基づく主要侵入経路の記述。",
    tag: "経営情報システム",
  });
  const choices = [
    { badge: "ア", text: "EDRは、マルウェア感染防止や外部からの攻撃通信ブロックなど、ランサムウェア攻撃による侵入の事前防止を担う。", h: 0.55 },
    { badge: "イ", text: "EPPは、PCやサーバに侵入してしまったランサムウェアを検知し、異常や不審な挙動があればシステム担当者に通知するなど侵入後の事後対処を担う。", h: 0.6 },
    { badge: "ウ", text: "ランサムウェアに感染した際に早期復旧できるように、バックアップデータを保存した機器は、常にネットワークに接続しておく。", h: 0.55 },
    { badge: "エ", text: "ランサムウェアに感染した場合は、速やかに感染した端末の電源を切り、システム担当者やセキュリティベンダに報告する。", h: 0.55 },
    { badge: "オ", text: "ランサムウェアの主要な侵入経路は、VPN機器、リモートデスクトップ、不審メールやその添付ファイルである。", h: 0.5 },
  ];
  let cy = addExamQuestion(s, { choices, correctIndex: 4 });
  cy += 0.06;
  s.addShape("line", { x: 0.55, y: cy, w: 12.25, h: 0, line: { color: INK, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "正解：オ", options: { bold: true, color: RED } },
    { text: "。IPAの「情報セキュリティ10大脅威」でもVPN機器の脆弱性・RDPの不正利用・不審メールが継続して挙げられており記述と一致。ア・イ：EPP（事前防止）とEDR（事後の検知・対応）の役割が", options: {} },
    { text: "入れ替わっている", options: { bold: true } },
    { text: "。ウ：バックアップ機器を常時接続すると", options: {} },
    { text: "バックアップ自体も暗号化される危険", options: { bold: true } },
    { text: "があり、平常時は切り離すべき。エ：まず優先すべきは", options: {} },
    { text: "ネットワークからの切り離し", options: { bold: true } },
    { text: "（電源を切るとフォレンジック調査に必要な情報が失われうる）。", options: {} },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.85, fontFace: F_BODY, fontSize: 9, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.12 });
  cy += 0.9;
  s.addText([
    { text: "具体例\n", options: { fontFace: F_MONO, fontSize: 8.5, bold: true, color: INK_SOFT, breakLine: true } },
    { text: "感染時はLANケーブルを抜く等でネットワークから切り離し、その後システム担当者・セキュリティベンダへ報告する初動対応が推奨される。", options: { fontFace: F_BODY, fontSize: 9, color: INK } },
  ], { x: 0.55, y: cy, w: 12.25, h: 0.4, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  cy += 0.45;
  s.addText("関連知識：EPP＝侵入させない予防、EDR＝侵入後の検知・対応という補完関係。バックアップのオフライン・オフサイト保管が基本。", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_BODY, fontSize: 9, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  cy += 0.35;
  s.addText("出典：past_exams/1st_stage/1ji2025/F1JI2025.pdf（令和7年度第1次試験）第19問／正解：past_exams/1st_stage_answers/r07/f_v2_20250902.pdf", {
    x: 0.55, y: cy, w: 12.25, h: 0.3, fontFace: F_MONO, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
}

// ---------- Slide 36: F-21 情報システムの評価 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-21 ／ 情報システムの評価（品質評価、価値評価）",
    title: "品質評価と経済性評価は別の評価軸",
    overview: "品質特性6観点は名前が似ているため混同しやすい。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "品質特性", v: "機能性・信頼性・使用性・効率性・保守性・移植性" },
    { k: "効率性", v: "限られた資源（処理時間・メモリ）で効率よく動作するか" },
    { k: "ROI", v: "投資額に対する利益の割合（経済性評価の指標）" },
  ], { fontSize: 11.5, labelW: 2.0, gap: 0.45 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「効率性」（資源利用の良さ）と「効果性」（目的達成度）を取り違えない。品質面と経済面は異なる評価軸。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "B", rankLabel: "頻出論点",
    related: "関連：F-11 システム評価（信頼性・経済性の指標と重複）。",
    years: mkYears(new Set(["'19", "'23"])),
  });
}

// ---------- Slide 37: F-22 意思決定支援：問題分析・意思決定技法 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-22 ／ 意思決定支援：問題分析・意思決定技法",
    title: "デシジョンツリーは□が意思決定、○が確率分岐",
    overview: "線形計画法は利益最大化だけでなくコスト最小化にも使える。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "線形計画法", v: "制約条件下で利益最大化・コスト最小化する生産量の組合せを求める" },
    { k: "デシジョンツリー分析", v: "選択肢と不確実な事象を樹木状に図示し期待値で判断" },
    { k: "待ち行列理論", v: "窓口数やサービス時間から行列の長さ・待ち時間を予測" },
  ], { fontSize: 11, labelW: 2.4, gap: 0.45 });
  cy += 0.1;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "意思決定者が選べる分岐点は□（四角）、確率に左右される分岐点は○（丸）。線形計画法を「利益最大化限定」と覚えない。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "C", rankLabel: "出題実績あり",
    related: "関連：F-18 組織と人材（PERT図・クリティカルパスは本論点とも共通）。",
    years: mkYears(new Set(["'18"])),
  });
}

// ---------- Slide 38: F-23 意思決定支援：データ分析技術と活用 ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-23 ／ 意思決定支援：データ分析技術と活用",
    title: "相関関係は因果関係を証明しない",
    overview: "主成分分析（要約）と因子分析（背後の共通因子推定）は目的が異なる。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "回帰分析", v: "説明変数から目的変数を予測する式を求める。2変数以上は重回帰分析" },
    { k: "主成分分析", v: "多数の変数を情報量を保ったまま少数の合成変数に要約" },
    { k: "因子分析", v: "観測変数の背後にある共通の要因（因子）を推定" },
    { k: "クラスター分析", v: "性質の似ているもの同士をグループに分類" },
  ], { fontSize: 10.5, labelW: 2.0, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「相関があれば必ず因果関係がある」は誤り（見せかけの相関の可能性）。主成分分析と因子分析の目的の違いに注意。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：F-24 データサイエンス（教師なし学習としてのクラスター分析）。",
    years: mkYears(new Set(["'16", "'17", "'19", "'20", "'21", "'22", "'23", "'25"])),
  });
}

// ---------- Slide 39: F-24 意思決定支援：データサイエンス ----------
{
  const s = pres.addSlide();
  addHeader(s, {
    kicker: "F-24 ／ 意思決定支援：データサイエンス（AI、確率・統計）",
    title: "教師あり学習は「正解データ」の有無で区別する",
    overview: "ハルシネーションは情報漏洩ではなく、事実に基づかない生成のこと。",
    tag: "経営情報システム",
  });
  const proseX = 0.55, proseW = 12.25;
  let cy = 1.85;
  cy = addTermRows(s, proseX, cy, proseW, [
    { k: "教師あり学習", v: "正解データを使って学習し、新しいデータを予測（分類・回帰）" },
    { k: "教師なし学習", v: "正解データなしにデータ自体のパターン・構造を発見" },
    { k: "LLM", v: "大量の文章データを学習し文章生成・要約・翻訳を行うAIモデル" },
    { k: "ハルシネーション", v: "AIが事実に基づかない誤情報をもっともらしく生成する現象" },
  ], { fontSize: 10.5, labelW: 1.8, gap: 0.36 });
  cy += 0.08;
  s.addShape("line", { x: proseX, y: cy, w: proseW, h: 0, line: { color: LINE, width: 1 } });
  cy += 0.08;
  s.addText([
    { text: "ひっかけ：", options: { bold: true, color: RED } },
    { text: "「教師あり学習＝正解データなし」は誤り（それは教師なし学習）。ハルシネーションは情報漏洩とは別の問題。", options: { color: RED } },
  ], { x: proseX, y: cy, w: proseW, h: 0.5, fontFace: F_BODY, fontSize: 9.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  addFreqBar(s, {
    y: 6.55, rank: "A", rankLabel: "最頻出論点",
    related: "関連：F-13 経営革新と情報システム（AI活用はDXの重要な手段）。",
    years: mkYears(new Set(["'16", "'17", "'19", "'20", "'21", "'22", "'23", "'24", "'25"])),
  });
}

pres.writeFile({ fileName: path.join(__dirname, "../../slides/1st_stage/F_information_systems.pptx") })
  .then(() => console.log("wrote", "slides/1st_stage/F_information_systems.pptx"));
