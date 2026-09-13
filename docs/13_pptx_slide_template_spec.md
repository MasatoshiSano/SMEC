# PowerPointスライドテンプレート仕様書（pptxgenjs / Swiss Grid）

`docs/11_slide_template_spec.md`（HTML版）と同じSwiss Gridデザインを、PowerPoint（.pptx）でダウンロードしたいというユーザー要望を受けて`pptxgenjs`で実装した再利用可能な部品ライブラリ。実体は`scripts/pptx_slidegen/`。

## 位置づけ・経緯

- `scripts/slidegen/`（python-pptx製）は先行して試作したが、生成物がPowerPointで「修復が必要」エラーになる不具合が未解決のまま放棄された（CLAUDE.md参照）。今回の`scripts/pptx_slidegen/`（pptxgenjs製）はそれとは別の新規実装であり、`scripts/slidegen/`を置き換えるもの。
- HTML版（`slides/1st_stage/*.html`）が先にあり、PowerPoint版はその内容をそのまま別フォーマットに移植したもの。**内容の正はHTML版**：今後HTML版の内容を変更したら、対応するPowerPoint版もこの仕組みで再生成すること。

## ディレクトリ構成

```
scripts/pptx_slidegen/
  package.json         # pptxgenjsの依存を宣言（node_modulesはgit管理外）
  lib.js               # 再利用可能な部品ライブラリ（このファイルの本体）
  build_C_business_administration.js   # 企業経営理論デッキの内容定義（本ファイルが唯一のコンテンツ）
```

新しい科目・デッキを作る場合は`build_<科目名>.js`を追加し、`lib.js`をrequireして内容を書く。`lib.js`自体は編集せず、機能追加が必要なら関数を増やす形にする。

## セットアップと実行

```bash
cd scripts/pptx_slidegen
npm install                                   # 初回のみ（pptxgenjsをこのディレクトリにインストール）
node build_C_business_administration.js       # slides/1st_stage/C_business_administration.pptx を生成
```

## 配色・フォント（`lib.js`からexport）

```js
INK="171717" INK_SOFT="8A8A8A" RED="FF3B1F" LINE="E6E6E6" GHOST="F4F4F4" WHITE="FFFFFF"
F_HEAD="Yu Gothic"  F_BODY="Yu Gothic"  F_MONO="Consolas"
```

HTML版は日本語見出しに「Zen Kaku Gothic New」、等幅に「Space Mono」を使うが、PowerPointではユーザー環境にWebフォントは入っていないため、**日本語Windows/Officeにほぼ確実に入っている「Yu Gothic」**を採用した。等幅（符番・年度ラベル等）はラテン文字・数字のみなので「Consolas」で代用している。

### なぜ`transparency`オプションを使わないか（重要）

表紙・区切りスライドの半透明の白文字（`W_10`〜`W_94`という名前の定数）は、`color: WHITE, transparency: N`ではなく**あらかじめ背景色`#171717`と混ぜた実色**を使っている。

理由：pptxgenjsの`transparency`はOOXML上`<a:alpha>`として正しく書き込まれる（`scripts/office/validate.py`も通る）が、**Google Slidesがテキストの`<a:alpha>`を描画しない**（常に完全不透明で表示する）ことが実機検証で判明した。実際のPowerPointでは正しく半透明に見える可能性が高いが、閲覧環境を選ばず同じ見た目にするため、背景が確定している`#171717`に対してのみ有効なこの「事前混色」方式を採用している。**別の背景色に半透明白文字を置きたい場合は、その背景色に対して同じ計算（`blend = bg*(transparency/100) + white*(1-transparency/100)`）をして新しい定数を追加すること。** 詳細経緯は本ファイル末尾の「実装上の教訓」参照。

## レイアウトの基本座標（16:9、13.333×7.5インチ）

| 要素 | 座標 |
|---|---|
| ヘッダー kicker | x=0.55, y=0.35 |
| ヘッダー title | x=0.53, y=0.62, 25pt bold |
| ヘッダー overview | x=0.55, y=1.18, 13pt |
| ヘッダー下罫線 | y=1.62 |
| 本文エリア開始 | y≈1.85 |
| 頻出データ帯 | y=6.55（`addFreqBar`） |

本文を1カラム（全幅）にする場合：`proseX=0.55, proseW=12.25`
本文を2カラム（左テキスト2/3・右図解1/3）にする場合：`proseX=0.55, proseW=7.4` ／ `diagX=8.25, diagW=4.05`

## `lib.js`が提供する部品

| 関数 | 用途 |
|---|---|
| `newPres()` | 16:9のPresentationを作成 |
| `addCoverSlide(pres, opts)` | D1：科目の表紙（黒背景） |
| `addDividerSlide(pres, opts)` | D2：大分類の区切り（黒背景＋巨大ゴースト番号） |
| `addHeader(slide, opts)` | 全内容スライド共通のヘッダー（kicker/title/overview/タグ/罫線） |
| `addFreqBar(slide, opts)` + `mkYears(onSet)` | 頻出データ帯（頻出ランク・関連論点・10年分の出題年表）。**実データのみ使用、捏造禁止** |
| `addTermRows(slide, x, y, w, rows, opts)` | 2〜4件の「太字ラベル｜グレー説明文」の用語リスト |
| `addRowList(slide, x, y, w, rows, opts)` | N件の「太字名称｜等幅タグ｜グレー説明文」の行リスト（アンゾフの成長マトリクス等） |
| `addExamQuestion(slide, opts)` | 過去問チェックの設問／解答＆解説（`correctIndex`の有無で出し分け） |
| `drawScopeBars(slide, x, y, w, h, opts)` | 図解：狭い/広いの幅比較バー |
| `drawVennOverlap(slide, x, y, w, h, opts)` | 図解：2円の重なり（関連性・シナジー等） |

## 右側に図解を置くかどうかの判断基準

**内容ごとに判断し、デフォルトで全スライドに図解を付けない。** 具体的には、内容が「本質的に空間的・比較的な構造」を持つ場合にのみ右1/3に図解を置き、そうでない場合は全幅テキストにする。

判断の実例（企業経営理論C-1〜C-3で検証済み）：

- **C-1（VRIO分析）→ 図解なし・全幅**：V/R/I/Oの4基準はすでに左の用語リスト（`addTermRows`）で十分に一覧化できており、右に同じ4項目を円で並べ直しても情報が増えない（単なる重複表示）。
- **C-2（ドメインの2つの定義）→ 図解あり（`drawScopeBars`）**：「物理的定義（狭い）」と「機能的定義（広い）」という**視野の広さの違い**は文章だけでは伝わりにくく、幅の違うバーで見せると一目で分かる。図解が独自の情報を追加している。
- **C-3（多角化とシナジー）→ 図解あり（`drawVennOverlap`）**：「既存事業との重なり＝シナジー」はまさに円の重なりという空間的構造そのものであり、文章より図の方が本質を表す。

新しい論点を追加する際は、この3例と同じ問いを立てること：**「右に図を置いたら、左の文章だけでは伝わらない情報が本当に増えるか？」** 増えないなら置かない。

## 文字がボックスをはみ出さないための検証手順（レンダリング目視ができない環境向け）

この環境では`soffice`（LibreOffice）による`.pptx`のレンダリングが機能しない（`Error: source file could not be loaded`、あらゆる入力ファイルで再現）ため、HTML版のようにPlaywrightで見た目を目視確認することができない。代わりに以下の機械的な検証を必須とする：

1. **スキーマ検証**：`python scripts/office/validate.py <file>.pptx`（pptxgenjs付属のスキルディレクトリ内）
2. **境界チェック**：`python-pptx`で全シェイプの`left+width`・`top+height`がスライド範囲（13.333×7.5インチ）に収まっているかを確認
3. **折り返し行数の見積もりチェック**：全角文字は`フォントサイズ(pt)/72`インチ、半角文字はその半分の幅として概算し、テキストボックスの幅から必要行数を逆算、`行数×行高（フォントサイズ×1.35/72インチ目安）`がボックスの高さを超えていないか確認する

3の見積もりは実際に複数回、本番デッキ作成中の高さ不足を検出した（`具体例`ボックス等）。**新しいスライドを作ったら必ずこの3点セットを実行し、指摘があれば高さを調整してから納品すること。** 検証用スクリプトの雛形はこの仕様書のGitログ（該当コミット）のPythonコードを参照。

## 実装上の教訓

- **`element.scrollHeight === element.clientHeight`同様の罠はPowerPointにはない**が、代わりに「テキストボックスの`h`はレイアウトの目安に過ぎず、実際の文字がそれを超えても自動でクリップされない（PowerPoint側は溢れたまま表示する）」という別の罠がある。目視できない以上、上記の折り返し行数見積もりで正しいサイズを計算してから`h`を指定すること。
- **Google Slidesはテキストの`<a:alpha>`（透明度）を描画しない。** 半透明が必要な箇所は背景色との事前混色で実色を作ること（上記「なぜtransparencyを使わないか」参照）。
- **この環境では`soffice`によるpptx→PDF変換が機能しない**（`.txt`ファイルの変換すら失敗することを確認済み）。視覚検証はできない前提で、構造的な検証（スキーマ・境界・行数見積もり）を徹底すること。
- **図解は内容ごとに要否を判断する。** 「とりあえず全部のスライドに右カラムの図を置く」という機械的な適用はしない（上記「右側に図解を置くかどうかの判断基準」参照）。
