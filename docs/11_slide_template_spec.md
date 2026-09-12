# スライドテンプレート仕様書（Swiss Grid）

`docs/textbook/` の教科書コンテンツを、図解の多いスライドに仕立てるためのデザインテンプレート仕様。過去問完全マスター方式の教科書と対になる「見せ方」のルールをここに一元化する。実物のサンプルは `docs/slide_template/*.html`（ブラウザで直接開ける単体HTML）を参照。

## 採用した方向性

5案（Swiss Grid／Soft Rounded Pastel／Editorial Serif／Blueprint Technical／Neubrutalist／Ink & Washi）を試作し、**Swiss Grid**（スイス国際様式：モノクロ＋差し色1色、罫線グリッド）を採用した。理由：
- 図解の種類（概念図・座標軸グラフ・組織図・記号記法）が科目によって大きく異なるため、色数を絞ったニュートラルな様式の方が科目をまたいで破綻しにくい。
- 罫線・余白・番号による構成は、詰め込みたい情報量が多い試験対策資料と相性が良い。

## 配色・タイポグラフィ

```css
--ink:      #171717;  /* 本文・見出し・主要な線 */
--ink-soft: #8a8a8a;  /* 補助テキスト・キッカー */
--red:      #FF3B1F;  /* 差し色。強調・警告・ひっかけポイントにのみ使う */
--line:     #e6e6e6;  /* 罫線・淡い区切り */
--ghost:    #f4f4f4;  /* 背景に沈めるゴースト数字など */
--green:    #2f8f5b;  /* 科目別グラフでの「増加・利益・良い」の意味づけにのみ使用 */
```

- 見出し・数字：**Zen Kaku Gothic New**（700〜900）
- 符番・英語ラベル・軸ラベル：**Space Mono**
- 本文・説明文：**Noto Sans JP**
- 背景は白（`#fff`）で統一。色数を増やさない。

## 絶対ルール（構成）

1. **すべてのスライドは、タイトル直下に1文の概要文（`.overview`）を置く。** 図だけで完結させない。
2. **図を使う場合は、必ず上下左右いずれかの位置に説明の文章・箇条書きを併記する。** 「図のみのスライド」は作らない。逆に、図解に向かない論点（用語の羅列・注意点の整理など）は無理に図解化せず、後述の「文章主体パターン」を使う。
3. **囲み枠（ボーダー）は、構造上ノードや領域を区別する必要がある図解（マトリクス・ツリー・ネットワーク・比較表など）にのみ使う。** 単なる要素の列挙は、罫線・余白・番号で区切り、四角で囲まない（Swiss Gridは「箱で囲う」スタイルではない）。
4. 配置（構図）は内容によって変える。基本パターンは3つ：
   - 図は左（メイン）＋右に凡例・補足文（`.gside` + `.tside`）
   - 図を中央〜大きめに＋下にキャプション1行（`.gcap-wrap` / `.gcap-graph` + `.gcap-text`）
   - 左右2枚を並べて対比（`.split-wrap` / `.split-panel`、Before/After・2つの構造の比較などに使う）

## 実装上の注意点（重要・過去のバグからの教訓）

このテンプレートは実装時に2つの罠にはまり、モバイル実機で「グラフがヘッダーに重なる」「文字がはみ出す」という崩れを繰り返し起こした。再実装する際は必ず以下を守ること。

### 1. カード自体の16:9固定は `aspect-ratio` プロパティ単体に頼らない

`aspect-ratio:16/9` を直接カードに指定すると、環境によってレイアウト計算が安定しないことがあった。**`padding-top` パーセントトリック**（全ブラウザで確実に効く古典的手法）で高さを確定させる。

```css
.mock-frame{position:relative;width:100%;}
.mock-frame::before{content:"";display:block;padding-top:56.25%;} /* 9/16 */
.mock{position:absolute;inset:0;overflow:hidden;background:#fff;
  border:1px solid var(--line);box-shadow:0 1px 3px rgba(0,0,0,.04);
  display:flex;flex-direction:column;container-type:inline-size;}
```

HTML側は必ず `.mock` を `.mock-frame` で包む：

```html
<div class="mock-frame"><div class="mock"> ... </div></div>
```

### 2. カード内の文字サイズは固定pxではなく、コンテナクエリ（`cqw`）で指定する

カードはページ幅に応じて縮む（特にスマホ）。文字が固定px指定だと、カードだけが縮んで文字がはみ出し・重なりの原因になる。`.mock` に `container-type:inline-size` を指定し、内部の `font-size` はすべて **カード幅に対する%（`cqw`）** で書く。デザイン基準幅は約972px（`Ncqw = Npx / 972 * 100`）。

```css
.head .title{font-size:1.7cqw;}      /* 元16.5px相当 */
.head .overview{font-size:1.03cqw;}  /* 元10px相当 */
```

### 3. グラフ（SVG）の枠は、高さを強制せず「収まるだけ縮む」方式にする

SVGを `padding-top` で強制的に一定の高さにする実装は、見出しが2行になるなどで親のスペースが狭くなったときに**縮まずにあふれ、中央寄せの結果として上下（特に上＝ヘッダー側）にはみ出す**という不具合を起こした。正しくは、画像の `object-fit:contain` と同じ考え方で「親が確保できたスペースの中で、SVG自身のアスペクト比を保ったまま縮小する」実装にする。

```css
.chart-frame{width:100%;height:100%;flex:1;display:flex;
  align-items:center;justify-content:center;min-width:0;min-height:0;}
.chart-frame svg{display:block;width:auto;height:auto;
  max-width:100%;max-height:100%;}
```

親側のflexコンテナ（`.gside` `.gcap-graph` `.split-panel` 等）には必ず `min-height:0` を入れること（入れないとflexアイテムが縮まずあふれる）。

## 12種類の汎用図解パターン

`docs/slide_template/patterns_12.html` 参照。科目を問わず使える基本部品。

| # | パターン名 | 構図 |
|---|---|---|
| 1 | 定義・要素カード | 罫線区切りの横並び（並列する複数要素） |
| 2 | 左右比較 | 中央にVSマーク、左右対称スプリット |
| 3 | プロセス・手順 | 番号付きステップ＋矢印 |
| 4 | 因果関係 | 左に原因、右に複数の結果を枝分かれ |
| 5 | 循環モデル | 中央に円環図＋下にキャプション |
| 6 | 階層・ピラミッド | 左に図、右に各階層の説明リスト |
| 7 | タイムライン | 横一直線の時系列 |
| 8 | 2×2マトリクス | 左に図、右に軸の意味の説明 |
| 9 | ツリー・分類 | 上から下へ枝分かれ＋具体例1行 |
| 10 | ネットワーク・関係図 | 左に凡例（線種の意味）、右に図 |
| 11 | 数式・投入産出モデル | 上に数式、下に入力→処理→出力 |
| 12 | グラフ・表・指標比較 | 上に簡易グラフ、下に正確な表 |

## 科目別専用グラフ（5科目・計21パターン）

12の汎用パターンではカバーしきれない、科目固有の作法・記法を持つ図解。試験の過去問での描かれ方に合わせている。

### 経済学・経済政策（A） — `economics_6.html`
需要供給曲線 ／ 曲線のシフトと均衡変化（Before/After2枚） ／ 45度線分析（ケインズ・クロス） ／ IS-LM分析（政策シフト付き） ／ 余剰・死荷重（ハッチングで死荷重を表現） ／ 時系列・景気循環（好況・不況を面で色分け）

### 財務・会計（B） — `finance_3.html`
損益分岐点図（CVP分析） ／ キャッシュフロー（ウォーターフォール図） ／ 財務諸表の連動（三表連動）

### 企業経営理論（C） — `business_admin_3.html`
SWOT分析 ／ 5フォース分析（強度バー付き） ／ 組織構造の比較（実際の組織図で職能別×事業部制を対比）

### 運営管理（D） — `operations_3.html`
工程図（JIS工程分析記号：○加工・□検査・→運搬・D停滞・▽貯蔵） ／ レイアウトの比較（機能別×製品別、2枚並列） ／ 在庫推移（のこぎり刃モデル）

### 経営情報システム（F） — `information_3.html`
OSI参照モデル（7階層の積み上げ図） ／ ER図（実体・関係・多重度） ／ 処理フロー（JIS流れ図記号：開始/処理/判断/入出力）

## 文章主体パターン（3種）— 図解に向かない論点向け

`docs/slide_template/text_primary_3.html` 参照。

### なぜ必要か

上記の図解パターンは試作段階では見栄えの検証を優先し、グラフや図解が大きく載る構成中心になっていた。しかし実際の `docs/textbook/` の論点は、座標軸グラフ・組織図のように「絵で描ける」ものは一部で、大半は用語の定義・ひっかけポイントの整理・複数概念の列挙など文章中心の内容である。すべての論点を無理に図解パターンへ当てはめると、意味の薄い箱や矢印を量産することになる。本番スライド作成では、論点の性質に応じて「図解主体」（12汎用＋科目別専用グラフ）と「文章主体」（T1〜T3）を使い分けること。

| # | パターン名 | 構図 |
|---|---|---|
| T1 | 概念解説（文章primary＋小さいアイコン） | 左に小さいアイコン1つ、右にほとんどのスペースを説明文にあてる。図解しにくい抽象的な概念向き |
| T2 | よくある混同の整理（文章primary＋警告マーク） | 左に小さい警告アイコン、右に対比構造の説明文。似た用語・紛らわしい選択肢を区別させたい場合向き |
| T3 | 箇条書きリスト（図なし、罫線＋番号のみで構成） | 図解を一切持たず、罫線区切り＋番号＋名称＋説明の列。対比すべき複数の理論・法令・制度などの列挙向き |

### 実装上の注意点（追加の教訓：小アイコン・箇条書きも「縮むがあふれない」原則で）

T1/T2の小アイコン（`.icon-mark`）やT3の各行（`.rows .row`）でも、グラフの`.chart-frame`と全く同じクラスのバグが再発した。修正方針も同じ「縮むが、あふれない」原則で統一する。

- `.icon-mark`の高さを固定`cqw`にすると、カード高さが縮んだときにアイコン自身は縮まずあふれる → `height:100%;max-height:6.5cqw`として親の実高さに追従させ、中のSVGは`width:auto;height:auto;max-width:100%;max-height:100%`で「収まるだけ縮む」。
- `.prose`（説明文）や`.rows`（箇条書き）を`justify-content:center`にすると、内容が入りきらないときに上下対称にあふれ、上側がヘッダーに重なる → `justify-content:flex-start`＋`overflow:hidden`＋`min-height:0`にして、万一入りきらない場合も下端で安全に切れるようにする（上＝ヘッダー側には絶対にあふれさせない）。

```css
.icon-wrap{flex:1;display:flex;gap:2cqw;min-height:0;min-width:0;overflow:hidden;}
.icon-mark{flex:0 0 auto;width:6.5cqw;height:100%;max-height:6.5cqw;
  display:flex;align-items:center;justify-content:center;min-width:0;min-height:0;}
.icon-mark svg{display:block;width:auto;height:auto;max-width:100%;max-height:100%;}
.prose{flex:1;min-width:0;min-height:0;display:flex;flex-direction:column;
  justify-content:flex-start;gap:0.9cqw;overflow:hidden;}
.rows{flex:1;display:flex;flex-direction:column;justify-content:flex-start;
  min-height:0;overflow:hidden;}
```

## 頻出データ帯（標準コンポーネント・全パターン共通）

`docs/slide_template/freq_data_bar.html` 参照。

### 経緯：16:9固定と実際の文章量のミスマッチ

図解主体・文章主体のパターンを一通り試作した後、`docs/textbook/C_business_administration_textbook.md` のC-7・C-8を**要約せず全文そのまま**流し込んで実際の見え方を検証したところ、16:9カードの下半分近くが白紙のまま埋まらないことが分かった。これは「モック文章が短すぎた」からではなく、**教科書1論点あたりの文章量に対して16:9という箱が構造的に大きい**という、パターンの試作だけでは気づけなかった問題だった。

対策として「16:9は維持し、文字を拡大するのではなく、情報量そのものを増やす」方針を採用。ただし内容を水増し（捏造）するのではなく、**リポジトリに既にある実データ**を使う：`problem_sets/1st_stage/<科目記号>_*.md` に全論点の頻出ランク（A/B/C）と出題年度（2016〜2025年度、過去10年）が既に集計済みなので、これをスライド下部の帯として機械的に追加する。

### 構成要素

| 要素 | 内容 |
|---|---|
| 頻出ランクチップ | A/B/Cを丸チップで表示（`problem_sets/1st_stage/<科目>.md`の「頻出ランク」列） |
| 関連論点 | 同じ分野・セットで出題されやすい論点への相互参照（例：C-7とC-8は対になる論点） |
| 出題年度タイムライン | 2016〜2025年度の10年分を●（出題あり）／○（出題なし）で並べる年表。「'16 '17 '19…」と年度を羅列するだけより、パッと見て出題頻度のムラが分かり、かつ視覚的に必要な面積を稼げる |

```css
.freq-section{flex:0 0 auto;padding-top:1cqw;margin-top:1cqw;border-top:1px solid var(--ink);
  display:flex;flex-direction:column;gap:0.75cqw;}
.freq-top{display:flex;align-items:center;gap:1.5cqw;}
.freq-top .rank-chip{flex:0 0 auto;font-family:"Space Mono",monospace;font-weight:700;font-size:1.15cqw;
  color:#fff;background:var(--red);width:1.9cqw;height:1.9cqw;border-radius:50%;
  display:flex;align-items:center;justify-content:center;}
.freq-timeline{display:flex;align-items:center;gap:0.2cqw;}
.freq-timeline .yr{flex:1;display:flex;flex-direction:column;align-items:center;gap:0.35cqw;}
.freq-timeline .dot{width:1.5cqw;height:1.5cqw;border-radius:50%;border:1.4px solid var(--line);background:#fff;}
.freq-timeline .dot.on{background:var(--red);border-color:var(--red);}
```

`.body`を`display:flex;flex-direction:column`にし、図・文章部分（`flex:1`）の下に`.freq-section`（`flex:0 0 auto`）を並べる。図・文章部分側の実装教訓（縮むがあふれない原則）はそのまま維持すること。

### 検証結果と、詰め込みパターン（T4/T5）を不採用にした経緯

同じ白空間問題への対策として、文章主体パターンを左右2列・上下2段に分割して情報量を倍にする「T4：2列密集リスト」「T5：概念＋ひっかけ上下結合」も試作した。しかしコンテナクエリの基準（カード幅 or カード自体）を列・段に対して正しく再設定してもなお、文字を大幅に縮小しないと収まらず、可読性を大きく犠牲にすることが分かった。実データによる頻出データ帯の追加だけで白空間問題は十分に解消できたため、**T4/T5の詰め込み路線は不採用**とした。1論点1枚を無理に2論点分に増量するより、素直に2枚に分けた方が読みやすい。

## 未着手・今後の拡張候補

- **経営法務（E）・中小企業経営政策（G）の専用パターン**：現時点では作成していない。E科目は法令の存続期間比較や手続きフロー、G科目は中小企業基本法の分類基準表など、法律・政策系ならではの見せ方を検討する余地がある。
- **実データでの本番スライド作成**：白空間問題の検証でC-7・C-8の2論点は教科書全文＋頻出データ帯で組んで問題ないことを確認したが、これは検証用の個別サンプルであり、全226論点を通しで作る本番作業はまだ行っていない。
- **PowerPoint(.pptx)版**：`scripts/slidegen/`（python-pptx製）で先行して試作したが、生成物がPowerPointで「修復が必要」というエラーになる不具合が未解決のまま、HTML版に方針転換した経緯がある。pptx版が必要になった場合は、この不具合の原因調査から再開すること。
