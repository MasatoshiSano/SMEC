# 情報源一覧（入手元）

このリポジトリに収録した情報・データファイルの出典を、取得日・URLとともに記録する。
すべて公的機関・試験実施団体が無償で公開している一次情報のみを対象とし、市販の教材・書籍そのものは複製していない（市販教材は `docs/04_study_materials_guide.md` にタイトルと入手先のみ記載）。

## 1. 試験実施団体・制度全般

| 項目 | 発行元 | URL | 取得日 | 保存先 |
|---|---|---|---|---|
| 令和8年度 中小企業診断士 第1次試験案内（全文） | 一般社団法人 日本中小企業診断士協会連合会（JF-CMCA、旧・中小企業診断協会） | https://www.jf-cmca.jp/attach/test/r08/r08_1ji_annai.pdf | 2026-09-05 | `docs/official_guides/R08_1ji_shiken_annai.pdf` |
| 令和8年度 中小企業診断士 第2次試験案内（全文） | 同上 | https://www.jf-cmca.jp/attach/test/r08/r08_2ji_annai.pdf | 2026-09-05 | `docs/official_guides/R08_2ji_shiken_annai.pdf` |
| 試験問題ダウンロードページ（過去問一覧） | 同上 | https://www.jf-cmca.jp/contents/010_c_/shikenmondai.html | 2026-09-05 | リンク一覧は本ドキュメント末尾参照 |
| 申込者・合格者統計資料 一覧ページ（1次） | 同上 | https://www.jf-cmca.jp/contents/010_c_/001_shiken_kakokekka_1.html | 2026-09-05 | - |
| 申込者・合格者統計資料 一覧ページ（2次） | 同上 | https://www.jf-cmca.jp/contents/010_c_/001_shiken_kakokekka_2.html | 2026-09-05 | - |
| 令和8年度からの試験改正点について（受験料改定・口述試験廃止） | 中小企業庁 | https://www.chusho.meti.go.jp/shindanshi/2026/260205shindanshi.html | 2026-09-05 | 内容は `docs/00_exam_system_overview.md` に反映 |
| 中小企業診断士試験に関するFAQ | JF-CMCA | https://www.jf-cmca.jp/contents/013_c_faq/001_faq_shiken.html | 2026-09-05 | - |

**組織名について**：試験実施団体は令和6年（2024年）10月に「一般社団法人中小企業診断協会」から「一般社団法人日本中小企業診断士協会連合会（JF-CMCA）」に改称された。ドメインも `j-smeca.jp` から `jf-cmca.jp` に変更されている。市販教材サイトや古い解説記事では旧称・旧ドメインの記載が残っている場合があるため注意。

## 2. 1次試験 科目シラバス（試験科目設置の目的および内容）

- 出典：`docs/official_guides/R08_1ji_shiken_annai.pdf` の「13．試験科目設置の目的と内容」（12～20ページ）
- 発行元・URLは上記1次試験案内と同じ
- 全文を `docs/01_1st_stage_syllabus.md` に転記済み

## 3. 過去問題（第1次試験・第2次試験）

- 一覧ページ：https://www.jf-cmca.jp/contents/010_c_/shikenmondai.html
- ファイル本体の格納パターン：`https://www.jf-cmca.jp/attach/test/shikenmondai/<年度ディレクトリ>/<ファイル名>.pdf`
- 取得日：2026-09-05
- 保存先：
  - 第1次試験（平成19年度〜令和8年度、科目A〜G）：`past_exams/1st_stage/1ji<年>/`
  - 第2次試験（平成19年度〜令和7年度、事例A〜D）：`past_exams/2nd_stage/questions/2ji<年>/`
  - 令和5年度第1次試験（再試験）：`past_exams/1st_stage/1ji(sai)2023/`
- 令和8年度第2次試験は2026年10月25日実施予定のため、本リポジトリ作成時点（2026年9月5日）では未実施・問題未公開。

## 3.5 第1次試験 正解・配点

- 格納パターン：`https://www.jf-cmca.jp/attach/test/<年度>/1ji_seikai/<西暦><科目記号>.pdf`
- 取得日：2026-09-05（令和8年度分のみ）
- 保存先：`past_exams/1st_stage_answers/r08/`

## 4. 第2次試験「出題の趣旨」

- 公式に模範解答は公表されないが、「出題の趣旨」（各事例の出題意図の要約）が年度ごとに公表されている。
- 格納パターン：`https://www.jf-cmca.jp/attach/test/<年度(r02〜r07)>/2ji_shushi/<年度>_2ji_shushi_jirei<1〜4>.pdf`
- 取得日：2026-09-05
- 保存先：`past_exams/2nd_stage/shushi/<年度>/`
- 令和元年度以前は同一パターンでの公開が確認できなかった（未収録）。

## 5. 合格率・受験者数などの公式統計

- 出典：JF-CMCA「申込者・合格者にかかる統計資料」
- 取得日：2026-09-05
- 保存先：`docs/official_stats/1st/`（第1次試験、令和2〜7年度）、`docs/official_stats/2nd/`（第2次試験、令和2〜7年度）
- 抽出した数値は `docs/03_exam_statistics.md` に一覧化

## 6. 中小企業白書・中小企業政策

| 資料 | 発行元 | URL | 取得日 | 保存先 |
|---|---|---|---|---|
| 2026年版中小企業白書 概要 | 中小企業庁 | https://www.chusho.meti.go.jp/pamflet/hakusyo/2026/PDF/chusho/01Hakuksyo_gaiyo_web.pdf | 2026-09-05 | `references/whitepaper/R08_chusho_hakusho_gaiyo.pdf` |
| 2026年版中小企業白書 目次・凡例 | 同上 | https://www.chusho.meti.go.jp/pamflet/hakusyo/2026/PDF/chusho/02Hakusyo_mokuji_hanrei_web.pdf | 2026-09-05 | `references/whitepaper/R08_chusho_hakusho_mokuji.pdf` |
| 中小企業施策（令和7年度実施・令和8年度予算等） | 同上 | https://www.chusho.meti.go.jp/pamflet/hakusyo/2026/PDF/chusho/05Hakusyo_R7sesaku_web.pdf | 2026-09-05 | `references/whitepaper/R08_chusho_seisaku_chusho.pdf` |
| 小規模企業施策 | 同上 | https://www.chusho.meti.go.jp/pamflet/hakusyo/2026/PDF/chusho/07Hakusyo_R8sesaku_web.pdf | 2026-09-05 | `references/whitepaper/R08_chusho_seisaku_shokibo.pdf` |
| 中小企業白書索引ページ | 中小企業庁 | https://www.chusho.meti.go.jp/pamflet/hakusyo/index.html | 2026-09-05 | - |

白書全文（`00Hakusyo_zentai.pdf`）はファイルサイズが大きいため未取得。必要な場合は上記索引ページから各章のPDFを個別取得すること。

## 7. 市販教材・学習法（リンクのみ、著作物は複製せず）

`docs/04_study_materials_guide.md` にまとめて記載。検索により存在を確認した情報源：

- TAC出版オンラインストア（スピードテキスト、みんなが欲しかった！シリーズ、最速合格のための過去問題集など）
- アガルートアカデミー、伊藤塾、STUDYing、資格の学校TACの各コラム記事（試験制度解説）
- 中小企業診断士試験コラム各種（勉強時間・合格率の考察記事）

## 8. 著作権・利用上の注意

- 過去問題・出題の趣旨・統計資料は、JF-CMCA公式サイトで無償公開されている一次資料をそのまま保存したものである。個人の学習目的での保管・参照を前提とし、再配布・商用利用は行わないこと。
- 中小企業白書は中小企業庁が公開する政府刊行物であり、出典を明記の上での引用・利用が想定されている。
- 市販の参考書・問題集（TAC、LEC、スタディング、同友館「ふぞろいな合格答案」など）は著作物であるため、本リポジトリには実体を含めず、書誌情報とリンクのみを記録する。

## 9. 追加収集分（欠落補完）

初回収集で未収録だった一次資料を追加取得した。すべてJF-CMCA・中小企業庁が無償公開する一次資料である。取得日：2026-09-05。

| 追加資料 | 発行元 | URL（パターン／一覧ページ） | 保存先 |
|---|---|---|---|
| 第1次試験 正解と配点（令和5〜7年度＝R05〜R07） | JF-CMCA | 一覧: `https://www.jf-cmca.jp/contents/010_c_/010_c_r0X_shiken/R0X_1ji_shiken_kaitou.html`／本体: `https://www.jf-cmca.jp/attach/test/r0X/1ji_seikai/<西暦><a〜g>.pdf` | `past_exams/1st_stage_answers/r05〜r07/` |
| 第2次試験「出題の趣旨」（平成19〜令和元年度＝H19〜R01） | JF-CMCA | 一覧: `https://www.jf-cmca.jp/contents/010_c_/001_shiken_kakokekka_syusi.html`／本体: `https://www.jf-cmca.jp/attach/test/<年度>/<年度>_2ji_shushi/<年度>_2ji_shushi_jirei1〜4.pdf` | `past_exams/2nd_stage/shushi/h19〜h30, r01/` |
| 公式統計（第1次・第2次、平成19〜31年度＝H19〜H31） | JF-CMCA | 一覧: `001_shiken_kakokekka_1.html`（1次）／`001_shiken_kakokekka_2.html`（2次）／本体: `https://www.jf-cmca.jp/attach/test/<年度>/<年度>_1ji_toukei.pdf`・`_2ji_toukei.pdf` | `docs/official_stats/1st/`・`docs/official_stats/2nd/` |
| 中小企業白書・小規模企業白書 概要版（2022〜2025年版） | 中小企業庁 | `https://www.chusho.meti.go.jp/pamflet/hakusyo/<西暦>/PDF/<西暦>gaiyou.pdf` | `references/whitepaper/<西暦>_chusho_shokibo_hakusyo_gaiyou.pdf` |

これにより、第2次「出題の趣旨」は**平成19年度〜令和7年度の全年度**、公式統計は**平成19年度〜令和7年度**、白書概要は**2022〜2026年版**を収録し、初回収集時の未収録範囲を解消した。

### 収集の再現用スクリプト

収集手順を再現・検証できるよう、取得に用いたスクリプトとURL一覧を `scripts/` に保存している。

- `scripts/crawl_pdfs.py` … 年度別HTMLページを巡回し、相対リンクを絶対URLに解決してPDFリンクを抽出する。
- `scripts/download_mondai.sh` … URL一覧を受け取り、レート制御（0.3秒間隔）付きで一括ダウンロードする。
- `scripts/mondai_urls.txt` … 過去問（第1次148・第2次76＝計224）の全ソースURL一覧。
- `scripts/ref_urls.txt` … 統計・試験案内・受験者数推移PDFのソースURL一覧。
- `scripts/misc_urls.txt` … 出題の趣旨・正解表・試験案内のソースURL一覧。

**ネットワーク上の注意**：旧ドメイン `www.j-smeca.jp` は本収集環境のDNSで解決できず、現行ドメイン `www.jf-cmca.jp` から取得した（組織改称・ドメイン移行のため）。中小企業庁サイトはブラウザ相当のUser-Agentを付与しないと403を返す場合がある。
