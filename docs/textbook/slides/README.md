# 教科書 図解スライド

各教科書の全論点を、全体構造・セクション別要点・学習のツボを図解でコンパクトに整理したPowerPoint資料です。

## 標準図解テンプレート

| ファイル | 内容 |
|---|---|
| [`SMEC_visual_diagram_template.pptx`](SMEC_visual_diagram_template.pptx) | 汎用図解12パターン＋経済学専用グラフ6パターン（全21枚、ガイド3枚を含む） |
| [`TEMPLATE_GUIDE.md`](TEMPLATE_GUIDE.md) | 配色・文字・余白・パターン選択基準・禁止事項 |

テンプレートはすべてPowerPoint上で編集できる図形で作成しています。

```bash
python3 scripts/generate_smec_visual_template.py
```

## 科目別スライド

| 科目 | 元の教科書 | スライド | 論点数 | 構成 |
|---|---|---|---:|---|
| A. 経済学・経済政策 | [`A_economics_textbook.md`](../A_economics_textbook.md) | [`A_economics_visual_slides.pptx`](A_economics_visual_slides.pptx) | 36 | 表紙・全体マップ・8セクション要点・まとめ (計11枚) |

## 再生成

```bash
python3 -m pip install -r scripts/requirements-slides.txt
python3 scripts/generate_textbook_slides.py
```

特定科目だけ生成する場合、またはモードを切り替える場合:

```bash
# 要点ダイジェスト版（デフォルト、10〜11枚）
python3 scripts/generate_textbook_slides.py --subject A --mode summary

# 各論点ごとの図解スライド版（36〜46枚）
python3 scripts/generate_textbook_slides.py --subject A --mode visual

# 教科書全文詳細つき完全版（133枚）
python3 scripts/generate_textbook_slides.py --subject A --mode full
```

スライドは教科書本文を基に生成しています。法令・統計・制度の数値は、受験年度の公式情報も確認してください。
