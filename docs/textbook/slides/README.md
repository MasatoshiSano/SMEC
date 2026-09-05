# 教科書 図解スライド

教科書の全論点を図解し、基本概念・表・公式・具体例・ひっかけポイント・過去問参照を省略せず詳細スライドに収録したPowerPoint資料です。

| 科目 | 元の教科書 | スライド | 論点数 |
|---|---|---|---:|
| A. 経済学・経済政策 | [`A_economics_textbook.md`](../A_economics_textbook.md) | [`A_economics_visual_slides.pptx`](A_economics_visual_slides.pptx) | 36 |

## 再生成

```bash
python3 -m pip install -r scripts/requirements-slides.txt
python3 scripts/generate_textbook_slides.py
```

特定科目だけ生成する場合は、科目記号を指定します。

```bash
python3 scripts/generate_textbook_slides.py --subject C
```

スライドは教科書本文を基に生成しています。法令・統計・制度の数値は、受験年度の公式情報も確認してください。
