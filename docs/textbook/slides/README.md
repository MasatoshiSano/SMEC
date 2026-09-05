# 教科書 図解スライド

教科書の全論点を、比較・フロー・マトリクス・階層・循環などの図解で整理したPowerPoint資料です。現在はA科目を収録しています。

| 科目 | 元の教科書 | スライド | 論点数 |
|---|---|---|---:|
| A. 経済学・経済政策 | [`A_economics_textbook.md`](../A_economics_textbook.md) | [`A_economics_visual_slides.pptx`](A_economics_visual_slides.pptx) | 36 |

## 再生成

```bash
python3 -m pip install -r scripts/requirements-slides.txt
python3 scripts/generate_textbook_slides.py
```

次の科目を追加する場合は、科目記号を指定します。

```bash
python3 scripts/generate_textbook_slides.py --subject B
```

スライドは教科書本文を基に生成しています。法令・統計・制度の数値は、受験年度の公式情報も確認してください。
