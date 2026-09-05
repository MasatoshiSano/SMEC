# 教科書 図解スライド

各教科書の全論点を、比較・フロー・マトリクス・階層・循環などの図解で整理したPowerPoint資料です。

| 科目 | 元の教科書 | スライド | 論点数 |
|---|---|---|---:|
| A. 経済学・経済政策 | [`A_economics_textbook.md`](../A_economics_textbook.md) | [`A_economics_visual_slides.pptx`](A_economics_visual_slides.pptx) | 36 |
| B. 財務・会計 | [`B_finance_accounting_textbook.md`](../B_finance_accounting_textbook.md) | [`B_finance_accounting_visual_slides.pptx`](B_finance_accounting_visual_slides.pptx) | 34 |
| C. 企業経営理論 | [`C_business_administration_textbook.md`](../C_business_administration_textbook.md) | [`C_business_administration_visual_slides.pptx`](C_business_administration_visual_slides.pptx) | 42 |
| D. 運営管理 | [`D_operations_management_textbook.md`](../D_operations_management_textbook.md) | [`D_operations_management_visual_slides.pptx`](D_operations_management_visual_slides.pptx) | 35 |
| E. 経営法務 | [`E_business_law_textbook.md`](../E_business_law_textbook.md) | [`E_business_law_visual_slides.pptx`](E_business_law_visual_slides.pptx) | 27 |
| F. 経営情報システム | [`F_information_systems_textbook.md`](../F_information_systems_textbook.md) | [`F_information_systems_visual_slides.pptx`](F_information_systems_visual_slides.pptx) | 24 |
| G. 中小企業経営・中小企業政策 | [`G_sme_management_policy_textbook.md`](../G_sme_management_policy_textbook.md) | [`G_sme_management_policy_visual_slides.pptx`](G_sme_management_policy_visual_slides.pptx) | 28 |

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
