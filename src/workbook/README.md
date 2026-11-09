# 📂 Workbook

`KPR_Template.xlsx` is the macro-free template of the K-PRICING workbook: a
**Home** sheet and the **K-PRICING** Ribbon (`src/ribbon/customUI14.xml`, byte
for byte). It holds no VBA project and no document properties.

The Home sheet follows the K-SACCR workbook template: Aptos; column A is the
black navigation column, with **Home** highlighted and a link to **KPR Dates
Demo**; row 1 is the black band with the gold wordmark, row 2 the blue band
(`#265A99`) with the sheet title and row 3 the grey subtitle; rows 1 to 3 and
column A are frozen. Sections in navy (`#1F3864`) hold bold labels in column C
and wrapped text in column D: what K-PRICING does, how to use it, its sheets
and the commit it was built from.

It is generated, never edited in Excel:

```bash
python3 tools/workbook_template.py --write   # after changing the Ribbon or the Home sheet text
python3 tools/workbook_template.py --check   # CI: fails when the file differs from its source
```

`tools/Check-KPR.ps1` saves it as an `.xlsm`, imports one commit's modules,
compiles them, writes the commit on the Home sheet (cell `D22`) and builds the
**KPR Dates Demo** sheet into it with `KPR_Demo_BuildDates`, so the demo stays
generated from `examples/modules/` rather than stored twice. No workbook
document modules (`ThisWorkbook`, sheet modules) exist yet; they would live
here too.
