# 📂 Workbook

`KPR_Template.xlsx` is the macro-free template of the K-PRICING workbook: a
**Home** sheet and the **K-PRICING** Ribbon (`src/ribbon/customUI14.xml`, byte
for byte). It holds no VBA project and no document properties.

It is generated, never edited in Excel:

```bash
python3 tools/workbook_template.py --write   # after changing the Ribbon or the Home sheet text
python3 tools/workbook_template.py --check   # CI: fails when the file differs from its source
```

`tools/Check-KPR.ps1` saves it as an `.xlsm`, imports one commit's modules,
compiles them, writes the commit on the Home sheet (cell `B13`) and builds the
**KPR Dates Demo** sheet into it with `KPR_Demo_BuildDates`, so the demo stays
generated from `examples/modules/` rather than stored twice. No workbook
document modules (`ThisWorkbook`, sheet modules) exist yet; they would live
here too.
