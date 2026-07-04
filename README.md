# Landmark Channel Toolkit V3

V3 combines:

1. PPT核对Report: audit Excel vs PPT without modifying PPT.
2. PPT价格更新: update only existing Unit prices in PPT and generate a report.

Notes:
- Missing units are not automatically generated; they are listed in the report.
- Price update tries to preserve formatting by replacing only the price text inside existing text runs.
- Supports `.xlsx` and many `.xls` exports. If an `.xls` fails, save it as `.xlsx` and upload again.
