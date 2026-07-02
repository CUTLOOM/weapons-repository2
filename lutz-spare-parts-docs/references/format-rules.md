# Format Rules

## 补件唛头
Cols: A=20 B=35 C=24 D=20 E=28
R1: Yellow bg. R3: C3:D3 merged. All bold headers.
All black text. Thin borders. No empty rows.
Images: OneCellAnchor centered 70% of cell.
Qty: {n}SETS newline {n}X{per_set}={total}PCS

## Spare Parts List
Cols: A=25 B=27.82 C=28 D=20 E=30 F=12
C1:D1 C2:D2 C3:D3 merged. All headers bold.
Aggregate: normalize desc, merge Feet/FEET, Price=Qty*0.001
Images: same centering as 补件唛头

## CI Multi-Container
COLUMNS: A-L (PL), A-H (CI)
HEADER STABLE: PL R1-R6, CI R1-R8 (never delete these rows)
CONTAINER DELETE: PL `delete_rows(7,12)`, CI `delete_rows(9,15)`

Each container section: Container info → Products → Spare → Subtotal
- Subtotal formulas: E=SUM, F=SUM, H=SUM(J), J=SUM(J), L=SUM(L) for PL
- Subtotal formulas: E=SUM, F=SUM, H=SUM (G NOT summed) for CI
- "In total": sums all containers' data
- CI final: =In_total_H * 0.955

SPECIAL ITEMS: Credit note / 罚款 → alert user first

SPARE EXTRACTION: Check column S for Pos.No (product row), column A for 'Spareparts' (spare row)