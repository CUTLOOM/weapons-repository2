---
name: lutz-container-loader
description: >
  Create or update LUTZ container loading data (装柜数据) Excel files.
  Use when the user provides a PI (Proforma Invoice) .xlsx file and asks to make
  its container loading data with net/gross weight calculations.
  Reads LUTZ产品净毛重.xlsx from the workspace or configurable path for weight lookups.
  Produces .xlsx output matching the reference template format.
---

# LUTZ Container Loading Data Generator

## Workflow

When user provides a PI .xlsx path, do:

1. **Read the PI** using openpyxl (Sheet `Proforma Invoice`)
2. **Extract**: Invoice No, Order No, Date, Notify Party, Container Type, Shipment Date, all product rows
3. **Lookup weights** from `LUTZ产品净毛重.xlsx` sheet `XXXLUTZ` or `Home 24` or `Momax` by matching model number (e.g., DC-24065A). Search all three sheets.
4. **Copy the template** from `assets/LUTZ装柜数据模板.xlsx`, modify in place

## Template Structure (R=row, C=col)

```
R1:   BAZHOU XINWEI FURNITURE CO. LTD. (merged A1:N1, sz20 bold, center wrap)
R2:   Supplier Adress (B2:N2 merged)
R3:   Telephone (B3:C3, F3:H3 merged)
R5:   箱号 | (value) | ETD | (value)
R6:   封号 | (value) | ETA | (value)
R7:   提单号 | (value) | 材质 | 布+金属
R8:   目的港 | (value)
R9:   船名航次 | (value)
R10:  总箱数 | =J31
R11:  总毛重 | =R31
R12:  总净重 | =P31
R13:  总体积 | =L31
R15:  PROFORMA INVOICE (merged A15:N15, sz16 bold center)
R16:  Notify Party | BDSK/XXXLutz | Date: | (date)
R17:  Invoice No.: | (invoice)
R18:  Order No.: | (order)
R20:  Photo | Item No. | Design No. | Description | Material | Qty | Price | Total | PU | CTN | CTN | Total | Logo | EAN Code | 净重 | 总净重 | 毛重 | 总毛重
R21:  (CBM sub-header: K21=CBM, L21=CBM)
R22+: PRODUCTS start here - each product occupies 1 row
R28:  Spareparts (FOC, CTN=1, CBM=0.014, 净重=1, 毛重=2)
R29:  (amount formula =SUM(H22:H27))
R30:  discounts | 0.045
R31:  TOTAL (formulas SUM for F,J,L,P,R columns)
R32:  TOTAL: (amount in words)
R34:  Date of Shipment:
R35-40: Payment, Incoterm, Port, etc.
R49-55: Bank info
R57:  Buyer | Suppliername
```

## Merged Cells (critical)

```
A1:N1, B2:N2, B3:C3, F3:H3,
B5:C5 through B14:C14 (container values),
A15:N15 (PROFORMA),
B16:D18 (Notify/Invoice/Order),
M16:N16, M17:N17, M18:N18,
K19:L19 (CBM sub-header),
A20:A21 through R20:R21 (Photo header + weight headers),
A28:E28, G28:H28, A29:G29, I29:N29 (discount/TOTAL rows),
A31:E31, B32:N32,
A34:B34 through A40:B40, C34:D34 through C40:D40, J36:L36, J37:L37, J39:L39, J40:L40 (footer),
A43:B45, C43:E43 through C45:E45, A47:I47,
B50:N50, B52:N52 through B55:N55, A57:H57
```

## Product Row Layout (each row, starting at R22)

Column | Content | Format
--- | --- | ---
B(2) | Item No. (e.g., 10780066/01) | sz10, left
C(3) | "Name\nModel" (e.g., "MATTI\nDC-24020C") | sz10, center, wrap=True
D(4) | Description | sz10, **left**, **wrap=True**
E(5) | Material | sz10, **left**, **wrap=True**
F(6) | Qty | sz10, center
G(7) | Unit price | sz10, center
H(8) | Amount (=F*G) | sz10, center
I(9) | Packing unit (2) | sz10, center
J(10) | CTN (=F/I) | sz10, center
K(11) | CBM per piece | sz10, center
L(12) | Total CBM (=F*K) | sz10, center
M(13) | Logo (Livetastic or MID.YOU) | sz10, center
N(14) | EAN Code | sz10, center
O(15) | **净重** per-unit (hardcoded value, NO formula) | sz10, center, thin border
P(16) | **总净重** (=O*F, formula) | sz10, center, thin border
Q(17) | **毛重** per-unit (hardcoded value, NO formula) | sz10, center, thin border
R(18) | **总毛重** (=Q*F, formula) | sz10, center, thin border
S(19) | **唛箱号** (e.g., 8NBU6-1) | sz10, center, NO border

## Modification Rules

### Products
- Products start at R22
- Variable count (6, 8, etc.)
- R22 + n products
- **Do NOT copy** the PI's Photo column or any numbering (1,2,3...) - those are user row markers, IGNORE

### Spareparts (R28)
- Always reserve below the last product
- `C1="Spareparts"`, `C6=(qty by user)`, `C7="FOC"`, `C10=1`, `C12=0.014`, `O(15)=1`, `Q(17)=2`
- Only qty comes from user; rest is fixed

### TOTAL Row (R31)
- F: `=SUM(F22:F27)` (adjust range for actual products)
- J: `=SUM(J22:J27)`
- L: `=SUM(L22:L27)`
- P(16): `=SUM(P22:P27)`
- R(18): `=SUM(R22:R27)`
- H(8): `=H30*95.5%` (discount)

### Summary References
- R10: `=J{total_row}` (总箱数)
- R11: `=R{total_row}` (总毛重)
- R12: `=P{total_row}` (总净重)
- R13: `=L{total_row}` (总体积)

### Container Info Labels (R5-R13)
- Labels (C1): sz14 bold center, thin borders
- 目的港(R8) C1: no bottom border
- 总箱数(R10) C1: no top border
- 总毛重(R11) C1: no top border

### Container Info Values (C2)
- All values: sz12, left aligned
- 箱号(R5): sz12 bold
- 封号(R6): sz12 bold (but value left-aligned)
- 提单号(R7): sz11
- 目的港(R7): sz12, Border(left=thin, right=thin) only
- 船名航次(R9): sz11.25
- 总箱数(R10): sz12, no top border, number_format='General"CTNS"'
- 总毛重(R11-R13): sz12, number_format='0.000_ ' for gross, '0.000_);[Red]\(0.000\)' for net/vol

### ETD/ETA (C4-C5)
- C4 labels (ETD/ETA/材质): sz14 bold center
- C5 values: sz12, center (布+金属 is sz12 left)

## Weight Lookup
From `LUTZ产品净毛重.xlsx`:
- Sheet `XXXLUTZ`: {model: net, gross}
- Sheet `Home 24`: {model: net, gross}
- Sheet `Momax`: {model: net, gross}
Search all three sheets; first match wins.

## Notify Party Mapping
Based on PI's Notify Party:
- If contains "XXXLutz KG": use "XXXLutz KG\nRömerstraße 39\nA-4600 Wels\nUID-Nr.: ATU65296645"
- If contains "BDSK": use "BDSK Handels GmbH & Co. KG\n(Billing address)\nMergentheimer Straße 59\nDE-97084 Würzburg\nGermany\nUID: DE279448078"

## 唛箱号 Pattern
`{order_no}-{n}` (e.g., 8NBU6-1, 8NBU6-2, ...) starting from 1.

## Template File
A ready-to-use template is bundled at `assets/LUTZ装柜数据模板.xlsx`.
Copy it as the starting point and modify data/values/formulas as needed.
Deviate from it only when the explicit instructions above require it.
