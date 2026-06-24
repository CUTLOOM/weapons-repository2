---
name: chuhuoshujuzhizuo
description: Generate 5 export shipping documents from container loading PI data.
metadata:
  short-description: Generate 5 export shipping documents from PI data
---

# Shipping Documents Generator

Generate 5 export shipping documents from a Proforma Invoice (PI / 装柜数据) Excel file.

## Workflow

For each new order, generate 5 files in order.

Output folder: 出货文件{BillOfLadingNo}

### 0. Read PI Data
Key cells: B5(Container), B6(Seal), B7(BL), B9(Vessel), B10(CTN), B11(GW), B12(NW), B13(CBM), B15(Customer), M16-M18(Invoice/Order), rows 21-31(Products), spare parts row, TOTAL row.

### 1. Create Output Folder
出货文件{BL No} in the order directory.

### 2. Generate 1. customer_sample.xlsx
Update: F6(BL), A16(customer), A22(vessel), C28(CTN), D28(goods), F28(GW), G28(CBM), A37(container), C37(seal), D37(type), E37(CTN), F37(GW), G37(CBM), J37(餐椅).

### 3. Generate 2. ICS2.xlsx
Always use assets/ICS2_Template.xlsx as base. Never modify Row 5 (56 sample cells).
Fill Row 6: E6(voyage), S6=1, U6(GW), buyer/shipper info.
Goods sheet: Row6 chair (CTN=total-spare, GW=total-2), Row7 spare parts (1 CTN, 2 GW, 0.014 CBM).

### 4. Generate 3. packing_list.xlsx
Update: A10(customer), D15(invoice), E15(date=first gen), A19(PCS), B19(CTNS), C19-E19(GW/NW/CBM).

### 5. Generate 4. invoice.xlsx
Filename: 4. 发票{InvoiceNo}.xlsx. Update: A8(customer), C12(invoice), F12(date), D23(PCS), E23(avg price), G23(amount), G25(USD{amount}).

### 6. Generate 5. customs_declaration.docx
Update table cells with newlines (NO | chars). Row1 Cell4: label only. Row1 Cell10: 提运单号\n{BL}. Other fields with proper newlines.

## Rules
1. Date = first generation date (箱单+发票 only).
2. ICS2 Row 5: never modify.
3. Chair CTN = total CTN - spare CTN.
4. 报关单: newlines, NO pipe chars.
5. 运输工具名称: label only.
6. Clean up temp files.
