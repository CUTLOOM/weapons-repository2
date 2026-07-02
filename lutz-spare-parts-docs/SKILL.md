---
name: lutz-spare-parts-docs
description: |
  Generate 补件唛头 (Spare Part Shipping Mark) and Spare Parts List
  Excel files from LUTZ Proforma Invoices. Use when creating
  foreign trade order documents: reading PI data, looking up
  spare parts from master reference, generating formatted
  补件唛头 with photos, generating aggregated SPL.
  Trigger: generate spare parts docs, make shipping marks,
  PI to documents, 补件唛头, Spare Parts List
---

# LUTZ Spare Parts Document Generator

Generate 补件唛头 and Spare Parts List from PI.

## Quick Start
Run: python scripts/generate_docs.py --pi path/to/PI.xlsx

Output: 补件唛头{ORDER}.xlsx in PI folder, SPL in 请款资料/

## Critical Rules
- Read ALL PI rows: for r in range(12, ws_pi.max_row + 1), check item_no and qty\n- NO blank rows between product groups (cr += len(parts), no +1)\n\n## 二、Type A: 请款资料工作流

### 2.1 读取装柜数据
Read the 装柜数据 xlsx with openpyxl. Sheets:
- `Drop Down` — Buyer info
- `Proforma Invoice` or `{Order}总数据` — Main order data
- `{ContainerN}` — 每个柜的明细（以柜号命名）

### 2.2 模板来源
模板文件：`8M73B` 订单的请款资料文件夹（已清理图片）
**中文路径问题**：先复制到 `$env:TEMP` 处理，再复制回目标文件夹

### 2.3 关键映射规则
- **DATE** = 装柜数据文件名中的日期
- **Buyer Address** = 标准地址（全大写，ß→SS，加国家，UID）
- **Route** = `FROM XINGANG TIANJIN CHINA TO {港口},{国家} BY SEA`
- **Discounts** = 默认 4.5%（0.045，`0.0%` 格式）

### 2.4 行数调整
模板有 3 个产品行，根据实际产品数增删行。操作前先取消所有合并单元格，调整后重新合并。

### 2.5 产品数据提取
- 产品行识别：Column S(Pos.No) **或** Column B(Article No)
- 毛重/净重/体积：以总数据 sheet 为准（`{Order}总数据`），不用柜子 sheet 的数据
- 装箱数公式：`=E{n}/{pu}`（实际装箱数），不硬编码 `/2`
- 所有数值：`str()` 自然精度，不强制小数位

### 2.6 图片规则
C 列 PICTURE 始终留空，`ws._images.clear()`

### 2.7 Declaration.docx
用 `.replace()` 保留标签文字，替换 BL/Container/Invoice/Order 字段

---
## 三、多柜/多订单处理

### 3.1 装柜数据文件结构
每个柜有一个独立 sheet（以柜号命名），每个订单有一个总数据 sheet。
柜数和订单数无固定对应关系：
- 1 个订单 + 多个柜
- 多个订单 + 1 个柜
- 多个订单 + 多个柜
具体情况具体分析。

### 3.2 PL 结构（固定）
```
R1-R6: 固定表头（永不删除）
R7+: 数据区
  每个柜:
    Container info → 产品行 → Spare part → Total(小计)
  最后:
    In total（各柜 Total 直接相加）
    Summary text（自然精度 str()）
```
**删除 PL 数据区：** `ws.delete_rows(7,12)`

### 3.3 CI 结构（固定）
```
R1-R8: 固定表头（永不删除）
R9+: 数据区
  每个柜:
    Container info → 产品行 → Spare part(FOC) → Total(小计 E/F/H)
  最后:
    discounts(4.5%, 0.0% 格式)
    In total(E=各柜Total E相加, F=各柜Total F相加, H=(各柜Total H相加)*0.955)
    Bank info
```
**删除 CI 数据区：** `ws.delete_rows(9,15)`（保留 R8 表头）

### 3.4 容器信息
- 柜号/封号：从柜名 sheet 的 R5(箱号) 和 R6(封号) 读取
- BL/ETD/ETA/港口：从柜名 sheet 读取（所有柜共用同一 BL）

### 3.5 PL Summary 文本格式
```
  TOTAL:   GROSS WEIGHT:{gw}KGS     ETD:{etd}
           NET WEIGHT:{nw}KGS       ETA:{eta}
           VLM:{cbm}CBM
           PCS:{pcs}PCS
           CTNS:{ctns}CTNS
```
- 字体：宋体 20, 对齐：left/top/wrap
- 合并范围：A{ssu}:E{ssu+2}（3行5列 A-E）
- 行高：汇总文字行 45.0

### 3.6 待办提醒
装柜数据中如有 `Credit note` / `罚款` 条目，必须先提醒用户。
discounts 默认 4.5%，除非用户告知变动。

---
## 四、参考资料
- scripts/generate_docs.py: Spare parts docs generation
- references/format-rules.md: 完整格式规格（含 PL/CI/Declaration）
- 8M73B 订单请款资料文件夹: 模板文件（已清理图片）