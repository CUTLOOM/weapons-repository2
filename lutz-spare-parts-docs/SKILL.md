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
- Read ALL PI rows: for r in range(12, ws_pi.max_row + 1), check item_no and qty\n- NO blank rows between product groups (cr += len(parts), no +1)\n\n## Resources
- scripts/generate_docs.py: Main generation script
- references/format-rules.md: Formatting specifications
