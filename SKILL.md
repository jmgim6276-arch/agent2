---
name: caishui-agent2-import
description: Import Agent1-generated three-sheet Excel into Caishuitong with a fixed production sequence (Step1 + Step2 + Step2.5 + Step3). Use when users need one-click system import for newbies, including preflight checks, employee import, fee template/fee-role binding, workflow reuse, template creation, and fee-limit branching.
---

# Caishui Agent2 Import

Run this skill to execute system import directly from Agent1 output (`01_添加员工`, `02_费用科目配置`, `03_单据表`).

## Execute

1. Run preflight first:
```bash
python3 scripts/preflight_check.py
```

2. Import from Agent1 xlsx:
```bash
python3 scripts/import_from_agent1.py \
  --xlsx "/path/to/三表联动_客户模板_xxxx.xlsx" \
  --output "./agent2_import_report.json"
```

## Enforce Workflow Order

1. Import Sheet1 employees.
2. Process Sheet2 fee templates and fee-role relation.
3. Reuse/create workflow in Step2.5 and get `workflowId`.
4. Create Sheet3 templates with visibility mapping and fee-limit branch.

## Enforce Hard Rules

- Keep Sheet2 `归属单据名称` and Sheet3 `单据模板名称` one-to-one.
- Only allow Step3 doc categories: 报销单 / 借款单 / 批量付款单 / 申请单.
- Handle fee-limit logic only for 报销单 / 批量付款单.
- Skip fee-limit branch for 借款单 / 申请单 (do not mark as “unlimited”).
- Use fee-role branch only when Sheet2 row has both:
  - `归属单据名称` non-empty
  - `单据适配人员` non-empty
- Otherwise use direct leaf-fee restriction branch.
- Never create workflow inside Step3.

## Read References When Needed

- Newbie operation: `references/小白执行手册.md`
- Field and boundary rules: `references/字段与边界速查.md`
- Full battle playbook: `references/作战说明书-小白版.md`
- Troubleshooting quick card: `references/故障排查速查卡.md`
