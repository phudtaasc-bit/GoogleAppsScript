# Audit baseline — Financial model logic

Date: 2026-07-10
Branch: `audit-financial-logic-20260710`

## Scope

Audit and patch the existing Apps Script source without rebuilding the model, deleting input data, or changing the user interface unless an entry-point conflict prevents the current code from running.

## Data-preservation invariants

1. Do not modify or clear `01. Đầu vào`.
2. Do not rename sheets, blocks, or source columns during logic patches.
3. Preserve current public function names unless a duplicate global definition is the defect being fixed.
4. Keep backward-compatible aliases when standardizing terminology.
5. Patch one logical defect group per commit.
6. Preserve the current final-period VAT refund treatment: negative VAT payable represents the refund of the remaining deductible VAT balance.

## Canonical terminology

- Tiền SDĐ
- Tiền thuê đất
- Chi phí xây dựng & thiết bị
- Chi phí GPMB
- Chi phí HTKT
- Chi phí dự phòng
- Chi phí lãi vay
- Chi phí bán hàng
- Chi phí vận hành
- Vốn chủ sở hữu
- Thuế GTGT
- Thuế TNDN
- FCFF
- FCFE
- NPV
- IRR

Legacy labels remain accepted as input aliases to avoid data loss.

## Confirmed model decisions

### VAT refund

- The final-period negative VAT payable is intentional.
- It represents the cash refund of the remaining deductible VAT balance.
- The patch must preserve this treatment in Sheet 03, Sheet 04, summary reports, and sensitivity calculations.
- Validation only reports an error when the refund amount does not equal the remaining deductible VAT balance.

## High-priority findings

### P0 — execution and global-function conflicts

- Multiple global definitions exist for `onOpen`, `FS05B_DoNhay_CSH`, `FS_chayMoHinh_Buoc1`, and `FS_chayMoHinh_Buoc2`.
- The `Menu.js` implementation of `FS05B_DoNhay_CSH` calls an undefined function and overrides the valid batch implementation.
- Existing full-model runners use inconsistent dependency orders although Sheet 02 and Sheet 03 depend on each other through allocated interest.

### P0 — VAT

- The final-period VAT refund treatment is approved and must be retained.
- The VAT checker previously rejected negative VAT payable, contradicting the approved generation logic; the independent audit checker now accepts and validates it.
- VAT input on selling costs is omitted in Sheet 03.
- Direct construction VAT is initialized at zero even though the common-cost block supports a VAT rate.
- The sensitivity engine must reproduce the same final-period VAT refund treatment as the base model.

### P0 — land costs

- Separate land-cost schedules can each apply the full aggregate land-cost pool because all specific names are matched to one merged `Tiền SDĐ/thuê đất` category.
- This creates a material double-counting risk when Liền kề, Chung cư, TMDV, and Chợ have separate schedule rows.
- Product classification depends on free-text product names; canonical aliases are required without changing existing data.

### P1 — interest, FCFF and FCFE

- Interest is calculated from opening debt, then allocated back to products through a fixed two-pass rebuild; no convergence control exists.
- FCFF is the pre-financing operating cash flow and FCFE is shareholder distributions less equity contributions, but checkers use stale column mappings.
- The Sheet 04 checker reads only 32 of 39 columns and misidentifies financing columns.

### P1 — NPV and IRR

- FCFF and FCFE are discounted using the same input rate.
- The model does not distinguish project discount rate/WACC from cost of equity.
- Monthly IRR annualization is consistent, but the sensitivity engines must use the same timing convention as the base model.

### P1 — Thuế TNDN

- Taxable profit is floored at zero by month/product with no loss carryforward balance.
- The sensitivity engine uses a simplified revenue-weighted tax allocation that does not reproduce the base tax logic.

### P1 — selling costs

- Selling costs are calculated only for sale products, which is appropriate, but VAT treatment differs between the base model and the sensitivity engine.

## Planned commit sequence

1. Audit baseline and invariants.
2. Remove broken duplicate entry points while preserving the visible menu.
3. Add non-destructive validation and record approved VAT refund treatment.
4. Repair selling-cost VAT and construction VAT while preserving final-period refund logic.
5. Separate land-cost pools and schedules by canonical category.
6. Stabilize interest iteration and financing checks.
7. Align FCFF/FCFE, NPV/IRR rates, timing, and validation.
8. Align Thuế TNDN and sensitivity calculations with the base model.
9. Add regression checks and final audit report.
