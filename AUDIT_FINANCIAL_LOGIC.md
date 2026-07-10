# Audit status — Financial model logic

Date: 2026-07-10
Branch: `audit-financial-logic-20260710`

## Scope

Audit and patch the existing Apps Script source without rebuilding the model, deleting input data, or changing the user interface unless an entry-point conflict or a validation requirement makes a change necessary.

## Data-preservation invariants

1. Do not modify or clear `01. Đầu vào`.
2. Do not rename sheets, blocks, or source columns during logic patches.
3. Preserve current public function names unless a duplicate global definition is the defect being fixed.
4. Keep backward-compatible aliases when standardizing terminology.
5. Patch one logical defect group per commit.
6. Preserve the current final-period VAT refund treatment: negative VAT payable represents the refund of the remaining deductible VAT balance.
7. Thuế TNDN is calculated separately by output product; taxable profit and Thuế TNDN are never negative and no loss offset is applied between products or months.

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
- Validation reports an error only when the refund occurs before the final period or the amount does not reconcile with the remaining deductible VAT balance.

### Thuế TNDN

- Thuế TNDN is determined for each product output.
- Taxable profit is `MAX(0, revenue - deductible product cost)`.
- No negative Thuế TNDN is generated.
- No loss carryforward or cross-product loss offset is applied in this model.
- Sheet 03 and Sheet 04 receive the monthly aggregate from Sheet 02.

### Discount rates

- FCFF / NPV project uses WACC.
- FCFE / NPV equity uses the cost of equity input (`Tỷ suất chiết khấu`).
- Monthly discounting uses the effective monthly equivalent of the annual rate.
- Monthly IRR is annualized on an effective basis.

## Completed patches

### Execution and entry points

- Removed broken duplicate global entry points.
- Preserved backward-compatible aliases.
- Standardized the model runner order.

### VAT

- Preserved final-period VAT refund treatment.
- Added input VAT on construction & equipment costs.
- Added input VAT on selling costs.
- Added cross-sheet VAT validation.

### Tiền SDĐ / Tiền thuê đất

- Separated pools for Liền kề, Chung cư, TMDV, and Chợ.
- Added fallback compatibility for the legacy aggregate land-cost row.
- Added a blocking guard for duplicate or over-allocated land-cost configurations.

### Interest and financing

- Replaced the fixed two-pass rebuild with convergence iteration.
- Convergence tracks loan drawdown, interest, principal repayment, and closing debt by month.
- The model stops before summary generation if convergence is not reached.

### Thuế TNDN

- Added product-level formula validation.
- Added monthly reconciliation between Sheet 02, Sheet 03, and Sheet 04.
- Added a blocking guard before summary generation.

### FCFF / FCFE / NPV / IRR

- Corrected stale Sheet 04 column mappings.
- Validated FCFF against pre-financing cash flow.
- Validated FCFE against shareholder cash movements used by the current model.
- Changed NPV project to WACC and NPV equity to cost of equity.
- Preserved effective annualization of monthly IRR.

### Sensitivity

- Added NPV/IRR equity sensitivity to loan interest rate.
- Added NPV/IRR equity sensitivity to investment cost.
- Both new groups read the same scenario levels used by the existing project sensitivity table.
- Input values and formulas are restored in `finally` blocks after each sensitivity run.

### Regression validation

- Added end-to-end checks for Sheet 00, 02, 03, and 04.
- Added checks for VAT, Thuế TNDN, financing, FCFF, FCFE, NPV, and IRR formulas.
- The full model runner blocks completion when regression validation fails.
- Added menu command `7. Kiểm thử hồi quy toàn mô hình` for manual validation.

## Remaining before merge

1. Run the full model and the regression suite on the live Google Sheet project.
2. Run both existing and new sensitivity modules on the live model.
3. Confirm no formula or input restoration errors occur after sensitivity runs.
4. Consolidate stable patch logic into core files where this can be done without introducing replacement risk.
5. Remove superseded patch files only after the live-model tests pass.
6. Merge to `main` only after all blocking checks pass.
