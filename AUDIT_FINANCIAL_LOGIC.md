# Audit status — Financial model logic

Date: 2026-07-13
Branch: `audit-financial-logic-20260710`
Status: **Sprint 1 source logic complete; live Apps Script runtime validation pending**

## Scope

Audit and patch the existing Apps Script source without rebuilding the model, deleting input data, or changing the user interface unless an entry-point conflict or a validation requirement makes a change necessary.

## Data-preservation invariants

1. Do not modify or clear `01. Đầu vào`.
2. Do not rename sheets, blocks, or source columns during logic patches.
3. Preserve public function names unless a duplicate global definition is the defect being fixed.
4. Keep input aliases where needed to avoid data loss.
5. Patch one logical defect group per commit.
6. Preserve the final-period VAT refund treatment: negative VAT payable represents the refund of the remaining deductible VAT balance.
7. Thuế TNDN is calculated separately by output product; taxable profit and Thuế TNDN are never negative and no loss offset is applied between products or months.

## Confirmed model decisions

### VAT refund

- Final-period negative VAT payable is intentional.
- It represents the cash refund of the remaining deductible VAT balance.
- Validation reports an error only when the refund occurs before the final period or does not reconcile with the remaining deductible VAT balance.

### Thuế TNDN

- Thuế TNDN is determined for each product output.
- Taxable profit is `MAX(0, revenue - deductible product cost)`.
- No negative Thuế TNDN is generated.
- No loss carryforward or cross-product loss offset is applied in this model.
- Sheet 03 and Sheet 04 receive the monthly aggregate from Sheet 02.

### Financing waterfall

- Cash retained from the previous month is used before new capital is raised.
- When cash is insufficient, each funding event follows the input ratio between debt and equity (**Method A**).
- Interest is calculated on opening debt and capitalized into debt.
- Surplus operating cash is used to repay debt.
- Cash remaining after debt repayment is retained for later months.
- Distribution to shareholders occurs only at the final model period.
- Closing debt is:

  `Opening debt + Loan drawdown + Capitalized interest - Principal repayment`

- Sheet 04 is the single calculation source for financing; Sheet 03 receives the synchronized financing block `U:AE`.

### FCFF / FCFE / valuation

- FCFF is the project cash flow before financing.
- FCFE represents cash flow available to equity after debt effects and equity funding, regardless of whether cash is physically distributed in that month.
- NPV project uses FCFF and WACC.
- NPV equity uses FCFE and the cost of equity input (`Tỷ suất chiết khấu`).
- Monthly discounting uses the effective monthly equivalent of the annual rate.
- Monthly IRR is annualized on an effective basis.

## Completed Sprint 1 patches

### Execution and entry points

- Removed broken duplicate global entry points.
- Standardized the model runner order.
- Sheet 00 valuation logic now resides in the core `00.Tonghop.js`; temporary valuation patch and compatibility alias were removed.

### Sheet 03 cost and financing separation

- Sheet 03 initially calculates only cost, VAT, and pre-financing cash flow.
- The initial financing block `U:AE` is zeroed before the first financing run.
- Sheet 04 calculates financing and synchronizes the complete state back to Sheet 03.
- Removed duplicate finance-sync and cost-only entry files.

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

- Replaced the fixed rebuild with convergence iteration.
- Convergence tracks loan drawdown, interest, principal repayment, closing debt, and retained closing cash by month.
- Maximum iterations: 50.
- The model stops before summary generation if convergence is not reached.
- Added an independent interest audit and complete Sheet 03–04 financing reconciliation.

### Thuế TNDN

- Added product-level formula validation.
- Added monthly reconciliation between Sheet 02, Sheet 03, and Sheet 04.
- Added a blocking guard before summary generation.

### FCFF / FCFE / NPV / IRR

- Corrected stale Sheet 04 column mappings.
- Rebuilt the cash waterfall using retained cash, capitalized interest, debt repayment, and final-period distribution.
- Validated FCFF against pre-financing cash flow.
- Validated FCFE against the approved equity cash-flow definition.
- Integrated project NPV/WACC and equity NPV/cost-of-equity logic into the core Sheet 00 builder.

### Sensitivity

- Added NPV/IRR equity sensitivity to loan interest rate.
- Added NPV/IRR equity sensitivity to investment cost.
- Both groups read the same scenario levels used by the existing project sensitivity table.
- Removed the superseded fixed-level sensitivity entry point.
- Input values, formulas, and the complete baseline model are restored even when a scenario fails.

### Regression validation

- Added end-to-end checks for Sheet 00, 02, 03, and 04.
- Added checks for VAT, Thuế TNDN, financing, retained cash, FCFF, FCFE, NPV, and IRR formulas.
- Added full `U:AE` Sheet 03–04 reconciliation, including cumulative values.
- The full model runner blocks completion when regression validation fails.
- Added menu command `7. Kiểm thử hồi quy toàn mô hình` for manual validation.

## Sprint 1 completion boundary

The source-level financial logic audit and patch set is complete. The branch is **not yet approved for merge** because the connected GitHub environment cannot execute the container-bound Apps Script against the live spreadsheet.

## Required live validation before merge

1. Synchronize this branch to the Apps Script project attached to the live Google Sheet.
2. Run `4.3. Chạy toàn bộ mô hình - có cập nhật lãi vay`.
3. Run `7. Kiểm thử hồi quy toàn mô hình`.
4. Run existing project sensitivity tables.
5. Run `6.1 Độ nhạy vốn CSH - Lãi suất & Vốn đầu tư`.
6. Confirm that all formulas and input values return to the baseline after the sensitivity run.
7. Review final debt, retained cash, total interest, FCFF, FCFE, NPV, and IRR.
8. Merge to `main` only after all blocking checks pass.

## Sprint 2 scope after runtime approval

- Consolidate the stable Sheet 03 patch into the core Sheet 03 file.
- Rename helper files to final production names.
- Remove remaining compatibility-only code and dead functions.
- Optimize spreadsheet reads/writes without changing financial results.
- Prepare the final merge and release notes.
