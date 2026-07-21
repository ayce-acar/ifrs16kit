# Vignette — Re-performing a Lessee's IFRS 16 Calculation with ifrs16kit (LeaseAuditKit)

This vignette walks the CALISMA golden case end-to-end through the tool, mapping each step to the requirement in IFRS 16 it implements and the auditing standard (ISA) it supports. It is written from the perspective of an audit senior re-performing management's lease estimate as a substantive procedure.

## The engagement scenario

Example Co. Limited (Ireland, EUR) leases a company car:

| Input | Value |
|---|---|
| Commencement date | 1 January 2025 |
| Lease term | 2 years (24 months) |
| Payment | €650 per month, **in advance** |
| Annual incremental borrowing rate (IBR) | 6.00% |
| Initial direct costs (IDC) | €600 |
| Prepayments / incentives / restoration | nil |
| Depreciation | Straight-line; useful life 60 months |

The client has recognised a lease liability and right-of-use (ROU) asset. Under ISA (Ireland) 540, the discounted liability is an accounting estimate; the auditor responds by **re-performing** the calculation independently — a form of audit evidence explicitly recognised in ISA (Ireland) 500 (re-performance, para A21). The tool documents that re-performance in a form retainable on file per ISA (Ireland) 230.

## Step 1 — Configure

```bash
ifrs16kit
```

(Programmatic alternative for a scripted audit file: construct
`ifrs16kit.LeaseInputs(...)` directly and call
`build_calculation_workbook()` — see the README's Python API section.
The vignette follows the interactive route an audit senior would use.)

Three questions set the engagement conventions:

1. **Country:** Ireland → the Setup lookup drives the currency (EUR €), reporting framework (IFRS, EU-adopted), lease standard (IFRS 16), audit standards (ISA (Ireland) 540/500/230), and an illustrative 12.5% tax rate. Choosing Türkiye, the UK, the UAE, or Australia swaps every one of these automatically (TFRS 16/BDS, UK-adopted IFRS, AASB 16/ASA, etc.).
2. **Timing:** Advance.
3. **Frequency:** Monthly (12/year).

## Step 2 — The input template

The tool writes `IFRS16_Input.xlsx`. Only dark-red cells are inputs; the sheet is already "liquid":

- **Total term (B11)** `=B9*12+B10` — enter 2 years + 0 months and 24 appears.
- **N (B14)** `=B11/B13` — 24 monthly periods.
- **Periodic rate (B21)** `=(1+B20)^(1/B12)-1` = 0.4868% — an *effective* monthly conversion of the 6% annual IBR, not 0.5% simple pro-rata. This is the first common client error the re-performance can surface.
- Two green checks guard the template: capacity (N ≤ 120) and whole-number-of-periods (a 25-month term at quarterly frequency would fail here, before the auditor wastes any time downstream).

The auditor fills entity, asset, date, term, payment, IBR, IDC, and useful life, saves, and returns to the script.

## Step 3 — Validation

`read_template()` re-reads the raw inputs and enforces the same rules in Python: valid date, positive amounts, IBR between 0 and 1, whole N, N ≤ 120. Validation failures are listed cell-by-cell (e.g. "IBR (B20) looks like a percentage — enter 0.06 or 6%") so the reviewer can trace every rejection to a specific cell — an ISA 230 documentation property.

## Step 4 — Initial measurement (IFRS 16 ¶¶22–28)

The **Initial Measurement** sheet discounts each payment at the periodic rate. Because payments are in advance, the discount exponent is `p − 1`, so Period 1's €650 is undiscounted:

| Period | Payment | Factor | PV |
|---|---|---|---|
| 1 | 650.00 | 1.000000 | 650.00 |
| 2 | 650.00 | 0.995156 | 646.85 |
| … | … | … | … |
| 24 | 650.00 | 0.894329 | 581.31 |
| **Total (¶26)** | 15,600.00 | | **14,761.09** |

The ROU asset build-up applies ¶24 line by line:

| | |
|---|---|
| Lease liability at commencement | 14,761.09 |
| Add: initial direct costs | 600.00 |
| Add: prepaid payments | 0.00 |
| Less: incentives received | (0.00) |
| Add: restoration provision | 0.00 |
| **ROU asset at commencement (¶24)** | **15,361.09** |

Every one of these is an Excel formula referencing the Inputs sheet — the auditor (or reviewer) can change IDC to €1,000 and watch the ROU asset move to 15,761.09 with no code involved.

## Step 5 — Subsequent measurement (¶36, ¶31)

The **Lease Schedule** sheet unwinds the liability using the effective interest method. For advance payments, interest accrues on the opening balance *net of* the period's payment:

`Interest_p = (Opening_p − Payment) × 0.48676%`

| Period | Opening | Interest | Principal | Closing | Depreciation |
|---|---|---|---|---|---|
| 1 | 14,761.09 | 68.69 | 581.31 | 14,179.78 | 640.05 |
| 2 | 14,179.78 | 65.86 | 584.14 | 13,595.63 | 640.05 |
| … | … | … | … | … | … |
| 24 | 650.00 | 0.00 | 650.00 | **0.00** | 640.05 |

Depreciation is straight-line over min(24, 60) = 24 periods: 15,361.09 ÷ 24 = 640.05 per period (¶31, no ownership transfer). The schedule is row-guarded to 120 periods — rows beyond N blank themselves out, so the same fabric serves a 6-month equipment lease or a 10-year property lease.

## Step 6 — The audit assertions, automated

The **Annual Summary** sheet aggregates by reporting year and runs five reconciliation checks that correspond to the arithmetic assertions of the re-performance:

| Check | Balance | Status |
|---|---|---|
| Closing lease liability at end of term | ~0.00 | OK |
| Closing ROU asset at end of term | ~0.00 | OK |
| Total principal repaid − initial liability | 0.00 | OK |
| Total expense − (total payments + IDC) | 0.00 | OK |
| Sum of front-loading differences | ~0.00 | OK |

The front-loading memo quantifies the IFRS 16 expense profile against a straight-line equivalent: 2025 carries €214.97 *more* expense than straight-line, 2026 exactly €214.97 less — the characteristic front-loading of interest-bearing lease accounting, useful when explaining the P&L effect to the client or in an audit committee summary.

Any failed check propagates automatically to the **Findings** sheet, which flips the corresponding row to Adjustment = "Y" and switches the recommendation from "No adjustment required." to an investigation prompt.

## Step 7 — Journals, tax, and the file

- **Journals** posts the commencement entry (Dr ROU 15,361.09; Cr liability 14,761.09; Cr cash — IDC 600.00) with a live debits-equals-credits check, plus representative Period 1 entries for the payment, interest unwind, and depreciation.
- **Tax_Reconciliation** contrasts the IFRS 16 P&L charge with a cash-rental deduction regime and computes the year-1 deferred tax at the Setup rate (12.5% for Ireland) — the temporary-difference mechanics an Irish or UK tax reviewer will ask about.
- **Audit_Procedures** (P-1 to P-20) and **PBC_List** (18 items) scope the wider IFRS 16 audit response beyond the single re-performed lease: completeness search via AP listings, IBR benchmarking, modifications, variable payments, expedients, FX, impairment, subleases, and sale-and-leaseback.

## Step 8 — The two-engine control

Before writing the workbook, the tool prints an independent Python re-computation of the same figures (available programmatically as `ifrs16kit.cross_check(inp)`, which is how the package's own verification tests assert the golden values):

```
Lease Liability (Day 1)  : EUR      14,761.09   (IFRS 16, para 26)
ROU Asset (Day 1)        : EUR      15,361.09   (IFRS 16, para 24)
Total Interest Expense   : EUR         838.91
Closing liability        : 0.000000   ✓ zero
Reconciliation: Liability + Interest = 15,600.00 = Total Payments   ✓ RECONCILED
```

The Excel model and the Python engine are two separate implementations of ¶¶24, 26, and 36. Agreement to the cent between them — and between both and the client's figure — is the re-performance evidence; disagreement with the client's figure is a finding, quantified and carried to the Findings sheet. This dual-implementation design addresses the reliability of the auditor's own tool (ISA 500's accuracy-and-completeness consideration for information produced by the auditor).

## Standards map

| Tool element | IFRS 16 | ISA (Ireland) |
|---|---|---|
| PV of payments, discount rate | ¶26–¶27 | 540 ¶¶22–23 (estimate re-performance) |
| ROU cost build-up | ¶24 | 540 |
| Effective interest unwinding | ¶36 | 500 A21 (re-performance) |
| Depreciation basis | ¶31–¶32 | 500 |
| Reconciliation checks & findings | — | 540 ¶23, 450 (misstatement evaluation) |
| Templates, validation trail, workbook | — | 230 (documentation) |
| Procedures P-1–P-20, PBC list | ¶9, ¶42, ¶47, ¶B58, ¶99–103 | 500, 330 |
