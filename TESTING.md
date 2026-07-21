# TESTING — the ifrs16kit automated verification protocol

The suite verifies the package in six layers, from pure arithmetic up to a
fully automated version of the two-engine control described in VIGNETTE.md.

| Layer | File | What it proves |
|---|---|---|
| 1. Engine | `tests/test_engine.py` | The CALISMA golden benchmark (liability 14,761.09 · ROU 15,361.09 · interest 838.91); the periodic rate is an effective conversion; the liability matches an independent closed-form annuity ("third engine"); and across the full grid of frequency × timing × method × term × useful life, the accounting identities hold: closing liability = 0, liability + interest = payments, ROU = liability + IDC ± ¶24 items, total depreciation = ROU, closing ROU = 0, total P&L = payments + IDC. Includes the fractional useful-life regression. |
| 2. Template | `tests/test_template.py` | Required cells are generated blank (the cannot-run-on-example-values control); derived cells are live formulas including `B33 = MIN(B14,INT(B32))`; the four dropdowns exist; a user-style fill round-trips through `read_template` to the golden figures; and every validation rule rejects cell-by-cell (empty cells, IBR as whole percentage, non-whole N, N > 120, bad country/frequency/method/timing/date, non-positive amounts). |
| 3. Workbook | `tests/test_workbook.py` | Exactly ten sheets in documented order; the four core calculation sheets contain **zero** hard-coded numbers (the "liquid" claim); key outputs are cross-sheet links; the schedule is 120-row-capacity and row-guarded; the depreciation formula divides by the floored period count; Setup mirrors `COUNTRIES` exactly; Journals and Findings are wired across sheets; P-1…P-20 and the 18 PBC items are present; the workbook builds for all five jurisdictions. |
| 4. CLI & flow | `tests/test_cli.py` | `--demo` writes only `_DEMO` files and never touches real workbooks (regression); `python -m ifrs16kit` works; the full scripted interview → fill → validate → confirm → build flow completes; an invalid fill is rejected and the retry loop recovers; an existing filled template is detected and reused. |
| 5. API | `tests/test_api.py` | Every `__all__` export resolves; `__version__` equals the installed package metadata (single-sourcing); `create_input` / `create_calculation` work and reject bad arguments. |
| 6. Two-engine | `tests/test_excel_recalc.py` | LibreOffice recalculates the generated workbooks head-lessly and Excel's computed figures must equal `cross_check` within half a cent, for four scenarios (golden, arrears-quarterly, annual-UK-reducing-balance, fractional-life regression). All five Annual Summary reconciliation checks must read "OK". Skips when LibreOffice is absent; CI installs it. |

## Running

```bash
pip install -e .[dev]
pytest                                  # layers 1–5 (~15 s); layer 6 auto-skips
pytest --cov=ifrs16kit --cov-report=term-missing
pytest tests/test_excel_recalc.py      # layer 6, needs LibreOffice (soffice)
```

If `soffice` is not on PATH, point `IFRS16KIT_SOFFICE` at a LibreOffice
binary (or a `.py` wrapper) before running layer 6.

## Continuous integration

`ifrs16kit/.github/workflows/ci.yml` runs three jobs on every push and pull request:
layers 1–5 on Python 3.9 / 3.11 / 3.13 with a 90% coverage gate, layer 6 on
Ubuntu with `libreoffice-calc` installed, and a packaging job (`python -m
build` + `twine check`). A release is publishable only when all three are
green.

## Extending

New calculation features need, at minimum: a golden value or identity in
layer 1, a formula/structure assertion in layer 3, and a recalc scenario in
layer 6 — so every figure stays proven twice, in Python and in Excel.
