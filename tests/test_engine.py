"""Layer 1 — the Python cross-check engine (IFRS 16 ¶¶24, 26, 31, 36).

Verifies the golden benchmark, checks the engine against an independent
closed-form annuity implementation (a *third* engine), and asserts the
accounting identities that must hold for every input combination.
"""
import pytest

import ifrs16kit as kit
from conftest import GOLDEN, MONEY


# ── Golden benchmark ────────────────────────────────────────────────────

def test_golden_case_matches_documented_values():
    cc = kit.cross_check(kit.LeaseInputs())
    assert cc["liability"] == pytest.approx(GOLDEN["liability"], **MONEY)
    assert cc["rou"] == pytest.approx(GOLDEN["rou"], **MONEY)
    assert cc["total_interest"] == pytest.approx(GOLDEN["total_interest"], **MONEY)
    assert cc["total_depreciation"] == pytest.approx(GOLDEN["total_depreciation"], **MONEY)
    assert cc["total_payments"] == pytest.approx(GOLDEN["total_payments"], **MONEY)
    assert cc["n"] == GOLDEN["n"]
    assert cc["closing_liability"] == pytest.approx(0, abs=1e-6)
    assert cc["closing_rou"] == pytest.approx(0, abs=1e-6)


def test_periodic_rate_is_effective_not_pro_rata():
    inp = kit.LeaseInputs()
    assert inp.periodic_rate == pytest.approx((1.06) ** (1 / 12) - 1, rel=1e-12)
    assert inp.periodic_rate != pytest.approx(0.06 / 12, abs=1e-6)


# ── Third engine: closed-form annuity PV ────────────────────────────────

@pytest.mark.parametrize("freq", [12, 4, 2, 1])
@pytest.mark.parametrize("advance", [True, False])
@pytest.mark.parametrize("ibr", [0.06, 0.12, 0.0001])
def test_liability_matches_closed_form_annuity(freq, advance, ibr):
    inp = kit.LeaseInputs(freq=freq, is_advance=advance, ibr_annual=ibr,
                          term_years=2, useful_life_months=60)
    r, n = inp.periodic_rate, int(inp.num_periods)
    pv = inp.payment * (1 - (1 + r) ** -n) / r          # ordinary annuity
    if advance:
        pv *= (1 + r)                                    # annuity-due
    assert kit.cross_check(inp)["liability"] == pytest.approx(pv, rel=1e-9)


def test_advance_liability_exceeds_arrears():
    adv = kit.cross_check(kit.LeaseInputs(is_advance=True))["liability"]
    arr = kit.cross_check(kit.LeaseInputs(is_advance=False))["liability"]
    assert adv > arr


# ── Accounting identities across the whole input grid ───────────────────

@pytest.mark.parametrize("freq", [12, 4, 2, 1])
@pytest.mark.parametrize("advance", [True, False])
@pytest.mark.parametrize("method", ["Straight-line", "Reducing balance"])
@pytest.mark.parametrize("term_years", [2, 5])
@pytest.mark.parametrize("life", [60, 18])
def test_engine_invariants(freq, advance, method, term_years, life):
    inp = kit.LeaseInputs(freq=freq, is_advance=advance, method=method,
                          term_years=term_years, useful_life_months=life)
    cc = kit.cross_check(inp)

    # Liability fully unwinds: closing balance is nil (¶36).
    assert cc["closing_liability"] == pytest.approx(0, abs=1e-6)
    # Principal + interest identity: liability + interest = cash paid.
    assert cc["liability"] + cc["total_interest"] == pytest.approx(
        cc["total_payments"], abs=1e-6)
    # ROU build-up (¶24) with the golden default components.
    assert cc["rou"] == pytest.approx(cc["liability"] + inp.idc, abs=1e-6)
    # The asset fully depreciates within min(term, life) for BOTH methods
    # (straight-line exhausts it; reducing balance writes off the final
    # period) — so closing ROU is nil and total depreciation equals cost.
    assert cc["total_depreciation"] == pytest.approx(cc["rou"], abs=1e-6)
    assert cc["closing_rou"] == pytest.approx(0, abs=1e-6)
    # Total P&L over the term = cash payments + IDC (¶¶31, 36 combined).
    assert cc["total_interest"] + cc["total_depreciation"] == pytest.approx(
        cc["total_payments"] + inp.idc, abs=1e-6)


def test_rou_buildup_uses_all_para24_components():
    inp = kit.LeaseInputs(idc=1000, prepaid=200, incentives=150, restoration=75)
    cc = kit.cross_check(inp)
    assert cc["rou"] == pytest.approx(
        cc["liability"] + 1000 + 200 - 150 + 75, abs=1e-9)


# ── Regression: fractional useful-life-in-periods (life-limited lease) ──
# Excel previously divided by an *unrounded* MIN(B14, B32) while Python
# floored — the engines diverged and the asset never fully depreciated.
# Both engines must now use the floored period count.

@pytest.mark.parametrize("method", ["Straight-line", "Reducing balance"])
def test_fractional_life_periods_fully_depreciate(method):
    inp = kit.LeaseInputs(freq=4, term_years=5, useful_life_months=50,
                          method=method)          # 50/3 = 16.67 → 16 periods
    cc = kit.cross_check(inp)
    assert cc["total_depreciation"] == pytest.approx(cc["rou"], abs=1e-6)
    assert cc["closing_rou"] == pytest.approx(0, abs=1e-6)


def test_depreciation_stops_at_useful_life_before_term_end():
    # 12-month life inside a 24-month lease: dep runs 12 periods only.
    inp = kit.LeaseInputs(useful_life_months=12)
    cc = kit.cross_check(inp)
    per_period = cc["rou"] / 12
    assert per_period * 12 == pytest.approx(cc["total_depreciation"], abs=1e-9)


# ── Data-model properties ───────────────────────────────────────────────

def test_term_and_period_properties():
    inp = kit.LeaseInputs(term_years=3, term_extra_months=6, freq=2)
    assert inp.term_months == 42
    assert inp.months_per_period == 6
    assert inp.num_periods == 7
    assert kit.LeaseInputs(is_advance=False).timing_flag == 0
    assert kit.LeaseInputs(is_advance=True).timing_flag == 1
