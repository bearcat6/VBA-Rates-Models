# Roadmap

This roadmap defines the intended order of development and migration.

## Phase 1: Repository Structure

- Prepare purpose-based documentation folders.
- Define module boundaries and dependency direction.
- Keep existing code stable while the target layout is documented.

## Phase 2: Discount Curve Foundation

Target module: `01_discount_curve`

- Stabilize `clsDiscountCurve` as the curve interface.
- Stabilize `clsOISStepForwardCurve`.
- Stabilize `clsOISZeroLinearCurve`.
- Keep curve construction independent from trade valuation.
- Confirm the time-based interface required by Hull-White:
  - `DF_T(T)`
  - `ForwardRate_T(T1, T2)`
  - `InstantaneousForward(T)`

## Phase 3: OIS Swap Valuation

Target module: `02_ois_swap`

- Keep OIS swap product terms and cashflow generation outside the curve classes.
- Implement and test fixed leg PV, floating leg PV, NPV and par rate.
- Confirm sign convention:
  - PAYER: floating leg PV - fixed leg PV
  - RECEIVER: fixed leg PV - floating leg PV

## Phase 4: SABR Model

Target module: `03_sabr`

- Implement Normal SABR with beta fixed at 0 as the initial production-style prototype.
- Add parameter validation for alpha, rho and nu.
- Add smile fitting from strike-volatility points.
- Add price curve smoothing if needed.
- Add density calculation and strike-to-percentile / percentile-to-strike conversion.

## Phase 5: Hull-White 1F

Target module: `04_hull_white_1f`

- Keep model logic independent from Excel worksheet I/O.
- Use the discount curve through a time-based interface.
- Initial calibration scope:
  - mean reversion `a` is externally supplied;
  - volatility `sigma` is fitted to ATM normal swaption volatility.
- Simulate future short-rate factors and future curves.
- Produce percentile curves and stress curves for risk-management use.

## Phase 6: CMS Spread Valuation

Target module: `05_cms_spread`

- Keep ATM swaption volatility handling here while CMS valuation is still under development.
- Later add:
  - CMS leg representation;
  - CMS spread leg representation;
  - convexity adjustment logic;
  - valuation workflow using discount curve and swaption volatility / SABR inputs.

## Commit Policy

Use small commits by purpose.

Recommended commit sequence:

1. Documentation structure.
2. Move common utilities.
3. Move discount curve files.
4. Move OIS swap files.
5. Move Hull-White files.
6. Add SABR files.
7. Add CMS spread files.
8. Refactor internal references after import testing.

Avoid moving and refactoring the same file in the same commit.
