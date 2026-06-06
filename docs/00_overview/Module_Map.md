# Module Map

This document defines the final purpose-based module structure for `VBA-Rates-Models`.

The repository is moving from a flat `src/classes` and `src/modules` layout to a structure organized by analytical purpose.

## Final Module Structure

```text
src/
├─ common/
│  ├─ classes/
│  └─ modules/
├─ 01_discount_curve/
│  ├─ classes/
│  └─ modules/
├─ 02_ois_swap/
│  ├─ classes/
│  └─ modules/
├─ 03_sabr/
│  ├─ classes/
│  └─ modules/
├─ 04_hull_white_1f/
│  ├─ classes/
│  └─ modules/
└─ 05_cms_spread/
   ├─ classes/
   └─ modules/
```

## Responsibility of Each Module

| Module | Responsibility |
|---|---|
| `common` | Date utilities, day count, business day adjustment, array utilities, numerical methods, random numbers, normal distribution and Bachelier utilities |
| `01_discount_curve` | JPY TONA/OIS discount curve construction and curve interface |
| `02_ois_swap` | OIS swap product terms, cashflow generation, PV, NPV and par rate |
| `03_sabr` | Normal SABR parameters, smile fitting, price curve smoothing, density calculation and strike/quantile conversion |
| `04_hull_white_1f` | Hull-White 1F model, calibration to volatility, Monte Carlo simulation and future curve generation |
| `05_cms_spread` | CMS and CMS spread valuation components; currently mainly swaption volatility handling |

## Dependency Direction

Dependencies should point only from downstream modules to upstream modules.

```text
common
  ↓
01_discount_curve
  ↓
02_ois_swap

common + 01_discount_curve
  ↓
04_hull_white_1f

common + 01_discount_curve + 03_sabr
  ↓
05_cms_spread
```

Avoid reverse dependencies. For example, `01_discount_curve` must not call `04_hull_white_1f` or `05_cms_spread`.

## Migration Map from Current Flat Structure

| Current Path | Target Path | Notes |
|---|---|---|
| `src/classes/clsDiscountCurve.cls` | `src/01_discount_curve/classes/clsDiscountCurve.cls` | Curve interface |
| `src/classes/clsOISStepForwardCurve.cls` | `src/01_discount_curve/classes/clsOISStepForwardCurve.cls` | Step forward OIS curve |
| `src/classes/clsOISZeroLinearCurve.cls` | `src/01_discount_curve/classes/clsOISZeroLinearCurve.cls` | Zero rate linear curve |
| `src/modules/mdl_CurveMath.bas` | `src/01_discount_curve/modules/mdl_CurveMath.bas` | Curve-specific math |
| `src/classes/clsRandomNormal.cls` | `src/common/classes/clsRandomNormal.cls` | Common Monte Carlo utility |
| `src/classes/clsVolSurface.cls` | `src/04_hull_white_1f/classes/clsVolSurface.cls` | Generic surface currently used for HW input |
| `src/classes/clsHullWhite1F.cls` | `src/04_hull_white_1f/classes/clsHullWhite1F.cls` | HW model |
| `src/classes/clsHWCalibrator.cls` | `src/04_hull_white_1f/classes/clsHWCalibrator.cls` | HW calibration |
| `src/classes/clsHWSimulator.cls` | `src/04_hull_white_1f/classes/clsHWSimulator.cls` | HW simulation |
| `src/modules/mdl_HullWhiteMath.bas` | `src/04_hull_white_1f/modules/mdl_HullWhiteMath.bas` | HW math utilities |
| `src/modules/mdl_HullWhiteWorkFlow.bas` | `src/04_hull_white_1f/modules/mdl_HullWhiteWorkflow.bas` | HW Excel workflow/orchestration |
| `src/classes/clsATMSwaptionVol.cls` | `src/05_cms_spread/classes/clsATMSwaptionVol.cls` | ATM swaption vol for CMS-related valuation |

## Naming Notes

- Class modules should start with `cls`.
- Standard modules should start with `mdl_`.
- Function arguments should use the `in_` prefix for input parameters.
- Object variables should use a `c` prefix where practical, for example `in_cCurve`.

## Migration Rule

When moving VBA files, preserve file contents exactly first. Refactoring and renaming inside the files should be done in a separate commit after confirming that the moved files can be imported into Excel/VBA.
