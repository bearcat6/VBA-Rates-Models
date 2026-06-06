# Discount Curve Design

## Purpose

`01_discount_curve` contains the basic JPY TONA/OIS discount curve components.

This module is the foundation for OIS swap valuation, Hull-White calibration and simulation, and future CMS spread valuation.

## Responsibilities

- Define the discount curve interface.
- Build JPY OIS discount curves.
- Return discount factors, zero rates and forward rates.
- Provide a time-based interface for interest-rate models such as Hull-White 1F.

## Main Classes

| Class | Role |
|---|---|
| `clsDiscountCurve` | Interface-style class for discount curves |
| `clsOISStepForwardCurve` | OIS curve based on stepwise instantaneous forward rates |
| `clsOISZeroLinearCurve` | OIS curve based on linear interpolation of continuously compounded zero rates |

## Main Modules

| Module | Role |
|---|---|
| `mdl_CurveMath` | Curve-related mathematical utilities such as zero-rate and forward-rate conversion |

## Interface

Date-based interface:

```vb
DF(in_TargetDate)
ZeroRateCont(in_TargetDate)
ForwardRate(in_StartDate, in_EndDate)
```

Time-based interface for model use:

```vb
DF_T(T)
ZeroRate_T(T)
ForwardRate_T(T1, T2)
InstantaneousForward(T)
```

## Dependency Rule

This module may depend on `common`, but must not depend on product valuation modules or model modules.

Allowed:

```text
01_discount_curve -> common
```

Not allowed:

```text
01_discount_curve -> 02_ois_swap
01_discount_curve -> 04_hull_white_1f
01_discount_curve -> 05_cms_spread
```

## Notes

OIS swap par-rate helper functions may exist temporarily in the curve class during prototyping, but the long-term direction is to keep trade valuation logic in `02_ois_swap`.
