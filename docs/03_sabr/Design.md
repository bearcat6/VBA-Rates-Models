# SABR Design

## Purpose

`03_sabr` contains Normal SABR smile modelling components.

The initial target is the interest-rate option setting where market data are given as normal volatilities by expiry, tenor and strike/moneyness.

## Responsibilities

- Represent SABR parameters.
- Validate parameter ranges.
- Fit SABR parameters to observed normal volatilities.
- Convert normal volatility to option prices.
- Smooth the strike-price curve where necessary.
- Calculate implied density from option prices.
- Convert strike to percentile and percentile to strike.

## Initial Model Scope

- Normal SABR.
- `beta = 0` fixed.
- Parameters:
  - `alpha`
  - `rho`
  - `nu`

## Candidate Classes

| Class | Role |
|---|---|
| `clsSABRNormal` | Normal SABR model and volatility function |
| `clsSABRCalibrator` | Parameter fitting from strike-volatility points |
| `clsSmileSlice` | One expiry/tenor smile slice |

## Candidate Modules

| Module | Role |
|---|---|
| `mdl_SABRMath` | SABR formula and helper functions |
| `mdl_SABRDensity` | Density and quantile calculation |
| `mdl_Bachelier` | Normal option pricing; may live in `common` if used broadly |

## Dependency Rule

Allowed:

```text
03_sabr -> common
```

Avoid dependency on CMS, Hull-White or OIS swap valuation.

Not allowed:

```text
03_sabr -> 05_cms_spread
03_sabr -> 04_hull_white_1f
```

## Notes

SABR should not be placed under CMS because the same smile, density and quantile logic can be reused for stress testing and general risk analysis.
