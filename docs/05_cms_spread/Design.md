# CMS Spread Valuation Design

## Purpose

`05_cms_spread` contains CMS and CMS spread valuation components.

The current practical focus is swaption volatility handling. Full CMS spread valuation is a future extension.

## Responsibilities

- Hold ATM swaption volatility matrices.
- Interpolate expiry x tenor volatility points.
- Provide volatility inputs for CMS convexity adjustment.
- Later represent CMS and CMS spread trades.
- Later implement convexity adjustment and valuation workflows.

## Current Main Class

| Class | Role |
|---|---|
| `clsATMSwaptionVol` | ATM swaption volatility matrix and interpolation class |

## Future Candidate Classes

| Class | Role |
|---|---|
| `clsCMSSpreadTrade` | CMS spread trade terms |
| `clsCMSLeg` | CMS leg representation |
| `clsCMSSpreadLeg` | Spread leg representation |
| `clsCMSConvexityAdjustment` | Convexity adjustment logic |

## Dependency Rule

Allowed:

```text
05_cms_spread -> common
05_cms_spread -> 01_discount_curve
05_cms_spread -> 03_sabr
```

Possible later:

```text
05_cms_spread -> 02_ois_swap
```

Only add this dependency if swap construction or annuity calculation is explicitly reused.

## Notes

ATM swaption volatility handling is placed here for now because its current use is CMS-related. If it becomes a broadly reused market-data component, it may later be moved to a dedicated `market_data` or `volatility` module.
