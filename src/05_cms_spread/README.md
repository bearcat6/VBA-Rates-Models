# 05_cms_spread

CMS and CMS spread valuation module.

## Current Focus

Currently this area mainly holds swaption volatility handling for future CMS-related valuation.

## Intended Contents

```text
classes/
  clsATMSwaptionVol.cls
  clsCMSSpreadTrade.cls
  clsCMSLeg.cls
  clsCMSSpreadLeg.cls
modules/
  mdl_CMSVolUtils.bas
  mdl_CMSSpreadPricer.bas
```

## Dependency Rule

This module may depend on:

```text
src/common
src/01_discount_curve
src/03_sabr
```

A dependency on `src/02_ois_swap` should be added only if swap construction or annuity logic is explicitly reused.
