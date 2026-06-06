# 02_ois_swap

OIS swap product representation, cashflow generation and valuation.

## Intended Contents

```text
classes/
  clsOISSwap.cls
  clsOISLeg.cls
  clsCashflow.cls
modules/
  mdl_OISSwapPricer.bas
```

## Dependency Rule

This module may depend on:

```text
src/common
src/01_discount_curve
```

It should not depend on SABR, Hull-White or CMS spread modules.
