# 01_discount_curve

JPY TONA/OIS discount curve construction and curve interface.

## Intended Contents

```text
classes/
  clsDiscountCurve.cls
  clsOISStepForwardCurve.cls
  clsOISZeroLinearCurve.cls
modules/
  mdl_CurveMath.bas
```

## Dependency Rule

This module may use `src/common`, but should not depend on OIS swap valuation, SABR, Hull-White, or CMS spread modules.
