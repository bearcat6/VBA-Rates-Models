# 03_sabr

Normal SABR smile, density and quantile module.

## Intended Contents

```text
classes/
  clsSABRNormal.cls
  clsSABRCalibrator.cls
  clsSmileSlice.cls
modules/
  mdl_SABRMath.bas
  mdl_SABRDensity.bas
```

## Initial Scope

- Normal SABR
- beta = 0
- alpha, rho, nu parameter validation
- smile fitting
- density and quantile calculation

## Dependency Rule

This module may depend on `src/common` only.
