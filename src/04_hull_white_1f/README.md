# 04_hull_white_1f

Hull-White 1F model, calibration and Monte Carlo simulation.

## Intended Contents

```text
classes/
  clsHullWhite1F.cls
  clsHWCalibrator.cls
  clsHWSimulator.cls
  clsVolSurface.cls
modules/
  mdl_HullWhiteMath.bas
  mdl_HullWhiteWorkflow.bas
```

## Dependency Rule

This module may depend on:

```text
src/common
src/01_discount_curve
```

Model classes should not directly depend on Excel worksheet names or ranges.
