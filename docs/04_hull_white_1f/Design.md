# Hull-White 1F Design

## Purpose

`04_hull_white_1f` contains the 1-factor Hull-White model, calibration helpers and Monte Carlo simulation workflow.

The main business objective is to generate future JPY interest-rate curves from a discount curve and swaption volatility inputs, then produce percentile curves and stress curves.

## Model

```text
dr(t) = { theta(t) - a r(t) } dt + sigma dW(t)
```

Implementation uses the shifted short-rate representation:

```text
r(t) = phi(t) + x(t)
dx(t) = -a x(t) dt + sigma dW(t)
```

`theta(t)` is not treated as a direct input parameter. It is implied by the initial discount curve.

## Responsibilities

- Store Hull-White 1F parameters.
- Use discount curve input through a time-based interface.
- Calibrate `sigma` to volatility data under the initial simple scope.
- Simulate future short-rate factors.
- Generate future discount, zero-rate and forward-rate curves.
- Produce percentile and stress curve outputs.

## Main Classes

| Class | Role |
|---|---|
| `clsHullWhite1F` | Model parameter and analytical helper class |
| `clsHWCalibrator` | Calibration to volatility data |
| `clsHWSimulator` | Monte Carlo simulation |
| `clsVolSurface` | Generic volatility surface used as HW input |

## Main Modules

| Module | Role |
|---|---|
| `mdl_HullWhiteMath` | Model-level mathematical utilities |
| `mdl_HullWhiteWorkflow` | Thin Excel workflow/orchestration layer |

## Required Curve Interface

The curve object is expected to expose:

```vb
DF_T(T)
ForwardRate_T(T1, T2)
InstantaneousForward(T)
```

## Initial Calibration Scope

- Mean reversion `a` is externally supplied.
- Volatility `sigma` is fitted to ATM normal swaption volatility.
- Full swaption valuation and exact multi-point calibration are future extensions.

## Dependency Rule

Allowed:

```text
04_hull_white_1f -> 01_discount_curve
04_hull_white_1f -> common
```

Avoid:

```text
04_hull_white_1f -> 02_ois_swap
04_hull_white_1f -> 05_cms_spread
```

## Notes

Excel worksheet reading/writing should be kept in workflow modules. The model classes should not directly depend on sheet names or ranges.
