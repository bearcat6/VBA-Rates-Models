# Hull-White 1F Curve Interface

## Purpose

The Hull-White 1F model is implemented on a year-fraction based time axis such as `t` and `T`.
Existing OIS curve classes mainly expose date-based methods such as `DF(Date)`, `ZeroRateCont(Date)`, and `ForwardRate(Date, Date)`.
For Hull-White calibration and simulation, each concrete discount curve class should also expose a common time-based interface.

This document defines the curve interface expected by the Hull-White 1F components.

## Design principle

The Hull-White model should not depend on the concrete curve construction method.
It should not know whether the source curve is built with stepwise instantaneous forwards, zero-rate linear interpolation, or another method.

The Hull-White model should call only the following methods:

```vb
Public Function DF_T(ByVal T As Double) As Double
Public Function ZeroRate_T(ByVal T As Double) As Double
Public Function ForwardRate_T(ByVal T1 As Double, ByVal T2 As Double) As Double
Public Function InstantaneousForward(ByVal T As Double, Optional ByVal eps As Double = 0.0001) As Double
Public Function YearFracFromValDate(ByVal targetDate As Date) As Double
```

## Required methods

### DF_T(T)

Returns the discount factor `P(0,T)` from the valuation date to time `T` years.

Implementation policy:

- Convert `T` to a target date by `DateFromT(T)`.
- Call the existing date-based `DF(Date)` method.
- `T = 0` should return `1`.
- Negative `T` should raise an error.

### ZeroRate_T(T)

Returns the continuously compounded zero rate:

```text
z(0,T) = -ln(P(0,T)) / T
```

Implementation policy:

- Use `mdl_CurveMath.ZeroRateFromDF`.
- `T = 0` should return `0`.
- Non-positive discount factors should raise an error in the common utility function.

### ForwardRate_T(T1, T2)

Returns the simple forward rate between `T1` and `T2`:

```text
F(0;T1,T2) = {P(0,T1) / P(0,T2) - 1} / (T2 - T1)
```

Implementation policy:

- Use `mdl_CurveMath.ForwardRateFromDFs`.
- `T2 <= T1` should raise an error.
- Non-positive discount factors should raise an error in the common utility function.

### InstantaneousForward(T)

Returns the initial instantaneous forward rate:

```text
f(0,T) = - d ln P(0,T) / dT
```

Implementation policy:

- Use a finite-difference approximation based on `DF_T`.
- The default `eps` is `0.0001` years.
- Around `T = 0`, use a forward difference to avoid negative time.
- For `T > eps`, use a central difference.

The value may differ between `clsOISStepForwardCurve` and `clsOISZeroLinearCurve`.
This is natural because the interpolation method is different.
For example, a step-forward curve may produce piecewise-flat instantaneous forwards, while a zero-linear curve produces forwards implied by the slope of the zero-rate curve.

### YearFracFromValDate(date)

Returns the ACT/365F year fraction from the valuation date to the target date.

Implementation policy:

- Use the same day count basis as the existing curve implementation, currently `YearFracAct365F`.
- Dates before the valuation date should raise an error.

## VBA implementation note

VBA does not provide standard class inheritance in the same way as many object-oriented languages.
Although `clsDiscountCurve` is used with `Implements` for the date-based interface, the Hull-White time-based methods can be treated as a convention-based interface.

That means `clsHullWhite1F`, `clsHWCalibrator`, and `clsHWSimulator` may receive the curve as `Object` and call the required same-name methods.
Concrete curve classes such as `clsOISStepForwardCurve` and `clsOISZeroLinearCurve` must therefore implement the same public method names.

## Concrete curve classes

The following classes should expose the Hull-White time-based interface:

- `clsOISStepForwardCurve`
- `clsOISZeroLinearCurve`

`clsDiscountCurve` may keep only the existing date-based `Implements` interface.
If desired, comments can be added to `clsDiscountCurve` explaining that Hull-White uses an additional convention-based time interface.

## Temporary DateFromT policy

`DateFromT` may initially convert time to date using:

```vb
DateFromT = DateAdd("d", CLng(Round(T * 365#, 0)), mValuationDate)
```

This is acceptable for the first Hull-White implementation step.
Later, it should be aligned with the repository's day-count, calendar, and business-day conventions.

## Separation of responsibilities

Hull-White model classes should not read Excel sheets directly.
Excel input/output should remain in `modExcelIO`.

Recommended split:

- Curve classes: provide discount factors, zero rates, forwards, and instantaneous forwards.
- `clsHullWhite1F`: hold model parameters and model formula logic.
- `clsHWCalibrator`: calibrate `a` and `sigma` from volatility data.
- `clsHWSimulator`: generate Monte Carlo paths and future curves.
- `modExcelIO`: read inputs and write outputs.
- `modHWWorkflow`: orchestrate high-level workflow from Excel/VBA.
