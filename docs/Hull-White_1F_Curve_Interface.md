# Hull-White 1F Curve Interface

## 1. Purpose

This document defines the curve interface required by the Hull-White 1F model implementation.

The repository already contains date-based OIS curve classes such as:

* `clsOISStepForwardCurve`
* `clsOISZeroLinearCurve`

These classes mainly expose date-based methods such as:

```vb
Public Function DF(ByVal in_TargetDate As Date) As Double
Public Function ZeroRateCont(ByVal in_TargetDate As Date) As Double
Public Function ForwardRate(ByVal in_StartDate As Date, _
                            ByVal in_EndDate As Date) As Double
```

However, the Hull-White 1F model is naturally expressed on a time axis using year fractions such as `t` and `T`.

Therefore, each concrete curve class used by Hull-White should also expose a common time-based interface.

## 2. Background

The Hull-White 1F model is written as:

```text
dr(t) = { theta(t) - a r(t) } dt + sigma dW(t)
```

For implementation, this repository uses the shifted short-rate representation:

```text
r(t) = f(0,t) + x(t)

dx(t) = -a x(t) dt + sigma dW(t)
```

In this setup:

* the discount curve reproduces the current market curve;
* `a` controls mean reversion;
* `sigma` controls short-rate volatility;
* `theta(t)` is not stored directly as an input parameter;
* the model requires access to the initial discount curve and initial instantaneous forward curve.

Therefore, the Hull-White model should be able to call:

```vb
DF_T(T)
InstantaneousForward(T)
```

without knowing how the underlying curve was constructed.

## 3. Design Principle

The Hull-White implementation should not depend on a specific curve construction method.

For example, the model should not know whether the curve is:

* a stepwise instantaneous forward curve;
* a zero-rate linear interpolation curve;
* or another curve type added later.

Instead, concrete curve classes should provide the same public time-based methods.

The Hull-White classes can then receive the curve as `Object` and call same-name methods.

This is a practical design in VBA because VBA does not provide ordinary class inheritance in the same way as many object-oriented languages.

## 4. Required Time-Based Interface

Concrete curve classes used by Hull-White should implement the following public methods:

```vb
Public Function YearFracFromValDate(ByVal targetDate As Date) As Double

Public Function DateFromT(ByVal T As Double) As Date

Public Function DF_T(ByVal T As Double) As Double

Public Function ZeroRate_T(ByVal T As Double) As Double

Public Function ForwardRate_T(ByVal T1 As Double, _
                              ByVal T2 As Double) As Double

Public Function InstantaneousForward(ByVal T As Double, _
                                     Optional ByVal eps As Double = 0.0001) As Double
```

## 5. Method Details

### 5.1 YearFracFromValDate

```vb
Public Function YearFracFromValDate(ByVal targetDate As Date) As Double
```

Returns the year fraction from the valuation date to `targetDate`.

Current implementation policy:

```vb
YearFracFromValDate = YearFracAct365F(mValuationDate, targetDate)
```

Expected behavior:

* if `targetDate` is before the valuation date, raise an error;
* if `targetDate` equals the valuation date, return `0`;
* use the same day-count basis as the existing curve implementation.

### 5.2 DateFromT

```vb
Public Function DateFromT(ByVal T As Double) As Date
```

Converts a year-fraction time `T` into a date.

Temporary implementation policy:

```vb
DateFromT = DateAdd("d", CLng(Round(T * 365#, 0)), mValuationDate)
```

This is acceptable for the first Hull-White implementation step.

Future improvement:

* align this conversion with the repository's day-count convention;
* review consistency with `YearFracAct365F`;
* consider whether business-day adjustment should be applied depending on the intended use.

Expected behavior:

* if `T < 0`, raise an error;
* if `T = 0`, return the valuation date.

### 5.3 DF_T

```vb
Public Function DF_T(ByVal T As Double) As Double
```

Returns the discount factor `P(0,T)`.

Implementation policy:

```vb
DF_T = Me.DF(DateFromT(T))
```

or, for classes whose date-based method is named `df`:

```vb
DF_T = Me.df(DateFromT(T))
```

Expected behavior:

* if `T < 0`, raise an error;
* if `T = 0`, return `1`;
* otherwise, convert `T` to a date and call the existing date-based discount factor method.

### 5.4 ZeroRate_T

```vb
Public Function ZeroRate_T(ByVal T As Double) As Double
```

Returns the continuously compounded zero rate:

```text
z(0,T) = -ln(P(0,T)) / T
```

Implementation policy:

```vb
ZeroRate_T = ZeroRateFromDF(DF_T(T), T)
```

The common function `ZeroRateFromDF` is defined in `mdl_CurveMath.bas`.

Expected behavior:

* if `T = 0`, return `0`;
* if the discount factor is non-positive, raise an error through `mdl_CurveMath.ZeroRateFromDF`.

### 5.5 ForwardRate_T

```vb
Public Function ForwardRate_T(ByVal T1 As Double, _
                              ByVal T2 As Double) As Double
```

Returns the simple forward rate between `T1` and `T2`:

```text
F(0;T1,T2) = { P(0,T1) / P(0,T2) - 1 } / (T2 - T1)
```

Implementation policy:

```vb
ForwardRate_T = ForwardRateFromDFs(DF_T(T1), DF_T(T2), T1, T2)
```

The common function `ForwardRateFromDFs` is defined in `mdl_CurveMath.bas`.

Expected behavior:

* if `T2 <= T1`, raise an error;
* if either discount factor is non-positive, raise an error through `mdl_CurveMath.ForwardRateFromDFs`.

### 5.6 InstantaneousForward

```vb
Public Function InstantaneousForward(ByVal T As Double, _
                                     Optional ByVal eps As Double = 0.0001) As Double
```

Returns the initial instantaneous forward rate:

```text
f(0,T) = - d ln P(0,T) / dT
```

Implementation policy:

Use a finite-difference approximation based on `DF_T`.

For `T` close to zero, use a forward difference:

```text
f(0,T) ≈ - { ln P(0,T+eps) - ln P(0,0) } / (T+eps)
```

For normal positive `T`, use a central difference:

```text
f(0,T) ≈ - { ln P(0,T+eps) - ln P(0,T-eps) } / (2 eps)
```

Expected behavior:

* if `T < 0`, raise an error;
* if `eps <= 0`, raise an error;
* if either discount factor is non-positive, raise an error.

## 6. clsOISStepForwardCurve Implementation Policy

`clsOISStepForwardCurve` represents a curve where the instantaneous forward rate is stepwise constant by segment.

For this class:

```vb
DF_T(T)
```

should convert `T` into a date and call the existing:

```vb
df(Date)
```

method.

```vb
ZeroRate_T(T)
```

should call:

```vb
ZeroRateFromDF(DF_T(T), T)
```

```vb
ForwardRate_T(T1, T2)
```

should call:

```vb
ForwardRateFromDFs(DF_T(T1), DF_T(T2), T1, T2)
```

```vb
InstantaneousForward(T)
```

should be calculated by finite difference using `DF_T`.

Because this curve is based on stepwise forwards, `InstantaneousForward(T)` may show discontinuities around segment boundaries.

This is not an error. It reflects the curve construction method.

## 7. clsOISZeroLinearCurve Implementation Policy

`clsOISZeroLinearCurve` represents a curve where continuously compounded zero rates are linearly interpolated against year fraction.

For this class:

```vb
DF_T(T)
```

should convert `T` into a date and call the existing:

```vb
DF(Date)
```

method.

Other methods should follow the same interface as `clsOISStepForwardCurve`.

The shape of `InstantaneousForward(T)` will generally differ from `clsOISStepForwardCurve`.

This is natural because a zero-linear interpolation curve and a step-forward curve imply different instantaneous forward curves.

## 8. Relationship with clsDiscountCurve

`clsDiscountCurve` is currently used as a date-based interface through `Implements`.

It defines:

```vb
Public Function DF(ByVal in_TargetDate As Date) As Double

Public Function ZeroRateCont(ByVal in_TargetDate As Date) As Double

Public Function ForwardRate(ByVal in_StartDate As Date, _
                            ByVal in_EndDate As Date) As Double
```

The Hull-White time-based methods do not need to be added to `clsDiscountCurve` at this stage.

Reason:

* adding these methods to `clsDiscountCurve` would require all implementing classes to define corresponding private interface methods;
* the Hull-White model can instead receive the curve as `Object`;
* the required methods can be treated as a convention-based public interface.

Therefore, `clsDiscountCurve` remains the date-based interface.

The Hull-White curve interface is documented here as a convention-based interface.

## 9. Use from Hull-White Classes

Hull-White classes should receive a curve object such as:

```vb
Private mCurve As Object
```

and call:

```vb
mCurve.DF_T(T)
mCurve.InstantaneousForward(T)
```

This keeps the Hull-White implementation independent from the concrete curve construction method.

Example:

```vb
Public Sub Init(ByVal in_Curve As Object, _
                ByVal in_a As Double, _
                ByVal in_sigma As Double)

  Set mCurve = in_Curve
  mA = in_a
  mSigma = in_sigma

End Sub
```

## 10. Separation of Responsibilities

Hull-White model classes should not read Excel sheets directly.

Recommended responsibility split:

* curve classes
  provide discount factors, zero rates, forward rates, and instantaneous forwards;

* `mdl_CurveMath.bas`
  provides common curve-related mathematical functions;

* `clsHullWhite1F`
  holds Hull-White model parameters and model formula logic;

* `clsHWCalibrator`
  calibrates `a` and `sigma` from volatility data;

* `clsHWSimulator`
  generates Monte Carlo paths and future yield curves;

* `modExcelIO`
  reads inputs from Excel and writes outputs to Excel;

* `modHWWorkflow`
  orchestrates the workflow from Excel/VBA.

## 11. Current Status

Implemented or planned components:

```text
src/
  classes/
    clsDiscountCurve.cls
    clsOISStepForwardCurve.cls
    clsOISZeroLinearCurve.cls
    clsVolSurface.cls
    clsHullWhite1F.cls
    clsHWCalibrator.cls
    clsHWSimulator.cls
    clsRandomNormal.cls

  modules/
    mdl_CurveMath.bas
    modHWMath.bas
    modHWWorkflow.bas
    modExcelIO.bas
    modDiagnostics.bas
```

Current implementation step:

1. keep `mdl_CurveMath.bas` as the common utility module;
2. add the time-based interface to `clsOISStepForwardCurve`;
3. add the same time-based interface to `clsOISZeroLinearCurve`;
4. keep `clsDiscountCurve` as the existing date-based interface;
5. implement `clsHullWhite1F` using the time-based curve interface.