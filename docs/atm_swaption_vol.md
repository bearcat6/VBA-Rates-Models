# ATM Swaption Volatility Class

## Purpose

`clsATMSwaptionVol` stores an ATM swaption volatility matrix and returns interpolated ATM volatilities for arbitrary Expiry × Tenor points.

The previous class name `clsSwaptionVol` was too broad because the class handles ATM volatilities only. The new name makes the scope explicit and leaves room for future classes such as smile-aware swaption volatility surfaces.

## File

```text
src/classes/clsATMSwaptionVol.cls
```

The old file path:

```text
src/classes/clsSwaptionVol.cls
```

has been removed.

## Input matrix

The class expects an Excel range with tenors in the first row and expiries in the first column.

```text
       1Y    2Y    5Y    10Y
1M     vol   vol   vol   vol
3M     vol   vol   vol   vol
6M     vol   vol   vol   vol
1Y     vol   vol   vol   vol
```

## Interpolation policy

* Tenor direction: linear interpolation on variance, `Vol^2`.
* Expiry direction: linear interpolation on total variance, `Vol^2 × Expiry`.

## Main methods

```vb
Public Sub InitializeFromRange(ByVal in_MatrixRange As Range, _
                               Optional ByVal in_AllowFlatExtrapolation As Boolean = False)

Public Function Vol(ByVal in_ExpiryYears As Double, _
                    ByVal in_TenorText As String) As Double

Public Function VolByYears(ByVal in_ExpiryYears As Double, _
                           ByVal in_TenorYears As Double) As Double
```

## Naming rule

Use `clsATMSwaptionVol` for ATM volatility matrix access.

Do not use `clsSwaptionVol` for this class. A non-ATM or smile-aware swaption volatility class should be added separately with a more specific name.
