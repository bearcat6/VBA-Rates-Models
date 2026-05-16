Option Explicit

'====================================================
' Module: MCapNormalFormula
' Purpose:
'   Bachelier / Normal model functions for caplet pricing
'====================================================

Private Const PI As Double = 3.14159265358979

Public Function NormalPDF(ByVal x As Double) As Double
    NormalPDF = Exp(-0.5 * x * x) / Sqr(2# * PI)
End Function

Public Function NormalCDF(ByVal x As Double) As Double
    ' Excelの標準関数を利用
    NormalCDF = Application.WorksheetFunction.Norm_S_Dist(x, True)
End Function

Public Function BachelierCallValue( _
    ByVal ForwardRate As Double, _
    ByVal strikeRate As Double, _
    ByVal normalVol As Double, _
    ByVal expiryYear As Double _
) As Double

    Dim stdDev As Double
    Dim d As Double

    If expiryYear <= 0# Or normalVol <= 0# Then
        BachelierCallValue = Application.WorksheetFunction.Max(ForwardRate - strikeRate, 0#)
        Exit Function
    End If

    stdDev = normalVol * Sqr(expiryYear)

    If stdDev <= 0# Then
        BachelierCallValue = Application.WorksheetFunction.Max(ForwardRate - strikeRate, 0#)
        Exit Function
    End If

    d = (ForwardRate - strikeRate) / stdDev

    BachelierCallValue = _
        (ForwardRate - strikeRate) * NormalCDF(d) _
        + stdDev * NormalPDF(d)

End Function

Public Function CapletPV_Normal( _
    ByVal ForwardRate As Double, _
    ByVal strikeRate As Double, _
    ByVal normalVol As Double, _
    ByVal expiryYear As Double, _
    ByVal accrualYear As Double, _
    ByVal discountFactor As Double, _
    ByVal Notional As Double _
) As Double

    Dim optionValue As Double

    optionValue = BachelierCallValue( _
                    ForwardRate, _
                    strikeRate, _
                    normalVol, _
                    expiryYear)

    CapletPV_Normal = Notional * accrualYear * discountFactor * optionValue

End Function

