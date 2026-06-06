Option Explicit

'====================================================
' Module: MCapNormalFormula
' Purpose:
'   Bachelier / Normal model functions for caplet pricing
'====================================================

Private Const PI As Double = 3.14159265358979

Public Function NormalPDF(ByVal in_X As Double) As Double
    NormalPDF = Exp(-0.5 * in_X * in_X) / Sqr(2# * PI)
End Function

Public Function NormalCDF(ByVal in_X As Double) As Double
    ' Excelの標準関数を利用
    NormalCDF = Application.WorksheetFunction.Norm_S_Dist(in_X, True)
End Function

Public Function BachelierCallValue( _
    ByVal in_ForwardRate As Double, _
    ByVal in_StrikeRate As Double, _
    ByVal in_NormalVol As Double, _
    ByVal in_ExpiryYear As Double _
) As Double

    Dim stdDev As Double
    Dim d As Double

    If in_ExpiryYear <= 0# Or in_NormalVol <= 0# Then
        BachelierCallValue = Application.WorksheetFunction.Max(in_ForwardRate - in_StrikeRate, 0#)
        Exit Function
    End If

    stdDev = in_NormalVol * Sqr(in_ExpiryYear)

    If stdDev <= 0# Then
        BachelierCallValue = Application.WorksheetFunction.Max(in_ForwardRate - in_StrikeRate, 0#)
        Exit Function
    End If

    d = (in_ForwardRate - in_StrikeRate) / stdDev

    BachelierCallValue = _
        (in_ForwardRate - in_StrikeRate) * NormalCDF(d) _
        + stdDev * NormalPDF(d)

End Function

Public Function CapletPV_Normal( _
    ByVal in_ForwardRate As Double, _
    ByVal in_StrikeRate As Double, _
    ByVal in_NormalVol As Double, _
    ByVal in_ExpiryYear As Double, _
    ByVal in_AccrualYear As Double, _
    ByVal in_DiscountFactor As Double, _
    ByVal in_Notional As Double _
) As Double

    Dim optionValue As Double

    optionValue = BachelierCallValue( _
                    in_ForwardRate, _
                    in_StrikeRate, _
                    in_NormalVol, _
                    in_ExpiryYear)

    CapletPV_Normal = in_Notional * in_AccrualYear * in_DiscountFactor * optionValue

End Function