Option Explicit

' =============================================================================
' Curve Math Utilities
' =============================================================================
'
' Common curve-related mathematical functions.
'
' These functions are intentionally independent from specific curve classes such
' as clsOISStepForwardCurve or clsOISZeroLinearCurve.
'
' Concrete curve classes should call these functions to avoid duplicating
' zero-rate and forward-rate calculation logic.
'
' =============================================================================

Public Function ZeroRateFromDF(ByVal df As Double, ByVal T As Double) As Double

    If T <= 0# Then
        ZeroRateFromDF = 0#
        Exit Function
    End If

    If df <= 0# Then
        Err.Raise vbObjectError + 5501, "mdl_CurveMath.ZeroRateFromDF", _
                  "Discount factor must be positive."
    End If

    ZeroRateFromDF = -Log(df) / T

End Function


Public Function ForwardRateFromDFs(ByVal df1 As Double, ByVal df2 As Double, _
                                   ByVal T1 As Double, ByVal T2 As Double) As Double

    If T2 <= T1 Then
        Err.Raise vbObjectError + 5502, "mdl_CurveMath.ForwardRateFromDFs", _
                  "T2 must be greater than T1."
    End If

    If df1 <= 0# Or df2 <= 0# Then
        Err.Raise vbObjectError + 5503, "mdl_CurveMath.ForwardRateFromDFs", _
                  "Discount factors must be positive."
    End If

    ForwardRateFromDFs = (df1 / df2 - 1#) / (T2 - T1)

End Function
