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

Public Function ZeroRateFromDF(ByVal in_df As Double, ByVal in_T As Double) As Double

    If in_T <= 0# Then
        ZeroRateFromDF = 0#
        Exit Function
    End If

    If in_df <= 0# Then
        Err.Raise vbObjectError + 5501, "mdl_CurveMath.ZeroRateFromDF", _
                  "Discount factor must be positive."
    End If

    ZeroRateFromDF = -Log(in_df) / in_T

End Function


Public Function ForwardRateFromDFs(ByVal in_df1 As Double, ByVal in_df2 As Double, _
                                    ByVal in_T1 As Double, ByVal in_T2 As Double) As Double

    If in_T2 <= in_T1 Then
        Err.Raise vbObjectError + 5502, "mdl_CurveMath.ForwardRateFromDFs", _
                  "T2 must be greater than T1."
    End If

    If in_df1 <= 0# Or in_df2 <= 0# Then
        Err.Raise vbObjectError + 5503, "mdl_CurveMath.ForwardRateFromDFs", _
                  "Discount factors must be positive."
    End If

    ForwardRateFromDFs = (in_df1 / in_df2 - 1#) / (in_T2 - in_T1)

End Function