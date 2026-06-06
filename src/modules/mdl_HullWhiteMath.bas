Option Explicit

' =============================================================================
' mdl_HullWhiteMath
' =============================================================================
'
' Hull-White 1F common mathematical functions.
'
' This module intentionally contains only model-level mathematical utilities.
' It does not read Excel sheets and does not depend on a concrete curve class.
'
' Curve-dependent functions receive the discount curve as Object and expect the
' curve to expose the time-based interface defined in docs:
'
'   DF_T(T)
'   InstantaneousForward(T)
'
' Model convention:
'
'   r(t) = phi(t) + x(t)
'   dx(t) = -a x(t) dt + sigma dW(t)
'
' where x(0) = 0 and phi(t) is chosen to reproduce the initial discount curve.
'
' In the first simple workflow, users may also inspect f(0,t) + x(t), but for
' arbitrage-consistent discount bond generation the phi adjustment should be used.
'
' =============================================================================

Private Const HW_EPS As Double = 0.000000000001
Private Const HW_ERR_BASE As Long = 7000

' =============================================================================
' Validation helpers
' =============================================================================

Public Sub HW_ValidateMeanReversion(ByVal in_a As Double, _
                                    Optional ByVal in_SourceName As String = "mdl_HullWhiteMath")

    If in_a < 0# Then
        Err.Raise vbObjectError + HW_ERR_BASE + 1, in_SourceName, _
                  "Mean reversion a must be non-negative."
    End If

End Sub

Public Sub HW_ValidateSigma(ByVal in_Sigma As Double, _
                            Optional ByVal in_SourceName As String = "mdl_HullWhiteMath")

    If in_Sigma < 0# Then
        Err.Raise vbObjectError + HW_ERR_BASE + 2, in_SourceName, _
                  "Sigma must be non-negative."
    End If

End Sub

Public Sub HW_ValidateTime(ByVal in_Time As Double, _
                           Optional ByVal in_SourceName As String = "mdl_HullWhiteMath")

    If in_Time < 0# Then
        Err.Raise vbObjectError + HW_ERR_BASE + 3, in_SourceName, _
                  "Time must be non-negative."
    End If

End Sub

Public Sub HW_ValidateTimeOrder(ByVal in_StartTime As Double, _
                                ByVal in_EndTime As Double, _
                                Optional ByVal in_SourceName As String = "mdl_HullWhiteMath")

    If in_StartTime < 0# Or in_EndTime < 0# Then
        Err.Raise vbObjectError + HW_ERR_BASE + 4, in_SourceName, _
                  "Times must be non-negative."
    End If

    If in_EndTime < in_StartTime Then
        Err.Raise vbObjectError + HW_ERR_BASE + 5, in_SourceName, _
                  "T must be greater than or equal to t."
    End If

End Sub

' =============================================================================
' Hull-White B function
' =============================================================================
'
' B(t,T) = (1 - exp(-a * (T - t))) / a
'
' If a is close to zero, B(t,T) tends to T - t.
'
' =============================================================================

Public Function HW_B(ByVal in_a As Double, _
                     ByVal in_StartTime As Double, _
                     ByVal in_EndTime As Double) As Double

    Dim tau As Double

    Call HW_ValidateMeanReversion(in_a, "mdl_HullWhiteMath.HW_B")
    Call HW_ValidateTimeOrder(in_StartTime, in_EndTime, "mdl_HullWhiteMath.HW_B")

    tau = in_EndTime - in_StartTime

    If Abs(in_a) < HW_EPS Then
        HW_B = tau
    Else
        HW_B = (1# - Exp(-in_a * tau)) / in_a
    End If

End Function

' =============================================================================
' OU process moments for x(t)
' =============================================================================
'
' dx(t) = -a x(t) dt + sigma dW(t)
'
' E[x(t+dt) | x(t)]
' Var[x(t+dt) | x(t)]
'
' =============================================================================

Public Function HW_XNextMean(ByVal in_x As Double, _
                             ByVal in_a As Double, _
                             ByVal in_dt As Double) As Double

    Call HW_ValidateMeanReversion(in_a, "mdl_HullWhiteMath.HW_XNextMean")
    Call HW_ValidateTime(in_dt, "mdl_HullWhiteMath.HW_XNextMean")

    If in_dt = 0# Then
        HW_XNextMean = in_x
    ElseIf Abs(in_a) < HW_EPS Then
        HW_XNextMean = in_x
    Else
        HW_XNextMean = in_x * Exp(-in_a * in_dt)
    End If

End Function

Public Function HW_XNextVariance(ByVal in_a As Double, _
                                 ByVal in_Sigma As Double, _
                                 ByVal in_dt As Double) As Double

    Call HW_ValidateMeanReversion(in_a, "mdl_HullWhiteMath.HW_XNextVariance")
    Call HW_ValidateSigma(in_Sigma, "mdl_HullWhiteMath.HW_XNextVariance")
    Call HW_ValidateTime(in_dt, "mdl_HullWhiteMath.HW_XNextVariance")

    If in_dt = 0# Or in_Sigma = 0# Then
        HW_XNextVariance = 0#
    ElseIf Abs(in_a) < HW_EPS Then
        HW_XNextVariance = in_Sigma * in_Sigma * in_dt
    Else
        HW_XNextVariance = in_Sigma * in_Sigma * (1# - Exp(-2# * in_a * in_dt)) / (2# * in_a)
    End If

End Function

Public Function HW_XNextStdDev(ByVal in_a As Double, _
                               ByVal in_Sigma As Double, _
                               ByVal in_dt As Double) As Double

    HW_XNextStdDev = Sqr(HW_XNextVariance(in_a, in_Sigma, in_dt))

End Function

Public Function HW_XNext(ByVal in_x As Double, _
                         ByVal in_a As Double, _
                         ByVal in_Sigma As Double, _
                         ByVal in_dt As Double, _
                         ByVal in_z As Double) As Double

    HW_XNext = HW_XNextMean(in_x, in_a, in_dt) + HW_XNextStdDev(in_a, in_Sigma, in_dt) * in_z

End Function

Public Function HW_XVarianceFromZero(ByVal in_a As Double, _
                                     ByVal in_Sigma As Double, _
                                     ByVal in_Time As Double) As Double

    HW_XVarianceFromZero = HW_XNextVariance(in_a, in_Sigma, in_Time)

End Function

Public Function HW_XStdDevFromZero(ByVal in_a As Double, _
                                   ByVal in_Sigma As Double, _
                                   ByVal in_Time As Double) As Double

    HW_XStdDevFromZero = Sqr(HW_XVarianceFromZero(in_a, in_Sigma, in_Time))

End Function

' =============================================================================
' Variance of the future integral of x(u)
' =============================================================================
'
' For conditional pricing at time t, x(t) is known.
'
'   q(t,T) = Var_t[ Integral_t^T x(u) du ]
'          = sigma^2 / a^2 * [ tau - 2 B(t,T)
'              + (1 - exp(-2 a tau)) / (2 a) ]
'
' where tau = T - t.
'
' If a is close to zero:
'   q(t,T) -> sigma^2 * tau^3 / 3
' =============================================================================

Public Function HW_IntegralXVariance(ByVal in_a As Double, _
                                     ByVal in_Sigma As Double, _
                                     ByVal in_StartTime As Double, _
                                     ByVal in_EndTime As Double) As Double

    Dim tau As Double
    Dim b As Double

    Call HW_ValidateMeanReversion(in_a, "mdl_HullWhiteMath.HW_IntegralXVariance")
    Call HW_ValidateSigma(in_Sigma, "mdl_HullWhiteMath.HW_IntegralXVariance")
    Call HW_ValidateTimeOrder(in_StartTime, in_EndTime, "mdl_HullWhiteMath.HW_IntegralXVariance")

    tau = in_EndTime - in_StartTime

    If tau = 0# Or in_Sigma = 0# Then
        HW_IntegralXVariance = 0#
    ElseIf Abs(in_a) < HW_EPS Then
        HW_IntegralXVariance = in_Sigma * in_Sigma * tau * tau * tau / 3#
    Else
        b = HW_B(in_a, in_StartTime, in_EndTime)
        HW_IntegralXVariance = in_Sigma * in_Sigma / (in_a * in_a) * _
            (tau - 2# * b + (1# - Exp(-2# * in_a * tau)) / (2# * in_a))
    End If

End Function

' =============================================================================
' Phi adjustment and short rate
' =============================================================================
'
' With r(t) = phi(t) + x(t), the deterministic shift is:
'
'   phi(t) = f(0,t) + sigma^2 / (2 a^2) * (1 - exp(-a t))^2
'
' for a > 0.
'
' If a is close to zero, the adjustment tends to 0.5 * sigma^2 * t^2.
' =============================================================================

Public Function HW_PhiAdjustment(ByVal in_a As Double, _
                                 ByVal in_Sigma As Double, _
                                 ByVal in_Time As Double) As Double

    Dim tmp As Double

    Call HW_ValidateMeanReversion(in_a, "mdl_HullWhiteMath.HW_PhiAdjustment")
    Call HW_ValidateSigma(in_Sigma, "mdl_HullWhiteMath.HW_PhiAdjustment")
    Call HW_ValidateTime(in_Time, "mdl_HullWhiteMath.HW_PhiAdjustment")

    If in_Sigma = 0# Or in_Time = 0# Then
        HW_PhiAdjustment = 0#
    ElseIf Abs(in_a) < HW_EPS Then
        HW_PhiAdjustment = 0.5 * in_Sigma * in_Sigma * in_Time * in_Time
    Else
        tmp = 1# - Exp(-in_a * in_Time)
        HW_PhiAdjustment = in_Sigma * in_Sigma * tmp * tmp / (2# * in_a * in_a)
    End If

End Function

Public Function HW_Phi(ByVal in_InitialInstantaneousForward As Double, _
                       ByVal in_a As Double, _
                       ByVal in_Sigma As Double, _
                       ByVal in_Time As Double) As Double

    HW_Phi = in_InitialInstantaneousForward + HW_PhiAdjustment(in_a, in_Sigma, in_Time)

End Function

Public Function HW_ShortRateFromX(ByVal in_InitialInstantaneousForward As Double, _
                                  ByVal in_x As Double, _
                                  ByVal in_a As Double, _
                                  ByVal in_Sigma As Double, _
                                  ByVal in_Time As Double) As Double

    HW_ShortRateFromX = HW_Phi(in_InitialInstantaneousForward, in_a, in_Sigma, in_Time) + in_x

End Function

' =============================================================================
' Discount bond from x(t) and the initial discount curve
' =============================================================================
'
' Under r(t) = phi(t) + x(t):
'
'   P(t,T) = P(0,T) / P(0,t)
'            * exp( -B(t,T) * x(t)
'                   + 0.5 * q(t,T)
'                   - 0.5 * { q(0,T) - q(0,t) } )
'
' where:
'   q(s,u) = Var_s[ Integral_s^u x(v) du ]
'
' This form is convenient because it uses only the initial discount curve and x(t).
'
' The curve object must expose:
'   DF_T(T)
' =============================================================================

Public Function HW_DiscountBondFromX(ByVal in_Curve As Object, _
                                     ByVal in_a As Double, _
                                     ByVal in_Sigma As Double, _
                                     ByVal in_StartTime As Double, _
                                     ByVal in_EndTime As Double, _
                                     ByVal in_x As Double) As Double

    Dim df0Start As Double
    Dim df0End As Double
    Dim b As Double
    Dim q_startEnd As Double
    Dim q_0End As Double
    Dim q_0Start As Double
    Dim exponentValue As Double

    If in_Curve Is Nothing Then
        Err.Raise vbObjectError + HW_ERR_BASE + 40, _
                  "mdl_HullWhiteMath.HW_DiscountBondFromX", _
                  "Curve object is Nothing."
    End If

    Call HW_ValidateTimeOrder(in_StartTime, in_EndTime, "mdl_HullWhiteMath.HW_DiscountBondFromX")

    df0Start = in_Curve.DF_T(in_StartTime)
    df0End = in_Curve.DF_T(in_EndTime)

    If df0Start <= 0# Or df0End <= 0# Then
        Err.Raise vbObjectError + HW_ERR_BASE + 41, _
                  "mdl_HullWhiteMath.HW_DiscountBondFromX", _
                  "Initial discount factors must be positive."
    End If

    If in_EndTime = in_StartTime Then
        HW_DiscountBondFromX = 1#
        Exit Function
    End If

    b = HW_B(in_a, in_StartTime, in_EndTime)
    q_startEnd = HW_IntegralXVariance(in_a, in_Sigma, in_StartTime, in_EndTime)
    q_0End = HW_IntegralXVariance(in_a, in_Sigma, 0#, in_EndTime)
    q_0Start = HW_IntegralXVariance(in_a, in_Sigma, 0#, in_StartTime)

    exponentValue = -b * in_x + 0.5 * q_startEnd - 0.5 * (q_0End - q_0Start)

    HW_DiscountBondFromX = (df0End / df0Start) * Exp(exponentValue)

End Function

Public Function HW_ZeroRateFromBond(ByVal in_BondPrice As Double, _
                                    ByVal in_StartTime As Double, _
                                    ByVal in_EndTime As Double) As Double

    Dim tau As Double

    Call HW_ValidateTimeOrder(in_StartTime, in_EndTime, "mdl_HullWhiteMath.HW_ZeroRateFromBond")

    tau = in_EndTime - in_StartTime

    If tau <= 0# Then
        HW_ZeroRateFromBond = 0#
        Exit Function
    End If

    If in_BondPrice <= 0# Then
        Err.Raise vbObjectError + HW_ERR_BASE + 42, _
                  "mdl_HullWhiteMath.HW_ZeroRateFromBond", _
                  "Bond price must be positive."
    End If

    HW_ZeroRateFromBond = -Log(in_BondPrice) / tau

End Function

Public Function HW_ForwardRateFromBonds(ByVal in_BondStart As Double, _
                                        ByVal in_BondEnd As Double, _
                                        ByVal in_StartTime As Double, _
                                        ByVal in_EndTime As Double) As Double

    If in_EndTime <= in_StartTime Then
        Err.Raise vbObjectError + HW_ERR_BASE + 43, _
                  "mdl_HullWhiteMath.HW_ForwardRateFromBonds", _
                  "endTime must be greater than startTime."
    End If

    If in_BondStart <= 0# Or in_BondEnd <= 0# Then
        Err.Raise vbObjectError + HW_ERR_BASE + 44, _
                  "mdl_HullWhiteMath.HW_ForwardRateFromBonds", _
                  "Bond prices must be positive."
    End If

    HW_ForwardRateFromBonds = (in_BondStart / in_BondEnd - 1#) / (in_EndTime - in_StartTime)

End Function

' =============================================================================
' Simple diagnostics helpers
' =============================================================================

Public Function HW_IsNearZero(ByVal in_x As Double, _
                              Optional ByVal in_Tolerance As Double = HW_EPS) As Boolean

    If in_Tolerance <= 0# Then
        Err.Raise vbObjectError + HW_ERR_BASE + 50, _
                  "mdl_HullWhiteMath.HW_IsNearZero", _
                  "Tolerance must be positive."
    End If

    HW_IsNearZero = (Abs(in_x) <= in_Tolerance)

End Function

Public Function HW_Max(ByVal in_x As Double, ByVal in_y As Double) As Double

    If in_x >= in_y Then
        HW_Max = in_x
    Else
        HW_Max = in_y
    End If

End Function

Public Function HW_Min(ByVal in_x As Double, ByVal in_y As Double) As Double

    If in_x <= in_y Then
        HW_Min = in_x
    Else
        HW_Min = in_y
    End If

End Function
