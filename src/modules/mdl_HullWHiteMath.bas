Option Explicit

' =============================================================================
' mdl_HullWHiteMath
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

Public Sub HW_ValidateMeanReversion(ByVal a As Double, _
                                    Optional ByVal sourceName As String = "mdl_HullWHiteMath")

    If a < 0# Then
        Err.Raise vbObjectError + HW_ERR_BASE + 1, sourceName, _
                  "Mean reversion a must be non-negative."
    End If

End Sub

Public Sub HW_ValidateSigma(ByVal sigma As Double, _
                            Optional ByVal sourceName As String = "mdl_HullWHiteMath")

    If sigma < 0# Then
        Err.Raise vbObjectError + HW_ERR_BASE + 2, sourceName, _
                  "Sigma must be non-negative."
    End If

End Sub

Public Sub HW_ValidateTime(ByVal T As Double, _
                           Optional ByVal sourceName As String = "mdl_HullWHiteMath")

    If T < 0# Then
        Err.Raise vbObjectError + HW_ERR_BASE + 3, sourceName, _
                  "Time must be non-negative."
    End If

End Sub

Public Sub HW_ValidateTimeOrder(ByVal t As Double, ByVal T As Double, _
                                Optional ByVal sourceName As String = "mdl_HullWHiteMath")

    If t < 0# Or T < 0# Then
        Err.Raise vbObjectError + HW_ERR_BASE + 4, sourceName, _
                  "Times must be non-negative."
    End If

    If T < t Then
        Err.Raise vbObjectError + HW_ERR_BASE + 5, sourceName, _
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

Public Function HW_B(ByVal a As Double, _
                     ByVal t As Double, _
                     ByVal T As Double) As Double

    Dim tau As Double

    Call HW_ValidateMeanReversion(a, "mdl_HullWHiteMath.HW_B")
    Call HW_ValidateTimeOrder(t, T, "mdl_HullWHiteMath.HW_B")

    tau = T - t

    If Abs(a) < HW_EPS Then
        HW_B = tau
    Else
        HW_B = (1# - Exp(-a * tau)) / a
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

Public Function HW_XNextMean(ByVal x As Double, _
                             ByVal a As Double, _
                             ByVal dt As Double) As Double

    Call HW_ValidateMeanReversion(a, "mdl_HullWHiteMath.HW_XNextMean")
    Call HW_ValidateTime(dt, "mdl_HullWHiteMath.HW_XNextMean")

    If dt = 0# Then
        HW_XNextMean = x
    ElseIf Abs(a) < HW_EPS Then
        HW_XNextMean = x
    Else
        HW_XNextMean = x * Exp(-a * dt)
    End If

End Function

Public Function HW_XNextVariance(ByVal a As Double, _
                                 ByVal sigma As Double, _
                                 ByVal dt As Double) As Double

    Call HW_ValidateMeanReversion(a, "mdl_HullWHiteMath.HW_XNextVariance")
    Call HW_ValidateSigma(sigma, "mdl_HullWHiteMath.HW_XNextVariance")
    Call HW_ValidateTime(dt, "mdl_HullWHiteMath.HW_XNextVariance")

    If dt = 0# Or sigma = 0# Then
        HW_XNextVariance = 0#
    ElseIf Abs(a) < HW_EPS Then
        HW_XNextVariance = sigma * sigma * dt
    Else
        HW_XNextVariance = sigma * sigma * (1# - Exp(-2# * a * dt)) / (2# * a)
    End If

End Function

Public Function HW_XNextStdDev(ByVal a As Double, _
                               ByVal sigma As Double, _
                               ByVal dt As Double) As Double

    HW_XNextStdDev = Sqr(HW_XNextVariance(a, sigma, dt))

End Function

Public Function HW_XNext(ByVal x As Double, _
                         ByVal a As Double, _
                         ByVal sigma As Double, _
                         ByVal dt As Double, _
                         ByVal z As Double) As Double

    HW_XNext = HW_XNextMean(x, a, dt) + HW_XNextStdDev(a, sigma, dt) * z

End Function

Public Function HW_XVarianceFromZero(ByVal a As Double, _
                                     ByVal sigma As Double, _
                                     ByVal t As Double) As Double

    HW_XVarianceFromZero = HW_XNextVariance(a, sigma, t)

End Function

Public Function HW_XStdDevFromZero(ByVal a As Double, _
                                   ByVal sigma As Double, _
                                   ByVal t As Double) As Double

    HW_XStdDevFromZero = Sqr(HW_XVarianceFromZero(a, sigma, t))

End Function

' =============================================================================
' Variance of the future integral of x(u)
' =============================================================================
'
' For conditional pricing at time t, x(t) is known.
'
'
'   q(t,T) = Var_t[ Integral_t^T x(u) du ]
'
'          = sigma^2 / a^2 * [ tau - 2 B(t,T)
'              + (1 - exp(-2 a tau)) / (2 a) ]
'
' where tau = T - t.
'
' If a is close to zero:
'
'   q(t,T) -> sigma^2 * tau^3 / 3
'
' =============================================================================

Public Function HW_IntegralXVariance(ByVal a As Double, _
                                     ByVal sigma As Double, _
                                     ByVal t As Double, _
                                     ByVal T As Double) As Double

    Dim tau As Double
    Dim b As Double

    Call HW_ValidateMeanReversion(a, "mdl_HullWHiteMath.HW_IntegralXVariance")
    Call HW_ValidateSigma(sigma, "mdl_HullWHiteMath.HW_IntegralXVariance")
    Call HW_ValidateTimeOrder(t, T, "mdl_HullWHiteMath.HW_IntegralXVariance")

    tau = T - t

    If tau = 0# Or sigma = 0# Then
        HW_IntegralXVariance = 0#
    ElseIf Abs(a) < HW_EPS Then
        HW_IntegralXVariance = sigma * sigma * tau * tau * tau / 3#
    Else
        b = HW_B(a, t, T)
        HW_IntegralXVariance = sigma * sigma / (a * a) * _
            (tau - 2# * b + (1# - Exp(-2# * a * tau)) / (2# * a))
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
'
' =============================================================================

Public Function HW_PhiAdjustment(ByVal a As Double, _
                                 ByVal sigma As Double, _
                                 ByVal t As Double) As Double

    Dim tmp As Double

    Call HW_ValidateMeanReversion(a, "mdl_HullWHiteMath.HW_PhiAdjustment")
    Call HW_ValidateSigma(sigma, "mdl_HullWHiteMath.HW_PhiAdjustment")
    Call HW_ValidateTime(t, "mdl_HullWHiteMath.HW_PhiAdjustment")

    If sigma = 0# Or t = 0# Then
        HW_PhiAdjustment = 0#
    ElseIf Abs(a) < HW_EPS Then
        HW_PhiAdjustment = 0.5 * sigma * sigma * t * t
    Else
        tmp = 1# - Exp(-a * t)
        HW_PhiAdjustment = sigma * sigma * tmp * tmp / (2# * a * a)
    End If

End Function

Public Function HW_Phi(ByVal initialInstantaneousForward As Double, _
                       ByVal a As Double, _
                       ByVal sigma As Double, _
                       ByVal t As Double) As Double

    HW_Phi = initialInstantaneousForward + HW_PhiAdjustment(a, sigma, t)

End Function

Public Function HW_ShortRateFromX(ByVal initialInstantaneousForward As Double, _
                                  ByVal x As Double, _
                                  ByVal a As Double, _
                                  ByVal sigma As Double, _
                                  ByVal t As Double) As Double

    HW_ShortRateFromX = HW_Phi(initialInstantaneousForward, a, sigma, t) + x

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
'   q(s,u) = Var_s[ Integral_s^u x(v) dv ]
'
' This form is convenient because it uses only the initial discount curve and x(t).
'
' The curve object must expose:
'   DF_T(T)
'
' =============================================================================

Public Function HW_DiscountBondFromX(ByVal curve As Object, _
                                     ByVal a As Double, _
                                     ByVal sigma As Double, _
                                     ByVal t As Double, _
                                     ByVal T As Double, _
                                     ByVal x As Double) As Double

    Dim df0t As Double
    Dim df0T As Double
    Dim b As Double
    Dim q_tT As Double
    Dim q_0T As Double
    Dim q_0t As Double
    Dim exponentValue As Double

    If curve Is Nothing Then
        Err.Raise vbObjectError + HW_ERR_BASE + 40, _
                  "mdl_HullWHiteMath.HW_DiscountBondFromX", _
                  "Curve object is Nothing."
    End If

    Call HW_ValidateTimeOrder(t, T, "mdl_HullWHiteMath.HW_DiscountBondFromX")

    df0t = curve.DF_T(t)
    df0T = curve.DF_T(T)

    If df0t <= 0# Or df0T <= 0# Then
        Err.Raise vbObjectError + HW_ERR_BASE + 41, _
                  "mdl_HullWHiteMath.HW_DiscountBondFromX", _
                  "Initial discount factors must be positive."
    End If

    If T = t Then
        HW_DiscountBondFromX = 1#
        Exit Function
    End If

    b = HW_B(a, t, T)
    q_tT = HW_IntegralXVariance(a, sigma, t, T)
    q_0T = HW_IntegralXVariance(a, sigma, 0#, T)
    q_0t = HW_IntegralXVariance(a, sigma, 0#, t)

    exponentValue = -b * x + 0.5 * q_tT - 0.5 * (q_0T - q_0t)

    HW_DiscountBondFromX = (df0T / df0t) * Exp(exponentValue)

End Function

Public Function HW_ZeroRateFromBond(ByVal bondPrice As Double, _
                                    ByVal t As Double, _
                                    ByVal T As Double) As Double

    Dim tau As Double

    Call HW_ValidateTimeOrder(t, T, "mdl_HullWHiteMath.HW_ZeroRateFromBond")

    tau = T - t

    If tau <= 0# Then
        HW_ZeroRateFromBond = 0#
        Exit Function
    End If

    If bondPrice <= 0# Then
        Err.Raise vbObjectError + HW_ERR_BASE + 42, _
                  "mdl_HullWHiteMath.HW_ZeroRateFromBond", _
                  "Bond price must be positive."
    End If

    HW_ZeroRateFromBond = -Log(bondPrice) / tau

End Function

Public Function HW_ForwardRateFromBonds(ByVal bondStart As Double, _
                                        ByVal bondEnd As Double, _
                                        ByVal startTime As Double, _
                                        ByVal endTime As Double) As Double

    If endTime <= startTime Then
        Err.Raise vbObjectError + HW_ERR_BASE + 43, _
                  "mdl_HullWHiteMath.HW_ForwardRateFromBonds", _
                  "endTime must be greater than startTime."
    End If

    If bondStart <= 0# Or bondEnd <= 0# Then
        Err.Raise vbObjectError + HW_ERR_BASE + 44, _
                  "mdl_HullWHiteMath.HW_ForwardRateFromBonds", _
                  "Bond prices must be positive."
    End If

    HW_ForwardRateFromBonds = (bondStart / bondEnd - 1#) / (endTime - startTime)

End Function

' =============================================================================
' Simple diagnostics helpers
' =============================================================================

Public Function HW_IsNearZero(ByVal x As Double, _
                              Optional ByVal tolerance As Double = HW_EPS) As Boolean

    If tolerance <= 0# Then
        Err.Raise vbObjectError + HW_ERR_BASE + 50, _
                  "mdl_HullWHiteMath.HW_IsNearZero", _
                  "Tolerance must be positive."
    End If

    HW_IsNearZero = (Abs(x) <= tolerance)

End Function

Public Function HW_Max(ByVal x As Double, ByVal y As Double) As Double

    If x >= y Then
        HW_Max = x
    Else
        HW_Max = y
    End If

End Function

Public Function HW_Min(ByVal x As Double, ByVal y As Double) As Double

    If x <= y Then
        HW_Min = x
    Else
        HW_Min = y
    End If

End Function
