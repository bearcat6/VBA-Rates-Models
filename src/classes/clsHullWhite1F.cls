Option Explicit

'====================================================
' clsHullWhite1F
'
' 1-factor Hull-White model with constant mean reversion and volatility.
'
' Model:
'   dr(t) = {theta(t) - a * r(t)} dt + sigma dW(t)
'
' Equivalent shifted-state representation:
'   r(t) = x(t) + phi(t)
'   dx(t) = -a * x(t) dt + sigma dW(t), x(0) = 0
'
' This class stores:
'   ・initial discount curve
'   ・valuation date
'   ・mean reversion a
'   ・short-rate volatility sigma
'
' Main uses:
'   ・A(t,T), B(t,T) calculation
'   ・discount bond P(t,T) from short rate r(t)
'   ・exact one-step simulation of shifted state x(t)
'   ・initial forward / fitted shift phi(t)
'
' Notes:
'   ・Times are year fractions measured from valuation date.
'   ・The input curve must implement clsDiscountCurve.
'   ・Instantaneous forward is estimated by finite difference of ln DF.
'====================================================

Private Const DAYS_PER_YEAR As Double = 365.25
Private Const EPS As Double = 0.0000000001

Private mCurve As clsDiscountCurve
Private mValuationDate As Date
Private mA As Double
Private mSigma As Double
Private mIsInitialized As Boolean

'====================================================
' Initializer
'====================================================
Public Sub Initialize( _
  ByVal in_DiscountCurve As clsDiscountCurve, _
  ByVal in_ValuationDate As Date, _
  ByVal in_MeanReversion As Double, _
  ByVal in_Sigma As Double)

  If in_DiscountCurve Is Nothing Then
    Err.Raise vbObjectError + 1000, "clsHullWhite1F.Initialize", _
          "discount curve is Nothing."
  End If
  
  If in_MeanReversion <= 0 Then
    Err.Raise vbObjectError + 1001, "clsHullWhite1F.Initialize", _
          "mean reversion must be positive."
  End If
  
  If in_Sigma < 0 Then
    Err.Raise vbObjectError + 1002, "clsHullWhite1F.Initialize", _
          "sigma must be non-negative."
  End If
  
  Set mCurve = in_DiscountCurve
  mValuationDate = in_ValuationDate
  mA = in_MeanReversion
  mSigma = in_Sigma
  mIsInitialized = True

End Sub

'====================================================
' 初期化済み判定
'====================================================
Public Property Get IsInitialized() As Boolean
  IsInitialized = mIsInitialized
End Property

'====================================================
' Valuation Date
'====================================================
Public Property Get ValuationDate() As Date
  EnsureInitialized "ValuationDate"
  ValuationDate = mValuationDate
End Property

'====================================================
' Mean Reversion a
'====================================================
Public Property Get MeanReversion() As Double
  EnsureInitialized "MeanReversion"
  MeanReversion = mA
End Property

'====================================================
' Short-rate volatility sigma
'====================================================
Public Property Get Sigma() As Double
  EnsureInitialized "Sigma"
  Sigma = mSigma
End Property

'====================================================
' B(t,T) = (1 - exp(-a(T-t))) / a
'====================================================
Public Function B( _
  ByVal in_T As Double, _
  ByVal in_TerminalT As Double) As Double

  Dim tau As Double
  
  EnsureInitialized "B"
  ValidateTimeOrder in_T, in_TerminalT, "B"
  
  tau = in_TerminalT - in_T
  
  B = (1# - Exp(-mA * tau)) / mA

End Function

'====================================================
' A(t,T)
'
' P(t,T) = A(t,T) * exp(-B(t,T) * r(t))
'
' A(t,T) is chosen so that the model fits the initial curve.
'====================================================
Public Function A( _
  ByVal in_T As Double, _
  ByVal in_TerminalT As Double) As Double

  Dim p0Start As Double
  Dim p0Maturity As Double
  Dim bVal As Double
  Dim f0Start As Double
  Dim varianceAdj As Double
  
  EnsureInitialized "A"
  ValidateTimeOrder in_T, in_TerminalT, "A"
  
  p0Start = DFByYears(in_T)
  p0Maturity = DFByYears(in_TerminalT)
  bVal = B(in_T, in_TerminalT)
  f0Start = InstantForward(in_T)
  
  varianceAdj = (mSigma * mSigma / (4# * mA)) _
              * (1# - Exp(-2# * mA * in_T)) _
              * bVal * bVal
  
  A = (p0Maturity / p0Start) * Exp(bVal * f0Start - varianceAdj)

End Function

'====================================================
' P(t,T) from short rate r(t)
'====================================================
Public Function DiscountBondByShortRate( _
  ByVal in_T As Double, _
  ByVal in_TerminalT As Double, _
  ByVal in_ShortRate As Double) As Double

  Dim aVal As Double
  Dim bVal As Double
  
  EnsureInitialized "DiscountBondByShortRate"
  ValidateTimeOrder in_T, in_TerminalT, "DiscountBondByShortRate"
  
  aVal = A(in_T, in_TerminalT)
  bVal = B(in_T, in_TerminalT)
  
  DiscountBondByShortRate = aVal * Exp(-bVal * in_ShortRate)

End Function

'====================================================
' P(t,T) from shifted state x(t)
'====================================================
Public Function DiscountBondByState( _
  ByVal in_T As Double, _
  ByVal in_TerminalT As Double, _
  ByVal in_StateX As Double) As Double

  Dim r As Double
  
  EnsureInitialized "DiscountBondByState"
  r = ShortRateFromState(in_T, in_StateX)
  
  DiscountBondByState = DiscountBondByShortRate(in_T, in_TerminalT, r)

End Function

'====================================================
' Initial discount factor P(0,T)
'====================================================
Public Function DFByYears(ByVal in_T As Double) As Double

  Dim targetDate As Date
  
  EnsureInitialized "DFByYears"
  ValidateNonNegativeTime in_T, "DFByYears"
  
  targetDate = YearsToDate(in_T)
  DFByYears = mCurve.DF(targetDate)
  
  If DFByYears <= 0 Then
    Err.Raise vbObjectError + 1100, "clsHullWhite1F.DFByYears", _
          "discount factor must be positive."
  End If

End Function

'====================================================
' Initial zero rate, continuously compounded
'====================================================
Public Function ZeroRateContByYears(ByVal in_T As Double) As Double

  EnsureInitialized "ZeroRateContByYears"
  ValidatePositiveTime in_T, "ZeroRateContByYears"
  
  ZeroRateContByYears = -Log(DFByYears(in_T)) / in_T

End Function

'====================================================
' Initial instantaneous forward f(0,t)
'
' Estimated from:
'   f(0,t) = - d ln P(0,t) / dt
'====================================================
Public Function InstantForward(ByVal in_T As Double) As Double

  Dim h As Double
  Dim t1 As Double
  Dim t2 As Double
  Dim p1 As Double
  Dim p2 As Double
  
  EnsureInitialized "InstantForward"
  ValidateNonNegativeTime in_T, "InstantForward"
  
  h = ForwardDiffStep(in_T)
  
  If in_T <= h Then
    t1 = 0#
    t2 = in_T + h
  Else
    t1 = in_T - h
    t2 = in_T + h
  End If
  
  p1 = DFByYears(t1)
  p2 = DFByYears(t2)
  
  InstantForward = -(Log(p2) - Log(p1)) / (t2 - t1)

End Function

'====================================================
' phi(t)
'
' r(t) = x(t) + phi(t)
'
' For constant a and sigma:
'   phi(t) = f(0,t) + sigma^2 / (2a^2) * (1 - exp(-a t))^2
'====================================================
Public Function Phi(ByVal in_T As Double) As Double

  Dim adj As Double
  
  EnsureInitialized "Phi"
  ValidateNonNegativeTime in_T, "Phi"
  
  adj = (mSigma * mSigma / (2# * mA * mA)) _
      * (1# - Exp(-mA * in_T)) _
      * (1# - Exp(-mA * in_T))
  
  Phi = InstantForward(in_T) + adj

End Function

'====================================================
' r(t) = x(t) + phi(t)
'====================================================
Public Function ShortRateFromState( _
  ByVal in_T As Double, _
  ByVal in_StateX As Double) As Double

  EnsureInitialized "ShortRateFromState"
  ValidateNonNegativeTime in_T, "ShortRateFromState"
  
  ShortRateFromState = in_StateX + Phi(in_T)

End Function

'====================================================
' x(t+dt) exact step under risk-neutral measure
'
' x(t+dt) = x(t) exp(-a dt)
'         + sigma sqrt((1 - exp(-2a dt)) / (2a)) Z
'====================================================
Public Function NextStateExact( _
  ByVal in_StateX As Double, _
  ByVal in_Dt As Double, _
  ByVal in_NormalZ As Double) As Double

  Dim mean As Double
  Dim stdev As Double
  
  EnsureInitialized "NextStateExact"
  ValidatePositiveTime in_Dt, "NextStateExact"
  
  mean = in_StateX * Exp(-mA * in_Dt)
  stdev = StateStepStdDev(in_Dt)
  
  NextStateExact = mean + stdev * in_NormalZ

End Function

'====================================================
' r(t+dt) exact step via shifted state
'====================================================
Public Function NextShortRateExact( _
  ByVal in_T As Double, _
  ByVal in_StateX As Double, _
  ByVal in_Dt As Double, _
  ByVal in_NormalZ As Double) As Double

  Dim nextT As Double
  Dim nextX As Double
  
  EnsureInitialized "NextShortRateExact"
  ValidateNonNegativeTime in_T, "NextShortRateExact"
  ValidatePositiveTime in_Dt, "NextShortRateExact"
  
  nextT = in_T + in_Dt
  nextX = NextStateExact(in_StateX, in_Dt, in_NormalZ)
  
  NextShortRateExact = ShortRateFromState(nextT, nextX)

End Function

'====================================================
' Conditional variance of x(t+dt) given x(t)
'====================================================
Public Function StateStepVariance(ByVal in_Dt As Double) As Double

  EnsureInitialized "StateStepVariance"
  ValidatePositiveTime in_Dt, "StateStepVariance"
  
  StateStepVariance = (mSigma * mSigma) _
                    * (1# - Exp(-2# * mA * in_Dt)) _
                    / (2# * mA)

End Function

'====================================================
' Conditional standard deviation of x(t+dt) given x(t)
'====================================================
Public Function StateStepStdDev(ByVal in_Dt As Double) As Double

  Dim v As Double
  
  EnsureInitialized "StateStepStdDev"
  v = StateStepVariance(in_Dt)
  
  If v < 0 Then
    Err.Raise vbObjectError + 1200, "clsHullWhite1F.StateStepStdDev", _
          "state variance is negative."
  End If
  
  StateStepStdDev = Sqr(v)

End Function

'====================================================
' Unconditional variance of x(t), x(0)=0
'====================================================
Public Function StateVariance(ByVal in_T As Double) As Double

  EnsureInitialized "StateVariance"
  ValidateNonNegativeTime in_T, "StateVariance"
  
  If in_T = 0# Then
    StateVariance = 0#
  Else
    StateVariance = (mSigma * mSigma) _
                  * (1# - Exp(-2# * mA * in_T)) _
                  / (2# * mA)
  End If

End Function

'====================================================
' Unconditional standard deviation of x(t), x(0)=0
'====================================================
Public Function StateStdDev(ByVal in_T As Double) As Double

  Dim v As Double
  
  EnsureInitialized "StateStdDev"
  v = StateVariance(in_T)
  
  If v < 0 Then
    Err.Raise vbObjectError + 1210, "clsHullWhite1F.StateStdDev", _
          "state variance is negative."
  End If
  
  StateStdDev = Sqr(v)

End Function

'====================================================
' Forward measure bond volatility integral helper
'
' Integrated variance of ln P(t,T) over [0, optionExpiry]
' for bond maturity T.
'
' Useful for analytic bond option / swaption calibration components.
'====================================================
Public Function BondLogVariance( _
  ByVal in_OptionExpiry As Double, _
  ByVal in_BondMaturity As Double) As Double

  Dim b0 As Double
  
  EnsureInitialized "BondLogVariance"
  ValidateTimeOrder 0#, in_OptionExpiry, "BondLogVariance"
  ValidateTimeOrder in_OptionExpiry, in_BondMaturity, "BondLogVariance"
  
  b0 = B(0#, in_BondMaturity - in_OptionExpiry)
  
  ' Equivalent to:
  ' sigma^2 / (2a) * (1 - exp(-2aS)) * B(S,T)^2
  ' where S = option expiry.
  BondLogVariance = (mSigma * mSigma) _
                  * (1# - Exp(-2# * mA * in_OptionExpiry)) _
                  / (2# * mA) _
                  * b0 * b0

End Function

'====================================================
' Utility: year fraction to date
'====================================================
Private Function YearsToDate(ByVal in_T As Double) As Date

  Dim daysToAdd As Long
  
  daysToAdd = CLng(in_T * DAYS_PER_YEAR + 0.5)
  YearsToDate = DateAdd("d", daysToAdd, mValuationDate)

End Function

'====================================================
' Utility: finite-difference step for instantaneous forward
'====================================================
Private Function ForwardDiffStep(ByVal in_T As Double) As Double

  ' About one week, but not too large for very short expiries.
  If in_T < 0.05 Then
    ForwardDiffStep = 1# / 365.25
  Else
    ForwardDiffStep = 7# / 365.25
  End If

End Function

'====================================================
' Validation helpers
'====================================================
Private Sub EnsureInitialized(ByVal in_MethodName As String)

  If Not mIsInitialized Then
    Err.Raise vbObjectError + 9000, "clsHullWhite1F." & in_MethodName, _
          "Hull-White 1F model is not initialized."
  End If

End Sub

Private Sub ValidateNonNegativeTime( _
  ByVal in_T As Double, _
  ByVal in_MethodName As String)

  If in_T < 0# Then
    Err.Raise vbObjectError + 9010, "clsHullWhite1F." & in_MethodName, _
          "time must be non-negative."
  End If

End Sub

Private Sub ValidatePositiveTime( _
  ByVal in_T As Double, _
  ByVal in_MethodName As String)

  If in_T <= 0# Then
    Err.Raise vbObjectError + 9011, "clsHullWhite1F." & in_MethodName, _
          "time must be positive."
  End If

End Sub

Private Sub ValidateTimeOrder( _
  ByVal in_T1 As Double, _
  ByVal in_T2 As Double, _
  ByVal in_MethodName As String)

  If in_T1 < 0# Or in_T2 < 0# Then
    Err.Raise vbObjectError + 9020, "clsHullWhite1F." & in_MethodName, _
          "time must be non-negative."
  End If
  
  If in_T2 < in_T1 Then
    Err.Raise vbObjectError + 9021, "clsHullWhite1F." & in_MethodName, _
          "terminal time must be greater than or equal to start time."
  End If

End Sub
