Option Explicit

'============================================================
' clsHWCalibrator
'
' Hull-White 1F Calibrator
'
' 目的：
'   ・Discount curve と volatility surface を使って、
'     1-factor Hull-White model のパラメータを推定する。
'
' 初期スコープ：
'   ・平均回帰 a は外部から固定値として与える。
'   ・sigma を ATM normal swaption vol に最小二乗でフィットする。
'
' 前提：
'   ・Discount curve は Object として受け取る。
'   ・Curve object は以下の Public method を持つ前提：
'       DF_T(T)
'       ForwardRate_T(T1, T2)
'       InstantaneousForward(T)
'
'   ・Vol surface は clsVolSurface として受け取る。
'   ・初期実装では NormalVol(expiry, tenor) を市場ボラとして使う。
'
' モデル：
'   dr(t) = {theta(t) - a r(t)} dt + sigma dW(t)
'
'   実装上は shifted short-rate representation を想定：
'
'       r(t) = f(0,t) + x(t)
'       dx(t) = -a x(t) dt + sigma dW(t)
'
'   x(T) の標準偏差：
'
'       Std[x(T)] = sigma * sqrt((1 - exp(-2aT)) / (2a))
'
'   a -> 0 の極限：
'
'       Std[x(T)] = sigma * sqrt(T)
'
' Swaption normal vol 近似：
'   ・Expiry 時点の swap rate を x(T) の一次関数として近似する。
'   ・Model normal vol は以下で近似する。
'
'       modelVol = |dS/dx| * Std[x(expiry)] / sqrt(expiry)
'
'   ・dS/dx は、expiry時点の forward discount factor に対する
'     Hull-White bond sensitivity B(T,U) を使って計算する。
'
' 注意：
'   ・本クラスは初期キャリブレーション用の簡易・実務的実装。
'   ・厳密な Jamshidian decomposition による swaption pricing ではない。
'   ・ただし、sigma が normal vol に与える一次的な影響を捉えるため、
'     初期パラメータ推定、シミュレーション用のラフなsigma推定には有用。
'   ・将来的には clsHullWhite1F / modHWMath 側に厳密な bond option / swaption
'     pricing を実装し、本クラスの objective から呼ぶ設計に拡張する。
'
' 依存：
'   ・clsVolSurface
'   ・Curve object with DF_T
'============================================================

Private mCurve As Object
Private mVolSurface As clsVolSurface

Private mPaymentFrequencyYears As Double
Private mIsInitialized As Boolean

Private mQuoteCount As Long
Private mQuoteExpiry() As Double
Private mQuoteTenor() As Double

Private mLastA As Double
Private mLastSigma As Double
Private mLastObjective As Double
Private mHasCalibrationResult As Boolean

Private Const EPS_A As Double = 0.00000001
Private Const EPS_TIME As Double = 0.0000000001

'============================================================
' 初期化
'
' in_Curve:
'   DF_T(T) を持つ curve object
'
' in_VolSurface:
'   NormalVol(expiry, tenor) を持つ clsVolSurface
'
' in_PaymentFrequencyYears:
'   underlying swap の固定レグ支払間隔。
'   初期値は 1.0 年。
'============================================================
Public Sub Init( _
  ByVal in_Curve As Object, _
  ByVal in_VolSurface As clsVolSurface, _
  Optional ByVal in_PaymentFrequencyYears As Double = 1# _
)

  If in_Curve Is Nothing Then
    Err.Raise vbObjectError + 7001, _
      "clsHWCalibrator.Init", _
      "Curve object is Nothing."
  End If

  If in_VolSurface Is Nothing Then
    Err.Raise vbObjectError + 7002, _
      "clsHWCalibrator.Init", _
      "Vol surface is Nothing."
  End If

  If Not in_VolSurface.IsInitialized Then
    Err.Raise vbObjectError + 7003, _
      "clsHWCalibrator.Init", _
      "Vol surface is not initialized."
  End If

  If in_PaymentFrequencyYears <= 0# Then
    Err.Raise vbObjectError + 7004, _
      "clsHWCalibrator.Init", _
      "Payment frequency must be positive."
  End If

  Set mCurve = in_Curve
  Set mVolSurface = in_VolSurface
  mPaymentFrequencyYears = in_PaymentFrequencyYears

  ClearCalibrationQuotes

  mHasCalibrationResult = False
  mIsInitialized = True

End Sub

'============================================================
' 初期化済み判定
'============================================================
Public Property Get IsInitialized() As Boolean
  IsInitialized = mIsInitialized
End Property

'============================================================
' Quote数
'============================================================
Public Property Get QuoteCount() As Long
  QuoteCount = mQuoteCount
End Property

'============================================================
' 固定レグ支払間隔
'============================================================
Public Property Get PaymentFrequencyYears() As Double
  PaymentFrequencyYears = mPaymentFrequencyYears
End Property

'============================================================
' 最後に推定した a
'============================================================
Public Property Get LastA() As Double
  If Not mHasCalibrationResult Then
    Err.Raise vbObjectError + 7011, _
      "clsHWCalibrator.LastA", _
      "Calibration result is not available."
  End If

  LastA = mLastA
End Property

'============================================================
' 最後に推定した sigma
'============================================================
Public Property Get LastSigma() As Double
  If Not mHasCalibrationResult Then
    Err.Raise vbObjectError + 7012, _
      "clsHWCalibrator.LastSigma", _
      "Calibration result is not available."
  End If

  LastSigma = mLastSigma
End Property

'============================================================
' 最後の目的関数値
'============================================================
Public Property Get LastObjective() As Double
  If Not mHasCalibrationResult Then
    Err.Raise vbObjectError + 7013, _
      "clsHWCalibrator.LastObjective", _
      "Calibration result is not available."
  End If

  LastObjective = mLastObjective
End Property

'============================================================
' Calibration quote をクリア
'============================================================
Public Sub ClearCalibrationQuotes()

  mQuoteCount = 0
  Erase mQuoteExpiry
  Erase mQuoteTenor

End Sub

'============================================================
' Calibration quote を1点追加
'
' in_ExpiryYears:
'   swaption expiry in years
'
' in_TenorYears:
'   underlying swap tenor in years
'============================================================
Public Sub AddCalibrationQuote( _
  ByVal in_ExpiryYears As Double, _
  ByVal in_TenorYears As Double _
)

  CheckInitialized

  If in_ExpiryYears <= 0# Then
    Err.Raise vbObjectError + 7021, _
      "clsHWCalibrator.AddCalibrationQuote", _
      "Expiry must be positive."
  End If

  If in_TenorYears <= 0# Then
    Err.Raise vbObjectError + 7022, _
      "clsHWCalibrator.AddCalibrationQuote", _
      "Tenor must be positive."
  End If

  mQuoteCount = mQuoteCount + 1

  ReDim Preserve mQuoteExpiry(1 To mQuoteCount)
  ReDim Preserve mQuoteTenor(1 To mQuoteCount)

  mQuoteExpiry(mQuoteCount) = in_ExpiryYears
  mQuoteTenor(mQuoteCount) = in_TenorYears

End Sub

'============================================================
' 配列から Calibration quote を作成
'============================================================
Public Sub BuildCalibrationQuotes( _
  ByVal in_ExpiryYears As Variant, _
  ByVal in_TenorYears As Variant _
)

  Dim nExp As Long
  Dim nTenor As Long
  Dim i As Long

  CheckInitialized

  nExp = GetVariantItemCount(in_ExpiryYears)
  nTenor = GetVariantItemCount(in_TenorYears)

  If nExp <> nTenor Then
    Err.Raise vbObjectError + 7023, _
      "clsHWCalibrator.BuildCalibrationQuotes", _
      "Expiry array and tenor array must have the same length."
  End If

  ClearCalibrationQuotes

  For i = 1 To nExp
    AddCalibrationQuote _
      CDbl(GetVariantItem(in_ExpiryYears, i)), _
      CDbl(GetVariantItem(in_TenorYears, i))
  Next i

End Sub

'============================================================
' 固定 a に対して sigma をキャリブレーション
'
' 探索方法：
'   ・golden section search
'   ・目的関数は normal vol error の二乗和
'
' 戻り値：
'   calibrated sigma
'============================================================
Public Function CalibrateSigmaFixedA( _
  ByVal in_a As Double, _
  Optional ByVal in_SigmaLower As Double = 0.000001, _
  Optional ByVal in_SigmaUpper As Double = 0.05, _
  Optional ByVal in_Tolerance As Double = 0.0000000001, _
  Optional ByVal in_MaxIterations As Long = 100 _
) As Double

  Dim lower As Double
  Dim upper As Double
  Dim x1 As Double
  Dim x2 As Double
  Dim f1 As Double
  Dim f2 As Double
  Dim gr As Double
  Dim i As Long
  Dim bestSigma As Double
  Dim bestObj As Double

  CheckReadyForCalibration

  If in_a < 0# Then
    Err.Raise vbObjectError + 7031, _
      "clsHWCalibrator.CalibrateSigmaFixedA", _
      "Mean reversion a must be non-negative."
  End If

  If in_SigmaLower < 0# Then
    Err.Raise vbObjectError + 7032, _
      "clsHWCalibrator.CalibrateSigmaFixedA", _
      "Sigma lower bound must be non-negative."
  End If

  If in_SigmaUpper <= in_SigmaLower Then
    Err.Raise vbObjectError + 7033, _
      "clsHWCalibrator.CalibrateSigmaFixedA", _
      "Sigma upper bound must be greater than lower bound."
  End If

  If in_Tolerance <= 0# Then
    Err.Raise vbObjectError + 7034, _
      "clsHWCalibrator.CalibrateSigmaFixedA", _
      "Tolerance must be positive."
  End If

  If in_MaxIterations <= 0 Then
    Err.Raise vbObjectError + 7035, _
      "clsHWCalibrator.CalibrateSigmaFixedA", _
      "MaxIterations must be positive."
  End If

  lower = in_SigmaLower
  upper = in_SigmaUpper

  gr = (Sqr(5#) - 1#) / 2#

  x1 = upper - gr * (upper - lower)
  x2 = lower + gr * (upper - lower)

  f1 = ObjectiveSigmaFixedA(x1, in_a)
  f2 = ObjectiveSigmaFixedA(x2, in_a)

  For i = 1 To in_MaxIterations

    If Abs(upper - lower) <= in_Tolerance Then
      Exit For
    End If

    If f1 > f2 Then
      lower = x1
      x1 = x2
      f1 = f2
      x2 = lower + gr * (upper - lower)
      f2 = ObjectiveSigmaFixedA(x2, in_a)
    Else
      upper = x2
      x2 = x1
      f2 = f1
      x1 = upper - gr * (upper - lower)
      f1 = ObjectiveSigmaFixedA(x1, in_a)
    End If

  Next i

  bestSigma = 0.5 * (lower + upper)
  bestObj = ObjectiveSigmaFixedA(bestSigma, in_a)

  mLastA = in_a
  mLastSigma = bestSigma
  mLastObjective = bestObj
  mHasCalibrationResult = True

  CalibrateSigmaFixedA = bestSigma

End Function

'============================================================
' 指定 a, sigma に対する目的関数値を返す
'============================================================
Public Function ObjectiveValue( _
  ByVal in_a As Double, _
  ByVal in_sigma As Double _
) As Double

  CheckReadyForCalibration

  ObjectiveValue = ObjectiveSigmaFixedA(in_sigma, in_a)

End Function

'============================================================
' Hull-White 近似 normal swaption vol
'
' 戻り値：
'   annualized normal vol
'============================================================
Public Function ModelNormalVol( _
  ByVal in_a As Double, _
  ByVal in_sigma As Double, _
  ByVal in_ExpiryYears As Double, _
  ByVal in_TenorYears As Double _
) As Double

  Dim sens As Double
  Dim stdX As Double

  CheckInitialized

  If in_a < 0# Then
    Err.Raise vbObjectError + 7041, _
      "clsHWCalibrator.ModelNormalVol", _
      "Mean reversion a must be non-negative."
  End If

  If in_sigma < 0# Then
    Err.Raise vbObjectError + 7042, _
      "clsHWCalibrator.ModelNormalVol", _
      "Sigma must be non-negative."
  End If

  If in_ExpiryYears <= 0# Then
    Err.Raise vbObjectError + 7043, _
      "clsHWCalibrator.ModelNormalVol", _
      "Expiry must be positive."
  End If

  If in_TenorYears <= 0# Then
    Err.Raise vbObjectError + 7044, _
      "clsHWCalibrator.ModelNormalVol", _
      "Tenor must be positive."
  End If

  sens = Abs(SwapRateSensitivityToX(in_a, in_ExpiryYears, in_TenorYears))
  stdX = StdDevX(in_a, in_sigma, in_ExpiryYears)

  ModelNormalVol = sens * stdX / Sqr(in_ExpiryYears)

End Function

'============================================================
' Forward swap rate を返す
'
' expiry時点スタート、expiry + tenor 満期の forward swap rate。
' 固定レグは mPaymentFrequencyYears 間隔。
'============================================================
Public Function ForwardSwapRate( _
  ByVal in_ExpiryYears As Double, _
  ByVal in_TenorYears As Double _
) As Double

  Dim startT As Double
  Dim endT As Double
  Dim annuity As Double
  Dim numerator As Double
  Dim payTimes() As Double
  Dim accruals() As Double
  Dim n As Long
  Dim i As Long

  CheckInitialized

  If in_ExpiryYears < 0# Then
    Err.Raise vbObjectError + 7051, _
      "clsHWCalibrator.ForwardSwapRate", _
      "Expiry must be non-negative."
  End If

  If in_TenorYears <= 0# Then
    Err.Raise vbObjectError + 7052, _
      "clsHWCalibrator.ForwardSwapRate", _
      "Tenor must be positive."
  End If

  startT = in_ExpiryYears
  endT = in_ExpiryYears + in_TenorYears

  BuildPaymentSchedule startT, endT, payTimes, accruals, n

  annuity = 0#

  For i = 1 To n
    annuity = annuity + accruals(i) * ForwardDF(startT, payTimes(i))
  Next i

  If annuity <= 0# Then
    Err.Raise vbObjectError + 7053, _
      "clsHWCalibrator.ForwardSwapRate", _
      "Swap annuity must be positive."
  End If

  numerator = 1# - ForwardDF(startT, endT)

  ForwardSwapRate = numerator / annuity

End Function

'============================================================
' Calibration report を返す
'
' 列：
'   1 Quote No
'   2 Expiry
'   3 Tenor
'   4 Market Normal Vol
'   5 Model Normal Vol
'   6 Error
'   7 Squared Error
'============================================================
Public Function GetCalibrationReport() As Variant

  Dim arr() As Variant
  Dim i As Long
  Dim marketVol As Double
  Dim modelVol As Double
  Dim errVol As Double

  CheckReadyForCalibration

  If Not mHasCalibrationResult Then
    Err.Raise vbObjectError + 7061, _
      "clsHWCalibrator.GetCalibrationReport", _
      "Calibration result is not available."
  End If

  ReDim arr(1 To mQuoteCount + 1, 1 To 7)

  arr(1, 1) = "Quote No"
  arr(1, 2) = "Expiry"
  arr(1, 3) = "Tenor"
  arr(1, 4) = "Market Normal Vol"
  arr(1, 5) = "Model Normal Vol"
  arr(1, 6) = "Error"
  arr(1, 7) = "Squared Error"

  For i = 1 To mQuoteCount
    marketVol = mVolSurface.NormalVol(mQuoteExpiry(i), mQuoteTenor(i))
    modelVol = ModelNormalVol(mLastA, mLastSigma, mQuoteExpiry(i), mQuoteTenor(i))
    errVol = modelVol - marketVol

    arr(i + 1, 1) = i
    arr(i + 1, 2) = mQuoteExpiry(i)
    arr(i + 1, 3) = mQuoteTenor(i)
    arr(i + 1, 4) = marketVol
    arr(i + 1, 5) = modelVol
    arr(i + 1, 6) = errVol
    arr(i + 1, 7) = errVol * errVol
  Next i

  GetCalibrationReport = arr

End Function

'============================================================
' Summary
'============================================================
Public Function Summary() As String

  CheckInitialized

  Summary = "clsHWCalibrator" & vbCrLf & _
            "Initialized = " & CStr(mIsInitialized) & vbCrLf & _
            "QuoteCount = " & CStr(mQuoteCount) & vbCrLf & _
            "PaymentFrequencyYears = " & CStr(mPaymentFrequencyYears) & vbCrLf & _
            "HasCalibrationResult = " & CStr(mHasCalibrationResult)

  If mHasCalibrationResult Then
    Summary = Summary & vbCrLf & _
              "LastA = " & CStr(mLastA) & vbCrLf & _
              "LastSigma = " & CStr(mLastSigma) & vbCrLf & _
              "LastObjective = " & CStr(mLastObjective)
  End If

End Function

'============================================================
' 内部処理：目的関数
'============================================================
Private Function ObjectiveSigmaFixedA( _
  ByVal in_sigma As Double, _
  ByVal in_a As Double _
) As Double

  Dim i As Long
  Dim marketVol As Double
  Dim modelVol As Double
  Dim errVol As Double
  Dim total As Double

  If in_sigma < 0# Then
    Err.Raise vbObjectError + 7101, _
      "clsHWCalibrator.ObjectiveSigmaFixedA", _
      "Sigma must be non-negative."
  End If

  total = 0#

  For i = 1 To mQuoteCount
    marketVol = mVolSurface.NormalVol(mQuoteExpiry(i), mQuoteTenor(i))
    modelVol = ModelNormalVol(in_a, in_sigma, mQuoteExpiry(i), mQuoteTenor(i))
    errVol = modelVol - marketVol
    total = total + errVol * errVol
  Next i

  ObjectiveSigmaFixedA = total

End Function

'============================================================
' 内部処理：swap rate の x に対する感応度 dS/dx
'
' expiry時点での forward discount factor を基準に計算する。
'
' P_T(U) = P(0,U) / P(0,T)
' dP_T(U)/dx = -B(T,U) * P_T(U)
'
' S = N / A
' N = 1 - P_T(end)
' A = sum_i accrual_i * P_T(pay_i)
'
' dS/dx = (dN * A - N * dA) / A^2
'============================================================
Private Function SwapRateSensitivityToX( _
  ByVal in_a As Double, _
  ByVal in_ExpiryYears As Double, _
  ByVal in_TenorYears As Double _
) As Double

  Dim startT As Double
  Dim endT As Double
  Dim payTimes() As Double
  Dim accruals() As Double
  Dim n As Long
  Dim i As Long

  Dim pEnd As Double
  Dim pPay As Double
  Dim bEnd As Double
  Dim bPay As Double

  Dim numerator As Double
  Dim annuity As Double
  Dim dNumerator As Double
  Dim dAnnuity As Double

  startT = in_ExpiryYears
  endT = in_ExpiryYears + in_TenorYears

  BuildPaymentSchedule startT, endT, payTimes, accruals, n

  pEnd = ForwardDF(startT, endT)
  bEnd = HW_B(in_a, startT, endT)

  numerator = 1# - pEnd
  dNumerator = bEnd * pEnd

  annuity = 0#
  dAnnuity = 0#

  For i = 1 To n
    pPay = ForwardDF(startT, payTimes(i))
    bPay = HW_B(in_a, startT, payTimes(i))

    annuity = annuity + accruals(i) * pPay
    dAnnuity = dAnnuity - accruals(i) * bPay * pPay
  Next i

  If annuity <= 0# Then
    Err.Raise vbObjectError + 7111, _
      "clsHWCalibrator.SwapRateSensitivityToX", _
      "Swap annuity must be positive."
  End If

  SwapRateSensitivityToX = (dNumerator * annuity - numerator * dAnnuity) / (annuity * annuity)

End Function

'============================================================
' 内部処理：x(T) の標準偏差
'============================================================
Private Function StdDevX( _
  ByVal in_a As Double, _
  ByVal in_sigma As Double, _
  ByVal in_T As Double _
) As Double

  Dim v As Double

  If in_T <= 0# Then
    Err.Raise vbObjectError + 7121, _
      "clsHWCalibrator.StdDevX", _
      "T must be positive."
  End If

  If Abs(in_a) < EPS_A Then
    v = in_sigma * in_sigma * in_T
  Else
    v = in_sigma * in_sigma * (1# - Exp(-2# * in_a * in_T)) / (2# * in_a)
  End If

  If v < 0# Then
    Err.Raise vbObjectError + 7122, _
      "clsHWCalibrator.StdDevX", _
      "Variance is negative."
  End If

  StdDevX = Sqr(v)

End Function

'============================================================
' 内部処理：Hull-White B(t,T)
'
' B(t,T) = (1 - exp(-a(T-t))) / a
' a -> 0 の極限は T - t
'============================================================
Private Function HW_B( _
  ByVal in_a As Double, _
  ByVal in_t As Double, _
  ByVal in_T As Double _
) As Double

  Dim dt As Double

  If in_T < in_t Then
    Err.Raise vbObjectError + 7131, _
      "clsHWCalibrator.HW_B", _
      "T must be greater than or equal to t."
  End If

  dt = in_T - in_t

  If Abs(dt) < EPS_TIME Then
    HW_B = 0#
  ElseIf Abs(in_a) < EPS_A Then
    HW_B = dt
  Else
    HW_B = (1# - Exp(-in_a * dt)) / in_a
  End If

End Function

'============================================================
' 内部処理：Forward discount factor P(0,T2) / P(0,T1)
'============================================================
Private Function ForwardDF( _
  ByVal in_T1 As Double, _
  ByVal in_T2 As Double _
) As Double

  Dim df1 As Double
  Dim df2 As Double

  If in_T2 < in_T1 Then
    Err.Raise vbObjectError + 7141, _
      "clsHWCalibrator.ForwardDF", _
      "T2 must be greater than or equal to T1."
  End If

  df1 = CurveDF(in_T1)
  df2 = CurveDF(in_T2)

  If df1 <= 0# Or df2 <= 0# Then
    Err.Raise vbObjectError + 7142, _
      "clsHWCalibrator.ForwardDF", _
      "Discount factors must be positive."
  End If

  ForwardDF = df2 / df1

End Function

'============================================================
' 内部処理：Curve DF_T wrapper
'============================================================
Private Function CurveDF(ByVal in_T As Double) As Double

  If in_T < 0# Then
    Err.Raise vbObjectError + 7151, _
      "clsHWCalibrator.CurveDF", _
      "T must be non-negative."
  End If

  CurveDF = CDbl(mCurve.DF_T(in_T))

  If CurveDF <= 0# Then
    Err.Raise vbObjectError + 7152, _
      "clsHWCalibrator.CurveDF", _
      "Curve returned non-positive DF."
  End If

End Function

'============================================================
' 内部処理：固定レグ支払スケジュールを作成
'
' startT から endT まで、mPaymentFrequencyYears 間隔。
' 最終期間はスタブを許容する。
'============================================================
Private Sub BuildPaymentSchedule( _
  ByVal in_StartT As Double, _
  ByVal in_EndT As Double, _
  ByRef out_PayTimes() As Double, _
  ByRef out_Accruals() As Double, _
  ByRef out_Count As Long _
)

  Dim nFull As Long
  Dim i As Long
  Dim prevT As Double
  Dim payT As Double

  If in_EndT <= in_StartT Then
    Err.Raise vbObjectError + 7161, _
      "clsHWCalibrator.BuildPaymentSchedule", _
      "EndT must be greater than StartT."
  End If

  nFull = CLng(Int((in_EndT - in_StartT) / mPaymentFrequencyYears))

  If in_StartT + CDbl(nFull) * mPaymentFrequencyYears < in_EndT - EPS_TIME Then
    out_Count = nFull + 1
  Else
    out_Count = nFull
  End If

  If out_Count <= 0 Then
    out_Count = 1
  End If

  ReDim out_PayTimes(1 To out_Count)
  ReDim out_Accruals(1 To out_Count)

  prevT = in_StartT

  For i = 1 To out_Count
    payT = in_StartT + CDbl(i) * mPaymentFrequencyYears

    If payT > in_EndT Or i = out_Count Then
      payT = in_EndT
    End If

    If payT <= prevT Then
      Err.Raise vbObjectError + 7162, _
        "clsHWCalibrator.BuildPaymentSchedule", _
        "Invalid payment schedule."
    End If

    out_PayTimes(i) = payT
    out_Accruals(i) = payT - prevT

    prevT = payT
  Next i

End Sub

'============================================================
' 内部処理：キャリブレーション可能状態チェック
'============================================================
Private Sub CheckReadyForCalibration()

  CheckInitialized

  If mQuoteCount <= 0 Then
    Err.Raise vbObjectError + 7171, _
      "clsHWCalibrator", _
      "No calibration quotes. Add quotes first."
  End If

End Sub

'============================================================
' 内部処理：初期化チェック
'============================================================
Private Sub CheckInitialized()

  If Not mIsInitialized Then
    Err.Raise vbObjectError + 7181, _
      "clsHWCalibrator", _
      "clsHWCalibrator is not initialized. Call Init first."
  End If

End Sub

'============================================================
' 内部処理：Variant要素数を取得
'============================================================
Private Function GetVariantItemCount(ByVal in_v As Variant) As Long

  On Error GoTo TryTwoDim

  If IsArray(in_v) Then
    GetVariantItemCount = UBound(in_v) - LBound(in_v) + 1
  Else
    GetVariantItemCount = 1
  End If

  Exit Function

TryTwoDim:
  On Error GoTo SingleValue

  If IsArray(in_v) Then
    If UBound(in_v, 1) - LBound(in_v, 1) + 1 >= _
       UBound(in_v, 2) - LBound(in_v, 2) + 1 Then

      GetVariantItemCount = UBound(in_v, 1) - LBound(in_v, 1) + 1

    Else

      GetVariantItemCount = UBound(in_v, 2) - LBound(in_v, 2) + 1

    End If
  Else
    GetVariantItemCount = 1
  End If

  Exit Function

SingleValue:
  GetVariantItemCount = 1

End Function

'============================================================
' 内部処理：Variantからi番目の値を取得
'============================================================
Private Function GetVariantItem( _
  ByVal in_v As Variant, _
  ByVal in_Index As Long _
) As Variant

  On Error GoTo TryTwoDimColumn

  If IsArray(in_v) Then
    GetVariantItem = in_v(LBound(in_v) + in_Index - 1)
  Else
    GetVariantItem = in_v
  End If

  Exit Function

TryTwoDimColumn:
  On Error GoTo TryTwoDimRow

  GetVariantItem = in_v(LBound(in_v, 1) + in_Index - 1, LBound(in_v, 2))
  Exit Function

TryTwoDimRow:
  On Error GoTo SingleValue

  GetVariantItem = in_v(LBound(in_v, 1), LBound(in_v, 2) + in_Index - 1)
  Exit Function

SingleValue:
  GetVariantItem = in_v

End Function
