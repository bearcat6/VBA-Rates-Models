Option Explicit

'============================================================
' clsHWSimulator
'
' Hull-White 1F Monte Carlo Simulator
'
' 目的：
'   ・Discount curve と Hull-White 1F parameters (a, sigma) を使って、
'     shifted short-rate factor x(t) を Monte Carlo simulation する。
'
'   ・将来時点 horizon における short rate、zero rate curve、
'     discount factor curve、forward rate curve を生成する。
'
' モデル：
'   r(t) = f(0,t) + x(t)
'
'   dx(t) = -a x(t) dt + sigma dW(t)
'
' Exact discretization:
'   x(t+dt) = x(t) exp(-a dt)
'             + sigma sqrt((1 - exp(-2a dt)) / (2a)) Z
'
'   a -> 0 の極限：
'   x(t+dt) = x(t) + sigma sqrt(dt) Z
'
' Curve interface:
'   in_Curve は Object として受け取る。
'   以下の Public method を持つ前提：
'       DF_T(T)
'       InstantaneousForward(T)
'
' 注意：
'   ・初期実装では theta(t) は明示的に扱わない。
'   ・将来時点 horizon のカーブは、初期カーブを基準に
'     Hull-White の bond sensitivity B(horizon,T) と x(horizon) を使って
'     簡易的にシフトする。
'   ・厳密な無裁定将来カーブ生成というより、ストレスカーブ・分位点カーブ
'     作成のための実務的な初期版である。
'
' 依存：
'   ・Curve object with DF_T and InstantaneousForward
'============================================================

Private mCurve As Object
Private mA As Double
Private mSigma As Double
Private mDt As Double
Private mSteps As Long
Private mHorizonYears As Double

Private mNumPaths As Long
Private mIsInitialized As Boolean
Private mHasSimulation As Boolean

Private mXPaths() As Double
Private mShortRatePaths() As Double
Private mTimes() As Double

Private Const EPS_A As Double = 0.00000001
Private Const EPS_TIME As Double = 0.0000000001
Private Const PI_VALUE As Double = 3.14159265358979

'============================================================
' 初期化
'
' in_Curve:
'   DF_T, InstantaneousForward を持つ curve object
'
' in_a:
'   mean reversion
'
' in_sigma:
'   short-rate factor volatility
'
' in_HorizonYears:
'   simulation horizon
'
' in_Dt:
'   time step in years
'============================================================
Public Sub Init( _
  ByVal in_Curve As Object, _
  ByVal in_a As Double, _
  ByVal in_sigma As Double, _
  ByVal in_HorizonYears As Double, _
  Optional ByVal in_Dt As Double = 0.00396825396825397 _
)

  If in_Curve Is Nothing Then
    Err.Raise vbObjectError + 8001, _
      "clsHWSimulator.Init", _
      "Curve object is Nothing."
  End If

  If in_a < 0# Then
    Err.Raise vbObjectError + 8002, _
      "clsHWSimulator.Init", _
      "Mean reversion a must be non-negative."
  End If

  If in_sigma < 0# Then
    Err.Raise vbObjectError + 8003, _
      "clsHWSimulator.Init", _
      "Sigma must be non-negative."
  End If

  If in_HorizonYears <= 0# Then
    Err.Raise vbObjectError + 8004, _
      "clsHWSimulator.Init", _
      "HorizonYears must be positive."
  End If

  If in_Dt <= 0# Then
    Err.Raise vbObjectError + 8005, _
      "clsHWSimulator.Init", _
      "Dt must be positive."
  End If

  Set mCurve = in_Curve
  mA = in_a
  mSigma = in_sigma
  mHorizonYears = in_HorizonYears
  mDt = in_Dt

  mSteps = CLng(Application.WorksheetFunction.RoundUp(mHorizonYears / mDt, 0))

  If mSteps <= 0 Then
    Err.Raise vbObjectError + 8006, _
      "clsHWSimulator.Init", _
      "Invalid number of simulation steps."
  End If

  BuildTimeGrid

  mNumPaths = 0
  mHasSimulation = False
  mIsInitialized = True

End Sub

'============================================================
' 初期化済み判定
'============================================================
Public Property Get IsInitialized() As Boolean
  IsInitialized = mIsInitialized
End Property

'============================================================
' Simulation済み判定
'============================================================
Public Property Get HasSimulation() As Boolean
  HasSimulation = mHasSimulation
End Property

'============================================================
' Mean reversion a
'============================================================
Public Property Get MeanReversion() As Double
  MeanReversion = mA
End Property

'============================================================
' Sigma
'============================================================
Public Property Get Sigma() As Double
  Sigma = mSigma
End Property

'============================================================
' Horizon years
'============================================================
Public Property Get HorizonYears() As Double
  HorizonYears = mHorizonYears
End Property

'============================================================
' Time step
'============================================================
Public Property Get Dt() As Double
  Dt = mDt
End Property

'============================================================
' Steps
'============================================================
Public Property Get Steps() As Long
  Steps = mSteps
End Property

'============================================================
' Number of paths
'============================================================
Public Property Get NumPaths() As Long
  NumPaths = mNumPaths
End Property

'============================================================
' Monte Carlo simulation 実行
'
' in_NumPaths:
'   path数
'
' in_Seed:
'   0以下なら Randomize のみ
'   正なら Rnd のシードに使う
'============================================================
Public Sub Simulate( _
  ByVal in_NumPaths As Long, _
  Optional ByVal in_Seed As Long = 0 _
)

  Dim p As Long
  Dim s As Long
  Dim z As Double
  Dim xPrev As Double
  Dim xNext As Double
  Dim dtStep As Double
  Dim tNext As Double

  CheckInitialized

  If in_NumPaths <= 0 Then
    Err.Raise vbObjectError + 8011, _
      "clsHWSimulator.Simulate", _
      "NumPaths must be positive."
  End If

  mNumPaths = in_NumPaths

  ReDim mXPaths(1 To mNumPaths, 0 To mSteps)
  ReDim mShortRatePaths(1 To mNumPaths, 0 To mSteps)

  If in_Seed > 0 Then
    Rnd -CDbl(in_Seed)
    Randomize in_Seed
  Else
    Randomize
  End If

  For p = 1 To mNumPaths

    mXPaths(p, 0) = 0#
    mShortRatePaths(p, 0) = InitialForward(0#)

    For s = 1 To mSteps
      dtStep = mTimes(s) - mTimes(s - 1)
      tNext = mTimes(s)
      z = RandomNormal()

      xPrev = mXPaths(p, s - 1)
      xNext = EvolveX(xPrev, dtStep, z)

      mXPaths(p, s) = xNext
      mShortRatePaths(p, s) = InitialForward(tNext) + xNext
    Next s

  Next p

  mHasSimulation = True

End Sub

'============================================================
' Time grid table を返す
'============================================================
Public Function GetTimeGrid() As Variant

  Dim arr() As Variant
  Dim s As Long

  CheckInitialized

  ReDim arr(1 To mSteps + 1, 1 To 2)

  arr(1, 1) = "Step"
  arr(1, 2) = "Time"

  For s = 0 To mSteps
    arr(s + 2, 1) = s
    arr(s + 2, 2) = mTimes(s)
  Next s

  GetTimeGrid = arr

End Function

'============================================================
' x path table を返す
'
' 行：Step
' 列：Path
'============================================================
Public Function GetXPathTable() As Variant

  Dim arr() As Variant
  Dim p As Long
  Dim s As Long

  CheckSimulationReady

  ReDim arr(1 To mSteps + 2, 1 To mNumPaths + 1)

  arr(1, 1) = "Time \ Path"

  For p = 1 To mNumPaths
    arr(1, p + 1) = "Path " & CStr(p)
  Next p

  For s = 0 To mSteps
    arr(s + 2, 1) = mTimes(s)

    For p = 1 To mNumPaths
      arr(s + 2, p + 1) = mXPaths(p, s)
    Next p
  Next s

  GetXPathTable = arr

End Function

'============================================================
' short rate path table を返す
'
' 行：Step
' 列：Path
'============================================================
Public Function GetShortRatePathTable() As Variant

  Dim arr() As Variant
  Dim p As Long
  Dim s As Long

  CheckSimulationReady

  ReDim arr(1 To mSteps + 2, 1 To mNumPaths + 1)

  arr(1, 1) = "Time \ Path"

  For p = 1 To mNumPaths
    arr(1, p + 1) = "Path " & CStr(p)
  Next p

  For s = 0 To mSteps
    arr(s + 2, 1) = mTimes(s)

    For p = 1 To mNumPaths
      arr(s + 2, p + 1) = mShortRatePaths(p, s)
    Next p
  Next s

  GetShortRatePathTable = arr

End Function

'============================================================
' 指定pathのhorizon時点 x を返す
'============================================================
Public Function TerminalX(ByVal in_PathIndex As Long) As Double

  CheckSimulationReady
  CheckPathIndex in_PathIndex

  TerminalX = mXPaths(in_PathIndex, mSteps)

End Function

'============================================================
' 指定pathのhorizon時点 short rate を返す
'============================================================
Public Function TerminalShortRate(ByVal in_PathIndex As Long) As Double

  CheckSimulationReady
  CheckPathIndex in_PathIndex

  TerminalShortRate = mShortRatePaths(in_PathIndex, mSteps)

End Function

'============================================================
' 将来時点 horizon から見た discount factor P(horizon, maturity)
'
' in_MaturityYears:
'   valuation date から見た maturity time
'
' 注意：
'   maturity は horizon より後である必要がある。
'============================================================
Public Function FutureDF( _
  ByVal in_PathIndex As Long, _
  ByVal in_MaturityYears As Double _
) As Double

  Dim h As Double
  Dim xH As Double
  Dim fwdDF As Double
  Dim b As Double

  CheckSimulationReady
  CheckPathIndex in_PathIndex

  h = mHorizonYears

  If in_MaturityYears < h Then
    Err.Raise vbObjectError + 8021, _
      "clsHWSimulator.FutureDF", _
      "MaturityYears must be greater than or equal to horizon."
  End If

  If Abs(in_MaturityYears - h) < EPS_TIME Then
    FutureDF = 1#
    Exit Function
  End If

  xH = TerminalX(in_PathIndex)
  fwdDF = ForwardDF(h, in_MaturityYears)
  b = HW_B(mA, h, in_MaturityYears)

  ' Approximation: P_h(T) = P(0,T)/P(0,h) * exp(-B(h,T) * x_h)
  FutureDF = fwdDF * Exp(-b * xH)

  If FutureDF <= 0# Then
    Err.Raise vbObjectError + 8022, _
      "clsHWSimulator.FutureDF", _
      "Future DF is non-positive."
  End If

End Function

'============================================================
' 将来時点 horizon から見た zero rate
'
' zeroRate = -ln(P(h,T)) / (T-h)
'============================================================
Public Function FutureZeroRate( _
  ByVal in_PathIndex As Long, _
  ByVal in_MaturityYears As Double _
) As Double

  Dim h As Double
  Dim tau As Double
  Dim df As Double

  CheckSimulationReady

  h = mHorizonYears
  tau = in_MaturityYears - h

  If tau <= 0# Then
    Err.Raise vbObjectError + 8031, _
      "clsHWSimulator.FutureZeroRate", _
      "MaturityYears must be greater than horizon."
  End If

  df = FutureDF(in_PathIndex, in_MaturityYears)

  FutureZeroRate = -Log(df) / tau

End Function

'============================================================
' 将来時点 horizon から見た forward rate
'
' F = {P(h,T1)/P(h,T2) - 1} / (T2-T1)
'============================================================
Public Function FutureForwardRate( _
  ByVal in_PathIndex As Long, _
  ByVal in_T1 As Double, _
  ByVal in_T2 As Double _
) As Double

  Dim df1 As Double
  Dim df2 As Double

  CheckSimulationReady

  If in_T2 <= in_T1 Then
    Err.Raise vbObjectError + 8041, _
      "clsHWSimulator.FutureForwardRate", _
      "T2 must be greater than T1."
  End If

  If in_T1 < mHorizonYears Then
    Err.Raise vbObjectError + 8042, _
      "clsHWSimulator.FutureForwardRate", _
      "T1 must be greater than or equal to horizon."
  End If

  df1 = FutureDF(in_PathIndex, in_T1)
  df2 = FutureDF(in_PathIndex, in_T2)

  If df1 <= 0# Or df2 <= 0# Then
    Err.Raise vbObjectError + 8043, _
      "clsHWSimulator.FutureForwardRate", _
      "Future discount factors must be positive."
  End If

  FutureForwardRate = (df1 / df2 - 1#) / (in_T2 - in_T1)

End Function

'============================================================
' 指定maturity gridに対する将来zero rate tableを返す
'
' 行：Path
' 列：Maturity
'============================================================
Public Function GetFutureZeroCurveTable(ByVal in_MaturityYears As Variant) As Variant

  Dim nMat As Long
  Dim mats() As Double
  Dim arr() As Variant
  Dim i As Long
  Dim p As Long

  CheckSimulationReady

  nMat = GetVariantItemCount(in_MaturityYears)

  If nMat <= 0 Then
    Err.Raise vbObjectError + 8051, _
      "clsHWSimulator.GetFutureZeroCurveTable", _
      "Maturity grid is empty."
  End If

  ReDim mats(1 To nMat)

  For i = 1 To nMat
    mats(i) = CDbl(GetVariantItem(in_MaturityYears, i))

    If mats(i) <= mHorizonYears Then
      Err.Raise vbObjectError + 8052, _
        "clsHWSimulator.GetFutureZeroCurveTable", _
        "All maturities must be greater than horizon."
    End If
  Next i

  ReDim arr(1 To mNumPaths + 1, 1 To nMat + 1)

  arr(1, 1) = "Path \ Maturity"

  For i = 1 To nMat
    arr(1, i + 1) = mats(i)
  Next i

  For p = 1 To mNumPaths
    arr(p + 1, 1) = p

    For i = 1 To nMat
      arr(p + 1, i + 1) = FutureZeroRate(p, mats(i))
    Next i
  Next p

  GetFutureZeroCurveTable = arr

End Function

'============================================================
' 指定maturity gridに対する将来DF tableを返す
'
' 行：Path
' 列：Maturity
'============================================================
Public Function GetFutureDFCurveTable(ByVal in_MaturityYears As Variant) As Variant

  Dim nMat As Long
  Dim mats() As Double
  Dim arr() As Variant
  Dim i As Long
  Dim p As Long

  CheckSimulationReady

  nMat = GetVariantItemCount(in_MaturityYears)

  If nMat <= 0 Then
    Err.Raise vbObjectError + 8061, _
      "clsHWSimulator.GetFutureDFCurveTable", _
      "Maturity grid is empty."
  End If

  ReDim mats(1 To nMat)

  For i = 1 To nMat
    mats(i) = CDbl(GetVariantItem(in_MaturityYears, i))

    If mats(i) < mHorizonYears Then
      Err.Raise vbObjectError + 8062, _
        "clsHWSimulator.GetFutureDFCurveTable", _
        "All maturities must be greater than or equal to horizon."
    End If
  Next i

  ReDim arr(1 To mNumPaths + 1, 1 To nMat + 1)

  arr(1, 1) = "Path \ Maturity"

  For i = 1 To nMat
    arr(1, i + 1) = mats(i)
  Next i

  For p = 1 To mNumPaths
    arr(p + 1, 1) = p

    For i = 1 To nMat
      arr(p + 1, i + 1) = FutureDF(p, mats(i))
    Next i
  Next p

  GetFutureDFCurveTable = arr

End Function

'============================================================
' maturity gridごとの分位点zero curveを返す
'
' in_Percentile:
'   0.99, 0.95, 0.5 など
'
' 戻り値：
'   1行目 header
'   2行目 percentile zero rate
'============================================================
Public Function GetFutureZeroPercentileCurve( _
  ByVal in_MaturityYears As Variant, _
  ByVal in_Percentile As Double _
) As Variant

  Dim nMat As Long
  Dim mats() As Double
  Dim values() As Double
  Dim arr() As Variant
  Dim i As Long
  Dim p As Long

  CheckSimulationReady

  If in_Percentile < 0# Or in_Percentile > 1# Then
    Err.Raise vbObjectError + 8071, _
      "clsHWSimulator.GetFutureZeroPercentileCurve", _
      "Percentile must be between 0 and 1."
  End If

  nMat = GetVariantItemCount(in_MaturityYears)
  ReDim mats(1 To nMat)
  ReDim arr(1 To 2, 1 To nMat + 1)

  arr(1, 1) = "Percentile"
  arr(2, 1) = in_Percentile

  For i = 1 To nMat
    mats(i) = CDbl(GetVariantItem(in_MaturityYears, i))

    If mats(i) <= mHorizonYears Then
      Err.Raise vbObjectError + 8072, _
        "clsHWSimulator.GetFutureZeroPercentileCurve", _
        "All maturities must be greater than horizon."
    End If

    ReDim values(1 To mNumPaths)

    For p = 1 To mNumPaths
      values(p) = FutureZeroRate(p, mats(i))
    Next p

    SortDoubleArray values

    arr(1, i + 1) = mats(i)
    arr(2, i + 1) = PercentileFromSorted(values, in_Percentile)
  Next i

  GetFutureZeroPercentileCurve = arr

End Function

'============================================================
' Summary
'============================================================
Public Function Summary() As String

  CheckInitialized

  Summary = "clsHWSimulator" & vbCrLf & _
            "a = " & CStr(mA) & vbCrLf & _
            "sigma = " & CStr(mSigma) & vbCrLf & _
            "horizon = " & CStr(mHorizonYears) & vbCrLf & _
            "dt = " & CStr(mDt) & vbCrLf & _
            "steps = " & CStr(mSteps) & vbCrLf & _
            "numPaths = " & CStr(mNumPaths) & vbCrLf & _
            "hasSimulation = " & CStr(mHasSimulation)

End Function

'============================================================
' 内部処理：time grid 作成
'============================================================
Private Sub BuildTimeGrid()

  Dim s As Long
  Dim t As Double

  ReDim mTimes(0 To mSteps)

  mTimes(0) = 0#

  For s = 1 To mSteps
    t = CDbl(s) * mDt

    If t > mHorizonYears Or s = mSteps Then
      t = mHorizonYears
    End If

    mTimes(s) = t
  Next s

End Sub

'============================================================
' 内部処理：xを1ステップ進める
'============================================================
Private Function EvolveX( _
  ByVal in_X As Double, _
  ByVal in_Dt As Double, _
  ByVal in_Z As Double _
) As Double

  Dim decay As Double
  Dim stdev As Double

  If in_Dt <= 0# Then
    Err.Raise vbObjectError + 8101, _
      "clsHWSimulator.EvolveX", _
      "Dt must be positive."
  End If

  If Abs(mA) < EPS_A Then
    decay = 1#
    stdev = mSigma * Sqr(in_Dt)
  Else
    decay = Exp(-mA * in_Dt)
    stdev = mSigma * Sqr((1# - Exp(-2# * mA * in_Dt)) / (2# * mA))
  End If

  EvolveX = in_X * decay + stdev * in_Z

End Function

'============================================================
' 内部処理：初期瞬間フォワード
'============================================================
Private Function InitialForward(ByVal in_T As Double) As Double

  If in_T < 0# Then
    Err.Raise vbObjectError + 8111, _
      "clsHWSimulator.InitialForward", _
      "T must be non-negative."
  End If

  InitialForward = CDbl(mCurve.InstantaneousForward(in_T))

End Function

'============================================================
' 内部処理：Curve DF_T wrapper
'============================================================
Private Function CurveDF(ByVal in_T As Double) As Double

  If in_T < 0# Then
    Err.Raise vbObjectError + 8121, _
      "clsHWSimulator.CurveDF", _
      "T must be non-negative."
  End If

  CurveDF = CDbl(mCurve.DF_T(in_T))

  If CurveDF <= 0# Then
    Err.Raise vbObjectError + 8122, _
      "clsHWSimulator.CurveDF", _
      "Curve returned non-positive DF."
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
    Err.Raise vbObjectError + 8131, _
      "clsHWSimulator.ForwardDF", _
      "T2 must be greater than or equal to T1."
  End If

  df1 = CurveDF(in_T1)
  df2 = CurveDF(in_T2)

  ForwardDF = df2 / df1

End Function

'============================================================
' 内部処理：Hull-White B(t,T)
'============================================================
Private Function HW_B(ByVal in_a As Double, _
                      ByVal in_StartT As Double, _
                      ByVal in_EndT As Double) As Double

  Dim dt As Double

  If in_EndT < in_StartT Then
    Err.Raise vbObjectError + 7131, _
      "clsHWCalibrator.HW_B", _
      "T must be greater than or equal to t."
  End If

  dt = in_EndT - in_StartT

  If Abs(dt) < EPS_TIME Then
    HW_B = 0#
  ElseIf Abs(in_a) < EPS_A Then
    HW_B = dt
  Else
    HW_B = (1# - Exp(-in_a * dt)) / in_a
  End If

End Function

'============================================================
' 内部処理：標準正規乱数 Box-Muller
'============================================================
Private Function RandomNormal() As Double

  Dim u1 As Double
  Dim u2 As Double

  u1 = Rnd
  u2 = Rnd

  If u1 <= 0# Then u1 = 0.000000000001

  RandomNormal = Sqr(-2# * Log(u1)) * Cos(2# * PI_VALUE * u2)

End Function

'============================================================
' 内部処理：Path index check
'============================================================
Private Sub CheckPathIndex(ByVal in_PathIndex As Long)

  If in_PathIndex < 1 Or in_PathIndex > mNumPaths Then
    Err.Raise vbObjectError + 8151, _
      "clsHWSimulator.CheckPathIndex", _
      "Path index is out of range."
  End If

End Sub

'============================================================
' 内部処理：simulation ready check
'============================================================
Private Sub CheckSimulationReady()

  CheckInitialized

  If Not mHasSimulation Then
    Err.Raise vbObjectError + 8161, _
      "clsHWSimulator", _
      "Simulation has not been run. Call Simulate first."
  End If

End Sub

'============================================================
' 内部処理：初期化チェック
'============================================================
Private Sub CheckInitialized()

  If Not mIsInitialized Then
    Err.Raise vbObjectError + 8171, _
      "clsHWSimulator", _
      "clsHWSimulator is not initialized. Call Init first."
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

'============================================================
' 内部処理：配列ソート 昇順
'============================================================
Private Sub SortDoubleArray(ByRef io_values() As Double)

  Dim i As Long
  Dim j As Long
  Dim tmp As Double

  For i = LBound(io_values) To UBound(io_values) - 1
    For j = i + 1 To UBound(io_values)
      If io_values(j) < io_values(i) Then
        tmp = io_values(i)
        io_values(i) = io_values(j)
        io_values(j) = tmp
      End If
    Next j
  Next i

End Sub

'============================================================
' 内部処理：ソート済み配列からpercentileを線形補間で取得
'============================================================
Private Function PercentileFromSorted( _
  ByRef in_values() As Double, _
  ByVal in_p As Double _
) As Double

  Dim n As Long
  Dim pos As Double
  Dim idxLow As Long
  Dim idxHigh As Long
  Dim w As Double

  n = UBound(in_values) - LBound(in_values) + 1

  If n <= 0 Then
    Err.Raise vbObjectError + 8181, _
      "clsHWSimulator.PercentileFromSorted", _
      "Array is empty."
  End If

  If n = 1 Then
    PercentileFromSorted = in_values(LBound(in_values))
    Exit Function
  End If

  If in_p <= 0# Then
    PercentileFromSorted = in_values(LBound(in_values))
    Exit Function
  End If

  If in_p >= 1# Then
    PercentileFromSorted = in_values(UBound(in_values))
    Exit Function
  End If

  pos = CDbl(LBound(in_values)) + in_p * CDbl(n - 1)
  idxLow = CLng(Int(pos))
  idxHigh = idxLow + 1

  If idxHigh > UBound(in_values) Then
    PercentileFromSorted = in_values(UBound(in_values))
    Exit Function
  End If

  w = pos - CDbl(idxLow)

  PercentileFromSorted = in_values(idxLow) + w * (in_values(idxHigh) - in_values(idxLow))

End Function
