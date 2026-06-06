Option Explicit

'============================================================
' mdl_HullWhiteWorkFlow
'
' Hull-White 1F workflow helper module.
'
' 目的：
'   ・Excel上の入力範囲から clsVolSurface を作成する
'   ・clsHWCalibrator を作成し、固定 a に対して sigma を推定する
'   ・clsHWSimulator を作成し、将来金利カーブのMCシミュレーションを実行する
'   ・結果をExcelシートへ出力する
'
' 設計方針：
'   ・Hull-White本体、キャリブレーター、シミュレーターに
'     Excelシート読込ロジックを混ぜない
'   ・Excel I/O と全体実行の薄い orchestration を本モジュールに集約する
'   ・引数名は入力系を in_ で統一する
'
' 主な依存：
'   ・curve object with DF_T(T), InstantaneousForward(T)
'   ・clsVolSurface
'   ・clsHWCalibrator
'   ・clsHWSimulator
'============================================================

'============================================================
' Excel Range から clsVolSurface を作成
'
' in_ExpiryRange:
'   Expiry grid。縦・横どちらでも可。
'
' in_TenorRange:
'   Tenor grid。縦・横どちらでも可。
'
' in_VolRange:
'   Vol matrix。
'   行方向 = Expiry、列方向 = Tenor。
'
' in_VolType:
'   初期想定は "NORMAL"。
'
' in_AllowFlatExtrapolation:
'   False: 範囲外はエラー
'   True : 範囲外は端点固定
'============================================================
Public Function HW_CreateVolSurfaceFromRanges( _
  ByVal in_ExpiryRange As Range, _
  ByVal in_TenorRange As Range, _
  ByVal in_VolRange As Range, _
  Optional ByVal in_VolType As String = "NORMAL", _
  Optional ByVal in_AllowFlatExtrapolation As Boolean = False _
) As clsVolSurface

  Dim volSurface As clsVolSurface
  Dim expiryCount As Long
  Dim tenorCount As Long

  ValidateRangeNotNothing in_ExpiryRange, "HW_CreateVolSurfaceFromRanges", "ExpiryRange"
  ValidateRangeNotNothing in_TenorRange, "HW_CreateVolSurfaceFromRanges", "TenorRange"
  ValidateRangeNotNothing in_VolRange, "HW_CreateVolSurfaceFromRanges", "VolRange"

  expiryCount = RangeVectorCount(in_ExpiryRange)
  tenorCount = RangeVectorCount(in_TenorRange)

  If in_VolRange.Rows.Count <> expiryCount Then
    Err.Raise vbObjectError + 9001, _
      "mdl_HullWhiteWorkFlow.HW_CreateVolSurfaceFromRanges", _
      "VolRange row count must match ExpiryRange item count."
  End If

  If in_VolRange.Columns.Count <> tenorCount Then
    Err.Raise vbObjectError + 9002, _
      "mdl_HullWhiteWorkFlow.HW_CreateVolSurfaceFromRanges", _
      "VolRange column count must match TenorRange item count."
  End If

  Set volSurface = New clsVolSurface

  volSurface.Init _
    in_ExpiryYears:=in_ExpiryRange.Value, _
    in_TenorYears:=in_TenorRange.Value, _
    in_Vols:=in_VolRange.Value, _
    in_VolType:=in_VolType, _
    in_AllowFlatExtrapolation:=in_AllowFlatExtrapolation

  Set HW_CreateVolSurfaceFromRanges = volSurface

End Function

'============================================================
' 配列・Variant から clsVolSurface を作成
'
' Unit test や、modExcelIO 側で既に配列化済みの場合に使う。
'============================================================
Public Function HW_CreateVolSurfaceFromArrays( _
  ByVal in_ExpiryYears As Variant, _
  ByVal in_TenorYears As Variant, _
  ByVal in_Vols As Variant, _
  Optional ByVal in_VolType As String = "NORMAL", _
  Optional ByVal in_AllowFlatExtrapolation As Boolean = False _
) As clsVolSurface

  Dim volSurface As clsVolSurface

  Set volSurface = New clsVolSurface

  volSurface.Init _
    in_ExpiryYears:=in_ExpiryYears, _
    in_TenorYears:=in_TenorYears, _
    in_Vols:=in_Vols, _
    in_VolType:=in_VolType, _
    in_AllowFlatExtrapolation:=in_AllowFlatExtrapolation

  Set HW_CreateVolSurfaceFromArrays = volSurface

End Function

'============================================================
' clsHWCalibrator を作成
'
' in_Curve:
'   DF_T(T) を持つ curve object。
'
' in_VolSurface:
'   clsVolSurface。
'
' in_QuoteExpiryRange / in_QuoteTenorRange:
'   calibration に使う quote grid。
'   同じ件数である必要がある。
'============================================================
Public Function HW_CreateCalibratorFromRanges( _
  ByVal in_Curve As Object, _
  ByVal in_VolSurface As clsVolSurface, _
  ByVal in_QuoteExpiryRange As Range, _
  ByVal in_QuoteTenorRange As Range, _
  Optional ByVal in_PaymentFrequencyYears As Double = 1# _
) As clsHWCalibrator

  Dim calibrator As clsHWCalibrator

  ValidateObjectNotNothing in_Curve, "HW_CreateCalibratorFromRanges", "Curve"
  ValidateVolSurfaceReady in_VolSurface, "HW_CreateCalibratorFromRanges"
  ValidateRangeNotNothing in_QuoteExpiryRange, "HW_CreateCalibratorFromRanges", "QuoteExpiryRange"
  ValidateRangeNotNothing in_QuoteTenorRange, "HW_CreateCalibratorFromRanges", "QuoteTenorRange"

  If RangeVectorCount(in_QuoteExpiryRange) <> RangeVectorCount(in_QuoteTenorRange) Then
    Err.Raise vbObjectError + 9011, _
      "mdl_HullWhiteWorkFlow.HW_CreateCalibratorFromRanges", _
      "Quote expiry count and tenor count must match."
  End If

  Set calibrator = New clsHWCalibrator

  calibrator.Init _
    in_Curve:=in_Curve, _
    in_VolSurface:=in_VolSurface, _
    in_PaymentFrequencyYears:=in_PaymentFrequencyYears

  calibrator.BuildCalibrationQuotes _
    in_ExpiryYears:=in_QuoteExpiryRange.Value, _
    in_TenorYears:=in_QuoteTenorRange.Value

  Set HW_CreateCalibratorFromRanges = calibrator

End Function

'============================================================
' 配列・Variant から clsHWCalibrator を作成
'============================================================
Public Function HW_CreateCalibratorFromArrays( _
  ByVal in_Curve As Object, _
  ByVal in_VolSurface As clsVolSurface, _
  ByVal in_QuoteExpiryYears As Variant, _
  ByVal in_QuoteTenorYears As Variant, _
  Optional ByVal in_PaymentFrequencyYears As Double = 1# _
) As clsHWCalibrator

  Dim calibrator As clsHWCalibrator

  ValidateObjectNotNothing in_Curve, "HW_CreateCalibratorFromArrays", "Curve"
  ValidateVolSurfaceReady in_VolSurface, "HW_CreateCalibratorFromArrays"

  Set calibrator = New clsHWCalibrator

  calibrator.Init _
    in_Curve:=in_Curve, _
    in_VolSurface:=in_VolSurface, _
    in_PaymentFrequencyYears:=in_PaymentFrequencyYears

  calibrator.BuildCalibrationQuotes _
    in_ExpiryYears:=in_QuoteExpiryYears, _
    in_TenorYears:=in_QuoteTenorYears

  Set HW_CreateCalibratorFromArrays = calibrator

End Function

'============================================================
' 固定 a に対して sigma をキャリブレーション
'
' 戻り値：
'   1行目 header
'   2行目 result
'============================================================
Public Function HW_CalibrateSigmaFixedA( _
  ByVal in_Curve As Object, _
  ByVal in_VolSurface As clsVolSurface, _
  ByVal in_QuoteExpiryYears As Variant, _
  ByVal in_QuoteTenorYears As Variant, _
  ByVal in_FixedA As Double, _
  Optional ByVal in_PaymentFrequencyYears As Double = 1#, _
  Optional ByVal in_SigmaLower As Double = 0.000001, _
  Optional ByVal in_SigmaUpper As Double = 0.05, _
  Optional ByVal in_Tolerance As Double = 0.0000000001, _
  Optional ByVal in_MaxIterations As Long = 100 _
) As Variant

  Dim calibrator As clsHWCalibrator
  Dim sigma As Double
  Dim arr(1 To 2, 1 To 4) As Variant

  Set calibrator = HW_CreateCalibratorFromArrays( _
    in_Curve:=in_Curve, _
    in_VolSurface:=in_VolSurface, _
    in_QuoteExpiryYears:=in_QuoteExpiryYears, _
    in_QuoteTenorYears:=in_QuoteTenorYears, _
    in_PaymentFrequencyYears:=in_PaymentFrequencyYears _
  )

  sigma = calibrator.CalibrateSigmaFixedA( _
    in_a:=in_FixedA, _
    in_SigmaLower:=in_SigmaLower, _
    in_SigmaUpper:=in_SigmaUpper, _
    in_Tolerance:=in_Tolerance, _
    in_MaxIterations:=in_MaxIterations _
  )

  arr(1, 1) = "a"
  arr(1, 2) = "sigma"
  arr(1, 3) = "objective"
  arr(1, 4) = "quote count"

  arr(2, 1) = in_FixedA
  arr(2, 2) = sigma
  arr(2, 3) = calibrator.LastObjective
  arr(2, 4) = calibrator.QuoteCount

  HW_CalibrateSigmaFixedA = arr

End Function

'============================================================
' Excel Range 入力から sigma calibration を実行し、結果を出力
'
' 出力：
'   in_OutputTopLeft から summary
'   その下に calibration report
'============================================================
Public Sub HW_RunCalibrationFixedAToSheet( _
  ByVal in_Curve As Object, _
  ByVal in_ExpiryRange As Range, _
  ByVal in_TenorRange As Range, _
  ByVal in_VolRange As Range, _
  ByVal in_QuoteExpiryRange As Range, _
  ByVal in_QuoteTenorRange As Range, _
  ByVal in_FixedA As Double, _
  ByVal in_OutputTopLeft As Range, _
  Optional ByVal in_VolType As String = "NORMAL", _
  Optional ByVal in_AllowFlatExtrapolation As Boolean = False, _
  Optional ByVal in_PaymentFrequencyYears As Double = 1#, _
  Optional ByVal in_SigmaLower As Double = 0.000001, _
  Optional ByVal in_SigmaUpper As Double = 0.05, _
  Optional ByVal in_Tolerance As Double = 0.0000000001, _
  Optional ByVal in_MaxIterations As Long = 100 _
)

  Dim volSurface As clsVolSurface
  Dim calibrator As clsHWCalibrator
  Dim sigma As Double
  Dim summaryArr(1 To 2, 1 To 4) As Variant
  Dim reportArr As Variant

  ValidateRangeNotNothing in_OutputTopLeft, "HW_RunCalibrationFixedAToSheet", "OutputTopLeft"

  Set volSurface = HW_CreateVolSurfaceFromRanges( _
    in_ExpiryRange:=in_ExpiryRange, _
    in_TenorRange:=in_TenorRange, _
    in_VolRange:=in_VolRange, _
    in_VolType:=in_VolType, _
    in_AllowFlatExtrapolation:=in_AllowFlatExtrapolation _
  )

  Set calibrator = HW_CreateCalibratorFromRanges( _
    in_Curve:=in_Curve, _
    in_VolSurface:=volSurface, _
    in_QuoteExpiryRange:=in_QuoteExpiryRange, _
    in_QuoteTenorRange:=in_QuoteTenorRange, _
    in_PaymentFrequencyYears:=in_PaymentFrequencyYears _
  )

  sigma = calibrator.CalibrateSigmaFixedA( _
    in_a:=in_FixedA, _
    in_SigmaLower:=in_SigmaLower, _
    in_SigmaUpper:=in_SigmaUpper, _
    in_Tolerance:=in_Tolerance, _
    in_MaxIterations:=in_MaxIterations _
  )

  summaryArr(1, 1) = "a"
  summaryArr(1, 2) = "sigma"
  summaryArr(1, 3) = "objective"
  summaryArr(1, 4) = "quote count"

  summaryArr(2, 1) = in_FixedA
  summaryArr(2, 2) = sigma
  summaryArr(2, 3) = calibrator.LastObjective
  summaryArr(2, 4) = calibrator.QuoteCount

  WriteArrayToRange summaryArr, in_OutputTopLeft

  reportArr = calibrator.GetCalibrationReport()
  WriteArrayToRange reportArr, in_OutputTopLeft.Offset(4, 0)

End Sub

'============================================================
' clsHWSimulator を作成
'============================================================
Public Function HW_CreateSimulator( _
  ByVal in_Curve As Object, _
  ByVal in_a As Double, _
  ByVal in_sigma As Double, _
  ByVal in_HorizonYears As Double, _
  Optional ByVal in_Dt As Double = 0.00396825396825397 _
) As clsHWSimulator

  Dim simulator As clsHWSimulator

  ValidateObjectNotNothing in_Curve, "HW_CreateSimulator", "Curve"

  Set simulator = New clsHWSimulator

  simulator.Init _
    in_Curve:=in_Curve, _
    in_a:=in_a, _
    in_sigma:=in_sigma, _
    in_HorizonYears:=in_HorizonYears, _
    in_Dt:=in_Dt

  Set HW_CreateSimulator = simulator

End Function

'============================================================
' MC simulation を実行し、将来 zero rate percentile curve を出力
'
' in_MaturityYearsRange:
'   valuation date から見た maturity grid。
'   各 maturity は horizon より後である必要がある。
'
' in_Percentile:
'   0.99, 0.95, 0.5 など。
'============================================================
Public Sub HW_RunSimulationPercentileToSheet( _
  ByVal in_Curve As Object, _
  ByVal in_a As Double, _
  ByVal in_sigma As Double, _
  ByVal in_HorizonYears As Double, _
  ByVal in_NumPaths As Long, _
  ByVal in_MaturityYearsRange As Range, _
  ByVal in_Percentile As Double, _
  ByVal in_OutputTopLeft As Range, _
  Optional ByVal in_Dt As Double = 0.00396825396825397, _
  Optional ByVal in_Seed As Long = 0 _
)

  Dim simulator As clsHWSimulator
  Dim percentileArr As Variant

  ValidateRangeNotNothing in_MaturityYearsRange, "HW_RunSimulationPercentileToSheet", "MaturityYearsRange"
  ValidateRangeNotNothing in_OutputTopLeft, "HW_RunSimulationPercentileToSheet", "OutputTopLeft"

  Set simulator = HW_CreateSimulator( _
    in_Curve:=in_Curve, _
    in_a:=in_a, _
    in_sigma:=in_sigma, _
    in_HorizonYears:=in_HorizonYears, _
    in_Dt:=in_Dt _
  )

  simulator.Simulate _
    in_NumPaths:=in_NumPaths, _
    in_Seed:=in_Seed

  percentileArr = simulator.GetFutureZeroPercentileCurve( _
    in_MaturityYears:=in_MaturityYearsRange.Value, _
    in_Percentile:=in_Percentile _
  )

  WriteArrayToRange percentileArr, in_OutputTopLeft

End Sub

'============================================================
' MC simulation を実行し、全パスの将来 zero curve table を出力
'============================================================
Public Sub HW_RunSimulationZeroCurvesToSheet( _
  ByVal in_Curve As Object, _
  ByVal in_a As Double, _
  ByVal in_sigma As Double, _
  ByVal in_HorizonYears As Double, _
  ByVal in_NumPaths As Long, _
  ByVal in_MaturityYearsRange As Range, _
  ByVal in_OutputTopLeft As Range, _
  Optional ByVal in_Dt As Double = 0.00396825396825397, _
  Optional ByVal in_Seed As Long = 0 _
)

  Dim simulator As clsHWSimulator
  Dim curveArr As Variant

  ValidateRangeNotNothing in_MaturityYearsRange, "HW_RunSimulationZeroCurvesToSheet", "MaturityYearsRange"
  ValidateRangeNotNothing in_OutputTopLeft, "HW_RunSimulationZeroCurvesToSheet", "OutputTopLeft"

  Set simulator = HW_CreateSimulator( _
    in_Curve:=in_Curve, _
    in_a:=in_a, _
    in_sigma:=in_sigma, _
    in_HorizonYears:=in_HorizonYears, _
    in_Dt:=in_Dt _
  )

  simulator.Simulate _
    in_NumPaths:=in_NumPaths, _
    in_Seed:=in_Seed

  curveArr = simulator.GetFutureZeroCurveTable( _
    in_MaturityYears:=in_MaturityYearsRange.Value _
  )

  WriteArrayToRange curveArr, in_OutputTopLeft

End Sub

'============================================================
' 既存の simulator から percentile curve を出力
'============================================================
Public Sub HW_WriteSimulatorPercentileCurve( _
  ByVal in_Simulator As clsHWSimulator, _
  ByVal in_MaturityYearsRange As Range, _
  ByVal in_Percentile As Double, _
  ByVal in_OutputTopLeft As Range _
)

  Dim percentileArr As Variant

  ValidateSimulatorReady in_Simulator, "HW_WriteSimulatorPercentileCurve"
  ValidateRangeNotNothing in_MaturityYearsRange, "HW_WriteSimulatorPercentileCurve", "MaturityYearsRange"
  ValidateRangeNotNothing in_OutputTopLeft, "HW_WriteSimulatorPercentileCurve", "OutputTopLeft"

  percentileArr = in_Simulator.GetFutureZeroPercentileCurve( _
    in_MaturityYears:=in_MaturityYearsRange.Value, _
    in_Percentile:=in_Percentile _
  )

  WriteArrayToRange percentileArr, in_OutputTopLeft

End Sub

'============================================================
' 既存の calibrator から report を出力
'============================================================
Public Sub HW_WriteCalibrationReport( _
  ByVal in_Calibrator As clsHWCalibrator, _
  ByVal in_OutputTopLeft As Range _
)

  Dim reportArr As Variant

  ValidateCalibratorReady in_Calibrator, "HW_WriteCalibrationReport"
  ValidateRangeNotNothing in_OutputTopLeft, "HW_WriteCalibrationReport", "OutputTopLeft"

  reportArr = in_Calibrator.GetCalibrationReport()
  WriteArrayToRange reportArr, in_OutputTopLeft

End Sub

'============================================================
' Variant配列をRangeへ出力
'============================================================
Public Sub HW_WriteArrayToRange( _
  ByVal in_Array As Variant, _
  ByVal in_OutputTopLeft As Range _
)

  ValidateRangeNotNothing in_OutputTopLeft, "HW_WriteArrayToRange", "OutputTopLeft"
  WriteArrayToRange in_Array, in_OutputTopLeft

End Sub

'============================================================
' 内部処理：配列をRangeへ出力
'============================================================
Private Sub WriteArrayToRange( _
  ByVal in_Array As Variant, _
  ByVal in_OutputTopLeft As Range _
)

  Dim rowCount As Long
  Dim colCount As Long

  If Not IsArray(in_Array) Then
    in_OutputTopLeft.Value = in_Array
    Exit Sub
  End If

  rowCount = UBound(in_Array, 1) - LBound(in_Array, 1) + 1
  colCount = UBound(in_Array, 2) - LBound(in_Array, 2) + 1

  in_OutputTopLeft.Resize(rowCount, colCount).Value = in_Array

End Sub

'============================================================
' 内部処理：Range vector の要素数
'============================================================
Private Function RangeVectorCount(ByVal in_Range As Range) As Long

  If in_Range.Rows.Count >= in_Range.Columns.Count Then
    RangeVectorCount = in_Range.Rows.Count
  Else
    RangeVectorCount = in_Range.Columns.Count
  End If

End Function

'============================================================
' 内部処理：Range Nothing check
'============================================================
Private Sub ValidateRangeNotNothing( _
  ByVal in_Range As Range, _
  ByVal in_MethodName As String, _
  ByVal in_ArgumentName As String _
)

  If in_Range Is Nothing Then
    Err.Raise vbObjectError + 9901, _
      "mdl_HullWhiteWorkFlow." & in_MethodName, _
      in_ArgumentName & " is Nothing."
  End If

End Sub

'============================================================
' 内部処理：Object Nothing check
'============================================================
Private Sub ValidateObjectNotNothing( _
  ByVal in_Object As Object, _
  ByVal in_MethodName As String, _
  ByVal in_ArgumentName As String _
)

  If in_Object Is Nothing Then
    Err.Raise vbObjectError + 9902, _
      "mdl_HullWhiteWorkFlow." & in_MethodName, _
      in_ArgumentName & " is Nothing."
  End If

End Sub

'============================================================
' 内部処理：VolSurface check
'============================================================
Private Sub ValidateVolSurfaceReady( _
  ByVal in_VolSurface As clsVolSurface, _
  ByVal in_MethodName As String _
)

  If in_VolSurface Is Nothing Then
    Err.Raise vbObjectError + 9903, _
      "mdl_HullWhiteWorkFlow." & in_MethodName, _
      "VolSurface is Nothing."
  End If

  If Not in_VolSurface.IsInitialized Then
    Err.Raise vbObjectError + 9904, _
      "mdl_HullWhiteWorkFlow." & in_MethodName, _
      "VolSurface is not initialized."
  End If

End Sub

'============================================================
' 内部処理：Calibrator check
'============================================================
Private Sub ValidateCalibratorReady( _
  ByVal in_Calibrator As clsHWCalibrator, _
  ByVal in_MethodName As String _
)

  If in_Calibrator Is Nothing Then
    Err.Raise vbObjectError + 9905, _
      "mdl_HullWhiteWorkFlow." & in_MethodName, _
      "Calibrator is Nothing."
  End If

  If Not in_Calibrator.IsInitialized Then
    Err.Raise vbObjectError + 9906, _
      "mdl_HullWhiteWorkFlow." & in_MethodName, _
      "Calibrator is not initialized."
  End If

End Sub

'============================================================
' 内部処理：Simulator check
'============================================================
Private Sub ValidateSimulatorReady( _
  ByVal in_Simulator As clsHWSimulator, _
  ByVal in_MethodName As String _
)

  If in_Simulator Is Nothing Then
    Err.Raise vbObjectError + 9907, _
      "mdl_HullWhiteWorkFlow." & in_MethodName, _
      "Simulator is Nothing."
  End If

  If Not in_Simulator.IsInitialized Then
    Err.Raise vbObjectError + 9908, _
      "mdl_HullWhiteWorkFlow." & in_MethodName, _
      "Simulator is not initialized."
  End If

  If Not in_Simulator.HasSimulation Then
    Err.Raise vbObjectError + 9909, _
      "mdl_HullWhiteWorkFlow." & in_MethodName, _
      "Simulation has not been run."
  End If

End Sub
