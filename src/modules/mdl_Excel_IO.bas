Option Explicit

'============================================================
' mdl_Excel_IO
'
' Excel input / output helper module.
'
' 目的：
'   ・Excel Range から配列・Variantを取得する
'   ・Excel Range から clsVolSurface を作成する
'   ・クラスが返す Variant 配列を Excel Range へ出力する
'
' 設計方針：
'   ・モデルクラス、キャリブレーター、シミュレーターには
'     Excelシート読込ロジックを混ぜない
'   ・Excel I/O は本モジュールに寄せる
'   ・引数名は入力系を in_ で統一する
'
' 主な依存：
'   ・clsVolSurface
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
Public Function ExcelIO_CreateVolSurfaceFromRanges( _
  ByVal in_ExpiryRange As Range, _
  ByVal in_TenorRange As Range, _
  ByVal in_VolRange As Range, _
  Optional ByVal in_VolType As String = "NORMAL", _
  Optional ByVal in_AllowFlatExtrapolation As Boolean = False _
) As clsVolSurface

  Dim volSurface As clsVolSurface
  Dim expiryYears As Variant
  Dim tenorYears As Variant
  Dim vols As Variant
  Dim expiryCount As Long
  Dim tenorCount As Long

  ExcelIO_ValidateRangeNotNothing in_ExpiryRange, "ExcelIO_CreateVolSurfaceFromRanges", "ExpiryRange"
  ExcelIO_ValidateRangeNotNothing in_TenorRange, "ExcelIO_CreateVolSurfaceFromRanges", "TenorRange"
  ExcelIO_ValidateRangeNotNothing in_VolRange, "ExcelIO_CreateVolSurfaceFromRanges", "VolRange"

  expiryCount = ExcelIO_RangeVectorCount(in_ExpiryRange)
  tenorCount = ExcelIO_RangeVectorCount(in_TenorRange)

  If in_VolRange.Rows.Count <> expiryCount Then
    Err.Raise vbObjectError + 9101, _
      "mdl_Excel_IO.ExcelIO_CreateVolSurfaceFromRanges", _
      "VolRange row count must match ExpiryRange item count."
  End If

  If in_VolRange.Columns.Count <> tenorCount Then
    Err.Raise vbObjectError + 9102, _
      "mdl_Excel_IO.ExcelIO_CreateVolSurfaceFromRanges", _
      "VolRange column count must match TenorRange item count."
  End If

  expiryYears = ExcelIO_ReadVectorFromRange(in_ExpiryRange)
  tenorYears = ExcelIO_ReadVectorFromRange(in_TenorRange)
  vols = ExcelIO_ReadMatrixFromRange(in_VolRange)

  Set volSurface = New clsVolSurface

  volSurface.Init _
    in_ExpiryYears:=expiryYears, _
    in_TenorYears:=tenorYears, _
    in_Vols:=vols, _
    in_VolType:=in_VolType, _
    in_AllowFlatExtrapolation:=in_AllowFlatExtrapolation

  Set ExcelIO_CreateVolSurfaceFromRanges = volSurface

End Function

'============================================================
' Excel Range から1次元配列を作成
'
' 戻り値：
'   1-based 1次元配列
'
' 対応：
'   ・縦方向 Range
'   ・横方向 Range
'   ・1セル Range
'============================================================
Public Function ExcelIO_ReadVectorFromRange( _
  ByVal in_Range As Range _
) As Variant

  Dim arr() As Variant
  Dim i As Long
  Dim n As Long

  ExcelIO_ValidateRangeNotNothing in_Range, "ExcelIO_ReadVectorFromRange", "Range"

  n = ExcelIO_RangeVectorCount(in_Range)
  ReDim arr(1 To n)

  If in_Range.Rows.Count >= in_Range.Columns.Count Then
    For i = 1 To n
      arr(i) = in_Range.Cells(i, 1).Value
    Next i
  Else
    For i = 1 To n
      arr(i) = in_Range.Cells(1, i).Value
    Next i
  End If

  ExcelIO_ReadVectorFromRange = arr

End Function

'============================================================
' Excel Range から2次元配列を作成
'
' 戻り値：
'   1-based 2次元配列
'============================================================
Public Function ExcelIO_ReadMatrixFromRange( _
  ByVal in_Range As Range _
) As Variant

  Dim arr() As Variant
  Dim i As Long
  Dim j As Long
  Dim rowCount As Long
  Dim colCount As Long

  ExcelIO_ValidateRangeNotNothing in_Range, "ExcelIO_ReadMatrixFromRange", "Range"

  rowCount = in_Range.Rows.Count
  colCount = in_Range.Columns.Count

  If rowCount <= 0 Or colCount <= 0 Then
    Err.Raise vbObjectError + 9111, _
      "mdl_Excel_IO.ExcelIO_ReadMatrixFromRange", _
      "Range size is invalid."
  End If

  ReDim arr(1 To rowCount, 1 To colCount)

  For i = 1 To rowCount
    For j = 1 To colCount
      arr(i, j) = in_Range.Cells(i, j).Value
    Next j
  Next i

  ExcelIO_ReadMatrixFromRange = arr

End Function

'============================================================
' Matrix形式の vol table から clsVolSurface を作成
'
' in_SurfaceRange:
'   左上セルは見出し。
'   1行目の2列目以降が Tenor。
'   1列目の2行目以降が Expiry。
'   本体部分が Vol matrix。
'
' 例：
'          1Y    2Y    5Y
'   1Y     vol   vol   vol
'   2Y     vol   vol   vol
'   5Y     vol   vol   vol
'============================================================
Public Function ExcelIO_CreateVolSurfaceFromTableRange( _
  ByVal in_SurfaceRange As Range, _
  Optional ByVal in_VolType As String = "NORMAL", _
  Optional ByVal in_AllowFlatExtrapolation As Boolean = False _
) As clsVolSurface

  Dim expiryRange As Range
  Dim tenorRange As Range
  Dim volRange As Range
  Dim rowCount As Long
  Dim colCount As Long

  ExcelIO_ValidateRangeNotNothing in_SurfaceRange, "ExcelIO_CreateVolSurfaceFromTableRange", "SurfaceRange"

  rowCount = in_SurfaceRange.Rows.Count
  colCount = in_SurfaceRange.Columns.Count

  If rowCount < 2 Or colCount < 2 Then
    Err.Raise vbObjectError + 9121, _
      "mdl_Excel_IO.ExcelIO_CreateVolSurfaceFromTableRange", _
      "SurfaceRange must contain header row, header column, and vol matrix."
  End If

  Set tenorRange = in_SurfaceRange.Cells(1, 2).Resize(1, colCount - 1)
  Set expiryRange = in_SurfaceRange.Cells(2, 1).Resize(rowCount - 1, 1)
  Set volRange = in_SurfaceRange.Cells(2, 2).Resize(rowCount - 1, colCount - 1)

  Set ExcelIO_CreateVolSurfaceFromTableRange = ExcelIO_CreateVolSurfaceFromRanges( _
    in_ExpiryRange:=expiryRange, _
    in_TenorRange:=tenorRange, _
    in_VolRange:=volRange, _
    in_VolType:=in_VolType, _
    in_AllowFlatExtrapolation:=in_AllowFlatExtrapolation _
  )

End Function

'============================================================
' clsVolSurface の surface table を Excel へ出力
'============================================================
Public Sub ExcelIO_WriteVolSurfaceToRange( _
  ByVal in_VolSurface As clsVolSurface, _
  ByVal in_OutputTopLeft As Range _
)

  ExcelIO_ValidateVolSurfaceReady in_VolSurface, "ExcelIO_WriteVolSurfaceToRange"
  ExcelIO_ValidateRangeNotNothing in_OutputTopLeft, "ExcelIO_WriteVolSurfaceToRange", "OutputTopLeft"

  ExcelIO_WriteArrayToRange _
    in_Array:=in_VolSurface.GetSurfaceTable(), _
    in_OutputTopLeft:=in_OutputTopLeft

End Sub

'============================================================
' clsVolSurface の summary を Excel へ出力
'============================================================
Public Sub ExcelIO_WriteVolSurfaceSummaryToRange( _
  ByVal in_VolSurface As clsVolSurface, _
  ByVal in_OutputTopLeft As Range _
)

  Dim arr(1 To 4, 1 To 2) As Variant

  ExcelIO_ValidateVolSurfaceReady in_VolSurface, "ExcelIO_WriteVolSurfaceSummaryToRange"
  ExcelIO_ValidateRangeNotNothing in_OutputTopLeft, "ExcelIO_WriteVolSurfaceSummaryToRange", "OutputTopLeft"

  arr(1, 1) = "VolType"
  arr(1, 2) = in_VolSurface.VolType
  arr(2, 1) = "NumExpiries"
  arr(2, 2) = in_VolSurface.NumExpiries
  arr(3, 1) = "NumTenors"
  arr(3, 2) = in_VolSurface.NumTenors
  arr(4, 1) = "AllowFlatExtrapolation"
  arr(4, 2) = in_VolSurface.AllowFlatExtrapolation

  ExcelIO_WriteArrayToRange _
    in_Array:=arr, _
    in_OutputTopLeft:=in_OutputTopLeft

End Sub

'============================================================
' Variant配列をRangeへ出力
'
' in_ClearOutput:
'   True の場合、出力予定範囲を ClearContents してから書き込む。
'============================================================
Public Sub ExcelIO_WriteArrayToRange( _
  ByVal in_Array As Variant, _
  ByVal in_OutputTopLeft As Range, _
  Optional ByVal in_ClearOutput As Boolean = False _
)

  Dim rowCount As Long
  Dim colCount As Long
  Dim outputRange As Range

  ExcelIO_ValidateRangeNotNothing in_OutputTopLeft, "ExcelIO_WriteArrayToRange", "OutputTopLeft"

  If Not IsArray(in_Array) Then
    If in_ClearOutput Then
      in_OutputTopLeft.ClearContents
    End If

    in_OutputTopLeft.Value = in_Array
    Exit Sub
  End If

  rowCount = ExcelIO_ArrayRowCount(in_Array)
  colCount = ExcelIO_ArrayColumnCount(in_Array)

  Set outputRange = in_OutputTopLeft.Resize(rowCount, colCount)

  If in_ClearOutput Then
    outputRange.ClearContents
  End If

  outputRange.Value = in_Array

End Sub

'============================================================
' Range vector の要素数
'============================================================
Public Function ExcelIO_RangeVectorCount( _
  ByVal in_Range As Range _
) As Long

  ExcelIO_ValidateRangeNotNothing in_Range, "ExcelIO_RangeVectorCount", "Range"

  If in_Range.Rows.Count >= in_Range.Columns.Count Then
    ExcelIO_RangeVectorCount = in_Range.Rows.Count
  Else
    ExcelIO_RangeVectorCount = in_Range.Columns.Count
  End If

End Function

'============================================================
' 2次元配列の行数
'============================================================
Public Function ExcelIO_ArrayRowCount( _
  ByVal in_Array As Variant _
) As Long

  If Not IsArray(in_Array) Then
    ExcelIO_ArrayRowCount = 1
    Exit Function
  End If

  On Error GoTo OneDimensional

  ExcelIO_ArrayRowCount = UBound(in_Array, 1) - LBound(in_Array, 1) + 1
  Exit Function

OneDimensional:
  ExcelIO_ArrayRowCount = UBound(in_Array) - LBound(in_Array) + 1

End Function

'============================================================
' 2次元配列の列数
'============================================================
Public Function ExcelIO_ArrayColumnCount( _
  ByVal in_Array As Variant _
) As Long

  If Not IsArray(in_Array) Then
    ExcelIO_ArrayColumnCount = 1
    Exit Function
  End If

  On Error GoTo OneDimensional

  ExcelIO_ArrayColumnCount = UBound(in_Array, 2) - LBound(in_Array, 2) + 1
  Exit Function

OneDimensional:
  ExcelIO_ArrayColumnCount = 1

End Function

'============================================================
' Range Nothing check
'============================================================
Public Sub ExcelIO_ValidateRangeNotNothing( _
  ByVal in_Range As Range, _
  ByVal in_MethodName As String, _
  ByVal in_ArgumentName As String _
)

  If in_Range Is Nothing Then
    Err.Raise vbObjectError + 9191, _
      "mdl_Excel_IO." & in_MethodName, _
      in_ArgumentName & " is Nothing."
  End If

End Sub

'============================================================
' Object Nothing check
'============================================================
Public Sub ExcelIO_ValidateObjectNotNothing( _
  ByVal in_Object As Object, _
  ByVal in_MethodName As String, _
  ByVal in_ArgumentName As String _
)

  If in_Object Is Nothing Then
    Err.Raise vbObjectError + 9192, _
      "mdl_Excel_IO." & in_MethodName, _
      in_ArgumentName & " is Nothing."
  End If

End Sub

'============================================================
' clsVolSurface check
'============================================================
Public Sub ExcelIO_ValidateVolSurfaceReady( _
  ByVal in_VolSurface As clsVolSurface, _
  ByVal in_MethodName As String _
)

  If in_VolSurface Is Nothing Then
    Err.Raise vbObjectError + 9193, _
      "mdl_Excel_IO." & in_MethodName, _
      "VolSurface is Nothing."
  End If

  If Not in_VolSurface.IsInitialized Then
    Err.Raise vbObjectError + 9194, _
      "mdl_Excel_IO." & in_MethodName, _
      "VolSurface is not initialized."
  End If

End Sub
