
Option Explicit

'============================================================
' clsVolSurface
'
' Generic Volatility Surface
'
' 目的：
'   ・Hull-White 1F 等のモデル入力として利用する
'     汎用ボラティリティ・サーフェスを保持する
'
'   ・Expiry × Tenor の volatility matrix を保持し、
'     任意の Expiry × Tenor に対する補間済み vol を返す
'
'   ・Excel Range には直接依存しない
'     Excelシートからの読み込みは modExcelIO 側で行う
'
' 想定用途：
'   ・Hull-White 1F calibration
'   ・将来の金利モデル calibration
'
' 初期スコープ：
'   ・ATM swaption normal vol surface を主対象とする
'   ・ただしクラス名は汎用的に clsVolSurface とする
'
' 補間方針：
'   ・Tenor方向  ：Variance = Vol^2 で線形補間
'   ・Expiry方向 ：Total Variance = Vol^2 × Expiry で線形補間
'
' VolType：
'   ・"NORMAL"      : normal volatility
'   ・"LOGNORMAL"   : lognormal volatility
'   ・"SHIFTED_LOGNORMAL" などは将来拡張用
'
' 注意：
'   ・初期の Hull-White 1F では "NORMAL" を想定する
'   ・NormalVol メソッドは VolType が "NORMAL" の場合のみ使用可能
'============================================================

Private mExpiries() As Double
Private mTenors() As Double
Private mVols() As Double

Private mNumExpiries As Long
Private mNumTenors As Long

Private mVolType As String
Private mAllowFlatExtrapolation As Boolean
Private mIsInitialized As Boolean

'============================================================
' 初期化
'
' in_ExpiryYears:
'   Expiry grid
'   例：Array(0.25, 0.5, 1, 2, 5)
'
' in_TenorYears:
'   Tenor grid
'   例：Array(1, 2, 5, 10, 20)
'
' in_Vols:
'   Vol matrix
'   行方向：Expiry
'   列方向：Tenor
'
'   例：
'     vols(1,1) = 1Y x 1Y vol
'     vols(1,2) = 1Y x 2Y vol
'
' in_VolType:
'   "NORMAL" を初期想定
'
' in_AllowFlatExtrapolation:
'   False または省略：範囲外はエラー
'   True          ：範囲外は端点固定
'============================================================
Public Sub Init( _
  ByVal in_ExpiryYears As Variant, _
  ByVal in_TenorYears As Variant, _
  ByVal in_Vols As Variant, _
  Optional ByVal in_VolType As String = "NORMAL", _
  Optional ByVal in_AllowFlatExtrapolation As Boolean = False _
)

  Dim i As Long
  Dim j As Long

  mIsInitialized = False

  mNumExpiries = GetVariantItemCount(in_ExpiryYears)
  mNumTenors = GetVariantItemCount(in_TenorYears)

  If mNumExpiries <= 0 Then
    Err.Raise vbObjectError + 6001, _
      "clsVolSurface.Init", _
      "Expiry grid is empty."
  End If

  If mNumTenors <= 0 Then
    Err.Raise vbObjectError + 6002, _
      "clsVolSurface.Init", _
      "Tenor grid is empty."
  End If

  ReDim mExpiries(1 To mNumExpiries)
  ReDim mTenors(1 To mNumTenors)
  ReDim mVols(1 To mNumExpiries, 1 To mNumTenors)

  For i = 1 To mNumExpiries
    mExpiries(i) = CDbl(GetVariantItem(in_ExpiryYears, i))

    If mExpiries(i) <= 0# Then
      Err.Raise vbObjectError + 6003, _
        "clsVolSurface.Init", _
        "Expiry must be positive."
    End If
  Next i

  For j = 1 To mNumTenors
    mTenors(j) = CDbl(GetVariantItem(in_TenorYears, j))

    If mTenors(j) <= 0# Then
      Err.Raise vbObjectError + 6004, _
        "clsVolSurface.Init", _
        "Tenor must be positive."
    End If
  Next j

  ValidateStrictlyIncreasing mExpiries, "Expiry"
  ValidateStrictlyIncreasing mTenors, "Tenor"

  For i = 1 To mNumExpiries
    For j = 1 To mNumTenors
      mVols(i, j) = CDbl(GetMatrixItem(in_Vols, i, j))

      If mVols(i, j) < 0# Then
        Err.Raise vbObjectError + 6005, _
          "clsVolSurface.Init", _
          "Volatility must be non-negative."
      End If
    Next j
  Next i

  mVolType = NormalizeVolType(in_VolType)
  mAllowFlatExtrapolation = in_AllowFlatExtrapolation

  mIsInitialized = True

End Sub

'============================================================
' 初期化済み判定
'============================================================
Public Property Get IsInitialized() As Boolean
  IsInitialized = mIsInitialized
End Property

'============================================================
' Vol Type
'============================================================
Public Property Get VolType() As String
  VolType = mVolType
End Property

'============================================================
' Expiry数
'============================================================
Public Property Get NumExpiries() As Long
  NumExpiries = mNumExpiries
End Property

'============================================================
' Tenor数
'============================================================
Public Property Get NumTenors() As Long
  NumTenors = mNumTenors
End Property

'============================================================
' Flat extrapolation を許容するか
'============================================================
Public Property Get AllowFlatExtrapolation() As Boolean
  AllowFlatExtrapolation = mAllowFlatExtrapolation
End Property

'============================================================
' Expiry grid を返す
'============================================================
Public Function ExpiryAt(ByVal in_Index As Long) As Double

  CheckInitialized

  If in_Index < 1 Or in_Index > mNumExpiries Then
    Err.Raise vbObjectError + 6011, _
      "clsVolSurface.ExpiryAt", _
      "Expiry index is out of range."
  End If

  ExpiryAt = mExpiries(in_Index)

End Function

'============================================================
' Tenor grid を返す
'============================================================
Public Function TenorAt(ByVal in_Index As Long) As Double

  CheckInitialized

  If in_Index < 1 Or in_Index > mNumTenors Then
    Err.Raise vbObjectError + 6012, _
      "clsVolSurface.TenorAt", _
      "Tenor index is out of range."
  End If

  TenorAt = mTenors(in_Index)

End Function

'============================================================
' 元のグリッド上の Vol を返す
'============================================================
Public Function GridVol( _
  ByVal in_ExpiryIndex As Long, _
  ByVal in_TenorIndex As Long _
) As Double

  CheckInitialized

  If in_ExpiryIndex < 1 Or in_ExpiryIndex > mNumExpiries Then
    Err.Raise vbObjectError + 6013, _
      "clsVolSurface.GridVol", _
      "Expiry index is out of range."
  End If

  If in_TenorIndex < 1 Or in_TenorIndex > mNumTenors Then
    Err.Raise vbObjectError + 6014, _
      "clsVolSurface.GridVol", _
      "Tenor index is out of range."
  End If

  GridVol = mVols(in_ExpiryIndex, in_TenorIndex)

End Function

'============================================================
' 補間済み Vol を返す
'
' in_ExpiryYears:
'   オプション満期。年数。
'
' in_TenorYears:
'   原資産スワップ期間。年数。
'
' 戻り値：
'   mVolType に対応する年率vol
'============================================================
Public Function VolByYears( _
  ByVal in_ExpiryYears As Double, _
  ByVal in_TenorYears As Double _
) As Double

  Dim iLow As Long
  Dim iHigh As Long
  Dim jLow As Long
  Dim jHigh As Long

  Dim e1 As Double
  Dim e2 As Double
  Dim t1 As Double
  Dim t2 As Double

  Dim v11 As Double
  Dim v12 As Double
  Dim v21 As Double
  Dim v22 As Double

  Dim varAtLowExpiry As Double
  Dim varAtHighExpiry As Double
  Dim totalVariance As Double

  CheckInitialized

  If in_ExpiryYears <= 0# Then
    Err.Raise vbObjectError + 6021, _
      "clsVolSurface.VolByYears", _
      "Expiry must be positive."
  End If

  If in_TenorYears <= 0# Then
    Err.Raise vbObjectError + 6022, _
      "clsVolSurface.VolByYears", _
      "Tenor must be positive."
  End If

  FindBracket mExpiries, in_ExpiryYears, iLow, iHigh, mAllowFlatExtrapolation
  FindBracket mTenors, in_TenorYears, jLow, jHigh, mAllowFlatExtrapolation

  e1 = mExpiries(iLow)
  e2 = mExpiries(iHigh)
  t1 = mTenors(jLow)
  t2 = mTenors(jHigh)

  v11 = mVols(iLow, jLow)
  v12 = mVols(iLow, jHigh)
  v21 = mVols(iHigh, jLow)
  v22 = mVols(iHigh, jHigh)

  '----------------------------------------
  ' 1. Tenor方向の補間
  '    同一Expiry上で variance = vol^2 を線形補間
  '----------------------------------------
  varAtLowExpiry = InterpVariance(in_TenorYears, t1, t2, v11, v12)
  varAtHighExpiry = InterpVariance(in_TenorYears, t1, t2, v21, v22)

  '----------------------------------------
  ' 2. Expiry方向の補間
  '    total variance = variance × expiry を線形補間
  '----------------------------------------
  totalVariance = InterpTotalVariance( _
    in_ExpiryYears, _
    e1, _
    e2, _
    varAtLowExpiry, _
    varAtHighExpiry _
  )

  If totalVariance < 0# Then
    Err.Raise vbObjectError + 6023, _
      "clsVolSurface.VolByYears", _
      "Interpolated total variance is negative."
  End If

  VolByYears = Sqr(totalVariance / in_ExpiryYears)

End Function

'============================================================
' Normal Vol を返す
'
' Hull-White 1F の初期キャリブレーションでは
' このメソッドを使う想定。
'============================================================
Public Function NormalVol( _
  ByVal in_ExpiryYears As Double, _
  ByVal in_TenorYears As Double _
) As Double

  CheckInitialized

  If UCase$(mVolType) <> "NORMAL" Then
    Err.Raise vbObjectError + 6031, _
      "clsVolSurface.NormalVol", _
      "VolType must be NORMAL."
  End If

  NormalVol = VolByYears(in_ExpiryYears, in_TenorYears)

End Function

'============================================================
' Lognormal Vol を返す
'
' 将来拡張用。
' 現時点では必要に応じて使用する。
'============================================================
Public Function LognormalVol( _
  ByVal in_ExpiryYears As Double, _
  ByVal in_TenorYears As Double _
) As Double

  CheckInitialized

  If UCase$(mVolType) <> "LOGNORMAL" Then
    Err.Raise vbObjectError + 6032, _
      "clsVolSurface.LognormalVol", _
      "VolType must be LOGNORMAL."
  End If

  LognormalVol = VolByYears(in_ExpiryYears, in_TenorYears)

End Function

'============================================================
' Surface table を返す
'
' 行：Expiry
' 列：Tenor
'============================================================
Public Function GetSurfaceTable() As Variant

  Dim arr() As Variant
  Dim i As Long
  Dim j As Long

  CheckInitialized

  ReDim arr(1 To mNumExpiries + 1, 1 To mNumTenors + 1)

  arr(1, 1) = "Expiry \ Tenor"

  For j = 1 To mNumTenors
    arr(1, j + 1) = mTenors(j)
  Next j

  For i = 1 To mNumExpiries
    arr(i + 1, 1) = mExpiries(i)

    For j = 1 To mNumTenors
      arr(i + 1, j + 1) = mVols(i, j)
    Next j
  Next i

  GetSurfaceTable = arr

End Function

'============================================================
' Surface summary を返す
'============================================================
Public Function Summary() As String

  CheckInitialized

  Summary = "clsVolSurface" & vbCrLf & _
            "VolType = " & mVolType & vbCrLf & _
            "NumExpiries = " & CStr(mNumExpiries) & vbCrLf & _
            "NumTenors = " & CStr(mNumTenors) & vbCrLf & _
            "AllowFlatExtrapolation = " & CStr(mAllowFlatExtrapolation)

End Function

'============================================================
' 内部処理：VolTypeの正規化
'============================================================
Private Function NormalizeVolType(ByVal in_VolType As String) As String

  Dim s As String

  s = UCase$(Trim$(in_VolType))
  s = Replace(s, " ", "_")
  s = Replace(s, "-", "_")

  Select Case s
    Case "", "NORMAL", "NORMAL_VOL", "BACHELIER"
      NormalizeVolType = "NORMAL"

    Case "LOGNORMAL", "BLACK", "BLACK_VOL"
      NormalizeVolType = "LOGNORMAL"

    Case "SHIFTED_LOGNORMAL", "SHIFTED_BLACK"
      NormalizeVolType = "SHIFTED_LOGNORMAL"

    Case Else
      Err.Raise vbObjectError + 6041, _
        "clsVolSurface.NormalizeVolType", _
        "Unsupported VolType: " & in_VolType
  End Select

End Function

'============================================================
' 内部処理：Variance補間
'
' in_X  : 補間対象
' in_X1 : 左グリッド
' in_X2 : 右グリッド
' in_V1 : 左Vol
' in_V2 : 右Vol
'
' 戻り値：
'   variance = vol^2
'============================================================
Private Function InterpVariance( _
  ByVal in_X As Double, _
  ByVal in_X1 As Double, _
  ByVal in_X2 As Double, _
  ByVal in_V1 As Double, _
  ByVal in_V2 As Double _
) As Double

  Dim w As Double
  Dim var1 As Double
  Dim var2 As Double

  var1 = in_V1 * in_V1
  var2 = in_V2 * in_V2

  If in_X1 = in_X2 Then
    InterpVariance = var1
    Exit Function
  End If

  w = (in_X - in_X1) / (in_X2 - in_X1)

  InterpVariance = var1 + w * (var2 - var1)

End Function

'============================================================
' 内部処理：Total Variance補間
'
' in_Variance1, in_Variance2 は、
' tenor方向補間後の variance
'
' 戻り値：
'   total variance = vol^2 × expiry
'============================================================
Private Function InterpTotalVariance( _
  ByVal in_Expiry As Double, _
  ByVal in_E1 As Double, _
  ByVal in_E2 As Double, _
  ByVal in_Variance1 As Double, _
  ByVal in_Variance2 As Double _
) As Double

  Dim w As Double
  Dim tv1 As Double
  Dim tv2 As Double

  If in_E1 = in_E2 Then
    InterpTotalVariance = in_Variance1 * in_Expiry
    Exit Function
  End If

  tv1 = in_Variance1 * in_E1
  tv2 = in_Variance2 * in_E2

  w = (in_Expiry - in_E1) / (in_E2 - in_E1)

  InterpTotalVariance = tv1 + w * (tv2 - tv1)

End Function

'============================================================
' 内部処理：Bracket検索
'
' in_AllowFlatExtrapolation = False:
'   範囲外ならエラー
'
' in_AllowFlatExtrapolation = True:
'   範囲外なら端点固定
'============================================================
Private Sub FindBracket( _
  ByRef in_Grid() As Double, _
  ByVal in_X As Double, _
  ByRef out_IdxLow As Long, _
  ByRef out_IdxHigh As Long, _
  ByVal in_AllowFlatExtrapolation As Boolean _
)

  Dim i As Long
  Dim n As Long

  n = UBound(in_Grid)

  If in_X < in_Grid(1) Then
    If in_AllowFlatExtrapolation Then
      out_IdxLow = 1
      out_IdxHigh = 1
      Exit Sub
    Else
      Err.Raise vbObjectError + 6051, _
        "clsVolSurface.FindBracket", _
        "Value is below surface range: " & CStr(in_X)
    End If
  End If

  If in_X > in_Grid(n) Then
    If in_AllowFlatExtrapolation Then
      out_IdxLow = n
      out_IdxHigh = n
      Exit Sub
    Else
      Err.Raise vbObjectError + 6052, _
        "clsVolSurface.FindBracket", _
        "Value is above surface range: " & CStr(in_X)
    End If
  End If

  If in_X = in_Grid(1) Then
    out_IdxLow = 1
    out_IdxHigh = 1
    Exit Sub
  End If

  If in_X = in_Grid(n) Then
    out_IdxLow = n
    out_IdxHigh = n
    Exit Sub
  End If

  For i = 1 To n - 1
    If in_Grid(i) <= in_X And in_X <= in_Grid(i + 1) Then
      out_IdxLow = i
      out_IdxHigh = i + 1
      Exit Sub
    End If
  Next i

  Err.Raise vbObjectError + 6053, _
    "clsVolSurface.FindBracket", _
    "Could not find bracket for value: " & CStr(in_X)

End Sub

'============================================================
' 内部処理：昇順チェック
'============================================================
Private Sub ValidateStrictlyIncreasing( _
  ByRef in_Grid() As Double, _
  ByVal in_Label As String _
)

  Dim i As Long

  For i = LBound(in_Grid) + 1 To UBound(in_Grid)
    If in_Grid(i) <= in_Grid(i - 1) Then
      Err.Raise vbObjectError + 6061, _
        "clsVolSurface.ValidateStrictlyIncreasing", _
        in_Label & " grid must be strictly increasing."
    End If
  Next i

End Sub

'============================================================
' 内部処理：Variant要素数を取得
'
' 1次元配列、縦Range由来配列、横Range由来配列を
' できるだけ吸収する。
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
'
' 対応：
'   ・1次元配列
'   ・縦方向2次元配列
'   ・横方向2次元配列
'   ・単一値
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
' 内部処理：Matrixから(i,j)要素を取得
'
' in_Vols は基本的に2次元配列を想定する。
' 1-based / 0-based の両方に対応する。
'============================================================
Private Function GetMatrixItem( _
  ByVal in_v As Variant, _
  ByVal in_Row As Long, _
  ByVal in_Col As Long _
) As Variant

  On Error GoTo MatrixError

  If Not IsArray(in_v) Then
    Err.Raise vbObjectError + 6071, _
      "clsVolSurface.GetMatrixItem", _
      "Vol matrix must be an array."
  End If

  GetMatrixItem = in_v(LBound(in_v, 1) + in_Row - 1, _
                       LBound(in_v, 2) + in_Col - 1)

  Exit Function

MatrixError:
  Err.Raise vbObjectError + 6072, _
    "clsVolSurface.GetMatrixItem", _
    "Could not read vol matrix item at row=" & CStr(in_Row) & _
    ", col=" & CStr(in_Col) & "."

End Function

'============================================================
' 内部処理：初期化チェック
'============================================================
Private Sub CheckInitialized()

  If Not mIsInitialized Then
    Err.Raise vbObjectError + 6081, _
      "clsVolSurface", _
      "clsVolSurface is not initialized. Call Init first."
  End If

End Sub
