
Option Explicit

'====================================================
' clsSwaptionVol
'
' ATM Swaption Volatility Matrix を保持し、
' 任意の Expiry × Tenor の ATM Vol を返すクラス。
'
' 主用途：
'   ・CMS convexity adjustment
'   ・CMS spread swap 等で使う ATM swaption vol の取得
'
' 入力マトリックス前提：
'
'        1Y    2Y    5Y    10Y
' 1M     vol   vol   vol   vol
' 3M     vol   vol   vol   vol
' 6M     vol   vol   vol   vol
' 1Y     vol   vol   vol   vol
'
' 補間方針：
'   ・Tenor方向  ：Variance = Vol^2 で線形補間
'   ・Expiry方向 ：Total Variance = Vol^2 × Expiry で線形補間
'
'====================================================

Private mExpiries() As Double
Private mTenors() As Double
Private mVols() As Double

Private mNumExpiries As Long
Private mNumTenors As Long

Private mIsInitialized As Boolean
Private mAllowFlatExtrapolation As Boolean

'====================================================
' 初期化
'
' in_MatrixRange:
'   1行目に tenor、1列目に expiry を持つATM vol matrix。
'
' in_AllowFlatExtrapolation:
'   False または省略：範囲外はエラー
'   True          ：範囲外は端点固定
'====================================================
Public Sub InitializeFromRange( _
  ByVal in_MatrixRange As Range, _
  Optional ByVal in_AllowFlatExtrapolation As Boolean = False)

  Dim i As Long
  Dim j As Long
  
  If in_MatrixRange Is Nothing Then
    Err.Raise vbObjectError + 1000, "clsSwaptionVol.InitializeFromRange", _
          "matrixRange is Nothing."
  End If
  
  If in_MatrixRange.Rows.Count < 2 Or in_MatrixRange.Columns.Count < 2 Then
    Err.Raise vbObjectError + 1001, "clsSwaptionVol.InitializeFromRange", _
          "Vol matrix range is too small."
  End If
  
  mNumExpiries = in_MatrixRange.Rows.Count - 1
  mNumTenors = in_MatrixRange.Columns.Count - 1
  
  ReDim mExpiries(1 To mNumExpiries)
  ReDim mTenors(1 To mNumTenors)
  ReDim mVols(1 To mNumExpiries, 1 To mNumTenors)
  
  mAllowFlatExtrapolation = in_AllowFlatExtrapolation
  
  '----------------------------------------
  ' Tenor headers
  ' 1行目、2列目以降
  '----------------------------------------
  For j = 1 To mNumTenors
    mTenors(j) = SwaptionTenorToYears(CStr(in_MatrixRange.Cells(1, j + 1).Value))
  Next j
  
  '----------------------------------------
  ' Expiry headers and Vol matrix
  ' 1列目、2行目以降
  '----------------------------------------
  For i = 1 To mNumExpiries
  
    mExpiries(i) = SwaptionTenorToYears(CStr(in_MatrixRange.Cells(i + 1, 1).Value))
    
    For j = 1 To mNumTenors
      
      If IsEmpty(in_MatrixRange.Cells(i + 1, j + 1).Value) Then
        Err.Raise vbObjectError + 1002, "clsSwaptionVol.InitializeFromRange", _
              "Vol matrix contains empty cell."
      End If
      
      If Not IsNumeric(in_MatrixRange.Cells(i + 1, j + 1).Value) Then
        Err.Raise vbObjectError + 1003, "clsSwaptionVol.InitializeFromRange", _
              "Vol matrix contains non-numeric value."
      End If
      
      mVols(i, j) = CDbl(in_MatrixRange.Cells(i + 1, j + 1).Value)
      
      If mVols(i, j) < 0 Then
        Err.Raise vbObjectError + 1004, "clsSwaptionVol.InitializeFromRange", _
              "Volatility must be non-negative."
      End If
      
    Next j
  Next i
  
  ValidateStrictlyIncreasing mExpiries, "Expiry"
  ValidateStrictlyIncreasing mTenors, "Tenor"
  
  mIsInitialized = True

End Sub

'====================================================
' Main function
'
' in_ExpiryYears:
'   オプション期間。年数で指定する。
'   例：0.5, 1, 1.5
'
' in_TenorText:
'   原資産スワップ期間。文字列で指定する。
'   例："10Y", "20Y", "18M"
'====================================================
Public Function Vol( _
  ByVal in_ExpiryYears As Double, _
  ByVal in_TenorText As String) As Double

  Dim tenorYears As Double
  
  tenorYears = SwaptionTenorToYears(in_TenorText)
  
  Vol = VolByYears(in_ExpiryYears, tenorYears)

End Function

'====================================================
' 年数指定でVolを取得
'
' in_ExpiryYears:
'   オプション期間。年数で指定する。
'   例：0.5, 1, 1.5
'
' in_TenorYears:
'   原資産スワップ期間。年数で指定する。
'   例：1, 5, 10, 20
'====================================================
Public Function VolByYears( _
  ByVal in_ExpiryYears As Double, _
  ByVal in_TenorYears As Double) As Double

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
  
  If Not mIsInitialized Then
    Err.Raise vbObjectError + 1100, "clsSwaptionVol.VolByYears", _
          "Swaption vol surface is not initialized."
  End If
  
  If in_ExpiryYears <= 0 Then
    Err.Raise vbObjectError + 1101, "clsSwaptionVol.VolByYears", _
          "Expiry must be positive."
  End If
  
  If in_TenorYears <= 0 Then
    Err.Raise vbObjectError + 1102, "clsSwaptionVol.VolByYears", _
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
            varAtHighExpiry)
  
  '----------------------------------------
  ' 3. Volに戻す
  '----------------------------------------
  If totalVariance < 0 Then
    Err.Raise vbObjectError + 1103, "clsSwaptionVol.VolByYears", _
          "Interpolated total variance is negative."
  End If
  
  VolByYears = Sqr(totalVariance / in_ExpiryYears)

End Function

'====================================================
' 初期化済み判定
'====================================================
Public Property Get IsInitialized() As Boolean
  IsInitialized = mIsInitialized
End Property

'====================================================
' Expiry数
'====================================================
Public Property Get NumExpiries() As Long
  NumExpiries = mNumExpiries
End Property

'====================================================
' Tenor数
'====================================================
Public Property Get NumTenors() As Long
  NumTenors = mNumTenors
End Property

'====================================================
' Expiry gridを返す
'====================================================
Public Function ExpiryAt(ByVal in_Index As Long) As Double

  If Not mIsInitialized Then
    Err.Raise vbObjectError + 1200, "clsSwaptionVol.ExpiryAt", _
          "Swaption vol surface is not initialized."
  End If
  
  If in_Index < 1 Or in_Index > mNumExpiries Then
    Err.Raise vbObjectError + 1201, "clsSwaptionVol.ExpiryAt", _
          "Expiry index is out of range."
  End If
  
  ExpiryAt = mExpiries(in_Index)

End Function

'====================================================
' Tenor gridを返す
'====================================================
Public Function TenorAt(ByVal in_Index As Long) As Double

  If Not mIsInitialized Then
    Err.Raise vbObjectError + 1210, "clsSwaptionVol.TenorAt", _
          "Swaption vol surface is not initialized."
  End If
  
  If in_Index < 1 Or in_Index > mNumTenors Then
    Err.Raise vbObjectError + 1211, "clsSwaptionVol.TenorAt", _
          "Tenor index is out of range."
  End If
  
  TenorAt = mTenors(in_Index)

End Function

'====================================================
' 元のグリッド上のVolを返す
'====================================================
Public Function GridVol( _
  ByVal in_ExpiryIndex As Long, _
  ByVal in_TenorIndex As Long) As Double

  If Not mIsInitialized Then
    Err.Raise vbObjectError + 1220, "clsSwaptionVol.GridVol", _
          "Swaption vol surface is not initialized."
  End If
  
  If in_ExpiryIndex < 1 Or in_ExpiryIndex > mNumExpiries Then
    Err.Raise vbObjectError + 1221, "clsSwaptionVol.GridVol", _
          "Expiry index is out of range."
  End If
  
  If in_TenorIndex < 1 Or in_TenorIndex > mNumTenors Then
    Err.Raise vbObjectError + 1222, "clsSwaptionVol.GridVol", _
          "Tenor index is out of range."
  End If
  
  GridVol = mVols(in_ExpiryIndex, in_TenorIndex)

End Function

'====================================================
' Variance補間
'
' in_X  : 補間対象
' in_X1 : 左グリッド
' in_X2 : 右グリッド
' in_V1 : 左Vol
' in_V2 : 右Vol
'
' 戻り値：
'   variance = vol^2
'====================================================
Private Function InterpVariance( _
  ByVal in_X As Double, _
  ByVal in_X1 As Double, _
  ByVal in_X2 As Double, _
  ByVal in_V1 As Double, _
  ByVal in_V2 As Double) As Double

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

'====================================================
' Total Variance補間
'
' in_Variance1, in_Variance2 は、すでに tenor方向補間後の variance
'
' 戻り値：
'   total variance = vol^2 × expiry
'====================================================
Private Function InterpTotalVariance( _
  ByVal in_Expiry As Double, _
  ByVal in_E1 As Double, _
  ByVal in_E2 As Double, _
  ByVal in_Variance1 As Double, _
  ByVal in_Variance2 As Double) As Double

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

'====================================================
' Bracket検索
'
' in_AllowFlatExtrapolation = False:
'   範囲外ならエラー
'
' in_AllowFlatExtrapolation = True:
'   範囲外なら端点固定
'====================================================
Private Sub FindBracket( _
  ByRef in_Grid() As Double, _
  ByVal in_X As Double, _
  ByRef out_IdxLow As Long, _
  ByRef out_IdxHigh As Long, _
  ByVal in_AllowFlatExtrapolation As Boolean)

  Dim i As Long
  Dim n As Long
  
  n = UBound(in_Grid)
  
  If in_X < in_Grid(1) Then
    If in_AllowFlatExtrapolation Then
      out_IdxLow = 1
      out_IdxHigh = 1
      Exit Sub
    Else
      Err.Raise vbObjectError + 1300, "clsSwaptionVol.FindBracket", _
            "Value is below surface range: " & CStr(in_X)
    End If
  End If
  
  If in_X > in_Grid(n) Then
    If in_AllowFlatExtrapolation Then
      out_IdxLow = n
      out_IdxHigh = n
      Exit Sub
    Else
      Err.Raise vbObjectError + 1301, "clsSwaptionVol.FindBracket", _
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
  
  Err.Raise vbObjectError + 1302, "clsSwaptionVol.FindBracket", _
        "Could not find bracket for value: " & CStr(in_X)

End Sub

'====================================================
' 昇順チェック
'====================================================
Private Sub ValidateStrictlyIncreasing( _
  ByRef in_Grid() As Double, _
  ByVal in_Label As String)

  Dim i As Long
  
  For i = LBound(in_Grid) + 1 To UBound(in_Grid)
    If in_Grid(i) <= in_Grid(i - 1) Then
      Err.Raise vbObjectError + 1400, "clsSwaptionVol.ValidateStrictlyIncreasing", _
            in_Label & " grid must be strictly increasing."
    End If
  Next i

End Sub

