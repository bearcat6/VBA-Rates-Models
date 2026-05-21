
Option Explicit

'====================================================
' mdl_SwaptionVol
'
' clsSwaptionVol 用の標準モジュール。
'
' 役割：
'   ・Tenor / Expiry 文字列の年数変換
'   ・Excel関数ラッパー
'   ・テスト用プロシージャ
'
'====================================================

'====================================================
' Tenor文字列を年数に変換
'
' 対応例：
'   "1D"   -> 1 / 365
'   "1W"   -> 1 / 52
'   "1M"   -> 1 / 12
'   "3M"   -> 0.25
'   "6M"   -> 0.5
'   "1Y"   -> 1
'   "10Y"  -> 10
'   "18M"  -> 1.5
'   "1.5Y" -> 1.5
'
' 注意：
'   本関数は、期間の年数換算のみを行う。
'   日付計算、営業日調整、Actual/365等の日数計算は行わない。
'====================================================
Public Function SwaptionTenorToYears(ByVal in_TenorText As String) As Double

  Dim s As String
  Dim unitText As String
  Dim numText As String
  Dim n As Double
  
  s = UCase$(Trim$(in_TenorText))
  
  If Len(s) = 0 Then
    Err.Raise vbObjectError + 2000, "mdl_SwaptionVol.SwaptionTenorToYears", _
          "Tenor text is empty."
  End If
  
  ' 数値だけ渡された場合は年数とみなす
  If IsNumeric(s) Then
    SwaptionTenorToYears = CDbl(s)
    Exit Function
  End If
  
  If Len(s) < 2 Then
    Err.Raise vbObjectError + 2001, "mdl_SwaptionVol.SwaptionTenorToYears", _
          "Invalid tenor text: " & in_TenorText
  End If
  
  unitText = Right$(s, 1)
  numText = Left$(s, Len(s) - 1)
  
  If Not IsNumeric(numText) Then
    Err.Raise vbObjectError + 2002, "mdl_SwaptionVol.SwaptionTenorToYears", _
          "Invalid tenor number: " & in_TenorText
  End If
  
  n = CDbl(numText)
  
  If n <= 0 Then
    Err.Raise vbObjectError + 2003, "mdl_SwaptionVol.SwaptionTenorToYears", _
          "Tenor must be positive: " & in_TenorText
  End If
  
  Select Case unitText
    Case "D"
      SwaptionTenorToYears = n / 365#
      
    Case "W"
      SwaptionTenorToYears = n / 52#
      
    Case "M"
      SwaptionTenorToYears = n / 12#
      
    Case "Y"
      SwaptionTenorToYears = n
      
    Case Else
      Err.Raise vbObjectError + 2004, "mdl_SwaptionVol.SwaptionTenorToYears", _
            "Invalid tenor unit: " & in_TenorText
  End Select

End Function

'====================================================
' Excel関数：
' ATMスワップションボラを返す
'
' 使用例：
'   =ATM_SWAPTION_VOL(1.5,"10Y",SwaptionVol!A1:F7)
'   =ATM_SWAPTION_VOL(1.5,"10Y",SwaptionVol!A1:F7,TRUE)
'
' in_ExpiryYears:
'   オプション期間。年数で指定する。
'   例：0.5, 1, 1.5
'
' in_TenorText:
'   原資産スワップ期間。文字列で指定する。
'   例："1Y", "5Y", "10Y", "20Y"
'
' in_MatrixRange:
'   1行目に tenor、1列目に expiry を持つATM vol matrix。
'
' in_AllowFlatExtrapolation:
'   False または省略：範囲外はエラー
'   True          ：範囲外は端点固定
'====================================================
Public Function ATM_SWAPTION_VOL( _
  ByVal in_ExpiryYears As Double, _
  ByVal in_TenorText As String, _
  ByVal in_MatrixRange As Range, _
  Optional ByVal in_AllowFlatExtrapolation As Boolean = False) As Variant

  On Error GoTo ErrHandler
  
  Dim cSurf As clsSwaptionVol
  
  Set cSurf = New clsSwaptionVol
  
  cSurf.InitializeFromRange in_MatrixRange, in_AllowFlatExtrapolation
  
  ATM_SWAPTION_VOL = cSurf.Vol(in_ExpiryYears, in_TenorText)
  
  Exit Function

ErrHandler:
  ATM_SWAPTION_VOL = CVErr(xlErrValue)

End Function

'====================================================
' Excel関数：
' Expiryも文字列で指定できる版
'
' 使用例：
'   =ATM_SWAPTION_VOL_TEXT("18M","10Y",SwaptionVol!A1:F7)
'   =ATM_SWAPTION_VOL_TEXT("1.5Y","10Y",SwaptionVol!A1:F7)
'   =ATM_SWAPTION_VOL_TEXT("18M","10Y",SwaptionVol!A1:F7,TRUE)
'
' in_ExpiryText:
'   オプション期間。文字列で指定する。
'   例："6M", "1Y", "18M", "1.5Y"
'
' in_TenorText:
'   原資産スワップ期間。文字列で指定する。
'   例："1Y", "5Y", "10Y", "20Y"
'
' in_MatrixRange:
'   1行目に tenor、1列目に expiry を持つATM vol matrix。
'
' in_AllowFlatExtrapolation:
'   False または省略：範囲外はエラー
'   True          ：範囲外は端点固定
'====================================================
Public Function ATM_SWAPTION_VOL_TEXT( _
  ByVal in_ExpiryText As String, _
  ByVal in_TenorText As String, _
  ByVal in_MatrixRange As Range, _
  Optional ByVal in_AllowFlatExtrapolation As Boolean = False) As Variant

  On Error GoTo ErrHandler
  
  Dim cSurf As clsSwaptionVol
  Dim expiryYears As Double
  
  expiryYears = SwaptionTenorToYears(in_ExpiryText)
  
  Set cSurf = New clsSwaptionVol
  
  cSurf.InitializeFromRange in_MatrixRange, in_AllowFlatExtrapolation
  
  ATM_SWAPTION_VOL_TEXT = cSurf.Vol(expiryYears, in_TenorText)
  
  Exit Function

ErrHandler:
  ATM_SWAPTION_VOL_TEXT = CVErr(xlErrValue)

End Function

'====================================================
' Excel関数：
' Tenor文字列を年数に変換して返す確認用関数
'
' 使用例：
'   =SWAPTION_TENOR_TO_YEARS("18M")
'====================================================
Public Function SWAPTION_TENOR_TO_YEARS(ByVal in_TenorText As String) As Variant

  On Error GoTo ErrHandler
  
  SWAPTION_TENOR_TO_YEARS = SwaptionTenorToYears(in_TenorText)
  
  Exit Function

ErrHandler:
  SWAPTION_TENOR_TO_YEARS = CVErr(xlErrValue)

End Function

'====================================================
' VBAテスト用
'
' 前提：
'   SwaptionVol シートの A1:F7 にボラマトリックスがある
'====================================================
Public Sub Test_clsSwaptionVol()

  Dim cSurf As clsSwaptionVol
  Dim ws As Worksheet
  Dim rng As Range
  
  Dim v1 As Double
  Dim v2 As Double
  Dim v3 As Double
  
  Set ws = ThisWorkbook.Worksheets("SwaptionVol")
  Set rng = ws.Range("A1:F7")
  
  Set cSurf = New clsSwaptionVol
  
  ' 第2引数 False：
  ' 範囲外の場合はエラーにする
  cSurf.InitializeFromRange rng, False
  
  v1 = cSurf.Vol(1#, "10Y")
  v2 = cSurf.Vol(1.5, "10Y")
  v3 = cSurf.Vol(2#, "20Y")
  
  Debug.Print "1Y x 10Y Vol  = "; Format(v1, "0.0000%")
  Debug.Print "1.5Y x 10Y Vol= "; Format(v2, "0.0000%")
  Debug.Print "2Y x 20Y Vol  = "; Format(v3, "0.0000%")

End Sub

'====================================================
' 外挿ありのテスト
'
' 例えば、マトリックスが最大 20Y tenor までしかないとき、
' "30Y" を指定したら 20Y のボラを使う。
'====================================================
Public Sub Test_clsSwaptionVol_WithFlatExtrapolation()

  Dim cSurf As clsSwaptionVol
  Dim ws As Worksheet
  Dim rng As Range
  
  Dim v As Double
  
  Set ws = ThisWorkbook.Worksheets("SwaptionVol")
  Set rng = ws.Range("A1:F7")
  
  Set cSurf = New clsSwaptionVol
  
  ' 第2引数 True：
  ' 範囲外は端点固定
  cSurf.InitializeFromRange rng, True
  
  v = cSurf.Vol(10#, "30Y")
  
  Debug.Print "10Y x 30Y Vol with flat extrapolation = "; Format(v, "0.0000%")

End Sub

