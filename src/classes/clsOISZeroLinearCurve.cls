Option Explicit

Implements clsDiscountCurve

'============================================================
' clsOISZeroLinearCurve
'
' JPY OIS Zero Rate Linear Interpolation Curve
'
' 目的：
'   ・評価日から各満期までの連続複利ゼロレートを保持する
'   ・ゼロレートを評価日からの年数に対して直線補間する
'   ・Discount Factor:
'
'       DF(0,T) = Exp( - z(T) * T )
'
'     ここで z(T) は連続複利ゼロレート
'
' 前提：
'   ・入力レートは連続複利ゼロレートとする
'   ・ForwardRate は DF 比から算出する期間フォワードレート
'   ・ForwardRate は ACT/365 ベースの単利年率とする
'
' 依存：
'   ・clsDiscountCurve
'   ・clsHolidayCalendar
'   ・mdl_BusinessDay
'   ・mdl_Common / YearFracAct365F
'============================================================

Private mValuationDate As Date
Private mSpotDate As Date
Private mCalendar As clsHolidayCalendar
Private mBusinessDayConvention As String
Private mPaymentLag As Long

Private mPointCount As Long
Private mPointDate() As Date
Private mPointZeroRate() As Double

Private mIsInitialized As Boolean

'============================================================
' 初期化
'============================================================
Public Sub Init( _
  ByVal in_ValuationDate As Date, _
  Optional ByVal in_Calendar As clsHolidayCalendar = Nothing, _
  Optional ByVal in_spotLag As Long = 2, _
  Optional ByVal in_PaymentLag As Long = 2, _
  Optional ByVal in_BusinessDayConvention As String = "MF" _
)

  If in_spotLag < 0 Then
    Err.Raise vbObjectError + 4001, _
      "clsOISZeroLinearCurve.Init", _
      "SpotLag は0以上にしてください。"
  End If

  If in_PaymentLag < 0 Then
    Err.Raise vbObjectError + 4002, _
      "clsOISZeroLinearCurve.Init", _
      "PaymentLag は0以上にしてください。"
  End If

  mValuationDate = DateValue(in_ValuationDate)
  Set mCalendar = in_Calendar

  mSpotDate = GetWorkday(mValuationDate, in_spotLag, mCalendar)
  mPaymentLag = in_PaymentLag
  mBusinessDayConvention = in_BusinessDayConvention

  ClearPoints

  mIsInitialized = True

End Sub

'============================================================
' Property
'============================================================
Public Property Get ValuationDate() As Date
  ValuationDate = mValuationDate
End Property

Public Property Get SpotDate() As Date
  SpotDate = mSpotDate
End Property

Public Property Get PointCount() As Long
  PointCount = mPointCount
End Property

Public Property Get Calendar() As clsHolidayCalendar
  Set Calendar = mCalendar
End Property

Public Property Set Calendar(ByVal in_Calendar As clsHolidayCalendar)
  Set mCalendar = in_Calendar
End Property

Public Property Get PaymentLag() As Long
  PaymentLag = mPaymentLag
End Property

Public Property Get BusinessDayConvention() As String
  BusinessDayConvention = mBusinessDayConvention
End Property

'============================================================
' ゼロレート点をクリア
'============================================================
Public Sub ClearPoints()

  mPointCount = 0

  Erase mPointDate
  Erase mPointZeroRate

End Sub

'============================================================
' ゼロレート点を追加
'
' in_ZeroRateCont は連続複利ゼロレート
'============================================================
Public Sub AddPoint( _
  ByVal in_TargetDate As Date, _
  ByVal in_ZeroRateCont As Double _
)

  CheckInitialized

  in_TargetDate = DateValue(in_TargetDate)

  If in_TargetDate <= mValuationDate Then
    Err.Raise vbObjectError + 4003, _
      "clsOISZeroLinearCurve.AddPoint", _
      "TargetDate は評価日より後の日付にしてください。"
  End If

  If mPointCount > 0 Then
    If in_TargetDate <= mPointDate(mPointCount) Then
      Err.Raise vbObjectError + 4004, _
        "clsOISZeroLinearCurve.AddPoint", _
        "ゼロレート点は日付昇順で追加してください。"
    End If
  End If

  mPointCount = mPointCount + 1

  ReDim Preserve mPointDate(1 To mPointCount)
  ReDim Preserve mPointZeroRate(1 To mPointCount)

  mPointDate(mPointCount) = in_TargetDate
  mPointZeroRate(mPointCount) = in_ZeroRateCont

End Sub

'============================================================
' tenor / zero rate 配列からカーブを構築
'
' in_Tenors:
'   "ON", "O/N", "1M", "2M", "3M", "6M", "1Y", "2Y", ...
'
' in_ZeroRates:
'   連続複利ゼロレート
'
' 注意：
'   本関数は OIS quote からのブートストラップではない。
'   入力されたレートをゼロレートとして扱う。
'============================================================
Public Sub BuildFromZeroRates( _
  ByVal in_Tenors As Variant, _
  ByVal in_ZeroRates As Variant _
)

  Dim quoteDates() As Date
  Dim quoteRates() As Double
  Dim quoteCount As Long
  Dim i As Long

  CheckInitialized

  ClearPoints

  ExtractAndSortZeroRateQuotes _
    in_Tenors, _
    in_ZeroRates, _
    quoteDates, _
    quoteRates, _
    quoteCount

  If quoteCount = 0 Then
    Err.Raise vbObjectError + 4005, _
      "clsOISZeroLinearCurve.BuildFromZeroRates", _
      "有効なゼロレート点がありません。"
  End If

  For i = 1 To quoteCount
    AddPoint quoteDates(i), quoteRates(i)
  Next i

End Sub

'============================================================
' DF
'
' targetDate = valuationDate なら 1
' targetDate < valuationDate ならエラー
' targetDate > 最終点 + PaymentLag ならエラー
'============================================================
Public Function DF(ByVal in_TargetDate As Date) As Double

  Dim z As Double
  Dim t As Double
  Dim extrapolationLimitDate As Date

  CheckInitialized

  in_TargetDate = DateValue(in_TargetDate)

  If in_TargetDate = mValuationDate Then
    DF = 1#
    Exit Function
  End If

  If in_TargetDate < mValuationDate Then
    Err.Raise vbObjectError + 4006, _
      "clsOISZeroLinearCurve.DF", _
      "評価日より前の日付に対するDFは計算できません。"
  End If

  If mPointCount = 0 Then
    Err.Raise vbObjectError + 4007, _
      "clsOISZeroLinearCurve.DF", _
      "ゼロレート点が設定されていません。"
  End If

  extrapolationLimitDate = GetWorkday(mPointDate(mPointCount), mPaymentLag, mCalendar)

  If in_TargetDate > extrapolationLimitDate Then
    Err.Raise vbObjectError + 4008, _
      "clsOISZeroLinearCurve.DF", _
      "targetDate がカーブ最終日および支払ラグ許容範囲を超えています。"
  End If

  t = YearFracAct365F(mValuationDate, in_TargetDate)
  z = ZeroRateCont(in_TargetDate)

  DF = Exp(-z * t)

End Function

'============================================================
' 連続複利ゼロレート
'
' 評価日からの ACT/365 年数に対して直線補間する。
'============================================================
Public Function ZeroRateCont(ByVal in_TargetDate As Date) As Double

  Dim i As Long
  Dim t As Double
  Dim t1 As Double
  Dim t2 As Double
  Dim z1 As Double
  Dim z2 As Double
  Dim extrapolationLimitDate As Date

  CheckInitialized

  in_TargetDate = DateValue(in_TargetDate)

  If in_TargetDate <= mValuationDate Then
    Err.Raise vbObjectError + 4009, _
      "clsOISZeroLinearCurve.ZeroRateCont", _
      "targetDate は評価日より後の日付にしてください。"
  End If

  If mPointCount = 0 Then
    Err.Raise vbObjectError + 4010, _
      "clsOISZeroLinearCurve.ZeroRateCont", _
      "ゼロレート点が設定されていません。"
  End If

  ' 最初の点より手前は、最初のゼロレートを横置きする
  If in_TargetDate <= mPointDate(1) Then
    ZeroRateCont = mPointZeroRate(1)
    Exit Function
  End If

  ' 最後の点より後は、PaymentLag 許容範囲内に限り最後のゼロレートを横置きする
  If in_TargetDate >= mPointDate(mPointCount) Then

    extrapolationLimitDate = GetWorkday(mPointDate(mPointCount), mPaymentLag, mCalendar)

    If in_TargetDate > extrapolationLimitDate Then
      Err.Raise vbObjectError + 4017, _
        "clsOISZeroLinearCurve.ZeroRateCont", _
        "targetDate がカーブ最終日および支払ラグ許容範囲を超えています。"
    End If

    ZeroRateCont = mPointZeroRate(mPointCount)
    Exit Function

  End If

  t = YearFracAct365F(mValuationDate, in_TargetDate)

  For i = 1 To mPointCount - 1

    If in_TargetDate >= mPointDate(i) And in_TargetDate <= mPointDate(i + 1) Then

      t1 = YearFracAct365F(mValuationDate, mPointDate(i))
      t2 = YearFracAct365F(mValuationDate, mPointDate(i + 1))

      z1 = mPointZeroRate(i)
      z2 = mPointZeroRate(i + 1)

      If t2 <= t1 Then
        Err.Raise vbObjectError + 4011, _
          "clsOISZeroLinearCurve.ZeroRateCont", _
          "ゼロレート点の年数が不正です。"
      End If

      ZeroRateCont = z1 + (z2 - z1) * (t - t1) / (t2 - t1)
      Exit Function

    End If

  Next i

  Err.Raise vbObjectError + 4012, _
    "clsOISZeroLinearCurve.ZeroRateCont", _
    "補間対象区間が見つかりませんでした。"

End Function

'============================================================
' 期間フォワードレート
'
' 単純年率として返す。
'
' ForwardRate(S,E)
'   = ( DF(S) / DF(E) - 1 ) / YearFrac(S,E)
'============================================================
Public Function ForwardRate( _
  ByVal in_StartDate As Date, _
  ByVal in_EndDate As Date _
) As Double

  Dim dfStart As Double
  Dim dfEnd As Double
  Dim yf As Double

  CheckInitialized

  in_StartDate = DateValue(in_StartDate)
  in_EndDate = DateValue(in_EndDate)

  If in_EndDate <= in_StartDate Then
    Err.Raise vbObjectError + 4013, _
      "clsOISZeroLinearCurve.ForwardRate", _
      "EndDate は StartDate より後の日付にしてください。"
  End If

  dfStart = DF(in_StartDate)
  dfEnd = DF(in_EndDate)
  yf = YearFracAct365F(in_StartDate, in_EndDate)

  ForwardRate = (dfStart / dfEnd - 1#) / yf

End Function

'============================================================
' ゼロレート点表を返す
'============================================================
Public Function GetPointTable() As Variant

  Dim arr() As Variant
  Dim i As Long

  CheckInitialized

  If mPointCount = 0 Then
    Err.Raise vbObjectError + 4014, _
      "clsOISZeroLinearCurve.GetPointTable", _
      "ゼロレート点がありません。"
  End If

  ReDim arr(1 To mPointCount + 1, 1 To 5)

  arr(1, 1) = "Point No"
  arr(1, 2) = "Date"
  arr(1, 3) = "YearFrac"
  arr(1, 4) = "Zero Rate Cont"
  arr(1, 5) = "DF"

  For i = 1 To mPointCount
    arr(i + 1, 1) = i
    arr(i + 1, 2) = mPointDate(i)
    arr(i + 1, 3) = YearFracAct365F(mValuationDate, mPointDate(i))
    arr(i + 1, 4) = mPointZeroRate(i)
    arr(i + 1, 5) = DF(mPointDate(i))
  Next i

  GetPointTable = arr

End Function

'============================================================
' ゼロレート点表をシートに出力
'============================================================
Public Sub ExportPointsToSheet( _
  ByVal in_SheetName As String, _
  ByVal in_StartCell As String _
)

  Dim ws As Worksheet
  Dim arr As Variant

  Set ws = ThisWorkbook.Worksheets(in_SheetName)

  arr = GetPointTable()

  ws.Range(in_StartCell).Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr

End Sub

'============================================================
' 内部処理：Quote抽出・ソート
'============================================================
Private Sub ExtractAndSortZeroRateQuotes( _
  ByVal in_Tenors As Variant, _
  ByVal in_ZeroRates As Variant, _
  ByRef out_quoteDates() As Date, _
  ByRef out_quoteRates() As Double, _
  ByRef out_quoteCount As Long _
)

  Dim n As Long
  Dim i As Long
  Dim tenor As String
  Dim rate As Double
  Dim targetDate As Date

  n = GetVariantItemCount(in_Tenors)

  If GetVariantItemCount(in_ZeroRates) <> n Then
    Err.Raise vbObjectError + 4015, _
      "clsOISZeroLinearCurve.ExtractAndSortZeroRateQuotes", _
      "Tenor配列とRate配列の要素数が一致していません。"
  End If

  out_quoteCount = 0

  For i = 1 To n

    tenor = CStr(GetVariantItem(in_Tenors, i))

    If Trim$(tenor) <> "" Then

      rate = CDbl(GetVariantItem(in_ZeroRates, i))
      targetDate = TenorToTargetDate(tenor)

      out_quoteCount = out_quoteCount + 1

      ReDim Preserve out_quoteDates(1 To out_quoteCount)
      ReDim Preserve out_quoteRates(1 To out_quoteCount)

      out_quoteDates(out_quoteCount) = targetDate
      out_quoteRates(out_quoteCount) = rate

    End If

  Next i

  If out_quoteCount > 1 Then
    SortQuotesByDate out_quoteDates, out_quoteRates, out_quoteCount
    RemoveDuplicateQuoteDates out_quoteDates, out_quoteRates, out_quoteCount
  End If

End Sub

'============================================================
' 内部処理：テナーを対象日付に変換
'============================================================
Private Function TenorToTargetDate(ByVal in_Tenor As String) As Date

  Dim s As String

  s = UCase$(Trim$(in_Tenor))
  s = Replace(s, " ", "")

  Select Case s
    Case "ON", "O/N"
      TenorToTargetDate = GetWorkday(mValuationDate, 1, mCalendar)

    Case Else
      TenorToTargetDate = GetBusinessDay( _
        mSpotDate, _
        s, _
        mBusinessDayConvention, _
        mCalendar _
      )
  End Select

End Function

'============================================================
' 内部処理：Quoteを日付順にソート
'============================================================
Private Sub SortQuotesByDate( _
  ByRef io_dates() As Date, _
  ByRef io_rates() As Double, _
  ByVal in_count As Long _
)

  Dim i As Long
  Dim j As Long
  Dim tmpD As Date
  Dim tmpR As Double

  For i = 1 To in_count - 1
    For j = i + 1 To in_count
      If io_dates(j) < io_dates(i) Then

        tmpD = io_dates(i)
        io_dates(i) = io_dates(j)
        io_dates(j) = tmpD

        tmpR = io_rates(i)
        io_rates(i) = io_rates(j)
        io_rates(j) = tmpR

      End If
    Next j
  Next i

End Sub

'============================================================
' 内部処理：同じ日付のQuoteがある場合、後勝ちで1つにする
'============================================================
Private Sub RemoveDuplicateQuoteDates( _
  ByRef io_dates() As Date, _
  ByRef io_rates() As Double, _
  ByRef io_count As Long _
)

  Dim newDates() As Date
  Dim newRates() As Double
  Dim newCount As Long
  Dim i As Long

  newCount = 0

  For i = 1 To io_count

    If i < io_count Then
      If io_dates(i) = io_dates(i + 1) Then
        ' 後勝ちにするため、ここではスキップ
      Else
        newCount = newCount + 1
        ReDim Preserve newDates(1 To newCount)
        ReDim Preserve newRates(1 To newCount)
        newDates(newCount) = io_dates(i)
        newRates(newCount) = io_rates(i)
      End If
    Else
      newCount = newCount + 1
      ReDim Preserve newDates(1 To newCount)
      ReDim Preserve newRates(1 To newCount)
      newDates(newCount) = io_dates(i)
      newRates(newCount) = io_rates(i)
    End If

  Next i

  io_count = newCount
  io_dates = newDates
  io_rates = newRates

End Sub

'============================================================
' 内部処理：配列要素数を取得
'============================================================
Private Function GetVariantItemCount(ByVal in_v As Variant) As Long

  On Error GoTo SingleValue

  If IsArray(in_v) Then
    GetVariantItemCount = UBound(in_v, 1) - LBound(in_v, 1) + 1
  Else
    GetVariantItemCount = 1
  End If

  Exit Function

SingleValue:
  GetVariantItemCount = 1

End Function

'============================================================
' 内部処理：配列からi番目の値を取得
'============================================================
Private Function GetVariantItem( _
  ByVal in_v As Variant, _
  ByVal in_Index As Long _
) As Variant

  On Error GoTo TryOneDim

  If IsArray(in_v) Then
    GetVariantItem = in_v(LBound(in_v, 1) + in_Index - 1, 1)
  Else
    GetVariantItem = in_v
  End If

  Exit Function

TryOneDim:
  On Error GoTo SingleValue

  GetVariantItem = in_v(LBound(in_v) + in_Index - 1)
  Exit Function

SingleValue:
  GetVariantItem = in_v

End Function

'============================================================
' 内部処理：初期化チェック
'============================================================
Private Sub CheckInitialized()

  If Not mIsInitialized Then
    Err.Raise vbObjectError + 4016, _
      "clsOISZeroLinearCurve", _
      "clsOISZeroLinearCurve が初期化されていません。Init を先に実行してください。"
  End If

End Sub

'============================================================
' clsDiscountCurve Interface Implementation
'
' Implements clsDiscountCurve の実装は、必ず Private で書く。
' メソッド名は clsDiscountCurve_メソッド名 にする。
'============================================================
Private Function clsDiscountCurve_DF(ByVal in_TargetDate As Date) As Double
  clsDiscountCurve_DF = Me.DF(in_TargetDate)
End Function

Private Function clsDiscountCurve_ZeroRateCont(ByVal in_TargetDate As Date) As Double
  clsDiscountCurve_ZeroRateCont = Me.ZeroRateCont(in_TargetDate)
End Function

Private Function clsDiscountCurve_ForwardRate( _
  ByVal in_StartDate As Date, _
  ByVal in_EndDate As Date _
) As Double

  clsDiscountCurve_ForwardRate = Me.ForwardRate(in_StartDate, in_EndDate)

End Function
