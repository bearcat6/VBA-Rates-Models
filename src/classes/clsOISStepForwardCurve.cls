Option Explicit

Implements clsDiscountCurve

'============================================================
' clsOISStepForwardCurve
'
' JPY OIS Step Forward Curve
'
' 目的：
'   ・瞬間ONフォワード金利 f(t) が階段状のOISカーブを構築する
'   ・Discount Factor:
'
'       DF(0,T) = Exp( - Integral[0,T] f(u) du )
'
'   ・f(t) は各セグメントで一定
'
' 前提：
'   ・評価日からスポット日まではONを横置き
'   ・スポット日は評価日+2営業日
'   ・1M,2M,3M,6Mは、前回点から満期までを1区間フラットで解く
'   ・6M以降は3M刻み
'   ・観測点間に複数の3M区間がある場合、
'       f_i = f_prev + i * Δf
'     となるように Δf を1つだけ解く
'
' 依存：
'   ・clsDiscountCurve
'   ・clsHolidayCalendar
'   ・clsOISSwap
'   ・mdl_BusinessDay
'============================================================

Private mValuationDate As Date
Private mSpotDate As Date
Private mCalendar As clsHolidayCalendar
Private mBusinessDayConvention As String
Private mPaymentLag As Long

Private mSegCount As Long
Private mSegStart() As Date
Private mSegEnd() As Date
Private mSegFwd() As Double

Private mIsInitialized As Boolean

Private Const ROOT_TOL As Double = 0.000000000001
Private Const ROOT_MAX_ITER As Long = 100

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
    Err.Raise vbObjectError + 3001, _
      "clsOISStepForwardCurve.Init", _
      "SpotLag は0以上にしてください。"
  End If

  If in_PaymentLag < 0 Then
    Err.Raise vbObjectError + 3002, _
      "clsOISStepForwardCurve.Init", _
      "PaymentLag は0以上にしてください。"
  End If

  mValuationDate = DateValue(in_ValuationDate)
  Set mCalendar = in_Calendar

  mSpotDate = GetWorkday(mValuationDate, in_spotLag, mCalendar)
  mPaymentLag = in_PaymentLag
  mBusinessDayConvention = in_BusinessDayConvention

  ClearSegments

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

Public Property Get SegmentCount() As Long
  SegmentCount = mSegCount
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
' セグメントをクリア
'============================================================
Public Sub ClearSegments()

  mSegCount = 0

  Erase mSegStart
  Erase mSegEnd
  Erase mSegFwd

End Sub

'============================================================
' セグメント追加
'
' in_forwardRate は、連続複利ベースの瞬間フォワードとして扱う
'============================================================
Public Sub AddSegment( _
  ByVal in_StartDate As Date, _
  ByVal in_EndDate As Date, _
  ByVal in_forwardRate As Double _
)

  If in_EndDate <= in_StartDate Then
    Err.Raise vbObjectError + 3003, _
      "clsOISStepForwardCurve.AddSegment", _
      "Segment End は Segment Start より後の日付にしてください。"
  End If

  If mSegCount > 0 Then
    If in_StartDate <> mSegEnd(mSegCount) Then
      Err.Raise vbObjectError + 3004, _
        "clsOISStepForwardCurve.AddSegment", _
        "セグメントが連続していません。前セグメントのEndと新セグメントのStartを一致させてください。"
    End If
  End If

  mSegCount = mSegCount + 1

  ReDim Preserve mSegStart(1 To mSegCount)
  ReDim Preserve mSegEnd(1 To mSegCount)
  ReDim Preserve mSegFwd(1 To mSegCount)

  mSegStart(mSegCount) = DateValue(in_StartDate)
  mSegEnd(mSegCount) = DateValue(in_EndDate)
  mSegFwd(mSegCount) = in_forwardRate

End Sub

'============================================================
' DF
'
' targetDate = valuationDate なら 1
' targetDate < valuationDate ならエラー
' targetDate > 最終セグメントEnd + PaymentLag ならエラー
'============================================================
Public Function df(ByVal in_TargetDate As Date) As Double

  Dim integral As Double
  Dim extrapolationLimitDate As Date

  CheckInitialized

  in_TargetDate = DateValue(in_TargetDate)

  If in_TargetDate = mValuationDate Then
    df = 1#
    Exit Function
  End If

  If in_TargetDate < mValuationDate Then
    Err.Raise vbObjectError + 3005, _
      "clsOISStepForwardCurve.DF", _
      "評価日より前の日付に対するDFは計算できません。"
  End If

  If mSegCount = 0 Then
    Err.Raise vbObjectError + 3006, _
      "clsOISStepForwardCurve.DF", _
      "フォワードセグメントが設定されていません。"
  End If

  extrapolationLimitDate = GetWorkday(mSegEnd(mSegCount), mPaymentLag, mCalendar)

  If in_TargetDate > extrapolationLimitDate Then
    Err.Raise vbObjectError + 3007, _
      "clsOISStepForwardCurve.DF", _
      "targetDate がカーブ最終日および支払ラグ許容範囲を超えています。"
  End If

  integral = IntegralForwardAllowPaymentLag(mValuationDate, in_TargetDate)

  df = Exp(-integral)

End Function

'============================================================
' 連続複利ゼロレート
'
' ZeroRateCont(T) = -Log(DF(T)) / YearFrac(ValDate,T)
'============================================================
Public Function ZeroRateCont(ByVal in_TargetDate As Date) As Double

  Dim yf As Double
  Dim dfTarget As Double

  CheckInitialized

  in_TargetDate = DateValue(in_TargetDate)

  If in_TargetDate <= mValuationDate Then
    Err.Raise vbObjectError + 3008, _
      "clsOISStepForwardCurve.ZeroRateCont", _
      "targetDate は評価日より後の日付にしてください。"
  End If

  yf = YearFracAct365F(mValuationDate, in_TargetDate)
  dfTarget = df(in_TargetDate)

  ZeroRateCont = -Log(dfTarget) / yf

End Function

'============================================================
' 期間フォワードレート
'
' 単純年率として返す。
'
' ForwardRate(S,E)
'   = ( DF(S) / DF(E) - 1 ) / YearFrac(S,E)
'============================================================
Public Function ForwardRate(ByVal in_StartDate As Date, _
                            ByVal in_EndDate As Date) As Double

  Dim dfStart As Double
  Dim dfEnd As Double
  Dim yf As Double

  CheckInitialized

  in_StartDate = DateValue(in_StartDate)
  in_EndDate = DateValue(in_EndDate)

  If in_EndDate <= in_StartDate Then
    Err.Raise vbObjectError + 3009, _
      "clsOISStepForwardCurve.ForwardRate", _
      "EndDate は StartDate より後の日付にしてください。"
  End If

  dfStart = df(in_StartDate)
  dfEnd = df(in_EndDate)
  yf = YearFracAct365F(in_StartDate, in_EndDate)

  ForwardRate = (dfStart / dfEnd - 1#) / yf

End Function


'============================================================
' Forward Par Rate
'
' Expiry, Tenor を指定して、Forward Starting OIS Swap の
' Par Rate を返す。
'
' 例：
'   ForwardParRate("6M", "5Y")
'
' StartDate = SpotDate + Expiry
' EndDate   = StartDate + Tenor
'
' ParRate =
'   ( DF(Start) - DF(End) ) / Sum[ yearFrac_i * DF(payDate_i) ]
'
' 固定レグは 1Y 間隔を基本とし、端数期間は最終期間として扱う。
'============================================================
Public Function ForwardParRate(ByVal in_Expiry As String, _
                               ByVal in_Tenor As String) As Double

  Dim startDate As Date
  Dim endDate As Date
  
  CheckInitialized
  
  startDate = GetBusinessDay( _
    mSpotDate, _
    in_Expiry, _
    mBusinessDayConvention, _
    mCalendar _
  )
  
  endDate = GetBusinessDay( _
    startDate, _
    in_Tenor, _
    mBusinessDayConvention, _
    mCalendar _
  )
  
  ForwardParRate = ForwardParRateByDates(startDate, endDate)

End Function


'============================================================
' Forward Par Rate（日付指定版）
'
' Swaption や CMS 等で、開始日・終了日を直接指定したい場合用。
'============================================================
Public Function ForwardParRateByDates(ByVal in_StartDate As Date, _
                                      ByVal in_EndDate As Date) As Double

  Dim periodStart As Date
  Dim periodEnd As Date
  Dim payDate As Date
  
  Dim dfStart As Double
  Dim dfEnd As Double
  Dim annuity As Double
  Dim yf As Double
  
  CheckInitialized
  
  in_StartDate = DateValue(in_StartDate)
  in_EndDate = DateValue(in_EndDate)
  
  If in_EndDate <= in_StartDate Then
    Err.Raise vbObjectError + 3026, _
      "clsOISStepForwardCurve.ForwardParRateByDates", _
      "EndDate は StartDate より後の日付にしてください。"
  End If
  
  dfStart = df(in_StartDate)
  dfEnd = df(in_EndDate)
  
  annuity = 0#
  periodStart = in_StartDate
  
  Do While periodStart < in_EndDate
  
    periodEnd = GetBusinessDay( _
      periodStart, _
      "1Y", _
      mBusinessDayConvention, _
      mCalendar _
    )
    
    If periodEnd > in_EndDate Then
      periodEnd = in_EndDate
    End If
    
    payDate = GetWorkday(periodEnd, mPaymentLag, mCalendar)
    yf = YearFracAct365F(periodStart, periodEnd)
    
    annuity = annuity + yf * df(payDate)
    
    periodStart = periodEnd
  
  Loop
  
  If annuity <= 0# Then
    Err.Raise vbObjectError + 3027, _
      "clsOISStepForwardCurve.ForwardParRateByDates", _
      "固定レグのAnnuityが0以下です。"
  End If
  
  ForwardParRateByDates = (dfStart - dfEnd) / annuity

End Function


'============================================================
' Quote配列からカーブを構築
'
' in_tenors:
'   "ON", "1M", "2M", "3M", "6M", "1Y", "2Y", ...
'
' in_rates:
'   OISレート
'============================================================
Public Sub BuildFromQuotes(ByVal in_Tenors As Variant, _
                           ByVal in_Rates As Variant)

  Dim onRate As Double
  Dim quoteMonths() As Long
  Dim quoteRates() As Double
  Dim quoteCount As Long

  Dim i As Long
  Dim nextBusinessDate As Date
  Dim onForward As Double

  CheckInitialized

  ClearSegments

  Call ExtractAndSortQuotes( _
    in_Tenors, _
    in_Rates, _
    onRate, _
    quoteMonths, _
    quoteRates, _
    quoteCount _
  )

  If quoteCount = 0 Then
    Err.Raise vbObjectError + 3010, _
      "clsOISStepForwardCurve.BuildFromQuotes", _
      "ON以外の有効なOISレートがありません。"
  End If

  nextBusinessDate = GetWorkday(mValuationDate, 1, mCalendar)
  onForward = SimpleRateToContinuousForward( _
    onRate, _
    YearFracAct365F(mValuationDate, nextBusinessDate) _
  )

  AddSegment mValuationDate, mSpotDate, onForward

  For i = 1 To quoteCount
    BootstrapOneQuote quoteMonths(i), quoteRates(i)
  Next i

End Sub

'============================================================
' セグメント表を返す
'============================================================
Public Function GetSegmentTable() As Variant

  Dim arr() As Variant
  Dim i As Long

  CheckInitialized

  If mSegCount = 0 Then
    Err.Raise vbObjectError + 3011, _
      "clsOISStepForwardCurve.GetSegmentTable", _
      "セグメントがありません。"
  End If

  ReDim arr(1 To mSegCount + 1, 1 To 7)

  arr(1, 1) = "Segment No"
  arr(1, 2) = "Start Date"
  arr(1, 3) = "End Date"
  arr(1, 4) = "YearFrac"
  arr(1, 5) = "Instantaneous Forward"
  arr(1, 6) = "DF Start"
  arr(1, 7) = "DF End"

  For i = 1 To mSegCount
    arr(i + 1, 1) = i
    arr(i + 1, 2) = mSegStart(i)
    arr(i + 1, 3) = mSegEnd(i)
    arr(i + 1, 4) = YearFracAct365F(mSegStart(i), mSegEnd(i))
    arr(i + 1, 5) = mSegFwd(i)
    arr(i + 1, 6) = df(mSegStart(i))
    arr(i + 1, 7) = df(mSegEnd(i))
  Next i

  GetSegmentTable = arr

End Function

'============================================================
' セグメント表をシートに出力
'============================================================
Public Sub ExportSegmentsToSheet( _
  ByVal in_SheetName As String, _
  ByVal in_StartCell As String _
)

  Dim ws As Worksheet
  Dim arr As Variant

  Set ws = ThisWorkbook.Worksheets(in_SheetName)

  arr = GetSegmentTable()

  ws.Range(in_StartCell).Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr

End Sub

'============================================================
' 内部処理：1Quote分をブートストラップ
'============================================================
Private Sub BootstrapOneQuote( _
  ByVal in_maturityMonths As Long, _
  ByVal in_FixedRate As Double _
)

  Dim maturityDate As Date
  Dim prevEndDate As Date
  Dim prevForward As Double
  Dim solvedValue As Double

  CheckInitialized

  maturityDate = GetBusinessDay( _
    mSpotDate, _
    CStr(in_maturityMonths) & "M", _
    mBusinessDayConvention, _
    mCalendar _
  )

  prevEndDate = mSegEnd(mSegCount)
  prevForward = mSegFwd(mSegCount)

  If maturityDate <= prevEndDate Then
    Exit Sub
  End If

  If in_maturityMonths <= 6 Then
    solvedValue = SolveFlatForward( _
      prevEndDate, _
      maturityDate, _
      in_FixedRate _
    )

    AddSegment prevEndDate, maturityDate, solvedValue
  Else
    solvedValue = SolveDeltaForward( _
      in_maturityMonths, _
      in_FixedRate, _
      prevForward _
    )

    AppendDeltaForwardSegments _
      in_maturityMonths, _
      prevForward, _
      solvedValue
  End If

End Sub

'============================================================
' 内部処理：フラットフォワードを解く
'============================================================
Private Function SolveFlatForward( _
  ByVal in_segmentStart As Date, _
  ByVal in_segmentEnd As Date, _
  ByVal in_FixedRate As Double _
) As Double

  SolveFlatForward = SolveRootBisection( _
    -0.1, _
    0.1, _
    in_FixedRate, _
    "FLAT", _
    in_segmentStart, _
    in_segmentEnd, _
    0, _
    0# _
  )

End Function

'============================================================
' 内部処理：Δfを解く
'============================================================
Private Function SolveDeltaForward( _
  ByVal in_maturityMonths As Long, _
  ByVal in_FixedRate As Double, _
  ByVal in_prevForward As Double _
) As Double

  SolveDeltaForward = SolveRootBisection( _
    -0.05, _
    0.05, _
    in_FixedRate, _
    "DELTA", _
    0, _
    0, _
    in_maturityMonths, _
    in_prevForward _
  )

End Function

'============================================================
' 内部処理：二分法
'============================================================
Private Function SolveRootBisection( _
  ByVal in_lower As Double, _
  ByVal in_upper As Double, _
  ByVal in_FixedRate As Double, _
  ByVal in_mode As String, _
  ByVal in_flatStart As Date, _
  ByVal in_flatEnd As Date, _
  ByVal in_maturityMonths As Long, _
  ByVal in_prevForward As Double _
) As Double

  Dim lower As Double
  Dim upper As Double
  Dim mid As Double

  Dim fLower As Double
  Dim fUpper As Double
  Dim fMid As Double

  Dim i As Long
  Dim expandCount As Long

  lower = in_lower
  upper = in_upper

  fLower = BootstrapObjective( _
    lower, _
    in_FixedRate, _
    in_mode, _
    in_flatStart, _
    in_flatEnd, _
    in_maturityMonths, _
    in_prevForward _
  )

  fUpper = BootstrapObjective( _
    upper, _
    in_FixedRate, _
    in_mode, _
    in_flatStart, _
    in_flatEnd, _
    in_maturityMonths, _
    in_prevForward _
  )

  expandCount = 0

  Do While fLower * fUpper > 0# And expandCount < 10
    lower = lower * 2#
    upper = upper * 2#

    fLower = BootstrapObjective( _
      lower, _
      in_FixedRate, _
      in_mode, _
      in_flatStart, _
      in_flatEnd, _
      in_maturityMonths, _
      in_prevForward _
    )

    fUpper = BootstrapObjective( _
      upper, _
      in_FixedRate, _
      in_mode, _
      in_flatStart, _
      in_flatEnd, _
      in_maturityMonths, _
      in_prevForward _
    )

    expandCount = expandCount + 1
  Loop

  If fLower * fUpper > 0# Then
    Err.Raise vbObjectError + 3012, _
      "clsOISStepForwardCurve.SolveRootBisection", _
      "ブートストラップで解を挟めませんでした。入力レートまたは初期範囲を確認してください。"
  End If

  For i = 1 To ROOT_MAX_ITER
    mid = 0.5 * (lower + upper)

    fMid = BootstrapObjective( _
      mid, _
      in_FixedRate, _
      in_mode, _
      in_flatStart, _
      in_flatEnd, _
      in_maturityMonths, _
      in_prevForward _
    )

    If Abs(fMid) < ROOT_TOL Then
      SolveRootBisection = mid
      Exit Function
    End If

    If fLower * fMid <= 0# Then
      upper = mid
      fUpper = fMid
    Else
      lower = mid
      fLower = fMid
    End If
  Next i

  SolveRootBisection = 0.5 * (lower + upper)

End Function

'============================================================
' 内部処理：ブートストラップ目的関数
'============================================================
Private Function BootstrapObjective(ByVal in_X As Double, _
                                    ByVal in_FixedRate As Double, _
                                    ByVal in_mode As String, _
                                    ByVal in_flatStart As Date, _
                                    ByVal in_flatEnd As Date, _
                                    ByVal in_maturityMonths As Long, _
                                    ByVal in_prevForward As Double) As Double

  Dim originalSegCount As Long
  Dim maturityDate As Date
  Dim swp As clsOISswap
  Dim cCurve As clsDiscountCurve

  originalSegCount = mSegCount

  Select Case UCase$(in_mode)
    Case "FLAT"
      AddSegment in_flatStart, in_flatEnd, in_X
      maturityDate = in_flatEnd

    Case "DELTA"
      AppendDeltaForwardSegments _
        in_maturityMonths, _
        in_prevForward, _
        in_X

      maturityDate = GetBusinessDay( _
        mSpotDate, _
        CStr(in_maturityMonths) & "M", _
        mBusinessDayConvention, _
        mCalendar _
      )

    Case Else
      Err.Raise vbObjectError + 3013, _
        "clsOISStepForwardCurve.BootstrapObjective", _
        "mode は FLAT または DELTA にしてください。"
  End Select

  Set swp = New clsOISswap

  swp.Init _
    in_StartDate:=mSpotDate, _
    in_EndDate:=maturityDate, _
    in_FixedRate:=in_FixedRate, _
    in_Notional:=1#, _
    in_PaymentLag:=mPaymentLag, _
    in_BusinessDayConvention:=mBusinessDayConvention, _
    in_Calendar:=mCalendar

  Set cCurve = Me
    
  BootstrapObjective = swp.NPV(cCurve)

  RestoreSegmentCount originalSegCount

End Function

'============================================================
' 内部処理：6M以降のΔfセグメントを追加
'============================================================
Private Sub AppendDeltaForwardSegments(ByVal in_maturityMonths As Long, _
                                       ByVal in_prevForward As Double, _
                                       ByVal in_deltaForward As Double)

  Dim currentEnd As Date
  Dim nextEnd As Date
  Dim K As Long
  Dim i As Long
  Dim F As Double

  currentEnd = mSegEnd(mSegCount)
  K = FindNextGridMonthAfterCurrentEnd(in_maturityMonths)
  i = 1

  Do While K <= in_maturityMonths
    nextEnd = GetBusinessDay( _
      mSpotDate, _
      CStr(K) & "M", _
      mBusinessDayConvention, _
      mCalendar _
    )

    If nextEnd > currentEnd Then
      F = in_prevForward + CDbl(i) * in_deltaForward
      AddSegment currentEnd, nextEnd, F

      currentEnd = nextEnd
      i = i + 1
    End If

    K = K + 3
  Loop

End Sub

'============================================================
' 内部処理：現在の最終日より後に来る次の3Mグリッド月数を探す
'============================================================
Private Function FindNextGridMonthAfterCurrentEnd(ByVal in_maturityMonths As Long) As Long

  Dim K As Long
  Dim d As Date
  Dim currentEnd As Date

  currentEnd = mSegEnd(mSegCount)
  K = 9

  Do While K <= in_maturityMonths
    d = GetBusinessDay( _
      mSpotDate, _
      CStr(K) & "M", _
      mBusinessDayConvention, _
      mCalendar _
    )

    If d > currentEnd Then
      FindNextGridMonthAfterCurrentEnd = K
      Exit Function
    End If

    K = K + 3
  Loop

  Err.Raise vbObjectError + 3014, _
    "clsOISStepForwardCurve.FindNextGridMonthAfterCurrentEnd", _
    "追加すべき3Mグリッドが見つかりませんでした。"

End Function

'============================================================
' 内部処理：指定区間の∫fdt
'============================================================
Private Function IntegralForward(ByVal in_StartDate As Date, _
                                 ByVal in_EndDate As Date) As Double

  Dim i As Long
  Dim s As Date
  Dim e As Date
  Dim overlapStart As Date
  Dim overlapEnd As Date
  Dim yf As Double
  Dim total As Double

  If in_EndDate < in_StartDate Then
    Err.Raise vbObjectError + 3015, _
      "clsOISStepForwardCurve.IntegralForward", _
      "EndDate は StartDate 以降にしてください。"
  End If

  If in_EndDate = in_StartDate Then
    IntegralForward = 0#
    Exit Function
  End If

  total = 0#

  For i = 1 To mSegCount
    s = mSegStart(i)
    e = mSegEnd(i)

    If e <= in_StartDate Then
      GoTo ContinueLoop
    End If

    If s >= in_EndDate Then
      Exit For
    End If

    overlapStart = MaxDate(s, in_StartDate)
    overlapEnd = MinDate(e, in_EndDate)

    If overlapEnd > overlapStart Then
      yf = YearFracAct365F(overlapStart, overlapEnd)
      total = total + mSegFwd(i) * yf
    End If

ContinueLoop:
  Next i

  IntegralForward = total

End Function

'============================================================
' 内部処理：PaymentLag範囲内の短期外挿を許容した∫fdt
'============================================================
Private Function IntegralForwardAllowPaymentLag(ByVal in_StartDate As Date, _
                                                ByVal in_EndDate As Date) As Double

  Dim baseEndDate As Date
  Dim total As Double
  Dim tailYearFrac As Double

  If in_EndDate < in_StartDate Then
    Err.Raise vbObjectError + 3025, _
      "clsOISStepForwardCurve.IntegralForwardAllowPaymentLag", _
      "EndDate は StartDate 以降にしてください。"
  End If

  If in_EndDate = in_StartDate Then
    IntegralForwardAllowPaymentLag = 0#
    Exit Function
  End If

  baseEndDate = mSegEnd(mSegCount)

  If in_EndDate <= baseEndDate Then
    IntegralForwardAllowPaymentLag = IntegralForward(in_StartDate, in_EndDate)
    Exit Function
  End If

  total = IntegralForward(in_StartDate, baseEndDate)
  tailYearFrac = YearFracAct365F(baseEndDate, in_EndDate)
  total = total + mSegFwd(mSegCount) * tailYearFrac

  IntegralForwardAllowPaymentLag = total

End Function

'============================================================
' 内部処理：Quote抽出・ソート
'============================================================
Private Sub ExtractAndSortQuotes(ByVal in_Tenors As Variant, _
                                 ByVal in_Rates As Variant, _
                                 ByRef out_onRate As Double, _
                                 ByRef out_quoteMonths() As Long, _
                                 ByRef out_quoteRates() As Double, _
                                 ByRef out_quoteCount As Long)

  Dim n As Long
  Dim i As Long
  Dim tenor As String
  Dim rate As Double
  Dim months As Long
  Dim hasON As Boolean

  n = GetVariantItemCount(in_Tenors)

  out_quoteCount = 0
  hasON = False

  For i = 1 To n
    tenor = CStr(GetVariantItem(in_Tenors, i))

    If Trim$(tenor) <> "" Then
      rate = CDbl(GetVariantItem(in_Rates, i))

      If UCase$(Trim$(tenor)) = "ON" Or UCase$(Trim$(tenor)) = "O/N" Then
        out_onRate = rate
        hasON = True
      Else
        months = TenorToMonths(tenor)

        If IsAcceptedQuoteMonth(months) Then
          out_quoteCount = out_quoteCount + 1

          ReDim Preserve out_quoteMonths(1 To out_quoteCount)
          ReDim Preserve out_quoteRates(1 To out_quoteCount)

          out_quoteMonths(out_quoteCount) = months
          out_quoteRates(out_quoteCount) = rate
        End If
      End If
    End If
  Next i

  If Not hasON Then
    Err.Raise vbObjectError + 3016, _
      "clsOISStepForwardCurve.ExtractAndSortQuotes", _
      "ONレートが見つかりません。"
  End If

  If out_quoteCount > 1 Then
    SortQuotesByMonths out_quoteMonths, out_quoteRates, out_quoteCount
    RemoveDuplicateQuoteMonths out_quoteMonths, out_quoteRates, out_quoteCount
  End If

End Sub

'============================================================
' 内部処理：採用するQuote月数か
'============================================================
Private Function IsAcceptedQuoteMonth(ByVal in_Months As Long) As Boolean

  Select Case in_Months
    Case 1, 2, 3, 6
      IsAcceptedQuoteMonth = True

    Case Else
      If in_Months > 6 And in_Months Mod 3 = 0 Then
        IsAcceptedQuoteMonth = True
      Else
        IsAcceptedQuoteMonth = False
      End If
  End Select

End Function

'============================================================
' 内部処理：テナーを月数に変換
'============================================================
Private Function TenorToMonths(ByVal in_Tenor As String) As Long

  Dim s As String
  Dim unitChar As String
  Dim numPart As String
  Dim n As Long

  s = UCase$(Trim$(in_Tenor))
  s = Replace(s, " ", "")

  If Len(s) < 2 Then
    Err.Raise vbObjectError + 3017, _
      "clsOISStepForwardCurve.TenorToMonths", _
      "テナー指定が不正です。例：1M, 6M, 1Y"
  End If

  unitChar = Right$(s, 1)
  numPart = Left$(s, Len(s) - 1)

  If Not IsNumeric(numPart) Then
    Err.Raise vbObjectError + 3018, _
      "clsOISStepForwardCurve.TenorToMonths", _
      "テナーの数値部分が不正です。"
  End If

  n = CLng(numPart)

  Select Case unitChar
    Case "M"
      TenorToMonths = n

    Case "Y"
      TenorToMonths = n * 12

    Case Else
      Err.Raise vbObjectError + 3019, _
        "clsOISStepForwardCurve.TenorToMonths", _
        "テナー単位は M または Y にしてください。"
  End Select

End Function

'============================================================
' 内部処理：Quoteを月数順にソート
'============================================================
Private Sub SortQuotesByMonths(ByRef io_months() As Long, _
                               ByRef io_rates() As Double, _
                               ByVal in_count As Long)

  Dim i As Long
  Dim j As Long
  Dim tmpM As Long
  Dim tmpR As Double

  For i = 1 To in_count - 1
    For j = i + 1 To in_count
      If io_months(j) < io_months(i) Then
        tmpM = io_months(i)
        io_months(i) = io_months(j)
        io_months(j) = tmpM

        tmpR = io_rates(i)
        io_rates(i) = io_rates(j)
        io_rates(j) = tmpR
      End If
    Next j
  Next i

End Sub

'============================================================
' 内部処理：同じ月数のQuoteがある場合、後勝ちで1つにする
'============================================================
Private Sub RemoveDuplicateQuoteMonths(ByRef io_months() As Long, _
                                       ByRef io_rates() As Double, _
                                       ByRef io_count As Long)

  Dim newMonths() As Long
  Dim newRates() As Double
  Dim newCount As Long
  Dim i As Long

  newCount = 0

  For i = 1 To io_count
    If i < io_count Then
      If io_months(i) = io_months(i + 1) Then
        ' 後勝ちにするため、ここではスキップ
      Else
        newCount = newCount + 1
        ReDim Preserve newMonths(1 To newCount)
        ReDim Preserve newRates(1 To newCount)
        newMonths(newCount) = io_months(i)
        newRates(newCount) = io_rates(i)
      End If
    Else
      newCount = newCount + 1
      ReDim Preserve newMonths(1 To newCount)
      ReDim Preserve newRates(1 To newCount)
      newMonths(newCount) = io_months(i)
      newRates(newCount) = io_rates(i)
    End If
  Next i

  io_count = newCount
  io_months = newMonths
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
Private Function GetVariantItem(ByVal in_v As Variant, _
                                ByVal in_Index As Long) As Variant

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
' 内部処理：セグメント数を元に戻す
'============================================================
Private Sub RestoreSegmentCount(ByVal in_segmentCount As Long)

  If in_segmentCount < 0 Or in_segmentCount > mSegCount Then
    Err.Raise vbObjectError + 3020, _
      "clsOISStepForwardCurve.RestoreSegmentCount", _
      "復元するセグメント数が不正です。"
  End If

  mSegCount = in_segmentCount

  If mSegCount = 0 Then
    Erase mSegStart
    Erase mSegEnd
    Erase mSegFwd
  Else
    ReDim Preserve mSegStart(1 To mSegCount)
    ReDim Preserve mSegEnd(1 To mSegCount)
    ReDim Preserve mSegFwd(1 To mSegCount)
  End If

End Sub

'============================================================
' 内部処理：単純年率を連続フォワードに変換
'============================================================
Private Function SimpleRateToContinuousForward(ByVal in_simpleRate As Double, _
                                               ByVal in_yearFrac As Double) As Double

  If in_yearFrac <= 0# Then
    Err.Raise vbObjectError + 3021, _
      "clsOISStepForwardCurve.SimpleRateToContinuousForward", _
      "YearFrac は正の値にしてください。"
  End If

  If 1# + in_simpleRate * in_yearFrac <= 0# Then
    Err.Raise vbObjectError + 3022, _
      "clsOISStepForwardCurve.SimpleRateToContinuousForward", _
      "1 + rate * yearFrac が0以下です。"
  End If

  SimpleRateToContinuousForward = Log(1# + in_simpleRate * in_yearFrac) / in_yearFrac

End Function


'============================================================
' 内部処理：Max Date
'============================================================
Private Function MaxDate(ByVal in_date1 As Date, _
                         ByVal in_date2 As Date) As Date

  If in_date1 >= in_date2 Then
    MaxDate = in_date1
  Else
    MaxDate = in_date2
  End If

End Function

'============================================================
' 内部処理：Min Date
'============================================================
Private Function MinDate(ByVal in_date1 As Date, _
                         ByVal in_date2 As Date) As Date

  If in_date1 <= in_date2 Then
    MinDate = in_date1
  Else
    MinDate = in_date2
  End If

End Function

'============================================================
' 内部処理：初期化チェック
'============================================================
Private Sub CheckInitialized()

  If Not mIsInitialized Then
    Err.Raise vbObjectError + 3024, _
      "clsOISStepForwardCurve", _
      "clsOISStepForwardCurve が初期化されていません。Init を先に実行してください。"
  End If

End Sub

'============================================================
' clsDiscountCurve Interface Implementation
'
' Implements clsDiscountCurve の実装は、必ず Private で書く。
' メソッド名は clsDiscountCurve_メソッド名 にする。
'============================================================
Private Function clsDiscountCurve_DF(ByVal in_TargetDate As Date) As Double
  clsDiscountCurve_DF = Me.df(in_TargetDate)
End Function

Private Function clsDiscountCurve_ZeroRateCont(ByVal in_TargetDate As Date) As Double
  clsDiscountCurve_ZeroRateCont = Me.ZeroRateCont(in_TargetDate)
End Function

Private Function clsDiscountCurve_ForwardRate(ByVal in_StartDate As Date, _
                                              ByVal in_EndDate As Date) As Double

  clsDiscountCurve_ForwardRate = Me.ForwardRate(in_StartDate, in_EndDate)

End Function


