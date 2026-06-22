
Option Explicit

'============================================================
' clsOISswap
'
' JPY OIS swap class
'
' 前提：
'   ・TONA複利相当の変動脚 vs 固定脚
'   ・Delay方式：利息期間 EndDate + PaymentLag 営業日を支払日とする
'   ・1年以内は満期一括、1年超は年1回払い
'   ・期日は Modified Following 等で営業日調整済みの日付を使う
'
' 主な修正点：
'   ・期間数を先に数える方式を廃止
'   ・営業日調整後の期日だけを積み上げる方式に変更
'   ・同じ日付の期間、EndDate <= StartDate の期間を作らない
'============================================================

Private Const ERR_BASE As Long = vbObjectError + 2000
Private Const MAX_SCHEDULE_YEARS As Long = 100

Private mStartDate As Date
Private mEndDate As Date
Private mFixedRate As Double
Private mNotional As Double
Private mPaymentLag As Long
Private mPayReceive As String
Private mBusinessDayConvention As String
Private mCalendar As clsHolidayCalendar
Private mIsInitialized As Boolean

'============================================================
' 初期化
'============================================================
Public Sub Init(ByVal in_StartDate As Date, _
                ByVal in_EndDate As Date, _
                ByVal in_FixedRate As Double, _
                ByVal in_Notional As Double, _
                Optional ByVal in_PaymentLag As Long = 2, _
                Optional ByVal in_BusinessDayConvention As String = "MF", _
                Optional ByVal in_Calendar As clsHolidayCalendar = Nothing, _
                Optional ByVal in_PayReceive As String = "PAYER")

  If DateValue(in_EndDate) <= DateValue(in_StartDate) Then
    Err.Raise ERR_BASE + 1, _
              "clsOISswap.Init", _
              "EndDate は StartDate より後の日付にしてください。"
  End If

  If in_Notional <= 0# Then
    Err.Raise ERR_BASE + 2, _
              "clsOISswap.Init", _
              "Notional は正の値にしてください。"
  End If

  If in_PaymentLag < 0 Then
    Err.Raise ERR_BASE + 3, _
              "clsOISswap.Init", _
              "PaymentLag は0以上にしてください。"
  End If

  mStartDate = DateValue(in_StartDate)
  mEndDate = DateValue(in_EndDate)
  mFixedRate = in_FixedRate
  mNotional = in_Notional
  mPaymentLag = in_PaymentLag
  mBusinessDayConvention = NormalizeBusinessDayConventionForSwap(in_BusinessDayConvention)
  mPayReceive = NormalizePayReceive(in_PayReceive)

  Set mCalendar = in_Calendar

  mIsInitialized = True

End Sub

'============================================================
' Property
'============================================================
Public Property Get StartDate() As Date
  StartDate = mStartDate
End Property

Public Property Get EndDate() As Date
  EndDate = mEndDate
End Property

Public Property Get FixedRate() As Double
  FixedRate = mFixedRate
End Property

Public Property Let FixedRate(ByVal in_FixedRate As Double)
  mFixedRate = in_FixedRate
End Property

Public Property Get Notional() As Double
  Notional = mNotional
End Property

Public Property Let Notional(ByVal in_Notional As Double)

  If in_Notional <= 0# Then
    Err.Raise ERR_BASE + 4, _
              "clsOISswap.Notional", _
              "Notional は正の値にしてください。"
  End If

  mNotional = in_Notional

End Property

Public Property Get PaymentLag() As Long
  PaymentLag = mPaymentLag
End Property

Public Property Let PaymentLag(ByVal in_PaymentLag As Long)

  If in_PaymentLag < 0 Then
    Err.Raise ERR_BASE + 5, _
              "clsOISswap.PaymentLag", _
              "PaymentLag は0以上にしてください。"
  End If

  mPaymentLag = in_PaymentLag

End Property

Public Property Get PayReceive() As String
  PayReceive = mPayReceive
End Property

Public Property Let PayReceive(ByVal in_PayReceive As String)
  mPayReceive = NormalizePayReceive(in_PayReceive)
End Property

Public Property Get BusinessDayConvention() As String
  BusinessDayConvention = mBusinessDayConvention
End Property

Public Property Let BusinessDayConvention(ByVal in_BusinessDayConvention As String)
  mBusinessDayConvention = NormalizeBusinessDayConventionForSwap(in_BusinessDayConvention)
End Property

Public Property Get Calendar() As clsHolidayCalendar
  Set Calendar = mCalendar
End Property

Public Property Set Calendar(ByVal in_Calendar As clsHolidayCalendar)
  Set mCalendar = in_Calendar
End Property

Public Property Get IsInitialized() As Boolean
  IsInitialized = mIsInitialized
End Property

'============================================================
' 固定脚PV
'
' FixedLegPV = Σ Notional * FixedRate * YearFrac_i * DF(PayDate_i)
'
' 脚単体のPVを正値で返す。
' PAYER / RECEIVER の符号は NPV 側で反映する。
'============================================================
Public Function FixedLegPV(ByVal in_cCurve As clsDiscountCurve) As Double

  Dim schedule As Variant
  Dim i As Long
  Dim n As Long
  Dim yearFrac As Double
  Dim payDate As Date
  Dim dfPay As Double
  Dim pv As Double

  CheckInitialized
  CheckCurve in_cCurve

  schedule = GetScheduleArray()
  n = UBound(schedule, 1)

  pv = 0#

  For i = 1 To n
    yearFrac = CDbl(schedule(i, 4))
    payDate = CDate(schedule(i, 3))

    dfPay = in_cCurve.df(payDate)
    pv = pv + mNotional * mFixedRate * yearFrac * dfPay
  Next i

  FixedLegPV = pv

End Function

'============================================================
' 変動脚PV
'
' FloatingLegPV = Σ Notional * ForwardRate(Start_i, End_i) * YearFrac_i * DF(PayDate_i)
'
' curve.ForwardRate は、当該期間の複利と整合する単純年率を返す前提。
' 脚単体のPVを正値で返す。
'============================================================
Public Function FloatingLegPV(ByVal in_cCurve As clsDiscountCurve) As Double

  Dim schedule As Variant
  Dim i As Long
  Dim n As Long
  Dim periodStart As Date
  Dim periodEnd As Date
  Dim payDate As Date
  Dim yearFrac As Double
  Dim fwdRate As Double
  Dim dfPay As Double
  Dim pv As Double

  CheckInitialized
  CheckCurve in_cCurve

  schedule = GetScheduleArray()
  n = UBound(schedule, 1)

  pv = 0#

  For i = 1 To n
    periodStart = CDate(schedule(i, 1))
    periodEnd = CDate(schedule(i, 2))
    payDate = CDate(schedule(i, 3))
    yearFrac = CDbl(schedule(i, 4))

    fwdRate = in_cCurve.ForwardRate(periodStart, periodEnd)
    dfPay = in_cCurve.df(payDate)

    pv = pv + mNotional * fwdRate * yearFrac * dfPay
  Next i

  FloatingLegPV = pv

End Function

'============================================================
' NPV
'
' FixedLegPV / FloatingLegPV は脚単体のPVを正値で返す。
'
'   PAYER   : FloatingLegPV - FixedLegPV
'   RECEIVER: FixedLegPV - FloatingLegPV
'============================================================
Public Function NPV(ByVal in_cCurve As clsDiscountCurve) As Double

  Dim fixedPV As Double
  Dim floatingPV As Double

  CheckInitialized
  CheckCurve in_cCurve

  fixedPV = FixedLegPV(in_cCurve)
  floatingPV = FloatingLegPV(in_cCurve)

  Select Case mPayReceive
    Case "PAYER"
      NPV = floatingPV - fixedPV

    Case "RECEIVER"
      NPV = fixedPV - floatingPV

    Case Else
      Err.Raise ERR_BASE + 10, _
                "clsOISswap.NPV", _
                "PayReceive が不正です。PAYER または RECEIVER を指定してください。"
  End Select

End Function

'============================================================
' Par Rate
'
' FloatingLegPV = FixedLegPV となる固定金利。
'
' ParRate = FloatingLegPV / FixedAnnuity
' FixedAnnuity = Σ Notional * YearFrac_i * DF(PayDate_i)
'
' ParRate は PayReceive に依存しない。
'============================================================
Public Function ParRate(ByVal in_cCurve As clsDiscountCurve) As Double

  Dim schedule As Variant
  Dim i As Long
  Dim n As Long
  Dim yearFrac As Double
  Dim payDate As Date
  Dim dfPay As Double
  Dim annuity As Double
  Dim floatPV As Double

  CheckInitialized
  CheckCurve in_cCurve

  schedule = GetScheduleArray()
  n = UBound(schedule, 1)

  annuity = 0#

  For i = 1 To n
    yearFrac = CDbl(schedule(i, 4))
    payDate = CDate(schedule(i, 3))
    dfPay = in_cCurve.df(payDate)

    annuity = annuity + mNotional * yearFrac * dfPay
  Next i

  If annuity = 0# Then
    Err.Raise ERR_BASE + 6, _
              "clsOISswap.ParRate", _
              "Fixed annuity がゼロです。"
  End If

  floatPV = FloatingLegPV(in_cCurve)

  ParRate = floatPV / annuity

End Function

'============================================================
' キャッシュフロー表を返す
'
' in_cCurve を渡した場合：
'   ・Forward Rate / DF / Amount / PV まで返す
'
' in_cCurve を省略した場合：
'   ・日付、YearFrac、固定金利、固定金額まで返す
'
' 返却列：
'   1  Period No
'   2  Start Date
'   3  End Date
'   4  Payment Date
'   5  YearFrac
'   6  PayReceive
'   7  Fixed Rate
'   8  Forward Rate
'   9  DF Payment
'   10 Fixed Amount
'   11 Floating Amount
'   12 Fixed PV
'   13 Floating PV
'   14 Net PV
'============================================================
Public Function GetCashflowTable(Optional ByVal in_cCurve As clsDiscountCurve = Nothing) As Variant

  Dim schedule As Variant
  Dim result() As Variant
  Dim i As Long
  Dim n As Long
  Dim hasCurve As Boolean

  Dim periodStart As Date
  Dim periodEnd As Date
  Dim payDate As Date
  Dim yearFrac As Double

  Dim fwdRate As Variant
  Dim dfPay As Variant
  Dim fixedAmount As Variant
  Dim floatingAmount As Variant
  Dim fixedPV As Variant
  Dim floatingPV As Variant
  Dim netPV As Variant

  CheckInitialized

  hasCurve = Not (in_cCurve Is Nothing)

  schedule = GetScheduleArray()
  n = UBound(schedule, 1)

  ReDim result(1 To n + 1, 1 To 14)

  result(1, 1) = "Period No"
  result(1, 2) = "Start Date"
  result(1, 3) = "End Date"
  result(1, 4) = "Payment Date"
  result(1, 5) = "YearFrac"
  result(1, 6) = "PayReceive"
  result(1, 7) = "Fixed Rate"
  result(1, 8) = "Forward Rate"
  result(1, 9) = "DF Payment"
  result(1, 10) = "Fixed Amount"
  result(1, 11) = "Floating Amount"
  result(1, 12) = "Fixed PV"
  result(1, 13) = "Floating PV"
  result(1, 14) = "Net PV"

  For i = 1 To n
    periodStart = CDate(schedule(i, 1))
    periodEnd = CDate(schedule(i, 2))
    payDate = CDate(schedule(i, 3))
    yearFrac = CDbl(schedule(i, 4))

    fixedAmount = mNotional * mFixedRate * yearFrac

    If Not hasCurve Then
      fwdRate = Empty
      dfPay = Empty
      floatingAmount = Empty
      fixedPV = Empty
      floatingPV = Empty
      netPV = Empty
    Else
      fwdRate = in_cCurve.ForwardRate(periodStart, periodEnd)
      dfPay = in_cCurve.df(payDate)
      floatingAmount = mNotional * CDbl(fwdRate) * yearFrac
      fixedPV = CDbl(fixedAmount) * CDbl(dfPay)
      floatingPV = CDbl(floatingAmount) * CDbl(dfPay)

      Select Case mPayReceive
        Case "PAYER"
          netPV = CDbl(floatingPV) - CDbl(fixedPV)

        Case "RECEIVER"
          netPV = CDbl(fixedPV) - CDbl(floatingPV)

        Case Else
          Err.Raise ERR_BASE + 12, _
                    "clsOISswap.GetCashflowTable", _
                    "PayReceive が不正です。PAYER または RECEIVER を指定してください。"
      End Select
    End If

    result(i + 1, 1) = i
    result(i + 1, 2) = periodStart
    result(i + 1, 3) = periodEnd
    result(i + 1, 4) = payDate
    result(i + 1, 5) = yearFrac
    result(i + 1, 6) = mPayReceive
    result(i + 1, 7) = mFixedRate
    result(i + 1, 8) = fwdRate
    result(i + 1, 9) = dfPay
    result(i + 1, 10) = fixedAmount
    result(i + 1, 11) = floatingAmount
    result(i + 1, 12) = fixedPV
    result(i + 1, 13) = floatingPV
    result(i + 1, 14) = netPV
  Next i

  GetCashflowTable = result

End Function

'============================================================
' 内部関数：支払スケジュール生成
'
' 返却列：
'   1 Start Date
'   2 End Date
'   3 Payment Date
'   4 YearFrac
'
' 修正ポイント：
'   ・期間数を先に決め打ちしない
'   ・営業日調整後の期日を使って、重複・逆転を除外する
'   ・最後は必ず mEndDate にする
'============================================================
Private Function GetScheduleArray() As Variant

  Dim periodEnds As Variant
  Dim schedule() As Variant
  Dim i As Long
  Dim n As Long

  Dim periodStart As Date
  Dim periodEnd As Date
  Dim payDate As Date
  Dim yearFrac As Double

  CheckInitialized

  periodEnds = BuildPeriodEndDates()
  n = UBound(periodEnds)

  ReDim schedule(1 To n, 1 To 4)

  periodStart = mStartDate

  For i = 1 To n
    periodEnd = CDate(periodEnds(i))

    If periodEnd <= periodStart Then
      Err.Raise ERR_BASE + 20, _
                "clsOISswap.GetScheduleArray", _
                "スケジュール生成で EndDate <= StartDate の期間が発生しました。"
    End If

    payDate = AddBusinessDays(periodEnd, mPaymentLag, mCalendar)
    yearFrac = YearFracAct365F(periodStart, periodEnd)

    schedule(i, 1) = periodStart
    schedule(i, 2) = periodEnd
    schedule(i, 3) = payDate
    schedule(i, 4) = yearFrac

    periodStart = periodEnd
  Next i

  GetScheduleArray = schedule

End Function

'============================================================
' 内部関数：期間終了日リストを作る
'
' 1年ごとの候補日を営業日調整して、mEndDate より前のものだけ採用する。
' 候補日が mEndDate と同じ場合は中間期日として採用せず、最後の期日だけにする。
'
' これにより、以下のようなゼロ日数期間を防ぐ。
'   2026/5/22 → 2027/5/24
'   2027/5/24 → 2027/5/24  ← 作らない
'============================================================
Private Function BuildPeriodEndDates() As Variant

  Dim periodEndCollection As Collection
  Dim result() As Date
  Dim y As Long
  Dim i As Long
  Dim candidate As Date
  Dim lastEnd As Date

  Set periodEndCollection = New Collection

  lastEnd = mStartDate
  y = 1

  Do
    candidate = AdjustBusinessDay( _
                  AddTenorMonths(mStartDate, 12 * y, True), _
                  mBusinessDayConvention, _
                  mCalendar _
                )

    If candidate >= mEndDate Then
      Exit Do
    End If

    If candidate > lastEnd Then
      periodEndCollection.Add candidate
      lastEnd = candidate
    End If

    y = y + 1

    If y > MAX_SCHEDULE_YEARS Then
      Err.Raise ERR_BASE + 21, _
                "clsOISswap.BuildPeriodEndDates", _
                "スケジュール生成の年数が上限を超えました。"
    End If
  Loop

  If mEndDate <= lastEnd Then
    Err.Raise ERR_BASE + 22, _
              "clsOISswap.BuildPeriodEndDates", _
              "最終期日が直前の期日以下になっています。"
  End If

  periodEndCollection.Add mEndDate

  ReDim result(1 To periodEndCollection.Count)

  For i = 1 To periodEndCollection.Count
    result(i) = CDate(periodEndCollection.Item(i))
  Next i

  BuildPeriodEndDates = result

End Function

'============================================================
' 内部関数：PayReceive 正規化
'============================================================
Private Function NormalizePayReceive(ByVal in_PayReceive As String) As String

  Dim s As String

  s = UCase$(Trim$(in_PayReceive))
  s = Replace(s, " ", "")
  s = Replace(s, "_", "")
  s = Replace(s, "-", "")

  Select Case s
    Case "PAYER", "PAY", "FIXEDPAYER", "PAYFIXED"
      NormalizePayReceive = "PAYER"

    Case "RECEIVER", "REC", "RECEIVE", "FIXEDRECEIVER", "RECEIVEFIXED"
      NormalizePayReceive = "RECEIVER"

    Case Else
      Err.Raise ERR_BASE + 11, _
                "clsOISswap.NormalizePayReceive", _
                "PayReceive は PAYER または RECEIVER を指定してください。"
  End Select

End Function

'============================================================
' 内部関数：Business Day Convention 正規化
'============================================================
Private Function NormalizeBusinessDayConventionForSwap(ByVal in_BusinessDayConvention As String) As String

  Dim s As String

  s = UCase$(Trim$(in_BusinessDayConvention))
  s = Replace(s, " ", "")
  s = Replace(s, "_", "")
  s = Replace(s, "-", "")

  Select Case s
    Case "F", "FOLLOWING"
      NormalizeBusinessDayConventionForSwap = "F"

    Case "MF", "MODIFIEDFOLLOWING"
      NormalizeBusinessDayConventionForSwap = "MF"

    Case "P", "PRECEDING"
      NormalizeBusinessDayConventionForSwap = "P"

    Case Else
      Err.Raise ERR_BASE + 13, _
                "clsOISswap.NormalizeBusinessDayConventionForSwap", _
                "BusinessDayConvention は F, Following, MF, ModifiedFollowing, P, Preceding のいずれかにしてください。"
  End Select

End Function

'============================================================
' 内部チェック：初期化済みか
'============================================================
Private Sub CheckInitialized()

  If Not mIsInitialized Then
    Err.Raise ERR_BASE + 8, _
              "clsOISswap", _
              "clsOISswap が初期化されていません。Init を先に実行してください。"
  End If

End Sub

'============================================================
' 内部チェック：カーブが渡されているか
'============================================================
Private Sub CheckCurve(ByVal in_cCurve As clsDiscountCurve)

  If in_cCurve Is Nothing Then
    Err.Raise ERR_BASE + 9, _
              "clsOISswap", _
              "カーブオブジェクトが指定されていません。"
  End If

End Sub
