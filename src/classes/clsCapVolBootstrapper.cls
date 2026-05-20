
Option Explicit

'============================================================
' clsCapVolBootstrapper
'
' Purpose:
'   Bootstrap caplet normal vols from quoted cap normal vols.
'
' Design:
'   - This class does not use clsCap.
'   - This class does not represent a Cap transaction.
'   - It directly generates caplet periods and prices caplet PVs.
'   - Caplet PV is calculated by mdl_CapFormula.CapletPV_Normal.
'   - Year fraction is ACT/365 Fixed via mdl_Common.YearFracAct365F.
'
' Notes:
'   - Normal vol is absolute value.
'       Example: 40bp normal vol = 0.0040
'   - The discount curve must expose:
'       DF(in_TargetDate)
'       ForwardRate(in_StartDate, in_EndDate)
'
' Bootstrap logic:
'   - For each quoted cap maturity, all newly added caplet periods
'     between the previous quote maturity and current quote maturity
'     are solved as one bucket vol.
'   - This avoids trying to solve multiple individual caplet vols from
'     one cap quote, which is underdetermined.
'============================================================

Private mValDate As Date
Private mStartDate As Date

Private mPaymentTenorMonths As Long
Private mStrike As Double
Private mNotional As Double

Private mCurve As Object
Private mCalendar As clsHolidayCalendar

Private mBusinessDayConvention As String
Private mPaymentLag As Long
Private mKeepEndOfMonth As Boolean

Private mRootTol As Double
Private mRootMaxIter As Long
Private mHighVolInitial As Double
Private mHighVolMax As Double

Private mIsInitialized As Boolean

'============================================================
' Initialization
'============================================================
Public Sub Init( _
  ByVal in_ValuationDate As Date, _
  ByVal in_StartDate As Date, _
  ByVal in_PaymentTenorMonths As Long, _
  ByVal in_StrikeRate As Double, _
  ByVal in_Notional As Double, _
  ByVal in_DiscountCurve As Object, _
  Optional ByVal in_Calendar As clsHolidayCalendar = Nothing, _
  Optional ByVal in_BusinessDayConvention As String = "MF", _
  Optional ByVal in_PaymentLag As Long = 2, _
  Optional ByVal in_KeepEndOfMonth As Boolean = True _
)

  If in_PaymentTenorMonths <= 0 Then
    Err.Raise vbObjectError + 5201, _
      "clsCapVolBootstrapper.Init", _
      "PaymentTenorMonths は1以上にしてください。"
  End If

  If in_Notional <= 0# Then
    Err.Raise vbObjectError + 5202, _
      "clsCapVolBootstrapper.Init", _
      "Notional は正の値にしてください。"
  End If

  If in_DiscountCurve Is Nothing Then
    Err.Raise vbObjectError + 5203, _
      "clsCapVolBootstrapper.Init", _
      "DiscountCurve が設定されていません。"
  End If

  If in_PaymentLag < 0 Then
    Err.Raise vbObjectError + 5204, _
      "clsCapVolBootstrapper.Init", _
      "PaymentLag は0以上にしてください。"
  End If

  mValDate = DateValue(in_ValuationDate)
  mStartDate = DateValue(in_StartDate)

  mPaymentTenorMonths = in_PaymentTenorMonths
  mStrike = in_StrikeRate
  mNotional = in_Notional

  Set mCurve = in_DiscountCurve
  Set mCalendar = in_Calendar

  mBusinessDayConvention = in_BusinessDayConvention
  mPaymentLag = in_PaymentLag
  mKeepEndOfMonth = in_KeepEndOfMonth

  mRootTol = 0.000001
  mRootMaxIter = 100

  ' 500bp normal vol
  mHighVolInitial = 0.05

  ' 5000bp normal vol
  mHighVolMax = 0.5

  mIsInitialized = True

End Sub

'============================================================
' Properties
'============================================================
Public Property Get ValuationDate() As Date
  CheckInitialized
  ValuationDate = mValDate
End Property

Public Property Get StartDate() As Date
  CheckInitialized
  StartDate = mStartDate
End Property

Public Property Get PaymentTenorMonths() As Long
  CheckInitialized
  PaymentTenorMonths = mPaymentTenorMonths
End Property

Public Property Get Strike() As Double
  CheckInitialized
  Strike = mStrike
End Property

Public Property Get Notional() As Double
  CheckInitialized
  Notional = mNotional
End Property

Public Property Let RootTolerance(ByVal in_RootTol As Double)

  If in_RootTol <= 0# Then
    Err.Raise vbObjectError + 5211, _
      "clsCapVolBootstrapper.RootTolerance", _
      "RootTolerance は正の値にしてください。"
  End If

  mRootTol = in_RootTol

End Property

Public Property Let RootMaxIterations(ByVal in_RootMaxIter As Long)

  If in_RootMaxIter <= 0 Then
    Err.Raise vbObjectError + 5212, _
      "clsCapVolBootstrapper.RootMaxIterations", _
      "RootMaxIterations は正の値にしてください。"
  End If

  mRootMaxIter = in_RootMaxIter

End Property

Public Property Let HighVolInitial(ByVal in_HighVolInitial As Double)

  If in_HighVolInitial <= 0# Then
    Err.Raise vbObjectError + 5213, _
      "clsCapVolBootstrapper.HighVolInitial", _
      "HighVolInitial は正の値にしてください。"
  End If

  mHighVolInitial = in_HighVolInitial

End Property

Public Property Let HighVolMax(ByVal in_HighVolMax As Double)

  If in_HighVolMax <= 0# Then
    Err.Raise vbObjectError + 5214, _
      "clsCapVolBootstrapper.HighVolMax", _
      "HighVolMax は正の値にしてください。"
  End If

  mHighVolMax = in_HighVolMax

End Property

'============================================================
' Bootstrap
'============================================================
Public Function Bootstrap( _
  ByRef in_CapMaturityDates() As Date, _
  ByRef in_CapNormalVols() As Double _
) As clsCapVolTermStructure

  Dim maturityDates() As Date
  Dim capVols() As Double
  Dim quoteCount As Long

  Dim cVolTS As clsCapVolTermStructure
  Dim i As Long

  CheckInitialized

  Call CopyAndSortQuotes( _
    in_CapMaturityDates, _
    in_CapNormalVols, _
    maturityDates, _
    capVols, _
    quoteCount)

  Set cVolTS = New clsCapVolTermStructure
  cVolTS.Init mValDate

  For i = 1 To quoteCount
    BootstrapOneCapQuote cVolTS, maturityDates(i), capVols(i)
  Next i

  Set Bootstrap = cVolTS

End Function

'============================================================
' Bootstrap one cap quote
'
' For one quoted cap maturity:
'   1. Generate all caplet periods up to maturity
'   2. Price the cap using flat quoted cap vol
'   3. Find periods not yet stored in vol term structure
'   4. Solve one bucket vol for all newly added periods
'   5. Store the solved bucket vol to all new periods
'============================================================
Private Sub BootstrapOneCapQuote( _
  ByVal in_cVolTS As clsCapVolTermStructure, _
  ByVal in_MaturityDate As Date, _
  ByVal in_CapNormalVol As Double _
)

  Dim periodCount As Long
  Dim periodStarts() As Date
  Dim periodEnds() As Date
  Dim fixingDates() As Date
  Dim paymentDates() As Date
  Dim isNewPeriod() As Boolean

  Dim targetPrice As Double
  Dim solvedVol As Double

  Dim i As Long
  Dim newCount As Long

  If in_CapNormalVol < 0# Then
    Err.Raise vbObjectError + 5221, _
      "clsCapVolBootstrapper.BootstrapOneCapQuote", _
      "Cap normal vol は0以上にしてください。"
  End If

  If DateValue(in_MaturityDate) <= mStartDate Then
    Err.Raise vbObjectError + 5222, _
      "clsCapVolBootstrapper.BootstrapOneCapQuote", _
      "MaturityDate は StartDate より後の日付にしてください。"
  End If

  Call GenerateCapletSchedule( _
    mStartDate, _
    DateValue(in_MaturityDate), _
    periodCount, _
    periodStarts, _
    periodEnds, _
    fixingDates, _
    paymentDates)

  If periodCount <= 0 Then
    Exit Sub
  End If

  ReDim isNewPeriod(1 To periodCount)

  newCount = 0

  For i = 1 To periodCount

    If fixingDates(i) <= mValDate Then
      isNewPeriod(i) = False
    ElseIf in_cVolTS.HasNormalVolByPeriodEnd(periodEnds(i)) Then
      isNewPeriod(i) = False
    Else
      isNewPeriod(i) = True
      newCount = newCount + 1
    End If

  Next i

  If newCount = 0 Then
    Exit Sub
  End If

  targetPrice = PriceCapletsWithFlatVol( _
                  periodCount, _
                  periodStarts, _
                  periodEnds, _
                  fixingDates, _
                  paymentDates, _
                  in_CapNormalVol)

  solvedVol = SolveNewBucketVol( _
                periodCount, _
                periodStarts, _
                periodEnds, _
                fixingDates, _
                paymentDates, _
                isNewPeriod, _
                in_cVolTS, _
                targetPrice)

  For i = 1 To periodCount

    If isNewPeriod(i) Then
      in_cVolTS.AddCapletVol _
        periodStarts(i), _
        periodEnds(i), _
        fixingDates(i), _
        paymentDates(i), _
        solvedVol
    End If

  Next i

End Sub

'============================================================
' Price caplets using one flat normal vol
'============================================================
Private Function PriceCapletsWithFlatVol( _
  ByVal in_PeriodCount As Long, _
  ByRef in_PeriodStarts() As Date, _
  ByRef in_PeriodEnds() As Date, _
  ByRef in_FixingDates() As Date, _
  ByRef in_PaymentDates() As Date, _
  ByVal in_NormalVol As Double _
) As Double

  Dim i As Long
  Dim pv As Double

  pv = 0#

  For i = 1 To in_PeriodCount

    If in_FixingDates(i) <= mValDate Then
      GoTo ContinueLoop
    End If

    pv = pv + CapletPVFromDates( _
                in_PeriodStarts(i), _
                in_PeriodEnds(i), _
                in_FixingDates(i), _
                in_PaymentDates(i), _
                in_NormalVol)

ContinueLoop:
  Next i

  PriceCapletsWithFlatVol = pv

End Function

'============================================================
' Price caplets using existing vols + one bucket vol
'
' in_IsNewPeriod(i) = True:
'   use in_NewVol
'
' in_IsNewPeriod(i) = False:
'   use existing vol term structure
'============================================================
Private Function PriceCapletsWithBucketVol( _
  ByVal in_PeriodCount As Long, _
  ByRef in_PeriodStarts() As Date, _
  ByRef in_PeriodEnds() As Date, _
  ByRef in_FixingDates() As Date, _
  ByRef in_PaymentDates() As Date, _
  ByRef in_IsNewPeriod() As Boolean, _
  ByVal in_cExistingVolTS As clsCapVolTermStructure, _
  ByVal in_NewVol As Double _
) As Double

  Dim i As Long
  Dim pv As Double
  Dim normalVol As Double

  pv = 0#

  For i = 1 To in_PeriodCount

    If in_FixingDates(i) <= mValDate Then
      GoTo ContinueLoop
    End If

    If in_IsNewPeriod(i) Then
      normalVol = in_NewVol
    Else
      normalVol = in_cExistingVolTS.GetNormalVolByPeriodEnd(in_PeriodEnds(i))
    End If

    pv = pv + CapletPVFromDates( _
                in_PeriodStarts(i), _
                in_PeriodEnds(i), _
                in_FixingDates(i), _
                in_PaymentDates(i), _
                normalVol)

ContinueLoop:
  Next i

  PriceCapletsWithBucketVol = pv

End Function

'============================================================
' Solve bucket vol by bisection
'============================================================
Private Function SolveNewBucketVol( _
  ByVal in_PeriodCount As Long, _
  ByRef in_PeriodStarts() As Date, _
  ByRef in_PeriodEnds() As Date, _
  ByRef in_FixingDates() As Date, _
  ByRef in_PaymentDates() As Date, _
  ByRef in_IsNewPeriod() As Boolean, _
  ByVal in_cExistingVolTS As clsCapVolTermStructure, _
  ByVal in_TargetPrice As Double _
) As Double

  Dim lowVol As Double
  Dim highVol As Double
  Dim midVol As Double

  Dim priceLow As Double
  Dim priceHigh As Double
  Dim priceMid As Double

  Dim k As Long

  lowVol = 0#
  highVol = mHighVolInitial

  priceLow = PriceCapletsWithBucketVol( _
              in_PeriodCount, _
              in_PeriodStarts, _
              in_PeriodEnds, _
              in_FixingDates, _
              in_PaymentDates, _
              in_IsNewPeriod, _
              in_cExistingVolTS, _
              lowVol)

  priceHigh = PriceCapletsWithBucketVol( _
                in_PeriodCount, _
                in_PeriodStarts, _
                in_PeriodEnds, _
                in_FixingDates, _
                in_PaymentDates, _
                in_IsNewPeriod, _
                in_cExistingVolTS, _
                highVol)

  Do While priceHigh < in_TargetPrice And highVol < mHighVolMax

    highVol = highVol * 2#

    If highVol > mHighVolMax Then
      highVol = mHighVolMax
    End If

    priceHigh = PriceCapletsWithBucketVol( _
                  in_PeriodCount, _
                  in_PeriodStarts, _
                  in_PeriodEnds, _
                  in_FixingDates, _
                  in_PaymentDates, _
                  in_IsNewPeriod, _
                  in_cExistingVolTS, _
                  highVol)

  Loop

  If in_TargetPrice < priceLow - mRootTol Then
    Err.Raise vbObjectError + 5231, _
      "clsCapVolBootstrapper.SolveNewBucketVol", _
      "Target price が lowVol=0 の価格を下回っています。入力volまたはstrikeを確認してください。"
  End If

  If priceHigh < in_TargetPrice - mRootTol Then
    Err.Raise vbObjectError + 5232, _
      "clsCapVolBootstrapper.SolveNewBucketVol", _
      "Upper vol でも target price に到達できません。HighVolMax を引き上げてください。"
  End If

  midVol = 0.5 * (lowVol + highVol)

  For k = 1 To mRootMaxIter

    midVol = 0.5 * (lowVol + highVol)

    priceMid = PriceCapletsWithBucketVol( _
                  in_PeriodCount, _
                  in_PeriodStarts, _
                  in_PeriodEnds, _
                  in_FixingDates, _
                  in_PaymentDates, _
                  in_IsNewPeriod, _
                  in_cExistingVolTS, _
                  midVol)

    If Abs(priceMid - in_TargetPrice) < mRootTol Then
      Exit For
    End If

    If priceMid > in_TargetPrice Then
      highVol = midVol
    Else
      lowVol = midVol
    End If

  Next k

  SolveNewBucketVol = midVol

End Function

'============================================================
' Caplet PV from dates
'============================================================
Private Function CapletPVFromDates( _
  ByVal in_PeriodStart As Date, _
  ByVal in_PeriodEnd As Date, _
  ByVal in_FixingDate As Date, _
  ByVal in_PaymentDate As Date, _
  ByVal in_NormalVol As Double _
) As Double

  Dim forwardRate As Double
  Dim expiryYear As Double
  Dim accrualYear As Double
  Dim df As Double

  forwardRate = mCurve.forwardRate(in_PeriodStart, in_PeriodEnd)
  expiryYear = YearFracAct365F(mValDate, in_FixingDate)
  accrualYear = YearFracAct365F(in_PeriodStart, in_PeriodEnd)
  df = mCurve.df(in_PaymentDate)

  CapletPVFromDates = CapletPV_Normal( _
                       forwardRate, _
                       mStrike, _
                       in_NormalVol, _
                       expiryYear, _
                       accrualYear, _
                       df, _
                       mNotional)

End Function

'============================================================
' Generate caplet schedule
'
' Output arrays are 1-based.
'
' Convention:
'   fixingDate  = periodStart
'   paymentDate = periodEnd + payment lag business days
'============================================================
Private Sub GenerateCapletSchedule( _
  ByVal in_StartDate As Date, _
  ByVal in_MaturityDate As Date, _
  ByRef out_PeriodCount As Long, _
  ByRef out_PeriodStarts() As Date, _
  ByRef out_PeriodEnds() As Date, _
  ByRef out_FixingDates() As Date, _
  ByRef out_PaymentDates() As Date _
)

  Dim periodStart As Date
  Dim periodEnd As Date
  Dim nextEnd As Date
  Dim fixingDate As Date
  Dim paymentDate As Date

  out_PeriodCount = 0

  Erase out_PeriodStarts
  Erase out_PeriodEnds
  Erase out_FixingDates
  Erase out_PaymentDates

  periodStart = DateValue(in_StartDate)
  in_MaturityDate = DateValue(in_MaturityDate)

  Do While periodStart < in_MaturityDate

    nextEnd = GetBusinessDay( _
                periodStart, _
                CStr(mPaymentTenorMonths) & "M", _
                mBusinessDayConvention, _
                mCalendar, _
                mKeepEndOfMonth)

    If nextEnd > in_MaturityDate Then
      periodEnd = in_MaturityDate
    Else
      periodEnd = nextEnd
    End If

    If periodEnd <= periodStart Then
      Err.Raise vbObjectError + 5241, _
        "clsCapVolBootstrapper.GenerateCapletSchedule", _
        "PeriodEnd が PeriodStart 以下になりました。日付条件を確認してください。"
    End If

    fixingDate = periodStart
    paymentDate = AddBusinessDays(periodEnd, mPaymentLag, mCalendar)

    out_PeriodCount = out_PeriodCount + 1

    ReDim Preserve out_PeriodStarts(1 To out_PeriodCount)
    ReDim Preserve out_PeriodEnds(1 To out_PeriodCount)
    ReDim Preserve out_FixingDates(1 To out_PeriodCount)
    ReDim Preserve out_PaymentDates(1 To out_PeriodCount)

    out_PeriodStarts(out_PeriodCount) = periodStart
    out_PeriodEnds(out_PeriodCount) = periodEnd
    out_FixingDates(out_PeriodCount) = fixingDate
    out_PaymentDates(out_PeriodCount) = paymentDate

    periodStart = periodEnd

  Loop

End Sub

'============================================================
' Copy and sort quote arrays
'============================================================
Private Sub CopyAndSortQuotes( _
  ByRef in_CapMaturityDates() As Date, _
  ByRef in_CapNormalVols() As Double, _
  ByRef out_MaturityDates() As Date, _
  ByRef out_CapNormalVols() As Double, _
  ByRef out_Count As Long _
)

  Dim lbD As Long
  Dim ubD As Long
  Dim lbV As Long
  Dim ubV As Long

  Dim i As Long
  Dim j As Long

  Dim tmpDate As Date
  Dim tmpVol As Double

  lbD = LBound(in_CapMaturityDates)
  ubD = UBound(in_CapMaturityDates)

  lbV = LBound(in_CapNormalVols)
  ubV = UBound(in_CapNormalVols)

  If (ubD - lbD) <> (ubV - lbV) Then
    Err.Raise vbObjectError + 5251, _
      "clsCapVolBootstrapper.CopyAndSortQuotes", _
      "CapMaturityDates と CapNormalVols の要素数が一致していません。"
  End If

  out_Count = ubD - lbD + 1

  If out_Count <= 0 Then
    Err.Raise vbObjectError + 5252, _
      "clsCapVolBootstrapper.CopyAndSortQuotes", _
      "Cap quote が空です。"
  End If

  ReDim out_MaturityDates(1 To out_Count)
  ReDim out_CapNormalVols(1 To out_Count)

  For i = 1 To out_Count

    out_MaturityDates(i) = DateValue(in_CapMaturityDates(lbD + i - 1))
    out_CapNormalVols(i) = in_CapNormalVols(lbV + i - 1)

    If out_MaturityDates(i) <= mStartDate Then
      Err.Raise vbObjectError + 5253, _
        "clsCapVolBootstrapper.CopyAndSortQuotes", _
        "Cap maturity date は StartDate より後の日付にしてください。"
    End If

    If out_CapNormalVols(i) < 0# Then
      Err.Raise vbObjectError + 5254, _
        "clsCapVolBootstrapper.CopyAndSortQuotes", _
        "Cap normal vol は0以上にしてください。"
    End If

  Next i

  ' Simple ascending sort by maturity date
  For i = 1 To out_Count - 1
    For j = i + 1 To out_Count

      If out_MaturityDates(j) < out_MaturityDates(i) Then

        tmpDate = out_MaturityDates(i)
        out_MaturityDates(i) = out_MaturityDates(j)
        out_MaturityDates(j) = tmpDate

        tmpVol = out_CapNormalVols(i)
        out_CapNormalVols(i) = out_CapNormalVols(j)
        out_CapNormalVols(j) = tmpVol

      End If

    Next j
  Next i

  For i = 2 To out_Count
    If out_MaturityDates(i) = out_MaturityDates(i - 1) Then
      Err.Raise vbObjectError + 5255, _
        "clsCapVolBootstrapper.CopyAndSortQuotes", _
        "同じ maturity date の quote が重複しています: " & CStr(out_MaturityDates(i))
    End If
  Next i

End Sub

'============================================================
' Internal checks
'============================================================
Private Sub CheckInitialized()

  If Not mIsInitialized Then
    Err.Raise vbObjectError + 5291, _
      "clsCapVolBootstrapper", _
      "Init が実行されていません。"
  End If

End Sub

