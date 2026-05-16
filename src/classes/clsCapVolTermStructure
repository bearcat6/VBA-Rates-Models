Option Explicit

'============================================================
' clsCapVolTermStructure
'
' Purpose:
'   Stores bootstrapped caplet normal volatility term structure.
'
' Design:
'   - This class does not represent a Cap transaction.
'   - It stores period-by-period caplet normal vols.
'   - Normal vols are absolute values.
'       Example: 40bp normal vol = 0.0040
'   - Year fraction is ACT/365 Fixed via mdl_Common.YearFracAct365F.
'============================================================

Private mValDate As Date

Private mPeriodStartDates() As Date
Private mPeriodEndDates() As Date
Private mFixingDates() As Date
Private mPaymentDates() As Date
Private mNormalVols() As Double

Private mCount As Long
Private mIsInitialized As Boolean

'============================================================
' Initialization
'============================================================
Public Sub Init(ByVal in_ValuationDate As Date)

  mValDate = DateValue(in_ValuationDate)

  mCount = 0

  Erase mPeriodStartDates
  Erase mPeriodEndDates
  Erase mFixingDates
  Erase mPaymentDates
  Erase mNormalVols

  mIsInitialized = True

End Sub

'============================================================
' Properties
'============================================================
Public Property Get ValuationDate() As Date
  CheckInitialized
  ValuationDate = mValDate
End Property

Public Property Get Count() As Long
  CheckInitialized
  Count = mCount
End Property

'============================================================
' Add caplet normal vol
'============================================================
Public Sub AddCapletVol( _
  ByVal in_PeriodStartDate As Date, _
  ByVal in_PeriodEndDate As Date, _
  ByVal in_FixingDate As Date, _
  ByVal in_PaymentDate As Date, _
  ByVal in_NormalVol As Double _
)

  CheckInitialized

  in_PeriodStartDate = DateValue(in_PeriodStartDate)
  in_PeriodEndDate = DateValue(in_PeriodEndDate)
  in_FixingDate = DateValue(in_FixingDate)
  in_PaymentDate = DateValue(in_PaymentDate)

  If in_PeriodEndDate <= in_PeriodStartDate Then
    Err.Raise vbObjectError + 5101, _
      "clsCapVolTermStructure.AddCapletVol", _
      "PeriodEndDate は PeriodStartDate より後の日付にしてください。"
  End If

  If in_NormalVol < 0# Then
    Err.Raise vbObjectError + 5102, _
      "clsCapVolTermStructure.AddCapletVol", _
      "NormalVol は0以上にしてください。"
  End If

  If HasNormalVolByPeriodEnd(in_PeriodEndDate) Then
    Err.Raise vbObjectError + 5103, _
      "clsCapVolTermStructure.AddCapletVol", _
      "同じ PeriodEndDate の caplet vol が既に登録されています: " & CStr(in_PeriodEndDate)
  End If

  mCount = mCount + 1

  ReDim Preserve mPeriodStartDates(1 To mCount)
  ReDim Preserve mPeriodEndDates(1 To mCount)
  ReDim Preserve mFixingDates(1 To mCount)
  ReDim Preserve mPaymentDates(1 To mCount)
  ReDim Preserve mNormalVols(1 To mCount)

  mPeriodStartDates(mCount) = in_PeriodStartDate
  mPeriodEndDates(mCount) = in_PeriodEndDate
  mFixingDates(mCount) = in_FixingDate
  mPaymentDates(mCount) = in_PaymentDate
  mNormalVols(mCount) = in_NormalVol

End Sub

'============================================================
' Existence check
'============================================================
Public Function HasNormalVolByPeriodEnd( _
  ByVal in_PeriodEndDate As Date _
) As Boolean

  Dim i As Long

  CheckInitialized

  in_PeriodEndDate = DateValue(in_PeriodEndDate)

  For i = 1 To mCount
    If mPeriodEndDates(i) = in_PeriodEndDate Then
      HasNormalVolByPeriodEnd = True
      Exit Function
    End If
  Next i

  HasNormalVolByPeriodEnd = False

End Function

'============================================================
' Get normal vol
'============================================================
Public Function GetNormalVolByPeriodEnd( _
  ByVal in_PeriodEndDate As Date _
) As Double

  Dim i As Long

  CheckInitialized

  in_PeriodEndDate = DateValue(in_PeriodEndDate)

  For i = 1 To mCount
    If mPeriodEndDates(i) = in_PeriodEndDate Then
      GetNormalVolByPeriodEnd = mNormalVols(i)
      Exit Function
    End If
  Next i

  Err.Raise vbObjectError + 5111, _
    "clsCapVolTermStructure.GetNormalVolByPeriodEnd", _
    "No caplet volatility found for period end date: " & CStr(in_PeriodEndDate)

End Function

Public Function GetNormalVolByFixingDate( _
  ByVal in_FixingDate As Date _
) As Double

  Dim i As Long

  CheckInitialized

  in_FixingDate = DateValue(in_FixingDate)

  For i = 1 To mCount
    If mFixingDates(i) = in_FixingDate Then
      GetNormalVolByFixingDate = mNormalVols(i)
      Exit Function
    End If
  Next i

  Err.Raise vbObjectError + 5112, _
    "clsCapVolTermStructure.GetNormalVolByFixingDate", _
    "No caplet volatility found for fixing date: " & CStr(in_FixingDate)

End Function

Public Function GetNormalVolByIndex( _
  ByVal in_Index As Long _
) As Double

  CheckInitialized
  CheckIndex in_Index

  GetNormalVolByIndex = mNormalVols(in_Index)

End Function

'============================================================
' Get dates
'============================================================
Public Function GetPeriodStartDate( _
  ByVal in_Index As Long _
) As Date

  CheckInitialized
  CheckIndex in_Index

  GetPeriodStartDate = mPeriodStartDates(in_Index)

End Function

Public Function GetPeriodEndDate( _
  ByVal in_Index As Long _
) As Date

  CheckInitialized
  CheckIndex in_Index

  GetPeriodEndDate = mPeriodEndDates(in_Index)

End Function

Public Function GetFixingDate( _
  ByVal in_Index As Long _
) As Date

  CheckInitialized
  CheckIndex in_Index

  GetFixingDate = mFixingDates(in_Index)

End Function

Public Function GetPaymentDate( _
  ByVal in_Index As Long _
) As Date

  CheckInitialized
  CheckIndex in_Index

  GetPaymentDate = mPaymentDates(in_Index)

End Function

'============================================================
' Total variance
'
' TotalVariance = Sum( normalVol_i ^ 2 * yearFraction_i )
'
' The period is truncated at targetDate if targetDate falls inside
' a registered caplet period.
'============================================================
Public Function TotalVarianceToDate( _
  ByVal in_TargetDate As Date _
) As Double

  Dim i As Long
  Dim t0 As Double
  Dim t1 As Double
  Dim v As Double

  CheckInitialized

  in_TargetDate = DateValue(in_TargetDate)

  If in_TargetDate <= mValDate Then
    TotalVarianceToDate = 0#
    Exit Function
  End If

  v = 0#

  For i = 1 To mCount

    If mPeriodStartDates(i) >= in_TargetDate Then
      Exit For
    End If

    t0 = YearFracAct365F(mValDate, mPeriodStartDates(i))
    t1 = YearFracAct365F(mValDate, mPeriodEndDates(i))

    If mPeriodEndDates(i) > in_TargetDate Then
      t1 = YearFracAct365F(mValDate, in_TargetDate)
    End If

    If t1 > t0 Then
      v = v + mNormalVols(i) * mNormalVols(i) * (t1 - t0)
    End If

    If mPeriodEndDates(i) >= in_TargetDate Then
      Exit For
    End If

  Next i

  TotalVarianceToDate = v

End Function

'============================================================
' Dump table
'============================================================
Public Function GetTable() As Variant

  Dim arr() As Variant
  Dim i As Long

  CheckInitialized

  ReDim arr(1 To mCount + 1, 1 To 8)

  arr(1, 1) = "No"
  arr(1, 2) = "Period Start"
  arr(1, 3) = "Period End"
  arr(1, 4) = "Fixing Date"
  arr(1, 5) = "Payment Date"
  arr(1, 6) = "YearFrac"
  arr(1, 7) = "Normal Vol"
  arr(1, 8) = "Normal Vol(bp)"

  For i = 1 To mCount
    arr(i + 1, 1) = i
    arr(i + 1, 2) = mPeriodStartDates(i)
    arr(i + 1, 3) = mPeriodEndDates(i)
    arr(i + 1, 4) = mFixingDates(i)
    arr(i + 1, 5) = mPaymentDates(i)
    arr(i + 1, 6) = YearFracAct365F(mPeriodStartDates(i), mPeriodEndDates(i))
    arr(i + 1, 7) = mNormalVols(i)
    arr(i + 1, 8) = mNormalVols(i) * 10000#
  Next i

  GetTable = arr

End Function

Public Sub DumpToSheet( _
  ByVal in_SheetName As String, _
  Optional ByVal in_StartCell As String = "A1" _
)

  Dim ws As Worksheet
  Dim arr As Variant

  CheckInitialized

  Set ws = ThisWorkbook.Worksheets(in_SheetName)

  arr = GetTable()

  ws.Range(in_StartCell).Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr

End Sub

'============================================================
' Internal checks
'============================================================
Private Sub CheckInitialized()

  If Not mIsInitialized Then
    Err.Raise vbObjectError + 5191, _
      "clsCapVolTermStructure", _
      "Init が実行されていません。"
  End If

End Sub

Private Sub CheckIndex(ByVal in_Index As Long)

  If in_Index < 1 Or in_Index > mCount Then
    Err.Raise vbObjectError + 5192, _
      "clsCapVolTermStructure", _
      "Index out of range."
  End If

End Sub
