
Option Explicit

'============================================================
' mdl_BusinessDay
'
' 役割：
'   ・営業日判定
'   ・営業日加算
'   ・Business Day Convention 調整
'   ・テナー文字列による日付計算
'
' 設計方針：
'   ・休日データは clsHolidayCalendar が保持する
'   ・土日判定と Business Day Convention は本モジュールで行う
'   ・月加算は DateAdd("m") を使わず、EDate ベースで行う
'   ・月末日からの月加算では、原則として月末維持を行う
'
' BusinessDayConvention:
'   "F"  / "Following"
'   "MF" / "ModifiedFollowing"
'   "P"  / "Preceding"
'
' Tenor:
'   "+1D", "-2D", "3M", "+6M", "1Y", "-1Y"
'
' 注意：
'   ・D は暦日加算
'   ・営業日加算は AddBusinessDays を使う
'   ・期間の絶対値は MAX_SHIFT 以下
'============================================================

Private Const MAX_SHIFT As Long = 600

'------------------------------------------------------------
' 営業日判定
'
' 土日でなく、かつ休日カレンダー上の休日でなければ True
'------------------------------------------------------------
Public Function IsBusinessDay(ByVal in_TargetDate As Date, _
                              Optional ByVal in_cCalendar As clsHolidayCalendar = Nothing) As Boolean

  Dim wd As VbDayOfWeek
  wd = Weekday(in_TargetDate, vbMonday)

  ' 土曜=6, 日曜=7
  If wd >= 6 Then
    IsBusinessDay = False
    Exit Function
  End If

  If Not in_cCalendar Is Nothing Then
    If in_cCalendar.IsHoliday(in_TargetDate) Then
      IsBusinessDay = False
      Exit Function
    End If
  End If

  IsBusinessDay = True

End Function

'------------------------------------------------------------
' 営業日数を加減算する
'
' in_BusinessDays:
'   +2 なら2営業日後
'   -2 なら2営業日前
'   0 なら基準日が営業日かどうかに関係なく基準日を返す
'
' 注意：
'   ・基準日はカウントしない
'------------------------------------------------------------
Public Function AddBusinessDays(ByVal in_StartDate As Date, _
                                ByVal in_BusinessDays As Long, _
                                Optional ByVal in_cCalendar As clsHolidayCalendar = Nothing) As Date

  Dim d As Date
  Dim stepDir As Long
  Dim cnt As Long

  If Abs(in_BusinessDays) > MAX_SHIFT Then
    Err.Raise vbObjectError + 1001, _
              "mdl_BusinessDay.AddBusinessDays", _
              "営業日数の絶対値は " & MAX_SHIFT & " 以下にしてください。"
  End If

  If in_BusinessDays = 0 Then
    AddBusinessDays = in_StartDate
    Exit Function
  End If

  If in_BusinessDays > 0 Then
    stepDir = 1
  Else
    stepDir = -1
  End If

  d = in_StartDate
  cnt = 0

  Do While cnt < Abs(in_BusinessDays)
    d = DateAdd("d", stepDir, d)

    If IsBusinessDay(d, in_cCalendar) Then
      cnt = cnt + 1
    End If
  Loop

  AddBusinessDays = d

End Function

'------------------------------------------------------------
' Business Day Convention を適用する
'
' Following:
'   非営業日なら翌営業日
'
' Modified Following:
'   非営業日なら翌営業日
'   ただし月をまたぐ場合は前営業日
'
' Preceding:
'   非営業日なら前営業日
'------------------------------------------------------------
Public Function AdjustBusinessDay(ByVal in_TargetDate As Date, _
                                  Optional ByVal in_BusinessDayConvention As String = "MF", _
                                  Optional ByVal in_cCalendar As clsHolidayCalendar = Nothing) As Date

  Dim conv As String
  Dim d As Date
  Dim followingDate As Date

  conv = NormalizeBusinessDayConvention(in_BusinessDayConvention)

  If IsBusinessDay(in_TargetDate, in_cCalendar) Then
    AdjustBusinessDay = in_TargetDate
    Exit Function
  End If

  Select Case conv

    Case "F"

      d = in_TargetDate

      Do
        d = DateAdd("d", 1, d)
      Loop Until IsBusinessDay(d, in_cCalendar)

      AdjustBusinessDay = d

    Case "MF"

      d = in_TargetDate

      Do
        d = DateAdd("d", 1, d)
      Loop Until IsBusinessDay(d, in_cCalendar)

      followingDate = d

      If Month(followingDate) = Month(in_TargetDate) Then
        AdjustBusinessDay = followingDate
      Else
        d = in_TargetDate

        Do
          d = DateAdd("d", -1, d)
        Loop Until IsBusinessDay(d, in_cCalendar)

        AdjustBusinessDay = d
      End If

    Case "P"

      d = in_TargetDate

      Do
        d = DateAdd("d", -1, d)
      Loop Until IsBusinessDay(d, in_cCalendar)

      AdjustBusinessDay = d

    Case Else

      Err.Raise vbObjectError + 1002, _
                "mdl_BusinessDay.AdjustBusinessDay", _
                "BusinessDayConvention は F, MF, P のいずれかにしてください。"

  End Select

End Function

'------------------------------------------------------------
' 月加算
'
' EDate を使って月加算する。
'
' keepEndOfMonth=True の場合、
' 基準日が月末であれば、加算後も月末に補正する。
'
' 例：
'   2024/1/31 + 1M = 2024/2/29
'   2024/1/31 + 2M = 2024/3/31
'   2024/2/29 + 1M = 2024/3/31
'
' keepEndOfMonth=False の場合、
' EDate の結果をそのまま返す。
'------------------------------------------------------------
Public Function AddTenorMonths(ByVal in_BaseDate As Date, _
                               ByVal in_Months As Long, _
                               Optional ByVal in_KeepEndOfMonth As Boolean = True) As Date

  Dim d As Date

  d = Application.WorksheetFunction.EDate(in_BaseDate, in_Months)

  If in_KeepEndOfMonth Then
    If in_BaseDate = Application.WorksheetFunction.EoMonth(in_BaseDate, 0) Then
      d = Application.WorksheetFunction.EoMonth(d, 0)
    End If
  End If

  AddTenorMonths = d

End Function

'------------------------------------------------------------
' テナー文字列で日付を進める
'
' 例：
'   AddTenor(#2025/1/31#, "3M")
'   AddTenor(#2025/1/31#, "+6M")
'   AddTenor(#2025/1/31#, "-1Y")
'
' in_Tenor:
'   "+1D", "-2D", "3M", "+6M", "1Y", "-1Y"
'
' 注意：
'   ・D は暦日加算
'   ・M/Y は AddTenorMonths を使う
'   ・営業日調整は行わない
'------------------------------------------------------------
Public Function AddTenor(ByVal in_BaseDate As Date, _
                         ByVal in_Tenor As String, _
                         Optional ByVal in_KeepEndOfMonth As Boolean = True) As Date

  Dim num As Long
  Dim unitChar As String

  Call ParseTenor(in_Tenor, num, unitChar)

  Select Case unitChar

    Case "D"
      AddTenor = DateAdd("d", num, in_BaseDate)

    Case "M"
      AddTenor = AddTenorMonths(in_BaseDate, num, in_KeepEndOfMonth)

    Case "Y"
      AddTenor = AddTenorMonths(in_BaseDate, num * 12, in_KeepEndOfMonth)

    Case Else
      Err.Raise vbObjectError + 1003, _
                "mdl_BusinessDay.AddTenor", _
                "テナー単位は D, M, Y のいずれかにしてください。"

  End Select

End Function

'------------------------------------------------------------
' テナー文字列で日付を進めたうえで、営業日調整する
'
' 例：
'   GetBusinessDay(#2025/1/31#, "3M", "MF", cal)
'   GetBusinessDay(#2025/1/31#, "+6M", "MF", cal)
'   GetBusinessDay(#2025/1/31#, "-1Y", "P", cal)
'
' in_Tenor:
'   "+1D", "-2D", "3M", "+6M", "1Y", "-1Y"
'
' in_BusinessDayConvention:
'   "F"  / "Following"
'   "MF" / "ModifiedFollowing"
'   "P"  / "Preceding"
'------------------------------------------------------------
Public Function GetBusinessDay(ByVal in_BaseDate As Date, _
                               ByVal in_Tenor As String, _
                               Optional ByVal in_BusinessDayConvention As String = "MF", _
                               Optional ByVal in_cCalendar As clsHolidayCalendar = Nothing, _
                               Optional ByVal in_KeepEndOfMonth As Boolean = True) As Date

  Dim unadjustedDate As Date

  unadjustedDate = AddTenor(in_BaseDate, in_Tenor, in_KeepEndOfMonth)

  GetBusinessDay = AdjustBusinessDay(unadjustedDate, _
                                     in_BusinessDayConvention, _
                                     in_cCalendar)

End Function

'------------------------------------------------------------
' 旧名互換：GetWorkday
'
' 旧コードとの互換用。
' 新規コードでは AddBusinessDays を使う。
'------------------------------------------------------------
Public Function GetWorkday(ByVal in_BaseDate As Date, _
                           ByVal in_BusinessDays As Long, _
                           Optional ByVal in_cCalendar As clsHolidayCalendar = Nothing) As Date

  GetWorkday = AddBusinessDays(in_BaseDate, _
                               in_BusinessDays, _
                               in_cCalendar)

End Function

'------------------------------------------------------------
' Business Day Convention 文字列を正規化する
'------------------------------------------------------------
Private Function NormalizeBusinessDayConvention(ByVal in_BusinessDayConvention As String) As String

  Dim s As String

  s = UCase$(Trim$(in_BusinessDayConvention))
  s = Replace(s, " ", "")
  s = Replace(s, "_", "")
  s = Replace(s, "-", "")

  Select Case s

    Case "F", "FOLLOWING"
      NormalizeBusinessDayConvention = "F"

    Case "MF", "MODIFIEDFOLLOWING"
      NormalizeBusinessDayConvention = "MF"

    Case "P", "PRECEDING"
      NormalizeBusinessDayConvention = "P"

    Case Else
      Err.Raise vbObjectError + 1004, _
                "mdl_BusinessDay.NormalizeBusinessDayConvention", _
                "BusinessDayConvention は F, Following, MF, ModifiedFollowing, P, Preceding のいずれかにしてください。"

  End Select

End Function

'------------------------------------------------------------
' テナー文字列を数値と単位に分解する
'
' 例：
'   "3M"   ->  3, "M"
'   "+3M"  ->  3, "M"
'   "-2D"  -> -2, "D"
'   "1Y"   ->  1, "Y"
'
' 制約：
'   ・単位は D, M, Y
'   ・数値部分は整数
'   ・絶対値は MAX_SHIFT 以下
'------------------------------------------------------------
Private Sub ParseTenor(ByVal in_Tenor As String, _
                       ByRef out_Number As Long, _
                       ByRef out_UnitChar As String)

  Dim s As String
  Dim numPart As String
  Dim unitPart As String

  s = UCase$(Trim$(in_Tenor))
  s = Replace(s, " ", "")

  If Len(s) < 2 Then
    Err.Raise vbObjectError + 1005, _
              "mdl_BusinessDay.ParseTenor", _
              "テナーは 3M, +6M, -1Y のように指定してください。"
  End If

  unitPart = Right$(s, 1)
  numPart = Left$(s, Len(s) - 1)

  If unitPart <> "D" And unitPart <> "M" And unitPart <> "Y" Then
    Err.Raise vbObjectError + 1006, _
              "mdl_BusinessDay.ParseTenor", _
              "テナー単位は D, M, Y のいずれかにしてください。"
  End If

  If Not IsNumeric(numPart) Then
    Err.Raise vbObjectError + 1007, _
              "mdl_BusinessDay.ParseTenor", _
              "テナーの数値部分が不正です。例：3M, +6M, -1Y"
  End If

  out_Number = CLng(numPart)

  If Abs(out_Number) > MAX_SHIFT Then
    Err.Raise vbObjectError + 1008, _
              "mdl_BusinessDay.ParseTenor", _
              "テナーの絶対値は " & MAX_SHIFT & " 以下にしてください。"
  End If

  out_UnitChar = unitPart

End Sub
