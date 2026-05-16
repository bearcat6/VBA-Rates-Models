'============================================================
' mdl_Common
'
' 共通ユーティリティ関数
'============================================================

Option Explicit

'//
'// データの個数を行方向に数える
'//
Public Function GetvNumber(ByVal in_Range As Range) As Long
    GetvNumber = 1
    Do While in_Range.Cells(GetvNumber + 1, 1) <> ""
        GetvNumber = GetvNumber + 1
    Loop
    GetvNumber = GetvNumber - 1
End Function

'============================================================
' ACT/365 Fixed
'============================================================
Public Function YearFracAct365F(ByVal in_StartDate As Date, _
                                ByVal in_EndDate As Date) As Double

  If in_EndDate <= in_StartDate Then
      Err.Raise vbObjectError + 1001, _
              "mdl_Common.YearFracAct365F", _
              "EndDate は StartDate より後の日付にしてください。"
  End If

  YearFracAct365F = CDbl(in_EndDate - in_StartDate) / 365#

End Function
