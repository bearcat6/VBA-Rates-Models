
Option Explicit

Public Sub TestBuildSwaptionSABR()

    Dim cCurve As clsOISStepForwardCurve
    Set cCurve = Get_OIS_Curve

    Dim params As Collection
    Set params = BuildSwaptionSABRFromSheet( _
                    ThisWorkbook.Worksheets("SwaptionVol"), _
                    cCurve, _
                    ThisWorkbook.Worksheets("SwaptionVol").Range("M4"))

    MsgBox "SABR calibration finished. Count = " & params.Count, vbInformation

End Sub

Public Sub TestOneSABRVol()

    Dim cCurve As clsOISStepForwardCurve
    Set cCurve = Get_OIS_Curve

    Dim params As Collection
    Set params = BuildSwaptionSABRFromSheet( _
                    ThisWorkbook.Worksheets("SwaptionVol"), _
                    cCurve, _
                    ThisWorkbook.Worksheets("SwaptionVol").Range("M4"))

    Dim p As clsSABRParams
    Set p = params("3M x 1Y")

    Debug.Print p.Expiry, p.Tenor, p.Forward, p.Alpha, p.Rho, p.Nu
    Debug.Print "ATM vol=", p.NormalVol(p.Forward)

End Sub
