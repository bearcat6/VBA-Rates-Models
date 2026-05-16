
Option Explicit

Sub Test_CapVolBootstrap()

  Dim valDate As Date
  valDate = DateSerial(2026, 5, 11)

  Dim startDate As Date
  startDate = valDate

  Dim cCurve As clsOISStepForwardCurve
  Set cCurve = Get_OIS_Curve

  Dim maturities(1 To 4) As Date
  Dim vols(1 To 4) As Double

  maturities(1) = GetBusinessDay(startDate, "1Y", "MF", Nothing)
  maturities(2) = GetBusinessDay(startDate, "2Y", "MF", Nothing)
  maturities(3) = GetBusinessDay(startDate, "3Y", "MF", Nothing)
  maturities(4) = GetBusinessDay(startDate, "5Y", "MF", Nothing)

  vols(1) = 0.004
  vols(2) = 0.0045
  vols(3) = 0.005
  vols(4) = 0.0055

  Dim bs As clsCapVolBootstrapper
  Set bs = New clsCapVolBootstrapper

  bs.Init _
    valDate, _
    startDate, _
    6, _
    0.005, _
    100000000#, _
    cCurve, _
    Nothing, _
    "MF", _
    2, _
    True

  Dim volTS As clsCapVolTermStructure
  Set volTS = bs.Bootstrap(maturities, vols)

  volTS.DumpToSheet "CapVol", "A1"

End Sub

