
Option Explicit

' ============================================================
' Swaption Vol Sheet -> beta=0 Normal SABR parameters
'
' 前提:
'   - SwaptionVol sheet:
'       B列: "3M x 1Y" のような Expiry x Tenor
'       C:K列: -200bps, -100bps, ... ATM, ... 200bps
'       数値: Normal Vol。Excelの%表示なら 0.5077% = 0.005077 として読む
'   - 金利カーブ:
'       clsOISStepForwardCurve.ForwardParRate(Expiry, Tenor) As Double
'
' 使い方:
'   Public Sub TestBuildSwaptionSABR()
'       Dim cCurve As clsOISStepForwardCurve
'       Set cCurve = Get_OIS_Curve
'
'       Dim params As Collection
'       Set params = BuildSwaptionSABRFromSheet( _
'                       ThisWorkbook.Worksheets("SwaptionVol"), _
'                       cCurve, _
'                       ThisWorkbook.Worksheets("SwaptionVol").Range("M4"))
'   End Sub
' ============================================================

Public Function BuildSwaptionSABRFromSheet(ByVal in_Ws As Worksheet, _
                                           ByVal in_CCurve As clsOISStepForwardCurve, _
                                           Optional ByVal in_OutTopLeft As Range = Nothing, _
                                           Optional ByVal in_FirstDataRow As Long = 5, _
                                           Optional ByVal in_LastDataRow As Long = 0, _
                                           Optional ByVal in_LabelCol As Long = 2, _
                                           Optional ByVal in_FirstVolCol As Long = 3, _
                                           Optional ByVal in_LastVolCol As Long = 11, _
                                           Optional ByVal in_HeaderRow As Long = 4) As Collection
    Dim results As New Collection
    Dim r As Long, c As Long, n As Long
    Dim labelText As String, expiry As String, tenor As String
    Dim fwd As Double, tExp As Double
    Dim shifts() As Double, vols() As Double
    Dim p As clsSABRParams

    If in_LastDataRow = 0 Then
        in_LastDataRow = in_Ws.Cells(in_Ws.Rows.Count, in_LabelCol).End(xlUp).Row
    End If

    n = in_LastVolCol - in_FirstVolCol + 1
    ReDim shifts(1 To n)
    ReDim vols(1 To n)

    For c = in_FirstVolCol To in_LastVolCol
        shifts(c - in_FirstVolCol + 1) = HeaderToShiftBps(CStr(in_Ws.Cells(in_HeaderRow, c).Value))
    Next c

    If Not in_OutTopLeft Is Nothing Then
        WriteSABRHeader in_OutTopLeft
    End If

    For r = in_FirstDataRow To in_LastDataRow
        labelText = Trim$(CStr(in_Ws.Cells(r, in_LabelCol).Value))
        If Len(labelText) > 0 Then
            ParseExpiryTenor labelText, expiry, tenor
            tExp = PeriodToYears(expiry)

            'ここがポイント：
            'clsOISStepForwardCurve側の ForwardParRate(Expiry, Tenor) を利用する
            fwd = in_CCurve.ForwardParRate(expiry, tenor)

            For c = in_FirstVolCol To in_LastVolCol
                vols(c - in_FirstVolCol + 1) = NormalizeVol(CDbl(in_Ws.Cells(r, c).Value))
            Next c

            Set p = CalibrateSABR_Beta0(expiry, tenor, fwd, tExp, shifts, vols)
            results.Add p, expiry & " x " & tenor

            If Not in_OutTopLeft Is Nothing Then
                WriteSABRRow in_OutTopLeft.Offset(r - in_FirstDataRow + 1, 0), p
            End If
        End If
    Next r

    Set BuildSwaptionSABRFromSheet = results
End Function

Public Function CalibrateSABR_Beta0(ByVal in_Expiry As String, _
                                    ByVal in_Tenor As String, _
                                    ByVal in_Fwd As Double, _
                                    ByVal in_TExp As Double, _
                                    ByRef in_ShiftsBps() As Double, _
                                    ByRef in_MarketVols() As Double) As clsSABRParams
    Dim atmVol As Double
    atmVol = GetATMVol(in_ShiftsBps, in_MarketVols)

    Dim seedAlpha(1 To 3) As Double
    Dim seedRho(1 To 5) As Double
    Dim seedNu(1 To 5) As Double

    seedAlpha(1) = Application.Max(0.000001, atmVol * 0.75)
    seedAlpha(2) = Application.Max(0.000001, atmVol)
    seedAlpha(3) = Application.Max(0.000001, atmVol * 1.25)

    seedRho(1) = -0.75
    seedRho(2) = -0.35
    seedRho(3) = 0#
    seedRho(4) = 0.35
    seedRho(5) = 0.75

    seedNu(1) = 0.05
    seedNu(2) = 0.15
    seedNu(3) = 0.35
    seedNu(4) = 0.7
    seedNu(5) = 1.2

    Dim bestObj As Double, obj As Double
    Dim bestY(0 To 2) As Double, y0(0 To 2) As Double, yOpt As Variant
    Dim ia As Long, ir As Long, inu As Long, iters As Long, bestIters As Long

    bestObj = 1E+99

    For ia = LBound(seedAlpha) To UBound(seedAlpha)
        For ir = LBound(seedRho) To UBound(seedRho)
            For inu = LBound(seedNu) To UBound(seedNu)
                y0(0) = Log(seedAlpha(ia))
                y0(1) = AtanhSafe(seedRho(ir))
                y0(2) = Log(seedNu(inu))

                yOpt = NelderMead3_Beta0(in_Fwd, in_TExp, in_ShiftsBps, in_MarketVols, y0, obj, iters)

                If obj < bestObj Then
                    bestObj = obj
                    bestY(0) = yOpt(0)
                    bestY(1) = yOpt(1)
                    bestY(2) = yOpt(2)
                    bestIters = iters
                End If
            Next inu
        Next ir
    Next ia

    Dim alpha As Double, rho As Double, nu As Double
    UnpackParams bestY, alpha, rho, nu

    Dim p As New clsSABRParams
    p.Expiry = in_Expiry
    p.Tenor = in_Tenor
    p.Forward = in_Fwd
    p.ExpiryYears = in_TExp
    p.Alpha = alpha
    p.Beta = 0#
    p.Rho = rho
    p.Nu = nu
    p.RMSE = CalcRMSE_Beta0(in_Fwd, in_TExp, in_ShiftsBps, in_MarketVols, alpha, rho, nu)
    p.MaxAbsError = CalcMaxAbsError_Beta0(in_Fwd, in_TExp, in_ShiftsBps, in_MarketVols, alpha, rho, nu)
    p.Iterations = bestIters
    p.Success = (p.RMSE < 0.0005) '5bp程度。必要なら判定基準は調整

    Set CalibrateSABR_Beta0 = p
End Function

Public Function SABR_NormalVol_Beta0(ByVal in_Fwd As Double, _
                                     ByVal in_Strike As Double, _
                                     ByVal in_TExp As Double, _
                                     ByVal in_Alpha As Double, _
                                     ByVal in_Rho As Double, _
                                     ByVal in_Nu As Double) As Double
    Const EPS As Double = 0.0000000001

    If in_Alpha <= 0# Or in_Nu < 0# Or Abs(in_Rho) >= 1# Then
        SABR_NormalVol_Beta0 = CVErr(xlErrNum)
        Exit Function
    End If

    Dim z As Double, xz As Double, ratio As Double, tmp As Double
    z = (in_Nu / in_Alpha) * (in_Fwd - in_Strike)

    If Abs(z) < EPS Then
        ratio = 1#
    Else
        tmp = Sqr(1# - 2# * in_Rho * z + z * z) + z - in_Rho
        If tmp <= 0# Or (1# - in_Rho) <= 0# Then
            SABR_NormalVol_Beta0 = CVErr(xlErrNum)
            Exit Function
        End If
        xz = Log(tmp / (1# - in_Rho))

        If Abs(xz) < EPS Then
            ratio = 1#
        Else
            ratio = z / xz
        End If
    End If

    SABR_NormalVol_Beta0 = in_Alpha * ratio * (1# + ((2# - 3# * in_Rho * in_Rho) * in_Nu * in_Nu * in_TExp / 24#))
End Function

Private Function NelderMead3_Beta0(ByVal in_Fwd As Double, _
                                   ByVal in_TExp As Double, _
                                   ByRef in_ShiftsBps() As Double, _
                                   ByRef in_MarketVols() As Double, _
                                   ByRef in_Y0() As Double, _
                                   ByRef out_BestObj As Double, _
                                   ByRef out_Iterations As Long) As Variant
    Const N As Long = 3
    Const MAX_ITER As Long = 300
    Const TOL As Double = 0.000000000001

    Dim simplex(0 To N, 0 To N - 1) As Double
    Dim val(0 To N) As Double
    Dim i As Long, j As Long

    For j = 0 To N - 1
        simplex(0, j) = in_Y0(j)
    Next j

    For i = 1 To N
        For j = 0 To N - 1
            simplex(i, j) = in_Y0(j)
        Next j
        simplex(i, i - 1) = simplex(i, i - 1) + 0.15
    Next i

    For i = 0 To N
        val(i) = Objective_Beta0(in_Fwd, in_TExp, in_ShiftsBps, in_MarketVols, simplex, i)
    Next i

    Dim lo As Long, hi As Long, nhi As Long
    Dim centroid(0 To N - 1) As Double
    Dim xr(0 To N - 1) As Double, xe(0 To N - 1) As Double, xc(0 To N - 1) As Double
    Dim fr As Double, fe As Double, fc As Double
    Dim iter As Long, spread As Double

    For iter = 1 To MAX_ITER
        SortSimplex val, simplex, lo, hi, nhi

        spread = 0#
        For i = 0 To N
            spread = spread + Abs(val(i) - val(lo))
        Next i
        If spread < TOL Then Exit For

        For j = 0 To N - 1
            centroid(j) = 0#
            For i = 0 To N
                If i <> hi Then centroid(j) = centroid(j) + simplex(i, j)
            Next i
            centroid(j) = centroid(j) / N
        Next j

        ' reflection
        For j = 0 To N - 1
            xr(j) = centroid(j) + (centroid(j) - simplex(hi, j))
        Next j
        fr = ObjectiveVector_Beta0(in_Fwd, in_TExp, in_ShiftsBps, in_MarketVols, xr)

        If fr < val(lo) Then
            ' expansion
            For j = 0 To N - 1
                xe(j) = centroid(j) + 2# * (xr(j) - centroid(j))
            Next j
            fe = ObjectiveVector_Beta0(in_Fwd, in_TExp, in_ShiftsBps, in_MarketVols, xe)

            If fe < fr Then
                ReplacePoint simplex, val, hi, xe, fe
            Else
                ReplacePoint simplex, val, hi, xr, fr
            End If

        ElseIf fr < val(nhi) Then
            ReplacePoint simplex, val, hi, xr, fr

        Else
            ' contraction
            If fr < val(hi) Then
                For j = 0 To N - 1
                    xc(j) = centroid(j) + 0.5 * (xr(j) - centroid(j))
                Next j
            Else
                For j = 0 To N - 1
                    xc(j) = centroid(j) + 0.5 * (simplex(hi, j) - centroid(j))
                Next j
            End If

            fc = ObjectiveVector_Beta0(in_Fwd, in_TExp, in_ShiftsBps, in_MarketVols, xc)

            If fc < val(hi) Then
                ReplacePoint simplex, val, hi, xc, fc
            Else
                ' shrink
                For i = 0 To N
                    If i <> lo Then
                        For j = 0 To N - 1
                            simplex(i, j) = simplex(lo, j) + 0.5 * (simplex(i, j) - simplex(lo, j))
                        Next j
                        val(i) = Objective_Beta0(in_Fwd, in_TExp, in_ShiftsBps, in_MarketVols, simplex, i)
                    End If
                Next i
            End If
        End If
    Next iter

    SortSimplex val, simplex, lo, hi, nhi

    Dim ans(0 To N - 1) As Double
    For j = 0 To N - 1
        ans(j) = simplex(lo, j)
    Next j

    out_BestObj = val(lo)
    out_Iterations = iter
    NelderMead3_Beta0 = ans
End Function

Private Function Objective_Beta0(ByVal in_Fwd As Double, _
                                 ByVal in_TExp As Double, _
                                 ByRef in_ShiftsBps() As Double, _
                                 ByRef in_MarketVols() As Double, _
                                 ByRef in_Simplex() As Double, _
                                 ByVal in_Idx As Long) As Double
    Dim y(0 To 2) As Double
    y(0) = in_Simplex(in_Idx, 0)
    y(1) = in_Simplex(in_Idx, 1)
    y(2) = in_Simplex(in_Idx, 2)
    Objective_Beta0 = ObjectiveVector_Beta0(in_Fwd, in_TExp, in_ShiftsBps, in_MarketVols, y)
End Function

Private Function ObjectiveVector_Beta0(ByVal in_Fwd As Double, _
                                       ByVal in_TExp As Double, _
                                       ByRef in_ShiftsBps() As Double, _
                                       ByRef in_MarketVols() As Double, _
                                       ByRef in_Y() As Double) As Double
    Dim alpha As Double, rho As Double, nu As Double
    UnpackParams in_Y, alpha, rho, nu

    If alpha <= 0# Or nu < 0# Or Abs(rho) >= 0.999999 Then
        ObjectiveVector_Beta0 = 1E+99
        Exit Function
    End If

    Dim i As Long, k As Double, mdl As Double, e As Double, sse As Double
    For i = LBound(in_ShiftsBps) To UBound(in_ShiftsBps)
        k = in_Fwd + in_ShiftsBps(i) / 10000#
        mdl = SABR_NormalVol_Beta0(in_Fwd, k, in_TExp, alpha, rho, nu)
        e = mdl - in_MarketVols(i)
        sse = sse + e * e
    Next i

    ObjectiveVector_Beta0 = sse
End Function

Private Function CalcRMSE_Beta0(ByVal in_Fwd As Double, _
                                ByVal in_TExp As Double, _
                                ByRef in_ShiftsBps() As Double, _
                                ByRef in_MarketVols() As Double, _
                                ByVal in_Alpha As Double, _
                                ByVal in_Rho As Double, _
                                ByVal in_Nu As Double) As Double
    Dim i As Long, n As Long, k As Double, mdl As Double, e As Double, sse As Double

    For i = LBound(in_ShiftsBps) To UBound(in_ShiftsBps)
        k = in_Fwd + in_ShiftsBps(i) / 10000#
        mdl = SABR_NormalVol_Beta0(in_Fwd, k, in_TExp, in_Alpha, in_Rho, in_Nu)
        e = mdl - in_MarketVols(i)
        sse = sse + e * e
        n = n + 1
    Next i

    CalcRMSE_Beta0 = Sqr(sse / n)
End Function

Private Function CalcMaxAbsError_Beta0(ByVal in_Fwd As Double, _
                                       ByVal in_TExp As Double, _
                                       ByRef in_ShiftsBps() As Double, _
                                       ByRef in_MarketVols() As Double, _
                                       ByVal in_Alpha As Double, _
                                       ByVal in_Rho As Double, _
                                       ByVal in_Nu As Double) As Double
    Dim i As Long, k As Double, mdl As Double, e As Double, mx As Double

    For i = LBound(in_ShiftsBps) To UBound(in_ShiftsBps)
        k = in_Fwd + in_ShiftsBps(i) / 10000#
        mdl = SABR_NormalVol_Beta0(in_Fwd, k, in_TExp, in_Alpha, in_Rho, in_Nu)
        e = Abs(mdl - in_MarketVols(i))
        If e > mx Then mx = e
    Next i

    CalcMaxAbsError_Beta0 = mx
End Function

Private Sub WriteSABRHeader(ByVal in_TopLeft As Range)
    in_TopLeft.Resize(1, 12).Value = Array( _
        "Expiry", "Tenor", "Forward", "Alpha", "Beta", "Rho", "Nu", _
        "RMSE", "RMSE(bp)", "MaxAbsErr", "MaxAbsErr(bp)", "Success")
End Sub

Private Sub WriteSABRRow(ByVal in_TopLeft As Range, ByVal in_P As clsSABRParams)
    in_TopLeft.Offset(0, 0).Value = in_P.Expiry
    in_TopLeft.Offset(0, 1).Value = in_P.Tenor
    in_TopLeft.Offset(0, 2).Value = in_P.Forward
    in_TopLeft.Offset(0, 3).Value = in_P.Alpha
    in_TopLeft.Offset(0, 4).Value = in_P.Beta
    in_TopLeft.Offset(0, 5).Value = in_P.Rho
    in_TopLeft.Offset(0, 6).Value = in_P.Nu
    in_TopLeft.Offset(0, 7).Value = in_P.RMSE
    in_TopLeft.Offset(0, 8).Value = in_P.RMSE * 10000#
    in_TopLeft.Offset(0, 9).Value = in_P.MaxAbsError
    in_TopLeft.Offset(0, 10).Value = in_P.MaxAbsError * 10000#
    in_TopLeft.Offset(0, 11).Value = in_P.Success
End Sub

Private Sub ParseExpiryTenor(ByVal in_LabelText As String, _
                             ByRef out_Expiry As String, _
                             ByRef out_Tenor As String)
    Dim s As String, a As Variant
    s = UCase$(Replace(in_LabelText, " ", ""))
    a = Split(s, "X")
    If UBound(a) <> 1 Then Err.Raise vbObjectError + 100, , "Expiry x Tenor の形式ではありません: " & in_LabelText

    out_Expiry = CStr(a(0))
    out_Tenor = CStr(a(1))
End Sub

Private Function HeaderToShiftBps(ByVal in_HeaderText As String) As Double
    Dim s As String
    s = UCase$(Trim$(in_HeaderText))

    If s = "ATM" Then
        HeaderToShiftBps = 0#
        Exit Function
    End If

    s = Replace(s, "BPS", "")
    s = Replace(s, "BP", "")
    s = Replace(s, "+", "")
    s = Trim$(s)

    If Len(s) = 0 Or Not IsNumeric(s) Then
        Err.Raise vbObjectError + 101, , "bpsヘッダーを読めません: " & in_HeaderText
    End If

    HeaderToShiftBps = CDbl(s)
End Function

Private Function PeriodToYears(ByVal in_PeriodText As String) As Double
    Dim s As String, unit As String, num As Double
    s = UCase$(Trim$(in_PeriodText))

    unit = Right$(s, 1)
    num = CDbl(Left$(s, Len(s) - 1))

    Select Case unit
        Case "D"
            PeriodToYears = num / 365#
        Case "W"
            PeriodToYears = num / 52#
        Case "M"
            PeriodToYears = num / 12#
        Case "Y"
            PeriodToYears = num
        Case Else
            Err.Raise vbObjectError + 102, , "期間を年換算できません: " & in_PeriodText
    End Select
End Function

Private Function NormalizeVol(ByVal in_X As Double) As Double
    If in_X < 0# Then Err.Raise vbObjectError + 103, , "Volがマイナスです: " & in_X

    'Excelの%表示なら 0.5077% はセル値 0.005077。
    '手入力で 0.5077 と入っている場合だけ 100で割る。
    If in_X > 0.2 Then
        NormalizeVol = in_X / 100#
    Else
        NormalizeVol = in_X
    End If
End Function

Private Function GetATMVol(ByRef in_ShiftsBps() As Double, _
                           ByRef in_Vols() As Double) As Double
    Dim i As Long, bestI As Long, bestAbs As Double
    bestAbs = 1E+99

    For i = LBound(in_ShiftsBps) To UBound(in_ShiftsBps)
        If Abs(in_ShiftsBps(i)) < bestAbs Then
            bestAbs = Abs(in_ShiftsBps(i))
            bestI = i
        End If
    Next i

    GetATMVol = in_Vols(bestI)
End Function

Private Sub UnpackParams(ByRef in_Y() As Double, _
                         ByRef out_Alpha As Double, _
                         ByRef out_Rho As Double, _
                         ByRef out_Nu As Double)
    out_Alpha = Exp(in_Y(0))
    out_Rho = TanhSafe(in_Y(1))
    out_Nu = Exp(in_Y(2))

    If out_Rho > 0.999 Then out_Rho = 0.999
    If out_Rho < -0.999 Then out_Rho = -0.999
End Sub

Private Function AtanhSafe(ByVal in_X As Double) As Double
    If in_X >= 0.999 Then in_X = 0.999
    If in_X <= -0.999 Then in_X = -0.999
    AtanhSafe = 0.5 * Log((1# + in_X) / (1# - in_X))
End Function

Private Function TanhSafe(ByVal in_X As Double) As Double
    If in_X > 20# Then
        TanhSafe = 1#
    ElseIf in_X < -20# Then
        TanhSafe = -1#
    Else
        TanhSafe = (Exp(2# * in_X) - 1#) / (Exp(2# * in_X) + 1#)
    End If
End Function

Private Sub SortSimplex(ByRef inout_Val() As Double, _
                        ByRef inout_Simplex() As Double, _
                        ByRef out_Lo As Long, _
                        ByRef out_Hi As Long, _
                        ByRef out_Nhi As Long)
    Dim i As Long
    out_Lo = LBound(inout_Val)
    out_Hi = LBound(inout_Val)

    For i = LBound(inout_Val) To UBound(inout_Val)
        If inout_Val(i) < inout_Val(out_Lo) Then out_Lo = i
        If inout_Val(i) > inout_Val(out_Hi) Then out_Hi = i
    Next i

    out_Nhi = out_Lo
    For i = LBound(inout_Val) To UBound(inout_Val)
        If i <> out_Hi Then
            If out_Nhi = out_Hi Or inout_Val(i) > inout_Val(out_Nhi) Then out_Nhi = i
        End If
    Next i
End Sub

Private Sub ReplacePoint(ByRef inout_Simplex() As Double, _
                         ByRef inout_Val() As Double, _
                         ByVal in_Idx As Long, _
                         ByRef in_Y() As Double, _
                         ByVal in_Obj As Double)
    Dim j As Long
    For j = 0 To 2
        inout_Simplex(in_Idx, j) = in_Y(j)
    Next j
    inout_Val(in_Idx) = in_Obj
End Sub