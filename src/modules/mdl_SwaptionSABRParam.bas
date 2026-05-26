
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

Public Function BuildSwaptionSABRFromSheet(ByVal ws As Worksheet, _
                                           ByVal cCurve As clsOISStepForwardCurve, _
                                           Optional ByVal outTopLeft As Range = Nothing, _
                                           Optional ByVal firstDataRow As Long = 5, _
                                           Optional ByVal lastDataRow As Long = 0, _
                                           Optional ByVal labelCol As Long = 2, _
                                           Optional ByVal firstVolCol As Long = 3, _
                                           Optional ByVal lastVolCol As Long = 11, _
                                           Optional ByVal headerRow As Long = 4) As Collection
    Dim results As New Collection
    Dim r As Long, c As Long, n As Long
    Dim labelText As String, expiry As String, tenor As String
    Dim fwd As Double, tExp As Double
    Dim shifts() As Double, vols() As Double
    Dim p As clsSABRParams

    If lastDataRow = 0 Then
        lastDataRow = ws.Cells(ws.Rows.Count, labelCol).End(xlUp).Row
    End If

    n = lastVolCol - firstVolCol + 1
    ReDim shifts(1 To n)
    ReDim vols(1 To n)

    For c = firstVolCol To lastVolCol
        shifts(c - firstVolCol + 1) = HeaderToShiftBps(CStr(ws.Cells(headerRow, c).Value))
    Next c

    If Not outTopLeft Is Nothing Then
        WriteSABRHeader outTopLeft
    End If

    For r = firstDataRow To lastDataRow
        labelText = Trim$(CStr(ws.Cells(r, labelCol).Value))
        If Len(labelText) > 0 Then
            ParseExpiryTenor labelText, expiry, tenor
            tExp = PeriodToYears(expiry)

            'ここがポイント：
            'clsOISStepForwardCurve側の ForwardParRate(Expiry, Tenor) を利用する
            fwd = cCurve.ForwardParRate(expiry, tenor)

            For c = firstVolCol To lastVolCol
                vols(c - firstVolCol + 1) = NormalizeVol(CDbl(ws.Cells(r, c).Value))
            Next c

            Set p = CalibrateSABR_Beta0(expiry, tenor, fwd, tExp, shifts, vols)
            results.Add p, expiry & " x " & tenor

            If Not outTopLeft Is Nothing Then
                WriteSABRRow outTopLeft.Offset(r - firstDataRow + 1, 0), p
            End If
        End If
    Next r

    Set BuildSwaptionSABRFromSheet = results
End Function

Public Function CalibrateSABR_Beta0(ByVal expiry As String, _
                                    ByVal tenor As String, _
                                    ByVal fwd As Double, _
                                    ByVal tExp As Double, _
                                    ByRef shiftsBps() As Double, _
                                    ByRef marketVols() As Double) As clsSABRParams
    Dim atmVol As Double
    atmVol = GetATMVol(shiftsBps, marketVols)

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

                yOpt = NelderMead3_Beta0(fwd, tExp, shiftsBps, marketVols, y0, obj, iters)

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
    p.Expiry = expiry
    p.Tenor = tenor
    p.Forward = fwd
    p.ExpiryYears = tExp
    p.Alpha = alpha
    p.Beta = 0#
    p.Rho = rho
    p.Nu = nu
    p.RMSE = CalcRMSE_Beta0(fwd, tExp, shiftsBps, marketVols, alpha, rho, nu)
    p.MaxAbsError = CalcMaxAbsError_Beta0(fwd, tExp, shiftsBps, marketVols, alpha, rho, nu)
    p.Iterations = bestIters
    p.Success = (p.RMSE < 0.0005) '5bp程度。必要なら判定基準は調整

    Set CalibrateSABR_Beta0 = p
End Function

Public Function SABR_NormalVol_Beta0(ByVal fwd As Double, _
                                     ByVal strike As Double, _
                                     ByVal tExp As Double, _
                                     ByVal alpha As Double, _
                                     ByVal rho As Double, _
                                     ByVal nu As Double) As Double
    Const EPS As Double = 0.0000000001

    If alpha <= 0# Or nu < 0# Or Abs(rho) >= 1# Then
        SABR_NormalVol_Beta0 = CVErr(xlErrNum)
        Exit Function
    End If

    Dim z As Double, xz As Double, ratio As Double, tmp As Double
    z = (nu / alpha) * (fwd - strike)

    If Abs(z) < EPS Then
        ratio = 1#
    Else
        tmp = Sqr(1# - 2# * rho * z + z * z) + z - rho
        If tmp <= 0# Or (1# - rho) <= 0# Then
            SABR_NormalVol_Beta0 = CVErr(xlErrNum)
            Exit Function
        End If
        xz = Log(tmp / (1# - rho))

        If Abs(xz) < EPS Then
            ratio = 1#
        Else
            ratio = z / xz
        End If
    End If

    SABR_NormalVol_Beta0 = alpha * ratio * (1# + ((2# - 3# * rho * rho) * nu * nu * tExp / 24#))
End Function

Private Function NelderMead3_Beta0(ByVal fwd As Double, _
                                   ByVal tExp As Double, _
                                   ByRef shiftsBps() As Double, _
                                   ByRef marketVols() As Double, _
                                   ByRef y0() As Double, _
                                   ByRef bestObj As Double, _
                                   ByRef iterations As Long) As Variant
    Const N As Long = 3
    Const MAX_ITER As Long = 300
    Const TOL As Double = 0.000000000001

    Dim simplex(0 To N, 0 To N - 1) As Double
    Dim val(0 To N) As Double
    Dim i As Long, j As Long

    For j = 0 To N - 1
        simplex(0, j) = y0(j)
    Next j

    For i = 1 To N
        For j = 0 To N - 1
            simplex(i, j) = y0(j)
        Next j
        simplex(i, i - 1) = simplex(i, i - 1) + 0.15
    Next i

    For i = 0 To N
        val(i) = Objective_Beta0(fwd, tExp, shiftsBps, marketVols, simplex, i)
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
        fr = ObjectiveVector_Beta0(fwd, tExp, shiftsBps, marketVols, xr)

        If fr < val(lo) Then
            ' expansion
            For j = 0 To N - 1
                xe(j) = centroid(j) + 2# * (xr(j) - centroid(j))
            Next j
            fe = ObjectiveVector_Beta0(fwd, tExp, shiftsBps, marketVols, xe)

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

            fc = ObjectiveVector_Beta0(fwd, tExp, shiftsBps, marketVols, xc)

            If fc < val(hi) Then
                ReplacePoint simplex, val, hi, xc, fc
            Else
                ' shrink
                For i = 0 To N
                    If i <> lo Then
                        For j = 0 To N - 1
                            simplex(i, j) = simplex(lo, j) + 0.5 * (simplex(i, j) - simplex(lo, j))
                        Next j
                        val(i) = Objective_Beta0(fwd, tExp, shiftsBps, marketVols, simplex, i)
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

    bestObj = val(lo)
    iterations = iter
    NelderMead3_Beta0 = ans
End Function

Private Function Objective_Beta0(ByVal fwd As Double, _
                                 ByVal tExp As Double, _
                                 ByRef shiftsBps() As Double, _
                                 ByRef marketVols() As Double, _
                                 ByRef simplex() As Double, _
                                 ByVal idx As Long) As Double
    Dim y(0 To 2) As Double
    y(0) = simplex(idx, 0)
    y(1) = simplex(idx, 1)
    y(2) = simplex(idx, 2)
    Objective_Beta0 = ObjectiveVector_Beta0(fwd, tExp, shiftsBps, marketVols, y)
End Function

Private Function ObjectiveVector_Beta0(ByVal fwd As Double, _
                                       ByVal tExp As Double, _
                                       ByRef shiftsBps() As Double, _
                                       ByRef marketVols() As Double, _
                                       ByRef y() As Double) As Double
    Dim alpha As Double, rho As Double, nu As Double
    UnpackParams y, alpha, rho, nu

    If alpha <= 0# Or nu < 0# Or Abs(rho) >= 0.999999 Then
        ObjectiveVector_Beta0 = 1E+99
        Exit Function
    End If

    Dim i As Long, k As Double, mdl As Double, e As Double, sse As Double
    For i = LBound(shiftsBps) To UBound(shiftsBps)
        k = fwd + shiftsBps(i) / 10000#
        mdl = SABR_NormalVol_Beta0(fwd, k, tExp, alpha, rho, nu)
        e = mdl - marketVols(i)
        sse = sse + e * e
    Next i

    ObjectiveVector_Beta0 = sse
End Function

Private Function CalcRMSE_Beta0(ByVal fwd As Double, _
                                ByVal tExp As Double, _
                                ByRef shiftsBps() As Double, _
                                ByRef marketVols() As Double, _
                                ByVal alpha As Double, _
                                ByVal rho As Double, _
                                ByVal nu As Double) As Double
    Dim i As Long, n As Long, k As Double, mdl As Double, e As Double, sse As Double

    For i = LBound(shiftsBps) To UBound(shiftsBps)
        k = fwd + shiftsBps(i) / 10000#
        mdl = SABR_NormalVol_Beta0(fwd, k, tExp, alpha, rho, nu)
        e = mdl - marketVols(i)
        sse = sse + e * e
        n = n + 1
    Next i

    CalcRMSE_Beta0 = Sqr(sse / n)
End Function

Private Function CalcMaxAbsError_Beta0(ByVal fwd As Double, _
                                       ByVal tExp As Double, _
                                       ByRef shiftsBps() As Double, _
                                       ByRef marketVols() As Double, _
                                       ByVal alpha As Double, _
                                       ByVal rho As Double, _
                                       ByVal nu As Double) As Double
    Dim i As Long, k As Double, mdl As Double, e As Double, mx As Double

    For i = LBound(shiftsBps) To UBound(shiftsBps)
        k = fwd + shiftsBps(i) / 10000#
        mdl = SABR_NormalVol_Beta0(fwd, k, tExp, alpha, rho, nu)
        e = Abs(mdl - marketVols(i))
        If e > mx Then mx = e
    Next i

    CalcMaxAbsError_Beta0 = mx
End Function

Private Sub WriteSABRHeader(ByVal topLeft As Range)
    topLeft.Resize(1, 12).Value = Array( _
        "Expiry", "Tenor", "Forward", "Alpha", "Beta", "Rho", "Nu", _
        "RMSE", "RMSE(bp)", "MaxAbsErr", "MaxAbsErr(bp)", "Success")
End Sub

Private Sub WriteSABRRow(ByVal topLeft As Range, ByVal p As clsSABRParams)
    topLeft.Offset(0, 0).Value = p.Expiry
    topLeft.Offset(0, 1).Value = p.Tenor
    topLeft.Offset(0, 2).Value = p.Forward
    topLeft.Offset(0, 3).Value = p.Alpha
    topLeft.Offset(0, 4).Value = p.Beta
    topLeft.Offset(0, 5).Value = p.Rho
    topLeft.Offset(0, 6).Value = p.Nu
    topLeft.Offset(0, 7).Value = p.RMSE
    topLeft.Offset(0, 8).Value = p.RMSE * 10000#
    topLeft.Offset(0, 9).Value = p.MaxAbsError
    topLeft.Offset(0, 10).Value = p.MaxAbsError * 10000#
    topLeft.Offset(0, 11).Value = p.Success
End Sub

Private Sub ParseExpiryTenor(ByVal labelText As String, _
                             ByRef expiry As String, _
                             ByRef tenor As String)
    Dim s As String, a As Variant
    s = UCase$(Replace(labelText, " ", ""))
    a = Split(s, "X")
    If UBound(a) <> 1 Then Err.Raise vbObjectError + 100, , "Expiry x Tenor の形式ではありません: " & labelText

    expiry = CStr(a(0))
    tenor = CStr(a(1))
End Sub

Private Function HeaderToShiftBps(ByVal headerText As String) As Double
    Dim s As String
    s = UCase$(Trim$(headerText))

    If s = "ATM" Then
        HeaderToShiftBps = 0#
        Exit Function
    End If

    s = Replace(s, "BPS", "")
    s = Replace(s, "BP", "")
    s = Replace(s, "+", "")
    s = Trim$(s)

    If Len(s) = 0 Or Not IsNumeric(s) Then
        Err.Raise vbObjectError + 101, , "bpsヘッダーを読めません: " & headerText
    End If

    HeaderToShiftBps = CDbl(s)
End Function

Private Function PeriodToYears(ByVal periodText As String) As Double
    Dim s As String, unit As String, num As Double
    s = UCase$(Trim$(periodText))

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
            Err.Raise vbObjectError + 102, , "期間を年換算できません: " & periodText
    End Select
End Function

Private Function NormalizeVol(ByVal x As Double) As Double
    If x < 0# Then Err.Raise vbObjectError + 103, , "Volがマイナスです: " & x

    'Excelの%表示なら 0.5077% はセル値 0.005077。
    '手入力で 0.5077 と入っている場合だけ 100で割る。
    If x > 0.2 Then
        NormalizeVol = x / 100#
    Else
        NormalizeVol = x
    End If
End Function

Private Function GetATMVol(ByRef shiftsBps() As Double, _
                           ByRef vols() As Double) As Double
    Dim i As Long, bestI As Long, bestAbs As Double
    bestAbs = 1E+99

    For i = LBound(shiftsBps) To UBound(shiftsBps)
        If Abs(shiftsBps(i)) < bestAbs Then
            bestAbs = Abs(shiftsBps(i))
            bestI = i
        End If
    Next i

    GetATMVol = vols(bestI)
End Function

Private Sub UnpackParams(ByRef y() As Double, _
                         ByRef alpha As Double, _
                         ByRef rho As Double, _
                         ByRef nu As Double)
    alpha = Exp(y(0))
    rho = TanhSafe(y(1))
    nu = Exp(y(2))

    If rho > 0.999 Then rho = 0.999
    If rho < -0.999 Then rho = -0.999
End Sub

Private Function AtanhSafe(ByVal x As Double) As Double
    If x >= 0.999 Then x = 0.999
    If x <= -0.999 Then x = -0.999
    AtanhSafe = 0.5 * Log((1# + x) / (1# - x))
End Function

Private Function TanhSafe(ByVal x As Double) As Double
    If x > 20# Then
        TanhSafe = 1#
    ElseIf x < -20# Then
        TanhSafe = -1#
    Else
        TanhSafe = (Exp(2# * x) - 1#) / (Exp(2# * x) + 1#)
    End If
End Function

Private Sub SortSimplex(ByRef val() As Double, _
                        ByRef simplex() As Double, _
                        ByRef lo As Long, _
                        ByRef hi As Long, _
                        ByRef nhi As Long)
    Dim i As Long
    lo = LBound(val)
    hi = LBound(val)

    For i = LBound(val) To UBound(val)
        If val(i) < val(lo) Then lo = i
        If val(i) > val(hi) Then hi = i
    Next i

    nhi = lo
    For i = LBound(val) To UBound(val)
        If i <> hi Then
            If nhi = hi Or val(i) > val(nhi) Then nhi = i
        End If
    Next i
End Sub

Private Sub ReplacePoint(ByRef simplex() As Double, _
                         ByRef val() As Double, _
                         ByVal idx As Long, _
                         ByRef y() As Double, _
                         ByVal obj As Double)
    Dim j As Long
    For j = 0 To 2
        simplex(idx, j) = y(j)
    Next j
    val(idx) = obj
End Sub
