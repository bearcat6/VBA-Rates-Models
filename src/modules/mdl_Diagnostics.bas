Option Explicit

' =============================================================================
' mdl_Diagnostics
' =============================================================================
'
' Diagnostic utilities for model development.
'
' Main targets:
'   - Hull-White curve time-based interface
'       DF_T(T)
'       ZeroRate_T(T)
'       ForwardRate_T(T1, T2)
'       InstantaneousForward(T, eps)
'       DateFromT(T)
'       YearFracFromValDate(Date)
'
'   - Volatility surface style objects
'       VolByYears(expiryYears, tenorYears)
'       NormalVol(expiryYears, tenorYears)
'
' Design policy:
'   - Model classes should not read Excel directly.
'   - This diagnostics module may write diagnostic tables to Excel sheets.
'   - Curve and vol surface objects are accepted as Object to support
'     convention-based interfaces in VBA.
'   - Public procedure arguments use the in_ prefix.
'
' =============================================================================

' =============================================================================
' Hull-White Curve Interface Check
' =============================================================================
Public Function CheckHWCurveInterface(ByVal in_Curve As Object) As Boolean

    On Error GoTo ErrHandler

    If in_Curve Is Nothing Then
        Err.Raise vbObjectError + 9001, "mdl_Diagnostics.CheckHWCurveInterface", _
                  "Curve object is Nothing."
    End If

    ' Minimal smoke test for convention-based HW curve interface.
    Call in_Curve.DateFromT(1#)
    Call in_Curve.DF_T(1#)
    Call in_Curve.ZeroRate_T(1#)
    Call in_Curve.ForwardRate_T(1#, 2#)
    Call in_Curve.InstantaneousForward(1#, 0.0001)

    CheckHWCurveInterface = True
    Exit Function

ErrHandler:
    CheckHWCurveInterface = False

End Function


Public Sub AssertHWCurveInterface(ByVal in_Curve As Object)

    On Error GoTo ErrHandler

    If in_Curve Is Nothing Then
        Err.Raise vbObjectError + 9011, "mdl_Diagnostics.AssertHWCurveInterface", _
                  "Curve object is Nothing."
    End If

    Call in_Curve.DateFromT(1#)
    Call in_Curve.DF_T(1#)
    Call in_Curve.ZeroRate_T(1#)
    Call in_Curve.ForwardRate_T(1#, 2#)
    Call in_Curve.InstantaneousForward(1#, 0.0001)

    Exit Sub

ErrHandler:
    Err.Raise vbObjectError + 9012, "mdl_Diagnostics.AssertHWCurveInterface", _
              "Curve object does not satisfy Hull-White time-based interface. " & _
              "Original error: " & Err.Description

End Sub


' =============================================================================
' Hull-White Curve Interface Diagnostic Table
' =============================================================================
Public Sub DiagnoseHWCurveInterface( _
    ByVal in_Curve As Object, _
    ByVal in_OutputSheetName As String, _
    Optional ByVal in_MaxT As Double = 30#, _
    Optional ByVal in_StepT As Double = 0.5, _
    Optional ByVal in_Eps As Double = 0.0001 _
)

    Dim n As Long
    Dim i As Long
    Dim t As Double
    Dim t2 As Double
    Dim arr() As Variant

    If in_Curve Is Nothing Then
        Err.Raise vbObjectError + 9021, "mdl_Diagnostics.DiagnoseHWCurveInterface", _
                  "Curve object is Nothing."
    End If

    If in_MaxT < 0# Then
        Err.Raise vbObjectError + 9022, "mdl_Diagnostics.DiagnoseHWCurveInterface", _
                  "in_MaxT must be non-negative."
    End If

    If in_StepT <= 0# Then
        Err.Raise vbObjectError + 9023, "mdl_Diagnostics.DiagnoseHWCurveInterface", _
                  "in_StepT must be positive."
    End If

    If in_Eps <= 0# Then
        Err.Raise vbObjectError + 9024, "mdl_Diagnostics.DiagnoseHWCurveInterface", _
                  "in_Eps must be positive."
    End If

    n = CLng(Fix(in_MaxT / in_StepT)) + 1
    ReDim arr(1 To n + 1, 1 To 8)

    arr(1, 1) = "T"
    arr(1, 2) = "DateFromT"
    arr(1, 3) = "DF_T"
    arr(1, 4) = "ZeroRate_T"
    arr(1, 5) = "InstantaneousForward"
    arr(1, 6) = "ForwardRate_T(T,T+Step)"
    arr(1, 7) = "NextT"
    arr(1, 8) = "Status"

    For i = 1 To n
        t = CDbl(i - 1) * in_StepT
        t2 = t + in_StepT

        arr(i + 1, 1) = t
        arr(i + 1, 7) = t2

        arr(i + 1, 2) = TryDateFromT(in_Curve, t)
        arr(i + 1, 3) = TryDF_T(in_Curve, t)
        arr(i + 1, 4) = TryZeroRate_T(in_Curve, t)
        arr(i + 1, 5) = TryInstantaneousForward(in_Curve, t, in_Eps)

        If t < in_MaxT Then
            arr(i + 1, 6) = TryForwardRate_T(in_Curve, t, t2)
        Else
            arr(i + 1, 6) = ""
        End If

        arr(i + 1, 8) = RowStatus(arr(i + 1, 2), arr(i + 1, 3), arr(i + 1, 4), arr(i + 1, 5))
    Next i

    WriteArrayToSheet arr, in_OutputSheetName, "A1"

End Sub


' =============================================================================
' Date-based vs Time-based Curve Consistency
' =============================================================================
Public Sub DiagnoseCurveDateTimeConsistency( _
    ByVal in_Curve As Object, _
    ByVal in_TargetDates As Variant, _
    ByVal in_OutputSheetName As String _
)

    Dim n As Long
    Dim i As Long
    Dim d As Date
    Dim t As Double
    Dim dfDate As Variant
    Dim dfTime As Variant
    Dim arr() As Variant

    If in_Curve Is Nothing Then
        Err.Raise vbObjectError + 9031, "mdl_Diagnostics.DiagnoseCurveDateTimeConsistency", _
                  "Curve object is Nothing."
    End If

    n = GetVectorItemCount(in_TargetDates)
    If n <= 0 Then
        Err.Raise vbObjectError + 9032, "mdl_Diagnostics.DiagnoseCurveDateTimeConsistency", _
                  "in_TargetDates is empty."
    End If

    ReDim arr(1 To n + 1, 1 To 7)

    arr(1, 1) = "TargetDate"
    arr(1, 2) = "T = YearFracFromValDate"
    arr(1, 3) = "DF(Date)"
    arr(1, 4) = "DF_T(T)"
    arr(1, 5) = "Difference"
    arr(1, 6) = "Absolute Difference"
    arr(1, 7) = "Status"

    For i = 1 To n
        d = CDate(GetVectorItem(in_TargetDates, i))
        arr(i + 1, 1) = d

        t = TryYearFracFromValDate(in_Curve, d)
        arr(i + 1, 2) = t

        dfDate = TryDF_Date(in_Curve, d)
        dfTime = TryDF_T(in_Curve, t)

        arr(i + 1, 3) = dfDate
        arr(i + 1, 4) = dfTime

        If IsNumeric(dfDate) And IsNumeric(dfTime) Then
            arr(i + 1, 5) = CDbl(dfTime) - CDbl(dfDate)
            arr(i + 1, 6) = Abs(CDbl(dfTime) - CDbl(dfDate))
            arr(i + 1, 7) = "OK"
        Else
            arr(i + 1, 5) = ""
            arr(i + 1, 6) = ""
            arr(i + 1, 7) = "ERROR"
        End If
    Next i

    WriteArrayToSheet arr, in_OutputSheetName, "A1"

End Sub


' =============================================================================
' Instantaneous Forward eps Sensitivity
' =============================================================================
Public Sub DiagnoseInstantaneousForwardEps( _
    ByVal in_Curve As Object, _
    ByVal in_Times As Variant, _
    ByVal in_EpsValues As Variant, _
    ByVal in_OutputSheetName As String _
)

    Dim nT As Long
    Dim nEps As Long
    Dim i As Long
    Dim j As Long
    Dim t As Double
    Dim eps As Double
    Dim arr() As Variant

    If in_Curve Is Nothing Then
        Err.Raise vbObjectError + 9041, "mdl_Diagnostics.DiagnoseInstantaneousForwardEps", _
                  "Curve object is Nothing."
    End If

    nT = GetVectorItemCount(in_Times)
    nEps = GetVectorItemCount(in_EpsValues)

    If nT <= 0 Or nEps <= 0 Then
        Err.Raise vbObjectError + 9042, "mdl_Diagnostics.DiagnoseInstantaneousForwardEps", _
                  "in_Times and in_EpsValues must not be empty."
    End If

    ReDim arr(1 To nT + 1, 1 To nEps + 1)

    arr(1, 1) = "T"
    For j = 1 To nEps
        eps = CDbl(GetVectorItem(in_EpsValues, j))
        arr(1, j + 1) = "eps=" & CStr(eps)
    Next j

    For i = 1 To nT
        t = CDbl(GetVectorItem(in_Times, i))
        arr(i + 1, 1) = t

        For j = 1 To nEps
            eps = CDbl(GetVectorItem(in_EpsValues, j))
            arr(i + 1, j + 1) = TryInstantaneousForward(in_Curve, t, eps)
        Next j
    Next i

    WriteArrayToSheet arr, in_OutputSheetName, "A1"

End Sub


' =============================================================================
' Vol Surface Interface Check
' =============================================================================
Public Function CheckVolSurfaceInterface(ByVal in_VolSurface As Object) As Boolean

    On Error GoTo ErrHandler

    If in_VolSurface Is Nothing Then
        Err.Raise vbObjectError + 9051, "mdl_Diagnostics.CheckVolSurfaceInterface", _
                  "Vol surface object is Nothing."
    End If

    Call in_VolSurface.VolByYears(1#, 5#)

    CheckVolSurfaceInterface = True
    Exit Function

ErrHandler:
    CheckVolSurfaceInterface = False

End Function


Public Sub AssertVolSurfaceInterface(ByVal in_VolSurface As Object)

    On Error GoTo ErrHandler

    If in_VolSurface Is Nothing Then
        Err.Raise vbObjectError + 9061, "mdl_Diagnostics.AssertVolSurfaceInterface", _
                  "Vol surface object is Nothing."
    End If

    Call in_VolSurface.VolByYears(1#, 5#)

    Exit Sub

ErrHandler:
    Err.Raise vbObjectError + 9062, "mdl_Diagnostics.AssertVolSurfaceInterface", _
              "Vol surface object does not satisfy expected interface. " & _
              "Original error: " & Err.Description

End Sub


' =============================================================================
' Vol Surface Diagnostic Table
' =============================================================================
Public Sub DiagnoseVolSurface( _
    ByVal in_VolSurface As Object, _
    ByVal in_ExpiryYears As Variant, _
    ByVal in_TenorYears As Variant, _
    ByVal in_OutputSheetName As String, _
    Optional ByVal in_UseNormalVol As Boolean = True _
)

    Dim nExp As Long
    Dim nTenor As Long
    Dim i As Long
    Dim j As Long
    Dim expiryYears As Double
    Dim tenorYears As Double
    Dim arr() As Variant

    If in_VolSurface Is Nothing Then
        Err.Raise vbObjectError + 9071, "mdl_Diagnostics.DiagnoseVolSurface", _
                  "Vol surface object is Nothing."
    End If

    nExp = GetVectorItemCount(in_ExpiryYears)
    nTenor = GetVectorItemCount(in_TenorYears)

    If nExp <= 0 Or nTenor <= 0 Then
        Err.Raise vbObjectError + 9072, "mdl_Diagnostics.DiagnoseVolSurface", _
                  "in_ExpiryYears and in_TenorYears must not be empty."
    End If

    ReDim arr(1 To nExp + 1, 1 To nTenor + 1)

    arr(1, 1) = "Expiry \ Tenor"
    For j = 1 To nTenor
        arr(1, j + 1) = CDbl(GetVectorItem(in_TenorYears, j))
    Next j

    For i = 1 To nExp
        expiryYears = CDbl(GetVectorItem(in_ExpiryYears, i))
        arr(i + 1, 1) = expiryYears

        For j = 1 To nTenor
            tenorYears = CDbl(GetVectorItem(in_TenorYears, j))
            If in_UseNormalVol Then
                arr(i + 1, j + 1) = TryNormalVol(in_VolSurface, expiryYears, tenorYears)
            Else
                arr(i + 1, j + 1) = TryVolByYears(in_VolSurface, expiryYears, tenorYears)
            End If
        Next j
    Next i

    WriteArrayToSheet arr, in_OutputSheetName, "A1"

End Sub


' =============================================================================
' Sheet Output Helper
' =============================================================================
Public Sub WriteArrayToSheet( _
    ByVal in_Data As Variant, _
    ByVal in_OutputSheetName As String, _
    Optional ByVal in_StartCellAddress As String = "A1" _
)

    Dim ws As Worksheet
    Dim rCount As Long
    Dim cCount As Long

    Set ws = PrepareOutputSheet(in_OutputSheetName)

    rCount = UBound(in_Data, 1) - LBound(in_Data, 1) + 1
    cCount = UBound(in_Data, 2) - LBound(in_Data, 2) + 1

    ws.Range(in_StartCellAddress).Resize(rCount, cCount).Value = in_Data
    ws.Columns.AutoFit

End Sub


Public Function PrepareOutputSheet(ByVal in_OutputSheetName As String) As Worksheet

    Dim ws As Worksheet

    If Trim$(in_OutputSheetName) = "" Then
        Err.Raise vbObjectError + 9081, "mdl_Diagnostics.PrepareOutputSheet", _
                  "in_OutputSheetName must not be empty."
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(in_OutputSheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = in_OutputSheetName
    Else
        ws.Cells.Clear
    End If

    Set PrepareOutputSheet = ws

End Function


' =============================================================================
' Private Try Helpers: Curve
' =============================================================================
Private Function TryDateFromT(ByVal in_Curve As Object, ByVal in_T As Double) As Variant

    On Error GoTo ErrHandler
    TryDateFromT = in_Curve.DateFromT(in_T)
    Exit Function

ErrHandler:
    TryDateFromT = "ERROR: " & Err.Description

End Function


Private Function TryYearFracFromValDate(ByVal in_Curve As Object, ByVal in_TargetDate As Date) As Variant

    On Error GoTo ErrHandler
    TryYearFracFromValDate = in_Curve.YearFracFromValDate(in_TargetDate)
    Exit Function

ErrHandler:
    TryYearFracFromValDate = "ERROR: " & Err.Description

End Function


Private Function TryDF_T(ByVal in_Curve As Object, ByVal in_T As Double) As Variant

    On Error GoTo ErrHandler
    TryDF_T = in_Curve.DF_T(in_T)
    Exit Function

ErrHandler:
    TryDF_T = "ERROR: " & Err.Description

End Function


Private Function TryZeroRate_T(ByVal in_Curve As Object, ByVal in_T As Double) As Variant

    On Error GoTo ErrHandler
    TryZeroRate_T = in_Curve.ZeroRate_T(in_T)
    Exit Function

ErrHandler:
    TryZeroRate_T = "ERROR: " & Err.Description

End Function


Private Function TryForwardRate_T(ByVal in_Curve As Object, ByVal in_T1 As Double, ByVal in_T2 As Double) As Variant

    On Error GoTo ErrHandler
    TryForwardRate_T = in_Curve.ForwardRate_T(in_T1, in_T2)
    Exit Function

ErrHandler:
    TryForwardRate_T = "ERROR: " & Err.Description

End Function


Private Function TryInstantaneousForward( _
    ByVal in_Curve As Object, _
    ByVal in_T As Double, _
    ByVal in_Eps As Double _
) As Variant

    On Error GoTo ErrHandler
    TryInstantaneousForward = in_Curve.InstantaneousForward(in_T, in_Eps)
    Exit Function

ErrHandler:
    TryInstantaneousForward = "ERROR: " & Err.Description

End Function


Private Function TryDF_Date(ByVal in_Curve As Object, ByVal in_TargetDate As Date) As Variant

    On Error GoTo ErrHandler
    TryDF_Date = in_Curve.DF(in_TargetDate)
    Exit Function

ErrHandler:
    TryDF_Date = "ERROR: " & Err.Description

End Function


' =============================================================================
' Private Try Helpers: Vol Surface
' =============================================================================
Private Function TryVolByYears( _
    ByVal in_VolSurface As Object, _
    ByVal in_ExpiryYears As Double, _
    ByVal in_TenorYears As Double _
) As Variant

    On Error GoTo ErrHandler
    TryVolByYears = in_VolSurface.VolByYears(in_ExpiryYears, in_TenorYears)
    Exit Function

ErrHandler:
    TryVolByYears = "ERROR: " & Err.Description

End Function


Private Function TryNormalVol( _
    ByVal in_VolSurface As Object, _
    ByVal in_ExpiryYears As Double, _
    ByVal in_TenorYears As Double _
) As Variant

    On Error GoTo TryVolByYearsFallback
    TryNormalVol = in_VolSurface.NormalVol(in_ExpiryYears, in_TenorYears)
    Exit Function

TryVolByYearsFallback:
    Err.Clear
    On Error GoTo ErrHandler
    TryNormalVol = in_VolSurface.VolByYears(in_ExpiryYears, in_TenorYears)
    Exit Function

ErrHandler:
    TryNormalVol = "ERROR: " & Err.Description

End Function


' =============================================================================
' Private General Helpers
' =============================================================================
Private Function RowStatus( _
    ByVal in_Value1 As Variant, _
    ByVal in_Value2 As Variant, _
    ByVal in_Value3 As Variant, _
    ByVal in_Value4 As Variant _
) As String

    If IsErrorText(in_Value1) Or IsErrorText(in_Value2) Or _
       IsErrorText(in_Value3) Or IsErrorText(in_Value4) Then
        RowStatus = "ERROR"
    Else
        RowStatus = "OK"
    End If

End Function


Private Function IsErrorText(ByVal in_Value As Variant) As Boolean

    If VarType(in_Value) = vbString Then
        IsErrorText = (Left$(CStr(in_Value), 6) = "ERROR:")
    Else
        IsErrorText = False
    End If

End Function


Private Function GetVectorItemCount(ByVal in_Value As Variant) As Long

    On Error GoTo SingleValue

    If TypeName(in_Value) = "Range" Then
        GetVectorItemCount = in_Value.Cells.Count
    ElseIf IsArray(in_Value) Then
        GetVectorItemCount = UBound(in_Value) - LBound(in_Value) + 1
    Else
        GetVectorItemCount = 1
    End If

    Exit Function

SingleValue:
    GetVectorItemCount = 1

End Function


Private Function GetVectorItem(ByVal in_Value As Variant, ByVal in_Index As Long) As Variant

    On Error GoTo TryArray2D

    If TypeName(in_Value) = "Range" Then
        GetVectorItem = in_Value.Cells(in_Index).Value
    ElseIf IsArray(in_Value) Then
        GetVectorItem = in_Value(LBound(in_Value) + in_Index - 1)
    Else
        If in_Index <> 1 Then
            Err.Raise vbObjectError + 9091, "mdl_Diagnostics.GetVectorItem", _
                      "Scalar value has only one item."
        End If
        GetVectorItem = in_Value
    End If

    Exit Function

TryArray2D:
    On Error GoTo SingleValue

    If IsArray(in_Value) Then
        GetVectorItem = in_Value(LBound(in_Value, 1) + in_Index - 1, LBound(in_Value, 2))
        Exit Function
    End If

SingleValue:
    If in_Index <> 1 Then
        Err.Raise vbObjectError + 9092, "mdl_Diagnostics.GetVectorItem", _
                  "Could not get vector item."
    End If
    GetVectorItem = in_Value

End Function
