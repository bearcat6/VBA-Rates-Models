Option Explicit

'============================================================
' Class: clsSplineCurve
'
' Purpose:
'   Natural cubic spline interpolation for numeric x-y data.
'
' Design policy:
'   - Self-contained class. No dependency on standard modules.
'   - No financial date logic is included.
'   - x values must be strictly increasing.
'   - Natural boundary condition is used:
'       second derivative at both endpoints = 0
'
' Public API:
'   Build
'   Value
'   TryValue
'
' Legacy API:
'   genSpline
'   calcSpline
'
' Notes:
'   This class is intended to be used as a small numerical
'   component under discount curves, volatility curves, and
'   other financial model classes.
'============================================================

Private mPointCount As Long

Private mX() As Double
Private mY() As Double

' Coefficients for each interval:
'   S_i(x) = A_i * (x - X_i)^3
'          + B_i * (x - X_i)^2
'          + C_i * (x - X_i)
'          + D_i
Private mA() As Double
Private mB() As Double
Private mC() As Double
Private mD() As Double

Private mIsBuilt As Boolean
Private mAllowFlatExtrapolation As Boolean
Private mLastError As String


'============================================================
' Public API
'============================================================

Public Function Build( _
    ByVal in_DataNumber As Long, _
    ByRef in_X() As Double, _
    ByRef in_Y() As Double _
) As Boolean

  On Error GoTo ErrHandler

  mIsBuilt = False
  mLastError = vbNullString

  If Not ValidateInput(in_DataNumber, in_X, in_Y) Then
    Build = False
    Exit Function
  End If

  CopyInputData in_DataNumber, in_X, in_Y
  BuildCoefficients

  mIsBuilt = True
  Build = True
  Exit Function

ErrHandler:
  mLastError = Err.Description
  Build = False

End Function


Public Function Value(ByVal in_XValue As Double) As Double

  Dim segmentIndex As Long
  Dim dx As Double

  If Not mIsBuilt Then
    Err.Raise vbObjectError + 1001, "clsSplineCurve.Value", _
              "Spline curve has not been built."
  End If

  If in_XValue < mX(1) Then
    If mAllowFlatExtrapolation Then
      Value = mY(1)
      Exit Function
    End If

    Err.Raise vbObjectError + 1002, "clsSplineCurve.Value", _
              "Input x is below the curve range."
  End If

  If in_XValue > mX(mPointCount) Then
    If mAllowFlatExtrapolation Then
      Value = mY(mPointCount)
      Exit Function
    End If

    Err.Raise vbObjectError + 1003, "clsSplineCurve.Value", _
              "Input x is above the curve range."
  End If

  If in_XValue = mX(mPointCount) Then
    Value = mY(mPointCount)
    Exit Function
  End If

  segmentIndex = FindSegmentIndex(in_XValue)
  dx = in_XValue - mX(segmentIndex)

  Value = mA(segmentIndex) * dx ^ 3 _
        + mB(segmentIndex) * dx ^ 2 _
        + mC(segmentIndex) * dx _
        + mD(segmentIndex)

End Function


Public Function TryValue( _
    ByVal in_XValue As Double, _
    ByRef out_YValue As Double _
) As Boolean

  On Error GoTo ErrHandler

  out_YValue = Value(in_XValue)
  TryValue = True
  Exit Function

ErrHandler:
  mLastError = Err.Description
  TryValue = False

End Function


Public Property Get IsBuilt() As Boolean
  IsBuilt = mIsBuilt
End Property


Public Property Get PointCount() As Long
  PointCount = mPointCount
End Property


Public Property Get AllowFlatExtrapolation() As Boolean
  AllowFlatExtrapolation = mAllowFlatExtrapolation
End Property


Public Property Let AllowFlatExtrapolation(ByVal in_Value As Boolean)
  mAllowFlatExtrapolation = in_Value
End Property


Public Property Get LastError() As String
  LastError = mLastError
End Property


Public Function XAt(ByVal in_Index As Long) As Double

  If in_Index < 1 Or in_Index > mPointCount Then
    Err.Raise vbObjectError + 1004, "clsSplineCurve.XAt", _
              "Index is out of range."
  End If

  XAt = mX(in_Index)

End Function


Public Function YAt(ByVal in_Index As Long) As Double

  If in_Index < 1 Or in_Index > mPointCount Then
    Err.Raise vbObjectError + 1005, "clsSplineCurve.YAt", _
              "Index is out of range."
  End If

  YAt = mY(in_Index)

End Function


'============================================================
' Legacy API
'============================================================
' Existing code can continue to call:
'   genSpline dataCount, xArray, yArray
'   calcSpline x
'
' New code should use:
'   Build
'   Value
'============================================================

Public Sub genSpline( _
    ByVal in_DataNumber As Long, _
    ByRef in_XNumber() As Double, _
    ByRef in_YNumber() As Double _
)

  If Not Build(in_DataNumber, in_XNumber, in_YNumber) Then
    Err.Raise vbObjectError + 1010, "clsSplineCurve.genSpline", _
              "Failed to build spline curve. " & mLastError
  End If

End Sub


Public Function calcSpline(ByVal in_Number As Double) As Double
  calcSpline = Value(in_Number)
End Function


'============================================================
' Internal: validation and data setup
'============================================================

Private Function ValidateInput( _
    ByVal in_DataNumber As Long, _
    ByRef in_X() As Double, _
    ByRef in_Y() As Double _
) As Boolean

  On Error GoTo ErrHandler

  Dim i As Long
  Dim lbX As Long
  Dim lbY As Long
  Dim ubX As Long
  Dim ubY As Long

  ValidateInput = False
  mLastError = vbNullString

  If in_DataNumber < 3 Then
    mLastError = "At least 3 data points are required."
    Exit Function
  End If

  lbX = LBound(in_X)
  ubX = UBound(in_X)
  lbY = LBound(in_Y)
  ubY = UBound(in_Y)

  If ubX - lbX + 1 < in_DataNumber Then
    mLastError = "x array size is smaller than in_DataNumber."
    Exit Function
  End If

  If ubY - lbY + 1 < in_DataNumber Then
    mLastError = "y array size is smaller than in_DataNumber."
    Exit Function
  End If

  For i = 2 To in_DataNumber
    If in_X(lbX + i - 1) <= in_X(lbX + i - 2) Then
      mLastError = "x values must be strictly increasing."
      Exit Function
    End If
  Next i

  ValidateInput = True
  Exit Function

ErrHandler:
  mLastError = Err.Description
  ValidateInput = False

End Function


Private Sub CopyInputData( _
    ByVal in_DataNumber As Long, _
    ByRef in_X() As Double, _
    ByRef in_Y() As Double _
)

  Dim i As Long
  Dim lbX As Long
  Dim lbY As Long

  lbX = LBound(in_X)
  lbY = LBound(in_Y)

  mPointCount = in_DataNumber

  ReDim mX(1 To mPointCount)
  ReDim mY(1 To mPointCount)

  For i = 1 To mPointCount
    mX(i) = in_X(lbX + i - 1)
    mY(i) = in_Y(lbY + i - 1)
  Next i

End Sub


'============================================================
' Internal: spline construction
'============================================================

Private Sub BuildCoefficients()

  Dim secondDerivative() As Double
  Dim i As Long
  Dim h As Double

  secondDerivative = BuildSecondDerivatives()

  ReDim mA(1 To mPointCount - 1)
  ReDim mB(1 To mPointCount - 1)
  ReDim mC(1 To mPointCount - 1)
  ReDim mD(1 To mPointCount - 1)

  For i = 1 To mPointCount - 1
    h = mX(i + 1) - mX(i)

    mA(i) = (secondDerivative(i + 1) - secondDerivative(i)) / (6# * h)
    mB(i) = secondDerivative(i) / 2#
    mC(i) = (mY(i + 1) - mY(i)) / h _
          - h * (2# * secondDerivative(i) + secondDerivative(i + 1)) / 6#
    mD(i) = mY(i)
  Next i

End Sub


Private Function BuildSecondDerivatives() As Double()

  Dim lower() As Double
  Dim diag() As Double
  Dim upper() As Double
  Dim rhs() As Double
  Dim solution() As Double
  Dim secondDerivative() As Double

  Dim unknownCount As Long
  Dim i As Long
  Dim k As Long
  Dim hPrev As Double
  Dim hNext As Double

  ReDim secondDerivative(1 To mPointCount)

  ' Natural boundary condition
  secondDerivative(1) = 0#
  secondDerivative(mPointCount) = 0#

  unknownCount = mPointCount - 2

  If unknownCount <= 0 Then
    BuildSecondDerivatives = secondDerivative
    Exit Function
  End If

  ReDim lower(1 To unknownCount)
  ReDim diag(1 To unknownCount)
  ReDim upper(1 To unknownCount)
  ReDim rhs(1 To unknownCount)
  ReDim solution(1 To unknownCount)

  ' Solve for second derivatives M(2) ... M(n-1)
  For i = 2 To mPointCount - 1
    k = i - 1

    hPrev = mX(i) - mX(i - 1)
    hNext = mX(i + 1) - mX(i)

    If k > 1 Then
      lower(k) = hPrev
    Else
      lower(k) = 0#
    End If

    diag(k) = 2# * (hPrev + hNext)

    If k < unknownCount Then
      upper(k) = hNext
    Else
      upper(k) = 0#
    End If

    rhs(k) = 6# * ( _
        (mY(i + 1) - mY(i)) / hNext _
      - (mY(i) - mY(i - 1)) / hPrev _
    )
  Next i

  If Not SolveTridiagonalSystem(lower, diag, upper, rhs, solution) Then
    Err.Raise vbObjectError + 1020, "clsSplineCurve.BuildSecondDerivatives", _
              "Failed to solve tridiagonal system."
  End If

  For i = 2 To mPointCount - 1
    secondDerivative(i) = solution(i - 1)
  Next i

  BuildSecondDerivatives = secondDerivative

End Function


'============================================================
' Internal: evaluation
'============================================================

Private Function FindSegmentIndex(ByVal in_XValue As Double) As Long

  Dim lowIndex As Long
  Dim highIndex As Long
  Dim midIndex As Long

  lowIndex = 1
  highIndex = mPointCount - 1

  Do While lowIndex <= highIndex
    midIndex = (lowIndex + highIndex) \ 2

    If in_XValue < mX(midIndex) Then
      highIndex = midIndex - 1
    ElseIf in_XValue >= mX(midIndex + 1) Then
      lowIndex = midIndex + 1
    Else
      FindSegmentIndex = midIndex
      Exit Function
    End If
  Loop

  Err.Raise vbObjectError + 1030, "clsSplineCurve.FindSegmentIndex", _
            "Failed to find spline segment."

End Function


'============================================================
' Internal: linear algebra
'============================================================
' Thomas algorithm for tridiagonal linear systems.
'
' lower(i) * x(i-1) + diag(i) * x(i) + upper(i) * x(i+1) = rhs(i)
'
' lower(1) and upper(n) are ignored.
'============================================================

Private Function SolveTridiagonalSystem( _
    ByRef in_Lower() As Double, _
    ByRef in_Diag() As Double, _
    ByRef in_Upper() As Double, _
    ByRef in_Rhs() As Double, _
    ByRef out_Solution() As Double _
) As Boolean

  On Error GoTo ErrHandler

  Dim n As Long
  Dim i As Long
  Dim factor As Double

  Dim cPrime() As Double
  Dim dPrime() As Double

  n = UBound(in_Diag)

  ReDim cPrime(1 To n)
  ReDim dPrime(1 To n)

  If Abs(in_Diag(1)) < 0.000000000001 Then
    SolveTridiagonalSystem = False
    Exit Function
  End If

  cPrime(1) = in_Upper(1) / in_Diag(1)
  dPrime(1) = in_Rhs(1) / in_Diag(1)

  For i = 2 To n
    factor = in_Diag(i) - in_Lower(i) * cPrime(i - 1)

    If Abs(factor) < 0.000000000001 Then
      SolveTridiagonalSystem = False
      Exit Function
    End If

    If i < n Then
      cPrime(i) = in_Upper(i) / factor
    Else
      cPrime(i) = 0#
    End If

    dPrime(i) = (in_Rhs(i) - in_Lower(i) * dPrime(i - 1)) / factor
  Next i

  out_Solution(n) = dPrime(n)

  For i = n - 1 To 1 Step -1
    out_Solution(i) = dPrime(i) - cPrime(i) * out_Solution(i + 1)
  Next i

  SolveTridiagonalSystem = True
  Exit Function

ErrHandler:
  SolveTridiagonalSystem = False

End Function
