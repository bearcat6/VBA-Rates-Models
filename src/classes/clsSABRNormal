
Option Explicit

Private Const C_BETA_TOL As Double = 0.0000001
Private Const C_RHO_BOUND As Double = 0.999999
Private Const C_EPS As Double = 0.000000000001
Private Const C_DEFAULT_GRID_WIDTH As Double = 0.05
Private Const C_DEFAULT_DK As Double = 0.0001
Private Const C_DEFAULT_KMIN_SENTINEL As Double = 1E+99
Private Const C_DEFAULT_KMAX_SENTINEL As Double = -1E+99

Private m_cCurve As Object

Private m_OpYears As Double
Private m_TenorYears As Double

Private m_Forward As Double
Private m_Annuity As Double

Private m_Alpha As Double
Private m_Beta As Double
Private m_Rho As Double
Private m_Nu As Double

Private m_HasError As Boolean
Private m_LastError As String

Private m_GridBuilt As Boolean

Private m_K() As Double
Private m_Vol() As Double
Private m_Price() As Double
Private m_Density() As Double
Private m_CDF() As Double

Private m_KMin As Double
Private m_KMax As Double
Private m_DK As Double

'============================================================
' clsSABRNormal
'
' beta = 0 の Normal SABR パラメータを保持し、
' vol / payer swaption price / density / CDF / percentile を返すクラス。
'
' 前提：
' - JPY TONA/OIS ベース
' - Normal Vol は絶対値表記（40bp = 0.0040）
' - NormalSABRVol / SwaptionUnitCallPrice は mdl_SwaptionSABR 側を利用する
' - スワップション取引そのものの商品クラスではない
'============================================================

'============================================================
' 初期化
'
' in_cCurve は以下を返すカーブオブジェクトを想定する。
' - ForwardSwapRate(in_OpYears, in_TenorYears)
' - SwapAnnuity(in_OpYears, in_TenorYears)
'============================================================
Public Sub Init( _
  ByVal in_cCurve As Object, _
  ByVal in_OpYears As Double, _
  ByVal in_TenorYears As Double, _
  ByVal in_Alpha As Double, _
  Optional ByVal in_Beta As Double = 0#, _
  Optional ByVal in_Rho As Double = 0#, _
  Optional ByVal in_Nu As Double = 0# _
)

  On Error GoTo EH

  ClearError

  Set m_cCurve = in_cCurve

  m_OpYears = in_OpYears
  m_TenorYears = in_TenorYears

  m_Forward = in_cCurve.ForwardSwapRate(in_OpYears, in_TenorYears)
  m_Annuity = in_cCurve.SwapAnnuity(in_OpYears, in_TenorYears)

  m_Alpha = in_Alpha
  m_Beta = in_Beta
  m_Rho = in_Rho
  m_Nu = in_Nu

  m_GridBuilt = False
  Erase m_K
  Erase m_Vol
  Erase m_Price
  Erase m_Density
  Erase m_CDF

  If Not IsValid Then
    SetError "SABRパラメータまたはカーブ情報が不正です。"
  End If

  Exit Sub

EH:
  SetError "clsSABRNormal.Initでエラー: " & Err.Description
End Sub

'============================================================
' プロパティ
'============================================================
Public Property Get Alpha() As Double
  Alpha = m_Alpha
End Property

Public Property Get Beta() As Double
  Beta = m_Beta
End Property

Public Property Get Rho() As Double
  Rho = m_Rho
End Property

Public Property Get Nu() As Double
  Nu = m_Nu
End Property

Public Property Get OpYears() As Double
  OpYears = m_OpYears
End Property

Public Property Get TenorYears() As Double
  TenorYears = m_TenorYears
End Property

Public Property Get ForwardRate() As Double
  ForwardRate = m_Forward
End Property

Public Property Get Annuity() As Double
  Annuity = m_Annuity
End Property

Public Property Get GridBuilt() As Boolean
  GridBuilt = m_GridBuilt
End Property

Public Property Get KMin() As Double
  KMin = m_KMin
End Property

Public Property Get KMax() As Double
  KMax = m_KMax
End Property

Public Property Get DK() As Double
  DK = m_DK
End Property

Public Property Get HasError() As Boolean
  HasError = m_HasError
End Property

Public Property Get LastError() As String
  LastError = m_LastError
End Property

Public Sub SetError(ByVal in_Message As String)
  m_HasError = True
  m_LastError = in_Message
End Sub

Public Sub ClearError()
  m_HasError = False
  m_LastError = vbNullString
End Sub

'============================================================
' パラメータ妥当性
'============================================================
Public Property Get IsValid() As Boolean

  If m_OpYears <= 0# Then Exit Property
  If m_TenorYears <= 0# Then Exit Property
  If m_Annuity <= 0# Then Exit Property

  If m_Alpha <= 0# Then Exit Property
  If Abs(m_Beta) > C_BETA_TOL Then Exit Property
  If m_Rho <= -C_RHO_BOUND Or m_Rho >= C_RHO_BOUND Then Exit Property
  If m_Nu < 0# Then Exit Property

  IsValid = True
End Property

'============================================================
' Normal SABR Vol
'============================================================
Public Function Vol(ByVal in_K As Double) As Variant

  Dim volN As Double

  If Not IsValid Then
    Vol = CVErr(xlErrNum)
    Exit Function
  End If

  volN = NormalSABRVol(m_Forward, in_K, m_OpYears, m_Alpha, m_Beta, m_Rho, m_Nu)

  If volN <= 0# Then
    Vol = CVErr(xlErrNum)
  Else
    Vol = volN
  End If
End Function

'============================================================
' Payer swaption 価格
'============================================================
Public Function CallPrice(ByVal in_K As Double) As Variant

  Dim volN As Variant

  If Not IsValid Then
    CallPrice = CVErr(xlErrNum)
    Exit Function
  End If

  volN = Vol(in_K)
  If IsError(volN) Then
    CallPrice = CVErr(xlErrNum)
    Exit Function
  End If

  CallPrice = m_Annuity * SwaptionUnitCallPrice(m_Forward, m_OpYears, in_K, CDbl(volN))
End Function

'============================================================
' アニュイティ除き価格
' 密度計算用
'============================================================
Private Function UnitCallPrice(ByVal in_K As Double) As Double

  Dim volN As Variant
  volN = Vol(in_K)

  If IsError(volN) Then
    UnitCallPrice = 0#
  Else
    UnitCallPrice = SwaptionUnitCallPrice(m_Forward, m_OpYears, in_K, CDbl(volN))
  End If
End Function

'============================================================
' 密度グリッド作成
'
' 例:
'   in_KMin = Forward - 0.05
'   in_KMax = Forward + 0.05
'   in_DK   = 0.0001
'============================================================
Public Function BuildDensityGrid( _
  Optional ByVal in_KMin As Double = C_DEFAULT_KMIN_SENTINEL, _
  Optional ByVal in_KMax As Double = C_DEFAULT_KMAX_SENTINEL, _
  Optional ByVal in_DK As Double = C_DEFAULT_DK _
) As Boolean

  On Error GoTo EH

  ClearError
  m_GridBuilt = False

  If Not IsValid Then
    SetError "不正なSABRパラメータのため密度を作成できません。"
    BuildDensityGrid = False
    Exit Function
  End If

  If in_DK <= 0# Then
    SetError "in_DKが不正です。"
    BuildDensityGrid = False
    Exit Function
  End If

  If in_KMin = C_DEFAULT_KMIN_SENTINEL Then in_KMin = m_Forward - C_DEFAULT_GRID_WIDTH
  If in_KMax = C_DEFAULT_KMAX_SENTINEL Then in_KMax = m_Forward + C_DEFAULT_GRID_WIDTH

  If in_KMin >= in_KMax Then
    SetError "in_KMin >= in_KMax です。"
    BuildDensityGrid = False
    Exit Function
  End If

  m_KMin = in_KMin
  m_KMax = in_KMax
  m_DK = in_DK

  Dim gridCount As Long
  gridCount = CLng((in_KMax - in_KMin) / in_DK) + 1

  If gridCount < 10 Then
    SetError "グリッド数が少なすぎます。"
    BuildDensityGrid = False
    Exit Function
  End If

  ReDim m_K(0 To gridCount - 1)
  ReDim m_Vol(0 To gridCount - 1)
  ReDim m_Price(0 To gridCount - 1)
  ReDim m_Density(0 To gridCount - 1)
  ReDim m_CDF(0 To gridCount - 1)

  Dim i As Long
  Dim volN As Variant

  For i = 0 To gridCount - 1
    m_K(i) = in_KMin + in_DK * i

    volN = Vol(m_K(i))
    If IsError(volN) Then
      SetError "Vol計算エラー。K=" & CStr(m_K(i))
      BuildDensityGrid = False
      Exit Function
    End If

    m_Vol(i) = CDbl(volN)
    m_Price(i) = SwaptionUnitCallPrice(m_Forward, m_OpYears, m_K(i), m_Vol(i))
  Next i

  CalculateDensityFromPriceGrid gridCount, in_DK

  If Not NormalizeDensityGrid(gridCount, in_DK) Then
    BuildDensityGrid = False
    Exit Function
  End If

  m_GridBuilt = True
  BuildDensityGrid = True
  Exit Function

EH:
  SetError "BuildDensityGridでエラー: " & Err.Description
  BuildDensityGrid = False
End Function

'============================================================
' 密度 p(K)
'============================================================
Public Function DensityAt(ByVal in_K As Double) As Variant

  If Not m_GridBuilt Then
    If Not BuildDensityGrid() Then
      DensityAt = CVErr(xlErrValue)
      Exit Function
    End If
  End If

  If in_K < m_KMin Or in_K > m_KMax Then
    DensityAt = CVErr(xlErrNA)
    Exit Function
  End If

  DensityAt = LinearInterpolate(in_K, m_K, m_Density)
End Function

'============================================================
' P(S_T <= K)
' つまり strike K が何％点かを返す
'============================================================
Public Function PercentileOfStrike(ByVal in_K As Double) As Variant

  If Not m_GridBuilt Then
    If Not BuildDensityGrid() Then
      PercentileOfStrike = CVErr(xlErrValue)
      Exit Function
    End If
  End If

  If in_K <= m_KMin Then
    PercentileOfStrike = 0#
    Exit Function
  End If

  If in_K >= m_KMax Then
    PercentileOfStrike = 1#
    Exit Function
  End If

  PercentileOfStrike = LinearInterpolate(in_K, m_K, m_CDF)
End Function

'============================================================
' p％点の strike
' in_Probability = 0.95 なら95%点
'============================================================
Public Function StrikeAtPercentile(ByVal in_Probability As Double) As Variant

  If in_Probability < 0# Or in_Probability > 1# Then
    StrikeAtPercentile = CVErr(xlErrNum)
    Exit Function
  End If

  If Not m_GridBuilt Then
    If Not BuildDensityGrid() Then
      StrikeAtPercentile = CVErr(xlErrValue)
      Exit Function
    End If
  End If

  StrikeAtPercentile = InverseLinearInterpolate(in_Probability, m_CDF, m_K)
End Function

'============================================================
' 密度の健全性チェック
'============================================================
Public Function DensityCheck() As String

  If Not m_GridBuilt Then
    If Not BuildDensityGrid() Then
      DensityCheck = "NG: " & m_LastError
      Exit Function
    End If
  End If

  Dim i As Long
  Dim minDensity As Double
  Dim meanValue As Double
  Dim totalProbability As Double

  minDensity = m_Density(LBound(m_Density))

  For i = LBound(m_Density) To UBound(m_Density)
    If m_Density(i) < minDensity Then minDensity = m_Density(i)
  Next i

  totalProbability = m_CDF(UBound(m_CDF))

  For i = 1 To UBound(m_K)
    meanValue = meanValue + 0.5 * _
      (m_K(i - 1) * m_Density(i - 1) + m_K(i) * m_Density(i)) * m_DK
  Next i

  DensityCheck = _
    "TotalProb=" & Format(totalProbability, "0.000000") & _
    ", Mean=" & Format(meanValue, "0.000000") & _
    ", Forward=" & Format(m_Forward, "0.000000") & _
    ", Mean-Forward=" & Format(meanValue - m_Forward, "0.000000") & _
    ", MinDensity=" & Format(minDensity, "0.000000")
End Function

'============================================================
' グリッドを Excel に出力しやすい2次元配列で返す
'
' columns:
' 1 K
' 2 Vol
' 3 UnitCallPrice
' 4 Density
' 5 CDF
'============================================================
Public Function DensityGridTable() As Variant

  If Not m_GridBuilt Then
    If Not BuildDensityGrid() Then
      DensityGridTable = CVErr(xlErrValue)
      Exit Function
    End If
  End If

  Dim rowCount As Long
  rowCount = UBound(m_K) - LBound(m_K) + 1

  Dim result() As Variant
  ReDim result(1 To rowCount, 1 To 5)

  Dim i As Long
  Dim rowIndex As Long

  rowIndex = 1
  For i = LBound(m_K) To UBound(m_K)
    result(rowIndex, 1) = m_K(i)
    result(rowIndex, 2) = m_Vol(i)
    result(rowIndex, 3) = m_Price(i)
    result(rowIndex, 4) = m_Density(i)
    result(rowIndex, 5) = m_CDF(i)
    rowIndex = rowIndex + 1
  Next i

  DensityGridTable = result
End Function

'============================================================
' call価格グリッドから密度を2階差分で計算
'============================================================
Private Sub CalculateDensityFromPriceGrid( _
  ByVal in_GridCount As Long, _
  ByVal in_DK As Double _
)

  Dim i As Long

  m_Density(0) = 0#
  m_Density(in_GridCount - 1) = 0#

  For i = 1 To in_GridCount - 2
    m_Density(i) = (m_Price(i + 1) - 2# * m_Price(i) + m_Price(i - 1)) / (in_DK * in_DK)

    If m_Density(i) < 0# And m_Density(i) > -0.00000001 Then
      m_Density(i) = 0#
    End If
  Next i
End Sub

'============================================================
' 密度とCDFを正規化
'============================================================
Private Function NormalizeDensityGrid( _
  ByVal in_GridCount As Long, _
  ByVal in_DK As Double _
) As Boolean

  Dim i As Long

  m_CDF(0) = 0#

  For i = 1 To in_GridCount - 1
    m_CDF(i) = m_CDF(i - 1) + 0.5 * (m_Density(i - 1) + m_Density(i)) * in_DK
  Next i

  Dim totalProbability As Double
  totalProbability = m_CDF(in_GridCount - 1)

  If totalProbability <= 0# Then
    SetError "密度の積分が0以下です。"
    NormalizeDensityGrid = False
    Exit Function
  End If

  For i = 0 To in_GridCount - 1
    m_Density(i) = m_Density(i) / totalProbability
    m_CDF(i) = m_CDF(i) / totalProbability
  Next i

  NormalizeDensityGrid = True
End Function

'============================================================
' 線形補間
'============================================================
Private Function LinearInterpolate( _
  ByVal in_X As Double, _
  ByRef in_Xs() As Double, _
  ByRef in_Ys() As Double _
) As Double

  Dim i As Long

  If in_X <= in_Xs(LBound(in_Xs)) Then
    LinearInterpolate = in_Ys(LBound(in_Ys))
    Exit Function
  End If

  If in_X >= in_Xs(UBound(in_Xs)) Then
    LinearInterpolate = in_Ys(UBound(in_Ys))
    Exit Function
  End If

  For i = LBound(in_Xs) To UBound(in_Xs) - 1
    If in_Xs(i) <= in_X And in_X <= in_Xs(i + 1) Then
      LinearInterpolate = in_Ys(i) + _
        (in_Ys(i + 1) - in_Ys(i)) * (in_X - in_Xs(i)) / (in_Xs(i + 1) - in_Xs(i))
      Exit Function
    End If
  Next i
End Function

'============================================================
' 逆線形補間
' in_Xs が CDF、in_Ys が K
'============================================================
Private Function InverseLinearInterpolate( _
  ByVal in_X As Double, _
  ByRef in_Xs() As Double, _
  ByRef in_Ys() As Double _
) As Double

  Dim i As Long

  If in_X <= in_Xs(LBound(in_Xs)) Then
    InverseLinearInterpolate = in_Ys(LBound(in_Ys))
    Exit Function
  End If

  If in_X >= in_Xs(UBound(in_Xs)) Then
    InverseLinearInterpolate = in_Ys(UBound(in_Ys))
    Exit Function
  End If

  For i = LBound(in_Xs) To UBound(in_Xs) - 1
    If in_Xs(i) <= in_X And in_X <= in_Xs(i + 1) Then

      If Abs(in_Xs(i + 1) - in_Xs(i)) < C_EPS Then
        InverseLinearInterpolate = in_Ys(i)
      Else
        InverseLinearInterpolate = in_Ys(i) + _
          (in_Ys(i + 1) - in_Ys(i)) * (in_X - in_Xs(i)) / (in_Xs(i + 1) - in_Xs(i))
      End If

      Exit Function
    End If
  Next i
End Function
