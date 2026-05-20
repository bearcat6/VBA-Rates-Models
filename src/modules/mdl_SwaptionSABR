Option Explicit

Private Const C_PI As Double = 3.14159265358979
Private Const C_EPS As Double = 0.000000000001
Private Const C_BETA_TOL As Double = 0.0000001

'============================================================
' mdl_SwaptionSABR
'
' JPY TONA/OIS ベースの swaption normal SABR 補助関数。
' - beta = 0 の Normal SABR を対象とする
' - Normal Vol は絶対値表記とする（40bp = 0.0040）
' - スワップションの商品クラスは作成しない前提
' - CMS convexity adjustment / CMS spread swap 等で参照する
'   smile・密度・percentile の取得を目的とする
'============================================================

'============================================================
' 標準正規 PDF
'============================================================
Public Function NormPDF(ByVal in_X As Double) As Double
  NormPDF = Exp(-0.5 * in_X * in_X) / Sqr(2# * C_PI)
End Function

'============================================================
' 標準正規 CDF
' Excel 関数を利用
'============================================================
Public Function NormCDF(ByVal in_X As Double) As Double
  NormCDF = WorksheetFunction.Norm_S_Dist(in_X, True)
End Function

'============================================================
' SABR パラメータ検証
'============================================================
Public Function IsValidNormalSABRParameters( _
  ByVal in_T As Double, _
  ByVal in_Alpha As Double, _
  Optional ByVal in_Beta As Double = 0#, _
  Optional ByVal in_Rho As Double = 0#, _
  Optional ByVal in_Nu As Double = 0# _
) As Boolean

  If in_T < 0# Then Exit Function
  If in_Alpha <= 0# Then Exit Function
  If Abs(in_Beta) > C_BETA_TOL Then Exit Function
  If in_Rho <= -0.999999 Or in_Rho >= 0.999999 Then Exit Function
  If in_Nu < 0# Then Exit Function

  IsValidNormalSABRParameters = True
End Function

'============================================================
' Normal SABR Vol
'
' Hagan 型の beta = 0 normal SABR 近似。
' beta <> 0 は本モジュールの対象外として 0 を返す。
'============================================================
Public Function NormalSABRVol( _
  ByVal in_F As Double, _
  ByVal in_K As Double, _
  ByVal in_T As Double, _
  ByVal in_Alpha As Double, _
  Optional ByVal in_Beta As Double = 0#, _
  Optional ByVal in_Rho As Double = 0#, _
  Optional ByVal in_Nu As Double = 0# _
) As Double

  On Error GoTo EH

  Dim z As Double
  Dim xz As Double
  Dim zOverXz As Double
  Dim adjustment As Double
  Dim tmp As Double

  If Not IsValidNormalSABRParameters(in_T, in_Alpha, in_Beta, in_Rho, in_Nu) Then GoTo EH

  If in_Nu = 0# Then
    NormalSABRVol = in_Alpha
    Exit Function
  End If

  z = (in_Nu / in_Alpha) * (in_F - in_K)

  If Abs(z) < 0.00000001 Then
    zOverXz = 1#
  Else
    tmp = Sqr(1# - 2# * in_Rho * z + z * z) + z - in_Rho
    If tmp <= 0# Or (1# - in_Rho) <= 0# Then GoTo EH

    xz = Log(tmp / (1# - in_Rho))

    If Abs(xz) < C_EPS Then
      zOverXz = 1#
    Else
      zOverXz = z / xz
    End If
  End If

  adjustment = 1# + ((2# - 3# * in_Rho * in_Rho) / 24#) * in_Nu * in_Nu * in_T
  NormalSABRVol = in_Alpha * zOverXz * adjustment

  If NormalSABRVol <= 0# Then GoTo EH
  Exit Function

EH:
  NormalSABRVol = 0#
End Function

'============================================================
' アニュイティ除き Bachelier payer swaption 価格
'============================================================
Public Function SwaptionUnitCallPrice( _
  ByVal in_F As Double, _
  ByVal in_T As Double, _
  ByVal in_K As Double, _
  ByVal in_VolN As Double _
) As Double

  Dim d As Double
  Dim sqrtT As Double

  If in_T <= 0# Or in_VolN <= 0# Then
    SwaptionUnitCallPrice = 0#
    Exit Function
  End If

  sqrtT = Sqr(in_T)
  d = (in_F - in_K) / (in_VolN * sqrtT)

  SwaptionUnitCallPrice = (in_F - in_K) * NormCDF(d) + in_VolN * sqrtT * NormPDF(d)
End Function

'============================================================
' アニュイティ除き SABR smile 付き payer swaption 価格
'============================================================
Public Function SwaptionSABRUnitCallPrice( _
  ByVal in_F As Double, _
  ByVal in_T As Double, _
  ByVal in_K As Double, _
  ByVal in_Alpha As Double, _
  Optional ByVal in_Beta As Double = 0#, _
  Optional ByVal in_Rho As Double = 0#, _
  Optional ByVal in_Nu As Double = 0# _
) As Double

  Dim volN As Double

  volN = NormalSABRVol(in_F, in_K, in_T, in_Alpha, in_Beta, in_Rho, in_Nu)
  If volN <= 0# Then Exit Function

  SwaptionSABRUnitCallPrice = SwaptionUnitCallPrice(in_F, in_T, in_K, volN)
End Function

'============================================================
' Bachelier payer swaption 価格
'
' in_cCurve は以下を返すカーブオブジェクトを想定する。
' - ForwardSwapRate(in_OpYears, in_TenorYears)
' - SwapAnnuity(in_OpYears, in_TenorYears)
'============================================================
Public Function SwaptionCallPrice( _
  ByVal in_cCurve As Object, _
  ByVal in_OpYears As Double, _
  ByVal in_TenorYears As Double, _
  ByVal in_K As Double, _
  ByVal in_Alpha As Double, _
  Optional ByVal in_Beta As Double = 0#, _
  Optional ByVal in_Rho As Double = 0#, _
  Optional ByVal in_Nu As Double = 0# _
) As Variant

  On Error GoTo EH

  Dim forwardRate As Double
  Dim annuity As Double
  Dim unitPrice As Double

  If in_OpYears <= 0# Or in_TenorYears <= 0# Then
    SwaptionCallPrice = CVErr(xlErrNum)
    Exit Function
  End If

  forwardRate = in_cCurve.ForwardSwapRate(in_OpYears, in_TenorYears)
  annuity = in_cCurve.SwapAnnuity(in_OpYears, in_TenorYears)

  If annuity <= 0# Then
    SwaptionCallPrice = CVErr(xlErrNum)
    Exit Function
  End If

  unitPrice = SwaptionSABRUnitCallPrice( _
    forwardRate, in_OpYears, in_K, in_Alpha, in_Beta, in_Rho, in_Nu)

  If unitPrice <= 0# Then
    SwaptionCallPrice = CVErr(xlErrNum)
    Exit Function
  End If

  SwaptionCallPrice = annuity * unitPrice
  Exit Function

EH:
  SwaptionCallPrice = CVErr(xlErrValue)
End Function

'============================================================
' SABR smile からリスク中立 CDF を数値微分で近似
'
' Call(K) = E[(F - K)+] より、dCall/dK = -P(F > K)
' したがって CDF(K) = 1 + dCall/dK
'============================================================
Public Function NormalSABRCDF( _
  ByVal in_F As Double, _
  ByVal in_T As Double, _
  ByVal in_K As Double, _
  ByVal in_Alpha As Double, _
  Optional ByVal in_Beta As Double = 0#, _
  Optional ByVal in_Rho As Double = 0#, _
  Optional ByVal in_Nu As Double = 0#, _
  Optional ByVal in_StrikeStep As Double = 0.00001 _
) As Variant

  On Error GoTo EH

  Dim h As Double
  Dim cPlus As Double
  Dim cMinus As Double
  Dim cdfValue As Double

  h = EffectiveStrikeStep(in_StrikeStep)

  cPlus = SwaptionSABRUnitCallPrice(in_F, in_T, in_K + h, in_Alpha, in_Beta, in_Rho, in_Nu)
  cMinus = SwaptionSABRUnitCallPrice(in_F, in_T, in_K - h, in_Alpha, in_Beta, in_Rho, in_Nu)

  cdfValue = 1# + (cPlus - cMinus) / (2# * h)
  NormalSABRCDF = Clamp01(cdfValue)
  Exit Function

EH:
  NormalSABRCDF = CVErr(xlErrValue)
End Function

'============================================================
' SABR smile からリスク中立密度を数値微分で近似
'
' density(K) = d2Call/dK2
'============================================================
Public Function NormalSABRDensity( _
  ByVal in_F As Double, _
  ByVal in_T As Double, _
  ByVal in_K As Double, _
  ByVal in_Alpha As Double, _
  Optional ByVal in_Beta As Double = 0#, _
  Optional ByVal in_Rho As Double = 0#, _
  Optional ByVal in_Nu As Double = 0#, _
  Optional ByVal in_StrikeStep As Double = 0.00001 _
) As Variant

  On Error GoTo EH

  Dim h As Double
  Dim cPlus As Double
  Dim c0 As Double
  Dim cMinus As Double
  Dim densityValue As Double

  h = EffectiveStrikeStep(in_StrikeStep)

  cPlus = SwaptionSABRUnitCallPrice(in_F, in_T, in_K + h, in_Alpha, in_Beta, in_Rho, in_Nu)
  c0 = SwaptionSABRUnitCallPrice(in_F, in_T, in_K, in_Alpha, in_Beta, in_Rho, in_Nu)
  cMinus = SwaptionSABRUnitCallPrice(in_F, in_T, in_K - h, in_Alpha, in_Beta, in_Rho, in_Nu)

  densityValue = (cPlus - 2# * c0 + cMinus) / (h * h)
  If densityValue < 0# Then densityValue = 0#

  NormalSABRDensity = densityValue
  Exit Function

EH:
  NormalSABRDensity = CVErr(xlErrValue)
End Function

'============================================================
' 指定 percentile に対応する strike を返す
'
' in_Probability: 0 < p < 1
'============================================================
Public Function NormalSABRPercentileStrike( _
  ByVal in_F As Double, _
  ByVal in_T As Double, _
  ByVal in_Probability As Double, _
  ByVal in_Alpha As Double, _
  Optional ByVal in_Beta As Double = 0#, _
  Optional ByVal in_Rho As Double = 0#, _
  Optional ByVal in_Nu As Double = 0#, _
  Optional ByVal in_MaxIterations As Long = 100, _
  Optional ByVal in_Tolerance As Double = 0.00000001 _
) As Variant

  On Error GoTo EH

  Dim lowK As Double
  Dim highK As Double
  Dim midK As Double
  Dim width As Double
  Dim i As Long
  Dim cdfMid As Variant

  If in_Probability <= 0# Or in_Probability >= 1# Then
    NormalSABRPercentileStrike = CVErr(xlErrNum)
    Exit Function
  End If

  If Not IsValidNormalSABRParameters(in_T, in_Alpha, in_Beta, in_Rho, in_Nu) Then
    NormalSABRPercentileStrike = CVErr(xlErrNum)
    Exit Function
  End If

  width = Max2(10# * in_Alpha * Sqr(Max2(in_T, 0.000001)), 0.01)
  lowK = in_F - width
  highK = in_F + width

  For i = 1 To 20
    If CDbl(NormalSABRCDF(in_F, in_T, lowK, in_Alpha, in_Beta, in_Rho, in_Nu)) <= in_Probability _
      And CDbl(NormalSABRCDF(in_F, in_T, highK, in_Alpha, in_Beta, in_Rho, in_Nu)) >= in_Probability Then
      Exit For
    End If
    width = width * 2#
    lowK = in_F - width
    highK = in_F + width
  Next i

  For i = 1 To in_MaxIterations
    midK = 0.5 * (lowK + highK)
    cdfMid = NormalSABRCDF(in_F, in_T, midK, in_Alpha, in_Beta, in_Rho, in_Nu)

    If IsError(cdfMid) Then
      NormalSABRPercentileStrike = CVErr(xlErrValue)
      Exit Function
    End If

    If Abs(CDbl(cdfMid) - in_Probability) < in_Tolerance Then
      NormalSABRPercentileStrike = midK
      Exit Function
    End If

    If CDbl(cdfMid) < in_Probability Then
      lowK = midK
    Else
      highK = midK
    End If
  Next i

  NormalSABRPercentileStrike = 0.5 * (lowK + highK)
  Exit Function

EH:
  NormalSABRPercentileStrike = CVErr(xlErrValue)
End Function

Private Function EffectiveStrikeStep(ByVal in_StrikeStep As Double) As Double
  If in_StrikeStep <= 0# Then
    EffectiveStrikeStep = 0.00001
  Else
    EffectiveStrikeStep = in_StrikeStep
  End If
End Function

Private Function Clamp01(ByVal in_Value As Double) As Double
  If in_Value < 0# Then
    Clamp01 = 0#
  ElseIf in_Value > 1# Then
    Clamp01 = 1#
  Else
    Clamp01 = in_Value
  End If
End Function

Private Function Max2(ByVal in_A As Double, ByVal in_B As Double) As Double
  If in_A >= in_B Then
    Max2 = in_A
  Else
    Max2 = in_B
  End If
End Function
