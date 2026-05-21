
Option Explicit

Private Const C_BIG_PENALTY As Double = 1E+99
Private Const C_RHO_BOUND As Double = 0.999
Private Const C_ALPHA_TOL As Double = 0.00000001
Private Const C_RHO_TOL As Double = 0.000001
Private Const C_NU_TOL As Double = 0.000001

'============================================================
' mdl_FitNormalSABR
'
' JPY TONA/OIS ベースの swaption normal volatility smile を
' beta = 0 の Normal SABR にフィットする標準モジュール。
'
' 前提：
' - market vol は Normal Vol の絶対値表記（40bp = 0.0040）
' - beta = 0 固定
' - NormalSABRVol は mdl_SwaptionSABR 側の関数を利用する
' - 返却クラスは assumptions.md に合わせて clsSABRNormal とする
'============================================================

'============================================================
' Normal SABR を market normal vol quotes にフィットして
' clsSABRNormal クラスを返す
'
' in_Strikes(): strike K_i
' in_Vols():    market normal vol_i
'
' in_cCurve は以下を返すカーブオブジェクトを想定する。
' - ForwardSwapRate(in_OpYears, in_TenorYears)
'============================================================
Public Function FitNormalSABR( _
  ByVal in_cCurve As Object, _
  ByVal in_OpYears As Double, _
  ByVal in_TenorYears As Double, _
  ByRef in_Strikes() As Double, _
  ByRef in_Vols() As Double, _
  Optional ByVal in_MaxIterations As Long = 300, _
  Optional ByVal in_InitialRho As Double = 0#, _
  Optional ByVal in_InitialNu As Double = 0.5 _
) As clsSABRNormal

  On Error GoTo EH

  Dim cSABR As clsSABRNormal
  Set cSABR = New clsSABRNormal

  If Not IsValidFitInput(in_OpYears, in_TenorYears, in_Strikes, in_Vols) Then
    cSABR.SetError "SABRフィットの入力が不正です。expiry、tenor、strike、volを確認してください。"
    Set FitNormalSABR = cSABR
    Exit Function
  End If

  Dim forwardRate As Double
  forwardRate = in_cCurve.ForwardSwapRate(in_OpYears, in_TenorYears)

  Dim atmVol As Double
  atmVol = GuessATMVol(forwardRate, in_Strikes, in_Vols)

  If atmVol <= 0# Then
    cSABR.SetError "ATM Volを推定できません。入力ボラを確認してください。"
    Set FitNormalSABR = cSABR
    Exit Function
  End If

  Dim alpha As Double
  Dim rho As Double
  Dim nu As Double

  alpha = atmVol
  rho = ClampRho(in_InitialRho)
  nu = Max2(in_InitialNu, 0#)

  Dim stepAlpha As Double
  Dim stepRho As Double
  Dim stepNu As Double

  stepAlpha = Max2(atmVol * 0.2, 0.000001)
  stepRho = 0.2
  stepNu = 0.2

  Dim bestObjective As Double
  bestObjective = SABRObjective( _
    forwardRate, in_OpYears, in_Strikes, in_Vols, alpha, rho, nu)

  Dim iter As Long
  Dim improved As Boolean

  For iter = 1 To in_MaxIterations

    improved = False

    improved = TryCandidate( _
      forwardRate, in_OpYears, in_Strikes, in_Vols, _
      alpha + stepAlpha, rho, nu, _
      alpha, rho, nu, bestObjective) Or improved

    If alpha - stepAlpha > 0# Then
      improved = TryCandidate( _
        forwardRate, in_OpYears, in_Strikes, in_Vols, _
        alpha - stepAlpha, rho, nu, _
        alpha, rho, nu, bestObjective) Or improved
    End If

    If rho + stepRho < C_RHO_BOUND Then
      improved = TryCandidate( _
        forwardRate, in_OpYears, in_Strikes, in_Vols, _
        alpha, rho + stepRho, nu, _
        alpha, rho, nu, bestObjective) Or improved
    End If

    If rho - stepRho > -C_RHO_BOUND Then
      improved = TryCandidate( _
        forwardRate, in_OpYears, in_Strikes, in_Vols, _
        alpha, rho - stepRho, nu, _
        alpha, rho, nu, bestObjective) Or improved
    End If

    improved = TryCandidate( _
      forwardRate, in_OpYears, in_Strikes, in_Vols, _
      alpha, rho, nu + stepNu, _
      alpha, rho, nu, bestObjective) Or improved

    If nu - stepNu >= 0# Then
      improved = TryCandidate( _
        forwardRate, in_OpYears, in_Strikes, in_Vols, _
        alpha, rho, nu - stepNu, _
        alpha, rho, nu, bestObjective) Or improved
    End If

    If Not improved Then
      stepAlpha = stepAlpha * 0.5
      stepRho = stepRho * 0.5
      stepNu = stepNu * 0.5
    End If

    If stepAlpha < C_ALPHA_TOL And stepRho < C_RHO_TOL And stepNu < C_NU_TOL Then
      Exit For
    End If

  Next iter

  cSABR.Init in_cCurve, in_OpYears, in_TenorYears, alpha, 0#, rho, nu

  If Not cSABR.IsValid Then
    cSABR.SetError "SABRパラメータが不正です。"
  End If

  Set FitNormalSABR = cSABR
  Exit Function

EH:
  If cSABR Is Nothing Then Set cSABR = New clsSABRNormal
  cSABR.SetError "FitNormalSABRでエラー: " & Err.Description
  Set FitNormalSABR = cSABR
End Function

'============================================================
' 目的関数
'
' market normal vol と model normal vol の加重二乗誤差。
' ATM近辺は CMS convexity 等への影響が大きいため少し重くする。
'============================================================
Private Function SABRObjective( _
  ByVal in_F As Double, _
  ByVal in_T As Double, _
  ByRef in_Strikes() As Double, _
  ByRef in_Vols() As Double, _
  ByVal in_Alpha As Double, _
  ByVal in_Rho As Double, _
  ByVal in_Nu As Double _
) As Double

  On Error GoTo EH

  If in_Alpha <= 0# Then
    SABRObjective = C_BIG_PENALTY
    Exit Function
  End If

  If in_Nu < 0# Then
    SABRObjective = C_BIG_PENALTY
    Exit Function
  End If

  If in_Rho <= -C_RHO_BOUND Or in_Rho >= C_RHO_BOUND Then
    SABRObjective = C_BIG_PENALTY
    Exit Function
  End If

  Dim i As Long
  Dim modelVol As Double
  Dim diff As Double
  Dim obj As Double
  Dim weight As Double

  obj = 0#

  For i = LBound(in_Strikes) To UBound(in_Strikes)

    If in_Vols(i) <= 0# Then
      SABRObjective = C_BIG_PENALTY
      Exit Function
    End If

    modelVol = NormalSABRVol(in_F, in_Strikes(i), in_T, in_Alpha, 0#, in_Rho, in_Nu)

    If modelVol <= 0# Then
      SABRObjective = C_BIG_PENALTY
      Exit Function
    End If

    weight = SmileFitWeight(in_F, in_Strikes(i))
    diff = modelVol - in_Vols(i)
    obj = obj + weight * diff * diff

  Next i

  SABRObjective = obj
  Exit Function

EH:
  SABRObjective = C_BIG_PENALTY
End Function

'============================================================
' 候補パラメータを評価し、改善していれば現在値を更新する
'============================================================
Private Function TryCandidate( _
  ByVal in_F As Double, _
  ByVal in_T As Double, _
  ByRef in_Strikes() As Double, _
  ByRef in_Vols() As Double, _
  ByVal in_CandidateAlpha As Double, _
  ByVal in_CandidateRho As Double, _
  ByVal in_CandidateNu As Double, _
  ByRef io_Alpha As Double, _
  ByRef io_Rho As Double, _
  ByRef io_Nu As Double, _
  ByRef io_BestObjective As Double _
) As Boolean

  Dim obj As Double

  obj = SABRObjective( _
    in_F, in_T, in_Strikes, in_Vols, _
    in_CandidateAlpha, in_CandidateRho, in_CandidateNu)

  If obj < io_BestObjective Then
    io_Alpha = in_CandidateAlpha
    io_Rho = in_CandidateRho
    io_Nu = in_CandidateNu
    io_BestObjective = obj
    TryCandidate = True
  End If
End Function

'============================================================
' ATMに最も近いquoteのvolを初期alphaに使う
'============================================================
Private Function GuessATMVol( _
  ByVal in_F As Double, _
  ByRef in_Strikes() As Double, _
  ByRef in_Vols() As Double _
) As Double

  On Error GoTo EH

  Dim i As Long
  Dim bestIndex As Long
  Dim bestDistance As Double
  Dim distance As Double

  bestIndex = LBound(in_Strikes)
  bestDistance = Abs(in_Strikes(bestIndex) - in_F)

  For i = LBound(in_Strikes) To UBound(in_Strikes)
    distance = Abs(in_Strikes(i) - in_F)
    If distance < bestDistance Then
      bestDistance = distance
      bestIndex = i
    End If
  Next i

  GuessATMVol = in_Vols(bestIndex)
  Exit Function

EH:
  GuessATMVol = 0#
End Function

'============================================================
' smile fit の重み
'============================================================
Private Function SmileFitWeight( _
  ByVal in_F As Double, _
  ByVal in_K As Double _
) As Double

  Dim atmDistance As Double
  atmDistance = Abs(in_K - in_F)

  If atmDistance < 0.000001 Then
    SmileFitWeight = 3#
  ElseIf atmDistance <= 0.005 Then
    SmileFitWeight = 2#
  ElseIf atmDistance <= 0.01 Then
    SmileFitWeight = 1.5
  Else
    SmileFitWeight = 1#
  End If
End Function

'============================================================
' 入力検証
'============================================================
Private Function IsValidFitInput( _
  ByVal in_OpYears As Double, _
  ByVal in_TenorYears As Double, _
  ByRef in_Strikes() As Double, _
  ByRef in_Vols() As Double _
) As Boolean

  On Error GoTo EH

  Dim lowerStrike As Long
  Dim upperStrike As Long
  Dim lowerVol As Long
  Dim upperVol As Long
  Dim i As Long

  If in_OpYears <= 0# Then Exit Function
  If in_TenorYears <= 0# Then Exit Function

  lowerStrike = LBound(in_Strikes)
  upperStrike = UBound(in_Strikes)
  lowerVol = LBound(in_Vols)
  upperVol = UBound(in_Vols)

  If lowerStrike <> lowerVol Then Exit Function
  If upperStrike <> upperVol Then Exit Function
  If upperStrike < lowerStrike Then Exit Function

  For i = lowerStrike To upperStrike
    If in_Vols(i) <= 0# Then Exit Function
  Next i

  IsValidFitInput = True
  Exit Function

EH:
  IsValidFitInput = False
End Function

Private Function ClampRho(ByVal in_Rho As Double) As Double
  If in_Rho >= C_RHO_BOUND Then
    ClampRho = C_RHO_BOUND - 0.000001
  ElseIf in_Rho <= -C_RHO_BOUND Then
    ClampRho = -C_RHO_BOUND + 0.000001
  Else
    ClampRho = in_Rho
  End If
End Function

Private Function Max2(ByVal in_A As Double, ByVal in_B As Double) As Double
  If in_A >= in_B Then
    Max2 = in_A
  Else
    Max2 = in_B
  End If
End Function
