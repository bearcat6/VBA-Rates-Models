Option Explicit

'============================================================
' clsRandomNormal
'
' Standard Normal Random Number Generator
'
' 目的：
'   ・Monte Carlo simulation で利用する標準正規乱数を生成する。
'   ・乱数生成を Hull-White simulator 等のモデルクラスから分離する。
'
' 実装：
'   ・Box-Muller 法を使用する。
'   ・1回のBox-Mullerで2つの独立な標準正規乱数を生成できるため、
'     2個目を cache して次回呼び出し時に返す。
'
' 主な用途：
'   ・clsHWSimulator
'   ・将来の Monte Carlo pricer
'
' 注意：
'   ・VBA標準の Rnd / Randomize を乱数源とする。
'   ・金融MCの厳密な乱数品質検証までは行っていない。
'   ・初期実装では、再現性確保とクラス分離を主目的とする。
'============================================================

Private mIsInitialized As Boolean
Private mHasSeed As Boolean
Private mSeed As Long

Private mHasCachedNormal As Boolean
Private mCachedNormal As Double

Private mUseAntithetic As Boolean
Private mHasAntithetic As Boolean
Private mAntitheticValue As Double

Private Const PI_VALUE As Double = 3.14159265358979
Private Const MIN_UNIFORM As Double = 0.000000000001

'============================================================
' 初期化
'
' in_Seed:
'   0以下なら Randomize のみ
'   正なら再現性のあるseedとして使う
'
' in_UseAntithetic:
'   True の場合、Normal() は z, -z, z, -z, ... の順に返す。
'============================================================
Public Sub Init( _
  Optional ByVal in_Seed As Long = 0, _
  Optional ByVal in_UseAntithetic As Boolean = False _
)

  mHasCachedNormal = False
  mCachedNormal = 0#

  mUseAntithetic = in_UseAntithetic
  mHasAntithetic = False
  mAntitheticValue = 0#

  If in_Seed > 0 Then
    mHasSeed = True
    mSeed = in_Seed

    Rnd -CDbl(in_Seed)
    Randomize in_Seed
  Else
    mHasSeed = False
    mSeed = 0

    Randomize
  End If

  mIsInitialized = True

End Sub

'============================================================
' 初期化済み判定
'============================================================
Public Property Get IsInitialized() As Boolean
  IsInitialized = mIsInitialized
End Property

'============================================================
' Seed指定ありか
'============================================================
Public Property Get HasSeed() As Boolean
  HasSeed = mHasSeed
End Property

'============================================================
' Seed値
'============================================================
Public Property Get Seed() As Long
  If Not mHasSeed Then
    Err.Raise vbObjectError + 9001, _
      "clsRandomNormal.Seed", _
      "Seed is not set."
  End If

  Seed = mSeed
End Property

'============================================================
' Antithetic を使用するか
'============================================================
Public Property Get UseAntithetic() As Boolean
  UseAntithetic = mUseAntithetic
End Property

'============================================================
' Antithetic 使用有無を変更する
'
' 変更時には内部cacheをクリアする。
'============================================================
Public Property Let UseAntithetic(ByVal in_UseAntithetic As Boolean)
  mUseAntithetic = in_UseAntithetic
  ClearCache
End Property

'============================================================
' 標準正規乱数を1つ返す
'
' 戻り値：
'   N(0,1)
'============================================================
Public Function Normal() As Double

  Dim z As Double

  CheckInitialized

  If mUseAntithetic Then

    If mHasAntithetic Then
      Normal = mAntitheticValue
      mHasAntithetic = False
      Exit Function
    End If

    z = NextIndependentNormal()
    mAntitheticValue = -z
    mHasAntithetic = True

    Normal = z
    Exit Function

  End If

  Normal = NextIndependentNormal()

End Function

'============================================================
' 独立な標準正規乱数を1つ返す
'
' Antithetic 設定は無視する。
'============================================================
Public Function IndependentNormal() As Double

  CheckInitialized

  IndependentNormal = NextIndependentNormal()

End Function

'============================================================
' 標準正規乱数ペアを返す
'
'
' out_Z1, out_Z2:
'   Box-Mullerで生成した独立な標準正規乱数
'
'============================================================
Public Sub NormalPair( _
  ByRef out_Z1 As Double, _
  ByRef out_Z2 As Double _
)

  Dim u1 As Double
  Dim u2 As Double
  Dim radius As Double
  Dim angle As Double

  CheckInitialized

  u1 = SafeUniform()
  u2 = SafeUniform()

  radius = Sqr(-2# * Log(u1))
  angle = 2# * PI_VALUE * u2

  out_Z1 = radius * Cos(angle)
  out_Z2 = radius * Sin(angle)

End Sub

'============================================================
' Uniform(0,1) を返す
'
' Rnd が0を返す可能性に備えて下限を置く。
'============================================================
Public Function Uniform01() As Double

  CheckInitialized

  Uniform01 = SafeUniform()

End Function

'============================================================
' cacheをクリアする
'============================================================
Public Sub ClearCache()

  mHasCachedNormal = False
  mCachedNormal = 0#

  mHasAntithetic = False
  mAntitheticValue = 0#

End Sub

'============================================================
' 同じseedで再初期化する
'
' Seed未指定の場合は、新しく Randomize する。
'============================================================
Public Sub Reset()

  If mHasSeed Then
    Init mSeed, mUseAntithetic
  Else
    Init 0, mUseAntithetic
  End If

End Sub

'============================================================
' Summary
'============================================================
Public Function Summary() As String

  CheckInitialized

  Summary = "clsRandomNormal" & vbCrLf & _
            "HasSeed = " & CStr(mHasSeed) & vbCrLf & _
            "UseAntithetic = " & CStr(mUseAntithetic) & vbCrLf & _
            "HasCachedNormal = " & CStr(mHasCachedNormal) & vbCrLf & _
            "HasAntithetic = " & CStr(mHasAntithetic)

  If mHasSeed Then
    Summary = Summary & vbCrLf & "Seed = " & CStr(mSeed)
  End If

End Function

'============================================================
' 内部処理：Antitheticを考慮しない標準正規乱数
'============================================================
Private Function NextIndependentNormal() As Double

  Dim z1 As Double
  Dim z2 As Double

  If mHasCachedNormal Then
    NextIndependentNormal = mCachedNormal
    mCachedNormal = 0#
    mHasCachedNormal = False
    Exit Function
  End If

  NormalPair z1, z2

  NextIndependentNormal = z1
  mCachedNormal = z2
  mHasCachedNormal = True

End Function

'============================================================
' 内部処理：安全な Uniform(0,1)
'============================================================
Private Function SafeUniform() As Double

  Dim u As Double

  u = Rnd

  If u <= 0# Then
    u = MIN_UNIFORM
  End If

  If u >= 1# Then
    u = 1# - MIN_UNIFORM
  End If

  SafeUniform = u

End Function

'============================================================
' 内部処理：初期化チェック
'============================================================
Private Sub CheckInitialized()

  If Not mIsInitialized Then
    Err.Raise vbObjectError + 9091, _
      "clsRandomNormal", _
      "clsRandomNormal is not initialized. Call Init first."
  End If

End Sub
