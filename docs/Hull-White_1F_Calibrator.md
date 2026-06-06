# Hull-White 1F キャリブレーター設計メモ

## 1. 目的

この文書は、`src/classes/clsHWCalibrator.cls` の役割と初期実装方針を整理するための設計メモである。

`clsHWCalibrator.cls` は、Discount curve と volatility surface を使って、1-factor Hull-White モデルのパラメータを推定するためのクラスである。

初期実装では、平均回帰パラメータ `a` は外部から固定値として与え、ボラティリティ・パラメータ `sigma` を ATM normal swaption volatility にフィットする。

## 2. 対象モデル

Hull-White 1F モデルは以下で表される。

```text
dr(t) = { theta(t) - a r(t) } dt + sigma dW(t)
```

ただし、本リポジトリの実装では、直接 `theta(t)` を入力・保存するのではなく、以下の shifted short-rate representation を基本にする。

```text
r(t) = f(0,t) + x(t)

dx(t) = -a x(t) dt + sigma dW(t)
```

ここで、

```text
f(0,t)
  = 初期カーブから得られる瞬間フォワード

x(t)
  = 平均回帰する確率要因

a
  = 平均回帰速度

sigma
  = short-rate factor のボラティリティ
```

`theta(t)` は、初期Discount curveを再現するための内部的な調整項であり、初期実装では明示的に保持しない。

## 3. clsHWCalibrator の初期スコープ

初期の `clsHWCalibrator.cls` は、厳密な swaption pricing engine ではなく、Hull-White 1F の初期パラメータ推定用の簡易キャリブレーターである。

初期スコープは以下である。

```text
・Discount curve object を受け取る
・clsVolSurface を受け取る
・Calibration quote として Expiry × Tenor の点を保持する
・平均回帰 a は固定値として与える
・sigma を1次元探索で推定する
・市場ボラは clsVolSurface.NormalVol(expiry, tenor) から取得する
・目的関数は market normal vol と model normal vol の二乗誤差合計
```

## 4. 依存関係

`clsHWCalibrator` は、Excelシートには直接依存しない。

想定する依存関係は以下である。

```text
clsHWCalibrator
  uses:
    ・Curve object
    ・clsVolSurface
```

Curve object は、以下の Public method を持つことを前提とする。

```vb
DF_T(T)
ForwardRate_T(T1, T2)
InstantaneousForward(T)
```

実際の初期実装で主に使うのは `DF_T(T)` である。

対象となるカーブクラスは、たとえば以下である。

```text
・clsOISStepForwardCurve
・clsOISZeroLinearCurve
```

これらのクラスには、Hull-White 用の time-based interface として `DF_T` などを追加している。

Volatility surface は `clsVolSurface` として受け取る。

```vb
Private mVolSurface As clsVolSurface
```

市場ボラティリティは以下で取得する。

```vb
marketVol = mVolSurface.NormalVol(expiryYears, tenorYears)
```

## 5. なぜ Excel Range に依存させないか

`clsHWCalibrator` はモデル層のクラスであり、Excelシート構造を直接知るべきではない。

Excelからの読み込みは、別途 `modExcelIO` が担当する。

想定する流れは以下である。

```text
Excel sheet
  → modExcelIO が discount curve input と volatility input を読み込む
  → curve object と clsVolSurface を生成する
  → clsHWCalibrator に渡す
```

この構成にすることで、以下の利点がある。

```text
・モデルクラスがExcelの表構造に依存しない
・テストしやすい
・将来、CSV、DB、Bloomberg API 等から入力する場合にも流用しやすい
・キャリブレーションロジックと入出力処理を分離できる
```

## 6. キャリブレーション対象

初期実装では、`a` を固定して `sigma` を推定する。

```text
a      : external input / fixed
sigma  : calibrated parameter
```

これは、2変数同時最適化よりも安定しており、初期実装として扱いやすい。

将来的には、以下の拡張が考えられる。

```text
・a と sigma の2次元最適化
・expiry bucketごとの piecewise sigma
・time-dependent sigma(t)
・Jamshidian decomposition による厳密な swaption pricing
```

## 7. 市場ボラティリティ

市場ボラティリティは、`clsVolSurface` から取得する。

初期実装では normal volatility を前提とする。

```vb
marketVol = mVolSurface.NormalVol(expiryYears, tenorYears)
```

ここで、ボラティリティは年率換算された normal volatility として扱う。

例：

```text
0.0050 = 50bp normal vol
```

## 8. モデルボラティリティの近似

初期実装では、厳密な Hull-White swaption pricing ではなく、swap rate の short-rate factor `x` に対する一次感応度を使って normal swaption vol を近似する。

基本式は以下である。

```text
modelVol = |dS/dx| × Std[x(T)] / sqrt(T)
```

ここで、

```text
T
  = swaption expiry

S
  = expiry時点スタートの forward swap rate

x(T)
  = shifted short-rate factor

dS/dx
  = swap rate の x に対する一次感応度

Std[x(T)]
  = x(T) の標準偏差
```

`x(T)` の分散は以下である。

```text
Var[x(T)] = sigma^2 × {1 - exp(-2aT)} / (2a)
```

したがって、標準偏差は以下となる。

```text
Std[x(T)] = sigma × sqrt({1 - exp(-2aT)} / (2a))
```

`a → 0` の極限では、以下を使う。

```text
Std[x(T)] = sigma × sqrt(T)
```

## 9. Swap rate sensitivity の考え方

Expiry時点 `T` から見た、将来時点 `U` の forward discount factor を以下とする。

```text
P_T(U) = P(0,U) / P(0,T)
```

Hull-White の bond sensitivity は以下である。

```text
B(T,U) = {1 - exp(-a(U-T))} / a
```

`a → 0` の場合は、

```text
B(T,U) = U - T
```

とする。

forward discount factor の `x` に対する一次感応度は、以下で近似する。

```text
dP_T(U)/dx = -B(T,U) × P_T(U)
```

forward swap rate を、

```text
S = N / A
```

とする。

ここで、

```text
N = 1 - P_T(end)
A = Σ accrual_i × P_T(pay_i)
```

である。

このとき、

```text
dS/dx = (dN × A - N × dA) / A^2
```

を使って swap rate sensitivity を求める。

## 10. 目的関数

`sigma` のキャリブレーションでは、以下の二乗誤差合計を最小化する。

```text
Objective(sigma | a)
  = Σ_i { ModelNormalVol_i(a, sigma) - MarketNormalVol_i }^2
```

ここで、各 quote は以下で表される。

```text
Quote_i = (expiry_i, tenor_i)
```

市場ボラティリティは、

```vb
mVolSurface.NormalVol(expiry_i, tenor_i)
```

で取得する。

モデルボラティリティは、

```vb
ModelNormalVol(a, sigma, expiry_i, tenor_i)
```

で計算する。

## 11. 最適化方法

初期実装では、`a` を固定し、`sigma` のみを golden section search で探索する。

主な理由は以下である。

```text
・1次元探索なので実装が単純
・Excel VBAで安定しやすい
・初期パラメータ推定として十分使いやすい
・後で2次元最適化に拡張しやすい
```

主なメソッドは以下である。

```vb
Public Function CalibrateSigmaFixedA( _
  ByVal in_a As Double, _
  Optional ByVal in_SigmaLower As Double = 0.000001, _
  Optional ByVal in_SigmaUpper As Double = 0.05, _
  Optional ByVal in_Tolerance As Double = 0.0000000001, _
  Optional ByVal in_MaxIterations As Long = 100 _
) As Double
```

戻り値は、推定された `sigma` である。

## 12. Public メソッド一覧

`clsHWCalibrator.cls` の主な Public メソッドは以下である。

```vb
Public Sub Init( _
  ByVal in_Curve As Object, _
  ByVal in_VolSurface As clsVolSurface, _
  Optional ByVal in_PaymentFrequencyYears As Double = 1# _
)

Public Sub ClearCalibrationQuotes()

Public Sub AddCalibrationQuote( _
  ByVal in_ExpiryYears As Double, _
  ByVal in_TenorYears As Double _
)

Public Sub BuildCalibrationQuotes( _
  ByVal in_ExpiryYears As Variant, _
  ByVal in_TenorYears As Variant _
)

Public Function CalibrateSigmaFixedA( _
  ByVal in_a As Double, _
  Optional ByVal in_SigmaLower As Double = 0.000001, _
  Optional ByVal in_SigmaUpper As Double = 0.05, _
  Optional ByVal in_Tolerance As Double = 0.0000000001, _
  Optional ByVal in_MaxIterations As Long = 100 _
) As Double

Public Function ObjectiveValue( _
  ByVal in_a As Double, _
  ByVal in_sigma As Double _
) As Double

Public Function ModelNormalVol( _
  ByVal in_a As Double, _
  ByVal in_sigma As Double, _
  ByVal in_ExpiryYears As Double, _
  ByVal in_TenorYears As Double _
) As Double

Public Function ForwardSwapRate( _
  ByVal in_ExpiryYears As Double, _
  ByVal in_TenorYears As Double _
) As Double

Public Function GetCalibrationReport() As Variant

Public Function Summary() As String
```

## 13. 使い方イメージ

典型的な使い方は以下である。

```vb
Dim calib As clsHWCalibrator
Dim sigma As Double

Set calib = New clsHWCalibrator

calib.Init curve, volSurface, 1#

calib.AddCalibrationQuote 1#, 5#
calib.AddCalibrationQuote 1#, 10#
calib.AddCalibrationQuote 2#, 5#
calib.AddCalibrationQuote 2#, 10#

sigma = calib.CalibrateSigmaFixedA( _
            in_a:=0.03, _
            in_SigmaLower:=0.000001, _
            in_SigmaUpper:=0.05 _
        )

Debug.Print calib.Summary
```

キャリブレーション後、以下で各quoteの市場ボラ、モデルボラ、誤差を確認できる。

```vb
Dim report As Variant
report = calib.GetCalibrationReport()
```

## 14. 限界と注意点

現在の `clsHWCalibrator.cls` は、初期推定用の簡易キャリブレーターである。

重要な注意点は以下である。

```text
・厳密な Hull-White swaption pricing ではない
・Jamshidian decomposition は未実装
・a は固定であり、同時最適化していない
・sigma は定数として扱う
・normal vol の近似式に基づく
・swap rate sensitivity は一次近似である
```

したがって、得られた `sigma` は、厳密な市場整合パラメータというより、Monte Carlo simulation や初期分析に使うための実務的な初期推定値と位置づける。

## 15. 今後の拡張候補

今後の拡張候補は以下である。

```text
・clsHullWhite1F.cls にモデル本体を分離する
・HW_B, StdDevX などを clsHullWhite1F または modHWMath に移す
・Jamshidian decomposition による swaption pricing を追加する
・a と sigma の2次元最適化を追加する
・sigma(t) の time-dependent calibration を追加する
・キャリブレーションquoteごとのウェイトを導入する
・Calibration report に relative error や weight を追加する
```

特に、次の段階では `clsHullWhite1F.cls` を作成し、以下のようなモデル本体ロジックを移すとよい。

```text
・B(t,T)
・StdDevX(t)
・ShortRate(t,x)
・DiscountBond 関連式
・将来カーブ生成用の補助関数
```

## 16. まとめ

`clsHWCalibrator.cls` は、Hull-White 1F 実装における初期キャリブレーション用クラスである。

初期実装では、以下を重視する。

```text
・Excelシートに依存しない
・Discount curve と clsVolSurface を受け取る
・a は固定、sigma を推定する
・ATM normal swaption vol に対して最小二乗でフィットする
・Monte Carlo simulation 用の初期パラメータ推定として使う
```

厳密な swaption pricing や time-dependent parameter calibration は、今後の拡張課題とする。
