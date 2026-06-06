# Hull-White 1F Model Design

## 1. Purpose

`clsHullWhite1F.cls` は、1ファクター Hull-White モデルの中核ロジックを保持するクラスである。

主な目的は以下のとおり。

- 初期 Discount Curve に整合する Hull-White 1F モデルを表現する
- 平均回帰パラメータ `a` と短期金利ボラティリティ `sigma` を保持する
- `A(t,T)`、`B(t,T)` を計算し、将来時点のゼロクーポン債価格を返す
- shifted state `x(t)` から短期金利 `r(t)` を復元する
- Monte Carlo simulation で利用する 1 step exact transition を提供する

このクラスは、モンテカルロのパス生成ループそのものを持たない。
パス生成は `clsHWSimulator.cls` に分離し、本クラスはモデル式・係数・状態遷移を担当する。

---

## 2. Model Definition

Hull-White 1F モデルは、リスク中立測度の下で短期金利 `r(t)` が以下に従うとする。

```text
dr(t) = { theta(t) - a * r(t) } dt + sigma dW(t)
```

where:

- `a` : mean reversion speed
- `sigma` : short-rate volatility
- `theta(t)` : initial discount curve にフィットするための time-dependent drift
- `W(t)` : Brownian motion

VBA 実装では、`theta(t)` を直接保持せず、以下の shifted-state 表現を使う。

```text
r(t) = x(t) + phi(t)

dx(t) = -a * x(t) dt + sigma dW(t)

x(0) = 0
```

この表現により、シミュレーション対象は Ornstein-Uhlenbeck process である `x(t)` となり、
初期カーブへのフィットは `phi(t)` によって表現される。

---

## 3. Initial Curve Fit

初期 Discount Curve を `P(0,T)` とする。

瞬間フォワードレートは以下で定義される。

```text
f(0,t) = - d log P(0,t) / dt
```

`clsHullWhite1F` では、`P(0,t)` を `clsDiscountCurve.DF(targetDate)` から取得し、
`f(0,t)` は有限差分で近似する。

shift function `phi(t)` は以下で与える。

```text
phi(t) = f(0,t) + sigma^2 / (2 * a^2) * { 1 - exp(-a * t) }^2
```

これにより、モデルは初期 Discount Curve に整合する。

---

## 4. Zero Coupon Bond Formula

Hull-White 1F では、時点 `t` における満期 `T` のゼロクーポン債価格は以下で表される。

```text
P(t,T) = A(t,T) * exp( -B(t,T) * r(t) )
```

`B(t,T)` は以下。

```text
B(t,T) = { 1 - exp( -a * (T - t) ) } / a
```

`A(t,T)` は初期カーブに整合するように以下で計算する。

```text
A(t,T)
  = P(0,T) / P(0,t)
    * exp(
        B(t,T) * f(0,t)
        - sigma^2 / (4 * a) * { 1 - exp(-2 * a * t) } * B(t,T)^2
      )
```

したがって、将来時点 `t` の short rate `r(t)` が与えられれば、
`P(t,T)` を計算できる。

---

## 5. Class Responsibility

### clsHullWhite1F.cls

`clsHullWhite1F` は、Hull-White 1F モデルそのものを表現する。

担当するもの：

- Discount Curve の保持
- valuation date の保持
- `a` と `sigma` の保持
- `A(t,T)`、`B(t,T)` の計算
- `phi(t)` の計算
- `x(t)` から `r(t)` への変換
- `x(t)` の exact transition
- `x(t)` の分散・標準偏差
- 債券オプション等で使う bond log variance の補助計算

担当しないもの：

- swaption volatility への calibration
- Monte Carlo path の大量生成
- 分位点カーブの作成
- Excel Range への出力
- random number generation

これらは、別クラス・別モジュールに分ける。

---

## 6. Related Classes

想定されるクラス分担は以下。

```text
clsDiscountCurve
  Discount Curve interface.
  DF, ZeroRateCont, ForwardRate を提供する。

clsATMSwaptionVol
  ATM swaption volatility matrix を保持し、expiry x tenor の vol を返す。

clsHullWhite1F
  Hull-White 1F model の中核。
  A/B係数、phi、状態遷移、discount bond price を提供する。

clsHWCalibrator
  swaption volatility 等に対して a, sigma を calibration する。
  calibration 対象の market vol は clsATMSwaptionVol から取得する想定。

clsHWSimulator
  clsHullWhite1F を使って Monte Carlo path を生成する。
  1年後などの将来時点における金利カーブを多数生成する。

clsRandomNormal
  Monte Carlo simulation 用の標準正規乱数を生成する。
```

---

## 7. Main Public Methods

### Initialize

```vb
Public Sub Initialize( _
  ByVal in_DiscountCurve As clsDiscountCurve, _
  ByVal in_ValuationDate As Date, _
  ByVal in_MeanReversion As Double, _
  ByVal in_Sigma As Double)
```

Discount Curve、評価日、平均回帰、ボラティリティを設定する。

Validation:

- Discount Curve が Nothing の場合はエラー
- `a <= 0` の場合はエラー
- `sigma < 0` の場合はエラー

---

### B

```vb
Public Function B( _
  ByVal in_T As Double, _
  ByVal in_TerminalT As Double) As Double
```

`B(t,T)` を返す。

```text
B(t,T) = { 1 - exp( -a * (T - t) ) } / a
```

---

### A

```vb
Public Function A( _
  ByVal in_T As Double, _
  ByVal in_TerminalT As Double) As Double
```

`A(t,T)` を返す。

`A(t,T)` は、初期 Discount Curve にフィットするように計算する。

---

### DiscountBondByShortRate

```vb
Public Function DiscountBondByShortRate( _
  ByVal in_T As Double, _
  ByVal in_TerminalT As Double, _
  ByVal in_ShortRate As Double) As Double
```

短期金利 `r(t)` が与えられたとき、`P(t,T)` を返す。

```text
P(t,T) = A(t,T) * exp( -B(t,T) * r(t) )
```

---

### DiscountBondByState

```vb
Public Function DiscountBondByState( _
  ByVal in_T As Double, _
  ByVal in_TerminalT As Double, _
  ByVal in_StateX As Double) As Double
```

shifted state `x(t)` が与えられたとき、

```text
r(t) = x(t) + phi(t)
```

として `P(t,T)` を返す。

---

### DFByYears

```vb
Public Function DFByYears(ByVal in_T As Double) As Double
```

valuation date から `in_T` 年後の初期 Discount Factor を返す。

内部的には、`in_T` を日付に変換し、`clsDiscountCurve.DF(targetDate)` を呼び出す。

---

### ZeroRateContByYears

```vb
Public Function ZeroRateContByYears(ByVal in_T As Double) As Double
```

初期 Discount Factor から連続複利ゼロレートを返す。

```text
z(0,T) = -log P(0,T) / T
```

---

### InstantForward

```vb
Public Function InstantForward(ByVal in_T As Double) As Double
```

初期 Discount Curve から瞬間フォワードレート `f(0,t)` を返す。

```text
f(0,t) = - d log P(0,t) / dt
```

VBA 実装では有限差分で近似する。

---

### Phi

```vb
Public Function Phi(ByVal in_T As Double) As Double
```

shift function `phi(t)` を返す。

```text
phi(t) = f(0,t) + sigma^2 / (2 * a^2) * { 1 - exp(-a * t) }^2
```

---

### ShortRateFromState

```vb
Public Function ShortRateFromState( _
  ByVal in_T As Double, _
  ByVal in_StateX As Double) As Double
```

shifted state `x(t)` から short rate `r(t)` を返す。

```text
r(t) = x(t) + phi(t)
```

---

### NextStateExact

```vb
Public Function NextStateExact( _
  ByVal in_StateX As Double, _
  ByVal in_Dt As Double, _
  ByVal in_NormalZ As Double) As Double
```

Ornstein-Uhlenbeck process の exact transition により、次時点の `x` を返す。

```text
x(t+dt)
  = x(t) * exp(-a * dt)
    + sigma * sqrt( {1 - exp(-2 * a * dt)} / (2 * a) ) * Z
```

where:

- `Z` : standard normal random variable

---

### NextShortRateExact

```vb
Public Function NextShortRateExact( _
  ByVal in_T As Double, _
  ByVal in_StateX As Double, _
  ByVal in_Dt As Double, _
  ByVal in_NormalZ As Double) As Double
```

`NextStateExact` により `x(t+dt)` を生成し、
`ShortRateFromState` により `r(t+dt)` を返す。

---

### StateVariance / StateStdDev

```vb
Public Function StateVariance(ByVal in_T As Double) As Double
Public Function StateStdDev(ByVal in_T As Double) As Double
```

`x(0)=0` とした場合の `x(t)` の分散・標準偏差を返す。

```text
Var[x(t)] = sigma^2 * {1 - exp(-2 * a * t)} / (2 * a)
```

---

### StateStepVariance / StateStepStdDev

```vb
Public Function StateStepVariance(ByVal in_Dt As Double) As Double
Public Function StateStepStdDev(ByVal in_Dt As Double) As Double
```

`x(t)` から `x(t+dt)` への conditional variance / standard deviation を返す。

```text
Var[x(t+dt) | x(t)]
  = sigma^2 * {1 - exp(-2 * a * dt)} / (2 * a)
```

---

### BondLogVariance

```vb
Public Function BondLogVariance( _
  ByVal in_OptionExpiry As Double, _
  ByVal in_BondMaturity As Double) As Double
```

債券オプション等で利用する bond log variance の補助関数。

将来、Jamshidian decomposition による swaption calibration を実装する際の部品として利用する。

---

## 8. Time Convention

`clsHullWhite1F` 内部では、時刻は valuation date からの年数で扱う。

```text
t = 0.0     valuation date
T = 1.0     valuation date から1年後
T = 5.0     valuation date から5年後
```

一方、Discount Curve は `Date` を引数に取るため、
`DFByYears` 内で年数を日付に変換する。

現在の実装では、以下を使用する。

```text
1 year = 365.25 days
```

この day count は簡易実装であり、将来的には ACT/365F、ACT/360、営業日調整などの専用クラスに分離する余地がある。

---

## 9. Simulation Image

1年後の将来金利カーブを Monte Carlo で生成する場合の流れは以下。

```text
1. Discount Curve を構築する

2. clsHullWhite1F を初期化する
   - curve
   - valuation date
   - a
   - sigma

3. clsHWSimulator が path を生成する
   - x(0) = 0
   - Z を生成
   - NextStateExact で x(t+dt) を更新

4. 1年後の state x(1Y) を得る

5. 各満期 T について P(1Y,T) を計算する
   - DiscountBondByState(1Y, T, x(1Y))

6. P(1Y,T) からゼロレートまたはフォワードレートを復元する

7. これを多数パスで繰り返し、分位点カーブを作る
```

---

## 10. Calibration Image

`clsHullWhite1F` 自体は calibration を行わない。

calibration は `clsHWCalibrator` の責務とする。

想定される calibration の流れは以下。

```text
1. 初期 Discount Curve を用意する

2. clsATMSwaptionVol から market ATM swaption vol を取得する

3. a, sigma の候補値を設定する

4. Hull-White 1F で model swaption value / model implied vol を計算する

5. market vol との差を評価する

6. 最小二乗等で a, sigma を決定する
```

初期段階では、以下のように単純な方針でよい。

```text
- a は固定値として与える
- sigma のみ ATM swaption vol に合わせる
```

その後、複数 expiry / tenor の swaption vol に対して、`a` と `sigma` を同時に calibration する。

---

## 11. Current Limitations

現在の `clsHullWhite1F` は、最小限のモデル中核クラスである。

制約・今後の改善点は以下。

- `a` と `sigma` は constant parameter
- time-dependent volatility には未対応
- `theta(t)` は明示的には保持しない
- instantaneous forward は有限差分近似
- day count は 365.25 日固定
- business day adjustment は未対応
- swaption pricing は未実装
- calibration は未実装
- full yield curve simulation は `clsHWSimulator` 側で実装予定

---

## 12. Design Policy

`clsHullWhite1F` は、次の方針で維持する。

- モデル式に対応する関数を明示的に持つ
- calibration や simulation loop を持たせすぎない
- `clsDiscountCurve` に依存し、具体的な curve construction には依存しない
- 将来の `clsHWCalibrator`、`clsHWSimulator` から再利用しやすい形にする
- Excel Range や Worksheet への直接依存を避ける

この分離により、VBA であっても、以下の責務分担を保つ。

```text
Curve construction
  -> Discount Curve class

Market volatility input
  -> clsATMSwaptionVol

Model definition
  -> clsHullWhite1F

Calibration
  -> clsHWCalibrator

Monte Carlo path generation
  -> clsHWSimulator

Output / report
  -> standard module or worksheet-facing procedure
```
