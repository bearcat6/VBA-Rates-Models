# Hull-White 1F シミュレーター設計メモ

## 1. 目的

この文書は、`src/classes/clsHWSimulator.cls` の役割と初期実装方針を整理するための設計メモである。

`clsHWSimulator.cls` は、Discount curve と Hull-White 1F パラメータ `a`, `sigma` を使って、将来時点の金利カーブを Monte Carlo simulation で生成するためのクラスである。

最終的な目的は、以下である。

```text
・1年後などの将来時点の金利カーブを多数生成する
・将来ゼロレートカーブを作る
・将来DFカーブを作る
・将来フォワードカーブを作る
・年限ごとの分位点カーブを作る
・ストレスカーブ作成の材料にする
```

## 2. 対象モデル

Hull-White 1F モデルは以下で表される。

```text
dr(t) = { theta(t) - a r(t) } dt + sigma dW(t)
```

本リポジトリでは、実装上は以下の shifted short-rate representation を基本にする。

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

`clsHWSimulator.cls` は、このうち `x(t)` を Monte Carlo simulation する。

## 3. 初期実装の位置づけ

現在の `clsHWSimulator.cls` は、`clsHullWhite1F.cls` が未作成でも動く初期版である。

そのため、モデル本体の一部に相当する関数、たとえば以下も一時的に内部に持っている。

```text
・HW_B(t,T)
・EvolveX
・ForwardDF
・FutureDF
・FutureZeroRate
```

将来的に `clsHullWhite1F.cls` を作成した後は、`HW_B` や `StdDevX` などのモデル式は `clsHullWhite1F` または `modHWMath` に移し、`clsHWSimulator` は simulation workflow に集中させる方針である。

## 4. 依存関係

`clsHWSimulator` は、Discount curve を `Object` として受け取る。

Curve object は以下の Public method を持つ前提である。

```vb
DF_T(T)
InstantaneousForward(T)
```

対象となるカーブクラスは、たとえば以下である。

```text
・clsOISStepForwardCurve
・clsOISZeroLinearCurve
```

これらのカーブクラスには、Hull-White 用の time-based interface として `DF_T` や `InstantaneousForward` を追加している。

`clsHWSimulator` は、Excelシートには直接依存しない。

Excelからの入力や出力は、別途 `modExcelIO` または workflow module が担当する。

## 5. 初期化

主な初期化メソッドは以下である。

```vb
Public Sub Init( _
  ByVal in_Curve As Object, _
  ByVal in_a As Double, _
  ByVal in_sigma As Double, _
  ByVal in_HorizonYears As Double, _
  Optional ByVal in_Dt As Double = 0.00396825396825397 _
)
```

引数の意味は以下である。

```text
in_Curve
  DF_T, InstantaneousForward を持つ discount curve object

in_a
  Hull-White 1F の平均回帰速度

in_sigma
  short-rate factor のボラティリティ

in_HorizonYears
  simulation horizon
  例：1年後なら 1.0

in_Dt
  time step in years
  初期値 0.00396825396825397 は、おおむね 1/252 年を意識した値
```

## 6. x(t) のシミュレーション

`x(t)` は以下の Ornstein-Uhlenbeck process に従う。

```text
dx(t) = -a x(t) dt + sigma dW(t)
```

初期実装では、Euler ではなく exact discretization を使う。

```text
x(t+dt)
  = x(t) exp(-a dt)
    + sigma sqrt({1 - exp(-2a dt)} / (2a)) Z
```

ここで、`Z` は標準正規乱数である。

`a → 0` の極限では、以下を使う。

```text
x(t+dt) = x(t) + sigma sqrt(dt) Z
```

乱数は、初期実装ではクラス内部の Box-Muller 法で生成する。

将来的には `clsRandomNormal.cls` を作成し、乱数生成を分離することも考えられる。

## 7. short rate path

short rate は以下で計算する。

```text
r(t) = f(0,t) + x(t)
```

ここで、`f(0,t)` はカーブオブジェクトの `InstantaneousForward(t)` から取得する。

```vb
mShortRatePaths(p, s) = InitialForward(tNext) + xNext
```

初期実装では、`theta(t)` を直接計算・保存しない。

## 8. 将来DFカーブの近似

horizon 時点を `h`、満期を `T`、horizon時点の確率要因を `x_h` とする。

初期実装では、将来時点 `h` から見た discount factor を以下で近似する。

```text
P_h(T) = P(0,T) / P(0,h) × exp(-B(h,T) × x_h)
```

ここで、

```text
B(h,T) = {1 - exp(-a(T-h))} / a
```

である。

`a → 0` の場合は、

```text
B(h,T) = T - h
```

とする。

この式は、初期カーブの forward discount factor に、horizon時点の Hull-White factor による変化を掛ける実務的な近似である。

## 9. 将来ゼロレート

将来時点 `h` から見た、満期 `T` の zero rate は以下で計算する。

```text
ZeroRate_h(T) = -ln(P_h(T)) / (T - h)
```

VBA上では以下の Public method で取得する。

```vb
Public Function FutureZeroRate( _
  ByVal in_PathIndex As Long, _
  ByVal in_MaturityYears As Double _
) As Double
```

`in_MaturityYears` は、valuation date から見た maturity time である。

したがって、horizon が 1年後で、horizonから見た10年金利を見たい場合は、`in_MaturityYears = 11` のように扱う。

## 10. 将来フォワードレート

将来時点 `h` から見た `T1` から `T2` までの forward rate は以下で計算する。

```text
F_h(T1,T2) = {P_h(T1) / P_h(T2) - 1} / (T2 - T1)
```

VBA上では以下で取得する。

```vb
Public Function FutureForwardRate( _
  ByVal in_PathIndex As Long, _
  ByVal in_T1 As Double, _
  ByVal in_T2 As Double _
) As Double
```

`T1`, `T2` は valuation date から見た時点である。

## 11. 主な Public メソッド

`clsHWSimulator.cls` の主な Public メソッドは以下である。

```vb
Public Sub Init( _
  ByVal in_Curve As Object, _
  ByVal in_a As Double, _
  ByVal in_sigma As Double, _
  ByVal in_HorizonYears As Double, _
  Optional ByVal in_Dt As Double = 0.00396825396825397 _
)

Public Sub Simulate( _
  ByVal in_NumPaths As Long, _
  Optional ByVal in_Seed As Long = 0 _
)

Public Function GetTimeGrid() As Variant

Public Function GetXPathTable() As Variant

Public Function GetShortRatePathTable() As Variant

Public Function TerminalX(ByVal in_PathIndex As Long) As Double

Public Function TerminalShortRate(ByVal in_PathIndex As Long) As Double

Public Function FutureDF( _
  ByVal in_PathIndex As Long, _
  ByVal in_MaturityYears As Double _
) As Double

Public Function FutureZeroRate( _
  ByVal in_PathIndex As Long, _
  ByVal in_MaturityYears As Double _
) As Double

Public Function FutureForwardRate( _
  ByVal in_PathIndex As Long, _
  ByVal in_T1 As Double, _
  ByVal in_T2 As Double _
) As Double

Public Function GetFutureZeroCurveTable(ByVal in_MaturityYears As Variant) As Variant

Public Function GetFutureDFCurveTable(ByVal in_MaturityYears As Variant) As Variant

Public Function GetFutureZeroPercentileCurve( _
  ByVal in_MaturityYears As Variant, _
  ByVal in_Percentile As Double _
) As Variant

Public Function Summary() As String
```

## 12. 出力テーブル

`clsHWSimulator` は、Excelに貼り付けやすいように Variant 配列を返す。

主な出力は以下である。

```text
GetXPathTable
  ・各pathの x(t) path

GetShortRatePathTable
  ・各pathの short rate path

GetFutureZeroCurveTable
  ・各pathの horizon時点将来zero curve

GetFutureDFCurveTable
  ・各pathの horizon時点将来DF curve

GetFutureZeroPercentileCurve
  ・指定percentileの将来zero curve
```

たとえば、1年後の将来カーブを1,000本生成し、99%点カーブを取得する用途を想定している。

## 13. 分位点カーブ

`GetFutureZeroPercentileCurve` は、各maturityごとに全pathの将来zero rateを集め、指定percentileを計算する。

例：

```vb
pctCurve = sim.GetFutureZeroPercentileCurve(maturityGrid, 0.99)
```

この結果は、以下のような用途に使う。

```text
・99%点金利カーブ
・95%点金利カーブ
・中央値カーブ
・ストレスカーブ候補
・市場リスクストレステスト用の参考カーブ
```

## 14. 使い方イメージ

典型的な使い方は以下である。

```vb
Dim sim As clsHWSimulator
Dim mats(1 To 5) As Double
Dim pctCurve As Variant

Set sim = New clsHWSimulator

sim.Init curve, 0.03, 0.01, 1#, 1# / 252#
sim.Simulate 10000, 12345

mats(1) = 2#
mats(2) = 3#
mats(3) = 5#
mats(4) = 7#
mats(5) = 10#

pctCurve = sim.GetFutureZeroPercentileCurve(mats, 0.99)
```

この例では、valuation date から見た maturity を指定している。

horizon が1年の場合、`mats(1)=2` は、horizonから見れば1年先の満期を意味する。

## 15. 注意点と限界

現在の `clsHWSimulator.cls` は、将来カーブ生成の初期実装である。

重要な注意点は以下である。

```text
・theta(t) は明示的に扱っていない
・将来DFカーブは簡易近似式で生成している
・厳密な無裁定将来カーブ生成ではない
・乱数生成はクラス内部の簡易Box-Muller法である
・antithetic variates や control variate は未実装
・pathwise discounting や商品評価は未実装
```

したがって、現時点では、主に以下を目的として使う。

```text
・将来金利カーブの分布を見る
・分位点カーブを作る
・ストレスカーブ候補を作る
・Hull-White simulation workflow のプロトタイプを作る
```

## 16. 今後の拡張候補

今後の拡張候補は以下である。

```text
・clsHullWhite1F.cls を作成し、モデル式を分離する
・HW_B, EvolveX, StdDevX などを clsHullWhite1F または modHWMath へ移す
・clsRandomNormal.cls を作成し、乱数生成を分離する
・antithetic variates を追加する
・pathごとの discount factor を蓄積する
・商品評価用の Monte Carlo pricer に接続する
・将来カーブ生成式をより厳密な無裁定式に改善する
・PCAやSABR分位点カーブとの比較機能を追加する
・Excel出力用の modExcelIO と接続する
```

## 17. 関連クラス

関連する主なクラスは以下である。

```text
clsOISStepForwardCurve
  ・DF_T
  ・InstantaneousForward
  を提供する curve object

clsOISZeroLinearCurve
  ・DF_T
  ・InstantaneousForward
  を提供する curve object

clsVolSurface
  ・normal swaption vol surface を提供する

clsHWCalibrator
  ・a を固定し sigma を推定する初期キャリブレーター

clsHWSimulator
  ・a, sigma, curve を使って将来カーブを生成する
```

## 18. まとめ

`clsHWSimulator.cls` は、Hull-White 1F による将来金利カーブ生成の初期実装である。

初期実装で重視することは以下である。

```text
・Excelに依存しない
・curve object の DF_T と InstantaneousForward を使う
・a, sigma を外部から受け取る
・x(t) を exact discretization でシミュレーションする
・horizon時点の将来zero curve / DF curve を返す
・年限ごとの分位点カーブを返す
```

今後、`clsHullWhite1F.cls` を追加した段階で、モデル式をそちらに移し、`clsHWSimulator` はシミュレーション実行と出力に集中させる。