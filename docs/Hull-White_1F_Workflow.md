# Hull-White 1F Workflow

## 1. 目的

本ドキュメントは、VBA-Rates-Models における Hull-White 1F モデルの利用手順を整理するものである。

最終目的は、JPY OIS カーブおよびスワップション・ボラティリティ等を入力として、1ファクター Hull-White モデルを用いて将来金利カーブをモンテカルロ・シミュレーションし、分位点カーブおよびストレスカーブを作成することである。

具体的には、以下の流れを想定する。

```text
1. Discount curve を準備する
2. Swaption volatility surface を準備する
3. Hull-White 1F のパラメータを決める
4. 将来時点の金利カーブを Monte Carlo simulation する
5. パス別カーブから分位点カーブを作成する
6. リスク管理用のストレスカーブとして利用する
```

本ドキュメントは、実装の使い方だけでなく、各 class / module の責務分担も合わせて整理する。

---

## 2. 対象範囲

### 2.1 初期実装の対象

初期実装では、以下を対象とする。

- JPY TONA / OIS ベースの discount curve
- ATM normal swaption volatility surface
- 1-factor Hull-White model
- 平均回帰パラメータ `a` は外部から固定値として与える
- ボラティリティ `sigma` は ATM normal swaption volatility へフィットする
- 将来時点、例えば 1年後の金利カーブを Monte Carlo simulation する
- シミュレーション結果から分位点カーブを作成する

### 2.2 初期実装では対象外とするもの

初期実装では、以下は対象外または将来拡張とする。

- 厳密な Jamshidian decomposition による swaption pricing
- `a` と `sigma` の同時最適化
- 非ATM swaption smile / SABR smile を直接使った Hull-White calibration
- 2ファクター Hull-White モデル
- 完全な無裁定条件を満たす将来カーブ生成の厳密実装
- 実取引評価用のフルスペック pricer

初期実装は、リスク管理上の分位点カーブ・ストレスカーブ作成を目的とした実務的なプロトタイプと位置づける。

---

## 3. 全体ワークフロー

全体の処理は、以下の順で行う。

```text
入力データ
  ├─ OIS discount curve
  ├─ ATM swaption normal vol surface
  ├─ calibration quote grid
  ├─ fixed mean reversion a
  ├─ simulation horizon
  ├─ number of paths
  └─ output tenors

処理
  ├─ clsVolSurface を作成
  ├─ clsHWCalibrator を作成
  ├─ fixed a に対して sigma を推定
  ├─ clsHWSimulator を作成
  ├─ x(t) を Monte Carlo simulation
  ├─ horizon 時点の将来カーブを生成
  └─ 分位点カーブ・ストレスカーブを出力

出力
  ├─ calibrated parameters
  ├─ calibration report
  ├─ path別 future curve
  ├─ percentile curve
  └─ stress curve
```

---

## 4. 入力データ

### 4.1 Discount curve

Hull-White 1F では、初期 discount curve が必要となる。

入力カーブは、具体的な構築方法に依存しないようにする。

想定する具体クラスは、例えば以下である。

```text
clsOISStepForwardCurve
clsOISZeroLinearCurve
```

Hull-White 関連クラスから利用するため、curve object は以下の time-based interface を持つ必要がある。

```vb
Public Function DF_T(ByVal T As Double) As Double

Public Function InstantaneousForward(ByVal T As Double, _
                                     Optional ByVal eps As Double = 0.0001) As Double
```

必要に応じて、以下も利用する。

```vb
Public Function ZeroRate_T(ByVal T As Double) As Double

Public Function ForwardRate_T(ByVal T1 As Double, _
                              ByVal T2 As Double) As Double
```

ここで `T` は、評価日からの year fraction を表す。

### 4.2 Swaption volatility surface

Hull-White calibration では、ATM normal swaption volatility surface を使用する。

入力形式は、Expiry × Tenor のマトリックスを想定する。

```text
          1Y      2Y      5Y      10Y
1Y        vol     vol     vol     vol
2Y        vol     vol     vol     vol
5Y        vol     vol     vol     vol
10Y       vol     vol     vol     vol
```

ボラティリティは normal volatility の絶対値表記とする。

```text
40bp normal vol = 0.0040
```

ボラティリティ・サーフェスは `clsVolSurface` に格納する。

初期実装では、以下の補間方針を採用する。

- Tenor方向は variance = vol^2 を線形補間する
- Expiry方向は total variance = vol^2 × expiry を線形補間する
- 範囲外は原則エラーとする
- 明示的に許可した場合のみ、端点ボラによる flat extrapolation を行う

### 4.3 Calibration quote grid

キャリブレーションでは、vol surface 全体を必ずしもすべて使う必要はない。

使用する quote を、Expiry と Tenor の組み合わせとして指定する。

例：

```text
Quote No.   Expiry   Tenor
1           1Y       5Y
2           2Y       5Y
3           5Y       5Y
4           1Y       10Y
5           2Y       10Y
6           5Y       10Y
```

初期実装では、これらの quote に対して model normal vol と market normal vol の差を評価し、`sigma` を推定する。

### 4.4 Mean reversion `a`

初期実装では、mean reversion `a` は外部から固定値として与える。

例：

```text
a = 0.03
```

`a` は平均回帰の速さを表す。

- `a` が大きいほど、短期金利ファクター `x(t)` は速く平均回帰する
- `a` が小さいほど、ショックは長く残りやすい

初期段階では、複数の `a` 候補に対して `sigma` を推定し、calibration error やシミュレーション結果の形状を比較する運用も考えられる。

### 4.5 Simulation horizon

将来カーブを作成する時点を指定する。

例：

```text
0.25 = 3か月後
0.50 = 6か月後
1.00 = 1年後
```

リスク管理資料では、1年後カーブを中心に利用する想定とする。

### 4.6 Number of paths

Monte Carlo simulation のパス数を指定する。

例：

```text
1,000 paths
10,000 paths
50,000 paths
```

初期検証では 1,000 paths 程度で動作確認し、分位点カーブ作成では 10,000 paths 以上を目安とする。

### 4.7 Output tenors

将来カーブを出力する年限を指定する。

例：

```text
0.25Y
0.50Y
1Y
2Y
3Y
5Y
7Y
10Y
15Y
20Y
30Y
```

この output tenors に対して、各パスの zero rate または forward rate を出力する。

---

## 5. 利用する class / module

### 5.1 `clsVolSurface`

`clsVolSurface` は、volatility surface を保持するクラスである。

主な責務は以下のとおり。

- Expiry grid の保持
- Tenor grid の保持
- Vol matrix の保持
- 任意の Expiry × Tenor に対する補間済み normal vol の返却
- calibration 用 market vol の供給

Hull-White calibration では、`NormalVol(expiry, tenor)` を通じて市場ボラティリティを取得する。

### 5.2 `clsHWCalibrator`

`clsHWCalibrator` は、Hull-White 1F のパラメータ推定を担当する。

初期実装では、以下の方針とする。

- `a` は外部から固定値として受け取る
- `sigma` を ATM normal swaption vol にフィットする
- market vol と model vol の二乗誤差を最小化する
- model vol は swap rate sensitivity を用いた近似により計算する

本クラスは、厳密な swaption pricer ではなく、初期パラメータ推定用の実務的な calibrator と位置づける。

将来的には、`clsHullWhite1F` または `mdl_HullWhiteMath` に実装する bond option / swaption pricing を利用して、より厳密な objective function に置き換える。

### 5.3 `clsHullWhite1F`

`clsHullWhite1F` は、Hull-White 1F モデル本体を表すクラスである。

主な責務は以下のとおり。

- discount curve の保持
- valuation date の保持
- mean reversion `a` の保持
- volatility `sigma` の保持
- `A(t,T)` / `B(t,T)` 等のモデル関数
- shifted short-rate representation に基づく計算

モデルは以下で表される。

```text
dr(t) = { theta(t) - a r(t) } dt + sigma dW(t)
```

実装上は、以下の shifted representation を利用する。

```text
r(t) = phi(t) + x(t)

dx(t) = -a x(t) dt + sigma dW(t)
```

ここで、`phi(t)` は初期 discount curve と整合するように決まる項であり、`theta(t)` を直接入力することはしない。

### 5.4 `clsHWSimulator`

`clsHWSimulator` は、Hull-White 1F の Monte Carlo simulation を担当する。

主な責務は以下のとおり。

- short-rate factor `x(t)` のパス生成
- exact discretization による OU process のシミュレーション
- horizon 時点の short rate の生成
- horizon 時点の zero curve / discount curve / forward curve の生成
- パス別カーブの出力

`x(t)` は以下の過程に従う。

```text
dx(t) = -a x(t) dt + sigma dW(t)
```

exact discretization は以下である。

```text
x(t + dt)
  = x(t) exp(-a dt)
    + sigma sqrt((1 - exp(-2a dt)) / (2a)) Z
```

`a -> 0` の極限では、以下を使う。

```text
x(t + dt) = x(t) + sigma sqrt(dt) Z
```

### 5.5 `clsRandomNormal`

`clsRandomNormal` は、標準正規乱数を生成するクラスである。

主な責務は以下のとおり。

- Box-Muller 法による標準正規乱数生成
- seed 指定による再現性確保
- 必要に応じた antithetic variates の利用

Monte Carlo simulation の乱数生成を `clsHWSimulator` から分離することで、将来的な乱数生成方法の差し替えを容易にする。

### 5.6 `mdl_HullWhiteMath`

`mdl_HullWhiteMath` は、Hull-White 1F の共通数理関数を置く標準モジュールである。

主な責務は以下のとおり。

- パラメータ検証
- `B(t,T)` 等の基本関数
- shifted representation に関する補助関数
- 将来拡張として bond option / swaption pricing に関する数理関数

この module は、Excel sheet の読み書きには依存しない。

### 5.7 `mdl_HullWhiteWorkFlow`

`mdl_HullWhiteWorkFlow` は、Excel/VBA から一連の処理を実行するための workflow module である。

主な責務は以下のとおり。

- Excel Range から `clsVolSurface` を作成する
- `clsHWCalibrator` を作成する
- fixed `a` に対して `sigma` を推定する
- `clsHWSimulator` を作成する
- simulation 結果を Excel sheet へ出力する

モデル本体や calibrator に Excel I/O を混ぜず、workflow module に集約する方針とする。

---

## 6. キャリブレーション手順

### 6.1 手順概要

キャリブレーションは、以下の順で行う。

```text
1. discount curve を作成する
2. volatility surface を作成する
3. calibration quote grid を指定する
4. fixed a を指定する
5. clsHWCalibrator を初期化する
6. market vol と model vol の誤差を評価する
7. sigma を探索する
8. calibration result を出力する
```

### 6.2 Market vol の取得

各 calibration quote について、`clsVolSurface` から market normal vol を取得する。

```text
marketVol_i = NormalVol(expiry_i, tenor_i)
```

### 6.3 Model vol の計算

初期実装では、swaption の厳密評価ではなく、swap rate の一次感応度を用いた近似で model normal vol を計算する。

考え方は以下である。

```text
modelVol_i
  = |dS/dx| × Std[x(expiry_i)] / sqrt(expiry_i)
```

ここで、`x(t)` は Hull-White の shifted short-rate factor である。

`x(T)` の標準偏差は以下である。

```text
Std[x(T)]
  = sigma × sqrt((1 - exp(-2aT)) / (2a))
```

`a -> 0` の極限では以下となる。

```text
Std[x(T)] = sigma × sqrt(T)
```

### 6.4 Objective function

目的関数は、market vol と model vol の二乗誤差とする。

```text
Objective(a, sigma)
  = Σ_i { modelVol_i(a, sigma) - marketVol_i }^2
```

初期実装では `a` を固定するため、探索対象は `sigma` のみである。

```text
Objective(sigma | fixed a)
  = Σ_i { modelVol_i(fixed a, sigma) - marketVol_i }^2
```

### 6.5 Calibration result

キャリブレーション結果として、以下を出力する。

```text
a
sigma
objective
quote count
```

また、quote 別の calibration report として、以下を出力する。

```text
quote no.
expiry
tenor
market normal vol
model normal vol
error
squared error
```

### 6.6 複数の `a` 候補を比較する運用

初期実装では `a` を固定するが、実務上は複数の `a` 候補を比較するとよい。

例：

```text
a = 0.01
a = 0.03
a = 0.05
a = 0.10
```

各 `a` について `sigma` を推定し、以下を比較する。

- objective の大きさ
- quote 別 error の偏り
- 1年後カーブ分布の形状
- 長期年限の分位点カーブの安定性
- ストレスカーブとしての説明可能性

---

## 7. シミュレーション手順

### 7.1 手順概要

シミュレーションは、以下の順で行う。

```text
1. calibrated a, sigma を取得する
2. simulation horizon を指定する
3. time step dt を指定する
4. number of paths を指定する
5. output tenors を指定する
6. clsHWSimulator を初期化する
7. x(t) を horizon まで simulation する
8. 各 path について horizon 時点の将来カーブを作成する
9. path別カーブを出力する
10. 分位点カーブを作成する
```

### 7.2 Time step

`dt` は年単位で指定する。

例：

```text
1営業日相当  = 1 / 252
1カレンダー日相当 = 1 / 365
1か月相当 = 1 / 12
```

初期実装では、日次程度の time step を基本とする。

### 7.3 Simulated factor

シミュレーション対象は、短期金利そのものではなく shifted short-rate factor `x(t)` である。

```text
r(t) = phi(t) + x(t)
```

ここで、`phi(t)` は初期カーブと整合するための deterministic shift である。

### 7.4 Future curve の作成

horizon 時点の将来カーブは、初期 discount curve と simulated factor `x(horizon)` を用いて作成する。

初期実装では、以下の考え方で将来カーブを生成する。

```text
初期カーブ
  + Hull-White bond sensitivity
  + simulated x(horizon)
  → horizon時点の将来 zero / forward curve
```

この方法は、分位点カーブ・ストレスカーブ作成のための実務的な初期版である。

将来的には、より厳密な条件付き discount bond formula に基づく実装へ拡張する。

---

## 8. 出力データ

### 8.1 Calibration summary

キャリブレーション結果の summary は、以下の形式を想定する。

```text
a        sigma      objective      quote count
0.0300   0.0065     0.00000012     9
```

### 8.2 Calibration report

quote 別の report は、以下の形式を想定する。

```text
No.   Expiry   Tenor   MarketVol   ModelVol   Error   SquaredError
1     1.0      5.0     0.0040      0.0039     -0.0001 0.00000001
2     2.0      5.0     0.0045      0.0046      0.0001 0.00000001
```

### 8.3 Path別 future curve

path 別の将来カーブは、以下の形式を想定する。

```text
Path   Tenor   ZeroRate   ForwardRate   DiscountFactor
1      1Y      0.0040     0.0045        0.9960
1      2Y      0.0050     0.0055        0.9900
1      5Y      0.0070     0.0080        0.9656
2      1Y      0.0035     0.0040        0.9965
2      2Y      0.0048     0.0053        0.9904
```

または、Excelで扱いやすいように、横持ち形式にしてもよい。

```text
Path   1Y      2Y      3Y      5Y      10Y
1      0.0040  0.0050  0.0058  0.0070  0.0100
2      0.0035  0.0048  0.0056  0.0068  0.0098
```

### 8.4 Percentile curve

分位点カーブは、各 tenor ごとに path 分布の分位点を取って作成する。

例：

```text
Tenor   P1      P5      P10     P50     P90     P95     P99
1Y      ...     ...     ...     ...     ...     ...     ...
2Y      ...     ...     ...     ...     ...     ...     ...
5Y      ...     ...     ...     ...     ...     ...     ...
10Y     ...     ...     ...     ...     ...     ...     ...
```

上方ストレスカーブを作る場合は、例えば P95 または P99 を各 tenor で連結する。

```text
Stress curve = P99 curve
```

下方ストレスカーブを作る場合は、P1 または P5 を各 tenor で連結する。

```text
Downside stress curve = P1 curve
```

---

## 9. 分位点カーブ・ストレスカーブの考え方

### 9.1 分位点カーブ

分位点カーブは、各 tenor の simulated rate distribution から作成する。

例えば、1年後における 5年金利の simulated distribution がある場合、その 99%点を 5年 tenor の P99 rate とする。

これを各 tenor について行い、各 tenor の P99 rate を連結したものが P99 curve である。

```text
P99 curve:
  1Y tenor  : 1Y rate distribution の 99%点
  2Y tenor  : 2Y rate distribution の 99%点
  5Y tenor  : 5Y rate distribution の 99%点
  10Y tenor : 10Y rate distribution の 99%点
```

### 9.2 注意点

各 tenor ごとの分位点を単純に連結した curve は、必ずしも同一 simulation path から得られた curve ではない。

したがって、P99 curve は以下の意味を持つ。

```text
各 tenor における個別分布の 99%点を連結した統計的ストレスカーブ
```

これは、金利カーブ全体の同時発生シナリオというより、tenor別の上方分位点をつないだストレス表現である。

同一 path に基づくストレスカーブを作る場合は、別途、path単位の損失額や代表指標で ranking し、その path の curve を抽出する必要がある。

### 9.3 実務利用上の整理

分位点カーブには、以下の2種類がある。

```text
1. Tenor-wise percentile curve
   各 tenor ごとに分位点を取り、連結した curve

2. Path-wise stress curve
   損失額や代表金利変化で path を順位付けし、特定 path の curve を抽出したもの
```

リスク管理資料で利用する際は、どちらの意味で作成した stress curve なのかを明記する。

---

## 10. Excel上の想定シート構成

初期実装では、以下のような sheet 構成を想定する。

```text
HW_Input
HW_VolSurface
HW_Calibration
HW_Simulation
HW_Output
```

### 10.1 `HW_Input`

モデル全体の設定を置く。

```text
Valuation Date
Mean Reversion a
Simulation Horizon
Number of Paths
Time Step dt
Random Seed
Use Antithetic
```

### 10.2 `HW_VolSurface`

ATM normal swaption volatility matrix を置く。

```text
Expiry × Tenor matrix
```

### 10.3 `HW_Calibration`

calibration quote grid と calibration result を置く。

```text
Quote Expiry
Quote Tenor
Market Vol
Model Vol
Error
```

### 10.4 `HW_Simulation`

simulation 実行条件を置く。

```text
Output Tenors
Output Curve Type
Number of Paths
Horizon
```

### 10.5 `HW_Output`

simulation 結果を出力する。

```text
Path別 future curve
Percentile curve
Stress curve
```

---

## 11. 実行イメージ

### 11.1 Calibration 実行

Excel/VBA からは、workflow module を通じて実行する。

概念的には、以下のような処理となる。

```vb
Dim cVolSurface As clsVolSurface
Dim cCalibrator As clsHWCalibrator
Dim sigma As Double

Set cVolSurface = HW_CreateVolSurfaceFromRanges( _
    in_ExpiryRange:=Range("B2:B6"), _
    in_TenorRange:=Range("C1:G1"), _
    in_VolRange:=Range("C2:G6"), _
    in_VolType:="NORMAL", _
    in_AllowFlatExtrapolation:=False)

Set cCalibrator = HW_CreateCalibratorFromRanges( _
    in_Curve:=cCurve, _
    in_VolSurface:=cVolSurface, _
    in_QuoteExpiryRange:=Range("J2:J10"), _
    in_QuoteTenorRange:=Range("K2:K10"))

sigma = cCalibrator.CalibrateSigmaFixedA( _
    in_a:=0.03, _
    in_SigmaLower:=0.000001, _
    in_SigmaUpper:=0.05)
```

### 11.2 Simulation 実行

概念的には、以下のような処理となる。

```vb
Dim cSimulator As clsHWSimulator

Set cSimulator = HW_CreateSimulator( _
    in_Curve:=cCurve, _
    in_a:=0.03, _
    in_sigma:=sigma, _
    in_HorizonYears:=1#)

cSimulator.RunSimulation in_NumPaths:=10000
```

実際の関数名・引数名は、実装の進捗に応じて更新する。

---

## 12. 検証観点

### 12.1 入力検証

以下を確認する。

- discount factor が正である
- zero rate / forward rate が異常値になっていない
- vol surface の行数・列数が expiry / tenor grid と一致している
- normal vol が負ではない
- `a` が負ではない
- `sigma` が負ではない
- simulation horizon が正である
- number of paths が正である

### 12.2 Calibration 検証

以下を確認する。

- calibration quote 数が想定どおりである
- market vol と model vol の差が極端に大きくない
- 特定 expiry / tenor に error が偏っていない
- 推定された `sigma` が探索範囲の上限・下限に張り付いていない
- `a` の候補を変えた場合の objective が比較可能である

### 12.3 Simulation 検証

以下を確認する。

- `x(t)` の平均が概ね 0 付近にある
- `x(t)` の分散が理論分散と大きく乖離していない
- path 数を増やすと分位点が安定する
- random seed を固定した場合に結果が再現する
- antithetic を使った場合に分布が不自然になっていない

### 12.4 Output 検証

以下を確認する。

- tenor ごとの rate distribution が極端に歪んでいない
- P50 curve が初期 forward curve と大きく乖離しすぎていない
- P95 / P99 curve がストレスカーブとして説明可能な形状になっている
- 長期 tenor の分位点が過度に不安定でない
- 出力単位が rate の絶対値か bp 変化か明確である

---

## 13. 今後の拡張

### 13.1 厳密な swaption pricing

将来的には、近似 model vol ではなく、Hull-White 1F の厳密な swaption pricing を実装する。

候補は以下である。

- bond option formula
- Jamshidian decomposition
- normal implied vol への逆算

これにより、`clsHWCalibrator` の objective function をより市場整合的にする。

### 13.2 `a` と `sigma` の同時最適化

初期実装では `a` を固定するが、将来的には `a` と `sigma` を同時に推定する。

```text
minimize Σ_i { modelVol_i(a, sigma) - marketVol_i }^2
```

ただし、1ファクター・定数パラメータでは、vol surface 全体への完全なフィットには限界がある。

そのため、目的に応じて以下を検討する。

- 特定 expiry / tenor を重視する weighted calibration
- short expiry を重視する calibration
- long tenor を重視する calibration
- リスク管理目的に合わせた calibration quote の選択

### 13.3 Smile / SABR との接続

将来的には、ATM vol だけでなく、strike方向の smile 情報も利用する。

候補は以下である。

- SABR normal vol surface
- strike別 swaption normal vol
- density / percentile 情報との接続

ただし、Hull-White 1F は基本的に Gaussian short-rate model であり、smile 全体を自然に表現するモデルではない。

そのため、smile 情報は以下のような使い方が考えられる。

- calibration quote の補正
- stress scenario の妥当性確認
- Hull-White MC 結果との比較
- 別モデルによる分布推定との比較

### 13.4 Path-wise stress curve

Tenor-wise percentile curve だけでなく、path-wise stress curve も作成する。

例：

```text
1. 各 path について portfolio loss を計算する
2. loss が大きい順に path を並べる
3. 99%損失に対応する path を選ぶ
4. その path の future curve を stress curve とする
```

この方法では、同一 path 上の curve を使うため、カーブ全体としての一貫性が高い。

### 13.5 複数 horizon 対応

初期実装では 1年後を中心にするが、将来的には複数 horizon に対応する。

```text
3か月後
6か月後
1年後
2年後
```

これにより、時間方向のストレス変化や、リスク量計測期間との整合性を確認できる。

---

## 14. 注意事項

本実装は、教育・プロトタイプ用途の VBA 実装である。

実務利用する場合は、以下を別途確認する必要がある。

- market data の品質
- OIS curve construction の妥当性
- volatility data の単位
- day count convention
- business day convention
- calibration quote の選択
- model risk
- parameter stability
- stress curve の説明可能性
- 監査・検証資料としての十分性

また、初期実装の Hull-White calibration は、厳密な swaption pricing ではなく近似的な sigma 推定である。

そのため、結果を利用する際は、以下のように位置づける。

```text
市場整合価格評価モデルではなく、
リスク管理上の将来金利カーブ分布・ストレスカーブ作成のための
初期プロトタイプである。
```

---

## 15. 完了条件

本 workflow の初期版は、以下を満たすことを完了条件とする。

```text
1. discount curve object が DF_T(T) と InstantaneousForward(T) を返せる
2. clsVolSurface が ATM normal swaption vol を返せる
3. clsHWCalibrator が fixed a に対して sigma を推定できる
4. clsHWSimulator が horizon 時点の future curve を path別に出力できる
5. path別 future curve から percentile curve を作成できる
6. P95 / P99 等の stress curve を出力できる
7. calibration report と simulation output が Excel 上で確認できる
```

この完了条件を満たした後、厳密 pricing、同時最適化、path-wise stress curve、複数 horizon 対応へ拡張する。
