# 標準正規乱数生成クラス設計メモ

## 1. 目的

この文書は、`src/classes/clsRandomNormal.cls` の役割と初期実装方針を整理するための設計メモである。

`clsRandomNormal.cls` は、Monte Carlo simulation で利用する標準正規乱数を生成するためのクラスである。

主な目的は、乱数生成処理を `clsHWSimulator` などのモデルクラスから分離することである。

## 2. 背景

Hull-White 1F の Monte Carlo simulation では、各時間ステップごとに標準正規乱数 `Z` が必要になる。

たとえば、shifted short-rate factor `x(t)` は以下の exact discretization で更新する。

```text
x(t+dt)
  = x(t) exp(-a dt)
    + sigma sqrt({1 - exp(-2a dt)} / (2a)) Z
```

ここで `Z` は標準正規乱数である。

初期の `clsHWSimulator.cls` では、Box-Muller 法による乱数生成をクラス内部に持っていた。

しかし、今後以下のような拡張を考えると、乱数生成は独立クラスに分けた方がよい。

```text
・複数のMonte Carloクラスで同じ乱数生成器を使う
・seedを明示して再現性を確保する
・antithetic variates を導入する
・乱数品質や生成方法を後から差し替える
・Sobol列などの準乱数に将来拡張する
```

このため、`clsRandomNormal.cls` を新規作成する。

## 3. 役割

`clsRandomNormal.cls` の役割は以下である。

```text
・Uniform(0,1) 乱数を生成する
・標準正規乱数 N(0,1) を生成する
・Box-Muller 法で標準正規乱数を生成する
・1回のBox-Mullerで生成した2個目の標準正規乱数をcacheする
・seed指定により再現性を確保する
・antithetic variates 用に z, -z のペアを返す
```

一方、以下はこのクラスの責務ではない。

```text
・Hull-White モデル式の計算
・金利カーブの生成
・キャリブレーション
・Excelシートへの入出力
・商品評価
```

## 4. 実装方式

標準正規乱数は Box-Muller 法で生成する。

2つの独立な一様乱数 `u1`, `u2` から、以下で2つの独立な標準正規乱数を得る。

```text
z1 = sqrt(-2 ln u1) cos(2πu2)
z2 = sqrt(-2 ln u1) sin(2πu2)
```

`clsRandomNormal` では、1回のBox-Mullerで得られた `z1`, `z2` のうち、まず `z1` を返し、`z2` は内部にcacheする。

次回 `Normal()` が呼ばれたとき、cacheがあれば新たな一様乱数を使わずに `z2` を返す。

これにより、乱数生成の効率を少し改善できる。

## 5. Seed と再現性

`Init` メソッドでは、seedを指定できる。

```vb
Public Sub Init( _
  Optional ByVal in_Seed As Long = 0, _
  Optional ByVal in_UseAntithetic As Boolean = False _
)
```

`in_Seed > 0` の場合、VBA標準の `Rnd` / `Randomize` にseedを設定する。

```vb
Rnd -CDbl(in_Seed)
Randomize in_Seed
```

これにより、同じseedを指定すれば、同じ乱数列を再現できる。

`in_Seed <= 0` の場合は、通常の `Randomize` を使う。

## 6. Antithetic variates

`clsRandomNormal` は、簡易的な antithetic variates をサポートする。

`UseAntithetic = True` の場合、`Normal()` は以下の順序で乱数を返す。

```text
z, -z, z, -z, ...
```

これは、Monte Carlo simulation の分散削減に使える。

ただし、初期実装では単純なペア生成であり、path単位での完全な antithetic path 管理までは行っていない。

path単位のantithetic simulationを行う場合は、`clsHWSimulator` 側でpathの組み方を設計する必要がある。

## 7. 主な Public メソッド

`clsRandomNormal.cls` の主な Public メソッドは以下である。

```vb
Public Sub Init( _
  Optional ByVal in_Seed As Long = 0, _
  Optional ByVal in_UseAntithetic As Boolean = False _
)

Public Function Normal() As Double

Public Function IndependentNormal() As Double

Public Sub NormalPair( _
  ByRef out_Z1 As Double, _
  ByRef out_Z2 As Double _
)

Public Function Uniform01() As Double

Public Sub ClearCache()

Public Sub Reset()

Public Function Summary() As String
```

## 8. Public Properties

主な Property は以下である。

```vb
Public Property Get IsInitialized() As Boolean

Public Property Get HasSeed() As Boolean

Public Property Get Seed() As Long

Public Property Get UseAntithetic() As Boolean

Public Property Let UseAntithetic(ByVal in_UseAntithetic As Boolean)
```

`UseAntithetic` を変更した場合、内部cacheはクリアされる。

## 9. Normal と IndependentNormal の違い

`Normal()` は、`UseAntithetic` の設定を考慮する。

```text
UseAntithetic = False
  → 独立な標準正規乱数を返す

UseAntithetic = True
  → z, -z の順で返す
```

一方、`IndependentNormal()` は、`UseAntithetic` の設定を無視して、独立な標準正規乱数を返す。

そのため、通常のMonte Carloでは `Normal()` を使い、antitheticを無視したい特殊用途では `IndependentNormal()` を使う。

## 10. 使い方イメージ

通常の標準正規乱数を生成する例は以下である。

```vb
Dim rng As clsRandomNormal
Dim z As Double

Set rng = New clsRandomNormal
rng.Init 12345, False

z = rng.Normal()
```

antitheticを使う例は以下である。

```vb
Dim rng As clsRandomNormal
Dim z1 As Double
Dim z2 As Double

Set rng = New clsRandomNormal
rng.Init 12345, True

z1 = rng.Normal()  ' z
z2 = rng.Normal()  ' -z
```

乱数ペアを直接取得する例は以下である。

```vb
Dim z1 As Double
Dim z2 As Double

rng.NormalPair z1, z2
```

## 11. clsHWSimulator との関係

`clsHWSimulator.cls` の初期実装では、内部に `RandomNormal()` を持っている。

今後は、以下のように `clsRandomNormal` を利用する形へリファクタリングする。

```text
現状：
  clsHWSimulator
    ・内部に RandomNormal() を持つ

将来：
  clsHWSimulator
    ・clsRandomNormal をメンバーとして持つ
    ・rng.Normal() を呼ぶ
```

想定されるメンバーは以下である。

```vb
Private mRng As clsRandomNormal
```

初期化時に乱数生成器を渡す設計も考えられる。

```vb
Public Sub SetRandomGenerator(ByVal in_Rng As clsRandomNormal)
```

または、`Simulate` の中で seed を指定して内部生成器を初期化する方法も考えられる。

## 12. 注意点と限界

現在の `clsRandomNormal.cls` は、VBA標準の `Rnd` / `Randomize` を乱数源としている。

したがって、以下の点には注意する。

```text
・金融MC専用の高品質乱数生成器ではない
・Mersenne Twister ではない
・Sobol列などの準乱数ではない
・乱数品質の統計検定は行っていない
・大規模・高精度な評価用途では、乱数生成器の改善余地がある
```

ただし、初期のHull-White simulation、分位点カーブ作成、ストレスカーブ試作には十分扱いやすい。

## 13. 今後の拡張候補

今後の拡張候補は以下である。

```text
・Mersenne Twister の実装
・Sobol sequence の実装
・Cholesky分解と組み合わせた多次元正規乱数
・antithetic path の明示的管理
・moment matching
・乱数列の保存・再利用
・標準正規分布の逆関数による生成方式
```

Hull-White 1F では1次元正規乱数で足りるが、将来の多因子モデルやPCAベースの金利カーブシミュレーションでは、多次元正規乱数が必要になる可能性がある。

## 14. まとめ

`clsRandomNormal.cls` は、Monte Carlo simulation 用の標準正規乱数生成クラスである。

初期実装では、以下を重視する。

```text
・乱数生成をモデルクラスから分離する
・Box-Muller 法で標準正規乱数を生成する
・seed指定により再現性を確保する
・cacheによりBox-Mullerの2個目の乱数を有効活用する
・antithetic variates の初期対応を行う
```

今後、`clsHWSimulator.cls` はこのクラスを利用する形にリファクタリングする。
