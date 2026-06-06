# Hull-White 1F カーブ・インターフェース

## 1. 目的

本ドキュメントは、VBA-Rates-Models における Hull-White 1F モデル実装で必要となる、ディスカウントカーブ側のインターフェースを定義するものである。

本リポジトリには、すでに日付ベースの OIS カーブクラスが存在する。

- `clsOISStepForwardCurve`
- `clsOISZeroLinearCurve`

これらのカーブクラスは、主に以下のような日付ベースのメソッドを提供する。

```vb
Public Function DF(ByVal in_TargetDate As Date) As Double

Public Function ZeroRateCont(ByVal in_TargetDate As Date) As Double

Public Function ForwardRate(ByVal in_StartDate As Date, _
                            ByVal in_EndDate As Date) As Double
```

一方、Hull-White 1F モデルは、評価日からの年数である `t` や `T` を使って自然に記述される。

そのため、Hull-White 1F モデルで利用する具体的なカーブクラスは、既存の日付ベースのメソッドに加えて、年数ベースの共通インターフェースも提供する必要がある。

---

## 2. 背景

Hull-White 1F モデルは、一般に以下の形で表される。

```text
dr(t) = { theta(t) - a r(t) } dt + sigma dW(t)
```

本リポジトリでは、実装上、以下の shifted short-rate representation を利用する。

```text
r(t) = phi(t) + x(t)

dx(t) = -a x(t) dt + sigma dW(t)
```

ここで、各要素の意味は以下のとおりである。

- 初期ディスカウントカーブは、現在の市場カーブを再現する。
- `a` は平均回帰速度を表す。
- `sigma` は短期金利ファクターのボラティリティを表す。
- `theta(t)` は直接の入力パラメータとして保持しない。
- `phi(t)` は初期ディスカウントカーブと整合するように決まるシフト項である。
- モデル計算では、初期ディスカウントファクターと初期瞬間フォワードレートが必要となる。

したがって、Hull-White 1F モデルは、カーブの具体的な構築方法を知らなくても、以下のメソッドを呼び出せる必要がある。

```vb
DF_T(T)
InstantaneousForward(T)
```

---

## 3. 設計方針

Hull-White 1F の実装は、特定のカーブ構築方法に依存させない。

例えば、モデル側は、利用するカーブが以下のどれであるかを意識しない。

- 瞬間フォワードレートを階段状に保持するカーブ
- 連続複利ゼロレートを直線補間するカーブ
- 将来追加される別方式のカーブ

代わりに、各具体カーブクラスが、同じ名前の年数ベース Public メソッドを提供する。

Hull-White 関連クラスは、カーブを `Object` として受け取り、同名のメソッドを呼び出す。

これは、VBA が一般的なオブジェクト指向言語のようなクラス継承を持たないことを踏まえた、実務的な設計である。

---

## 4. 必要となる年数ベース・インターフェース

Hull-White 1F で利用する具体カーブクラスは、以下の Public メソッドを実装する。

```vb
Public Function YearFracFromValDate(ByVal targetDate As Date) As Double

Public Function DateFromT(ByVal T As Double) As Date

Public Function DF_T(ByVal T As Double) As Double

Public Function ZeroRate_T(ByVal T As Double) As Double

Public Function ForwardRate_T(ByVal T1 As Double, _
                              ByVal T2 As Double) As Double

Public Function InstantaneousForward(ByVal T As Double, _
                                     Optional ByVal eps As Double = 0.0001) As Double
```

---

## 5. 各メソッドの仕様

### 5.1 `YearFracFromValDate`

```vb
Public Function YearFracFromValDate(ByVal targetDate As Date) As Double
```

評価日から `targetDate` までの年数を返す。

現時点の実装方針は以下のとおりとする。

```vb
YearFracFromValDate = YearFracAct365F(mValuationDate, targetDate)
```

期待する挙動は以下のとおりである。

- `targetDate` が評価日より前の場合はエラーとする。
- `targetDate` が評価日と同じ場合は `0` を返す。
- 既存カーブ実装で利用している day count と整合させる。

---

### 5.2 `DateFromT`

```vb
Public Function DateFromT(ByVal T As Double) As Date
```

評価日からの年数 `T` を日付に変換する。

初期実装では、以下の簡易実装を許容する。

```vb
DateFromT = DateAdd("d", CLng(Round(T * 365#, 0)), mValuationDate)
```

これは Hull-White 1F の初期実装としては十分である。

将来の改善候補は以下のとおりである。

- `YearFracAct365F` との整合性を確認する。
- day count convention と日付変換の一貫性を高める。
- 利用目的に応じて、営業日調整を行うべきか検討する。

期待する挙動は以下のとおりである。

- `T < 0` の場合はエラーとする。
- `T = 0` の場合は評価日を返す。

---

### 5.3 `DF_T`

```vb
Public Function DF_T(ByVal T As Double) As Double
```

評価日から年数 `T` 先のディスカウントファクター `P(0,T)` を返す。

実装方針は以下のとおりである。

```vb
DF_T = Me.DF(DateFromT(T))
```

日付ベースの既存メソッド名が `df` のクラスでは、以下のように既存メソッドを呼び出す。

```vb
DF_T = Me.df(DateFromT(T))
```

期待する挙動は以下のとおりである。

- `T < 0` の場合はエラーとする。
- `T = 0` の場合は `1` を返す。
- それ以外の場合は、`T` を日付に変換し、既存の日付ベースのディスカウントファクター計算を呼び出す。

---

### 5.4 `ZeroRate_T`

```vb
Public Function ZeroRate_T(ByVal T As Double) As Double
```

評価日から年数 `T` 先までの連続複利ゼロレートを返す。

```text
z(0,T) = -ln(P(0,T)) / T
```

実装方針は以下のとおりである。

```vb
ZeroRate_T = ZeroRateFromDF(DF_T(T), T)
```

共通関数 `ZeroRateFromDF` は、`mdl_CurveMath.bas` に定義する。

期待する挙動は以下のとおりである。

- `T = 0` の場合は `0` を返す。
- ディスカウントファクターが正でない場合は、`mdl_CurveMath.ZeroRateFromDF` 側でエラーとする。

---

### 5.5 `ForwardRate_T`

```vb
Public Function ForwardRate_T(ByVal T1 As Double, _
                              ByVal T2 As Double) As Double
```

年数 `T1` から `T2` までの単利フォワードレートを返す。

```text
F(0;T1,T2) = { P(0,T1) / P(0,T2) - 1 } / (T2 - T1)
```

実装方針は以下のとおりである。

```vb
ForwardRate_T = ForwardRateFromDFs(DF_T(T1), DF_T(T2), T1, T2)
```

共通関数 `ForwardRateFromDFs` は、`mdl_CurveMath.bas` に定義する。

期待する挙動は以下のとおりである。

- `T2 <= T1` の場合はエラーとする。
- いずれかのディスカウントファクターが正でない場合は、`mdl_CurveMath.ForwardRateFromDFs` 側でエラーとする。

---

### 5.6 `InstantaneousForward`

```vb
Public Function InstantaneousForward(ByVal T As Double, _
                                     Optional ByVal eps As Double = 0.0001) As Double
```

初期瞬間フォワードレートを返す。

```text
f(0,T) = - d ln P(0,T) / dT
```

実装方針としては、`DF_T` を使った有限差分近似を利用する。

`T` がゼロに近い場合は、前進差分を使う。

```text
f(0,T) ≈ - { ln P(0,T+eps) - ln P(0,0) } / (T+eps)
```

通常の正の `T` では、中心差分を使う。

```text
f(0,T) ≈ - { ln P(0,T+eps) - ln P(0,T-eps) } / (2 eps)
```

期待する挙動は以下のとおりである。

- `T < 0` の場合はエラーとする。
- `eps <= 0` の場合はエラーとする。
- 計算に利用するディスカウントファクターが正でない場合はエラーとする。

---

## 6. `clsOISStepForwardCurve` での実装方針

`clsOISStepForwardCurve` は、瞬間フォワードレートを区間ごとに階段状に保持するカーブである。

このクラスでは、`DF_T(T)` は `T` を日付に変換し、既存の日付ベースのディスカウントファクター計算を呼び出す。

```vb
DF_T(T)
```

は、以下の既存メソッドを利用する。

```vb
df(Date)
```

`ZeroRate_T(T)` は、以下を呼び出す。

```vb
ZeroRateFromDF(DF_T(T), T)
```

`ForwardRate_T(T1, T2)` は、以下を呼び出す。

```vb
ForwardRateFromDFs(DF_T(T1), DF_T(T2), T1, T2)
```

`InstantaneousForward(T)` は、`DF_T` を使った有限差分で計算する。

このカーブは瞬間フォワードレートを階段状に保持するため、`InstantaneousForward(T)` はセグメント境界付近で不連続に見える場合がある。

これはエラーではなく、カーブ構築方法を反映した自然な挙動である。

---

## 7. `clsOISZeroLinearCurve` での実装方針

`clsOISZeroLinearCurve` は、評価日からの年数に対して連続複利ゼロレートを直線補間するカーブである。

このクラスでは、`DF_T(T)` は `T` を日付に変換し、既存の日付ベースのディスカウントファクター計算を呼び出す。

```vb
DF_T(T)
```

は、以下の既存メソッドを利用する。

```vb
DF(Date)
```

その他のメソッドは、`clsOISStepForwardCurve` と同じ年数ベース・インターフェースに従う。

ただし、`InstantaneousForward(T)` の形状は、`clsOISStepForwardCurve` と異なる可能性がある。

これは、ゼロレート直線補間カーブと、瞬間フォワード階段状カーブでは、暗黙に含まれる瞬間フォワードカーブの形状が異なるためである。

---

## 8. `clsDiscountCurve` との関係

`clsDiscountCurve` は、現時点では `Implements` を使う日付ベースのインターフェースとして利用する。

`clsDiscountCurve` が定義する主なメソッドは以下である。

```vb
Public Function DF(ByVal in_TargetDate As Date) As Double

Public Function ZeroRateCont(ByVal in_TargetDate As Date) As Double

Public Function ForwardRate(ByVal in_StartDate As Date, _
                            ByVal in_EndDate As Date) As Double
```

Hull-White 1F 用の年数ベースメソッドは、現時点では `clsDiscountCurve` に追加しない。

理由は以下のとおりである。

- `clsDiscountCurve` に年数ベースメソッドを追加すると、`Implements` しているすべてのクラスで対応する private interface method を実装する必要がある。
- Hull-White 1F モデル側は、カーブを `Object` として受け取ればよい。
- 必要な年数ベースメソッドは、共通の命名規約に基づく convention-based interface として扱える。

したがって、`clsDiscountCurve` は既存の日付ベース・インターフェースのままとする。

Hull-White 1F 用のカーブ・インターフェースは、本ドキュメントで定義する convention-based interface とする。

---

## 9. Hull-White 関連クラスからの利用方法

Hull-White 関連クラスでは、カーブオブジェクトを以下のように `Object` として保持する。

```vb
Private mCurve As Object
```

そして、以下のような年数ベースメソッドを呼び出す。

```vb
mCurve.DF_T(T)
mCurve.InstantaneousForward(T)
```

これにより、Hull-White 1F の実装は、具体的なカーブ構築方法から独立する。

初期化メソッドの例は以下のとおりである。

```vb
Public Sub Init(ByVal in_Curve As Object, _
                ByVal in_a As Double, _
                ByVal in_sigma As Double)

  Set mCurve = in_Curve
  mA = in_a
  mSigma = in_sigma

End Sub
```

---

## 10. 責務分離

Hull-White 1F のモデルクラスは、Excel シートを直接読み込まない。

推奨する責務分離は以下のとおりである。

- カーブクラス  
  ディスカウントファクター、ゼロレート、フォワードレート、瞬間フォワードレートを返す。

- `mdl_CurveMath.bas`  
  カーブ関連の共通数理関数を提供する。

- `clsHullWhite1F`  
  Hull-White 1F のパラメータとモデル数式を保持する。

- `clsHWCalibrator`  
  ボラティリティデータから `a` と `sigma` をキャリブレーションする。初期実装では、`a` を外部固定し、`sigma` を推定する。

- `clsHWSimulator`  
  Monte Carlo path と将来金利カーブを生成する。

- `clsRandomNormal`  
  Monte Carlo simulation で利用する標準正規乱数を生成する。

- `clsVolSurface`  
  Hull-White 1F のキャリブレーション等で利用するボラティリティ・サーフェスを保持する。

- `mdl_HullWhiteMath.bas`  
  Hull-White 1F に関する共通数理関数を提供する。

- `mdl_HullWhiteWorkFlow.bas`  
  Excel/VBA から、ボラティリティサーフェス作成、キャリブレーション、シミュレーション、出力までの全体処理をつなぐ。

Excel 入出力専用の処理を追加する場合は、将来的に `mdl_ExcelIO.bas` のような標準モジュールへ分離する。

診断・検証用の処理を追加する場合は、将来的に `mdl_Diagnostics.bas` のような標準モジュールへ分離する。

---

## 11. 現在の構成

現時点で実装済み、または本設計で想定する主要コンポーネントは以下のとおりである。

```text
src/
  classes/
    clsDiscountCurve.cls
    clsOISStepForwardCurve.cls
    clsOISZeroLinearCurve.cls
    clsVolSurface.cls
    clsHullWhite1F.cls
    clsHWCalibrator.cls
    clsHWSimulator.cls
    clsRandomNormal.cls

  modules/
    mdl_CurveMath.bas
    mdl_HullWhiteMath.bas
    mdl_HullWhiteWorkFlow.bas
```

将来的に必要に応じて、以下を追加する。

```text
src/
  modules/
    mdl_ExcelIO.bas
    mdl_Diagnostics.bas
```

---

## 12. 現在の実装ステップ

現在の実装ステップは以下のとおりである。

1. `mdl_CurveMath.bas` をカーブ共通数理モジュールとして維持する。
2. `clsOISStepForwardCurve` に年数ベース・インターフェースを追加する。
3. `clsOISZeroLinearCurve` に同じ年数ベース・インターフェースを追加する。
4. `clsDiscountCurve` は既存の日付ベース・インターフェースのままとする。
5. `clsHullWhite1F` は、年数ベース・カーブ・インターフェースを利用して実装する。
6. `clsHWCalibrator` は、`clsVolSurface` とカーブオブジェクトを利用して、初期実装では `a` 固定・`sigma` 推定を行う。
7. `clsHWSimulator` は、キャリブレーション済みの `a` と `sigma` を使い、将来時点の金利カーブを Monte Carlo simulation で生成する。

---

## 13. 注意点

本インターフェースは、Hull-White 1F のモデル実装を特定のカーブ構築方法から分離するためのものである。

そのため、各カーブクラスは、日付ベースの既存メソッドを維持しつつ、Hull-White 1F で必要となる年数ベースメソッドを追加する。

また、`InstantaneousForward(T)` は有限差分で近似するため、補間方法やカーブ構築方法によって滑らかさが異なる。

特に、`clsOISStepForwardCurve` のように瞬間フォワードを階段状に保持するカーブでは、セグメント境界付近で不連続が生じることがある。

これは実装エラーではなく、カーブ構築方法の性質である。

Hull-White 1F の厳密な市場整合性や swaption pricing の高度化は、今後の拡張課題とする。
