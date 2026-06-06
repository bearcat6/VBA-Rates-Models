# ATM Swaption Volatility Class

## 1. 目的

`clsATMSwaptionVol` は、ATM Swaption Volatility Matrix を保持し、任意の `Expiry × Tenor` に対応する ATM ボラティリティを返すクラスである。

このクラスが扱う対象は、あくまで **ATM の swaption volatility** である。

旧クラス名 `clsSwaptionVol` は、swaption volatility 全般を扱うように見えるため、対象範囲が広すぎる。実際には ATM ボラティリティ行列のみを扱うため、クラス名を `clsATMSwaptionVol` とし、責務を明確化する。

将来的に、ストライク別ボラティリティ、スマイル、SABR、密度関数などを扱う場合は、別クラスとして追加する。

---

## 2. 対象ファイル

現在の実装ファイルは以下。

```text
src/classes/clsATMSwaptionVol.cls
```

旧ファイル名は以下。

```text
src/classes/clsSwaptionVol.cls
```

旧ファイル名は使用しない。

---

## 3. クラスの責務

`clsATMSwaptionVol` の責務は以下である。

- Excel Range から ATM Swaption Volatility Matrix を読み込む
- Expiry グリッドを保持する
- Tenor グリッドを保持する
- Volatility matrix を保持する
- 任意の `Expiry × Tenor` に対して ATM Vol を返す
- 必要に応じて範囲外を端点固定で外挿する

一方で、以下は本クラスの責務ではない。

- Swaption pricing
- SABR calibration
- Strike smile interpolation
- Density calculation
- Hull-White calibration
- Monte Carlo simulation

これらは、別クラスまたは別モジュールで扱う。

---

## 4. 入力マトリックス

本クラスは、1行目に Tenor、1列目に Expiry を持つ Excel Range を入力として想定する。

```text
       1Y    2Y    5Y    10Y
1M     vol   vol   vol   vol
3M     vol   vol   vol   vol
6M     vol   vol   vol   vol
1Y     vol   vol   vol   vol
```

例：

- 1行目、2列目以降：Tenor
- 1列目、2行目以降：Expiry
- 2行目・2列目以降：ATM Vol

Tenor / Expiry は、`1M`、`3M`、`6M`、`1Y`、`5Y` などの文字列表現を年数に変換して保持する。

---

## 5. 補間方針

補間は、単純に Vol を線形補間するのではなく、以下の方針で行う。

### Tenor 方向

Tenor 方向は、Variance を線形補間する。

```text
Variance = Vol^2
```

### Expiry 方向

Expiry 方向は、Total Variance を線形補間する。

```text
Total Variance = Vol^2 × Expiry
```

このため、補間の流れは以下となる。

```text
1. Tenor 方向に Vol^2 を線形補間する
2. Expiry 方向に Vol^2 × Expiry を線形補間する
3. 最後に Vol に戻す
```

---

## 6. 主な Public Method

### InitializeFromRange

```vb
Public Sub InitializeFromRange( _
  ByVal in_MatrixRange As Range, _
  Optional ByVal in_AllowFlatExtrapolation As Boolean = False)
```

ATM Volatility Matrix を Excel Range から読み込む。

Validation:

- Range が Nothing の場合はエラー
- Matrix が小さすぎる場合はエラー
- 空セルがある場合はエラー
- 数値でない Vol がある場合はエラー
- 負の Vol がある場合はエラー
- Expiry / Tenor が昇順でない場合はエラー

`in_AllowFlatExtrapolation` の意味は以下。

```text
False または省略：範囲外はエラー
True          ：範囲外は端点固定で外挿
```

---

### Vol

```vb
Public Function Vol( _
  ByVal in_ExpiryYears As Double, _
  ByVal in_TenorText As String) As Double
```

Expiry を年数、Tenor を文字列で指定し、ATM Vol を返す。

例：

```vb
vol = atmVol.Vol(1#, "10Y")
```

---

### VolByYears

```vb
Public Function VolByYears( _
  ByVal in_ExpiryYears As Double, _
  ByVal in_TenorYears As Double) As Double
```

Expiry と Tenor をいずれも年数で指定し、ATM Vol を返す。

例：

```vb
vol = atmVol.VolByYears(1#, 10#)
```

---

### ExpiryAt / TenorAt / GridVol

```vb
Public Function ExpiryAt(ByVal in_Index As Long) As Double
Public Function TenorAt(ByVal in_Index As Long) As Double
Public Function GridVol( _
  ByVal in_ExpiryIndex As Long, _
  ByVal in_TenorIndex As Long) As Double
```

読み込んだ元データのグリッドや Vol を確認するための関数である。

---

## 7. 命名ルール

ATM Swaption Volatility Matrix を扱う場合は、以下の名前を使用する。

```text
clsATMSwaptionVol
```

以下の旧名は使用しない。

```text
clsSwaptionVol
```

理由は、`clsSwaptionVol` という名前では、ATM だけでなく、ストライク別 volatility surface や smile-aware surface まで含むように見えるためである。

---

## 8. 将来拡張

将来的に、ATM 以外の swaption volatility を扱う場合は、別クラスを追加する。

想定例：

```text
clsSwaptionSmile
  Expiry × Tenor × Strike の smile を扱う

clsSABRSwaptionVol
  SABR parameter と implied volatility を扱う

clsSwaptionVolSurface
  ATM / smile / SABR などを統合的に扱う上位クラス
```

`clsATMSwaptionVol` は、今後も ATM Vol Matrix へのアクセスに責務を限定する。

---

## 9. Hull-White 1F との関係

Hull-White 1F の calibration では、市場の ATM Swaption Vol を参照する。

このとき、`clsHWCalibrator` は `clsATMSwaptionVol` から calibration 対象の ATM Vol を取得し、`clsHullWhite1F` のパラメータ `a` と `sigma` を調整する想定である。

```text
clsATMSwaptionVol
  market ATM swaption vol を返す

clsHWCalibrator
  market vol と model vol を比較し、a / sigma を決定する

clsHullWhite1F
  決定された a / sigma を使って金利モデルを表現する
```

このため、ATM Vol の入力・補間処理は `clsATMSwaptionVol` に閉じ込め、Hull-White 本体には持ち込まない。