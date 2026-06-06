# ボラティリティ・サーフェス設計メモ

## 1. 目的

この文書は、`clsVolSurface.cls` と既存の `clsSwaptionVol.cls` の役割分担を整理するための設計メモである。

本リポジトリには、すでに `clsSwaptionVol.cls` が存在する。
このクラスは、ATMスワップション・ボラティリティのマトリックスを保持し、任意の `Expiry × Tenor` に対応するATMボラティリティを返すためのクラスである。

一方、Hull-White 1F モデルのキャリブレーションでは、モデル側から利用しやすい、より汎用的なボラティリティ・サーフェスが必要になる。

その役割を担うクラスとして、`clsVolSurface.cls` を新たに作成する。

## 2. 既存クラス：clsSwaptionVol

`clsSwaptionVol.cls` は、ATMスワップション・ボラティリティ・マトリックスを扱う既存クラスである。

主な役割は以下のとおりである。

```text
Expiry × Tenor の ATM swaption vol matrix を保持し、
任意の Expiry × Tenor の ATM vol を返す。
```

主な用途は以下である。

```text
・CMS convexity adjustment
・CMS spread swap 等で使う ATM swaption vol の取得
```

入力は、以下のようなExcel上のマトリックスを前提としている。

```text
        1Y    2Y    5Y    10Y
1M      vol   vol   vol   vol
3M      vol   vol   vol   vol
6M      vol   vol   vol   vol
1Y      vol   vol   vol   vol
```

補間方針は以下である。

```text
・Tenor方向  ：variance = vol^2 で線形補間
・Expiry方向 ：total variance = vol^2 × expiry で線形補間
```

したがって、`clsSwaptionVol.cls` は有用な既存クラスであり、削除する必要はない。

ただし、役割は「ATMスワップション・ボラティリティの取得」に寄っており、Hull-White 1F のモデル入力として使うには、やや具体的すぎる。

## 3. 新規クラス：clsVolSurface

`clsVolSurface.cls` は、より汎用的なボラティリティ・サーフェス・クラスとして作成する。

主な役割は以下である。

```text
Hull-White 1F のキャリブレーションや、将来の金利モデル拡張で利用する
市場ボラティリティの入力インターフェースを提供する。
```

`clsVolSurface.cls` は、特定の商品評価、特定のExcelレンジ、特定のシート構造に依存しない。

Hull-White のキャリブレーターは、Excelシートを直接読みに行くのではなく、`clsVolSurface` オブジェクトを受け取り、必要なボラティリティを問い合わせる。

例：

```vb
marketVol = volSurface.NormalVol(expiryYears, tenorYears)
```

または、

```vb
marketVol = volSurface.VolByYears(expiryYears, tenorYears)
```

## 4. clsSwaptionVol と clsVolSurface の違い

### 4.1 clsSwaptionVol

`clsSwaptionVol` は、ATMスワップション・ボラティリティ・マトリックスを扱う具体クラスである。

以下のような場合に適している。

```text
・入力がATMスワップション・ボラティリティのマトリックスである
・ボラティリティ種別が既に明確である
・Excel Range から直接読み込んでもよい
・主な利用者が CMS 関連の評価ユーティリティである
```

このクラスが答える問いは、例えば以下である。

```text
1Y × 10Y の補間済みATMスワップション・ボラティリティはいくらか。
```

### 4.2 clsVolSurface

`clsVolSurface` は、モデル入力として利用する汎用的なボラティリティ・サーフェスである。

以下のような場合に適している。

```text
・モデルやキャリブレーターをExcelシート構造に依存させたくない
・ボラティリティ種別を明示的に管理したい
・将来、ATMだけでなく、ストライク別・マネーネス別のボラティリティに拡張したい
・Hull-White、SABR、その他のモデルから共通的に参照したい
```

このクラスが答える問いは、例えば以下である。

```text
このモデルが、指定されたExpiryとTenorに対して使うべき市場ボラティリティはいくらか。
```

## 5. 推奨する設計方針

当面は、`clsSwaptionVol.cls` と `clsVolSurface.cls` の両方を残す。

役割は以下のように分ける。

```text
clsSwaptionVol
  = 既存のATMスワップション・ボラティリティ・マトリックス用クラス

clsVolSurface
  = Hull-White等のモデル入力として使う汎用ボラティリティ・サーフェス
```

初期実装では、`clsVolSurface` の内部補間ロジックは、`clsSwaptionVol` とかなり近いものになってよい。

ただし、設計上の依存関係は以下のようにする。

```text
CMS関連ユーティリティ
  → clsSwaptionVol を使ってもよい

Hull-Whiteキャリブレーション
  → clsVolSurface を使う
```

## 6. clsSwaptionVol を Hull-White で直接使わない理由

`clsSwaptionVol` をそのまま Hull-White のキャリブレーションに使うことも技術的には可能である。

しかし、初期段階から `clsVolSurface` を作った方がよい。

理由は以下である。

### 6.1 クラス名が商品寄りである

`clsSwaptionVol` は、スワップション専用の補助クラスに見える。

Hull-White のキャリブレーションでは、最初はスワップション・ボラティリティを使うとしても、モデル側から見ると「市場ボラティリティの入力」である。

したがって、モデル側の入力としては `clsVolSurface` の方が自然である。

### 6.2 ATM専用である

`clsSwaptionVol` は、ATMスワップション・ボラティリティのマトリックスを前提としている。

Hull-White の初期実装ではATMボラティリティで十分だが、将来的には以下のような拡張が考えられる。

```text
・ストライク別ボラティリティ
・マネーネス別ボラティリティ
・SABRで平滑化したボラティリティ
・normal vol / lognormal vol / shifted lognormal vol の切替
```

これらは、汎用的な `clsVolSurface` の方が受け止めやすい。

### 6.3 Excel Range 依存を避けたい

`clsSwaptionVol` は、Excel Range からの初期化を前提としている。

一方、Hull-White のモデルクラスやキャリブレーションクラスには、Excelシート読込を混ぜない方針である。

Excelからの入力は `modExcelIO` に分離し、モデル側はクラスオブジェクトを受け取るだけにする。

### 6.4 ボラティリティ種別を明示したい

Hull-White 1F のJPY金利モデルでは、まず normal volatility を前提にするのが自然である。

そのため、`clsVolSurface` では、少なくとも以下の情報を持たせる。

```text
VolType = "NORMAL"
```

これにより、normal vol と lognormal vol を誤って混ぜるリスクを下げられる。

## 7. clsVolSurface の初期スコープ

最初の `clsVolSurface.cls` は、あまり大きくしない。

初期スコープは以下でよい。

```text
・Expiry grid を年数で保持する
・Tenor grid を年数で保持する
・Vol matrix を保持する
・VolType を保持する
・当初の VolType は "NORMAL" を想定する
・必要に応じて端点固定外挿を許容する
・Tenor方向は variance = vol^2 で補間する
・Expiry方向は total variance = vol^2 × expiry で補間する
・指定された Expiry × Tenor の補間済みvolを返す
```

想定する主なPublicメソッドは以下である。

```vb
Public Sub Init( _
    ByVal in_ExpiryYears As Variant, _
    ByVal in_TenorYears As Variant, _
    ByVal in_Vols As Variant, _
    Optional ByVal in_VolType As String = "NORMAL", _
    Optional ByVal in_AllowFlatExtrapolation As Boolean = False _
)

Public Function VolByYears( _
    ByVal in_ExpiryYears As Double, _
    ByVal in_TenorYears As Double _
) As Double

Public Function NormalVol( _
    ByVal in_ExpiryYears As Double, _
    ByVal in_TenorYears As Double _
) As Double

Public Property Get VolType() As String

Public Property Get IsInitialized() As Boolean
```

`clsVolSurface` は、Excel Range を直接受け取らない。

Excelシートからの読み込みは、別途 `modExcelIO` が担当する。

## 8. ボラティリティ種別

Hull-White 1F の初期実装では、normal volatility を前提とする。

推奨する表現は以下である。

```text
VolType = "NORMAL"
```

値は年率換算された normal volatility として扱う。

例：

```text
0.0050 = 50bp normal vol
```

`clsVolSurface` では、normal vol と lognormal vol を暗黙に混在させない。

たとえば `NormalVol` メソッドでは、`VolType` が `"NORMAL"` であることを確認する。

イメージは以下である。

```vb
Public Function NormalVol(ByVal in_ExpiryYears As Double, _
                          ByVal in_TenorYears As Double) As Double

    If UCase$(mVolType) <> "NORMAL" Then
        Err.Raise vbObjectError + ..., _
                  "clsVolSurface.NormalVol", _
                  "VolType must be NORMAL."
    End If

    NormalVol = VolByYears(in_ExpiryYears, in_TenorYears)

End Function
```

## 9. Hull-White キャリブレーションとの関係

`clsHWCalibrator` は、Excel Range ではなく、`clsVolSurface` を受け取る。

想定する依存関係は以下である。

```text
clsHWCalibrator
  uses:
    ・discount curve object
    ・clsVolSurface
    ・swaption pricer または calibration objective
```

キャリブレーターは、以下のように市場ボラティリティを取得する。

```vb
marketVol = volSurface.NormalVol(expiryYears, tenorYears)
```

キャリブレーターの中で、以下のようなExcelアクセスは行わない。

```vb
Worksheets(...)
Range(...)
```

これにより、モデル層とExcel入出力層を分離できる。

## 10. modExcelIO との関係

`modExcelIO` は、Excelシートからボラティリティ・データを読み込む責務を持つ。

想定する流れは以下である。

```text
Excel sheet
  → modExcelIO が expiry labels, tenor labels, vol matrix を読み込む
  → modExcelIO が clsVolSurface を生成する
  → clsHWCalibrator が clsVolSurface を受け取る
```

この構造にすると、`clsVolSurface` はExcelに依存せず、テストしやすくなる。

## 11. 推奨する移行方針

### Step 1

`clsSwaptionVol.cls` はそのまま残す。

既存のCMS関連計算で利用できるため、急いで置き換えない。

### Step 2

`clsVolSurface.cls` を新規作成する。

初期実装では、`clsSwaptionVol.cls` の補間ロジックを参考にしてよい。

ただし、Excel Range から直接読むのではなく、配列またはVariantを受け取る。

### Step 3

Hull-White キャリブレーションでは `clsVolSurface` を使う。

### Step 4

将来的に必要であれば、`clsSwaptionVol` を `clsVolSurface` の薄いラッパーにする。

例えば、将来的には以下のような構造もあり得る。

```text
clsSwaptionVol
  → CMS向けの名前・入力形式を持つラッパー
  → 内部的には clsVolSurface を使う
```

ただし、このリファクタリングは初期段階では不要である。

## 12. まとめ

`clsSwaptionVol.cls` と `clsVolSurface.cls` は、似ているが役割を分ける。

```text
clsSwaptionVol.cls
  ・商品寄り / 補助クラス
  ・ATM swaption vol matrix 用
  ・CMS関連ユーティリティ向け
  ・Excel Range から直接初期化してもよい

clsVolSurface.cls
  ・モデル入力用の汎用クラス
  ・Hull-White calibration 向け
  ・将来のモデル拡張にも使う
  ・Excel Range には直接依存させない
  ・VolType を明示する
```

したがって、`clsVolSurface.cls` は `clsSwaptionVol.cls` を置き換えるものではなく、Hull-White 1F 実装に向けた新しいモデル入力クラスとして作成する。
