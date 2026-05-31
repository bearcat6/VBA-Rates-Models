# Design

## 1. 目的

本ドキュメントは、VBA-Rates-Models における金利モデル関連コードの設計方針を整理するものである。

当面の目的は、JPY TONA/OIS ベースのディスカウントカーブを構築し、将来的に TONA 変動金利にキャップが付いた商品の評価や、キャップボラティリティ期間構造の構築へ拡張できる設計とすることである。

設計上の詳細な前提は、`docs/assumptions.md` に従う。

---

## 2. 全体構成

本プロジェクトは、以下の構成とする。

```text
VBA-Rates-Models
├─ docs
│  ├─ assumptions.md
│  └─ design.md
├─ src
│  ├─ classes
│  │  ├─ clsHolidayCalendar.cls
│  │  ├─ clsDiscountCurve.cls
│  │  ├─ clsOISStepForwardCurve.cls
│  │  ├─ clsOISZeroLinearCurve.cls
│  │  ├─ clsOISswap.cls
│  │  ├─ clsCapVolBootstrapper
│  │  ├─ clsCapVolTermStructure
│  │  └─ clsSwaptionVol
│  └─ modules
│     ├─ mdl_Common.bas
│     ├─ mdl_DayCount.bas
│     ├─ mdl_BusinessDay.bas
│     ├─ mdl_DiscountCurveFactory.bas
│     ├─ mdl_CapFormula.bas
│     └─ mdl_SwaptionVol.bas
└─ README.md
```

---

## 3. 設計方針

### 3.1 責務を分離する

日付処理、営業日判定、ディスカウントカーブ、ボラティリティ処理は、それぞれ責務を分離する。

休日データは `clsHolidayCalendar` が保持し、土日を含む営業日判定、営業日加算、Business Day Convention に基づく日付調整は `mdl_BusinessDay` が行う。

### 3.2 Excel/VBA で扱いやすい設計にする

本プロジェクトは Excel/VBA 上での利用を前提とする。

そのため、過度に抽象化しすぎず、Excel関数として呼び出す場合にも扱いやすい設計とする。

### 3.3 将来拡張を前提にする

当初は OIS Step Forward Curve を実装対象とする。

ただし、将来的に以下のような拡張が可能な構成とする。

- ゼロレート直線補間型のディスカウントカーブ
- キャップボラティリティ期間構造
- TONA変動金利にキャップが付いた商品の評価
- スワップションボラティリティサーフェス
- CMS関連商品の評価

---

## 4. 命名規則

### 4.1 クラス名

クラス名は `cls` で始める。

例：

```text
clsHolidayCalendar
clsDiscountCurve
clsOISStepForwardCurve
```

### 4.2 モジュール名

標準モジュール名は `mdl_` で始める。

例：

```text
mdl_BusinessDay
mdl_DayCount
mdl_DiscountCurveFactory
```

### 4.3 関数引数名

関数の引数には `in_` を付ける。

例：

```text
in_TargetDate
in_StartDate
in_EndDate
```

### 4.4 クラス型の引数名・変数名

クラス型の引数名・変数名には、オブジェクトであることが分かる接頭辞 `c` を使用する。

例：

```text
in_cCurve
in_cCalendar
```

---

## 5. クラス設計

### 5.1 clsHolidayCalendar

#### 役割

休日データを保持するクラス。

当面は Tokyo holiday をコード内に保持する。  
将来的に、Holidayシート A列から休日データを読み込む方式へ切替可能な設計とする。

#### 主な責務

- 休日判定
- 東京休日の保持
- 外部休日データ読込への拡張余地の確保

#### 主なメソッド

```text
AddHoliday(in_TargetDate, in_Name)
LoadHolidays(in_Holidays)
IsHoliday(in_TargetDate) As Boolean
Count As Long
```

#### 補足

土日判定、営業日加算、Business Day Convention の判定は、本クラスでは行わない。  
それらは `mdl_BusinessDay` で行う。

### 5.2 clsOISSwap

#### 役割

JPY OIS スワップの商品条件を保持し、固定脚・変動脚のキャッシュフロー、PV、NPV、ParRateを計算する。

#### 主な責務

- StartDate / EndDate / FixedRate / Notional / PaymentLag の保持
- PAYER / RECEIVER の保持
- 支払スケジュール生成
- 固定脚PVの計算
- 変動脚PVの計算
- PayReceive を反映した NPV の計算
- ParRate の計算
- キャッシュフロー表を Variant 配列として返す

#### 符号規約

FixedLegPV および FloatingLegPV は、脚単体のPVとして正値で返す。  
NPV は PayReceive に応じて符号を反映する。

```text
PAYER    : FloatingLegPV - FixedLegPV
RECEIVER : FixedLegPV - FloatingLegPV
```
### 5.3 clsOISZeroLinearCurve

#### 役割

JPY OIS / TONA ベースのゼロレート点を保持し、評価日からの年数に対して連続複利ゼロレートを直線補間するディスカウントカーブ。

`clsOISStepForwardCurve` が瞬間ONフォワードを階段状に保持するのに対し、`clsOISZeroLinearCurve` はゼロレートそのものを補間対象とする。

#### 主な責務

- ValuationDate / SpotDate / PaymentLag / Business Day Convention の保持
- tenor / zero rate 配列からゼロレート点を構築
- 連続複利ゼロレートの直線補間
- DF の計算
- DF比による期間フォワードレートの計算
- `clsDiscountCurve` インターフェースの実装

#### 主なメソッド

```md
Init(in_ValuationDate, Optional in_Calendar, Optional in_spotLag, Optional in_PaymentLag, Optional in_BusinessDayConvention)
AddPoint(in_TargetDate, in_ZeroRateCont)
BuildFromZeroRates(in_Tenors, in_ZeroRates)
DF(in_TargetDate) As Double
ZeroRateCont(in_TargetDate) As Double
ForwardRate(in_StartDate, in_EndDate) As Double
GetPointTable() As Variant
ExportPointsToSheet(in_SheetName, in_StartCell)

```
## 6. 標準モジュール設計

### 6.1 mdl_BusinessDay

#### 役割

営業日判定、営業日加算、Business Day Convention に基づく営業日調整、テナー文字列による日付計算を行う。

休日データそのものは `clsHolidayCalendar` が保持し、本モジュールでは土日判定と休日カレンダーを組み合わせて営業日判定を行う。

#### 主な関数

```text
IsBusinessDay(in_TargetDate, in_cCalendar) As Boolean
AddBusinessDays(in_StartDate, in_BusinessDays, in_cCalendar) As Date
AdjustBusinessDay(in_TargetDate, in_BusinessDayConvention, in_cCalendar) As Date
AddTenorMonths(in_BaseDate, in_Months, in_KeepEndOfMonth) As Date
AddTenor(in_BaseDate, in_Tenor, in_KeepEndOfMonth) As Date
GetBusinessDay(in_BaseDate, in_Tenor, in_BusinessDayConvention, in_cCalendar, in_KeepEndOfMonth) As Date
GetWorkday(in_BaseDate, in_BusinessDays, in_cCalendar) As Date
```

#### GetBusinessDay の定義

`GetBusinessDay` は、指定日が営業日かどうかを返す関数ではない。

`GetBusinessDay` は、基準日にテナーを加算し、その後 Business Day Convention を適用した日付を返す関数とする。

処理順序は以下とする。

```text
1. AddTenor により未調整日付を作成する
2. AdjustBusinessDay により営業日調整する
3. 調整後日付を返す
```

#### Business Day Convention

以下の文字列指定に対応する。

```text
F  / Following
MF / ModifiedFollowing
P  / Preceding
```

既定値は `MF` とする。

#### テナー日付計算

テナー文字列は以下の形式を基本とする。

```text
+1D
-2D
3M
+6M
1Y
-1Y
```

`D` は暦日加算とする。

`M` は月加算とする。

`Y` は 12か月単位の月加算として扱う。

月加算は `DateAdd("m")` ではなく、`EDate` ベースで行う。

基準日が月末の場合は、原則として加算後も月末を維持する。

#### 注意点

営業日加算では、基準日はカウントしない。

`AddTenor` は、テナー加算のみを行い、営業日調整は行わない。

`GetBusinessDay` は、まずテナー加算により未調整日付を作成し、その後 Business Day Convention を適用する。

`GetWorkday` は旧名互換用とし、新規コードでは `AddBusinessDays` を使用する。

---

## 7. 基本実装範囲

当初実装では、以下を対象とする。

```text
1. clsHolidayCalendar
2. mdl_BusinessDay
3. mdl_DayCount
4. clsOISStepForwardCurve
5. OISカーブのDF、ZeroRateCont、ForwardRate取得
6. clsOISSwap
7. OISカーブのDF、ZeroRateCont、ForwardRate取得
8. OIS Swap のPV、NPV、ParRate、キャッシュフロー表取得
```

## 8. Cap Volatility Bootstrap の設計方針

本プロジェクトでは、Cap取引そのものを商品クラスとして評価することは当面の目的としない。

目的は、OIS / TONA ベースの変動金利に上限が付いた商品の caplet 部分を評価するために、quoted cap normal volatility から caplet normal volatility term structure を構築することである。

そのため、Cap商品クラス `clsCap` は作成しない。

Caplet単体のクラスも作成せず、caplet PV は標準モジュール `mdl_CapFormula` の Bachelier式で計算する。

### clsCapVolBootstrapper

#### 役割

Quoted cap normal vol から、期間別 caplet normal vol をブートストラップする。

#### 主な責務

- Cap maturity ごとの quoted cap normal vol を受け取る
- 対象期間の caplet schedule を内部生成する
- 各capletの forward rate を discount curve から取得する
- Bachelier式により caplet PV を計算する
- 既知期間の caplet vol を固定し、新規期間の caplet vol を二分法で解く
- 結果を `clsCapVolTermStructure` に格納する

### clsCapVolTermStructure

#### 役割

ブートストラップ済みの caplet normal volatility term structure を保持する。

#### 主な責務

- caplet期間ごとの normal vol を保持する
- period end date または fixing date により normal vol を取得する
- total variance を計算する
- 結果をシート出力する

### mdl_CapFormula

#### 役割

Caplet評価に必要な数式関数を提供する。

#### 主な関数

- NormalPDF
- NormalCDF
- BachelierCallValue
- CapletPV_Normal

### mdl_CapBootstrap

#### 役割

Cap volatility bootstrap のテスト・サンプル実行用モジュール。

本モジュールは本番ロジックではなく、動作確認用の入口として位置づける。

## 9. Swaption Volatility Surface の設計方針

本プロジェクトでは、スワップション取引そのものを商品クラスとして評価することは当面の目的としない。

目的は、CMS convexity adjustment や CMS spread swap 等で参照する ATM swaption volatility を、expiry × tenor のマトリックスから取得することである。

そのため、Swaption商品クラスは作成しない。

### clsSwaptionVol

#### 役割

ATM swaption volatility matrix を保持し、任意の expiry × tenor に対応する ATM volatility を返す。

#### 主な責務

- expiry grid の保持
- underlying swap tenor grid の保持
- ATM volatility matrix の保持
- tenor方向の variance 補間
- expiry方向の total variance 補間
- 範囲外指定時のエラー制御または flat extrapolation

#### 主なメソッド

```text
InitializeFromRange(in_MatrixRange, Optional in_AllowFlatExtrapolation)
Vol(in_ExpiryYears, in_TenorText) As Double
VolByYears(in_ExpiryYears, in_TenorYears) As Double
ExpiryAt(in_Index) As Double
TenorAt(in_Index) As Double
GridVol(in_ExpiryIndex, in_TenorIndex) As Double
```

### mdl_SwaptionVol

#### 役割

clsSwaptionVol を利用するための補助関数およびExcel関数ラッパーを提供する。

#### 主な関数

```text
SwaptionTenorToYears(in_TenorText) As Double
ATM_SWAPTION_VOL(in_ExpiryYears, in_TenorText, in_MatrixRange, Optional in_AllowFlatExtrapolation) As Variant
ATM_SWAPTION_VOL_TEXT(in_ExpiryText, in_TenorText, in_MatrixRange, Optional in_AllowFlatExtrapolation) As Variant
SWAPTION_TENOR_TO_YEARS(in_TenorText) As Variant
```
