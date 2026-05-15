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
│  │  └─ clsOISStepForwardCurve.cls
│  └─ modules
│     ├─ mdl_DateUtil.bas
│     ├─ mdl_DayCount.bas
│     ├─ mdl_BusinessDay.bas
│     ├─ mdl_DiscountCurveFactory.bas
│     └─ mdl_CapVolUtil.bas
└─ README.md
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

```text
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

---

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

## 7. 当初実装範囲

当初実装では、以下を対象とする。

```text
1. clsHolidayCalendar
2. mdl_BusinessDay
3. mdl_DayCount
4. clsOISStepForwardCurve
5. OISカーブのDF、ZeroRateCont、ForwardRate取得
```

Cap Volatility 以降は、OISカーブの実装後に設計を具体化する。
