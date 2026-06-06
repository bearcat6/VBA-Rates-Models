# Design

## 1. 目的

本ドキュメントは、VBA-Rates-Models における金利モデル関連コードの設計方針を整理するものである。

当面の目的は、JPY TONA/OIS ベースのディスカウントカーブを構築し、将来的に TONA 変動金利にキャップが付いた商品の評価や、キャップボラティリティ期間構造の構築へ拡張できる設計とすることである。

設計上の詳細な前提は、`docs/Assumptions.md` に従う。

---

## 2. 全体構成

本プロジェクトは、以下の構成とする。

```text
VBA-Rates-Models
├─ docs
│  ├─ Assumptions.md
│  └─ Design.md
├─ src
│  ├─ classes
│  │  ├─ clsHolidayCalendar.cls
│  │  ├─ clsDiscountCurve.cls
│  │  ├─ clsOISStepForwardCurve.cls
│  │  ├─ clsOISZeroLinearCurve.cls
│  │  ├─ clsOISswap.cls
│  │  ├─ clsATMSwaptionVol.cls
│  │  ├─ clsCapVolBootstrapper
│  │  ├─ clsCapVolTermStructure
│  │  ├─ clsVolSurface.cls
│  │  ├─ clsSABRParams.cls
│  │  ├─ clsHullWhite1F.cls
│  │  ├─ clsHWCalibrator.cls
│  │  ├─ clsHWSimulator.cls
│  │  └─ clsRandomNormal.cls
│  └─ modules
│     ├─ mdl_Common.bas
│     ├─ mdl_DayCount.bas
│     ├─ mdl_BusinessDay.bas
│     ├─ mdl_CurveMath.bas
│     ├─ mdl_DiscountCurveFactory.bas
│     ├─ mdl_HullWhiteMath.bas
│     ├─ mdl_HullWhiteWorkFlow.bas
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

```text
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

## 9. ATM Swaption Volatility Surface の設計方針

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
GridVol(in_ExpiryIndex) As Double
```

### mdl_ATMSwaptionVol

#### 役割

clsATMSwaptionVol を利用するための補助関数およびExcel関数ラッパーを提供する。

#### 主な関数

```text
SwaptionTenorToYears(in_TenorText) As Double
ATM_SWAPTION_VOL(in_ExpiryYears, in_TenorText, in_MatrixRange, Optional in_AllowFlatExtrapolation) As Variant
ATM_SWAPTION_VOL_TEXT(in_ExpiryText, in_TenorText, in_MatrixRange, Optional in_AllowFlatExtrapolation) As Variant
SWAPTION_TENOR_TO_YEARS(in_TenorText) As Variant
```

## 10. Hull-White 1F 関連クラス

本章では、1ファクター Hull-White モデルに関連するクラスの責務を整理する。

Hull-White 1F は、JPY OISカーブとスワップション・ボラティリティ等を入力として、将来時点の金利カーブをモンテカルロ・シミュレーションし、分位点カーブやストレスカーブを作成するために利用する。

初期実装では、平均回帰パラメータ `a` は外部から固定値として与え、短期金利ボラティリティ `sigma` を ATM normal swaption volatility にフィットする。

モデル本体、キャリブレーション、シミュレーション、乱数生成の責務を分離し、Excelシートからの入出力やワークフロー制御は標準モジュール側で扱う。

---

### 10.1 clsHullWhite1F

#### 役割

`clsHullWhite1F` は、1ファクター Hull-White モデル本体を表すクラスである。

ディスカウントカーブ、評価日、平均回帰パラメータ `a`、短期金利ボラティリティ `sigma` を保持し、Hull-White 1F に必要なモデル計算を提供する。

対象モデルは以下とする。

```text
dr(t) = { theta(t) - a r(t) } dt + sigma dW(t)

```

実装上は、以下の shifted short-rate representation を利用する。

```text
r(t) = phi(t) + x(t)

dx(t) = -a x(t) dt + sigma dW(t)
```

ここで、`phi(t)` は初期ディスカウントカーブと整合するように定まるシフト項であり、`x(t)` は平均回帰する確率ファクターである。

#### 主な責務

- 初期ディスカウントカーブの参照
- 評価日の保持
- 平均回帰パラメータ `a` の保持
- 短期金利ボラティリティ `sigma` の保持
- Hull-White の `B(t,T)` の計算
- Hull-White の `A(t,T)` の計算
- 将来時点の割引債価格 `P(t,T)` の計算
- shifted short-rate factor `x(t)` の1ステップ更新
- 初期フォワードレートおよびシフト項の計算

#### 設計方針

`clsHullWhite1F` は、具体的なOISカーブ構築方法には依存しない。

Hull-White 1F は時間軸 `t`, `T` を用いるため、カーブオブジェクトには、評価日からの年数を引数とする以下の time-based interface を要求する。

```vb
DF_T(T)
InstantaneousForward(T)
```

これにより、`clsOISStepForwardCurve`、`clsOISZeroLinearCurve` など、異なるカーブ構築方法を持つクラスを同じHull-Whiteモデルから利用できるようにする。

#### 注意点

`theta(t)` は直接入力パラメータとして保持しない。

初期ディスカウントカーブと整合するように `phi(t)` を扱うことで、現在の市場カーブを再現する設計とする。

---

### 10.2 clsHWCalibrator

#### 役割

`clsHWCalibrator` は、ディスカウントカーブとボラティリティ・サーフェスを入力として、1ファクター Hull-White モデルのパラメータを推定するクラスである。

初期実装では、平均回帰パラメータ `a` は外部から固定値として与え、短期金利ボラティリティ `sigma` を ATM normal swaption volatility にフィットする。

#### 主な責務

- ディスカウントカーブの参照
- Hull-White 1F キャリブレーション用のボラティリティ・サーフェスの参照
- キャリブレーション対象となる expiry / tenor quote の保持
- 固定された `a` に対する `sigma` の推定
- 市場ボラティリティとモデルボラティリティの誤差計算
- キャリブレーション結果およびレポート用配列の生成

#### 初期スコープ

初期実装では、厳密な Jamshidian decomposition によるスワップション評価は行わない。

スワップレートの Hull-White factor に対する感応度を用いた簡易的な normal volatility 近似により、`sigma` を推定する。

この実装は、将来金利カーブのモンテカルロ・シミュレーションやストレスカーブ作成に用いる初期パラメータ推定を目的とする。

#### 将来拡張

将来的には、`clsHullWhite1F` または `mdl_HullWhiteMath` に厳密な bond option / swaption pricing を実装し、`clsHWCalibrator` の目的関数から呼び出す設計に拡張する。

また、初期実装では `a` を固定値として扱うが、将来的には `a` と `sigma` の同時推定、または expiry / tenor ごとのフィット状況を確認する診断機能を追加する。

---

### 10.3 clsHWSimulator

#### 役割

`clsHWSimulator` は、ディスカウントカーブと Hull-White 1F のパラメータ `a`, `sigma` を用いて、将来時点の金利カーブをモンテカルロ・シミュレーションするクラスである。

主に、1年後などの将来時点における short rate、zero rate curve、discount factor curve、forward rate curve を生成し、分位点カーブやストレスカーブ作成に利用する。

#### 主な責務

- ディスカウントカーブの参照
- 平均回帰パラメータ `a` の保持
- 短期金利ボラティリティ `sigma` の保持
- シミュレーション期間 horizon の保持
- タイムステップ `dt` の保持
- shifted short-rate factor `x(t)` のパス生成
- short rate path の生成
- horizon時点の discount factor curve の生成
- horizon時点の zero rate curve の生成
- horizon時点の forward rate curve の生成
- パス別結果および分位点計算用データの生成

#### シミュレーション方法

`x(t)` は Ornstein-Uhlenbeck 過程として、以下の exact discretization により更新する。

```text
x(t+dt) = x(t) exp(-a dt)
          + sigma sqrt((1 - exp(-2a dt)) / (2a)) Z
```

`a` がゼロに近い場合は、以下の極限形を用いる。

```text
x(t+dt) = x(t) + sigma sqrt(dt) Z
```

ここで、`Z` は標準正規乱数である。

#### 注意点

初期実装では、将来時点のカーブ生成は、初期カーブを基準に Hull-White の bond sensitivity とシミュレーションされた factor を用いて行う。

厳密な無裁定将来カーブ生成というより、リスク管理上の分位点カーブ・ストレスカーブ作成を目的とした実務的な初期版とする。

---

### 10.4 clsRandomNormal

#### 役割

`clsRandomNormal` は、モンテカルロ・シミュレーションで利用する標準正規乱数を生成するクラスである。

乱数生成ロジックを `clsHWSimulator` などのモデルクラスから分離することで、シミュレーション本体の責務を明確にする。

#### 主な責務

- 標準正規乱数の生成
- seed指定による再現性の確保
- Box-Muller法による正規乱数生成
- Box-Muller法で生成される2つ目の乱数のキャッシュ
- antithetic variates の利用

#### 設計方針

初期実装では、VBA標準の `Rnd` / `Randomize` を乱数源として利用する。

金融モンテカルロとしての厳密な乱数品質検証は、初期実装の対象外とする。

ただし、seedを指定できる設計とし、同じ入力条件で同じシミュレーション結果を再現できるようにする。

#### 将来拡張

将来的には、より高品質な乱数生成器、準乱数、パス単位の乱数管理、またはシナリオ再現用の乱数保存機能を追加する余地を残す。
