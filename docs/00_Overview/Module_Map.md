# モジュール構成図

このドキュメントは、`VBA-Rates-Models` の最終的なモジュール構成を整理するものです。

従来のように `src/classes` と `src/modules` にファイルを横並びで置くのではなく、分析目的ごとにフォルダを分けて管理します。

## 最終的な src 構成

```text
src/
├─ 00_Common/
│  ├─ README.md
│  └─ modules/
│     ├─ mdl_Common.bas
│     └─ mdl_BusinessDay.bas
├─ 01_Discount_Curve/
│  ├─ classes/
│  └─ modules/
├─ 02_OIS_Swap/
│  ├─ classes/
│  └─ modules/
├─ 03_SABR/
│  ├─ classes/
│  └─ modules/
├─ 04_Hull-White_1F/
│  ├─ classes/
│  └─ modules/
├─ 05_CMS-Spread/
│  ├─ classes/
│  └─ modules/
└─ 06_RandomNumber/
   ├─ classes/
   └─ modules/
```

## 各モジュールの役割

| モジュール | 役割 |
|---|---|
| `00_Common` | 日付処理、日数計算、営業日調整、配列処理、数値計算、乱数、正規分布、Bachelier関連などの共通部品 |
| `01_Discount_Curve` | JPY TONA/OIS の基本的な割引カーブ構築とカーブインターフェース |
| `02_OIS_Swap` | OISスワップの商品条件、キャッシュフロー生成、PV、NPV、パーレート計算 |
| `03_SABR` | Normal SABR のパラメータ、スマイルフィット、価格曲線の平滑化、密度計算、ストライクと分位点の変換 |
| `04_Hull-White_1F` | 1ファクター Hull-White モデル、ボラティリティへのキャリブレーション、モンテカルロシミュレーション、将来カーブ生成 |
| `05_CMS-Spread` | CMSおよびCMSスプレッド評価関連。現時点では主にスワップション・ボラティリティ処理 |
| `06_RandomNumber` | Hull-White 1Fなどのモンテカルロシミュレーションで利用する乱数生成ロジック |

## docs 側のフォルダ名

ドキュメント側では、見やすさを優先して以下の表記にします。

```text
docs/01_Discount_Curve/
docs/02_OIS_Swap/
docs/03_SABR/
docs/04_Hull-White_1F/
docs/05_CMS-Spread/
docs/06_RandomNumber/
```

VBAコード側の `src` フォルダについても、ドキュメント側と表記を合わせ、目的別のフォルダ名を使います。

## 依存関係の方向

依存関係は、原則として上流から下流へ一方向にします。

```text
00_Common
  ↓
01_Discount_Curve
  ↓
02_OIS_Swap

00_Common + 01_Discount_Curve
  ↓
04_Hull-White_1F

00_Common + 01_Discount_Curve + 03_SABR
  ↓
05_CMS-Spread
```

逆方向の依存は避けます。たとえば、`01_Discount_Curve` から `04_Hull-White_1F` や `05_CMS-Spread` を呼び出す構成にはしません。

## 現在のフラット構成からの移動方針

| 現在のパス | 移動後の想定パス | 補足 |
|---|---|---|
| `src/modules/mdl_Common.bas` | `src/00_Common/modules/mdl_Common.bas` | 共通ユーティリティ、ACT/365F日数計算。移動済み |
| `src/modules/mdl_BusinessDay.bas` | `src/00_Common/modules/mdl_BusinessDay.bas` | 営業日判定・営業日調整・テナー日付計算。移動済み |
| `src/classes/clsHolidayCalendar.cls` | `src/00_Common/classes/clsHolidayCalendar.cls` | 休日カレンダー。移動候補 |
| `src/classes/clsRandomNormal.cls` | `src/00_Common/classes/clsRandomNormal.cls` | モンテカルロ用の共通乱数部品。移動候補 |
| `src/classes/clsDiscountCurve.cls` | `src/01_Discount_Curve/classes/clsDiscountCurve.cls` | カーブインターフェース |
| `src/classes/clsOISStepForwardCurve.cls` | `src/01_Discount_Curve/classes/clsOISStepForwardCurve.cls` | Step Forward型OISカーブ |
| `src/classes/clsOISZeroLinearCurve.cls` | `src/01_Discount_Curve/classes/clsOISZeroLinearCurve.cls` | ゼロレート直線補間型カーブ |
| `src/modules/mdl_CurveMath.bas` | `src/01_Discount_Curve/modules/mdl_CurveMath.bas` | カーブ関連の数値計算 |
| `src/classes/clsVolSurface.cls` | `src/04_Hull-White_1F/classes/clsVolSurface.cls` | 現時点ではHull-White入力用の汎用ボラティリティ・サーフェス |
| `src/classes/clsHullWhite1F.cls` | `src/04_Hull-White_1F/classes/clsHullWhite1F.cls` | Hull-Whiteモデル本体 |
| `src/classes/clsHWCalibrator.cls` | `src/04_Hull-White_1F/classes/clsHWCalibrator.cls` | Hull-Whiteキャリブレーション |
| `src/classes/clsHWSimulator.cls` | `src/04_Hull-White_1F/classes/clsHWSimulator.cls` | Hull-Whiteシミュレーション |
| `src/modules/mdl_HullWhiteMath.bas` | `src/04_Hull-White_1F/modules/mdl_HullWhiteMath.bas` | Hull-White関連の数値計算 |
| `src/modules/mdl_HullWhiteWorkFlow.bas` | `src/04_Hull-White_1F/modules/mdl_HullWhiteWorkflow.bas` | Excel上でのHull-White実行フロー |
| `src/classes/clsATMSwaptionVol.cls` | `src/05_CMS-Spread/classes/clsATMSwaptionVol.cls` | CMS関連評価で使うATMスワップション・ボラティリティ |

## 命名ルール

- クラスモジュールは `cls` で始めます。
- 標準モジュールは `mdl_` で始めます。
- 関数の入力引数には `in_` を付けます。
- クラス型の変数や引数は、可能な範囲で `c` を付けます。例：`in_cCurve`。

## 移行時のルール

VBAファイルを移動するときは、まず中身を変えずに移動します。

ファイル移動と中身のリファクタリングを同じコミットで行うと、問題が起きたときに原因を追いにくくなるため、以下の順番を原則とします。

```text
1. 既存ファイルの内容をそのまま新しい場所へ移動
2. Excel/VBAへインポートできることを確認
3. 必要に応じて参照関係やコメントを修正
4. ロジックのリファクタリングは別コミットで実施
```
