# モジュール構成図

このドキュメントは、`VBA-Rates-Models` の最終的なモジュール構成を整理するものです。

従来のように `src/classes` と `src/modules` にファイルを横並びで置くのではなく、分析目的ごとにフォルダを分けて管理します。

## 最終的な src 構成

```text
src/
├─ common/
│  ├─ classes/
│  └─ modules/
├─ 01_discount_curve/
│  ├─ classes/
│  └─ modules/
├─ 02_ois_swap/
│  ├─ classes/
│  └─ modules/
├─ 03_sabr/
│  ├─ classes/
│  └─ modules/
├─ 04_hull_white_1f/
│  ├─ classes/
│  └─ modules/
└─ 05_cms_spread/
   ├─ classes/
   └─ modules/
```

## 各モジュールの役割

| モジュール | 役割 |
|---|---|
| `common` | 日付処理、日数計算、営業日調整、配列処理、数値計算、乱数、正規分布、Bachelier関連などの共通部品 |
| `01_discount_curve` | JPY TONA/OIS の基本的な割引カーブ構築とカーブインターフェース |
| `02_ois_swap` | OISスワップの商品条件、キャッシュフロー生成、PV、NPV、パーレート計算 |
| `03_sabr` | Normal SABR のパラメータ、スマイルフィット、価格曲線の平滑化、密度計算、ストライクと分位点の変換 |
| `04_hull_white_1f` | 1ファクター Hull-White モデル、ボラティリティへのキャリブレーション、モンテカルロシミュレーション、将来カーブ生成 |
| `05_cms_spread` | CMSおよびCMSスプレッド評価関連。現時点では主にスワップション・ボラティリティ処理 |

## 依存関係の方向

依存関係は、原則として上流から下流へ一方向にします。

```text
common
  ↓
01_discount_curve
  ↓
02_ois_swap

common + 01_discount_curve
  ↓
04_hull_white_1f

common + 01_discount_curve + 03_sabr
  ↓
05_cms_spread
```

逆方向の依存は避けます。たとえば、`01_discount_curve` から `04_hull_white_1f` や `05_cms_spread` を呼び出す構成にはしません。

## 現在のフラット構成からの移動方針

| 現在のパス | 移動後の想定パス | 補足 |
|---|---|---|
| `src/classes/clsDiscountCurve.cls` | `src/01_discount_curve/classes/clsDiscountCurve.cls` | カーブインターフェース |
| `src/classes/clsOISStepForwardCurve.cls` | `src/01_discount_curve/classes/clsOISStepForwardCurve.cls` | Step Forward型OISカーブ |
| `src/classes/clsOISZeroLinearCurve.cls` | `src/01_discount_curve/classes/clsOISZeroLinearCurve.cls` | ゼロレート直線補間型カーブ |
| `src/modules/mdl_CurveMath.bas` | `src/01_discount_curve/modules/mdl_CurveMath.bas` | カーブ関連の数値計算 |
| `src/classes/clsRandomNormal.cls` | `src/common/classes/clsRandomNormal.cls` | モンテカルロ用の共通乱数部品 |
| `src/classes/clsVolSurface.cls` | `src/04_hull_white_1f/classes/clsVolSurface.cls` | 現時点ではHull-White入力用の汎用ボラティリティ・サーフェス |
| `src/classes/clsHullWhite1F.cls` | `src/04_hull_white_1f/classes/clsHullWhite1F.cls` | Hull-Whiteモデル本体 |
| `src/classes/clsHWCalibrator.cls` | `src/04_hull_white_1f/classes/clsHWCalibrator.cls` | Hull-Whiteキャリブレーション |
| `src/classes/clsHWSimulator.cls` | `src/04_hull_white_1f/classes/clsHWSimulator.cls` | Hull-Whiteシミュレーション |
| `src/modules/mdl_HullWhiteMath.bas` | `src/04_hull_white_1f/modules/mdl_HullWhiteMath.bas` | Hull-White関連の数値計算 |
| `src/modules/mdl_HullWhiteWorkFlow.bas` | `src/04_hull_white_1f/modules/mdl_HullWhiteWorkflow.bas` | Excel上でのHull-White実行フロー |
| `src/classes/clsATMSwaptionVol.cls` | `src/05_cms_spread/classes/clsATMSwaptionVol.cls` | CMS関連評価で使うATMスワップション・ボラティリティ |

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
