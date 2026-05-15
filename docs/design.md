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
