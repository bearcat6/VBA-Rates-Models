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
