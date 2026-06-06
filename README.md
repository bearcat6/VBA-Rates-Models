# VBA Rates Models

Excel/VBAでJPY金利分析を試作するためのリポジトリです。

このリポジトリでは、JPY OIS割引カーブ、OISスワップ評価、SABRスマイル、1ファクター Hull-White モデル、CMSスプレッド評価への拡張を、目的別のモジュールとして管理します。

## 対象範囲

VBAクラスを単純に横並びで管理するのではなく、目的別に分けて管理します。

| 領域 | 目的 | 状況 |
|---|---|---|
| `01_Discount_Curve` | JPY TONA/OIS の基本的な割引カーブ構築 | 対応中 |
| `02_ois_swap` | OISスワップのキャッシュフロー生成と評価 | 対応中・拡張中 |
| `03_sabr` | Normal SABR のスマイル、密度、分位点計算 | 計画中・拡張中 |
| `04_hull_white_1f` | Hull-White 1F のキャリブレーションと将来カーブシミュレーション | 対応中 |
| `05_cms_spread` | CMSスプレッド評価。現時点では主にスワップション・ボラティリティ処理 | 計画中・一部対応 |

## 最終的なリポジトリ構成

```text
VBA-Rates-Models/
├─ docs/
│  ├─ 00_Overview/
│  │  ├─ Module_Map.md
│  │  └─ Roadmap.md
│  ├─ 01_Discount_Curve/
│  │  └─ Design.md
│  ├─ 02_ois_swap/
│  │  └─ Design.md
│  ├─ 03_sabr/
│  │  └─ Design.md
│  ├─ 04_hull_white_1f/
│  │  └─ Design.md
│  └─ 05_cms_spread/
│     └─ Design.md
├─ src/
│  ├─ common/
│  ├─ 01_discount_curve/
│  ├─ 02_ois_swap/
│  ├─ 03_sabr/
│  ├─ 04_hull_white_1f/
│  └─ 05_cms_spread/
├─ examples/
├─ tests/
└─ README.md
```

## 依存関係の方向

依存関係は、原則として一方向にします。

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

`03_sabr` はCMS専用ではなく、リスク分析やストレスシナリオ作成にも使えるため、CMSとは独立したモジュールとして管理します。

## 現在の移行方針

既存のVBAファイルは、内容を安全に保持したまま移動できる状態になるまで、当面は旧パスにも残します。

まず最終的なモジュール構成とドキュメントを整備し、その後、VBAファイルを1つずつ内容を壊さないように移動します。

旧来のフラット構成から新しい目的別構成への移動表は、`docs/00_Overview/Module_Map.md` を参照してください。

## 設計方針

- 共通部品と、商品評価・モデルロジックを分けます。
- 割引カーブ構築と取引評価を分けます。
- モデルのキャリブレーションやシミュレーションは、Excelシート入出力に直接依存しないようにします。
- Excel/VBAで扱いやすいインターフェースを優先します。
- 秘密情報、顧客情報、社内専有データは含めません。

## 注意事項

このリポジトリは、教育・検証・プロトタイピングを目的としたものです。

実務利用する場合は、モデル妥当性、数値精度、境界条件、監査対応、データ管理を別途確認する必要があります。
