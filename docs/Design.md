# Design

このドキュメントは、旧リンクとの互換性を残すためのトップページです。

現在の設計ドキュメントは、目的別モジュール構成に移行しています。

## 現在の設計ドキュメント

| ドキュメント | 内容 |
|---|---|
| `docs/00_Overview/Module_Map.md` | 全体のモジュール構成と移動方針 |
| `docs/00_Overview/Roadmap.md` | 開発順序と移行ロードマップ |
| `docs/01_Discount_Curve/Design.md` | 割引カーブ設計 |
| `docs/02_OIS_Swap/Design.md` | OISスワップ評価設計 |
| `docs/03_sabr/Design.md` | Normal SABR設計 |
| `docs/04_hull_white_1f/Design.md` | Hull-White 1F設計 |
| `docs/05_cms_spread/Design.md` | CMSスプレッド評価設計 |

## 設計方針

リポジトリは、以下の目的別モジュールで管理します。

```text
common
01_discount_curve
02_ois_swap
03_sabr
04_hull_white_1f
05_cms_spread
```

従来の `src/classes` と `src/modules` に横並びで置く構成から、`docs/00_Overview/Module_Map.md` に記載した目的別構成へ移行します。

## 依存関係の方向

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

逆方向の依存は避けます。

また、割引カーブ構築クラスに商品評価ロジックを過度に混ぜない方針とします。
