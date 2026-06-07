# 01_Discount_Curve

## 目的

`01_Discount_Curve` は、JPY TONA/OIS の割引カーブ構築と、割引カーブの共通インターフェースを管理する領域です。

OISスワップ評価、Hull-White 1F、CMSスプレッド評価などの上流にある基礎モジュールとして位置づけます。

## 想定する内容

```text
classes/
  clsDiscountCurve.cls
  clsOISStepForwardCurve.cls
  clsOISZeroLinearCurve.cls
modules/
  mdl_CurveMath.bas
```

## 主な役割

- Discount Factor を返す。
- 連続複利ゼロレートを返す。
- 期間フォワードレートを返す。
- 年数ベースの `DF_T(T)`、`ForwardRate_T(T1, T2)`、`InstantaneousForward(T)` などを提供する。
- 他の評価モデルが使いやすい形で、割引カーブを共通化する。

## 依存関係のルール

このモジュールは、共通部品として `src/00_Common` を利用してよいです。

許容する依存関係：

```text
01_Discount_Curve -> 00_Common
```

避ける依存関係：

```text
01_Discount_Curve -> 02_ois_swap
01_Discount_Curve -> 03_sabr
01_Discount_Curve -> 04_hull_white_1f
01_Discount_Curve -> 05_cms_spread
```

## 設計上の注意点

割引カーブは、商品評価やモデル評価の上流にある基礎部品です。

そのため、OISスワップの商品条件やキャッシュフロー評価、Hull-Whiteのシミュレーション、CMSスプレッドの評価ロジックを、このモジュールに直接混ぜない方針とします。

カーブ構築に必要な最小限の補助計算は `modules/mdl_CurveMath.bas` に置き、商品固有の評価処理は下流モジュールへ分離します。
