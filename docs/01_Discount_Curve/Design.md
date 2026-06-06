# 割引カーブ設計

## 目的

`01_Discount_Curve` は、JPY TONA/OIS ベースの基本的な割引カーブ構築を扱うモジュールです。

このモジュールは、OISスワップ評価、Hull-White 1F のキャリブレーション・シミュレーション、将来のCMSスプレッド評価の基盤になります。

## 役割

このモジュールの主な役割は以下のとおりです。

- 割引カーブの共通インターフェースを定義する。
- JPY OIS割引カーブを構築する。
- Discount Factor、ゼロレート、フォワードレートを返す。
- Hull-White 1F などの金利モデルで利用する time-based interface を提供する。

## 主なクラス

| クラス | 役割 |
|---|---|
| `clsDiscountCurve` | 割引カーブのインターフェース的なクラス |
| `clsOISStepForwardCurve` | 瞬間ONフォワード金利を階段状に保持するOISカーブ |
| `clsOISZeroLinearCurve` | 連続複利ゼロレートを直線補間するOISカーブ |

## 主な標準モジュール

| モジュール | 役割 |
|---|---|
| `mdl_CurveMath` | ゼロレート変換、フォワードレート変換など、カーブ関連の数値計算 |

## インターフェース

日付ベースのインターフェースは以下を基本とします。

```vb
DF(in_TargetDate)
ZeroRateCont(in_TargetDate)
ForwardRate(in_StartDate, in_EndDate)
```

金利モデルで利用する年数ベースのインターフェースは以下を基本とします。

```vb
DF_T(T)
ZeroRate_T(T)
ForwardRate_T(T1, T2)
InstantaneousForward(T)
```

## 依存関係のルール

このモジュールは `common` に依存してよいですが、商品評価モジュールやモデルモジュールには依存しない方針とします。

許容する依存関係：

```text
01_Discount_Curve -> common
```

避ける依存関係：

```text
01_Discount_Curve -> 02_ois_swap
01_Discount_Curve -> 04_hull_white_1f
01_Discount_Curve -> 05_cms_spread
```

## 設計上の注意点

OISカーブのブートストラップでは、プロトタイプ段階でOISスワップのパーレート計算に近い補助関数をカーブクラス内に置くことがあります。

ただし、長期的には以下を分離します。

```text
割引カーブ構築
商品条件・キャッシュフロー生成
商品評価
```

割引カーブクラスが商品評価ロジックを抱え込みすぎると、後からSABR、Hull-White、CMS評価へ拡張する際に依存関係が複雑になります。

そのため、`01_Discount_Curve` はあくまで「市場データから割引カーブを作り、DF・ゼロレート・フォワードレートを返す基礎部品」として管理します。
