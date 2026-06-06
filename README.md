# VBA Rates Models

JPY金利分析をExcel/VBAで試作するためのリポジトリです。

主な対象は、TONA/OISベースのディスカウントカーブ構築、キャップ・ボラティリティのブートストラップ、ATMスワップション・ボラティリティの補間、CMS関連評価ユーティリティ、および1ファクター Hull-White モデルによる将来金利カーブのシミュレーションです。

## 目的

このリポジトリの目的は、JPY OISカーブと金利ボラティリティを入力として、金利モデル・評価ロジック・リスク分析用のプロトタイプをVBAで実装することです。

特に、Hull-White 1Fモデルについては、以下を目標とします。

- ディスカウントカーブとスワップション・ボラティリティからモデルパラメータを推定する
- 1年後などの将来時点における金利カーブをモンテカルロ・シミュレーションする
- シミュレーション結果から分位点カーブを作成する
- リスク管理で利用するストレスカーブを作成する

## 主な機能

### OISカーブ

- JPY TONA/OISベースのディスカウントカーブ構築
- Step forward型のOISカーブ
- Zero rate linear interpolation型のOISカーブ
- ディスカウントファクター、ゼロレート、フォワードレートの計算

### ボラティリティ関連

- キャップ・ボラティリティのブートストラップ
- ATMスワップション・ボラティリティ行列の保持と補間
- Hull-White 1Fキャリブレーション用の汎用ボラティリティ・サーフェス
- Normal Volを中心とした初期実装

### Hull-White 1F

- 1ファクター Hull-White モデル本体
- 平均回帰パラメータ `a` とボラティリティ `sigma` の管理
- 初期実装では、`a` を外部から固定値として与え、`sigma` をATM normal swaption volatilityへフィットする
- shifted short-rate representation による実装
- 将来金利カーブのモンテカルロ・シミュレーション
- 分位点カーブ・ストレスカーブ作成に向けた出力

### CMS関連

- CMS convexity adjustment などで利用するATMスワップション・ボラティリティ参照
- CMS spread swap等への拡張を意識した設計

## ディレクトリ構成

```text
VBA-Rates-Models
├─ docs
│  ├─ Assumptions.md
│  ├─ Design.md
│  └─ Hull-White_1F_Curve_Interface.md
├─ src
│  ├─ classes
│  │  ├─ clsDiscountCurve.cls
│  │  ├─ clsOISStepForwardCurve.cls
│  │  ├─ clsOISZeroLinearCurve.cls
│  │  ├─ clsATMSwaptionVol.cls
│  │  ├─ clsVolSurface.cls
│  │  ├─ clsHullWhite1F.cls
│  │  ├─ clsHWCalibrator.cls
│  │  ├─ clsHWSimulator.cls
│  │  └─ clsRandomNormal.cls
│  └─ modules
│     ├─ mdl_CurveMath.bas
│     ├─ mdl_HullWhiteMath.bas
│     └─ mdl_HullWhiteWorkFlow.bas
└─ README.md
```

実際の実装ファイルは今後も追加・整理される可能性があります。詳細な設計方針は `docs/Design.md`、実装前提は `docs/Assumptions.md` を参照してください。

## クラス名に関する補足

`src/classes/clsATMSwaptionVol.cls` は、ATMスワップション・ボラティリティ行列を保持し、任意の Expiry × Tenor に対するATMボラティリティを返すクラスです。

このクラス名は、将来的に非ATMのスワップション・ボラティリティ・サーフェスやスマイルモデルを追加する可能性を踏まえ、あえてATM専用であることが分かる名前にしています。

一方、`src/classes/clsVolSurface.cls` は、Hull-White 1Fなどのモデル入力に利用する、より汎用的なボラティリティ・サーフェスとして位置づけます。

## Hull-White 1Fの実装方針

Hull-White 1Fでは、以下のモデルを対象とします。

```text
dr(t) = { theta(t) - a r(t) } dt + sigma dW(t)
```

実装上は、以下の shifted short-rate representation を利用します。

```text
r(t) = phi(t) + x(t)

dx(t) = -a x(t) dt + sigma dW(t)
```

`theta(t)` は直接入力パラメータとして扱わず、初期ディスカウントカーブと整合するように扱います。

Hull-White関連クラスは、カーブオブジェクトに対して、評価日からの年数 `T` を引数とする以下のような time-based interface を要求します。

```vb
DF_T(T)
InstantaneousForward(T)
```

この方針により、Hull-Whiteモデル本体は、OISカーブの具体的な構築方法に依存しない設計とします。

## 注意事項

このリポジトリは、教育・検証・プロトタイピングを目的としたものです。

- サンプルデータは架空データを使用します
- 秘密情報・顧客情報・社内専有データは含めません
- 実務利用する場合は、モデル妥当性、数値精度、境界条件、監査対応、データ管理を別途確認する必要があります
- 初期のHull-White 1Fキャリブレーションは簡易実装であり、厳密なスワップション評価モデルは今後の拡張対象です
