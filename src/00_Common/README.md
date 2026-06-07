# 00_Common

## 目的

`00_Common` は、割引カーブ、OISスワップ、SABR、Hull-White、CMSスプレッド評価などで共通利用するVBA部品を管理する領域です。

特定の商品評価や金利モデルに依存しない、汎用的な関数・クラスをここに集約します。

## 想定する内容

- 日付処理
- 日数計算
- 営業日調整
- 配列処理
- 数値計算
- 数値補間
- 乱数生成
- 正規分布関連の補助関数
- Bachelier / Normal Option Pricing 関連の補助関数
- Excel入出力に近いが、複数モジュールで共通利用する軽量ユーティリティ

## 現在の構成

```text
classes/
  clsSplineCurve.cls
modules/
  mdl_Common.bas
  mdl_BusinessDay.bas
```

## 現在の移動対象

| 旧パス | 新パス | 補足 |
|---|---|---|
| `src/modules/mdl_Common.bas` | `src/00_Common/modules/mdl_Common.bas` | 共通ユーティリティ、ACT/365F日数計算。移動済み |
| `src/modules/mdl_BusinessDay.bas` | `src/00_Common/modules/mdl_BusinessDay.bas` | 営業日判定・営業日調整・テナー日付計算。移動済み |
| `Financial_Code/cls/clsSplineCurve.cls` | `src/00_Common/classes/clsSplineCurve.cls` | Natural Cubic Spline補間。割引カーブ、ボラカーブ、SABR、Hull-Whiteで共通利用。移動済み |
| `src/classes/clsHolidayCalendar.cls` | `src/00_Common/classes/clsHolidayCalendar.cls` | 休日カレンダー。移動候補 |
| `src/classes/clsRandomNormal.cls` | `src/00_Common/classes/clsRandomNormal.cls` | 標準正規乱数生成。移動候補 |

## 依存関係のルール

`00_Common` は、他の下流モジュールに依存しない方針とします。

```text
00_Common
  ↓
01_Discount_Curve
02_OIS_Swap
03_SABR
04_Hull-White_1F
05_CMS-Spread
06_RandomNumber
```

ここに置くファイルは、できるだけ以下に依存しないようにします。

```text
特定の商品クラス
特定の金利モデルクラス
特定のExcelシート名
特定の市場データ形式
```

## 注意点

便利だからという理由だけで、何でも `00_Common` に置くと、後から責務が曖昧になります。

共通モジュールに置く基準は、複数の領域から自然に再利用され、かつ特定モデルや商品に依存しないことです。

`clsSplineCurve` は汎用的な数値補間部品です。金利カーブそのものではなく、割引カーブ、ボラティリティカーブ、モデルパラメータ曲線などの内部補間エンジンとして利用します。
