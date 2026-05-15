# Assumptions

## 全体前提
- 通貨は JPY
- 金利は TONA/OISベース
- 日付は Excel Date型
- 営業日調整は Modified Following
- 休日カレンダーは Holidayシート A列
- 関数の引数には、in_ を入れる
- クラスの頭文字は、cls
- モジュールの頭文字は、mdl_

## Holiday Calendaer
- クラス名：clsHolidayCalendar
- IsHoliday(day) を返す

## Discount Curve (OIS Curve)
- クラス名：clsOISStepForwardCurve
- DF(date), ZeroRateCont(date), ForwardRateBetween(d1,d2) を返す
- 評価日 ON、1M、2M、3M、6M、1Y、2Y…のOISレートからブートストラップ
- ONは評価日→翌営業日、スポット前まではON横置き
- OISはスポットT+2開始
- Delay方式2営業日
- 利息計算期間と金利参照期間は一致
- 支払日は各期間終了日の2営業日後
- ACT/365
- 1年以内は満期一括、1年超は年1回払い
- 6M以降は3か月刻みの瞬間ONフォワード
- 観測点間に複数の3M区間がある場合は、直前フォワード f_prev に対して、 f_prev + 1Δf, f_prev + 2Δf, ... として、Δfをブートストラップする
- 8Mなど中途半端なテナーは無視

## Cap Volatility
- キャップボラは ATM Normal Volを前提
- Total Varianceベースで補間・Bootstrapする

