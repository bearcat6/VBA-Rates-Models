# Assumptions

## 全体前提
- 通貨はJPY
- 金利はTONA/OISベース
- 日付はExcel Date型
- 営業日調整はModified Following
- 休日カレンダーはHolidayシートA列

## カーブ
- OISカーブは COISStepForwardCurve を使用
- DF(date), ZeroRateCont(date), ForwardRateBetween(d1,d2) を返す

## ボラティリティ
- キャップボラはATM Normal Volを前提
- Total Varianceベースで補間・Bootstrapする
