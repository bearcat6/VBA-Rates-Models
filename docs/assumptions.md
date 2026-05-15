# Assumptions

## 全体前提
- 通貨は JPY
- 金利は TONA/OISベース
- 日付は Excel Date型
- 関数の引数には、in_ を入れる
- クラス名は cls で始める
- クラス型の引数名・変数名には、オブジェクトであることが分かる接頭辞 c を使用する
例：in_cCurve や in_cCalendar
- モジュール名は、mdl_で始まる

## Business Day Convention
- 営業日調整の既定値は Modified Following
- ただし、将来的に Following / Preceding 等へ拡張できる設計とする
- Excel関数として利用する場合を考慮し、Enumだけに依存しすぎない

## Holiday Calendar
- クラス名：clsHolidayCalendar
- IsHoliday(in_TargetDate) を返す
- Business Day Convention の判定は別関数で行う
- 当面は Tokyo holiday をコード内に保持する
- 将来的に Holidayシート A列から読み込む方式へ切替可能な設計とする

## Discount Curve Interface
- クラス名：clsDiscountCurve
- DF(in_TargetDate), ZeroRateCont(in_TargetDate), ForwardRate(in_StartDate, in_EndDate) を返す
- ForwardRate(in_StartDate, in_EndDate) は、DF比から算出する期間フォワードレートとする
  ForwardRate は、原則として ACT/365 ベースの単利年率とする
- clsDiscountCurve は共通インターフェース的な位置づけとする
- 実際の構築手法は、clsOISStepForwardCurve 等の具体クラスで実装する
- 将来的に、ゼロレート直線補間型などの具体クラスを追加できる設計とする

## OIS共通前提
- ONは評価日→翌営業日、スポット前まではON横置き
- OISはスポットT+2開始
- Delay方式2営業日
- 利息計算期間と金利参照期間は一致
- 支払日は各期間終了日の2営業日後
- ACT/365
- 1年以内は満期一括、1年超は年1回払い

## OIS Step Forward Curve
- クラス名：clsOISStepForwardCurve
- DF(in_TargetDate), ZeroRateCont(in_TargetDate), ForwardRate(in_StartDate, in_EndDate) を返す
- ForwardRate(in_StartDate, in_EndDate) は、DF比から算出する期間フォワードレートとする
  ForwardRate は、原則として ACT/365 ベースの単利年率とする
- 評価日 ON、1M、2M、3M、6M、1Y、2Y…のOISレートからブートストラップ
- 6M以降は3か月刻みの瞬間ONフォワード
- 観測点間に複数の3M区間がある場合は、直前フォワード f_prev に対して、 f_prev + 1Δf, f_prev + 2Δf, ... として、Δfをブートストラップする
- 8Mなど中途半端なテナーは無視

## Cap Volatility
- キャップボラは ATM Normal Volを前提
- 対象は TONA変動金利にキャップが付いた商品評価を想定
- ATM Strike は、対象Capの forward par rate を基本とする
- Total Varianceベースで補間・Bootstrapする
- Caplet単体のクラスは原則作らない
- Bachelier式などの数式処理は標準モジュールに置く

