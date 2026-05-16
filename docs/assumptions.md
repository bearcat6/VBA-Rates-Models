# Assumptions

## 全体前提
- 通貨は JPY
- 金利は TONA/OISベース
- 日付は Excel Date型
- 関数の引数には、in_ を入れる
- クラス名は cls で始める
- クラス型の引数名・変数名には、オブジェクトであることが分かる接頭辞 c を使用する
  - 例：in_cCurve, in_cCalendar
- モジュール名は mdl_ で始める
- タブ間隔は半角2文字分

## Business Day Convention
- 営業日調整の既定値は Modified Following
- ただし、将来的に Following / Preceding 等へ拡張できる設計とする
- Excel関数として利用する場合を考慮し、Enumだけに依存しすぎない
- Business Day Convention は文字列でも指定可能とする
  - "F" / "Following"
  - "MF" / "ModifiedFollowing"
  - "P" / "Preceding"
- 営業日判定は、土日判定と clsHolidayCalendar の休日判定を組み合わせて行う
- 営業日加算では、基準日はカウントしない

## Tenor Date Calculation
- テナー文字列は、"+1D", "-2D", "3M", "+6M", "1Y", "-1Y" の形式を基本とする
- D は暦日加算とする
- M/Y は月加算として扱い、Y は 12M として処理する
- 月加算は DateAdd("m") ではなく、EDate ベースで行う
- 基準日が月末の場合は、原則として加算後も月末を維持する
- テナー加算のみを行う関数と、テナー加算後に営業日調整を行う関数は分ける
- テナー加算後に Business Day Convention を適用する
- Excel VBA 上での利用を前提とし、EDate / EoMonth は WorksheetFunction を使用する

## Holiday Calendar
- クラス名：clsHolidayCalendar
- IsHoliday(in_TargetDate) を返す
- Business Day Convention の判定は別関数で行う
- 当面は Tokyo holiday をコード内に保持する
- 休日データは、将来的に Holidayシート A列から読み込む方式へ切替可能な設計とする

## Discount Curve Interface
- クラス名：clsDiscountCurve
- DF(in_TargetDate), ZeroRateCont(in_TargetDate), ForwardRate(in_StartDate, in_EndDate) を返す
- ForwardRate(in_StartDate, in_EndDate) は、DF比から算出する期間フォワードレートとする
- ForwardRate は、原則として ACT/365 ベースの単利年率とする
- ForwardRate = (DF(in_StartDate) / DF(in_EndDate) - 1) / YearFraction(in_StartDate, in_EndDate)
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

## OIS Swap
- clsOISSwap は OIS スワップの商品条件を保持する商品クラスとする
- payer / receiver を商品条件として保持する
- PAYER は固定払い・変動受けを意味する
- RECEIVER は固定受け・変動払いを意味する
- FixedLegPV および FloatingLegPV は、脚単体のPVを正値として返す
- NPV は PayReceive に応じて以下のとおり計算する
  - PAYER: FloatingLegPV - FixedLegPV
  - RECEIVER: FixedLegPV - FloatingLegPV
- ParRate は、固定脚PVと変動脚PVが一致する固定金利として商品クラスで計算する
- ParRate は PayReceive に依存しない
- ACT/365 Fixed の YearFrac は共通モジュール `mdl_Common` で定義する

## OIS Step Forward Curve
- クラス名：clsOISStepForwardCurve
- DF(in_TargetDate), ZeroRateCont(in_TargetDate), ForwardRate(in_StartDate, in_EndDate) を返す
- ForwardRate(in_StartDate, in_EndDate) は、DF比から算出する期間フォワードレートとする
- ForwardRate は、原則として ACT/365 ベースの単利年率とする
- ForwardRate = (DF(in_StartDate) / DF(in_EndDate) - 1) / YearFraction(in_StartDate, in_EndDate)
- 評価日 ON、1M、2M、3M、6M、1Y、2Y…のOISレートからブートストラップ
- 6M以降は3か月刻みの瞬間ONフォワード
- 観測点間に複数の3M区間がある場合は、直前フォワード f_prev に対して、 f_prev + 1Δf, f_prev + 2Δf, ... として、Δfをブートストラップする
- 8Mなど中途半端なテナーは無視

## Cap Volatility

- 目的は、Cap取引そのものの評価ではなく、TONA 6M 等の caplet 部分を評価するための caplet normal volatility term structure を構築することである
- キャップボラは ATM Normal Vol を前提とする
- 対象は、OIS / TONA ベースの変動金利に上限が付いた商品の caplet 部分評価を想定する
- ATM Strike は、対象となる forward rate または評価対象商品の条件に応じた strike を使用する
- Caplet単体のクラスは作成しない
- Cap取引を表す clsCap も作成しない
- Quoted cap normal vol は、複数capletの合計価格を表す入力として扱う
- Bootstrapper は、quoted cap normal vol から期間別 caplet normal vol を逆算する
- Bootstrap後の caplet normal vol は clsCapVolTermStructure に格納する
- Bachelier式などの数式処理は標準モジュール `mdl_CapFormula` に置く
- ACT/365 Fixed の YearFrac は `mdl_Common.YearFracAct365F` に統一する
- Normal Vol は絶対値表記とする
  - 例：40bp normal vol = 0.0040
- 時間方向は total variance ベースで管理する
  - Total Variance = Σ normalVol_i^2 × yearFraction_i

