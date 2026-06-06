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
- ForwardParRate(in_Expiry, in_Tenor) は、Forward Starting OIS Swap の par rate を返す
- ForwardParRate の in_Expiry / in_Tenor は tenor文字列で指定する
  - 例：ForwardParRate("6M", "5Y")
- ForwardParRate の開始日は SpotDate + Expiry、終了日は開始日 + Tenor とする
- ForwardParRate は PayReceive に依存しない
- ForwardParRate は、固定脚PVと変動脚PVが一致する固定金利とする
- ForwardParRate = FloatingLegPV / Annuity
- FloatingLegPV = Σ (DF(periodStart_i) / DF(periodEnd_i) - 1) × DF(paymentDate_i)
- Annuity = Σ YearFraction(periodStart_i, periodEnd_i) × DF(paymentDate_i)
- 固定脚の YearFrac は ACT/365 とする
- 固定脚の支払日は各期間終了日の PaymentLag 営業日後とする
- 固定脚は OIS共通前提に合わせ、1年以内は満期一括、1年超は年1回払いとする
- ForwardParRateByDates(in_StartDate, in_EndDate) は、開始日・終了日を直接指定する日付指定版とする

## OIS Zero Linear Curve

- クラス名：clsOISZeroLinearCurve
- `DF(in_TargetDate)`, `ZeroRateCont(in_TargetDate)`, `ForwardRate(in_StartDate, in_EndDate)` を返す
- `clsDiscountCurve` を Implements する
- 入力レートは連続複利ゼロレートとする
- OIS quote からのブートストラップは行わない
- tenor / zero rate 配列を受け取り、各 tenor に対応する満期日をゼロレート点として保持する
- `"ON"` / `"O/N"` は、評価日翌営業日のゼロレート点として扱う
- `"1M"`, `"2M"`, `"3M"`, `"6M"`, `"1Y"`, `"2Y"` 等は、SpotDate + Tenor を Business Day Convention で調整した日付として扱う
- 補間対象は、連続複利ゼロレートとする
- 補間軸は、評価日から対象日までの ACT/365 年数とする
- 補間方法は、ゼロレートの直線補間とする
- DF は `DF(0,T) = Exp(-ZeroRateCont(T) * T)` で計算する
- ForwardRate は、DF 比から算出する期間フォワードレートとする
- ForwardRate は、原則として ACT/365 ベースの単利年率とする
- ForwardRate = `(DF(in_StartDate) / DF(in_EndDate) - 1) / YearFraction(in_StartDate, in_EndDate)`
- 最初のゼロレート点より手前は、最初のゼロレートを横置きする
- 最終ゼロレート点より後は、原則として外挿しない
- ただし、支払日計算への対応として、最終点 + PaymentLag までは最後のゼロレートを横置きする
- それを超える日付はエラーとする

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

## Swaption Volatility

- 目的は、スワップション取引そのものの評価ではなく、CMS convexity adjustment や CMS spread swap 等で参照する ATM swaption volatility surface を構築・参照することである
- スワップションの商品クラスは当面作成しない
- ATM swaption volatility は、expiry × underlying swap tenor のマトリックスとして入力する
- Expiry は年数 Double または tenor文字列で指定可能とする
  - 例：0.5, 1, 1.5, "6M", "1Y", "18M"
- Tenor は文字列指定を基本とする
  - 例："1Y", "5Y", "10Y", "20Y"
- 入力ボラティリティは、当面 ATM Normal Vol を前提とする
- Normal Vol は絶対値表記とする
  - 例：40bp normal vol = 0.0040
- Tenor方向は variance = vol^2 を線形補間する
- Expiry方向は total variance = vol^2 × expiry を線形補間する
- 範囲外の外挿は原則エラーとする
- ただし、明示的に allowFlatExtrapolation = True とした場合のみ、最短・最長グリッドの端点ボラを使用する
- 本サーフェスは、CMS関連評価のための実務的な補間サーフェスであり、完全な無裁定スワップション・ボラティリティモデルではない
- Black Vol / Shifted Black Vol を扱う場合は、別サーフェスまたは volatility type を明示的に分ける

## Hull-White 1F

- 対象モデルは 1-factor Hull-White model とする。
- 平均回帰パラメータ `a` とボラティリティ `sigma` は定数とする。
- 初期実装では、`a` は外部から固定値として与える。
- `sigma` は ATM normal swaption volatility へフィットする。
- 実装では shifted short-rate representation を使用する。
- `theta(t)` は直接入力せず、初期 discount curve と整合するように扱う。
- 時間軸は評価日からの year fraction `T` とする。
- Hull-White 関連クラスは curve object に対して `DF_T(T)` と `InstantaneousForward(T)` を要求する。
