
Option Explicit

'============================================================
' clsDiscountCurve
'
' Discount Curve Interface
'
' 目的：
'   ・割引カーブとして必要な最小メソッドを定義する
'   ・具体的な構築方法は、clsOISStepForwardCurve 等の実装クラスに委ねる
'
' 実装方法：
'   ・VBA では継承ではなく Implements を使用する
'   ・具体クラス側で以下を実装する
'       Private Function clsDiscountCurve_DF(...)
'       Private Function clsDiscountCurve_ZeroRateCont(...)
'       Private Function clsDiscountCurve_ForwardRate(...)
'
' 注意：
'   ・本クラスはインターフェース用途のため、計算ロジックは持たない
'   ・直接 New して使用しない
'============================================================

Public Function DF(ByVal in_TargetDate As Date) As Double
End Function

Public Function ZeroRateCont(ByVal in_TargetDate As Date) As Double
End Function

Public Function ForwardRate(ByVal in_StartDate As Date, _
                            ByVal in_EndDate As Date) As Double
End Function
