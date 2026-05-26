
Option Explicit

Private mExpiry As String
Private mTenor As String
Private mForward As Double
Private mT As Double
Private mAlpha As Double
Private mBeta As Double
Private mRho As Double
Private mNu As Double
Private mRMSE As Double
Private mMaxAbsError As Double
Private mIterations As Long
Private mSuccess As Boolean

Public Property Get Expiry() As String: Expiry = mExpiry: End Property
Public Property Let Expiry(ByVal v As String): mExpiry = v: End Property

Public Property Get Tenor() As String: Tenor = mTenor: End Property
Public Property Let Tenor(ByVal v As String): mTenor = v: End Property

Public Property Get Forward() As Double: Forward = mForward: End Property
Public Property Let Forward(ByVal v As Double): mForward = v: End Property

Public Property Get ExpiryYears() As Double: ExpiryYears = mT: End Property
Public Property Let ExpiryYears(ByVal v As Double): mT = v: End Property

Public Property Get Alpha() As Double: Alpha = mAlpha: End Property
Public Property Let Alpha(ByVal v As Double): mAlpha = v: End Property

Public Property Get Beta() As Double: Beta = mBeta: End Property
Public Property Let Beta(ByVal v As Double): mBeta = v: End Property

Public Property Get Rho() As Double: Rho = mRho: End Property
Public Property Let Rho(ByVal v As Double): mRho = v: End Property

Public Property Get Nu() As Double: Nu = mNu: End Property
Public Property Let Nu(ByVal v As Double): mNu = v: End Property

Public Property Get RMSE() As Double: RMSE = mRMSE: End Property
Public Property Let RMSE(ByVal v As Double): mRMSE = v: End Property

Public Property Get MaxAbsError() As Double: MaxAbsError = mMaxAbsError: End Property
Public Property Let MaxAbsError(ByVal v As Double): mMaxAbsError = v: End Property

Public Property Get Iterations() As Long: Iterations = mIterations: End Property
Public Property Let Iterations(ByVal v As Long): mIterations = v: End Property

Public Property Get Success() As Boolean: Success = mSuccess: End Property
Public Property Let Success(ByVal v As Boolean): mSuccess = v: End Property

Public Function NormalVol(ByVal Strike As Double) As Double
    NormalVol = SABR_NormalVol_Beta0(mForward, Strike, mT, mAlpha, mRho, mNu)
End Function
