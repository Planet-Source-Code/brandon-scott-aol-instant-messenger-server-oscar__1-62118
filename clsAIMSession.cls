VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAIMSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_iIndex As Integer
Private m_iWarningLevel As Integer
Private m_sScreenName As String
Private m_sFormattedScreenName As String
Private m_sEmailAddress As String
Private m_sProfile As String
Private m_sProfileEncoding As String
Private m_sAwayMessage As String
Private m_sAwayMessageEncoding As String
Private m_bSignedOn As Boolean
Private m_sCapabilities As String
Private m_sShortCapabilities As String
Private m_dtSignonTime As Date
Private m_dtIdleTime As Date
Private m_sBartInfo As String
Private m_sSubscriptions As String
Private m_sClientVersion As String
Private m_lClientID As Long
Private m_lClientMajorVersion As Long
Private m_lClientMinorVersion As Long
Private m_lClientPointVersion As Long
Private m_lClientBuildNumber As Long
Private m_dClientDistributionChannel As Double
Private m_sIPAddress As String
Private m_sPassword As String
Private m_sCookie As String
Private m_lAuthSocket As Long
Private m_sLoginTicket As String
Private m_bAuthorized As Boolean
Private m_lID As Long
Private m_lSignonTimestamp As Long
Private m_lUserClass As Long
Private m_sCerts As String
Private m_lServiceID As Long
Private m_colServiceCollection As Collection


Public Property Get Index() As Integer

    Index = m_iIndex

End Property

Public Property Let Index(ByVal iIndex As Integer)

    m_iIndex = iIndex

End Property

Public Property Get Cookie() As String

    Cookie = m_sCookie

End Property

Public Property Let Cookie(ByVal sCookie As String)

    m_sCookie = sCookie

End Property

Public Property Get WarningLevel() As Integer

    WarningLevel = m_iWarningLevel

End Property

Public Property Let WarningLevel(ByVal iWarningLevel As Integer)

    m_iWarningLevel = iWarningLevel

End Property

Public Property Get ScreenName() As String

    ScreenName = m_sScreenName

End Property

Public Property Let ScreenName(ByVal sScreenName As String)

    m_sScreenName = sScreenName

End Property

Public Property Get Password() As String

    Password = m_sPassword

End Property

Public Property Let Password(ByVal sPassword As String)

    m_sPassword = sPassword

End Property


Public Property Get FormattedScreenName() As String

    FormattedScreenName = m_sFormattedScreenName

End Property

Public Property Let FormattedScreenName(ByVal sFormattedScreenName As String)

    m_sFormattedScreenName = sFormattedScreenName

End Property

Public Property Get EmailAddress() As String

    EmailAddress = m_sEmailAddress

End Property

Public Property Let EmailAddress(ByVal sEmailAddress As String)

    m_sEmailAddress = sEmailAddress

End Property

Public Property Get Profile() As String

    Profile = m_sProfile

End Property

Public Property Let Profile(ByVal sProfile As String)

    m_sProfile = sProfile

End Property

Public Property Get ProfileEncoding() As String

    ProfileEncoding = m_sProfileEncoding

End Property

Public Property Let ProfileEncoding(ByVal sProfileEncoding As String)

    m_sProfileEncoding = sProfileEncoding

End Property

Public Property Get AwayMessage() As String

    AwayMessage = m_sAwayMessage

End Property

Public Property Let AwayMessage(ByVal sAwayMessage As String)

    m_sAwayMessage = sAwayMessage

End Property

Public Property Get AwayMessageEncoding() As String

    AwayMessageEncoding = m_sAwayMessageEncoding

End Property

Public Property Let AwayMessageEncoding(ByVal sAwayMessageEncoding As String)

    m_sAwayMessageEncoding = sAwayMessageEncoding

End Property

Public Property Get SignedOn() As Boolean

    SignedOn = m_bSignedOn

End Property

Public Property Let SignedOn(ByVal bSignedOn As Boolean)

    m_bSignedOn = bSignedOn

End Property

Public Property Get Capabilities() As String

    Capabilities = m_sCapabilities

End Property

Public Property Let Capabilities(ByVal sCapabilities As String)

    m_sCapabilities = sCapabilities

End Property

Public Property Get ShortCapabilities() As String

    ShortCapabilities = m_sShortCapabilities

End Property

Public Property Let ShortCapabilities(ByVal sShortCapabilities As String)

    m_sShortCapabilities = sShortCapabilities

End Property

Public Property Get SignonTime() As Date

    SignonTime = m_dtSignonTime

End Property

Public Property Let SignonTime(ByVal dtSignonTime As Date)

    m_dtSignonTime = dtSignonTime

End Property

Public Property Get IdleTime() As Date

    IdleTime = m_dtIdleTime

End Property

Public Property Let IdleTime(ByVal dtIdleTime As Date)

    m_dtIdleTime = dtIdleTime

End Property


Public Property Get BartInfo() As String

    BartInfo = m_sBartInfo

End Property

Public Property Let BartInfo(ByVal sBartInfo As String)

    m_sBartInfo = sBartInfo

End Property

Public Property Get Subscriptions() As String

    Subscriptions = m_sSubscriptions

End Property

Public Property Let Subscriptions(ByVal sSubscriptions As String)

    m_sSubscriptions = sSubscriptions

End Property

Public Property Get ClientVersion() As String

    ClientVersion = m_sClientVersion

End Property

Public Property Let ClientVersion(ByVal sClientVersion As String)

    m_sClientVersion = sClientVersion

End Property

Public Property Get ClientID() As Long

    ClientID = m_lClientID

End Property

Public Property Let ClientID(ByVal lClientID As Long)

    m_lClientID = lClientID

End Property

Public Property Get ClientMajorVersion() As Long

    ClientMajorVersion = m_lClientMajorVersion

End Property

Public Property Let ClientMajorVersion(ByVal lClientMajorVersion As Long)

    m_lClientMajorVersion = lClientMajorVersion

End Property

Public Property Get ClientMinorVersion() As Long

    ClientMinorVersion = m_lClientMinorVersion

End Property

Public Property Let ClientMinorVersion(ByVal lClientMinorVersion As Long)

    m_lClientMinorVersion = lClientMinorVersion

End Property

Public Property Get ClientPointVersion() As Long

    ClientPointVersion = m_lClientPointVersion

End Property

Public Property Let ClientPointVersion(ByVal lClientPointVersion As Long)

    m_lClientPointVersion = lClientPointVersion

End Property

Public Property Get ClientBuildNumber() As Long

    ClientBuildNumber = m_lClientBuildNumber

End Property

Public Property Let ClientBuildNumber(ByVal lClientBuildNumber As Long)

    m_lClientBuildNumber = lClientBuildNumber

End Property

Public Property Get ClientDistributionChannel() As Double

    ClientDistributionChannel = m_dClientDistributionChannel

End Property

Public Property Let ClientDistributionChannel(ByVal dClientDistributionChannel As Double)

    m_dClientDistributionChannel = dClientDistributionChannel

End Property

Public Property Get IPAddress() As String

    IPAddress = m_sIPAddress

End Property

Public Property Let IPAddress(ByVal sIPAddress As String)

    m_sIPAddress = sIPAddress

End Property


Public Property Get AuthSocket() As Long

    AuthSocket = m_lAuthSocket

End Property

Public Property Let AuthSocket(ByVal lAuthSocket As Long)

    m_lAuthSocket = lAuthSocket

End Property

Public Property Get LoginTicket() As String

    LoginTicket = m_sLoginTicket

End Property

Public Property Let LoginTicket(ByVal sLoginTicket As String)

    m_sLoginTicket = sLoginTicket

End Property

Public Property Get Authorized() As Boolean

    Authorized = m_bAuthorized

End Property

Public Property Let Authorized(ByVal bAuthorized As Boolean)

    m_bAuthorized = bAuthorized

End Property


Public Property Get ID() As Long

    ID = m_lID

End Property

Public Property Let ID(ByVal lID As Long)

    m_lID = lID

End Property


Public Property Get SignonTimestamp() As Long

    SignonTimestamp = m_lSignonTimestamp

End Property

Public Property Let SignonTimestamp(ByVal lSignonTimestamp As Long)

    m_lSignonTimestamp = lSignonTimestamp

End Property

Public Property Get UserClass() As Long

    UserClass = m_lUserClass

End Property

Public Property Let UserClass(ByVal lUserClass As Long)

    m_lUserClass = lUserClass

End Property


Public Property Get Certs() As String

    Certs = m_sCerts

End Property

Public Property Let Certs(ByVal sCerts As String)

    m_sCerts = sCerts

End Property

Public Function AddService(lGroup As Long, sCookie As String) As clsAIMService
    
    Dim oAIMService As clsAIMService
    Set oAIMService = New clsAIMService
    m_lServiceID = m_lServiceID + 1
    oAIMService.Group = lGroup
    oAIMService.Cookie = sCookie
    m_colServiceCollection.Add oAIMService, CStr(m_lServiceID)
    Set AddService = oAIMService
    Set oAIMService = Nothing

End Function

Private Sub Class_Initialize()
    
    Set m_colServiceCollection = New Collection
    
End Sub

Private Sub Class_Terminate()

    Set m_colServiceCollection = Nothing

End Sub

Public Property Get Services() As Collection

    Set Services = m_colServiceCollection

End Property
