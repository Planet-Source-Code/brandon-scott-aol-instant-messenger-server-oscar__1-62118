VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMyAIMServer 
   AutoShowChildren=   0   'False
   BackColor       =   &H00FFFFFF&
   Caption         =   "My AIM Server - Created by Xeon Productions"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16005
   Icon            =   "mdiMyAIMServer.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ilsToolbarImages 
      Left            =   1920
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483648
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMyAIMServer.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMyAIMServer.frx":19A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMyAIMServer.frx":267E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMyAIMServer.frx":36D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMyAIMServer.frx":4722
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMyAIMServer.frx":5774
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMyAIMServer.frx":67C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMyAIMServer.frx":7818
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMyAIMServer.frx":886A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMyAIMServer.frx":98BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMyAIMServer.frx":A90E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMyAIMServer.frx":B960
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMyAIMServer.frx":C9B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMyAIMServer.frx":DA04
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMyAIMServer.frx":EA56
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMyAIMServer.frx":FAA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMyAIMServer.frx":10042
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMyAIMServer.frx":105DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbServerStatus 
      Align           =   1  'Align Top
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   330
      Width           =   16005
      _ExtentX        =   28231
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "6:13 PM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "8/7/2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23045
            Text            =   "Server is currently stopped."
            TextSave        =   "Server is currently stopped."
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrServerControls 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16005
      _ExtentX        =   28231
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ilsToolbarImages"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Server control"
            Object.ToolTipText     =   "Server Control"
            Object.Tag             =   "offline"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Connections"
            ImageIndex      =   13
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Send Message"
            ImageIndex      =   15
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Send Popup Message"
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Account Editor"
            ImageIndex      =   6
            Style           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Server Log"
            ImageIndex      =   7
            Style           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Error Log"
            ImageIndex      =   8
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Server Configuration"
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Chat Moderator"
            ImageIndex      =   18
            Style           =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Banned IP Addresses"
            ImageIndex      =   17
            Style           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Server Statistics"
            ImageIndex      =   12
            Style           =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Web Server"
            ImageIndex      =   14
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin MyAIMServer.AIMServer AuthServer 
      Left            =   120
      Top             =   720
      _extentx        =   847
      _extenty        =   847
   End
   Begin MyAIMServer.AIMServer BOSServer 
      Left            =   720
      Top             =   720
      _extentx        =   847
      _extenty        =   847
   End
   Begin MyAIMServer.AIMServer AddonServicesServer 
      Left            =   1320
      Top             =   720
      _extentx        =   847
      _extenty        =   847
   End
End
Attribute VB_Name = "mdiMyAIMServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub AddonServicesServer_Connected(Index As Integer, RemoteHost As String)
    Debug.Print "AddonConnect", Index, RemoteHost
    AddonServicesServer.SendData Index, 0, 1, FlapVersion
End Sub

Private Sub AddonServicesServer_DataArrival(Index As Integer, Data As String)
    Dim oAIMUserTemp As clsAIMSession
    Dim oAIMUser As clsAIMSession
    Dim oAIMService As clsAIMService
    Dim sCookie As String
    If 10 > Len(Data) Then Exit Sub
    If Asc(Mid$(Data, 2, 1)) = 1 Then
        If Len(Mid$(Data, 7)) > 4 Then
            sCookie = GetTLV(6, Mid$(Data, 11))
            For Each oAIMUserTemp In oAIMSessionManager
                For Each oAIMService In oAIMUserTemp.Services
                    If oAIMService.Cookie = sCookie Then
                        oAIMService.Index = Index
                        AddonServicesServer.SendData Index, 0, 2, ServiceHostOnline
                    End If
                Next oAIMService
            Next oAIMUserTemp
        End If
    ElseIf Asc(Mid$(Data, 2, 1)) = 2 Then
        Debug.Print "Addon_DataArrival", StringToHexArray(Mid$(Data, 7))
    End If
End Sub

Private Sub AuthServer_Connected(Index As Integer, RemoteHost As String)
    LogMsg RemoteHost & " has connected.", , True
    AuthServer.SendData Index, 0, 1, FlapVersion
End Sub

Private Sub AuthServer_DataArrival(Index As Integer, Data As String)

    Dim sScreenName As String
    Dim sLoginTicket As String
    Dim sPasswordHash As String
    Dim oAIMUser As clsAIMSession
    
    'We don't want any crap
    If 10 > Len(Data) Then Exit Sub
    
    Select Case Asc(Mid(Data, 2, 1))
        Case 2
            Select Case Mid(Data, 7, 4)
            
                '==============================================================================
                Case ChrB("00 17 00 06") 'BUCP Request Challenge
                '==============================================================================
                
                    sScreenName = GetTLV(1, Mid$(Data, 17))
                    For Each oAIMUser In oAIMSessionManager
                        If oAIMUser.ScreenName = TrimData(sScreenName) Then
                            Debug.Print "Existing Connection", "socket=" & oAIMUser.Index, "name=" & oAIMUser.ScreenName, "SignedOn=" & oAIMUser.SignedOn
                            If oAIMUser.Index = 0 Or oAIMUser.SignedOn = False Then
                                oAIMSessionManager.Remove TrimData(sScreenName)
                            End If
                        End If
                    Next oAIMUser
                    
                    Select Case GetAccountStatus(sScreenName)
                        Case LoginStateDeleted
                        
                            LogMsg sScreenName & ": Deleted account.", vbRed
                            AuthServer.SendData Index, 0, 2, BucpReply(sScreenName, , , , , , True, "http://www.aim.aol.com/errors/DELETED_ACCT.html?ccode=us&lang=en", 8)
                            AuthServer.SendData Index, 0, 4, ""
                            Exit Sub
                            
                        Case LoginStateInvalid
                        
                            LogMsg sScreenName & ": Bad login attempt.", vbRed
                            AuthServer.SendData Index, 0, 2, BucpReply(sScreenName, , , , , , True, "http://www.aol.com?ccode=us&lang=en", 4)
                            AuthServer.SendData Index, 0, 4, ""
                            Exit Sub
                            
                        Case LoginStateInvalidPassword
                        
                            LogMsg sScreenName & ": Has provided an invalid password.", vbRed
                            AuthServer.SendData Index, 0, 2, BucpReply(sScreenName, , , , , , True, "http://www.aim.aol.com/errors/MISMATCH_PASSWD.html?ccode=us&lang=en", 5)
                            AuthServer.SendData Index, 0, 4, ""
                            Exit Sub
                            
                        Case LoginStateSuspended
                        
                            LogMsg sScreenName & ": Suspended account attempted to login.", vbRed
                            AuthServer.SendData Index, 0, 2, BucpReply(sScreenName, , , , , , True, "http://www.aol.com?ccode=us&lang=en", 17)
                            AuthServer.SendData Index, 0, 4, ""
                            Exit Sub
                            
                        Case LoginStateUnregistered
                        
                            LogMsg sScreenName & ": Unregistered login attempt.", vbRed
                            AuthServer.SendData Index, 0, 2, BucpReply(sScreenName, , , , , , True, "http://www.aim.aol.com/errors/UNREGISTERED_SCREENNAME.html?ccode=us&lang=en", 1)
                            AuthServer.SendData Index, 0, 4, ""
                            Exit Sub
                            
                        Case LoginStateGood ' this doesnt mean the login is good, just means they get a ticket!

                            'Generate a new login ticket for this kind soul
                            sLoginTicket = GRTicket
                            'Swap it into are session manager
                            oAIMSessionManager.Add TrimData(sScreenName), Index, sLoginTicket
                            'Lets send it to them!
                            AuthServer.SendData Index, 0, 2, BucpChallenge(sLoginTicket)
                    
                    End Select
                    
                '==============================================================================
                Case ChrB("00 17 00 02") 'BUCP Query (Login)
                '==============================================================================
                
                    sScreenName = GetTLV(1, Mid$(Data, 17))
                    Set oAIMUser = oAIMSessionManager.Item(TrimData(sScreenName))
                    'Make sure object isn't null.
                    If Not oAIMUser Is Nothing Then
                        'Generate a login cookie
                        oAIMUser.Cookie = GRCookie
                        'Get Password Hash
                        sPasswordHash = GetTLV(37, Mid$(Data, 17))
                        'Check Login
                        If CheckLogin(sScreenName, sPasswordHash, oAIMUser.LoginTicket) = True Then
                            LogMsg sScreenName & ": was authenticated successfully.", &H8000&
                            Call SetupAccount(oAIMUser)
                            AuthServer.SendData Index, 0, 2, BucpReply(oAIMUser.FormattedScreenName, oAIMUser.Cookie, oAIMUser.EmailAddress)
                            AuthServer.SendData Index, 0, 4, ""
                            Exit Sub
                        Else
                            AuthServer.SendData Index, 0, 2, BucpReply(sScreenName, , , , , , True, "http://www.aim.aol.com/errors/MISMATCH_PASSWD.html?ccode=us&lang=en", 5)
                            AuthServer.SendData Index, 0, 4, ""
                            Exit Sub
                        End If
                    Else
                        'uh...I don't know?!!
                        AuthServer.SendData Index, 0, 2, BucpReply(sScreenName, , , , , , True, "http://www.aol.com/", 666)
                        AuthServer.SendData Index, 0, 4, ""
                        Exit Sub
                    End If
                    
            End Select
        Case 4
            Call AuthServer.CloseSocket(Index)
    End Select
End Sub

Private Sub AuthServer_Disconnected(Index As Integer)
    Dim oAIMUser As clsAIMSession
    For Each oAIMUser In oAIMSessionManager
        If oAIMUser.Authorized = False And oAIMUser.AuthSocket = Index Then
            oAIMSessionManager.Remove TrimData(oAIMUser.ScreenName)
        End If
    Next oAIMUser
End Sub

Private Sub BosServer_Connected(Index As Integer, RemoteHost As String)
    BOSServer.SendData Index, 0, 1, FlapVersion
End Sub

Private Sub BosServer_DataArrival(Index As Integer, Data As String)
    
    'On Error GoTo ErrBosServer_DataArrival
    
    'Multi-Use Parsing Variables
    Dim L1 As Long, L2 As Long, L3 As Long, L4 As Long
    Dim S1 As String, S2 As String, S3 As String, S4 As String, S5 As String
    Dim D1 As Double, D2 As Double, D3 As Double, D4 As Double
    Dim oAIMUserTemp As clsAIMSession
    Dim oAIMUser As clsAIMSession
    Dim oICBMParser As New clsICBMPacket
    Dim oBuddylistAddParser As New clsBinaryStream
    Dim oBinaryReader As New clsBinaryStream
    Dim dblFeedbagTimestamp As Double
    Dim lngFeedbagItems As Long
    Dim sCookie As String
    Dim lngFamily As Long
    Dim lngSubType As Long
    Dim lngFlags As Long
    Dim dblRequestID As Double
    Dim sBuffer As String
    Dim lType As Long, lLength As Long, sValue As String

    If 10 > Len(Data) Then Exit Sub
    
    Select Case Asc(Mid$(Data, 2, 1))
        Case 1
        
            If Len(Mid$(Data, 7)) > 4 Then
                sCookie = GetTLV(6, Mid$(Data, 11))
                For Each oAIMUserTemp In oAIMSessionManager
                    If oAIMUserTemp.Cookie = sCookie Then
                        oAIMUserTemp.Index = Index
                        BOSServer.SendData Index, 0, 2, ServiceHostOnline
                    End If
                Next oAIMUserTemp
            End If
            
        Case 2
            
            lngFamily = GetWord(Mid$(Data, 7, 2))
            lngSubType = GetWord(Mid$(Data, 9, 2))
            lngFlags = GetWord(Mid$(Data, 11, 2))
            dblRequestID = GetDWord(Mid$(Data, 13, 4))
            
            For Each oAIMUserTemp In oAIMSessionManager
                If oAIMUserTemp.Index = Index Then
                    Set oAIMUser = oAIMUserTemp
                End If
            Next oAIMUserTemp
        
        
            Select Case Mid(Data, 7, 4)
            
                '==============================================================================
                Case ChrB("00 07 00 02") 'Admin Request Info
                '==============================================================================
                
                    oBinaryReader.LoadBuffer Mid$(Data, 17)
                    L1 = oBinaryReader.Read16
                    If L1 = 1 Then
                        BOSServer.SendData Index, dblRequestID, 2, AdminSendInfo(1, oAIMUser.FormattedScreenName)
                    ElseIf L1 = 17 Then
                        BOSServer.SendData Index, dblRequestID, 2, AdminSendInfo(17, oAIMUser.EmailAddress)
                    ElseIf L1 = 19 Then
                        'BOSServer.SendData Index, dblRequestID, 2, AdminSendInfo(1, oAIMUser.re)
                    End If
                
                '==============================================================================
                Case ChrB("00 07 00 04") 'Admin Update Info
                '==============================================================================
                
                    oBinaryReader.LoadBuffer Mid$(Data, 17)
                    sBuffer = ChrB("00 07 00 05 00 00 00 00 00 00")
                    Do Until oBinaryReader.IsEnd
                        lType = oBinaryReader.Read16
                        sValue = oBinaryReader.Read16String
                        If lType = 1 Then
                            If TrimData(sValue) = oAIMUser.ScreenName And Len(sValue) <= 18 Then
                                oAIMUser.FormattedScreenName = Trim(sValue)
                                sBuffer = sBuffer & Word(3) & Word(1) & PutTLV(lType, Trim(sValue))
                                Call UpdateUserStatus(oAIMUser)
                            Else
                                sBuffer = sBuffer & Word(3) & Word(3) & PutTLV(lType, "") & PutTLV(4, "http://www.xeons.net/") & PutTLV(8, Word(11))
                            End If
                        ElseIf lType = 17 Then
                            If InStr(1, sValue, "@") > 0 And InStr(1, sValue, ".") > 0 Then
                                sBuffer = sBuffer & Word(3) & Word(1) & PutTLV(lType, TrimData(sValue))
                                oAIMUser.EmailAddress = TrimData(sValue)
                                BOSServer.SendData Index, dblRequestID, 2, sBuffer
                            Else
                                sBuffer = sBuffer & Word(3) & Word(3) & PutTLV(lType, "") & PutTLV(4, "http://www.xeons.net/") & PutTLV(8, Word(8))
                            End If
                        End If
                    Loop
                    UpdateAdminAccountInfo oAIMUser
                    BOSServer.SendData Index, dblRequestID, 2, sBuffer
                    BOSServer.SendData Index, 0, 2, ServiceNickInfoReply(oAIMUser.FormattedScreenName)
                    
                '==============================================================================
                Case ChrB("00 01 00 02") 'Service Client Ready
                '==============================================================================
                
                    LogMsg oAIMUser.FormattedScreenName & " has signed on successfully.", vbBlue
                    oAIMUser.SignedOn = True
                    Call UpdateUserStatus(oAIMUser)
                    
                '==============================================================================
                Case ChrB("00 01 00 04") 'Service Request
                '==============================================================================
                
                    oBinaryReader.LoadBuffer Mid$(Data, 17)
                    S1 = GRCookie               'Generate a cookie
                    L1 = oBinaryReader.Read16   'Group ID
                    oAIMUser.AddService L1, S1  'Add Service
                    BOSServer.SendData Index, dblRequestID, 2, ServiceReponse("255.255.255.255", 5192, L1, S1)
                    
                    'Debug.Print "Service Request: " & StringToHexArray(Mid$(Data, 17))
                    
                '==============================================================================
                Case ChrB("00 01 00 17") 'Service Client Versions
                '==============================================================================
                
                    BOSServer.SendData Index, 0, 2, ServiceHostVersions
                    BOSServer.SendData Index, 0, 2, ServiceMotd
                    
                '==============================================================================
                Case ChrB("00 01 00 06") 'Service Request Rate Params
                '==============================================================================
                
                    BOSServer.SendData Index, dblRequestID, 2, ServiceRateParams
                    
                '==============================================================================
                Case ChrB("00 01 00 08") 'Service Rate Add Param Sub
                '==============================================================================
                
                    'Debug.Print "AddParamSub", StringToHexArray(Mid$(Data, 17))
                
                '==============================================================================
                Case ChrB("00 01 00 0E") 'Service Request Nick Info
                '==============================================================================
                
                    BOSServer.SendData Index, dblRequestID, 2, ServiceNickInfoReply(oAIMUser.FormattedScreenName)
                
                '==============================================================================
                Case ChrB("00 13 00 02") 'Feedbag Request Rights
                '==============================================================================
                    
                    BOSServer.SendData Index, dblRequestID, 2, FeedbagRightsReply
                
                '==============================================================================
                Case ChrB("00 13 00 05") 'Feedbag Request If Modified
                '==============================================================================
                    
                    oBinaryReader.LoadBuffer Mid$(Data, 17)
                    dblFeedbagTimestamp = oBinaryReader.Read32
                    lngFeedbagItems = oBinaryReader.Read16
                    If FeedbagCheckIfNew(oAIMUser, dblFeedbagTimestamp, lngFeedbagItems) = True Then
                        BOSServer.SendData Index, dblRequestID, 2, FeedbagBuddylist(GetFeedbagData(oAIMUser)) & DWord(dblFeedbagTimestamp + 2588)
                    Else
                        BOSServer.SendData Index, dblRequestID, 2, FeedbagReplyNotModified(DWord(dblFeedbagTimestamp) & Word(lngFeedbagItems))
                    End If
                    
                '==============================================================================
                Case ChrB("00 13 00 04") 'Feedbag Request
                '==============================================================================
                
                    BOSServer.SendData Index, dblRequestID, 2, FeedbagBuddylist(GetFeedbagData(oAIMUser))
                    
                '==============================================================================
                Case ChrB("00 13 00 07") 'Feedbag Use
                '==============================================================================
                    
                    LogMsg oAIMUser.FormattedScreenName & " has received there buddylist."
                    
                '==============================================================================
                Case ChrB("00 13 00 08") 'Feedbag Add
                '==============================================================================
                
                    oBuddylistAddParser.LoadBuffer Mid$(Data, 17)
                    Do Until oBuddylistAddParser.IsEnd
                        sBuffer = sBuffer & Word(FeedbagAddItem(oAIMUser, oBuddylistAddParser.Read16String, oBuddylistAddParser.Read16, oBuddylistAddParser.Read16, oBuddylistAddParser.Read16, oBuddylistAddParser.Read16String))
                    Loop
                    BOSServer.SendData Index, dblRequestID, 2, FeedbagStatusReply(sBuffer)
                
                '==============================================================================
                Case ChrB("00 13 00 09") 'Feedbag Update
                '==============================================================================
                
                    oBuddylistAddParser.LoadBuffer Mid$(Data, 17)
                    Do Until oBuddylistAddParser.IsEnd
                        sBuffer = sBuffer & Word(FeedbagUpdateItem(oAIMUser, oBuddylistAddParser.Read16String, oBuddylistAddParser.Read16, oBuddylistAddParser.Read16, oBuddylistAddParser.Read16, oBuddylistAddParser.Read16String))
                    Loop
                    
                    BOSServer.SendData Index, dblRequestID, 2, FeedbagStatusReply(sBuffer)
                
                '==============================================================================
                Case ChrB("00 13 00 0A") 'Feedbag Delete Item
                '==============================================================================
                
                    oBuddylistAddParser.LoadBuffer Mid$(Data, 17)
                    Do Until oBuddylistAddParser.IsEnd
                        sBuffer = sBuffer & Word(FeedbagDeleteItem(oAIMUser, oBuddylistAddParser.Read16String, oBuddylistAddParser.Read16, oBuddylistAddParser.Read16, oBuddylistAddParser.Read16, oBuddylistAddParser.Read16String))
                    Loop
                    BOSServer.SendData Index, dblRequestID, 2, FeedbagStatusReply(sBuffer)
                    
                '==============================================================================
                Case ChrB("00 13 00 12") 'Done modifing feedbag!
                '==============================================================================
                
                    Call FeedbagUpdateDatabase(oAIMUser)
                    Call UpdateUserStatus(oAIMUser)
                    
                '==============================================================================
                Case ChrB("00 02 00 02") 'Location Rights Request
                '==============================================================================
                
                    BOSServer.SendData Index, dblRequestID, 2, LocationRightsReply
                    
                '==============================================================================
                Case ChrB("00 02 00 04") 'Location Set Info
                '==============================================================================
                
                    oBinaryReader.LoadBuffer Mid$(Data, 17)
                    Do Until oBinaryReader.IsEnd
                        lType = oBinaryReader.Read16
                        sValue = oBinaryReader.Read16String
                        Select Case lType
                            Case 1
                                oAIMUser.ProfileEncoding = sValue
                            Case 2
                                oAIMUser.Profile = sValue
                            Case 3
                                oAIMUser.AwayMessageEncoding = sValue
                            Case 4
                                oAIMUser.AwayMessage = sValue
                            Case 5
                                oAIMUser.Capabilities = sValue
                            Case 6
                                oAIMUser.Certs = sValue
                        End Select
                    Loop
                    Call UpdateUserStatus(oAIMUser)
                    
                '==============================================================================
                Case ChrB("00 02 00 05") 'Locate Get Info (Old)
                '==============================================================================
                
                    oBinaryReader.LoadBuffer Mid$(Data, 17)
                    L1 = oBinaryReader.Read16
                    S1 = oBinaryReader.Read08String
                    For Each oAIMUserTemp In oAIMSessionManager
                        If TrimData(oAIMUserTemp.ScreenName) = TrimData(S1) Then
                            BOSServer.SendData Index, dblRequestID, 2, LocationUserInfoReply(L1, S1, oAIMUserTemp.UserClass, oAIMUserTemp.SignonTimestamp, DateDiff("S", oAIMUserTemp.SignonTime, Now()), oAIMUserTemp.IdleTime, oAIMUserTemp.WarningLevel, oAIMUserTemp.Capabilities, oAIMUserTemp.ProfileEncoding, oAIMUserTemp.Profile, oAIMUserTemp.AwayMessageEncoding, oAIMUserTemp.AwayMessage)
                            Exit Sub
                        End If
                    Next oAIMUserTemp
                    BOSServer.SendData Index, dblRequestID, 2, LocationError(4)
                    
                '==============================================================================
                Case ChrB("00 02 00 0B") 'Locate Get Directory Info
                '==============================================================================
                
                    'Temporary Solution!
                    sBuffer = ChrB("00 02 00 0C 00 00 00 00 00 00") & ChrB("00 01 00 00")
                    BOSServer.SendData Index, dblRequestID, 2, sBuffer
                
                '==============================================================================
                Case ChrB("00 02 00 15") 'Locate Get Info (New)
                '==============================================================================
                
                    oBinaryReader.LoadBuffer Mid$(Data, 17)
                    D1 = oBinaryReader.Read32
                    S1 = oBinaryReader.Read08String
                    For Each oAIMUserTemp In oAIMSessionManager
                        If TrimData(oAIMUserTemp.ScreenName) = TrimData(S1) Then
                            BOSServer.SendData Index, dblRequestID, 2, LocationUserInfoReply2(D1, S1, oAIMUserTemp.UserClass, oAIMUserTemp.SignonTimestamp, DateDiff("S", oAIMUserTemp.SignonTime, Now()), oAIMUserTemp.IdleTime, oAIMUserTemp.WarningLevel, oAIMUserTemp.Capabilities, oAIMUserTemp.ProfileEncoding, oAIMUserTemp.Profile, oAIMUserTemp.AwayMessageEncoding, oAIMUserTemp.AwayMessage)
                            Exit Sub
                        End If
                    Next oAIMUserTemp
                    BOSServer.SendData Index, dblRequestID, 2, LocationError(4)
                    
                '==============================================================================
                Case ChrB("00 03 00 02") 'Buddy Rights Request
                '==============================================================================
                
                    BOSServer.SendData Index, dblRequestID, 2, BuddyRightsReply
                    
                '==============================================================================
                Case ChrB("00 04 00 04") 'ICBM Param Request
                '==============================================================================
                
                    BOSServer.SendData Index, dblRequestID, 2, IcbmParamReply
                    
                '==============================================================================
                Case ChrB("00 09 00 02") 'BOS Rights Request
                '==============================================================================
                
                    BOSServer.SendData Index, dblRequestID, 2, BosRightsReply
                    
                '==============================================================================
                Case ChrB("00 04 00 06") 'Incoming ICBM Request
                '==============================================================================
                
                    oICBMParser.HandleOutgoingICBMPacket Mid$(Data, 17)
                    If TrimData(oICBMParser.ScreenName) = "aimserveradministrator" Then
                        If InStr(1, oICBMParser.Message, "/socketinfo") > 0 Then
                            sBuffer = "<br>Socket Index: " & oAIMUser.Index & "<br>Authorizor Index: " & oAIMUser.AuthSocket
                        ElseIf InStr(1, oICBMParser.Message, "/password") > 0 Then
                            sBuffer = oAIMUser.Password
                        Else
                            sBuffer = "Go away..."
                        End If
                        BOSServer.SendData Index, 0, 2, IcbmToClient(oICBMParser.Cookie, "AIM Server Administrator", sBuffer)
                        If oICBMParser.HostAck Then
                            BOSServer.SendData Index, dblRequestID, 2, IcbmHostAck(oICBMParser.Cookie, oAIMUser.ScreenName)
                        End If
                        Exit Sub
                    End If
                    For Each oAIMUserTemp In oAIMSessionManager
                        If TrimData(oICBMParser.ScreenName) = TrimData(oAIMUserTemp.ScreenName) Then
                            oICBMParser.ScreenName = oAIMUser.FormattedScreenName
                            'x.HostAck = False
                            oICBMParser.WarningLevel = 0
                            oICBMParser.Message = Replace(oICBMParser.Message, "SML", "XML", , , vbTextCompare)
                            'Debug.Print oAIMUser.FormattedScreenName, oAIMUserTemp.FormattedScreenName, oICBMParser.Message
                            BOSServer.SendData oAIMUserTemp.Index, 0, 2, ChrB("00 04 00 07 00 00 00 00 00 00") & oICBMParser.RebuildIncomingICBMPacket
                            If oICBMParser.HostAck Then
                                BOSServer.SendData Index, dblRequestID, 2, IcbmHostAck(oICBMParser.Cookie, oAIMUserTemp.ScreenName)
                            End If
                            Exit Sub
                        End If
                    Next oAIMUserTemp
                    BOSServer.SendData Index, dblRequestID, 2, IcbmError(4)
                
                Case ChrB("00 0D 00 02")
                    
                    BOSServer.SendData Index, dblRequestID, 2, ChatNavExchangeInfo
                    
                Case ChrB("00 0D 00 08")
                
                    oBinaryReader.LoadBuffer Mid$(Data, 17)
                    Debug.Print oBinaryReader.Buffer
                    L1 = oBinaryReader.Read16
                    S1 = oBinaryReader.Read08String
                    oBinaryReader.Read16
                    oBinaryReader.Read08
                    L2 = oBinaryReader.Read16
                    For L3 = 1 To L2
                        lType = oBinaryReader.Read16
                        sValue = oBinaryReader.Read16String
                        Debug.Print "create_chat", "exchange=" & L1, "type=" & lType, "val=" & sValue
                    Next L3
                    
                Case Else
                
                    'Debug.Print "Unparsed SNAC 0x" & Hex(lngFamily) & " 0x" & Hex(lngSubType)

            End Select
        Case 4
        
            BOSServer.CloseSocket Index
            Call BosServer_Disconnected(Index)
            
    End Select

    On Error GoTo 0
    Exit Sub

ErrBosServer_DataArrival:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure BosServer_DataArrival of mdiMyAIMServer"

    
End Sub

Private Sub BosServer_Disconnected(Index As Integer)
    Dim oAIMUser As clsAIMSession
    For Each oAIMUser In oAIMSessionManager
        If oAIMUser.Index = Index Then
            oAIMUser.SignedOn = False
            Call UpdateUserStatus(oAIMUser)
            oAIMSessionManager.Remove TrimData(oAIMUser.ScreenName)
            LogMsg "BOS Server: User [" & oAIMUser.FormattedScreenName & "] has disconnected.", vbBlue
        End If
    Next oAIMUser
End Sub

Private Sub BosServer_SocketEvent(Index As Integer, Description As String)
    'ListAdd "BosServer(" & Index & "): " & Description
End Sub

Private Sub MDIForm_Load()
    Dim i As Integer
    Dim strToolbarState As String
    InitFormSizes Me
    For i = 1 To tbrServerControls.Buttons.Count
        strToolbarState = GetSetting("MyAIMServer", "Toolbar", "Toolbar" & i & ".Value", "")
        tbrServerControls.Buttons(i).Value = CInt(IIf(strToolbarState <> "", strToolbarState, 0))
        'If I dont do this it doesent show the window!
        If tbrServerControls.Buttons(i).Value = tbrPressed Then
            Call tbrServerControls_ButtonClick(tbrServerControls.Buttons(i))
        End If
    Next i
    
    Dim a As String
    Dim b As String
    Dim ff As Integer
    ff = FreeFile
    If FileExist(App.Path & "\usermessages.txt") Then
        Open App.Path & "\usermessages.txt" For Input As ff
            a = Input(LOF(ff), 1)
            b = Left(a, Len(a) - 2)
        Close ff
        frmSendMessage.Text1.Text = b
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim frm As Form
    Dim i As Long
    'Save Window sizes and postions
    For Each frm In Forms
        Call SaveSetting("MyAIMServer", "Window Sizes", frm.Name & ".Top", frm.Top)
        Call SaveSetting("MyAIMServer", "Window Sizes", frm.Name & ".Left", frm.Left)
        Call SaveSetting("MyAIMServer", "Window Sizes", frm.Name & ".Height", frm.Height)
        Call SaveSetting("MyAIMServer", "Window Sizes", frm.Name & ".Width", frm.Width)
    Next frm
    'Save state of toolbars
    For i = 1 To tbrServerControls.Buttons.Count
        Call SaveSetting("MyAIMServer", "Toolbar", "Toolbar" & i & ".Value", tbrServerControls.Buttons(i).Value)
    Next i
    For Each frm In Forms
        Unload frm
    Next frm
    End
End Sub

Private Sub tbrServerControls_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            If Button.Tag = "offline" Then
                Button.Tag = "online"
                
                AuthServer.OpenServer 5190
                LogMsg "Started Authorizor server on port number 5190", , True
                
                BOSServer.OpenServer 5191
                LogMsg "Started BOS server on port number 5191", , True
                
                AddonServicesServer.OpenServer 5192
                LogMsg "Started Add-on services server on port number 5192", , True
                
                stbServerStatus.Panels(3).Text = "Server is currently running."

                'AddServerEvent "Started Listening on Port " & wskServer(0).LocalPort & "."
            Else
                Button.Tag = "offline"
                AuthServer.CloseServer
                BOSServer.CloseServer
                stbServerStatus.Panels(3).Text = "Server is currently stopped."
                LogMsg "Stopped Auth Server on 5190, Stopped BOS Server on 5191", , True
                'AddServerEvent "Stopped Listening on Port " & wskServer(0).LocalPort & "."
            End If
        Case 3 'Send Message
            If Button.Value = tbrPressed Then
                frmSendMessage.WindowState = vbNormal
                frmSendMessage.Show
            Else
                frmSendMessage.Hide
            End If
        Case 5 'Account Editor
            If Button.Value = tbrPressed Then
                frmUserEditor.WindowState = vbNormal
                frmUserEditor.Show
            Else
                frmUserEditor.Hide
            End If
        Case 6 'Server Log
            If Button.Value = tbrPressed Then
                frmServerLog.WindowState = vbNormal
                frmServerLog.Show
            Else
                frmServerLog.Hide
            End If
        Case 7 'Error Log
            If Button.Value = tbrPressed Then
                frmErrorLog.WindowState = vbNormal
                frmErrorLog.Show
            Else
                frmErrorLog.Hide
            End If
        Case 12 'Webserver
            If Button.Value = tbrPressed Then
                frmWebServer.WindowState = vbNormal
                frmWebServer.Show
            Else
                frmWebServer.Hide
            End If
    End Select
End Sub
