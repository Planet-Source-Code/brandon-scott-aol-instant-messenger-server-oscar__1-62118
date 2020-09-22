VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "My AIM Server"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDebug 
      Caption         =   "Debug"
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   345
      Left            =   4320
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   345
      Left            =   4320
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.ListBox lstEventLog 
      Height          =   2205
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin MyAIMServer.AIMServer AuthServer 
      Left            =   2760
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MyAIMServer.AIMServer BosServer 
      Left            =   3240
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AuthServer_Connected(Index As Integer, RemoteHost As String)
    AuthServer.SendData Index, 0, 1, flap_version
End Sub

Private Sub AuthServer_DataArrival(Index As Integer, Data As String)
    Debug.Print "INCOMING AUTH: " & DisplayFormat(Data)
    Dim A1, A2, A3, A4, A5, A6, A7, A8, A9
    If 10 > Len(Data) Then Exit Sub
    Select Case Asc(Mid(Data, 2, 1))
        Case 1
        Case 2
            Select Case Mid(Data, 7, 4)
                Case ChrB("00 17 00 06")
                    AuthServer.SendData Index, 0, 2, bucp_challenge(GRTicket)
                Case ChrB("00 17 00 02")
                    A1 = GetTLV(1, Mid(Data, 17))
                    A2 = GRCookie
                    AddSession CStr(A1), CStr(A2)
                    AuthServer.SendData Index, 0, 2, bucp_reply(CStr(A1), CStr(A2))
                    ListAdd CStr(A1) & " has signed on."
                    'MsgBox CStr(A1)
                Case Else
            End Select
        Case 4
            AuthServer.ResetSock Index
    End Select
End Sub

Private Sub AuthServer_SocketEvent(Index As Integer, Description As String)
    ListAdd "AuthServer(" & Index & "): " & Description
End Sub

Private Sub BosServer_Connected(Index As Integer, RemoteHost As String)
    ListAdd "Bos: " & RemoteHost & " - Connected"
    BosServer.SendData Index, 0, 1, flap_version
End Sub

Private Sub BosServer_DataArrival(Index As Integer, Data As String)
    Debug.Print "INCOMING BOS: " & DisplayFormat(Data)
    Dim A1, A2, A3, A4, A5, A6, A7, A8, A9
    Dim S1 As String, S2 As String, S3 As String, S4 As String, S5 As String
    Dim SessionI As Integer
    Dim x As New clsICBMPacket
    Dim i As Integer
    If 10 > Len(Data) Then Exit Sub
    Select Case Asc(Mid(Data, 2, 1))
        Case 1
            If Len(Mid(Data, 7)) > 4 Then
                'we have a cookie and shit attached
                A1 = FindIndexByCookie(GetTLV(6, Mid(Data, 11)))
                AIMSessions(A1).Index = Index
                BosServer.SendData Index, 0, 2, host_online
            End If
        Case 2
            SessionI = FindSession(Index)
            Select Case Mid(Data, 7, 4)
                Case ChrB("00 01 00 17")
                    BosServer.SendData SessionI, 0, 2, host_versions
                    BosServer.SendData SessionI, 0, 2, host_motd
                Case ChrB("00 01 00 06")
                    BosServer.SendData SessionI, GetDWord(Mid(Data, 13, 4)), 2, rate_params
                Case ChrB("00 01 00 08")
                Case ChrB("00 01 00 0E")
                    BosServer.SendData SessionI, GetDWord(Mid(Data, 13, 4)), 2, NickInfoReply(AIMSessions(SessionI).strScreenName)
                Case ChrB("00 13 00 02")
                    BosServer.SendData SessionI, GetDWord(Mid(Data, 13, 4)), 2, FeedbagRightsReply
                Case ChrB("00 13 00 05")
                    BosServer.SendData SessionI, GetDWord(Mid(Data, 13, 4)), 2, FeedbagReplyNotModified(Mid$(Data, 17, 6))
                    'BosServer.SendData SessionI, GetDWord(Mid(Data, 13, 4)), 2, FeedbagBuddylist
                    'BosServer.SendData SessionI, GetDWord(Mid(Data, 13, 4)), 2, FeedbagReplyNotModified
                    'BosServer.SendData SessionI, GetDWord(Mid(Data, 13, 4)), 2, FeedbagError
                Case ChrB("00 13 00 04")
                    BosServer.SendData SessionI, GetDWord(Mid(Data, 13, 4)), 2, FeedbagBuddylist
                    'BosServer.SendData SessionI, GetDWord(Mid(Data, 13, 4)), 2, FeedbagReplyNotModified
                    
                Case ChrB("00 02 00 02")
                    BosServer.SendData SessionI, GetDWord(Mid(Data, 13, 4)), 2, LocationRightsReply
                Case ChrB("00 02 00 04")
                    Debug.Print DisplayFormat(Mid$(Data, 17))
                    If GetTLV(1, Mid$(Data, 17)) = "<none>" Or GetTLV(2, Mid$(Data, 17)) = "<none>" Then Exit Sub
                    AIMSessions(SessionI).sProfileEncoding = GetTLV(1, Mid$(Data, 17))
                    AIMSessions(SessionI).sProfile = GetTLV(2, Mid$(Data, 17))
                Case ChrB("00 02 00 05")
                
                    S1 = CStr(GetSByte(Mid$(Data, 19)))
                    For i = 1 To UBound(AIMSessions)
                        If LCase(Replace(AIMSessions(i).strScreenName, " ", "")) = LCase(Replace(S1, " ", "")) Then
                            BosServer.SendData SessionI, GetDWord(Mid(Data, 13, 4)), 2, LocationProfile(S1, AIMSessions(i).dSignonTime, 0, 0, AIMSessions(i).sProfileEncoding, AIMSessions(i).sProfile)
                           'Exit Sub
                        End If
                    Next i
                    
                Case ChrB("00 03 00 02")
                    BosServer.SendData SessionI, GetDWord(Mid(Data, 13, 4)), 2, BuddyRightsReply
                Case ChrB("00 04 00 04")
                    BosServer.SendData SessionI, GetDWord(Mid(Data, 13, 4)), 2, IcbmParamReply
                Case ChrB("00 09 00 02")
                    BosServer.SendData SessionI, GetDWord(Mid(Data, 13, 4)), 2, BosRightsReply
                Case ChrB("00 04 00 06")
                    Call x.HandleOutgoingICBMPacket(Mid(Data, 17))
                    For i = 1 To UBound(AIMSessions)
                        If LCase(Replace(AIMSessions(i).strScreenName, " ", "")) = LCase(Replace(x.ScreenName, " ", "")) Then
                            'get are message data
                            S1 = GetTLV(1281, x.Message) 'whisc caps
                            S2 = GetTLV(257, x.Message) 'message data
                            S3 = Mid$(S2, 1, 4) 'message encoding
                            S4 = Mid$(S2, 5) 'message
                            BosServer.SendData i, 0, 2, SendIncomingICBM(x.Cookie, AIMSessions(Index).strScreenName, S4)
                            'Exit Sub
                        End If
                    Next i
                    BosServer.SendData SessionI, GetDWord(Mid(Data, 13, 4)), 2, SendIcbmHostAck(x.Cookie, x.ScreenName)
                Case Else
            End Select
        Case 4
            BosServer.ResetSock Index
    End Select
End Sub

Private Sub BosServer_SocketEvent(Index As Integer, Description As String)
    ListAdd "BosServer(" & Index & "): " & Description
End Sub

Private Sub cmdDebug_Click()
'    InitializeDatabase
 '   AddTest
  '  TerminateDatabase
  
    Dim i As Integer
    For i = 1 To UBound(AIMSessions)
    BosServer.SendData i, 0, 2, SendIncomingICBM("AAAAAAAA", "a", "hey")
    Next i
End Sub

Private Sub cmdStart_Click()
    AuthServer.OpenServer 4448
    BosServer.OpenServer 4449
End Sub

Private Sub cmdStop_Click()
    AuthServer.CloseServer
End Sub

Public Sub ListAdd(strData As String)
    Dim i As Integer
    Dim B1 As String, B2 As Integer, B3 As String, B4 As String
    For i = 0 To lstEventLog.ListCount - 1
        If Right(lstEventLog.List(i), Len(strData) + 2) = "x " & strData Then
            B1 = lstEventLog.List(i)
            B2 = InStr(1, B1, "x")
            B3 = Left(B1, B2 - 2)
            B4 = Right(B1, Len(B1) - (B2 + 1))
            B3 = B3 + 1
            lstEventLog.List(i) = B3 & " x " & B4
            Exit Sub
        End If
        DoEvents
    Next i
    lstEventLog.AddItem "1 x " & strData
End Sub

Private Sub Form_Load()
    Debug.Print Now
    ReDim Preserve AIMSessions(0)
End Sub
