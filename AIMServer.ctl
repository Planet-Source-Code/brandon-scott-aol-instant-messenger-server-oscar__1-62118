VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl AIMServer 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSWinsockLib.Winsock sckAIMServer 
      Index           =   0
      Left            =   600
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   0
      Picture         =   "AIMServer.ctx":0000
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "AIMServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event DataArrival(Index As Integer, Data As String)
Public Event Connected(Index As Integer, RemoteHost As String)
Public Event Disconnected(Index As Integer)
Public Event SocketEvent(Index As Integer, Description As String)

Private ServerSequence() As Long
Private LocalSequence() As Long

Public Function IsConnected(Index As Integer) As Boolean
    If sckAIMServer(Index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function

Public Function CloseSocket(Index As Integer)
    ServerSequence(Index) = 0
    LocalSequence(Index) = 0
    sckAIMServer(Index).Close
End Function

Public Sub SendData(Index As Integer, requestID As Double, Frame As Byte, Data As String)
    If sckAIMServer(Index).State = sckConnected Then
        ServerSequence(Index) = ServerSequence(Index) + 1
        If ServerSequence(Index) = 65535 Then ServerSequence(Index) = 0
        If Frame = 2 Then
            If requestID > 0 Then
                Mid$(Data, 7, 4) = DWord(requestID)
            End If
        End If
        sckAIMServer(Index).SendData "*" & Chr(Frame) & Word(ServerSequence(Index)) & Word(Len(Data)) & Data
    End If
End Sub

Public Function CreateSock() As Integer
    'On Error Resume Next
    Dim i As Integer
    For i = 1 To sckAIMServer.UBound
        If sckAIMServer(i).State <> sckConnected Then
            CreateSock = i
            Exit Function
        End If
    Next i
    ReDim Preserve ServerSequence(0 To UBound(ServerSequence) + 1)
    ReDim Preserve LocalSequence(0 To UBound(LocalSequence) + 1)
    CreateSock = sckAIMServer.UBound + 1
    Load sckAIMServer(CreateSock)
End Function

Public Sub OpenServer(Port As Integer)
    'On Error Resume Next
    sckAIMServer(0).Close
    sckAIMServer(0).LocalPort = Port
    sckAIMServer(0).Listen
End Sub

Public Sub CloseServer()
    'On Error Resume Next
    Dim i As Integer
    For i = 1 To sckAIMServer.UBound
        sckAIMServer(i).Close
        Unload sckAIMServer(i)
    Next i
    ReDim Preserve LocalSequence(0)
    ReDim Preserve ServerSequence(0)
    sckAIMServer(0).Close
End Sub

Private Sub sckAIMServer_Close(Index As Integer)
    ServerSequence(Index) = 0
    LocalSequence(Index) = 0
    sckAIMServer(Index).Close
    RaiseEvent Disconnected(Index)
End Sub

Private Sub sckAIMServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    'On Error Resume Next
    Dim i As Integer
    i = CreateSock
    ServerSequence(i) = 0
    LocalSequence(i) = 0
    sckAIMServer(i).Close
    sckAIMServer(i).Accept requestID
    RaiseEvent Connected(i, sckAIMServer(Index).RemoteHostIP)
End Sub

Private Sub sckAIMServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    'On Error Resume Next
    Dim strData As String
    Dim lngLength As Long
Split:
    sckAIMServer(Index).PeekData strData, vbString
    If Len(strData) = 0 Then Exit Sub
    If Mid(strData, 1, 1) <> "*" Then
        RaiseEvent SocketEvent(Index, "Non-FLAP based packet received!")
        Exit Sub
    End If
    lngLength = GetWord(Mid(strData, 5, 2))
    If bytesTotal >= lngLength + 6 Then
        sckAIMServer(Index).GetData strData, vbString, lngLength + 6
        RaiseEvent DataArrival(Index, Mid(strData, 1, lngLength + 6))
        bytesTotal = bytesTotal - (lngLength + 6)
        If bytesTotal >= 6 Then GoTo Split
    End If
End Sub

Private Sub sckAIMServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ServerSequence(Index) = 0
    LocalSequence(Index) = 0
    sckAIMServer(Index).Close
    RaiseEvent Disconnected(Index)
End Sub

Private Sub UserControl_Initialize()
'On Error Resume Next
    ReDim Preserve LocalSequence(0)
    ReDim Preserve ServerSequence(0)
End Sub

Private Sub UserControl_Resize()
'On Error Resume Next
    UserControl.Width = imgLogo.Width
    UserControl.Height = imgLogo.Height
End Sub
