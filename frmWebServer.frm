VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWebServer 
   Caption         =   "Web Server"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWebServer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   10005
   Begin MSComctlLib.ImageList imgWebLogs 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWebServer.frx":1042
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkListen 
      Caption         =   "Listen"
      Height          =   255
      Left            =   7440
      TabIndex        =   1
      Top             =   5280
      Width           =   735
   End
   Begin MSWinsockLib.Winsock wsk 
      Index           =   0
      Left            =   720
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbrEvents 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   5235
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   17119
            Text            =   "Total Requests: [0]"
            TextSave        =   "Total Requests: [0]"
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
   Begin MSComctlLib.ListView lstLogs 
      Height          =   3615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgWebLogs"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "?"
         Object.Width           =   494
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "socketid"
         Text            =   "Index"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "time"
         Text            =   "Date/Time"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "request"
         Text            =   "Request"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "ipaddress"
         Text            =   "IP Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "percent"
         Text            =   "% Done"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "sendbytes"
         Text            =   "Sent Bytes"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "filesize"
         Text            =   "File Size"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmWebServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const HTTP_SERVER_VERSION As String = "XeonWebServer/2.0"

Private m_intHttpSocketCount As Integer
Private m_lngHttpTotalRequests As Long
Private m_lngHttpTotalReceivedBytes As Long
Private m_lngHttpTotalBytes() As Long
Private m_lngHttpTotalSentBytes() As Long
Private m_blnHttpRequestCompleted As Boolean
Private m_strHttpRequestURI As String
Private m_strHttpHeaders() As HTTPHeaderTag
Private m_strHttpPostData As String

Private Type HTTPHeaderTag
    strIdentifier As String
    strData As String
End Type

Private MimeTypes() As MimeType
Private Type MimeType
    strExt As String
    strType As String
End Type

Private Function LookupMime(strExtension As String) As String
    Dim i As Integer
    For i = 0 To UBound(MimeTypes)
        If strExtension = MimeTypes(i).strExt Then
            LookupMime = MimeTypes(i).strType
            Exit Function
        End If
    Next i
    LookupMime = "text/plain"
End Function

Private Sub LoadMimeTypes()
    Dim strData As String
    Dim strPath As String
    Dim strMimeData() As String
    Dim iFile As Integer
    strPath = App.Path & "\MimeTypes.dat"
    If FileExist(App.Path & "\MimeTypes.dat") = False Then
        MsgBox "There was a problem loading mimetypes.dat" & vbCrLf & "Please make sure you unzipped all the files correctly and that none are corrupt.", vbCritical, "Error"
        End
    End If
    ReDim Preserve MimeTypes(0)
    iFile = FreeFile
    Open strPath For Input As #iFile
        While Not EOF(iFile)
            Input #iFile, strData
            strMimeData() = Split(strData, "ÿÿ")
            ReDim Preserve MimeTypes(0 To UBound(MimeTypes) + 1)
            MimeTypes(UBound(MimeTypes)).strType = strMimeData(0)
            MimeTypes(UBound(MimeTypes)).strExt = strMimeData(1)
        Wend
    Close #iFile
End Sub

Private Sub ParseWebRequest(ByVal Index As Integer, ByVal Data As String)

    On Error GoTo ErrParseWebRequest

    Dim arrArgs() As String, arrArgs2() As String
    Dim strRequest As String
    Dim intFreeFile As Integer
    Dim bytData() As Byte, strBuffer As String, strMimeType As String
    Dim strHeader As String
    Dim lngChunkSize As Long
    Dim sScreenName As String, sPassword As String, sConfirm As String, sEmail As String
    Dim sResult As String
    Dim i As Integer, j As Integer
    If InStr(1, Data, "HTTP") <> 0 And (InStr(1, Data, "GET") = 1 Or InStr(1, Data, "POST") = 1) Then
        'Check Socket
        If wsk(Index).State <> 7 Then Exit Sub
        If Left(Data, 3) = "GET" Then
            'Decode request
            strRequest = DecodeStr(Mid$(Data, 5, InStr(5, Data, " ") - 5))
            lstLogs.ListItems(Index).SubItems(3) = strRequest
            If Not InStr(1, strRequest, "./") > 0 And Not InStr(1, strRequest, ".\") > 0 Then
                'Flip the slashes around
                strRequest = Replace(strRequest, "/", "\")
                'Check if requesting index page
                If strRequest = "\" Then strRequest = "\index.html"
                'Get webserver path and add request onto it
                strRequest = App.Path & "\Web\" & strRequest
                'Get next free file number
                intFreeFile = FreeFile
                'Reset Total Byte Counter
                m_lngHttpTotalBytes(Index) = 0
                'Reset Total Send Bytes
                m_lngHttpTotalSentBytes(Index) = 0
                'Check if file exists
                If FileExist(strRequest) = True Then
                    strMimeType = LookupMime(LCase(Mid$(strRequest, InStrRev(strRequest, "."))))
                    If strMimeType = "text/html" Then
                        'Open file and put data into buffer
                        Open strRequest For Binary As #intFreeFile
                            ReDim bytData(FileLen(strRequest))
                            Get #intFreeFile, , bytData
                        Close #intFreeFile
                        'Convert to string
                        strBuffer = StrConv(bytData, vbUnicode)
                        'Add header
                        strHeader = BuildHeader(strMimeType, Len(strBuffer))
                        'Calculate total bytes
                        m_lngHttpTotalBytes(Index) = Len(strHeader) + Len(strBuffer)
                        'Send Data
                        wsk(Index).SendData strHeader
                        wsk(Index).SendData strBuffer
                    Else
                        'Add header
                        strHeader = BuildHeader(strMimeType, FileLen(strRequest))
                        'Calculate total bytes
                        m_lngHttpTotalBytes(Index) = Len(strHeader) + FileLen(strRequest)
                        'Send Header
                        wsk(Index).SendData strHeader
                        'Open and send file
                        Open strRequest For Binary As #intFreeFile
                            Do While Not EOF(intFreeFile)
                                'So we dont over read
                                If (LOF(intFreeFile) - Loc(intFreeFile)) < 8192 Then
                                    ReDim bytData((LOF(intFreeFile) - Loc(intFreeFile)))
                                Else
                                    ReDim bytData(8192)
                                End If
                                '8 K/b Buffer to save memory
                                Get #intFreeFile, , bytData
                                wsk(Index).SendData bytData
                            Loop
                        Close #intFreeFile
                    End If
                Else
                    '404 Error
                    bytData = LoadResData("DOCUMENT_NOT_FOUND", "WEBSERVER")
                    'Convert to string
                    strBuffer = StrConv(bytData, vbUnicode)
                    'Add header
                    strHeader = BuildHeader("text/html", Len(strBuffer))
                    'Calculate total bytes
                    m_lngHttpTotalBytes(Index) = Len(strHeader) + Len(strBuffer)
                    'Send Data
                    wsk(Index).SendData strHeader
                    wsk(Index).SendData bytData
                End If
            Else
                LogMsg "Web Server: Attempted hack attempt by " & wsk(Index).RemoteHostIP & ". Requested URL: " & strRequest
                '404 Error
                bytData = LoadResData("FORBIDDEN", "WEBSERVER")
                'Convert to string
                strBuffer = StrConv(bytData, vbUnicode)
                'Add header
                strHeader = BuildHeader("text/html", Len(strBuffer))
                'Calculate total bytes
                m_lngHttpTotalBytes(Index) = Len(strHeader) + Len(strBuffer)
                'Send Data
                wsk(Index).SendData strHeader
                wsk(Index).SendData bytData
            End If
        ElseIf Left(Data, 4) = "POST" Then
            'Decode request
            strRequest = DecodeStr(Mid$(Data, 6, InStr(6, Data, " ") - 5))
            lstLogs.ListItems(Index).SubItems(3) = strRequest
            Debug.Print strRequest
            If strRequest = "/create " Then
                arrArgs = Split(Data, vbCrLf & vbCrLf, 2)
                arrArgs2 = Split(arrArgs(1), "&")
                For i = 0 To UBound(arrArgs2)
                    arrArgs = Split(arrArgs2(i), "=", 2)
                    Select Case arrArgs(0)
                        Case "screenname"
                            sScreenName = arrArgs(1)
                        Case "password"
                            sPassword = arrArgs(1)
                        Case "confirm"
                            sConfirm = arrArgs(1)
                        Case "email"
                            sEmail = arrArgs(1)
                    End Select
                Next i
                
                sResult = RegisterName(wsk(Index).RemoteHostIP, sScreenName, sPassword, sConfirm, sEmail)
                Select Case sResult
                    Case "invalidname"
                        strBuffer = "Invalid screen name. Must be between 2-18 characters, and start with a letter."
                    Case "invalidpassword"
                        strBuffer = "Invalid password."
                    Case "invalidconfirm"
                        strBuffer = "Confirm doesnt match password."
                    Case "invalidemail"
                        strBuffer = "Invalid email address."
                    Case "alreadyexist"
                        strBuffer = "Oops! That name already exists."
                    Case "good"
                        strBuffer = "Congrats! Your new screen name is `" & sScreenName & "`. Please note this site is in no way affiliated with America Online."
                End Select
                'Add header
                strHeader = BuildHeader("text/html", Len(strBuffer))
                'Calculate total bytes
                m_lngHttpTotalBytes(Index) = Len(strHeader) + Len(strBuffer)
                'Send Data
                wsk(Index).SendData strHeader
                wsk(Index).SendData strBuffer
            Else
                Call CloseSocket(Index)
            End If
        End If
    End If

    On Error GoTo 0
    Exit Sub

ErrParseWebRequest:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure ParseWebRequest of frmWebServer"
    Call CloseSocket(Index)

End Sub

Public Function BuildHeader(strMimeType As String, lngContentLength As Long)
    Dim strHeader As String
    strHeader = strHeader & "HTTP/1.1 200 OK" & vbCrLf
    strHeader = strHeader & "Content-Type: " & strMimeType & vbCrLf
    strHeader = strHeader & "Content-Length: " & lngContentLength & vbCrLf
    strHeader = strHeader & "Server: " & HTTP_SERVER_VERSION & vbCrLf
    strHeader = strHeader & vbCrLf
    BuildHeader = strHeader
End Function

Private Sub CloseSocket(Index As Integer)
    On Error GoTo ErrCloseSocket

    If Index <> 0 Then
        'TODO: add some actions here!
        wsk(Index).Close
    End If

    On Error GoTo 0
    Exit Sub

ErrCloseSocket:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure CloseSocket of frmWebServer"
    
End Sub

Private Sub chkListen_Click()
    If chkListen.Value Then
        wsk(0).LocalPort = 4447
        wsk(0).Listen
        LogMsg "Webserver Started at " & wsk(0).LocalIP & ":" & wsk(0).LocalPort & ".", , True
    Else
        wsk(0).Close
        LogMsg "Webserver Stopped.", , True
    End If
End Sub

Private Sub Form_Load()
    InitFormSizes Me
    
    'Load Mime Types
    Call LoadMimeTypes
    
    ReDim m_lngTotalBytes(0)
    ReDim m_lngTotalSentBytes(0)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        lstLogs.Width = Me.Width - 120
        lstLogs.Height = Me.Height - 800
        chkListen.Top = Me.Height - 755
        chkListen.Left = (Me.Width - chkListen.Width) - 450
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    mdiMyAIMServer.tbrServerControls.Buttons(12).Value = tbrUnpressed
    Me.Hide
End Sub

Private Sub wsk_Close(Index As Integer)
    Call CloseSocket(Index)
End Sub

Private Sub wsk_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim i As Integer
    Dim intIPCount As Integer
    Dim lstItem As ListItem
    On Error GoTo Errwsk_ConnectionRequest

    For i = wsk.LBound To wsk.UBound
        If wsk(i).RemoteHostIP = wsk(Index).RemoteHostIP And wsk(i).State = 7 Then
            intIPCount = intIPCount + 1
            If intIPCount > 10 Then Exit Sub
        End If
    Next i
    m_lngHttpTotalRequests = m_lngHttpTotalRequests + 1
    m_intHttpSocketCount = m_intHttpSocketCount + 1
    Load wsk(m_intHttpSocketCount)
    ReDim Preserve m_lngHttpTotalBytes(m_intHttpSocketCount)
    ReDim Preserve m_lngHttpTotalSentBytes(m_intHttpSocketCount)
    wsk(m_intHttpSocketCount).Accept requestID
    Set lstItem = lstLogs.ListItems.Add(, , , , 1)
    lstItem.SubItems(1) = m_intHttpSocketCount
    lstItem.SubItems(2) = Now
    lstItem.SubItems(4) = wsk(0).RemoteHostIP
    sbrEvents.Panels(1).Text = "Total Requests: [" & m_lngHttpTotalRequests & "]"

    On Error GoTo 0
    Exit Sub

Errwsk_ConnectionRequest:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure wsk_ConnectionRequest of frmWebServer"
    Call CloseSocket(Index)

End Sub

Private Sub wsk_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    Dim strData As String
    Dim strTemp As String
    On Error GoTo Errwsk_DataArrival
    
    wsk(Index).GetData strData
    If Left(strData, 4) = "POST" And InStr(1, strData, "Content-Length", vbTextCompare) = 0 Then
        wsk(Index).Tag = wsk(Index).Tag & strData
        Exit Sub
    Else
        strData = wsk(Index).Tag & strData
        wsk(Index).Tag = vbNullString
    End If
    'Debug.Print DisplayFormat(strData)
    m_lngHttpTotalReceivedBytes = m_lngHttpTotalReceivedBytes + bytesTotal
    Call ParseWebRequest(Index, strData)

    On Error GoTo 0
    Exit Sub

Errwsk_DataArrival:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure wsk_DataArrival of frmWebServer"
    Call CloseSocket(Index)

End Sub

Private Sub wsk_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call CloseSocket(Index)
End Sub

Private Sub wsk_SendComplete(Index As Integer)
    If m_lngHttpTotalSentBytes(Index) >= m_lngHttpTotalBytes(Index) Then
        Call CloseSocket(Index)
        Exit Sub
    End If
End Sub

Private Sub wsk_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)

    Dim i As Integer
    On Error GoTo Errwsk_SendProgress

    m_lngHttpTotalSentBytes(Index) = m_lngHttpTotalSentBytes(Index) + bytesSent
    lstLogs.ListItems(Index).SubItems(5) = Round((m_lngHttpTotalSentBytes(Index) / m_lngHttpTotalBytes(Index)) * 100, 0) & "%"
    lstLogs.ListItems(Index).SubItems(6) = Round(m_lngHttpTotalSentBytes(Index) / 1024, 2) & " KB"
    lstLogs.ListItems(Index).SubItems(7) = Round(m_lngHttpTotalBytes(Index) / 1024, 2) & " KB"

    On Error GoTo 0
    Exit Sub

Errwsk_SendProgress:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure wsk_SendProgress of frmWebServer"
    Call CloseSocket(Index)

End Sub
