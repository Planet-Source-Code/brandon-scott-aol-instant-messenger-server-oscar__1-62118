Attribute VB_Name = "modServer"
Option Explicit

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hSessionKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'[Return Events]
Public Enum TrayRetunEventEnum
    MouseMove = &H200       'On Mousemove
    LeftUp = &H202          'Left Button Mouse Up
    LeftDown = &H201        'Left Button MouseDown
    LeftDbClick = &H203     'Left Button Double Click
    RightUp = &H205         'Right Button Up
    RightDown = &H204       'Right Button Down
    RightDbClick = &H206    'Right Button Double Click
    MiddleUp = &H208        'Middle Button Up
    MiddleDown = &H207      'Middle Button Down
    MiddleDbClick = &H209   'Middle Button Double Click
End Enum

'[Modify Items]
Public Enum ModifyItemEnum
    ToolTip = 1             'Modify ToolTip
    Icon = 2                'Modify Icon
End Enum

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2

Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206

Private nid As NOTIFYICONDATA

Private Const SERVICE_PROVIDER As String = "Microsoft Enhanced Cryptographic Provider v1.0" & vbNullChar
Private Const KEY_CONTAINER As String = "GCN SSL Container" & vbNullChar
Private Const HP_HASHVAL As Long = 2
Private Const PROV_RSA_FULL As Long = 1
Private Const CALG_MD5 As Long = 32771
Private Const CRYPT_VERIFYCONTEXT = &HF0000000
Private Const CRYPT_NEWKEYSET As Long = 8

Public Const FEEDBAG_CLASS_ID_BUDDY = 0
Public Const FEEDBAG_CLASS_ID_GROUP = 1
Public Const FEEDBAG_CLASS_ID_PERMIT = 2
Public Const FEEDBAG_CLASS_ID_DENY = 3
Public Const FEEDBAG_CLASS_ID_PDINFO = 4
Public Const FEEDBAG_CLASS_ID_BUDDY_PREFS = 5
Public Const FEEDBAG_CLASS_ID_NONBUDDY = 6
Public Const FEEDBAG_CLASS_ID_TPA_PROVIDER = 7
Public Const FEEDBAG_CLASS_ID_TPA_SUBSCRIPTION = 8
Public Const FEEDBAG_CLASS_ID_CLIENT_PREFS = 9
Public Const FEEDBAG_CLASS_ID_STOCK = 10
Public Const FEEDBAG_CLASS_ID_WEATHER = 11
Public Const FEEDBAG_CLASS_ID_WATCH_LIST = 13
Public Const FEEDBAG_CLASS_ID_IGNORE_LIST = 14
Public Const FEEDBAG_CLASS_ID_DATE_TIME = 15
Public Const FEEDBAG_CLASS_ID_EXTERNAL_USER = 16
Public Const FEEDBAG_CLASS_ID_ROOT_CREATOR = 17
Public Const FEEDBAG_CLASS_ID_FISH = 18
Public Const FEEDBAG_CLASS_ID_IMPORT_TIMESTAMP = 19
Public Const FEEDBAG_CLASS_ID_BART = 20

Public Const FEEDBAG_STATUS_CODES_SUCCESS = 0
Public Const FEEDBAG_STATUS_CODES_DB_ERROR = 1
Public Const FEEDBAG_STATUS_CODES_NOT_FOUND = 2
Public Const FEEDBAG_STATUS_CODES_ALREADY_EXISTS = 3
Public Const FEEDBAG_STATUS_CODES_BAD_REQUEST = 10
Public Const FEEDBAG_STATUS_CODES_DB_TIME_OUT = 11
Public Const FEEDBAG_STATUS_CODES_OVER_ROW_LIMIT = 12
Public Const FEEDBAG_STATUS_CODES_NOT_EXECUTED = 13
Public Const FEEDBAG_STATUS_CODES_AUTH_REQUIRED = 14
Public Const FEEDBAG_STATUS_CODES_AUTO_AUTH = 15

Public oAIMSessionManager As clsAIMSessionManager
Public oAIMChats As Collection

Public Sub LoadIcon(frmForm As Form, strToolTip As String)
    With nid
        .cbSize = Len(nid)
        .hwnd = frmForm.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = frmForm.Icon
        .szTip = " " & strToolTip & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid
End Sub

Public Sub UnloadIcon()
    Shell_NotifyIcon NIM_DELETE, nid
End Sub

Public Sub Main()

    On Error GoTo ErrMain

    InitializeDatabase
    Set oAIMSessionManager = New clsAIMSessionManager
    Set oAIMChats = New Collection
    Load mdiMyAIMServer
    mdiMyAIMServer.Show
    
    Dim oAdmin As clsAIMSession
    Set oAdmin = oAIMSessionManager.Add("aimserveradministrator", 0, "")
    oAdmin.ID = 0
    oAdmin.Index = 0
    oAdmin.SignonTime = Now
    oAdmin.SignonTimestamp = GetTimeStamp
    oAdmin.UserClass = &H12&
    oAdmin.ScreenName = "aimserveradministrator"
    oAdmin.FormattedScreenName = "AIM Server Administrator"
    oAdmin.EmailAddress = ""
    oAdmin.Password = ""
    oAdmin.WarningLevel = 0
    oAdmin.Authorized = True
    oAdmin.SignedOn = True

    On Error GoTo 0
    Exit Sub

ErrMain:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure Main of modServer"
    
End Sub

Public Sub UpdateAdminAccountInfo(oUser As clsAIMSession)
    
    On Error GoTo ErrUpdateAdminAccountInfo

    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    RS.Open "SELECT * FROM [Registration] WHERE [Screenname] = '" & oUser.ScreenName & "'", DB_Connection, adOpenKeyset, adLockOptimistic
    RS.Fields("Formatted ScreenName") = oUser.FormattedScreenName
    RS.Fields("Email Address") = oUser.EmailAddress
    RS.Update
    RS.Close
    Set RS = Nothing

    On Error GoTo 0
    Exit Sub

ErrUpdateAdminAccountInfo:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure UpdateAdminAccountInfo of modServer"

End Sub

Public Sub FeedbagCheckMasterGroup(oUser As clsAIMSession)
    Dim RS As ADODB.Recordset
    On Error GoTo ErrFeedbagCheckMasterGroup

    Set RS = New ADODB.Recordset
    RS.Open "SELECT * FROM [Feedbag] WHERE [ID] = " & oUser.ID, DB_Connection, adOpenKeyset, adLockOptimistic
    If RS.RecordCount = 0 Then
        'Master Group (ALWAYS NEEDED)
        RS.AddNew
        RS.Fields("ID") = oUser.ID
        RS.Fields("Name") = ""
        RS.Fields("Group ID") = 0
        RS.Fields("Buddy ID") = 0
        RS.Fields("Class ID") = 1
        RS.Fields("Attributes") = ChrB("00 C8 00 06 00 01 00 02 00 03")
        RS.Update
        
        'Buddy Prefs
        RS.AddNew
        RS.Fields("ID") = oUser.ID
        RS.Fields("Name") = ""
        RS.Fields("Group ID") = 0
        RS.Fields("Buddy ID") = 1
        RS.Fields("Class ID") = 5
        RS.Fields("Attributes") = ChrB("00 C9 00 04 00 61 E7 FF 00 D6 00 04 00 77 FF FF")
        RS.Update
        'Buddies
        RS.AddNew
        RS.Fields("ID") = oUser.ID
        RS.Fields("Name") = "Buddies"
        RS.Fields("Group ID") = 1
        RS.Fields("Buddy ID") = 0
        RS.Fields("Class ID") = 1
        RS.Fields("Attributes") = ""
        RS.Update
        'Family
        RS.AddNew
        RS.Fields("ID") = oUser.ID
        RS.Fields("Name") = "Family"
        RS.Fields("Group ID") = 2
        RS.Fields("Buddy ID") = 0
        RS.Fields("Class ID") = 1
        RS.Fields("Attributes") = ""
        RS.Update
        'Co-Workers
        RS.AddNew
        RS.Fields("ID") = oUser.ID
        RS.Fields("Name") = "Co-Workers"
        RS.Fields("Group ID") = 3
        RS.Fields("Buddy ID") = 0
        RS.Fields("Class ID") = 1
        RS.Fields("Attributes") = ""
        RS.Update
        RS.Close
        'Update Feedbag Info
        RS.Open "SELECT * FROM [Registration] WHERE [ScreenName] = '" & oUser.ScreenName & "'", DB_Connection, adOpenKeyset, adLockOptimistic
        RS.Fields("Feedbag Timestamp") = GetTimeStamp
        RS.Fields("Feedbag Items") = 5
        RS.Update
        RS.Close
        LogMsg "Setup buddylist groups for [" & oUser.ScreenName & "]."
        Exit Sub
    End If
    RS.Close

    On Error GoTo 0
    Exit Sub

ErrFeedbagCheckMasterGroup:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure FeedbagCheckMasterGroup of modServer"

End Sub

Public Function GetFeedbagData(oUser As clsAIMSession) As String
    Dim sBuffer As String
    Dim RS As ADODB.Recordset
    On Error GoTo ErrGetFeedbagData

    Set RS = New ADODB.Recordset
    RS.Open "SELECT * FROM [Feedbag] WHERE [ID] = " & oUser.ID & " ORDER BY [Group ID],[Buddy ID] ASC", DB_Connection, adOpenKeyset, adLockOptimistic
    sBuffer = sBuffer & Chr(0) & Word(RS.RecordCount)
    Do Until RS.EOF
        sBuffer = sBuffer & SWord(RS.Fields("Name")) & Word(RS.Fields("Group ID")) & Word(RS.Fields("Buddy ID")) & Word(RS.Fields("Class ID")) & SWord(RS.Fields("Attributes"))
        RS.MoveNext
    Loop
    RS.Close
    RS.Open "SELECT * FROM [Registration] WHERE [ScreenName] = '" & oUser.ScreenName & "'", DB_Connection, adOpenKeyset, adLockOptimistic
    If RS.RecordCount > 0 Then
        sBuffer = sBuffer & DWord(RS.Fields("Feedbag Timestamp"))
    End If
    RS.Close
    GetFeedbagData = sBuffer

    On Error GoTo 0
    Exit Function

ErrGetFeedbagData:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure GetFeedbagData of modServer"

End Function

Public Function FeedbagCheckIfNew(oUser As clsAIMSession, dblTimestamp As Double, lngItems As Long) As Boolean
    Dim RS As ADODB.Recordset
    On Error GoTo ErrFeedbagCheckIfNew

    Set RS = New ADODB.Recordset
    RS.Open "SELECT * FROM [Registration] WHERE [ScreenName] = '" & oUser.ScreenName & "'", DB_Connection, adOpenKeyset, adLockOptimistic
    'MsgBox "Database: " & StringToHexArray(DWord(RS.Fields("Feedbag Timestamp")))
    If dblTimestamp = RS.Fields("Feedbag Timestamp") And lngItems = RS.Fields("Feedbag Items") Then
        FeedbagCheckIfNew = False
        RS.Close
        Exit Function
    Else
        FeedbagCheckIfNew = True
        RS.Close
        Exit Function
    End If
    RS.Close

    On Error GoTo 0
    Exit Function

ErrFeedbagCheckIfNew:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure FeedbagCheckIfNew of modServer"

End Function

Public Function FeedbagAddItem(oUser As clsAIMSession, _
                               sName As String, _
                               lGroupID As Long, lBuddyID As Long, lClassID As Long, _
                               sAttributes As String) As Long
                                
    Dim RS As ADODB.Recordset
    On Error GoTo ErrFeedbagAddItem

    Set RS = New ADODB.Recordset
    Dim blnGroupExists As Boolean
    
    'Debug.Print "FEEDBAG_ADD " & sName, lGroupID, lBuddyID, lClassID
    
    RS.Open "SELECT * FROM [Feedbag] WHERE [ID] = " & oUser.ID, DB_Connection, adOpenKeyset, adLockOptimistic
    RS.AddNew
    RS.Fields("ID") = oUser.ID
    RS.Fields("Name") = sName
    RS.Fields("Group ID") = lGroupID
    RS.Fields("Buddy ID") = lBuddyID
    RS.Fields("Class ID") = lClassID
    RS.Fields("Attributes") = sAttributes
    RS.Update
    RS.Close
    FeedbagAddItem = FEEDBAG_STATUS_CODES_SUCCESS

    On Error GoTo 0
    Exit Function

ErrFeedbagAddItem:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure FeedbagAddItem of modServer"

End Function

Public Function FeedbagDeleteItem(oUser As clsAIMSession, _
                               sName As String, _
                               lGroupID As Long, lBuddyID As Long, lClassID As Long, _
                               sAttributes As String) As Long
                               
    Dim RS As ADODB.Recordset
    On Error GoTo ErrFeedbagDeleteItem
    
    'Debug.Print "FEEDBAG_DELETE " & sName, lGroupID, lBuddyID, lClassID

    Set RS = New ADODB.Recordset
    RS.Open "SELECT * FROM [Feedbag] WHERE [ID] = " & oUser.ID & " AND [Group ID] = " & lGroupID & " AND [Buddy ID] = " & lBuddyID & " AND [Class ID] = " & lClassID, DB_Connection, adOpenKeyset, adLockOptimistic
    If lGroupID <> 0 Then
        If RS.RecordCount > 0 Then
            RS.Delete
            RS.Close
            FeedbagDeleteItem = FEEDBAG_STATUS_CODES_SUCCESS
            Exit Function
        End If
    End If
    FeedbagDeleteItem = 16

    On Error GoTo 0
    Exit Function

ErrFeedbagDeleteItem:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure FeedbagDeleteItem of modServer"

End Function

Public Function FeedbagUpdateItem(oUser As clsAIMSession, _
                               sName As String, _
                               lGroupID As Long, lBuddyID As Long, lClassID As Long, _
                               sAttributes As String) As Long
                               
    'Debug.Print "FEEDBAG_UPDATE " & sName, lGroupID, lBuddyID, lClassID
                               
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    RS.Open "SELECT * FROM [Feedbag] WHERE [ID] = " & oUser.ID & " AND [Group ID] = " & lGroupID & " AND [Buddy ID] = " & lBuddyID & " AND [Class ID] = " & lClassID, DB_Connection, adOpenKeyset, adLockOptimistic
    If RS.RecordCount > 0 Then
        RS.Fields("Name") = sName
        RS.Fields("Group ID") = lGroupID
        RS.Fields("Buddy ID") = lBuddyID
        'RS.Fields("Class ID") = lClassID
        RS.Fields("Attributes") = sAttributes
        RS.Update
        RS.Close
        FeedbagUpdateItem = FEEDBAG_STATUS_CODES_SUCCESS
        Exit Function
    Else
        FeedbagUpdateItem = FEEDBAG_STATUS_CODES_BAD_REQUEST
        Exit Function
        RS.Close
    End If
End Function

Public Sub FeedbagUpdateDatabase(oUser As clsAIMSession)
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Dim lRecordCount As Long
    RS.Open "SELECT * FROM [Feedbag] WHERE [ID] = " & oUser.ID, DB_Connection, adOpenKeyset, adLockOptimistic
    lRecordCount = RS.RecordCount
    RS.Close
    'Update Feedbag Info
    RS.Open "SELECT * FROM [Registration] WHERE [ScreenName] = '" & oUser.ScreenName & "'", DB_Connection, adOpenKeyset, adLockOptimistic
    RS.Fields("Feedbag Timestamp") = GetTimeStamp + 300 'lets add 5 minutes, this is so AIM wont be a little asswipe
    RS.Fields("Feedbag Items") = lRecordCount
    RS.Update
    RS.Close
End Sub

Public Sub InitFormSizes(frmForm As Form)
    Dim WindowSizes(1 To 4) As String
    On Error GoTo ErrInitFormSizes

    WindowSizes(1) = GetSetting("MyAIMServer", "Window Sizes", frmForm.Name & ".Top", "")
    WindowSizes(2) = GetSetting("MyAIMServer", "Window Sizes", frmForm.Name & ".Left", "")
    WindowSizes(3) = GetSetting("MyAIMServer", "Window Sizes", frmForm.Name & ".Height", "")
    WindowSizes(4) = GetSetting("MyAIMServer", "Window Sizes", frmForm.Name & ".Width", "")
    If WindowSizes(1) <> vbNullString Then: frmForm.Top = WindowSizes(1)
    If WindowSizes(2) <> vbNullString Then: frmForm.Left = WindowSizes(2)
    If WindowSizes(3) <> vbNullString Then: frmForm.Height = WindowSizes(3)
    If WindowSizes(4) <> vbNullString Then: frmForm.Width = WindowSizes(4)

    On Error GoTo 0
    Exit Sub

ErrInitFormSizes:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure InitFormSizes of modServer"

End Sub

Public Function MD5(ByVal TheString As String) As String
    Dim TheAnswer As Long
    Dim lngReturnValue As Long
    Dim strHash As String
    Dim hCryptProv As Long
    Dim hHash As Long
    Dim lngHashLen As Long

    On Error GoTo ErrMD5

    lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, CRYPT_NEWKEYSET) 'try to make a new key container
    If lngReturnValue = 0 Then
        lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, 0) 'try to get a handle to a key container that already exists, and if it fails...
        If lngReturnValue = 0 Then Exit Function
    End If
    lngReturnValue = CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash)
    lngReturnValue = CryptHashData(hHash, TheString, Len(TheString), 0)
    lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, vbNull, lngHashLen, 0)
    strHash = String(lngHashLen, vbNullChar)
    lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, strHash, lngHashLen, 0)
    If hHash <> 0 Then CryptDestroyHash hHash
    If hCryptProv <> 0 Then CryptReleaseContext hCryptProv, 0
    MD5 = strHash

    On Error GoTo 0
    Exit Function

ErrMD5:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure MD5 of modServer"


End Function

Public Function ErrMsg(ByRef sMessage As String, _
                       Optional bBold As Boolean = False, _
                       Optional lSize As Long = 8)
                       
    frmErrorLog.rtbServerLog.SelStart = Len(frmErrorLog.rtbServerLog.Text)
    frmErrorLog.rtbServerLog.SelColor = &HFF&
    frmErrorLog.rtbServerLog.SelFontSize = lSize
    frmErrorLog.rtbServerLog.SelBold = bBold
    frmErrorLog.rtbServerLog.SelText = "[" & Format(Time, "H:MM:SS AM/PM") & "]: " & sMessage & vbCrLf
    frmErrorLog.sbServerLog.Panels(1).Text = "Error Log: [" & (Len(frmErrorLog.rtbServerLog.Text) \ 1024) & " kb]"

End Function

Public Function LogMsg(ByRef sMessage As String, _
                       Optional lColor As Long = vbBlack, _
                       Optional bBold As Boolean = False, _
                       Optional lSize As Long = 8)
                       
    frmServerLog.rtbServerLog.SelStart = Len(frmServerLog.rtbServerLog.Text)
    frmServerLog.rtbServerLog.SelColor = lColor
    frmServerLog.rtbServerLog.SelFontSize = lSize
    frmServerLog.rtbServerLog.SelBold = bBold
    frmServerLog.rtbServerLog.SelText = "[" & Format(Time, "H:MM:SS AM/PM") & "]: " & sMessage & vbCrLf
    frmServerLog.sbServerLog.Panels(1).Text = "Server Log: [" & (Len(frmServerLog.rtbServerLog.Text) \ 1024) & " kb]"

End Function

Public Function TrimData(sData As String) As String
    TrimData = Replace$(LCase$(sData), " ", vbNullString)
End Function

Public Function IsBannedIP(ByVal IP As String) As Boolean
    Dim RS As ADODB.Recordset
    On Error GoTo ErrIsBannedIP

    Set RS = New ADODB.Recordset
    RS.Open "SELECT * FROM [Banned] WHERE [IP Address] = '" & IP & "'", DB_Connection, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        IsBannedIP = True
    Else
        IsBannedIP = False
    End If
    RS.Close

    On Error GoTo 0
    Exit Function

ErrIsBannedIP:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure IsBannedIP of modServer"

End Function

Public Function IsValidScreenName(ByVal Nickname As String) As Boolean
    Dim i As Integer, n As Byte
    On Error GoTo ErrIsValidScreenName

    'is it the correct length
    If Not (Len(Nickname) >= 2 And Len(Nickname) <= 18) Then
        IsValidScreenName = False
        Exit Function
    End If
    'does it start with a letter?
    n = Asc(Mid$(Nickname, 1, 1))
    If Not (n >= 65 And n <= 90 Or n >= 97 And n <= 122) Then
        IsValidScreenName = False
        Exit Function
    End If
    'does it contain any bad chars
    For i = 1 To Len(Nickname)
        n = Asc(Mid(Nickname, i, 1))
        Select Case n
            Case 65 To 90, 97 To 122, 48 To 57, 32, 46, 64, 95
            Case Else
                IsValidScreenName = False
                Exit Function
        End Select
    Next i
    IsValidScreenName = True

    On Error GoTo 0
    Exit Function

ErrIsValidScreenName:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure IsValidScreenName of modServer"

End Function

Public Function RegisterName(sIP As String, sName As String, sPassword As String, sConfirm As String, sEmail As String) As String
    On Error GoTo ErrRegisterName

    sName = DecodeStr(sName)
    sPassword = DecodeStr(sPassword)
    sConfirm = DecodeStr(sConfirm)
    sEmail = DecodeStr(sEmail)
    If IsValidScreenName(TrimData(sName)) = False Then
        RegisterName = "invalidname"
        Exit Function
    End If
    If Len(sPassword) > 20 Or Len(sPassword) < 3 Then
        RegisterName = "invalidpassword"
        Exit Function
    End If
    If sPassword <> sConfirm Then
        RegisterName = "invalidconfirm"
        Exit Function
    End If
    If Len(sEmail) < 3 Then
        RegisterName = "invalidemail"
        Exit Function
    End If
    If InStr(1, sEmail, "@") = 0 Or InStr(1, sEmail, ".") = 0 Then
        RegisterName = "invalidemail"
        Exit Function
    End If
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    'sScreenName = TrimData(sScreenName)
    RS.Open "SELECT * FROM [Registration] WHERE [Screenname] = '" & TrimData(sName) & "'", DB_Connection, adOpenKeyset, adLockOptimistic
    If RS.RecordCount > 0 Then
        RegisterName = "alreadyexist"
        RS.Close
        Exit Function
    End If
    RS.Close
    Set RS = Nothing
    Set RS = New Recordset
    RS.Open "SELECT * FROM [Registration]", DB_Connection, adOpenKeyset, adLockOptimistic
    RS.AddNew
    RS.Fields("Screenname") = TrimData(sName)
    RS.Fields("Formatted ScreenName") = Trim(sName)
    RS.Fields("Password") = sPassword
    RS.Fields("Temporary Evil") = 0
    RS.Fields("Permanent Evil") = 0
    RS.Fields("Email Address") = sEmail
    RS.Fields("Confirmed") = True
    RS.Fields("Last Signon Date") = Now
    RS.Fields("Creation Date") = Now
    RS.Fields("Logged IP Addresses") = sIP & ";"
    RS.Fields("Registered IP Address") = sIP
    RS.Update
    LogMsg "A new account [" & Trim(sName) & ":" & sPassword & "] at [" & sEmail & "] was just registered."
    RegisterName = "good"
    RS.Close
    Exit Function

    On Error GoTo 0
    Exit Function

ErrRegisterName:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure RegisterName of modServer"

End Function

Public Function GetAccountStatus(ByVal sScreenName As String) As LoginState
    Dim RS As ADODB.Recordset
    On Error GoTo ErrGetAccountStatus

    Set RS = New ADODB.Recordset
    sScreenName = TrimData(sScreenName)
    RS.Open "SELECT * FROM [Registration] WHERE [Screenname] = '" & sScreenName & "'", DB_Connection, adOpenKeyset, adLockReadOnly
    If IsValidScreenName(sScreenName) = False Then
        GetAccountStatus = LoginStateInvalid
        Exit Function
    End If
    If RS.RecordCount <= 0 Then
        GetAccountStatus = LoginStateUnregistered
        Exit Function
    End If
    If RS.Fields("Suspended") = True Then
        GetAccountStatus = LoginStateSuspended
        Exit Function
    End If
    If RS.Fields("Deleted") = True Then
        GetAccountStatus = LoginStateDeleted
        Exit Function
    End If
    GetAccountStatus = LoginStateGood
    RS.Close

    On Error GoTo 0
    Exit Function

ErrGetAccountStatus:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure GetAccountStatus of modServer"

End Function

Public Function CheckLogin(ByVal sScreenName As String, ByVal sPassword As String, ByVal sLoginTicket As String) As Boolean
    Dim RS As ADODB.Recordset
    On Error GoTo ErrCheckLogin

    Set RS = New ADODB.Recordset
    Dim sSpecialHash As String
    Dim sSpecialHash2 As String
    sScreenName = TrimData(sScreenName)
    RS.Open "SELECT * FROM [Registration] WHERE [Screenname] = '" & sScreenName & "'", DB_Connection, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        'To avoid having to check a certain TLV...lets support both
        'kinds of hashes in one function! YAY! Thank god im a lazy
        'bastard.
        sSpecialHash = MD5(sLoginTicket & RS.Fields("Password") & "AOL Instant Messenger (SM)")
        sSpecialHash2 = MD5(sLoginTicket & MD5(RS.Fields("Password")) & "AOL Instant Messenger (SM)")
        If sPassword = sSpecialHash Or sPassword = sSpecialHash2 Then
            CheckLogin = True
            RS.Close
            Exit Function
        Else
            CheckLogin = False
            Exit Function
            RS.Close
        End If
    Else
        CheckLogin = False
        Exit Function
        RS.Close
    End If
    RS.Close

    On Error GoTo 0
    Exit Function

ErrCheckLogin:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckLogin of modServer"

End Function

Public Sub SetupAccount(oAIMUser As clsAIMSession)
    Dim RS As ADODB.Recordset
    On Error GoTo ErrSetupAccount

    Set RS = New ADODB.Recordset
    RS.Open "SELECT * FROM [Registration] WHERE [Screenname] = '" & TrimData(oAIMUser.ScreenName) & "'", DB_Connection, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        oAIMUser.ID = RS.Fields("ID")
        oAIMUser.SignonTime = Now
        oAIMUser.SignonTimestamp = GetTimeStamp
        If RS.Fields("Internal") = True Then
            oAIMUser.UserClass = &H12&
        Else
            oAIMUser.UserClass = &H10&
        End If
        oAIMUser.EmailAddress = RS.Fields("Email Address")
        oAIMUser.Password = RS.Fields("Password")
        oAIMUser.FormattedScreenName = RS.Fields("Formatted ScreenName")
        oAIMUser.WarningLevel = RS.Fields("Temporary Evil")
        oAIMUser.Authorized = True
        'Make sure buddylist is in order
        Call FeedbagCheckMasterGroup(oAIMUser)
    End If
    RS.Close

    On Error GoTo 0
    Exit Sub

ErrSetupAccount:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure SetupAccount of modServer"

End Sub

Public Sub UpdateUserStatus(oUser As clsAIMSession)
    Dim RS As ADODB.Recordset
    Dim oAIM As clsAIMSession
    Dim i As Integer
    On Error GoTo ErrUpdateUserStatus

    Set RS = New ADODB.Recordset
    If Len(oUser.AwayMessage) > 0 Then
        If Not (oUser.UserClass And &H20) = &H20 Then
            oUser.UserClass = oUser.UserClass Or &H20
        End If
    Else
        If (oUser.UserClass And &H20) = &H20 Then
            oUser.UserClass = oUser.UserClass Xor &H20
        End If
    End If
    RS.Open "SELECT * FROM [Feedbag] WHERE [Name] = '" & TrimData(oUser.ScreenName) & "' AND [Class ID] = 0", DB_Connection, adOpenKeyset, adLockReadOnly
    For i = 1 To RS.RecordCount
        For Each oAIM In oAIMSessionManager
            If oAIM.ID = RS.Fields("ID") Then
                If oUser.SignedOn = True Then
                    mdiMyAIMServer.BOSServer.SendData oAIM.Index, 0, 2, BuddyArrived(oUser.FormattedScreenName, oUser.WarningLevel, oUser.UserClass, oUser.ShortCapabilities, oUser.Capabilities, DateDiff("S", oAIM.SignonTime, Now()), oAIM.SignonTimestamp)
                Else
                    mdiMyAIMServer.BOSServer.SendData oAIM.Index, 0, 2, BuddyDeparted(oUser.FormattedScreenName)
                End If
            End If
        Next oAIM
        RS.MoveNext
    Next i
    RS.Close
    RS.Open "SELECT * FROM [Feedbag] WHERE [ID] = " & oUser.ID & " AND [Class ID] = 0", DB_Connection, adOpenKeyset, adLockReadOnly
    For i = 1 To RS.RecordCount
        For Each oAIM In oAIMSessionManager
            If oAIM.ScreenName = TrimData(RS.Fields("Name")) Then
                If oAIM.SignedOn = True Then
                    mdiMyAIMServer.BOSServer.SendData oUser.Index, 0, 2, BuddyArrived(oAIM.FormattedScreenName, oAIM.WarningLevel, oAIM.UserClass, oAIM.ShortCapabilities, oAIM.Capabilities, DateDiff("S", oAIM.SignonTime, Now()), oAIM.SignonTimestamp)
                Else
                    mdiMyAIMServer.BOSServer.SendData oUser.Index, 0, 2, BuddyDeparted(oAIM.FormattedScreenName)
                End If
            End If
        Next oAIM
        RS.MoveNext
    Next i
    RS.Close

    On Error GoTo 0
    Exit Sub

ErrUpdateUserStatus:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure UpdateUserStatus of modServer"

End Sub

Public Function GetTimeStamp() As Long
    GetTimeStamp = DateDiff("S", "1/1/1970 00:00:00", Now())
End Function
