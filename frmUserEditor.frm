VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Account Editor"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   8895
   Begin VB.CommandButton cmdDeleteAccount 
      Caption         =   "Delete Account"
      Height          =   375
      Left            =   4320
      TabIndex        =   39
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdResetAccount 
      Caption         =   "Reset Account"
      Height          =   375
      Left            =   5880
      TabIndex        =   38
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaveAccount 
      Caption         =   "Save Account"
      Height          =   375
      Left            =   7440
      TabIndex        =   37
      Top             =   5880
      Width           =   1335
   End
   Begin VB.PictureBox picTabs 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   5055
      Index           =   0
      Left            =   2880
      ScaleHeight     =   5055
      ScaleWidth      =   5775
      TabIndex        =   12
      Top             =   600
      Width           =   5775
      Begin VB.Frame Frame1 
         Caption         =   "Account Flags"
         Height          =   855
         Left            =   120
         TabIndex        =   33
         Top             =   3720
         Width           =   5535
         Begin VB.CheckBox chkInternal 
            Caption         =   "Internal (Administrator)"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   360
            Width           =   2175
         End
         Begin VB.CheckBox chkDeleted 
            Caption         =   "Deleted"
            Height          =   255
            Left            =   4320
            TabIndex        =   35
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chkSuspended 
            Caption         =   "Suspended"
            Height          =   255
            Left            =   2760
            TabIndex        =   34
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame grpUserInformation 
         Caption         =   "User Information"
         Height          =   2055
         Left            =   120
         TabIndex        =   22
         Top             =   0
         Width           =   5535
         Begin VB.TextBox txtPassword 
            Height          =   285
            Left            =   2160
            TabIndex        =   26
            Top             =   1170
            Width           =   2055
         End
         Begin VB.TextBox txtEmailAddress 
            Height          =   285
            Left            =   2160
            TabIndex        =   25
            Top             =   1575
            Width           =   2055
         End
         Begin VB.TextBox txtFormattedNick 
            Height          =   285
            Left            =   2160
            TabIndex        =   24
            Top             =   765
            Width           =   2055
         End
         Begin VB.TextBox txtNick 
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label lblPasswordLength 
            Alignment       =   1  'Right Justify
            Caption         =   "Password:"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   1185
            Width           =   1815
         End
         Begin VB.Label lblEmailAddress 
            Alignment       =   1  'Right Justify
            Caption         =   "Email Address:"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   1590
            Width           =   1815
         End
         Begin VB.Label lblFormattedScreenName 
            Alignment       =   1  'Right Justify
            Caption         =   "Formatted Screenname:"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   780
            Width           =   1815
         End
         Begin VB.Label lblScreenName 
            Alignment       =   1  'Right Justify
            Caption         =   "Screenname:"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   375
            Width           =   1815
         End
      End
      Begin VB.Frame grpAccountInfo 
         Caption         =   "Account Signon Information"
         Height          =   1455
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   5535
         Begin VB.Label lblAccountAge 
            Alignment       =   1  'Right Justify
            Caption         =   "Account Age:"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   600
            Width           =   2265
         End
         Begin VB.Label AccountAgeField 
            Caption         =   "Unknown..."
            Height          =   195
            Left            =   2640
            TabIndex        =   20
            Top             =   600
            Width           =   2700
         End
         Begin VB.Label CreatedOnField 
            Caption         =   "Unknown..."
            Height          =   195
            Left            =   2640
            TabIndex        =   19
            Top             =   360
            Width           =   2700
         End
         Begin VB.Label lblCreatedOn 
            Alignment       =   1  'Right Justify
            Caption         =   "Account Creation Date:"
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   2265
         End
         Begin VB.Label InactiveField 
            Caption         =   "Unknown..."
            Height          =   195
            Left            =   2640
            TabIndex        =   17
            Top             =   1080
            Width           =   2700
         End
         Begin VB.Label lblInactivity 
            Alignment       =   1  'Right Justify
            Caption         =   "Inactivity:"
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   1080
            Width           =   2265
         End
         Begin VB.Label LastLoginField 
            Caption         =   "Unknown..."
            Height          =   195
            Left            =   2640
            TabIndex        =   15
            Top             =   840
            Width           =   2700
         End
         Begin VB.Label lblLastLogin 
            Alignment       =   1  'Right Justify
            Caption         =   "Last Login Date:"
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   840
            Width           =   2265
         End
      End
   End
   Begin VB.PictureBox picTabs 
      BorderStyle     =   0  'None
      Height          =   5055
      Index           =   1
      Left            =   2880
      ScaleHeight     =   5055
      ScaleWidth      =   5775
      TabIndex        =   0
      Top             =   600
      Width           =   5775
      Begin VB.Frame grpIPRestrictions 
         Caption         =   "IP Access Restrictions"
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   5535
         Begin VB.CommandButton cmdRemoveAccessItem 
            Caption         =   "Remove"
            Height          =   345
            Left            =   3000
            TabIndex        =   4
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton cmdClearAccessList 
            Caption         =   "Clear"
            Height          =   345
            Left            =   3000
            TabIndex        =   3
            Top             =   360
            Width           =   2295
         End
         Begin VB.ListBox lstIPAccessRestrictions 
            Height          =   1695
            IntegralHeight  =   0   'False
            ItemData        =   "frmUserEditor.frx":0000
            Left            =   240
            List            =   "frmUserEditor.frx":0002
            TabIndex        =   2
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame grpLoggedIPs 
         Caption         =   "Logged IP Addresses"
         Height          =   2415
         Left            =   120
         TabIndex        =   5
         Top             =   2400
         Width           =   5535
         Begin VB.ListBox lstLoggedIPAddresses 
            Height          =   1815
            IntegralHeight  =   0   'False
            ItemData        =   "frmUserEditor.frx":0004
            Left            =   240
            List            =   "frmUserEditor.frx":0006
            TabIndex        =   11
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton cmdBanIP 
            Caption         =   "Ban IP []..."
            Height          =   345
            Left            =   3000
            TabIndex        =   10
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton cmdAddToIPRestrict 
            Caption         =   "Add to IP Restriction List..."
            Height          =   345
            Left            =   3000
            TabIndex        =   9
            Top             =   720
            Width           =   2295
         End
         Begin VB.CommandButton cmdClearLoggedList 
            Caption         =   "Clear Logged IP List"
            Height          =   345
            Left            =   3000
            TabIndex        =   8
            Top             =   1800
            Width           =   2295
         End
         Begin VB.CommandButton cmdRemoveIP 
            Caption         =   "Remove []..."
            Height          =   345
            Left            =   3000
            TabIndex        =   7
            Top             =   1440
            Width           =   2295
         End
         Begin VB.CommandButton cmdAddIPToLog 
            Caption         =   "Add IP to Log..."
            Height          =   345
            Left            =   3000
            TabIndex        =   6
            Top             =   1080
            Width           =   2295
         End
      End
   End
   Begin MSComctlLib.TabStrip AccountTab 
      Height          =   5655
      Left            =   2760
      TabIndex        =   31
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   9975
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Access Rights"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgAccounts 
      Left            =   360
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserEditor.frx":0008
            Key             =   "lvl1"
            Object.Tag             =   "lvl1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserEditor.frx":05A2
            Key             =   "lvl2"
            Object.Tag             =   "lvl2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserEditor.frx":0B3C
            Key             =   "lvl3"
            Object.Tag             =   "lvl3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserEditor.frx":10D6
            Key             =   "lvl4"
            Object.Tag             =   "lvl4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserEditor.frx":1670
            Key             =   "lvl5"
            Object.Tag             =   "lvl5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserEditor.frx":1C0A
            Key             =   "down"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserEditor.frx":205C
            Key             =   "right"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwAccounts 
      Height          =   6135
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   10821
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   0
      LabelEdit       =   1
      Style           =   1
      ImageList       =   "imgAccounts"
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   1
   End
End
Attribute VB_Name = "frmUserEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub tvwAccounts_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Bold = False
    Node.Image = "right"
End Sub

Private Sub tvwAccounts_Expand(ByVal Node As MSComctlLib.Node)
    Node.Bold = True
    Node.Image = "down"
End Sub

Private Sub AccountTab_Click()
    Dim i As Integer
    For i = picTabs.LBound To picTabs.UBound
        picTabs(i).Visible = False
    Next i
    picTabs(AccountTab.SelectedItem.Index - 1).Visible = True
End Sub

Public Sub RefreshAccountList()

    On Error GoTo ErrRefreshAccountList
    
    Dim i As Long
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    'Clear list
    tvwAccounts.Nodes.Clear
    'Add level headings
    tvwAccounts.Nodes.Add , , "accounts", "Accounts", "right"
    'Do the SQL query
    RS.Open "SELECT * FROM [Registration] ORDER BY [ScreenName]", DB_Connection, adOpenKeyset, adLockOptimistic
    'Cycle through records and add them
    For i = 1 To RS.RecordCount
        tvwAccounts.Nodes.Add "accounts", tvwChild, RS.Fields("ScreenName"), RS.Fields("Formatted ScreenName"), "lvl2"
        RS.MoveNext
    Next i
    'Close the record
    RS.Close
    'Expand the nodes for convenience
    For i = 1 To tvwAccounts.Nodes.Count
        tvwAccounts.Nodes(i).Expanded = True
    Next i
    'Clear out some memory
    Set RS = Nothing

    On Error GoTo 0
    Exit Sub

ErrRefreshAccountList:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure RefreshAccountList of frmUserEditor"

End Sub

Private Sub Form_Load()
    Call RefreshAccountList
End Sub

Private Sub tvwAccounts_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    RS.Open "SELECT * FROM [Registration] WHERE [ScreenName] = '" & Node.Key & "'", DB_Connection, adOpenKeyset, adLockOptimistic
    If RS.RecordCount > 0 Then
        txtNick.Text = RS.Fields("ScreenName")
        txtFormattedNick = RS.Fields("Formatted ScreenName")
        txtPassword.Text = RS.Fields("Password")
        txtEmailAddress = RS.Fields("Email Address")
    End If
EndItAll:
    RS.Close
    Set RS = Nothing
End Sub
