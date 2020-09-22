VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSendMessage 
   Caption         =   "Send Message"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7110
   Icon            =   "frmSendMessage.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5070
   ScaleWidth      =   7110
   Begin VB.Frame Frame2 
      Height          =   555
      Left            =   45
      TabIndex        =   4
      Top             =   4185
      Width           =   6900
      Begin VB.CommandButton cmdSendMessage 
         Caption         =   "Send Messages"
         Height          =   285
         Left            =   90
         TabIndex        =   7
         Top             =   180
         Width           =   1410
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   300
         Left            =   4365
         TabIndex        =   6
         Top             =   180
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   1530
         TabIndex        =   5
         Top             =   180
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Message:"
      Height          =   4245
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   6900
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Left            =   75
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "frmSendMessage.frx":1042
         Top             =   225
         Width           =   6675
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   1830
         Left            =   75
         TabIndex        =   2
         Top             =   2340
         Width           =   6735
         ExtentX         =   11880
         ExtentY         =   3228
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   4785
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   12012
            Text            =   "Awaiting Message Input"
            TextSave        =   "Awaiting Message Input"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSendMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Open App.Path & "\usermessages.txt" For Output As #1
    Print #1, Text1
    Close #1
    'End If
    Me.Hide
    mdiMyAIMServer.tbrServerControls.Buttons(3).Value = tbrUnpressed
End Sub

Private Sub cmdSendMessage_Click()
    If Text1.Text = "" Then Exit Sub
    If oAIMSessionManager.Count = 0 Then Exit Sub 'nobody on...no point
    Dim oAIM As clsAIMSession
    Dim i As Integer
    ProgressBar1.Max = oAIMSessionManager.Count
    For Each oAIM In oAIMSessionManager
        i = i + 1
        ProgressBar1.Value = i
        mdiMyAIMServer.BOSServer.SendData oAIM.Index, 0, 2, IcbmToClient(GRICBMCookie, "AIM Server Administrator", Text1.Text)
    Next oAIM
End Sub

Private Sub Form_Initialize()
    WebBrowser1.Navigate "about:blank"
    Do
        DoEvents
    Loop Until WebBrowser1.Busy = False
End Sub

Private Sub Form_Load()
    InitFormSizes Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Frame1.Width = Me.Width - 200
    Frame1.Height = Me.Height - 1300
    Text1.Width = Me.Width - 350
    Text1.Height = Me.Height / 2 - 450
    WebBrowser1.Top = Me.Height / 2 - 175
    WebBrowser1.Width = Me.Width - 350
    WebBrowser1.Height = Me.Height / 2 - 1200
    Frame2.Top = Me.Height - 1350
    Frame2.Width = Me.Width - 200
    cmdClose.Left = Me.Width - 1350
    ProgressBar1.Width = Me.Width - 2950
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Open App.Path & "\usermessages.txt" For Output As #1
        Print #1, Text1.Text
    Close #1
    'End If
    Me.Hide
    mdiMyAIMServer.tbrServerControls.Buttons(3).Value = tbrUnpressed
End Sub

Private Sub Text1_Change()
    On Error Resume Next
    WebBrowser1.Document.body.innerHTML = Text1.Text
End Sub
