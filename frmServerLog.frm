VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmServerLog 
   Caption         =   "Server Log"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11115
   Icon            =   "frmServerLog.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   11115
   Begin RichTextLib.RichTextBox rtbServerLog 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   7858
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmServerLog.frx":1042
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbServerLog 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   5640
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   19076
            Text            =   "Server Log: [0 kb]"
            TextSave        =   "Server Log: [0 kb]"
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
Attribute VB_Name = "frmServerLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    InitFormSizes Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    rtbServerLog.Width = Me.Width - 120
    rtbServerLog.Height = Me.Height - 800
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Me.Hide
    mdiMyAIMServer.tbrServerControls.Buttons(6).Value = tbrUnpressed
End Sub
