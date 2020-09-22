VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettings 
   Caption         =   "Settings"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
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
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   9375
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   240
      ScaleHeight     =   6015
      ScaleWidth      =   8895
      TabIndex        =   1
      Top             =   600
      Width           =   8895
      Begin VB.Frame Frame4 
         Caption         =   "Error URLs"
         Height          =   4095
         Left            =   3840
         TabIndex        =   26
         Top             =   0
         Width           =   4935
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   240
            TabIndex        =   32
            Text            =   "http://www.xeons.net/aimerror.php?id=1"
            Top             =   600
            Width           =   4455
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   240
            TabIndex        =   31
            Text            =   "http://www.xeons.net/aimerror.php?id=2"
            Top             =   1200
            Width           =   4455
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   240
            TabIndex        =   30
            Text            =   "http://www.xeons.net/aimerror.php?id=3"
            Top             =   1800
            Width           =   4455
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   240
            TabIndex        =   29
            Text            =   "http://www.xeons.net/aimerror.php?id=4"
            Top             =   2400
            Width           =   4455
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   240
            TabIndex        =   28
            Text            =   "http://www.xeons.net/aimerror.php?id=5"
            Top             =   3000
            Width           =   4455
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   240
            TabIndex        =   27
            Text            =   "http://www.xeons.net/aimerror.php?id=7"
            Top             =   3600
            Width           =   4455
         End
         Begin VB.Label lblSettings 
            Caption         =   "Deleted Account URL:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   38
            Top             =   360
            Width           =   3975
         End
         Begin VB.Label lblSettings 
            Caption         =   "Bad Login URL:"
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   37
            Top             =   960
            Width           =   3975
         End
         Begin VB.Label lblSettings 
            Caption         =   "Invalid Password URL:"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   36
            Top             =   1560
            Width           =   3975
         End
         Begin VB.Label lblSettings 
            Caption         =   "Suspended Account URL:"
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   35
            Top             =   2160
            Width           =   3975
         End
         Begin VB.Label lblSettings 
            Caption         =   "Unregistered URL:"
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   34
            Top             =   2760
            Width           =   3975
         End
         Begin VB.Label lblSettings 
            Caption         =   "Password Change URL:"
            Height          =   255
            Index           =   16
            Left            =   240
            TabIndex        =   33
            Top             =   3360
            Width           =   3975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Location Rights"
         Height          =   1935
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   3615
         Begin VB.TextBox txtMaxFindByEmail 
            Height          =   285
            Left            =   2640
            TabIndex        =   21
            Text            =   "0"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtMaxCapabilitiesLength 
            Height          =   285
            Left            =   2640
            TabIndex        =   20
            Text            =   "0"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtMaxProfileLength 
            Height          =   285
            Left            =   2640
            TabIndex        =   19
            Text            =   "0"
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtMaxCertsLength 
            Height          =   285
            Left            =   2640
            TabIndex        =   18
            Text            =   "0"
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label lblSettings 
            Alignment       =   1  'Right Justify
            Caption         =   "Max Find by Email List:"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   25
            Top             =   1095
            Width           =   2295
         End
         Begin VB.Label lblSettings 
            Alignment       =   1  'Right Justify
            Caption         =   "Max Capabilities Length:"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   24
            Top             =   735
            Width           =   2295
         End
         Begin VB.Label lblSettings 
            Alignment       =   1  'Right Justify
            Caption         =   "Max Profile Length:"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   23
            Top             =   375
            Width           =   2295
         End
         Begin VB.Label lblSettings 
            Alignment       =   1  'Right Justify
            Caption         =   "Max Certs Length:"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   22
            Top             =   1455
            Width           =   2295
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "MOTD"
         Height          =   1575
         Left            =   120
         TabIndex        =   9
         Top             =   4320
         Width           =   3135
         Begin VB.CheckBox chkDisableAIMToday 
            Caption         =   "Disable AIM Today Window"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   1080
            Width           =   2655
         End
         Begin VB.TextBox txtAdSyncInterval 
            Height          =   285
            Left            =   2160
            TabIndex        =   11
            Text            =   "0"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtAdRotationInterval 
            Height          =   285
            Left            =   2160
            TabIndex        =   10
            Text            =   "0"
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblSettings 
            Alignment       =   1  'Right Justify
            Caption         =   "Ad Sync Interval:"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   13
            Top             =   735
            Width           =   1815
         End
         Begin VB.Label lblSettings 
            Alignment       =   1  'Right Justify
            Caption         =   "Ad Rotation Interval:"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   12
            Top             =   375
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Connection Info"
         Height          =   2175
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   3615
         Begin VB.TextBox txtServerIP 
            Height          =   285
            Left            =   1680
            TabIndex        =   16
            Text            =   "255.255.255.255"
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtAuthorizorPort 
            Height          =   285
            Left            =   1680
            TabIndex        =   5
            Text            =   "5190"
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox txtServicePort 
            Height          =   285
            Left            =   1680
            TabIndex        =   4
            Text            =   "5191"
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txtAddonPort 
            Height          =   285
            Left            =   1680
            TabIndex        =   3
            Text            =   "5192"
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label lblSettings 
            Alignment       =   1  'Right Justify
            Caption         =   "Server IP:"
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   15
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblSettings 
            Alignment       =   1  'Right Justify
            Caption         =   "Authorizor Port:"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   8
            Top             =   855
            Width           =   1215
         End
         Begin VB.Label lblSettings 
            Alignment       =   1  'Right Justify
            Caption         =   "Service Port:"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   7
            Top             =   1215
            Width           =   1215
         End
         Begin VB.Label lblSettings 
            Alignment       =   1  'Right Justify
            Caption         =   "Addon Port:"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   6
            Top             =   1575
            Width           =   1215
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11668
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
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
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub
