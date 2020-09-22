VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAIMService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lGroup As Long
Private m_sCookie As String
Private m_iIndex As Integer

Public Property Get Group() As Long

    Group = m_lGroup

End Property

Public Property Let Group(ByVal lGroup As Long)

    m_lGroup = lGroup

End Property

Public Property Get Cookie() As String

    Cookie = m_sCookie

End Property

Public Property Let Cookie(ByVal sCookie As String)

    m_sCookie = sCookie

End Property

Public Property Get Index() As Integer

    Index = m_iIndex

End Property

Public Property Let Index(ByVal iIndex As Integer)

    m_iIndex = iIndex

End Property
