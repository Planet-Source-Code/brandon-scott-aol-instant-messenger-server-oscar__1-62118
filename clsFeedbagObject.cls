VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFeedbagObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sName As String
Private m_lParentID As Long
Private m_lBuddyID As Long
Private m_sExtraData As String
Private m_lClassID As Long

Public Property Get Name() As String

    Name = m_sName

End Property

Public Property Let Name(ByVal sName As String)

    m_sName = sName

End Property

Public Property Get ParentID() As Long

    ParentID = m_lParentID

End Property

Public Property Let ParentID(ByVal lParentID As Long)

    m_lParentID = lParentID

End Property

Public Property Get BuddyID() As Long

    BuddyID = m_lBuddyID

End Property

Public Property Let BuddyID(ByVal lBuddyID As Long)

    m_lBuddyID = lBuddyID

End Property

Public Property Get ClassID() As Long

    ClassID = m_lClassID

End Property

Public Property Let ClassID(ByVal lClassID As Long)

    m_lClassID = lClassID

End Property

Public Property Get ExtraData() As String

    ExtraData = m_sExtraData

End Property

Public Property Let ExtraData(ByVal sExtraData As String)

    m_sExtraData = sExtraData

End Property
