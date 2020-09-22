VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAIMChatRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_sName As String
Private m_sOccupants As String
Private m_lExchange As Long
Private m_lInstance As Long


Public Property Get Name() As String

    Name = m_sName

End Property

Public Property Let Name(ByVal sName As String)

    m_sName = sName

End Property

Public Property Get Occupants() As String

    Occupants = m_sOccupants

End Property

Public Property Let Occupants(ByVal sOccupants As String)

    m_sOccupants = sOccupants

End Property

Public Property Get Exchange() As Long

    Exchange = m_lExchange

End Property

Public Property Let Exchange(ByVal lExchange As Long)

    m_lExchange = lExchange

End Property

Public Property Get Instance() As Long

    Instance = m_lInstance

End Property

Public Property Let Instance(ByVal lInstance As Long)

    m_lInstance = lInstance

End Property
