VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAIMSessionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_colAIMSessions As Collection

Public Function Add(ByVal sScreenName As String, ByVal lAuthSocket As Long, ByVal sTicket As String) As clsAIMSession

    'create a new object
    Dim oAIMSession As clsAIMSession
    On Error GoTo ErrAdd

    Set oAIMSession = New clsAIMSession
    oAIMSession.ScreenName = TrimData(sScreenName)
    oAIMSession.AuthSocket = lAuthSocket
    oAIMSession.LoginTicket = sTicket
    ' add the object, with a key of the Socket name
    m_colAIMSessions.Add oAIMSession, CStr(TrimData(sScreenName))
    Set Add = oAIMSession
    Set oAIMSession = Nothing

    On Error GoTo 0
    Exit Function

ErrAdd:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure Add of clsAIMSessionManager"

    
End Function

Public Sub Remove(ByVal sScreenName As String)

    On Error GoTo ErrRemove

    m_colAIMSessions.Remove CStr(TrimData(sScreenName))

    On Error GoTo 0
    Exit Sub

ErrRemove:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure Remove of clsAIMSessionManager"


End Sub

Public Property Get Item(vntIndexKey As Variant) As clsAIMSession
    
    On Error Resume Next
    Set Item = m_colAIMSessions(vntIndexKey)
    
End Property


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = m_colAIMSessions.[_NewEnum]
    
End Property

Private Sub Class_Initialize()

    Set m_colAIMSessions = New Collection
    
End Sub

Private Sub Class_Terminate()

    Set m_colAIMSessions = Nothing
    
End Sub

Public Property Get Count() As Long

    On Error GoTo Err_Handler
    Count = m_colAIMSessions.Count
    Exit Property

Err_Handler:

    LogMsg "Error: Count() procedure in AIM Session Manager. Description: " & Err.Description
    
End Property
