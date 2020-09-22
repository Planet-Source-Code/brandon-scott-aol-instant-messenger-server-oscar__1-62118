VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBinaryStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sBuffer As String
Private m_lPosition As Long
Private m_lLength As Long

Public Property Get Buffer() As String
    Buffer = m_sBuffer
End Property

Public Property Let Buffer(ByVal sBuffer As String)
    m_sBuffer = sBuffer
End Property

Public Property Get Position() As Long
    Position = m_lPosition
End Property

Public Property Let Position(ByVal lPosition As Long)
    m_lPosition = lPosition
End Property

Public Property Get Length() As Long
    Length = m_lLength
End Property

Public Property Let Length(ByVal lLength As Long)
    m_lLength = lLength
End Property

Public Function Read08() As Byte
    Dim bResult As Byte
    If m_lPosition + 1 <= (m_lLength + 1) Then
        bResult = Asc(Mid$(m_sBuffer$, m_lPosition, 1))
        m_lPosition = m_lPosition + 1
    End If
    Read08 = bResult
End Function

Public Function Read16() As Long
    Dim lResult As Long
    If m_lPosition + 2 <= (m_lLength + 1) Then
        lResult = GetWord(Mid$(m_sBuffer$, m_lPosition, 2))
        m_lPosition = m_lPosition + 2
    End If
    Read16 = lResult
End Function

Public Function Read32() As Double
    Dim dResult As Double
    If m_lPosition + 4 <= (m_lLength + 1) Then
        dResult = GetDWord(Mid$(m_sBuffer$, m_lPosition, 4))
        m_lPosition = m_lPosition + 4
    End If
    Read32 = dResult
End Function

Public Function Read08String() As String
    Dim bLength As Byte
    Dim sResult As String
    If m_lPosition + 1 <= (m_lLength + 1) Then
        bLength = Asc(Mid$(m_sBuffer$, m_lPosition, 1))
        m_lPosition = m_lPosition + 1
        If m_lPosition + bLength <= (m_lLength + 1) Then
            sResult = Mid$(m_sBuffer$, m_lPosition, bLength)
            m_lPosition = m_lPosition + bLength
        End If
    End If
    Read08String = sResult
End Function

Public Function Read16String() As String
    Dim lLength As Long
    Dim sResult As String
    If m_lPosition + 2 <= (m_lLength + 1) Then
        lLength = GetWord(Mid$(m_sBuffer$, m_lPosition, 2))
        m_lPosition = m_lPosition + 2
        If m_lPosition + lLength <= (m_lLength + 1) Then
            sResult = Mid$(m_sBuffer$, m_lPosition, lLength)
            m_lPosition = m_lPosition + lLength
        End If
    End If
    Read16String = sResult
End Function

Public Function Read32String() As String
    Dim dLength As Double
    Dim sResult As String
    If m_lPosition + 4 <= (m_lLength + 1) Then
        dLength = GetDWord(Mid$(m_sBuffer$, m_lPosition, 4))
        m_lPosition = m_lPosition + 4
        If m_lPosition + dLength <= (m_lLength + 1) Then
            sResult = Mid$(m_sBuffer$, m_lPosition, dLength)
            m_lPosition = m_lPosition + dLength
        End If
    End If
    Read32String = sResult
End Function

Public Function ReadString(lLength As Long) As String
    Dim sResult As String
    If m_lPosition + lLength <= (m_lLength + 1) Then
        sResult = Mid$(m_sBuffer$, m_lPosition, lLength)
        m_lPosition = m_lPosition + lLength
    End If
    ReadString = sResult
End Function

Public Sub Write08(bValue As Byte)
    m_sBuffer = m_sBuffer & Chr(bValue)
    m_lPosition = m_lPosition + 1
    m_lLength = m_lLength + 1
End Sub

Public Sub Write16(lValue As Long)
    m_sBuffer = m_sBuffer & Word(lValue)
    m_lPosition = m_lPosition + 2
    m_lLength = m_lLength + 2
End Sub

Public Sub Write32(dValue As Double)
    m_sBuffer = m_sBuffer & DWord(dValue)
    m_lPosition = m_lPosition + 4
    m_lLength = m_lLength + 4
End Sub

Public Sub Write08String(sValue As String)
    If Len(sValue) <= 255 Then
        m_sBuffer = m_sBuffer & Chr(Len(sValue)) & sValue
        m_lPosition = m_lPosition + 1 + Len(sValue)
        m_lLength = m_lLength + 1 + Len(sValue)
    End If
End Sub

Public Sub Write16String(sValue As String)
    If Len(sValue) <= 65535 Then
        m_sBuffer = m_sBuffer & Word(Len(sValue)) & sValue
        m_lPosition = m_lPosition + 2 + Len(sValue)
        m_lLength = m_lLength + 2 + Len(sValue)
    End If
End Sub

Public Sub Write32String(sValue As String)
    If Len(sValue) <= 4294967295# Then
        m_sBuffer = m_sBuffer & DWord(Len(sValue)) & sValue
        m_lPosition = m_lPosition + 4 + Len(sValue)
        m_lLength = m_lLength + 4 + Len(sValue)
    End If
End Sub

Public Function IsEnd() As Boolean
    If m_lPosition >= (m_lLength + 1) Then
        IsEnd = True
    End If
End Function

Public Sub LoadBuffer(sBuffer As String)
    m_sBuffer = sBuffer
    m_lLength = Len(sBuffer)
    m_lPosition = 1
End Sub

Public Function LoadFile(sFile As String) As Boolean
    On Error GoTo ErrReadFile
    Dim iFreeFile As Integer
    iFreeFile = FreeFile
    Open sFile For Binary As #iFreeFile
        Get #iFreeFile, , m_sBuffer
    Close #iFreeFile
    m_lLength = FileLen(sFile)
    m_lPosition = 0
    LoadFile = True
    On Error GoTo 0
    Exit Function
ErrReadFile:
    LoadFile = False
End Function

Private Sub Class_Initialize()
    m_lPosition = 1
End Sub
