VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTlvBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lCount As Long
Private m_lLength As Long
Private m_oTLVs() As AIM_TLV

Public Property Get Count() As Long
    Count = m_lCount
End Property

Public Property Let Count(ByVal lCount As Long)
    m_lCount = lCount
End Property

Public Property Get Length() As Long
    Length = m_lLength
End Property

Public Property Let Length(ByVal lLength As Long)
    m_lLength = lLength
End Property

Public Sub AddTLV(lType As Long, sValue As String)
    m_lCount = m_lCount + 1
    ReDim Preserve m_oTLVs(0 To m_lCount)
    m_oTLVs(m_lCount).lngType = lType
    m_oTLVs(m_lCount).lngLength = Len(sValue)
    m_oTLVs(m_lCount).strValue = sValue
End Sub

Public Sub InitBlock(sData As clsBinaryStream)
    Dim i As Long
    m_lCount = sData.Read16
    ReDim m_oTLVs(0 To m_lCount)
    For i = 0 To m_lCount - 1
        m_oTLVs(i).lngType = sData.Read16
        m_oTLVs(i).lngLength = sData.Read16
        m_oTLVs(i).strValue = sData.ReadString(m_oTLVs(i).lngLength)
    Next i
End Sub

Public Sub InitChain(sData As clsBinaryStream)
    Dim i As Long
    Do Until sData.IsEnd = True
        ReDim Preserve m_oTLVs(0 To m_lCount)
        m_oTLVs(i).lngType = sData.Read16
        m_oTLVs(i).lngLength = sData.Read16
        m_oTLVs(i).strValue = sData.ReadString(m_oTLVs(i).lngLength)
        m_lCount = m_lCount + 1
    Loop
End Sub

Public Function IsTLVPresent(lType As Long) As Boolean
    Dim i As Long
    For i = 0 To m_lCount
        If m_oTLVs(i).lngType = lType Then
           IsTLVPresent = True
           Exit Function
        End If
    Next i
    IsTLVPresent = False
End Function

Public Function GetTLVAs08(lType As Long) As Byte
    Dim i As Long
    For i = 0 To m_lCount
        If m_oTLVs(i).lngType = lType Then
           GetTLVAs08 = Asc(m_oTLVs(i).strValue)
           Exit Function
        End If
    Next i
End Function

Public Function GetTLVAs16(lType As Long) As Long
    Dim i As Long
    For i = 0 To m_lCount
        If m_oTLVs(i).lngType = lType Then
           GetTLVAs16 = GetWord(m_oTLVs(i).strValue)
           Exit Function
        End If
    Next i

End Function

Public Function GetTLVAs32(lType As Long) As Double
    Dim i As Long
    For i = 0 To m_lCount
        If m_oTLVs(i).lngType = lType Then
           GetTLVAs32 = GetDWord(m_oTLVs(i).strValue)
           Exit Function
        End If
    Next i
End Function

Public Function GetTLVAsString(lType As Long) As String
    Dim i As Long
    For i = 0 To m_lCount
        If m_oTLVs(i).lngType = lType Then
           GetTLVAsString = m_oTLVs(i).strValue
           Exit Function
        End If
    Next i
End Function

Private Sub Class_Initialize()

    ReDim m_oTLVs(0)
    
End Sub
