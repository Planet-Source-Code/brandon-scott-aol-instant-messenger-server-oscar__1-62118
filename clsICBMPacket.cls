VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICBMPacket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_Cookie As String
Private m_Channel As Long
Private m_ScreenNameLength As Long
Private m_ScreenName As String
Private m_WarningLevel As Long
Private m_InfoBlockTLVCount As Long
Private m_InfoBlockTLVChain() As AIM_TLV
Private m_IcbmTLVChain() As AIM_TLV

Public Property Let Cookie(ByVal vData As String)
    m_Cookie = vData
End Property

Public Property Get Cookie() As String
    Cookie = m_Cookie
End Property

Public Property Let Channel(ByVal vData As Long)
    m_Channel = vData
End Property

Public Property Get Channel() As Long
    Channel = m_Channel
End Property

Public Property Let ScreenName(ByVal vData As String)
    m_ScreenName = vData
End Property

Public Property Get ScreenName() As String
    ScreenName = m_ScreenName
End Property

Public Property Let WarningLevel(ByVal vData As Long)
    m_WarningLevel = vData
End Property

Public Property Get WarningLevel() As Long
    WarningLevel = m_WarningLevel
End Property

Public Property Let Message(ByVal vData As String)
    Dim i As Integer
   On Error GoTo Err_Proc

    For i = 0 To UBound(m_IcbmTLVChain)
        If m_IcbmTLVChain(i).lngType = 2 Then
            m_IcbmTLVChain(i).lngLength = Len(vData)
            m_IcbmTLVChain(i).strValue = vData
            Exit Property
        End If
    Next i

Exit_Proc:
   Exit Property

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsICBMPacket", "Let Message"
   Err.Clear
   Resume Exit_Proc

End Property

Public Property Get Message() As String
    Dim i As Integer
   On Error GoTo Err_Proc

    For i = 0 To UBound(m_IcbmTLVChain)
        If m_IcbmTLVChain(i).lngType = 2 Then
            Message = m_IcbmTLVChain(i).strValue
            Exit Property
        End If
    Next i

Exit_Proc:
   Exit Property

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsICBMPacket", "Get Message"
   Err.Clear
   Resume Exit_Proc

End Property

Public Property Let HostAck(ByVal vData As Boolean)
    Dim i As Integer
   On Error GoTo Err_Proc

    For i = 0 To UBound(m_IcbmTLVChain)
        If m_IcbmTLVChain(i).lngType = 3 Then
            If vData = False Then
                m_IcbmTLVChain(i).lngType = 0
                Exit Property
            Else
                Exit Property
            End If
        End If
    Next i
    If vData = True Then
        ReDim Preserve m_IcbmTLVChain(UBound(m_IcbmTLVChain) + 1)
        m_IcbmTLVChain(UBound(m_IcbmTLVChain)).lngType = 3
        m_IcbmTLVChain(UBound(m_IcbmTLVChain)).lngLength = 0
        m_IcbmTLVChain(UBound(m_IcbmTLVChain)).strValue = vbNullString
    End If

Exit_Proc:
   Exit Property

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsICBMPacket", "Let HostAck"
   Err.Clear
   Resume Exit_Proc

End Property

Public Property Get HostAck() As Boolean
    Dim i As Integer
   On Error GoTo Err_Proc

    For i = 0 To UBound(m_IcbmTLVChain)
        If m_IcbmTLVChain(i).lngType = 3 Then
            HostAck = True
            Exit Property
        End If
    Next i

Exit_Proc:
   Exit Property

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsICBMPacket", "Get HostAck"
   Err.Clear
   Resume Exit_Proc

End Property

Public Property Let AutoResponse(ByVal vData As Boolean)
    Dim i As Integer
   On Error GoTo Err_Proc

    For i = 0 To UBound(m_IcbmTLVChain)
        If m_IcbmTLVChain(i).lngType = 4 Then
            If vData = False Then
                m_IcbmTLVChain(i).lngType = 0
                Exit Property
            Else
                Exit Property
            End If
        End If
    Next i
    If vData = True Then
        ReDim Preserve m_IcbmTLVChain(UBound(m_IcbmTLVChain) + 1)
        m_IcbmTLVChain(UBound(m_IcbmTLVChain)).lngType = 4
        m_IcbmTLVChain(UBound(m_IcbmTLVChain)).lngLength = 0
        m_IcbmTLVChain(UBound(m_IcbmTLVChain)).strValue = vbNullString
    End If

Exit_Proc:
   Exit Property

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsICBMPacket", "Let AutoResponse"
   Err.Clear
   Resume Exit_Proc

End Property

Public Property Get AutoResponse() As Boolean
    Dim i As Integer
   On Error GoTo Err_Proc

    For i = 0 To UBound(m_IcbmTLVChain)
        If m_IcbmTLVChain(i).lngType = 4 Then
            AutoResponse = True
            Exit Property
        End If
    Next i

Exit_Proc:
   Exit Property

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsICBMPacket", "Get AutoResponse"
   Err.Clear
   Resume Exit_Proc

End Property

Public Property Let Data(ByVal vData As String)
    Dim i As Integer
   On Error GoTo Err_Proc

    For i = 0 To UBound(m_IcbmTLVChain)
        If m_IcbmTLVChain(i).lngType = 5 Then
            m_IcbmTLVChain(i).lngLength = Len(vData)
            m_IcbmTLVChain(i).strValue = vData
            Exit Property
        End If
    Next i

Exit_Proc:
   Exit Property

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsICBMPacket", "Let Data"
   Err.Clear
   Resume Exit_Proc

End Property

Public Property Get Data() As String
    Dim i As Integer
   On Error GoTo Err_Proc

    For i = 0 To UBound(m_IcbmTLVChain)
        If m_IcbmTLVChain(i).lngType = 5 Then
            Data = m_IcbmTLVChain(i).strValue
            Exit Property
        End If
    Next i

Exit_Proc:
   Exit Property

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsICBMPacket", "Get Data"
   Err.Clear
   Resume Exit_Proc

End Property

Public Function HandleIncomingICBMPacket(ByVal Data As String) As Long
    Dim l As Long, i As Integer
   On Error GoTo Err_Proc

    l = 1
    m_Cookie = Mid$(Data, l, 8): l = l + 8
    m_Channel = GetWord(Mid$(Data, l, 2)): l = l + 2
    m_ScreenNameLength = Asc(Mid$(Data, l, 2)): l = l + 1
    m_ScreenName = Mid$(Data, l, m_ScreenNameLength): l = l + m_ScreenNameLength
    m_WarningLevel = GetWord(Mid$(Data, l, 2)): l = l + 2
    m_InfoBlockTLVCount = GetWord(Mid$(Data, l, 2)): l = l + 2
    ReDim m_InfoBlockTLVChain(1 To m_InfoBlockTLVCount)
    For i = 1 To m_InfoBlockTLVCount
        m_InfoBlockTLVChain(i).lngType = GetWord(Mid$(Data, l, 2)): l = l + 2
        m_InfoBlockTLVChain(i).lngLength = GetWord(Mid$(Data, l, 2)): l = l + 2
        m_InfoBlockTLVChain(i).strValue = Mid$(Data, l, m_InfoBlockTLVChain(i).lngLength): l = l + m_InfoBlockTLVChain(i).lngLength
    Next i
    Do While l <= Len(Data)
        ReDim Preserve m_IcbmTLVChain(UBound(m_IcbmTLVChain) + 1)
        m_IcbmTLVChain(UBound(m_IcbmTLVChain)).lngType = GetWord(Mid$(Data, l, 2)): l = l + 2
        m_IcbmTLVChain(UBound(m_IcbmTLVChain)).lngLength = GetWord(Mid$(Data, l, 2)): l = l + 2
        m_IcbmTLVChain(UBound(m_IcbmTLVChain)).strValue = Mid$(Data, l, m_IcbmTLVChain(UBound(m_IcbmTLVChain)).lngLength): l = l + m_IcbmTLVChain(UBound(m_IcbmTLVChain)).lngLength
    Loop

Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsICBMPacket", "HandleIncomingICBMPacket"
   Err.Clear
   Resume Exit_Proc

End Function

Public Function RebuildIncomingICBMPacket() As String
    Dim strRebuild As String
    Dim i As Integer
    Dim lngType As Long, lngLength As Long, strValue As String
   On Error GoTo Err_Proc

    strRebuild = m_Cookie & Word(Channel) & SByte(m_ScreenName) & Word(m_WarningLevel)
    strRebuild = strRebuild & Word(m_InfoBlockTLVCount)
    For i = 1 To m_InfoBlockTLVCount
        lngType = m_InfoBlockTLVChain(i).lngType
        lngLength = m_InfoBlockTLVChain(i).lngLength
        strValue = m_InfoBlockTLVChain(i).strValue
        strRebuild = strRebuild & Word(lngType) & Word(lngLength) & strValue
    Next i
    For i = 1 To UBound(m_IcbmTLVChain)
        lngType = m_IcbmTLVChain(i).lngType
        lngLength = m_IcbmTLVChain(i).lngLength
        strValue = m_IcbmTLVChain(i).strValue
        If lngType <> 0 Then
            strRebuild = strRebuild & Word(lngType) & Word(lngLength) & strValue
        End If
    Next i
    RebuildIncomingICBMPacket = strRebuild

Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsICBMPacket", "RebuildIncomingICBMPacket"
   Err.Clear
   Resume Exit_Proc

End Function

Public Function HandleOutgoingICBMPacket(ByVal Data As String) As Long
    Dim l As Long
   On Error GoTo Err_Proc

    l = 1
    m_Cookie = Mid$(Data, l, 8): l = l + 8
    m_Channel = GetWord(Mid$(Data, l, 2)): l = l + 2
    m_ScreenNameLength = Asc(Mid$(Data, l, 2)): l = l + 1
    m_ScreenName = Mid$(Data, l, m_ScreenNameLength): l = l + m_ScreenNameLength
    Do While l <= Len(Data)
        ReDim Preserve m_IcbmTLVChain(UBound(m_IcbmTLVChain) + 1)
        m_IcbmTLVChain(UBound(m_IcbmTLVChain)).lngType = GetWord(Mid$(Data, l, 2)): l = l + 2
        m_IcbmTLVChain(UBound(m_IcbmTLVChain)).lngLength = GetWord(Mid$(Data, l, 2)): l = l + 2
        m_IcbmTLVChain(UBound(m_IcbmTLVChain)).strValue = Mid$(Data, l, m_IcbmTLVChain(UBound(m_IcbmTLVChain)).lngLength): l = l + m_IcbmTLVChain(UBound(m_IcbmTLVChain)).lngLength
    Loop

Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsICBMPacket", "HandleOutgoingICBMPacket"
   Err.Clear
   Resume Exit_Proc

End Function

Public Function RebuildOutgoingICBMPacket() As String
    Dim strRebuild As String
    Dim i As Integer
    Dim lngType As Long, lngLength As Long, strValue As String
   On Error GoTo Err_Proc

    strRebuild = m_Cookie & Word(Channel) & SByte(m_ScreenName)
    For i = 1 To UBound(m_IcbmTLVChain)
        lngType = m_IcbmTLVChain(i).lngType
        lngLength = m_IcbmTLVChain(i).lngLength
        strValue = m_IcbmTLVChain(i).strValue
        If lngType <> 0 Then
            If AutoResponse = True Then
                If lngType = 2 Or lngType = 4 Then
                    strRebuild = strRebuild & Word(lngType) & Word(lngLength) & strValue
                End If
            Else
                strRebuild = strRebuild & Word(lngType) & Word(lngLength) & strValue
            End If
        End If
    Next i
    RebuildOutgoingICBMPacket = strRebuild

Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsICBMPacket", "RebuildOutgoingICBMPacket"
   Err.Clear
   Resume Exit_Proc

End Function

Private Sub Class_Initialize()
    ReDim m_IcbmTLVChain(0)
End Sub

Private Sub Err_Handler(Optional ByVal vblnDisplayError As Boolean = True, _
                       Optional ByVal vstrErrNumber As String = vbNullString, _
                       Optional ByVal vstrErrDescription As String = vbNullString, _
                       Optional ByVal vstrModuleName As String = vbNullString, _
                       Optional ByVal vstrProcName As String = vbNullString)

  Dim strTemp As String

   On Error Resume Next
   '// Purpose: Error handling - On Error

    strTemp = "Error occured: "
    If LenB(vstrErrNumber) Then strTemp = strTemp & vstrErrNumber & vbNewLine Else strTemp = strTemp & vbNewLine
    If LenB(vstrErrDescription) Then strTemp = strTemp & "Description: " & vstrErrDescription & vbNewLine
    If LenB(vstrModuleName) Then strTemp = strTemp & "Module: " & vstrModuleName & vbNewLine
    If LenB(vstrProcName) Then strTemp = strTemp & "Function: " & vstrProcName
    ErrMsg strTemp

End Sub
