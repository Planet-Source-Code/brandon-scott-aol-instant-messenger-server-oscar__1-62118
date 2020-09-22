Attribute VB_Name = "modAIM"
Option Explicit

Public Enum LoginState
    LoginStateGood = 0
    LoginStateInvalidPassword = 1
    LoginStateSuspended = 2
    LoginStateDeleted = 3
    LoginStateUnregistered = 4
    LoginStateInvalid = 5
End Enum

Public Type AIMRateClass
    ClassID As Long
    windowsize As Double
    clearthreshold As Double
    alertthreshold As Double
    limitthreshold As Double
    minaverage As Double
    maxaverage As Double
    lastaverage As Double
    delta As Double
    islimiting As Boolean
End Type

Public Type AIMClientInfo
    sVersion As String
    lClientID As Long
    lMajorVersion As Long
    lMinorVersion As Long
    lPointVersion As Long
    lBuildNumber As Long
    dDistributionChannel As Double
End Type

Public Type AIM_TLV
    lngType As Long
    lngLength As Long
    strValue As String
End Type

Public Function GRInteger(LowerBound As Long, UpperBound As Long) As Long
    Randomize
    GRInteger = Int((UpperBound - LowerBound + 1) * Rnd + LowerBound)
End Function

Public Function GRTicket() As String
    Dim i As Integer
    For i = 1 To 10
        GRTicket = GRTicket & Chr(GRInteger(48, 57))
    Next i
End Function

Public Function GRCookie() As String
    Dim i As Integer
    For i = 1 To 256
        GRCookie = GRCookie & Chr(GRInteger(0, 255))
    Next i
End Function

Public Function GRICBMCookie() As String
    Dim i As Integer
    For i = 1 To 8
        GRICBMCookie = GRICBMCookie & Chr(GRInteger(0, 255))
    Next i
End Function

Public Function ReplaceTLV(lType As Long, strData As String, strValue As String) As String
    On Error GoTo BadCrap
    Dim i As Long
    Dim lngType As Long
    Dim lngLength As Long
    i = 1
    Do While i < Len(strData)
        lngType = GetWord(Mid(strData, i, 2)): i = i + 2
        lngLength = GetWord(Mid(strData, i, 2)): i = i + 2
        If lngType = lType Then
            ReplaceTLV = Mid(strData, 1, (i - 5)) & PutTLV(lType, strValue) & Mid(strData, i + lngLength)
            Exit Function
        End If
        i = i + lngLength
    Loop
    ReplaceTLV = strData
    Exit Function
BadCrap:
    ReplaceTLV = strData
End Function

Public Function GetTLV(lType As Long, strData As String) As String
    On Error GoTo BadCrap
    
    Dim i As Long
    Dim lngType As Long
    Dim lngLength As Long
    i = 1
    Do While i < Len(strData)
        lngType = GetWord(Mid(strData, i, 2)): i = i + 2
        lngLength = GetWord(Mid(strData, i, 2)): i = i + 2
        If lngType = lType Then
            GetTLV = Mid(strData, i, lngLength)
            Exit Function
        End If
        i = i + lngLength
    Loop
    GetTLV = "<none>"
    Exit Function
BadCrap:
    GetTLV = "<none>"
End Function

Public Function PutTLV(lType As Long, strData As String) As String
    PutTLV = Word(lType) & Word(Len(strData)) & strData
End Function
