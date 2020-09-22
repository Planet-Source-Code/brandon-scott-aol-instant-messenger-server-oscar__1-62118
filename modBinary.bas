Attribute VB_Name = "modBinary"
Option Explicit

Private Declare Function ntohl Lib "wsock32.dll" (ByVal a As Long) As Long
Private Declare Function ntohs Lib "wsock32.dll" (ByVal a As Long) As Integer

Public Function FileExist(strPath As String) As Boolean
    If Dir(strPath, vbNormal) = vbNullString Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function

Public Function LoadBinaryFile(strPath As String)
    Dim fileData As String
    Open strPath For Binary Access Read As #1
        fileData = Space$(LOF(1))
        Get #1, , fileData
    Close #1
    LoadBinaryFile = fileData
End Function

Public Function DecodeStr(ByVal str) As String
    On Error GoTo DecodingError:
    Dim pos As Long
    Dim cgiCharHex As String
    While (InStr(str, "%") <> 0)
        pos = InStr(str, "%")
        cgiCharHex = Mid$(str, pos + 1, 2)
        str = Replace$(str, "%" & cgiCharHex, Chr("&H" & cgiCharHex))
    Wend
DecodingError:
    DecodeStr = Replace$(str, "+", " ")
End Function

Public Function DisplayFormat(ByVal strData As String) As String
    Dim b As Integer
    For b = 127 To 255
        strData = Replace(strData, Chr(b), "[0x" & Hex(b) & "]")
    Next b
    For b = 0 To 31
        strData = Replace(strData, Chr(b), "[0x" & Hex(b) & "]")
    Next b
    DisplayFormat = strData
End Function

Public Function Word(ByVal lngVal As Long) As String
    Dim Lo As Long
    Dim Hi As Long
    On Error GoTo ErrWord

    Lo = Fix(lngVal / 256)
    Hi = lngVal Mod 256
    Word = Chr(Lo) & Chr(Hi)

    On Error GoTo 0
    Exit Function

ErrWord:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure Word of modBinary"

End Function

Public Function GetWord(ByVal strVal As String) As Long
    Dim Lo As Long
    Dim Hi As Long
    On Error GoTo ErrGetWord

    Lo = Asc(Mid(strVal, 1, 1))
    Hi = Asc(Mid(strVal, 2, 1))
    GetWord = (Lo * 256) + Hi

    On Error GoTo 0
    Exit Function

ErrGetWord:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure GetWord of modBinary"

End Function

Public Function DWord(ByVal lngVal As Double) As String
    Dim Lo As Single
    Dim Hi As Single
    On Error GoTo ErrDWord

    Lo = Fix(lngVal / 65536)
    Hi = Modulus(lngVal, 65536)
    DWord = Word(Lo) & Word(Hi)

    On Error GoTo 0
    Exit Function

ErrDWord:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure DWord of modBinary"

End Function
'
Public Function GetDWord(ByVal strVal As String) As Double
    Dim LoWord As Single
    Dim HiWord As Single
    On Error GoTo ErrGetDWord

    LoWord = GetWord(Mid(strVal, 1, 2))
    HiWord = GetWord(Mid(strVal, 3, 2))
    GetDWord = (LoWord * 65536) + HiWord

    On Error GoTo 0
    Exit Function

ErrGetDWord:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure GetDWord of modBinary"

End Function

Public Function SByte(strData As String) As String
    On Error GoTo ErrSByte

    SByte = Chr(Len(strData)) & strData

    On Error GoTo 0
    Exit Function

ErrSByte:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure SByte of modBinary"

End Function

Public Function SWord(strVal As String) As String
    On Error GoTo ErrSWord

    SWord = Word(Len(strVal)) & strVal

    On Error GoTo 0
    Exit Function

ErrSWord:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure SWord of modBinary"

End Function

Public Function SDWord(strVal As String) As String
    On Error GoTo ErrSDWord

    SDWord = DWord(Len(strVal)) & strVal

    On Error GoTo 0
    Exit Function

ErrSDWord:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure SDWord of modBinary"

End Function

Public Function GetSByte(strData As String)
    On Error GoTo ErrGetSByte

    GetSByte = Mid(strData, 2, Asc(Mid(strData, 1, 1)))

    On Error GoTo 0
    Exit Function

ErrGetSByte:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure GetSByte of modBinary"

End Function

Public Function GetSWord(strData As String)
    On Error GoTo ErrGetSWord

    GetSWord = Mid(strData, 3, GetWord(Mid(strData, 1, 2)))

    On Error GoTo 0
    Exit Function

ErrGetSWord:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure GetSWord of modBinary"

End Function

Public Function GetSDWord(strData As String)
    On Error GoTo ErrGetSDWord

    GetSDWord = Mid(strData, 5, GetDWord(Mid(strData, 1, 4)))

    On Error GoTo 0
    Exit Function

ErrGetSDWord:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure GetSDWord of modBinary"

End Function

Public Function Modulus(ByVal xVal As Double, ByVal yVal As Double) As Double
    On Error GoTo ErrModulus

    Modulus = xVal - yVal * Int(xVal / yVal)

    On Error GoTo 0
    Exit Function

ErrModulus:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure Modulus of modBinary"

End Function

Public Function Divide(ByVal N1 As Double, ByVal N2 As Double) As Double
    On Error GoTo ErrDivide

    Divide = Int(N1 / N2)

    On Error GoTo 0
    Exit Function

ErrDivide:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure Divide of modBinary"

End Function

Public Function HexToString(strIn As String) As String
    Dim i As Long
    For i = 1 To Len(strIn) Step 2
        HexToString = HexToString & Chr(Val("&H" & UCase(Mid(strIn, i, 2))))
    Next i
End Function

Public Function StringToHex(strIn As String) As String
    Dim i As Long
    Dim hexBuff As String
    For i = 1 To Len(strIn)
        hexBuff = Hex(Asc(Mid(strIn, i, 1)))
        StringToHex = StringToHex & IIf(Len(hexBuff) = 2, hexBuff, "0" & hexBuff)
    Next i
End Function

Public Function DecToString(strIn As String) As String
    Dim i As Long
    For i = 1 To Len(strIn) Step 3
        DecToString = DecToString & Chr(CLng(Mid(strIn, i, 3)))
    Next i
End Function

Public Function StringToDec(strIn As String) As String
    Dim i As Long
    Dim decBuff As String
    For i = 1 To Len(strIn)
        StringToDec = StringToDec & Format(Asc(Mid(strIn, i, 1)), "000")
    Next i
End Function

Public Function DecToHex(lngVal As Long) As String
    DecToHex = IIf(lngVal >= 16, Hex(lngVal), "0" & Hex(lngVal))
End Function

Public Function HexToDec(HexVal As String) As Long
    HexToDec = Val("&H" & HexVal)
End Function

Public Function ByteArrayToString(bData() As Byte) As String
    ByteArrayToString = bData
End Function

Public Function StringToByteArray(strData As String, bData() As Byte)
    bData() = strData
End Function

Public Function LEndianToggle(lNum As Long) As Long
    LEndianToggle = ntohl(lNum)
End Function

Public Function IEndianToggle(iNum As Integer) As Integer
    IEndianToggle = ntohs(iNum)
End Function

Public Function StringToUnicode(strData As String) As String
    StringToUnicode = StrConv(strData, vbUnicode)
End Function

Public Function UnicodeToString(strData As String) As String
    UnicodeToString = StrConv(strData, vbFromUnicode)
End Function

Public Function IPToLong(strIP As String) As Double
    Dim aIP As String
    Dim sIP() As String
    sIP() = Split(strIP, ".")
    aIP = Chr(sIP(0)) & Chr(sIP(1)) & Chr(sIP(2)) & Chr(sIP(3))
    IPToLong = GetDWord(aIP)
End Function

Public Function LongToIP(dblIP As Double) As String
    Dim a As String
    a = DWord(dblIP)
    LongToIP = Asc(Mid(a, 1, 1)) & "." & Asc(Mid(a, 2, 1)) & "." & Asc(Mid(a, 3, 1)) & "." & Asc(Mid(a, 4, 1))
End Function

Public Function ChrA(strData As String) As String
    Dim C1() As String
    Dim i As Integer
    On Error GoTo ErrChrA

    C1() = Split(strData, " ")
    For i = 0 To UBound(C1): ChrA = ChrA & Chr(C1(i)): Next i

    On Error GoTo 0
    Exit Function

ErrChrA:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure ChrA of modBinary"

End Function

Public Function ChrB(strData As String) As String
    On Error GoTo ErrChrB
    
    Dim C1() As String
    Dim i As Integer
    C1() = Split(strData, " ")
    For i = 0 To UBound(C1): ChrB = ChrB & Chr("&H" & C1(i)): Next i

    On Error GoTo 0
    Exit Function

ErrChrB:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure ChrB of modBinary"

End Function

Public Function StringToHexArray(strIncoming As String) As String
    Dim i As Integer, l As Integer
    Dim temp As String
    For i = 1 To Len(strIncoming)
        For l = 0 To 255
            If Mid(strIncoming, i, 1) = Chr(l) Then
                temp = temp & IIf(Len(Hex(l)) <> 2, "0" & Hex(l) & " ", Hex(l) & " ")
            End If
        Next l
    Next i
    StringToHexArray = temp
End Function

