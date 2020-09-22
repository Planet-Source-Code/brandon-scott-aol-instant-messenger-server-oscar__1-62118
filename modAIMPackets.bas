Attribute VB_Name = "modAIMPackets"
Option Explicit

Dim SNAC As String

Public Function AdminSendInfo(lRequested As Long, sData As String) As String

    SNAC = ChrB("00 07 00 03 00 00 00 00 00 00")
    AdminSendInfo = SNAC & Word(3) & Word(1) & PutTLV(lRequested, sData)
    
End Function

Public Function FlapVersion() As String

    FlapVersion = ChrB("00 00 00 01")
    
End Function

Public Function BuddyArrived(sScreenName As String, _
                             iWarningLevel As Integer, _
                             lNickFlags As Long, _
                             sShortCaps As String, _
                             sOscarCaps As String, _
                             dOnlineTime As Double, _
                             dSignonTOD As Double) As String

    SNAC = ChrB("00 03 00 0B 00 00 00 00 00 00")
    BuddyArrived = SNAC & SByte(sScreenName) & _
                          Word(iWarningLevel) & _
                          Word(4) & _
                          PutTLV(1, Word(lNickFlags)) & _
                          PutTLV(3, DWord(dSignonTOD)) & _
                          PutTLV(13, sOscarCaps) & _
                          PutTLV(15, DWord(dOnlineTime))

    
End Function

Public Function BuddyDeparted(sScreenName As String) As String

    SNAC = ChrB("00 03 00 0C 00 00 00 00 00 00")
    BuddyDeparted = SNAC & SByte(sScreenName) & ChrB("00 00 00 01 00 01 00 02 00 00")
    
End Function

Public Function BucpChallenge(sChallenge As String) As String

    SNAC = ChrB("00 17 00 07 00 00 00 00 00 00")
    BucpChallenge = SNAC & SWord(sChallenge)
    
End Function

Public Function BucpReply(sScreenName As String, _
                            Optional sCookie As String = "", _
                            Optional sEmailAddress As String = "", _
                            Optional sBosHost As String = "255.255.255.255", _
                            Optional lPort As Integer = 5191, _
                            Optional sPasswordChangeURL As String = "http://www.xeons.net", _
                            Optional bBadLogin As Boolean = False, _
                            Optional sURL As String = "http://www.xeons.net/", _
                            Optional lError As Long = 0) As String

    SNAC = ChrB("00 17 00 03 00 00 00 00 00 00")
    If bBadLogin = False Then
        BucpReply = SNAC & PutTLV(1, sScreenName) & PutTLV(5, sBosHost & ":" & CStr(lPort)) & PutTLV(6, sCookie) & PutTLV(17, sEmailAddress) & PutTLV(84, sPasswordChangeURL) '& PutTLV(85, Word(65535))
    Else
        BucpReply = SNAC & PutTLV(1, sScreenName) & PutTLV(4, sURL) & PutTLV(8, Word(lError))
    End If
End Function

Public Function ServiceHostOnline() As String

    SNAC = ChrB("00 01 00 03 00 00 00 00 00 00")
    ServiceHostOnline = SNAC & ChrB("00 01 00 02 00 03 00 04 00 06 00 07 00 08 00 09 00 0A 00 0B 00 13 00 15 00 22")

End Function

Public Function ServiceHostVersions() As String

    SNAC = ChrB("00 01 00 18 00 00 00 00 00 00")
    ServiceHostVersions = SNAC & ChrB("00 01 00 04 00 02 00 01") & _
                                 ChrB("00 03 00 01 00 04 00 01") & _
                                 ChrB("00 06 00 01 00 08 00 01") & _
                                 ChrB("00 09 00 01 00 0A 00 01") & _
                                 ChrB("00 0B 00 01 00 0C 00 01") & _
                                 ChrB("00 13 00 04 00 15 00 01") & _
                                 ChrB("00 22 00 01 00 07 00 01")
End Function

Public Function ChatNavExchangeInfo() As String
    SNAC = ChrB("00 0D 00 09 00 00 00 00 00 00")
    ChatNavExchangeInfo = SNAC
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 02 00 02 00 01 03 00 03 00 3C 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("01 00 0A 00 03 00 01 14 00 04 00 02 20 00 00 C9")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 02 00 44 00 CA 00 04 00 00 00 00 00 D0 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D1 00 02 07 D0 00 D2 00 02 00 2F 00 D4 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D5 00 01 01 00 DA 00 02 00 66 00 03 00 3C 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("02 00 0A 00 03 00 01 0F 00 04 00 02 1E 00 00 C9")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 02 00 40 00 CA 00 04 00 00 00 00 00 D0 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D1 00 02 07 D0 00 D2 00 02 00 26 00 D4 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D5 00 01 01 00 DA 00 02 07 D0 00 03 00 3C 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("04 00 0A 00 03 00 01 0F 00 04 00 02 1E 00 00 C9")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 02 00 40 00 CA 00 04 00 00 00 00 00 D0 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D1 00 02 07 D0 00 D2 00 02 00 16 00 D4 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D5 00 01 01 00 DA 00 02 00 E8 00 03 00 3C 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("05 00 0A 00 03 00 01 0F 00 04 00 02 1E 00 00 C9")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 02 00 44 00 CA 00 04 00 00 00 00 00 D0 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D1 00 02 02 00 00 D2 00 02 00 26 00 D4 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D5 00 01 01 00 DA 00 02 02 00 00 03 00 3C 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("06 00 0A 00 03 00 01 0F 00 04 00 02 1E 00 00 C9")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 02 00 44 00 CA 00 04 00 00 00 00 00 D0 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D1 00 02 02 00 00 D2 00 02 00 26 00 D4 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D5 00 01 01 00 DA 00 02 02 00 00 03 00 3C 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("07 00 0A 00 03 00 01 0F 00 04 00 02 40 00 00 C9")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 02 00 44 00 CA 00 04 00 00 00 00 00 D0 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D1 00 02 07 D0 00 D2 00 02 00 19 00 D4 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D5 00 01 01 00 DA 00 02 00 E8 00 03 00 3C 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("08 00 0A 00 03 00 01 0F 00 04 00 02 1E 00 00 C9")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 02 00 40 00 CA 00 04 00 00 00 00 00 D0 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D1 00 02 07 D0 00 D2 00 02 00 31 00 D4 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D5 00 01 01 00 DA 00 02 00 E8 00 03 00 3C 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("09 00 0A 00 03 00 01 0F 00 04 00 02 1E 00 00 C9")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 02 00 40 00 CA 00 04 00 00 00 00 00 D0 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D1 00 02 07 D0 00 D2 00 02 00 31 00 D4 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D5 00 01 01 00 DA 00 02 00 E8 00 03 00 3C 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("0A 00 0A 00 03 00 01 0F 00 04 00 02 1E 00 00 C9")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 02 00 40 00 CA 00 04 00 00 00 00 00 D0 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D1 00 02 07 D0 00 D2 00 02 00 31 00 D4 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D5 00 01 01 00 DA 00 02 00 E8 00 03 00 3C 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("0B 00 0A 00 03 00 01 0F 00 04 00 02 1E 00 00 C9")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 02 00 40 00 CA 00 04 00 00 00 00 00 D0 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D1 00 02 07 D0 00 D2 00 02 00 31 00 D4 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D5 00 01 01 00 DA 00 02 00 E8 00 03 00 3C 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("0C 00 0A 00 03 00 01 0F 00 04 00 02 1E 00 00 C9")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 02 00 40 00 CA 00 04 00 00 00 00 00 D0 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D1 00 02 07 D0 00 D2 00 02 00 31 00 D4 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D5 00 01 01 00 DA 00 02 00 E8 00 03 00 3C 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("0D 00 0A 00 03 00 01 0F 00 04 00 02 1E 00 00 C9")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 02 00 40 00 CA 00 04 00 00 00 00 00 D0 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D1 00 02 07 D0 00 D2 00 02 00 31 00 D4 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D5 00 01 01 00 DA 00 02 00 E8 00 03 00 3C 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("0E 00 0A 00 03 00 01 0F 00 04 00 02 1E 00 00 C9")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 02 00 40 00 CA 00 04 00 00 00 00 00 D0 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D1 00 02 07 D0 00 D2 00 02 00 31 00 D4 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D5 00 01 01 00 DA 00 02 00 E8 00 03 00 3C 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("0F 00 0A 00 03 00 01 0F 00 04 00 02 1E 00 00 C9")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 02 00 40 00 CA 00 04 00 00 00 00 00 D0 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D1 00 02 07 D0 00 D2 00 02 00 31 00 D4 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D5 00 01 01 00 DA 00 02 00 E8 00 03 00 3C 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("10 00 0A 00 03 00 01 0F 00 04 00 02 1E 00 00 C9")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 02 00 40 00 CA 00 04 00 00 00 00 00 D0 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D1 00 02 07 D0 00 D2 00 02 00 31 00 D4 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D5 00 01 01 00 DA 00 02 00 E8 00 03 00 3C 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("14 00 0A 00 03 00 01 0F 00 04 00 02 40 00 00 C9")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 02 00 44 00 CA 00 04 00 00 00 00 00 D0 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D1 00 02 07 D0 00 D2 00 02 00 19 00 D4 00 00")
    ChatNavExchangeInfo = ChatNavExchangeInfo & ChrB("00 D5 00 01 01 00 DA 00 02 00 E8")
End Function

Public Function ServiceMotd() As String
 
    SNAC = ChrB("00 01 00 13 00 00 00 00 00 00")
    ServiceMotd = SNAC & ChrB("00 05 00 02 00 02 00 1E 00 03 00 02 04 B0")
    
End Function

Public Function ServiceRateParams() As String
    SNAC = ChrB("00 01 00 07 00 00 00 00 00 00")
    ServiceRateParams = SNAC
    ServiceRateParams = ServiceRateParams & ChrB("00 05 00 01 00 00 00 50 00 00")
    ServiceRateParams = ServiceRateParams & ChrB("09 C4 00 00 07 D0 00 00 05 DC 00 00 03 20 00 00")
    ServiceRateParams = ServiceRateParams & ChrB("16 DC 00 00 17 70 00 00 00 00 00 00 02 00 00 00")
    ServiceRateParams = ServiceRateParams & ChrB("50 00 00 0B B8 00 00 07 D0 00 00 05 DC 00 00 03")
    ServiceRateParams = ServiceRateParams & ChrB("E8 00 00 17 70 00 00 17 70 00 00 00 7B 00")

    ServiceRateParams = ServiceRateParams & Word(3) & DWord(1000) & DWord(30) & DWord(20) & DWord(10) & DWord(0) & DWord(65535) & DWord(65535) & DWord(0) & Chr(0)
    
    ServiceRateParams = ServiceRateParams & ChrB("00 04 00 00 00 14 00 00 15 7C 00 00 14 B4 00")
    ServiceRateParams = ServiceRateParams & ChrB("00 10 68 00 00 0B B8 00 00 17 70 00 00 1F 40 00")
    ServiceRateParams = ServiceRateParams & ChrB("00 00 7B 00 00 05 00 00 00 0A 00 00 15 7C 00 00")
    ServiceRateParams = ServiceRateParams & ChrB("14 B4 00 00 10 68 00 00 0B B8 00 00 17 70 00 00")
    ServiceRateParams = ServiceRateParams & ChrB("1F 40 00 00 00 7B 00 00 01 00 A6 00 01 00 01 00")
    ServiceRateParams = ServiceRateParams & ChrB("01 00 02 00 01 00 03 00 01 00 04 00 01 00 05 00")
    ServiceRateParams = ServiceRateParams & ChrB("01 00 06 00 01 00 07 00 01 00 08 00 01 00 09 00")
    ServiceRateParams = ServiceRateParams & ChrB("01 00 0A 00 01 00 0B 00 01 00 0C 00 01 00 0D 00")
    ServiceRateParams = ServiceRateParams & ChrB("01 00 0E 00 01 00 0F 00 01 00 10 00 01 00 11 00")
    ServiceRateParams = ServiceRateParams & ChrB("01 00 12 00 01 00 13 00 01 00 14 00 01 00 15 00")
    ServiceRateParams = ServiceRateParams & ChrB("01 00 16 00 01 00 17 00 01 00 18 00 01 00 19 00")
    ServiceRateParams = ServiceRateParams & ChrB("01 00 1A 00 01 00 1B 00 01 00 1C 00 01 00 1D 00")
    ServiceRateParams = ServiceRateParams & ChrB("01 00 1E 00 01 00 1F 00 01 00 20 00 01 00 21 00")
    ServiceRateParams = ServiceRateParams & ChrB("01 00 22 00 01 00 23 00 01 00 24 00 01 00 25 00")
    ServiceRateParams = ServiceRateParams & ChrB("02 00 01 00 02 00 02 00 02 00 03 00 02 00 04 00")
    ServiceRateParams = ServiceRateParams & ChrB("02 00 06 00 02 00 07 00 02 00 08 00 02 00 0A 00")
    ServiceRateParams = ServiceRateParams & ChrB("02 00 0C 00 02 00 0D 00 02 00 0E 00 02 00 0F 00")
    ServiceRateParams = ServiceRateParams & ChrB("02 00 10 00 02 00 11 00 02 00 12 00 02 00 13 00")
    ServiceRateParams = ServiceRateParams & ChrB("02 00 14 00 02 00 15 00 03 00 01 00 03 00 02 00")
    ServiceRateParams = ServiceRateParams & ChrB("03 00 03 00 03 00 06 00 03 00 07 00 03 00 08 00")
    ServiceRateParams = ServiceRateParams & ChrB("03 00 09 00 03 00 0A 00 03 00 0B 00 03 00 0C 00")
    ServiceRateParams = ServiceRateParams & ChrB("03 00 0D 00 03 00 0E 00 04 00 01 00 04 00 02 00")
    ServiceRateParams = ServiceRateParams & ChrB("04 00 03 00 04 00 04 00 04 00 05 00 04 00 07 00")
    ServiceRateParams = ServiceRateParams & ChrB("04 00 08 00 04 00 09 00 04 00 0A 00 04 00 0B 00")
    ServiceRateParams = ServiceRateParams & ChrB("04 00 0C 00 04 00 0D 00 04 00 0E 00 04 00 0F 00")
    ServiceRateParams = ServiceRateParams & ChrB("04 00 10 00 04 00 11 00 04 00 12 00 04 00 13 00")
    ServiceRateParams = ServiceRateParams & ChrB("04 00 14 00 04 00 15 00 06 00 01 00 06 00 02 00")
    ServiceRateParams = ServiceRateParams & ChrB("06 00 03 00 08 00 01 00 08 00 02 00 09 00 01 00")
    ServiceRateParams = ServiceRateParams & ChrB("09 00 02 00 09 00 03 00 09 00 04 00 09 00 09 00")
    ServiceRateParams = ServiceRateParams & ChrB("09 00 0A 00 09 00 0B 00 0A 00 01 00 0A 00 02 00")
    ServiceRateParams = ServiceRateParams & ChrB("0A 00 03 00 0B 00 01 00 0B 00 02 00 0B 00 03 00")
    ServiceRateParams = ServiceRateParams & ChrB("0B 00 04 00 0C 00 01 00 0C 00 02 00 0C 00 03 00")
    ServiceRateParams = ServiceRateParams & ChrB("13 00 01 00 13 00 02 00 13 00 03 00 13 00 04 00")
    ServiceRateParams = ServiceRateParams & ChrB("13 00 05 00 13 00 06 00 13 00 07 00 13 00 08 00")
    ServiceRateParams = ServiceRateParams & ChrB("13 00 09 00 13 00 0A 00 13 00 0B 00 13 00 0C 00")
    ServiceRateParams = ServiceRateParams & ChrB("13 00 0D 00 13 00 0E 00 13 00 0F 00 13 00 10 00")
    ServiceRateParams = ServiceRateParams & ChrB("13 00 11 00 13 00 12 00 13 00 13 00 13 00 14 00")
    ServiceRateParams = ServiceRateParams & ChrB("13 00 15 00 13 00 16 00 13 00 17 00 13 00 18 00")
    ServiceRateParams = ServiceRateParams & ChrB("13 00 19 00 13 00 1A 00 13 00 1B 00 13 00 1C 00")
    ServiceRateParams = ServiceRateParams & ChrB("13 00 1D 00 13 00 1E 00 13 00 1F 00 13 00 20 00")
    ServiceRateParams = ServiceRateParams & ChrB("13 00 21 00 13 00 22 00 13 00 23 00 13 00 24 00")
    ServiceRateParams = ServiceRateParams & ChrB("13 00 25 00 13 00 26 00 13 00 27 00 13 00 28 00")
    ServiceRateParams = ServiceRateParams & ChrB("13 00 29 00 13 00 2A 00 13 00 2B 00 13 00 2C 00")
    ServiceRateParams = ServiceRateParams & ChrB("13 00 2D 00 13 00 2E 00 13 00 2F 00 13 00 30 00")
    ServiceRateParams = ServiceRateParams & ChrB("13 00 31 00 13 00 32 00 13 00 33 00 13 00 34 00")
    ServiceRateParams = ServiceRateParams & ChrB("13 00 35 00 13 00 36 00 15 00 01 00 15 00 02 00")
    ServiceRateParams = ServiceRateParams & ChrB("15 00 03 00 02 00 06 00 03 00 04 00 03 00 05 00")
    ServiceRateParams = ServiceRateParams & ChrB("09 00 05 00 09 00 06 00 09 00 07 00 09 00 08 00")
    ServiceRateParams = ServiceRateParams & ChrB("03 00 02 00 02 00 05 00 04 00 06 00 04 00 02 00")
    ServiceRateParams = ServiceRateParams & ChrB("02 00 09 00 02 00 0B 00 05 00 00")
    
End Function

Public Function ServiceNickInfoReply(strName As String) As String

    SNAC = ChrB("00 01 00 0F 00 00 00 00 00 00")
    ServiceNickInfoReply = SNAC & SByte(strName) & ChrB("00 00 00 06 00 01 00 02 00 90 00 0F 00 04 00 00 00 00 00 03 00 04 41 E9 B4 BB 00 0A 00 04 44 E3 A7 35 00 1E 00 04 00 00 00 00 00 05 00 04 38 C4 76 E8")

End Function

Public Function FeedbagRightsReply() As String

    SNAC = ChrB("00 13 00 03 00 00 00 00 00 00")
    FeedbagRightsReply = SNAC & ChrB("00 04 00 34 01 90 00 3D 00 C8 00 C8 00 01 00 01 00 96 00 0C 00 0C 00 00 00 32 00 32 00 00 00 00 00 00 00 00 00 00 00 01 00 00 00 00 00 0F 00 01 00 28 00 01 00 0A 00 C8 00 02 00 02 00 FE 00 03 00 02 01 FC 00 05 00 02 00 64 00 06 00 02 00 61 00 07 00 02 00 C8 00 08 00 02 00 0A 00 09 00 04 00 06 97 80 00 0A 00 04 00 00 00 0E")

End Function

Public Function FeedbagError() As String

    SNAC = ChrB("00 13 00 01 00 00 00 00 00 00")
    FeedbagError = SNAC & ChrB("00 01 00 02 00 01")
    
End Function

Public Function FeedbagBuddylist(sFeedbagBuffer As String) As String

    SNAC = ChrB("00 13 00 06 00 00 00 00 00 00")
    FeedbagBuddylist = SNAC & sFeedbagBuffer
    
End Function

Public Function FeedbagStatusReply(sFeedbagBuffer As String) As String

    SNAC = ChrB("00 13 00 0E 80 00 00 00 00 00") & ChrB("00 06 00 01 00 02 00 03")
    FeedbagStatusReply = SNAC & sFeedbagBuffer
    
End Function

Public Function FeedbagReplyNotModified(sData As String) As String

    SNAC = ChrB("00 13 00 0F 00 00 00 00 00 00")
    FeedbagReplyNotModified = SNAC & sData
    
End Function

Public Function LocationRightsReply() As String

    SNAC = ChrB("00 02 00 03 00 00 00 00 00 00")
    LocationRightsReply = SNAC & ChrB("00 01 00 02 04 00 00 02 00 02 00 12 00 05 00 02 00 80 00 03 00 02 00 0A 00 04 00 02 10 00")

End Function

Public Function LocationUserInfoReply(lInfoType As Long, _
                                      sWho As String, _
                                      lUserClass As Long, _
                                      dSignonTOD As Double, _
                                      dOnlineTime As Double, _
                                      dIdleTime As Double, _
                                      iWarningLevel As Integer, _
                                      sCapabilities As String, _
                                      sProfileEncoding As String, _
                                      sProfile As String, _
                                      sAwayEncoding As String, _
                                      sAwayMessage As String) As String
                                      
    Dim sBuffer As String
    SNAC = ChrB("00 02 00 06 00 00 00 00 00 00")
    sBuffer = SNAC & SByte(sWho) & Word(iWarningLevel) & Word(3) & PutTLV(1, Word(lUserClass)) & PutTLV(&HF, DWord(dOnlineTime)) & PutTLV(&H3, DWord(dSignonTOD))
    Select Case lInfoType
        Case 1
            If Len(sProfile) > 0 Then sBuffer = sBuffer & PutTLV(1, sProfileEncoding) & PutTLV(2, sProfile)
        Case 3
            If Len(sAwayMessage) > 0 Then sBuffer = sBuffer & PutTLV(3, sAwayEncoding) & PutTLV(4, sAwayMessage)
        Case 4
            If Len(sCapabilities) > 0 Then sBuffer = sBuffer & PutTLV(5, sCapabilities)
        Case 5
            If Len(sProfile) > 0 Then sBuffer = sBuffer & PutTLV(1, sProfileEncoding) & PutTLV(2, sProfile)
            If Len(sAwayMessage) > 0 Then sBuffer = sBuffer & PutTLV(3, sAwayEncoding) & PutTLV(4, sAwayMessage)
            If Len(sCapabilities) > 0 Then sBuffer = sBuffer & PutTLV(5, sCapabilities)
    End Select
    
    LocationUserInfoReply = sBuffer
End Function

'There are two versions of the request
Public Function LocationUserInfoReply2(dFlags As Double, _
                                       sWho As String, _
                                       lUserClass As Long, _
                                       dSignonTOD As Double, _
                                       dOnlineTime As Double, _
                                       dIdleTime As Double, _
                                       iWarningLevel As Integer, _
                                       sCapabilities As String, _
                                       sProfileEncoding As String, _
                                       sProfile As String, _
                                       sAwayEncoding As String, _
                                       sAwayMessage As String) As String
                                      
    Dim sBuffer As String
    SNAC = ChrB("00 02 00 06 00 00 00 00 00 00")
    sBuffer = SNAC & SByte(sWho) & Word(iWarningLevel) & Word(3) & PutTLV(1, Word(lUserClass)) & PutTLV(&HF, DWord(dOnlineTime)) & PutTLV(&H3, DWord(dSignonTOD))
    
    If (dFlags And &H1) = &H1 Then
        If Len(sProfile) > 0 Then sBuffer = sBuffer & PutTLV(1, sProfileEncoding) & PutTLV(2, sProfile)
    End If
    If (dFlags And &H2) = &H2 Then
        If Len(sAwayMessage) > 0 Then sBuffer = sBuffer & PutTLV(3, sAwayEncoding) & PutTLV(4, sAwayMessage)
    End If
    If (dFlags And &H4) = &H4 Then
        If Len(sCapabilities) > 0 Then sBuffer = sBuffer & PutTLV(5, sCapabilities)
    End If
    'If (dFlags And &H8) = &H8 Then
    '    If Len(sProfile) > 0 Then sBuffer = sBuffer & PutTLV(1, sProfileEncoding) & PutTLV(2, sProfile)
    'End If
    
    LocationUserInfoReply2 = sBuffer
End Function

Public Function LocationError(lErrorCode As Long) As String

    SNAC = ChrB("00 02 00 01 00 00 00 00 00 00")
    LocationError = SNAC & Word(lErrorCode)

End Function

Public Function BuddyRightsReply() As String

    SNAC = ChrB("00 03 00 03 00 00 00 00 00 00")
    BuddyRightsReply = SNAC & ChrB("00 02 00 02 07 D0 00 01 00 02 00 DC 00 04 00 02 00 20")
    
End Function

Public Function BosRightsReply() As String

    SNAC = ChrB("00 09 00 03 00 00 00 00 00 00")
    BosRightsReply = SNAC & ChrB("00 02 00 02 00 DC 00 01 00 02 00 DC")
    
End Function

Public Function IcbmParamReply() As String

    SNAC = ChrB("00 04 00 05 00 00 00 00 00 00")
    IcbmParamReply = SNAC & ChrB("00 02 00 00 00 03 02 00 03 84 03 E7 00 00 03 E8")
    
End Function


Public Function IcbmToClient(ByVal strCookie As String, ByVal strScreenName As String, ByVal strMessage As String) As String

    SNAC = ChrB("00 04 00 07 00 00 00 00 00 00")
    IcbmToClient = SNAC & strCookie & ChrB("00 01") & SByte(strScreenName) & ChrB("00 00 00 04 00 01 00 02 00 10 00 06 00 04 00 00 01 00 00 0F 00 04 00 00 57 0B 00 03 00 04 40 E6 DA B8") & PutTLV(2, ChrB("05 01 00 03 01 01 02 01 01") & SWord(ChrB("00 00 00 00") & strMessage)) & ChrB("00 0B 00 00")

End Function

Public Function IcbmError(lErrorCode As Long) As String

    SNAC = ChrB("00 04 00 01 00 00 00 00 00 00")
    IcbmError = SNAC & Word(lErrorCode)
    
End Function

Public Function IcbmHostAck(ByVal sCookie As String, sScreenName As String) As String

    SNAC = ChrB("00 04 00 0C 00 00 00 00 00 00")
    IcbmHostAck = SNAC & sCookie & ChrB("00 01") & SByte(sScreenName)
    
End Function

Public Function ServiceBartReply() As String

    SNAC = ChrB("00 01 00 21 00 00 00 00 00 00")
    ServiceBartReply = SNAC & ChrB("04 00 00 05 2B 00 00 18 51 00 81 00 05 2B 00 00 11 F9 00 00 00 05 02 01 D2 04 72 00 01 41 10 F6 10 DC 05 3C 62 BF F2 87 74 23 BA DE 3C 58 82 00 03 00 07 05 69 6D 73 65 6E 64 00 60 00 07 05 69 6D 73 65 6E 64 00 83 00 07 05 69 6D 73 65 6E 64")

End Function

Public Function ServiceReponse(sHost As String, lPort As Long, lGroup As Long, sCookie As String) As String

    SNAC = ChrB("00 01 00 05 00 00 00 00 00 00")
    ServiceReponse = SNAC & PutTLV(13, Word(lGroup)) & PutTLV(5, sHost & ":" & lPort) & PutTLV(6, sCookie)
    
End Function



