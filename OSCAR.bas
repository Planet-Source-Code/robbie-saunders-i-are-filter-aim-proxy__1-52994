Attribute VB_Name = "OSCAR"
Dim A1, A2, A3, A4, A5, A6, A7, A8, A9

'login authorization
Public Function AuthLogin(strUser As String, strPassword As String, Optional intService As Integer = 1) As String

    A1 = ""
    AuthLogin = A1 & ChrA("0 0 0 1") & TLV(1, strUser) & TLV(2, EncryptPW(strPassword)) & _
        TLV(3, "TestBuddy AOL Instant Messenger (SM)") & _
        TLV(22, Word(intService)) & TLV(23, ChrA("0 1")) & TLV(24, ChrA("0 0")) & _
        TLV(25, ChrA("0 0")) & TLV(26, ChrA("0 1")) & TLV(14, "us") & TLV(15, "en") ' & TLV(9, ChrA("0 9"))

End Function

'set a profile/capabilities/with security
Public Function changeIdle(strIdleTime As String) As String

    A1 = ChrA("0 1 0 17 0 0 0 0 0 17")
    changeIdle = A1 & strIdleTime
        
End Function

'challenge request
Public Function ChallengeRequest(strUser As String, strSecureID As String) As String

    A1 = ChrA("0 23 0 6 0 0 0 0 0 0")
    ChallengeRequest = A1 & TLV(1, strUser) ' & TLV(75, strSecureID)

End Function

'bucp query
Public Function BUCPQuery(strUser As String, strPassword As String, strChallenge As String, intMultiConn As Integer) As String

    A1 = ChrA("0 23 0 2 0 0 0 0 0 0")
    BUCPQuery = A1 & TLV(1, strUser) & TLV(37, AIMEncryptPw(strChallenge, strPassword)) & _
        TLV(76, vbNullString) & _
        TLV(3, "AOL Instant Messenger, version 5.5.3145/WIN32") & _
        TLV(22, ChrH("ff 15")) & _
        TLV(23, Word(1)) & TLV(24, Word(75)) & TLV(25, Word(0)) & _
        TLV(26, Word(563)) & TLV(15, "en") & TLV(14, "us") & TLV(74, Chr(255))

End Function

Public Function auth_sendlogin(strUser As String, strPassword As String, strChallenge As String)
    A1 = ChrH("00 17 00 02 00 00 00 00 00 00")

    auth_sendlogin = A1 & _
                    TLV(1, strUser) & _
                     TLV(37, AIMEncryptPw(strChallenge, strPassword)) & _
                     TLV(76, vbNullString) & _
                     TLV(3, "AOL Instant Messenger, version 5.2.3277/WIN32") & _
                    TLV(22, "asdf") & _
                     TLV(23, Word(1)) & TLV(24, Word(75)) & _
                     TLV(25, Word(0)) & TLV(26, Word(563)) & _
                     TLV(14, "us") & TLV(15, "en")
End Function

Public Function removeMe(strUserName As String)
    
    A1 = ChrA("0 19 0 22 0 0 0 0 0 22")
    removeMe = A1 & OBL(strUserName)

End Function

'login authorization step 2
Public Function authLogin2(strCookie As String) As String

    A1 = ""
    authLogin2 = A1 & ChrA("0 0 0 1") & TLV(6, strCookie)

End Function

'required for login
Public Function addICBMParam() As String

    A1 = ChrA("0 4 0 2 0 0 0 0 0 2")
    addICBMParam = A1 & ChrA("0 0 0 0 0 11 31 64 3 231 3 231 0 0 0 0")
    
End Function

'some mysterious packet
Public Function loginPacket1() As String
    
    ' 0 1  - permit some
    ' 0 19 - permit some
    ' 0 2  - permit all
    ' 0 3  - permit all
    ' 0 4  - permit all
    ' 0 6  - permit all
    ' 0 8  - permit all
    ' 0 9  - permit all
    ' 0 10 - permit all
    ' 0 11 - permit all

    A1 = ChrA("0 1 0 23 0 0 0 0 0 23")
    loginPacket1 = A1 & ChrA("0 1 0 3 0 19 0 3 0 2 0 1 0 3 0 1 0 4 0 1 0 6 0 1 0 8 0 1 0 9 0 1 0 10 0 1 0 11 0 1")
    
End Function


'some mysterious packet
Public Function formatPacket1() As String

    A1 = ChrA("0 1 0 23 0 0 0 0 0 23")
    formatPacket1 = A1 & ChrA("0 1 0 3 0 7 0 1")
    
End Function

'some request packet
Public Function requestPacket1() As String

    A1 = ChrA("0 19 0 2 0 0 0 0 0 2")
    requestPacket1 = A1
    
End Function

'some request packet
Public Function requestPacket2() As String

    A1 = ChrA("0 19 0 5 0 0 0 32 0 5")
    requestPacket2 = A1 & ChrA("62 37 14 13 1 101")
    
End Function

'some request packet
Public Function requestPacket3() As String

    A1 = ChrA("0 2 0 2 0 0 0 0 0 2")
    requestPacket3 = A1
    
End Function

'some request packet
Public Function requestPacket4() As String

    A1 = ChrA("0 4 0 4 0 0 0 0 0 4")
    requestPacket4 = A1
    
End Function

'some request packet
Public Function requestPacket5() As String

    A1 = ChrA("0 19 0 7 0 0 0 0 0 7")
    requestPacket5 = A1
    
End Function

'request rate
Public Function requestRate() As String

    A1 = ChrA("0 1 0 6 0 0 0 0 0 6")
    requestRate = A1
    
End Function

'request personal info
Public Function requestPInfo() As String
    
    A1 = ChrA("0 1 0 14 0 0 0 0 0 14")
    requestPInfo = A1

End Function

'some login packet
Public Function someThing() As String

    A1 = ChrA("0 9 0 2 0 0 0 0 0 2")
    someThing = A1
    
End Function

'request buddy list?
Public Function requestList() As String

    A1 = ChrA("0 3 0 2 0 0 0 0 0 2")
    requestList = A1

End Function

'acknowledge rate
Public Function rateAck() As String

    A1 = ChrA("0 1 0 8 0 0 0 0 0 8")
    rateAck = A1 & ChrA("0 1 0 2 0 3 0 4 0 5")
    
End Function

'request privacy
Public Function requestPrivacy() As String

    A1 = ChrA("0 1 0 20 0 0 0 0 0 0")
    requestPrivacy = A1 & ChrA("0 0 0 3")
    
End Function

'stuff for the search? i dunno
Public Function watcherPacket1() As String

    A1 = ChrA("0 2 0 9 0 0 0 1 0 9")
    watcherPacket1 = A1
    
End Function

'stuff for the search? i dunno
Public Function watcherPacket2() As String

    A1 = ChrA("0 2 0 15 0 0 0 2 0 15")
    watcherPacket2 = A1
    
End Function

'stuff for the search? i dunno
Public Function watcherPacket3(strName As String) As String

    A1 = ChrA("0 2 0 11 0 0 0 3 0 11")
    watcherPacket3 = A1 & Chr(Len(strName)) & strName
    
End Function

'set a profile/capabilities/with security
Public Function changeProfile(strProfileText As String) As String

    A1 = ChrA("0 2 0 4 0 0 0 0 0 0")
    changeProfile = A1 & TLV(1, "oscSockv1.2") & _
        TLV(2, strProfileText) & TLV(5, FULLCAP) & _
        TLV(6, securityBlock)
        
End Function

Public Function securityBlock()
    
    securityBlock = _
        TLV(4, ChrA("0 1")) & _
        ChrH("00 01 04 68 30 82 04 64 30 82 03 CD A0 03 02 01") & _
        ChrH("02 02 10 12 84 ED 80 40 09 F3 99 73 4F AD F6 FB") & _
        ChrH("E1 A7 91 30 0D 06 09 2A 86 48 86 F7 0D 01 01 04") & _
        ChrH("05 00 30 81 CC") & _
        ChrH("31 17 30 15 06 03 55 04 0A 13") & _
        OBL("VeriSign, Inc.") & _
        ChrH("31 1F 30 1D 06 03 55 04 0B 13") & _
        OBL("VeriSign Trust Network") & _
        ChrH("31 46 30 44 06 03 55 04 0B 13") & _
        OBL("www.verisign.com/repository/RPA Incorp. By Ref.,LIAB.LTD(c)98") & _
        ChrH("31 48 30 46 06 03 55 04 03 13") & _
        OBL("VeriSign Class 1 CA Individual Subscriber-Persona Not Validated") & _
        ChrH("30 1E 17") & _
        OBL("030701000000Z") & _
        ChrH("17") & _
        OBL("030830235959Z") & _
        ChrH("30 82 01 08") & _
        ChrH("31 17 30 15 06 03 55 04 0A 13") & _
        OBL("VeriSign, Inc.") & _
        ChrH("31 1F 30 1D 06 03 55 04 0B 13") & _
        OBL("VeriSign Trust Network") & _
        ChrH("31 46 30 44 06 03 55 04 0B 13") & _
        OBL("www.verisign.com/repository/RPA Incorp. By Ref.,LIAB.LTD(c)98")

    securityBlock = securityBlock & _
        ChrH("31 1E 30 1C 06 03 55 04 0B 13") & _
        OBL("Persona Not Validated") & _
        ChrH("31 27 30 25 06 03 55 04 0B 13") & _
        OBL("Digital ID Class 1 - Microsoft") & _
        ChrH("31 17 30 15 06 03 55 04 03 14") & _
        OBL("oscSock version 1.2") & _
        ChrH("31 22 30 20 06 09 2A 86 48 86") & _
        ChrH("F7 0D 01 09 01 16") & _
        OBL("oscSock version 1.2") & _
        ChrH("30 81 9F") & _
        ChrH("30 0D 06 09 2A 86 48 86 F7 0D 01 01 01 05 00 03") & _
        ChrH("81 8D 00 30 81 89 02 81 81 00 92 C6 89 82 61 97") & _
        ChrH("9D E2 FB 02 A0 C1 70 F7 43 BC E6 68 8B 03 20 EB") & _
        ChrH("27 89 7F E5 4A 1A A2 9A CA F2 D5 12 3B 38 F3 5C") & _
        ChrH("01 74 F5 D7 98 47 20 CC 71 7F BF E1 30 49 69 9D") & _
        ChrH("24 0A B7 A5 36 DA D8 06 63 99 9B D8 7A 89 51 BC") & _
        ChrH("71 85 D3 F0 CB C2 46 12 41 0E 79 02 0A 02 70 64") & _
        ChrH("C1 9C B5 04 13 C4 49 DF 04 0A 58 22 F1 6A 5E C7") & _
        ChrH("7E 12 DB 6D 7F 9B C4 4B D8 7A 10 1E F9 A0 37 A4") & _
        ChrH("09 BD B3 01 AB 84 67 CE 5D 29 02 03 01 00 01 A3") & _
        ChrH("82 01 06 30 82 01 02 30 09 06 03 55 1D 13 04 02") & _
        ChrH("30 00 30 81 AC 06 03 55 1D 20 04 81 A4 30 81 A1") & _
        ChrH("30 81 9E 06 0B 60 86 48 01 86 F8 45 01 07 01 01") & _
        ChrH("30 81 8E 30 28 06 08 2B 06 01 05 05 07 02 01 16")
        
    securityBlock = securityBlock & _
        OBL("https://www.verisign.com/CPS") & _
        ChrH("30 62 06") & _
        ChrH("08 2B 06 01 05 05 07 02 02 30 56 30 15 16") & _
        OBL("VeriSign, Inc.") & _
        ChrH("30 03 02 01 01 1A") & _
        OBL("VeriSign's CPS incorp. by reference liab. ltd. (c)97 VeriSign") & _
        ChrH("30 11 06 09 60 86 48 01 86 F8 42 01 01 04 04") & _
        ChrH("03 02 07 80 30 33 06 03 55 1D 1F 04 2C 30 2A 30") & _
        ChrH("28 A0 26 A0 24 86") & _
        OBL("http://crl.verisign.com/class1.crl") & _
        ChrH("30 0D 06 09 2A 86 48") & _
        ChrH("86 F7 0D 01 01 04 05 00 03 81 81 00 6F 6C F7 23") & _
        ChrH("80 0E 50 6B F9 21 98 69 D2 61 05 45 42 CA 97 E4") & _
        ChrH("11 D6 EC 23 8E 3E 06 82 5B D2 54 13 81 D7 9C 82") & _
        ChrH("52 8B 52 84 CD B0 68 38 7D 37 59 F5 AE DA 72 92") & _
        ChrH("E8 68 10 16 39 25 B2 DF BA F0 B3 4A 27 19 1D 4F") & _
        ChrH("15 B5 24 EB AF 42 D8 19 74 8C 4C 1C E6 BC 68 E7") & _
        ChrH("CA 38 49 EC 70 91 57 F7 20 A9 87 A1 47 AB 27 6E") & _
        ChrH("96 C3 7E 2A 6F 15 73 2F CC 86 B8 E2 C2 F1 13 AB") & _
        ChrH("14 31 6B 5F E2 A8 B6 B4 61 79 F8 90 00 05 00 14") & _
        ChrH("04 02 01 10 1B 43 A9 9A 9B 98 43 74 F2 04 39 13") & _
        ChrH("24 7D 2E 34 00 06 00 14 04 03 01 10 1B 43 A9 9A") & _
        ChrH("9B 98 43 74 F2 04 39 13 24 7D 2E 34")
        
End Function


'set an away message
Public Function AwayMessage(strMessage As String) As String
    
    A1 = ChrA("0 2 0 4 0 0 0 0 0 4")
    AwayMessage = A1 & TLV(3, "text/aolrtf; charset=" & Chr(34) & "us-ascii" & Chr(34)) & TLV(4, strMessage)

End Function

'send an im
Public Function instantMessage(strRequestID As String, strUserName As String, strMessage As String, Optional boolConfirm As Boolean = False) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    If boolConfirm = False Then
        instantMessage = A1 & strRequestID & ChrA("0 1") & Chr(Len(strUserName)) & strUserName & _
            TLV(2, ChrA("5 1 0 3 1 1 2 1 1") & TwoByteLen(ChrA("0 0 0 0") & strMessage))
    Else
        instantMessage = A1 & strRequestID & ChrA("0 1") & Chr(Len(strUserName)) & strUserName & _
            TLV(2, ChrA("5 1 0 3 1 1 2 1 1") & TwoByteLen(ChrA("0 0 0 0") & strMessage)) & TLV(3, "")
    End If
    
End Function

'send an im
Public Function getBuddyIcon(strRequestID As String, strUserName As String, strMessage As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    getBuddyIcon = A1 & strRequestID & ChrA("0 1") & Chr(Len(strUserName)) & strUserName & _
        TLV(2, ChrA("5 1 0 3 1 1 2 1 1") & TwoByteLen(ChrA("0 0 0 0") & strMessage)) & ChrA("0 9 0 0")

End Function

'send a unicode im
Public Function unicodeMessage(strRequestID As String, strUserName As String, strMessage As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    unicodeMessage = A1 & strRequestID & ChrA("0 1") & Chr(Len(strUserName)) & strUserName & _
        TLV(2, ChrA("5 1 0 3 1 1 2 1 1") & TwoByteLen(ChrA("0 2 0 0") & strMessage))

End Function

'send a file
Public Function fileSend(strRequestID As String, strUserName As String, strFileName As String, strMessage As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    fileSend = A1 & strRequestID & ChrA("0 2") & Chr(Len(strUserName)) & strUserName & _
        TLV(5, ChrA("0 0") & strRequestID & ChrA("9 70 19 67 76 127 17 209 130 34 68 69 83 84 0 0 0 10 0 2 0 1 0 15 0 0 0 3 0 4 255 255 255 255 255 255 0 2 255 255") & TLV(12, strMessage) & ChrA("39 17") & TwoByteLen(ChrA("0 2 255 255 255 255 255 255") & strFileName)) & ChrA("0 3 0 0")

End Function

'Gets a file
Public Function FileGet(strRequestID As String, strUserName As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    FileGet = A1 & strRequestID & ChrA("0 2") & Chr(Len(strUserName)) & strUserName & _
        TLV(5, ChrA("0 0") & strRequestID & ChrA("9 70 19 72 76 127 17 209 130 34 68 69 83 84 0 0 0 10 0 2 0 1 0 15 0 0 0 3 0 4 H7F 0 0 1 0 5 0 2 HD H83 H27 H11 0 9 0 H12 0 2 0 0 0 1 0 H27 H12 0 8") & "us-ascii") & ChrA("0 3 0 0")

End Function

'Gets a file
Public Function bartGET(strUserName As String) As String

    A1 = ChrA("0 16 0 4 0 0 23 245 0 4")
    bartGET = A1 & Chr(Len(strUserName)) & strUserName & ChrA("1 0 2")
    
End Function

'Gets a file
Public Function bartPOST(intType As Integer, strData As String) As String

    A1 = ChrA("0 16 0 2 0 0 2 118 0 2")
    bartPOST = A1 & TLV(intType, strData)
    
End Function

'deny a file
Public Function talkDeny(strRequestID As String, strUserName As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    talkDeny = A1 & strRequestID & ChrA("0 2") & Chr(Len(strUserName)) & strUserName & _
        TLV(5, ChrA("0 1") & strRequestID & ChrH("09 8F 24 20 62 87 11 D1 82 22 44 45 53 54 00 00") & ChrA("0 11 0 2 0 1"))

End Function

'direct connect request
Public Function dcRequest(strRequestID As String, strUserName As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    dcRequest = A1 & strRequestID & ChrA("0 2") & Chr(Len(strUserName)) & strUserName & _
        TLV(5, ChrA("0 0") & strRequestID & ChrA("9 70 19 69 76 127 17 209 130 34 68 69 83 84 0 0 0 10 0 2 0 1 0 3 0 4 24 16 172 135 0 5 0 2 20 70 0 15 0 0")) & ChrA("0 3 0 0")

End Function

'direct connect request
Public Function vdRequest(strRequestID As String, strUserName As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    vdRequest = A1 & strRequestID & ChrA("0 2") & Chr(Len(strUserName)) & strUserName & _
        TLV(5, ChrA("0 0") & strRequestID & ChrA("9 70 1 1 76 127 17 209 130 34 68 69 83 84 0 0 0 10 0 2 0 1 0 15 0 0")) & ChrA("0 3 0 0")

End Function

'game invite
Public Function inviteGame(strRequestID As String, strUserName As String, strGameURL As String, strGameName As String, strMessage As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    inviteGame = A1 & strRequestID & ChrA("0 2") & Chr(Len(strUserName)) & strUserName & _
        TLV(5, ChrA("0 0") & strRequestID & ChrA("9 70 19 71 76 127 17 209 130 34 68 69 83 84 0 0 0 10 0 2 0 1 0 15 0 0 0 14") & TwoByteLen("en") & ChrA("0 13") & TwoByteLen("us-ascii") & ChrA("0 12") & TwoByteLen(strMessage) & ChrA("0 3 0 4 64 163 30 79 0 5 0 2 20 70 0 7") & TwoByteLen(strGameURL) & ChrA("39 17") & TwoByteLen(ChrA("0 0 2 0 5 7 76 127 17 209 130 34 68 69 83 84 0 0 0 11 0 9") & strGameName & Chr(0) & "Fuck you" & ChrA("0 0 0 0 0"))) & ChrA("0 3 0 0")

End Function

'chat invite
Public Function InviteChat(strRequestID As String, strUserName As String, strChatRoomURL As String, strMessage As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    InviteChat = A1 & strRequestID & ChrA("0 2") & Chr(Len(strUserName)) & strUserName & _
        TLV(5, ChrA("0 0") & strRequestID & ChrA("116 143 36 32 98 135 17 209 130 34 68 69 83 84 0 0 0 10 0 2 0 1 0 15 0 0") & TLV(14, "en") & TLV(13, "us-ascii") & TLV(12, strMessage) & ChrA("39 17") & TwoByteLen(ChrA("0 4") & Chr(Len(strChatRoomURL)) & strChatRoomURL & ChrA("0 0"))) & ChrA("0 3 0 0")

End Function

'buddy list send
Public Function buddyList(strRequestID As String, strUserName As String, strBuddyList As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    buddyList = A1 & strRequestID & ChrA("0 2") & Chr(Len(strUserName)) & strUserName & _
        TLV(5, ChrA("0 0") & strRequestID & ChrA("9 70 19 75 76 127 17 209 130 34 68 69 83 84 0 0 0 10 0 2 0 1 0 15 0 0 39 17") & TwoByteLen(strBuddyList)) & ChrA("0 3 0 0")

End Function

'warn a user
Public Function UserWarning(strUserName As String, intType As Integer) As String
    
    A1 = ChrA("0 4 0 8 0 0 0 9 0 8 0 " & intType)
    UserWarning = A1 & Chr(Len(strUserName)) & strUserName

End Function

'block/unblock a user
Public Function userBlock(strUserName As String, intType As Integer) As String
    
    A1 = ChrA("0 19 0 " & intType & " 0 0 0 6 0 " & intType)
    userBlock = A1 & TwoByteLen(strUserName) & ChrA("0 0 11 17 0 3 0 0")

End Function

'look online / offline
Public Function adjustPrivacy(boolOnline As Boolean) As String
    
    A1 = ChrA("0 1 0 30 0 0 0 0 0 30")
    If boolOnline Then
        adjustPrivacy = A1 & TLV(6, ChrA("0 0 0 0"))
    Else
        adjustPrivacy = A1 & TLV(6, ChrA("0 0 1 0"))
    End If

End Function

'talk request
Public Function requestTalk(strRequestID As String, strUserName As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    requestTalk = A1 & strRequestID & ChrA("0 2") & Chr(Len(strUserName)) & strUserName & _
        TLV(5, ChrA("0 0") & strRequestID & ChrA("9 70 19 65 76 127 17 209 130 34 68 69 83 84 0 0 0 10 0 2 0 1 0 3 0 4 127 0 0 1 0 255 0 0 39 17 0 4 0 0 0 1")) & ChrA("0 3 0 0")

End Function

'send a theme
Public Function SendTheme(strRequestID As String, strUserName As String, strMessage As String, strThemeName As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    SendTheme = ChrA("0 4 0 6 0 0 0 9 0 6") & strRequestID & ChrA("0 1") & Chr(Len(strUserName)) & strUserName & _
        TLV(2, ChrA("5 1 0 3 1 1 2 1 1") & TwoByteLen(ChrA("0 0 0 0") & strMessage)) & TLV(13, TLV(1, Chr(5) & strThemeName) & TLV(128, Chr(5) & strThemeName) & TLV(130, Chr(5) & strThemeName)) & TLV(9, "")

End Function

'send a theme
Public Function SendTheme2(strRequestID As String, strUserName As String, strMessage As String, strThemeName As String, strThemeName2, strThemeName3) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    SendTheme2 = ChrA("0 4 0 6 0 0 0 9 0 6") & strRequestID & ChrA("0 1") & Chr(Len(strUserName)) & strUserName & _
        TLV(2, ChrA("5 1 0 3 1 1 2 1 1") & TwoByteLen(ChrA("0 0 0 0") & strMessage)) & TLV(13, TLV(1, Chr(5) & strThemeName3) & TLV(128, Chr(5) & strThemeName) & TLV(130, Chr(5) & strThemeName2))

End Function

'istyping send
Public Function setTalk(intType As Integer, strUserName As String) As String

    A1 = ChrA("0 4 0 20 0 0 0 0 0 20")
    setTalk = A1 & ChrA("0 0 0 0 0 0 0 0 0 1") & Chr(Len(strUserName)) & strUserName & Chr(0) & Chr(intType)

End Function

'store a comment for a user
Public Function setComment(strName As String, strComment As String) As String

    A1 = ChrA("0 19 0 9 0 0 0 6")
    setComment = A1 & TLV(9, strName) & ChrA("96 147 93 230 0 0 0 7") & TLV(316, strComment)

End Function

'send a stupid email to someone
Public Function inviteFriend(strEmail As String, strMessage As String) As String

    A1 = ChrA("0 6 0 2 0 0 0 1 0 2")
    inviteFriend = A1 & TLV(17, strEmail) & TLV(21, "asdf")

End Function

'that little `createÿÿ` packet
Public Function chatCreate(strChannel As String) As String

    A1 = ChrA("0 13 0 8 0 0 0 1 0 8 0 4 6")
    chatCreate = A1 & "create" & ChrA("255 255 1 0 3") & TLV(215, "en") & TLV(214, "us-ascii") & TLV(211, strChannel)

End Function

'send something to a chat room
Public Function chatSend(strRequestID As String, strMessage As String) As String

    A1 = ChrA("0 14 0 5 0 0 0 0 0 5")
    chatSend = A1 & strRequestID & ChrA("0 3 0 1 0 0 0 6 0 0") & TLV(5, TLV(2, "us-ascii") & TLV(3, "en") & TLV(1, strMessage))

End Function

'get member info
Public Function getMInfo(strUserName As String, Optional intType As Integer = "1") As String
    
    A1 = ChrA("0 2 0 21 0 0 0 21 0 21")
    getMInfo = A1 & ChrA("0 0 0 " & intType) & Chr(Len(strUserName)) & strUserName

End Function

'get member info
Public Function getSNBE(strUserName As String, Optional intType As Integer = "1") As String
    
    A1 = ChrA("0 10 0 2 0 0 0 2 0 2")
    getSNBE = A1 & strUserName

End Function

'set directory info
Public Function setDir(strParams As String) As String

    A1 = ChrA("0 2 0 9 0 0 0 31 0 9")
    A1 = A1 & TLV(28, 0) & TLV(10, ChrA("0 2"))
    A2 = Split(strParams, Chr(1))
    For i = 0 To UBound(A2)
        A1 = A1 & TLV(i + 1, CStr(A2(i)))
        DoEvents
    Next i
    setDir = A1
    
End Function

Function GetDir(strUserName As String)
     
    A1 = ChrA("0 2 0 11 0 0 0 0 0 11")
    GetDir = A1 & Chr(Len(strUserName)) & strUserName
    
End Function

'get member info
Public Function nameFormat(strFormattedName As String, intType As Integer) As String
    
    A1 = ChrA("0 7 0 4 0 0 0 1 0 4")
    nameFormat = A1 & TLV(intType, strFormattedName)
    
End Function

Public Function infoQuery() As String

    A1 = ChrA("0 7 0 2 0 0 20 244 0 2")
    infoQuery = A1

End Function

'get member info
Public Function emailChange(strEmail As String) As String
    
    A1 = ChrA("0 7 0 4 0 0 0 1 0 4")
    emailChange = A1 & TLV(17, strEmail)
    
End Function

'get member info
Public Function passChange(strOld As String, strNew As String) As String
    
    A1 = ChrA("0 7 0 4 0 0 0 1 0 4")
    passChange = A1 & TLV(1, strNew) & TLV(18, strOld)
    
End Function

'format name server change
Public Function servFormat() As String

    A1 = ChrA("0 1 0 4 0 0 0 9 0 4 0 7")
    servFormat = A1

End Function

'format name server change
Public Function servLink(strUser As String, strPassword As String, strChallenge As String) As String

    A1 = ChrA("0 1 0 4 0 0 0 1 0 4 0 1")
    servLink = A1 & TLV(40, AIMEncryptPw(strChallenge, strPassword)) & TLV(1, strUser)

End Function

'chatting server change (1st one)
Public Function servChat() As String

    A1 = ChrA("0 1 0 4 0 0 0 3 0 4 0 13")
    servChat = A1

End Function

'try and connect to a new server for ims
Public Function servFriend(intASDFAsdf) As String

    A1 = ChrA("0 1 0 4 0 0 0 13 0 4" & intASDFAsdf)
    servFriend = A1

End Function

'try and connect to a new server for ims
Public Function servAny(intASDFAsdf) As String

    A1 = ChrA("0 1 0 4 0 0 0 13 0 4 0 15")
    servAny = A1

End Function

'chatting server change (2nd one)
Public Function servChat2(strChatURL As String) As String

    A1 = ChrA("0 1 0 4 0 0 0 4 0 4 0 14")
    servChat2 = A1 & TLV(1, ChrA("0 4") & Chr(Len(strChatURL)) & strChatURL & ChrA("0 0"))

End Function

'final login step
Public Function clientReady() As String

    A1 = ChrA("0 1 0 2 0 0 0 0 2 0")
    clientReady = A1 & _
        ChrA("0 1 0 3 1 16 6 41") & _
        ChrA("0 19 0 3 1 16 6 41") & _
        ChrA("0 2 0 1 1 16 6 41") & _
        ChrA("0 3 0 1 1 16 6 41") & _
        ChrA("0 4 0 1 1 16 6 41") & _
        ChrA("0 6 0 1 1 16 6 41") & _
        ChrA("0 8 0 1 1 16 6 41") & _
        ChrA("0 9 0 1 1 16 6 41") & _
        ChrA("0 10 0 1 1 16 6 41") & _
        ChrA("0 11 0 1 1 4 0 1")

End Function

'final login step
Public Function clientReady2() As String

    A1 = ChrA("0 1 0 2 0 0 0 0 2 0")
    clientReady2 = A1 & _
        ChrA("0 1 0 3 1 16 6 41") & _
        ChrA("0 19 0 3 1 16 6 41") & _
        ChrA("0 2 0 1 1 16 6 41") & _
        ChrA("0 3 0 1 1 16 6 41") & _
        ChrA("0 4 0 1 1 16 6 41") & _
        ChrA("0 6 0 1 1 16 6 41") & _
        ChrA("0 8 0 1 1 16 6 41") & _
        ChrA("0 9 0 1 1 16 6 41") & _
        ChrA("0 10 0 1 1 16 6 41") & _
        ChrA("0 11 0 1 1 4 0 1")

End Function
'final login step
Public Function clientReadyJacked() As String

    A1 = ChrA("0 1 0 2 0 0 0 0 2 0")
    clientReadyJacked = A1 & _
        ChrA("0 1 0 3 1 16 6 41") & _
        ChrA("0 19 0 3 1 16 6 41") & _
        ChrA("0 2 0 1 1 16 6 41") & _
        ChrA("0 3 0 1 1 16 6 41") & _
        ChrA("0 6 0 1 1 16 6 41") & _
        ChrA("0 8 0 1 1 16 6 41") & _
        ChrA("0 9 0 1 1 16 6 41") & _
        ChrA("0 10 0 1 1 16 6 41") & _
        ChrA("0 11 0 1 1 4 0 1")
        
End Function

'final format handshake step
Public Function formatReady() As String

    A1 = ChrA("0 1 0 2 0 0 0 0 0 2")
    formatReady = A1 & ChrA("0 1 0 4 0 16 8 214 0 7 0 1 0 16 8 214")
    
End Function

'final chat handshake step
Public Function chatReady() As String

    A1 = ChrA("0 1 0 2 0 0 0 0 0 2")
    chatReady = A1 & ChrA("0 1 0 3 0 16 6 208 0 13 0 1 0 16 6 208")
    
End Function

'final search handshake step
Public Function searchReady() As String

    searchReady = ChrA("0 1 0 2 0 0 0 0 0 2 0 1 0 3 0 16 8 63 0 15 0 1 0 16 8 63")
    
End Function

'final bart handshake step
Public Function bartReady() As String

    bartReady = ChrA("0 1 0 2 0 0 0 0 0 2 0 1 0 4 0 16 8 214 0 16 0 1 0 16 8 214")
    
End Function

'final chat handshake step 2
Public Function chatReady2() As String

    A1 = ChrA("0 1 0 2 0 0 0 0 0 2")
    chatReady2 = A1 & ChrA("0 1 0 3 0 16 6 208 0 14 0 1 0 16 6 208")
    
End Function

'add a buddy
Public Function buddyAdd(strUserName As String) As String
    
    A1 = ChrA("0 3 0 4 0 0 0 0 0 0")
    buddyAdd = A1 & Chr(Len(strUserName)) & strUserName
    
End Function

'password encryption
Public Function EncryptPW(ByRef strPass As String) As String
    Dim arrTable() As Variant
    Dim strEncrypted As String
    Dim lngX As Long
    Dim strHex As String
    
    arrTable = Array(243, 179, 108, 153, 149, 63, 172, 182, 197, 250, 107, 99, 105, 108, 195, 154)
    
    For lngX = 0 To Len(strPass$) - 1
        strHex = Chr(Asc(Mid(strPass, lngX + 1, 1)) Xor CLng(arrTable((lngX Mod 16))))
        strEncrypted = strEncrypted & strHex
    Next
    
    EncryptPW = strEncrypted
End Function

'capability block
Public Function FULLCAP()
FULLCAP = _
    ChrH("09 46 13 4A 4C 7F 11 D1 82 22 44 45 53 54 00 00") & _
    ChrH("09 46 13 4B 4C 7F 11 D1 82 22 44 45 53 54 00 00") & _
    ChrH("09 8F 24 20 62 87 11 D1 82 22 44 45 53 54 00 00") & _
    ChrH("09 46 13 4D 4C 7F 11 D1 82 22 44 45 53 54 00 00") & _
    ChrH("09 46 13 41 4C 7F 11 D1 82 22 44 45 53 54 00 00") & _
    ChrH("09 46 00 00 4C 7F 11 D1 82 22 44 45 53 54 00 00") & _
    ChrH("09 46 13 43 4C 7F 11 D1 82 22 44 45 53 54 00 00") & _
    ChrH("09 46 01 FF 4C 7F 11 D1 82 22 44 45 53 54 00 00") & _
    ChrH("09 46 00 01 4C 7F 11 D1 82 22 44 45 53 54 00 00") & _
    ChrH("09 46 13 45 4C 7F 11 D1 82 22 44 45 53 54 00 00") & _
    ChrH("09 46 13 46 4C 7F 11 D1 82 22 44 45 53 54 00 00") & _
    ChrH("09 46 13 47 4C 7F 11 D1 82 22 44 45 53 54 00 00")
    
'1  - ????
'2  - ????
'3  - chat
'4  - ????
'5  - talk
'6  - ????
'7  - file transfers
'8  - ????
'9  - secruity enabled
'10 - direct im
'11 - buddy icon
'12 - add-ins
End Function

'i can do anything! yaaay!
Public Function SUPERCAP()
    
    SUPERCAP = _
        CAPIT("5D5E1708-55AA-11D3-B143-0060B0FB1ECB") & _
        CAPIT("5D5E1709-55AA-11D3-B143-0060B0FB1ECB") & _
        CAPIT("84350CC8-E401-11D5-9754-0060B0EE0631") & _
        CAPIT("0946134E-4C7F-11D1-8222-444553540000") & _
        CAPIT("0946134D-4C7F-11D1-8222-444553540000") & _
        CAPIT("0946134B-4C7F-11D1-8222-444553540000") & _
        CAPIT("0946134A-4C7F-11D1-8222-444553540000") & _
        CAPIT("09461349-4C7F-11D1-8222-444553540000") & _
        CAPIT("09461348-4C7F-11D1-8222-444553540000") & _
        CAPIT("09461347-4C7F-11D1-8222-444553540000") & _
        CAPIT("09461346-4C7F-11D1-8222-444553540000") & _
        CAPIT("09461345-4C7F-11D1-8222-444553540000") & _
        CAPIT("09461344-4C7F-11D1-8222-444553540000") & _
        CAPIT("09461343-4C7F-11D1-8222-444553540000") & _
        CAPIT("09461342-4C7F-11D1-8222-444553540000") & _
        CAPIT("09461341-4C7F-11D1-8222-444553540000") & _
        CAPIT("0946F00F-4C7F-11D1-8222-444553540000") & _
        CAPIT("0946F000-4C7F-11D1-8222-444553540000") & _
        CAPIT("0946E0FF-4C7F-11D1-8222-444553540000") & _
        CAPIT("0946E000-4C7F-11D1-8222-444553540000") & _
        CAPIT("09460100-4C7F-11D1-8222-444553540000") & _
        CAPIT("09460001-4C7F-11D1-8222-444553540000") & _
        CAPIT("09460000-4C7F-11D1-8222-444553540000")
    
    SUPERCAP = SUPERCAP & _
        CAPIT("5D5E170A-55AA-11D3-B143-0060B0FB1ECB") & _
        CAPIT("5D5E170E-55AA-11D3-B143-0060B0FB1ECB") & _
        CAPIT("78382132-8E20-11D3-9830-8F6ED783EDD8") & _
        CAPIT("58CF279A-9911-11D3-A159-D57FA28A7112") & _
        CAPIT("9269D81C-A6A2-11D3-AEE2-E2489E91D6E7") & _
        CAPIT("33DD2CBA-0FDC-11D4-B073-BC1838CEF8C8") & _
        CAPIT("3B4056D2-7457-11D4-A1D5-DC243835AD24") & _
        CAPIT("B3809AD8-0DBA-11D5-9F8A-0060B0EE0631") & _
        CAPIT("3DFA0724-2AD7-11D5-9773-0060B0EE0631") & _
        CAPIT("67400b76-3516-11d5-8b6c-0060b0ee0631") & _
        CAPIT("200A0000-A289-11D3-A52D-001083341CFA") & _
        CAPIT("200A0001-A289-11D3-A52D-001083341CFA") & _
        CAPIT("200A000A-A289-11D3-A52D-001083341CFA") & _
        CAPIT("200A000B-A289-11D3-A52D-001083341CFA") & _
        CAPIT("748F2420-6287-11D1-8222-444553540000") & _
        CAPIT("97B12751-243C-4334-AD22-D6ABF73F1492") & _
        CAPIT("2E7A6475-FADF-4dc8-886F-EA3595FDB6DF") & _
        CAPIT("DD16F202-84E6-11D4-90DB-00104B9B4B7D") & _
        CAPIT("7f53f598-05aa-11d4-a5ad-001083341cfa") & _
        CAPIT("3e8001da-de70-11d3-a57d-001083341cfa") & _
        CAPIT("3e8001db-de70-11d3-a57d-001083341cfa") & _
        CAPIT("3ff9b60d-c55b-4d4f-b9d0-ca6a0f5783be") & _
        CAPIT("50000000-2a9a-11d5-808a-0060b0ee0631") & _
        CAPIT("50000001-2a9a-11d5-808a-0060b0ee0631")
        
    SUPERCAP = SUPERCAP & _
        CAPIT("c28e07ec-4486-11d5-9673-0060b0ee0631") & _
        CAPIT("50020002-2a9a-11d5-808a-0060b0ee0631") & _
        CAPIT("50010002-2a9a-11d5-808a-0060b0ee0631") & _
        CAPIT("50010001-2a9a-11d5-808a-0060b0ee0631") & _
        CAPIT("50010000-2a9a-11d5-808a-0060b0ee0631") & _
        CAPIT("50000002-2a9a-11d5-808a-0060b0ee0631")

End Function

Public Function nameDelete(strPass As String) As String
    
    A1 = ChrA("0 7 0 8 0 0 0 0 0 0")
    nameDelete = A1 & TLV(2, strPass) & TLV(18, strPass)
    
End Function

Function CAPIT(strStrin)
    
    Dim strGay
    strStrin = Replace(strStrin, "-", "")
    strGay = ""
    For i = 0 To (Len(strStrin) / 2) - 1
        strGay = strGay & Chr("&H" & Mid(strStrin, (i * 2) + 1, 2))
        DoEvents
    Next i
    CAPIT = strGay

End Function

Public Function TLV(intType As Integer, strData As String) As String
    TLV = IntegerToBase256(intType) & TwoByteLen(strData)
End Function

