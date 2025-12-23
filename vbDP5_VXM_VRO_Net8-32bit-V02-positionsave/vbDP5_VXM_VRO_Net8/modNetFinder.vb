Option Strict Off
Option Explicit On
Module modNetFinder

    '=========================================================
    'Purpose: <BR>
    '    The Netfinder protocol allows a PC application to search for embedded systems on a network.<BR>
    '    For USB, the NetFinder Packet is used to Identify DPP devices<BR>

    Public InetShareDpp As Boolean
    Public InetLockDpp As Boolean
    Public Const MAX_NETFINDER_ENTRIES As Short = 4
    Public Const NO_NETFINDER_ENTRIES As Short = -1
    Public Const MAX_NETFINDER_STRING As Short = 80
    Public Netfinder_active As Boolean
    Public Netfinder_Seq As Short
    Public NetfinderUnit As Short
    Public EthernetConnectionStatus As Short
    Public EthernetConnectionTicker As Short
   
    Public UnitIP(MAX_NETFINDER_ENTRIES - 1, 3) As Object ' capture up to 16 IPs per Netfinder?
    Public EthernetConnected As Boolean
    Public SearchIp As Integer

    Public Structure NETDISPLAY_ENTRY
        Dim alert_level As Byte
        Dim event1_days As Integer
        Dim event1_hours As Byte
        Dim event1_minutes As Byte
        Dim event1_seconds As Byte
        Dim event2_days As Integer
        Dim event2_hours As Byte
        Dim event2_minutes As Byte
        Dim event2_seconds As Byte
        <VBFixedArray(6)> Dim mac() As Byte
        <VBFixedArray(4)> Dim ip() As Byte
        Dim port As Short
        <VBFixedArray(4)> Dim subnet() As Byte
        <VBFixedArray(4)> Dim gateway() As Byte
        Dim str_type As String
        Dim str_description As String
        Dim str_ev1 As String
        Dim str_ev2 As String
        Dim str_display As String
        Dim ccConnectRGB As Integer 'COLORREF
        Dim HasData As Boolean
        Dim SockAddr As UInt32
       
        Public Sub Initialize()
            ReDim mac(6)
            ReDim ip(4)
            ReDim subnet(4)
            ReDim gateway(4)
        End Sub
    End Structure
   
    Public NfPktEntries(MAX_NETFINDER_ENTRIES - 1) As NETDISPLAY_ENTRY 'RS232,USB,ETHERNET,SPARE
    Public DppEntries(MAX_NETFINDER_ENTRIES - 1) As NETDISPLAY_ENTRY
   
    Public Sub AddStr(ByRef strStr As String, ByRef str1 As String, Optional ByRef str2 As String = "", Optional ByRef AddNL As Boolean = True)
        Dim strNL As String
        strNL = ""
        If (AddNL) Then strNL = vbCrLf
        strStr = strStr & str1 & str2 & strNL
    End Sub
   
    'IpAddrConvert converts (with ToByteArray = False)
    'long to string
    'byte array to string
    'string to long
    '
    'IpAddrConvert converts (with ToByteArray = True)
    'long to byte array
    'byte array to string (no change)
    'string to byte array
    Public Function IpAddrConvert(ByRef IPAddress As Object, Optional ByRef ToByteArray As Boolean = False) As Object
        Dim idxByte As Integer 'byte index
        Dim idxPos As Integer 'position in hex string
        Dim idxPosOld As Integer
        Dim strHex As String
        Dim strSep As String
        Dim ipByte As Byte
        Dim ipByteArr(3) As Byte
        Dim iAddr32 As UInt32
        Dim strIpAddr As String

        IpAddrConvert = IPAddress
        If IsNumeric(IPAddress) Then 'convert long
            strIpAddr = ""
            strSep = ""
            strHex = Hex(IPAddress)
            strHex = New String("0", 8 - Len(strHex)) & strHex
            For idxByte = 0 To 3
                idxPos = 1 + (idxByte * 2)
                ipByte = CByte("&H" & Mid(strHex, idxPos, 2)) 'convert hex to byte
                ipByteArr(idxByte) = ipByte
                strIpAddr = strIpAddr & strSep & CStr(ipByte)
                strSep = "."
            Next
            If (ToByteArray) Then
                IpAddrConvert = ipByteArr
            Else
                IpAddrConvert = strIpAddr
            End If
        ElseIf IsArray(IPAddress) Then  'ByteArray(0-3)
            If ((TypeName(IPAddress) = "Byte()") And (UBound(IPAddress) = 3)) Then
                strIpAddr = ""
                strSep = ""
                For idxByte = 0 To 3
                    ipByte = IPAddress(idxByte)
                    strIpAddr = strIpAddr & strSep & CStr(ipByte)
                    strSep = "."
                Next
                IpAddrConvert = strIpAddr
            End If
        ElseIf (TypeName(IPAddress) = "String") Then  'String 0.0.0.0
            idxPos = 1
            idxPosOld = 0
            strIpAddr = ""
            For idxByte = 0 To 3
                idxPos = InStr(idxPosOld + 1, IPAddress, ".")
                If (idxPos = 0) Then idxPos = Len(IPAddress) + 1 'end of string
                ipByteArr(idxByte) = Val(Mid(IPAddress, idxPosOld + 1, (idxPos - idxPosOld) - 1)) And &HFFS
                idxPosOld = idxPos
                strHex = Hex(ipByteArr(idxByte))
                strHex = New String("0", 2 - Len(strHex)) & strHex
                strIpAddr = strIpAddr & strHex
            Next
            If (ToByteArray) Then
                IpAddrConvert = ipByteArr
            Else
                iAddr32 = CULng("&H" & strIpAddr)
                IpAddrConvert = iAddr32
            End If
        End If
    End Function

    'IP Address Convert
    'long . string >. string=IpAddrConvert(long)
    'long . byte array >. byte array=IpAddrConvert(long,true)
    'string . long >. long=IpAddrConvert(string)
    'string . byte array >. byte array=IpAddrConvert(string,true)
    'byte array . string >. string=IpAddrConvert(byte array)
    '(2 steps for byte array to long)
    'byte array . long >. string=IpAddrConvert(byte array) >>. long=IpAddrConvert(string)

    Public Sub InitEntry(ByRef pEntry As NETDISPLAY_ENTRY)
        Dim idxInit As Integer
        pEntry.alert_level = 0

        pEntry.event1_days = 0
        pEntry.event1_hours = 0
        pEntry.event1_minutes = 0

        pEntry.event2_days = 0
        pEntry.event2_hours = 0
        pEntry.event2_minutes = 0

        pEntry.event1_seconds = 0
        pEntry.event2_seconds = 0

        pEntry.port = 0 ' Get port from UDP header

        pEntry.Initialize()
        For idxInit = 0 To MAX_NETFINDER_ENTRIES - 1
            pEntry.mac(idxInit) = 0
            pEntry.ip(idxInit) = 0
            pEntry.subnet(idxInit) = 0
            pEntry.gateway(idxInit) = 0
        Next

        pEntry.mac(4) = 0
        pEntry.mac(5) = 0

        pEntry.str_type = ""
        pEntry.str_description = ""
        pEntry.str_ev1 = ""
        pEntry.str_ev2 = ""

        pEntry.str_display = ""
        pEntry.ccConnectRGB = colorWhite

        pEntry.HasData = False

        pEntry.SockAddr = 0
    End Sub

    Public Sub InitAllEntries(ByRef DppEntries() As NETDISPLAY_ENTRY)
        Dim idxEntry As Integer
        Dim lMaxEntry As Integer
       
        lMaxEntry = UBound(DppEntrieS)
        For idxEntry = 0 To lMaxEntry
            Call InitEntry(DppEntries(idxEntry))
        Next
    End Sub
   
    Public Function AlertEntryToCOLORREF(ByRef pEntry As NETDISPLAY_ENTRY) As Integer 'COLORREF
        AlertEntryToCOLORREF = AlertByteToCOLORREF(pEntry.alert_level)
    End Function
   
    Public Function AlertByteToCOLORREF(ByRef alert_level As Byte) As Integer 'COLORREF
        Select Case alert_level ' Alert Level
            Case &H0s '0 = Interface is open (unconnected)
                AlertByteToCOLORREF = colorGreen
            Case &H1s '1 = Interface is connected (sharing is allowed)
                AlertByteToCOLORREF = colorYellow
            Case &H2s '2 = Interface is connected (sharing is not allowed)
                AlertByteToCOLORREF = colorRed
            Case &H3s '3 = Interface is locked
                AlertByteToCOLORREF = colorRed
            Case &H4s '4 = Interface is unavailable because USB is connected
                AlertByteToCOLORREF = colorSilver
            Case Else 'Interface configuration unknown
                AlertByteToCOLORREF = colorWhite
        End Select
    End Function
   
    Public Function FindEntry(ByRef SockAddr As Integer, ByRef DppEntries() As NETDISPLAY_ENTRY) As Integer
        Dim idxEntry As Integer
        Dim EntryFound As Integer
        Dim lMaxEntry As Integer
       
        lMaxEntry = UBound(DppEntrieS)
        EntryFound = NO_NETFINDER_ENTRIES
        For idxEntry = 0 To lMaxEntry
            If (SockAddr = DppEntries(idxEntry).SockAddr) Then
                EntryFound = idxEntry
                Exit For
            End If
        Next
        FindEntry = EntryFound
    End Function
   
    Public Sub AddEntry(ByRef pEntry As NETDISPLAY_ENTRY, ByRef Buffer() As Byte, ByRef destPort As Integer)
        Dim idxBuffer As Integer 'buffer index
        Dim nfVersion As Byte
        Dim RandSeqNumber As Integer
       
        nfVersion = Buffer(&H0S) 'byte 0 (for reference)
        pEntry.alert_level = Buffer(&H1S) 'byte 1
        RandSeqNumber = (CInt(Buffer(&H2S)) * CInt(256)) + CInt(Buffer(&H3S)) 'bytes 2 & 3 (for reference)
        pEntry.port = CShort("&H" & Hex(destPort And &HFFFFS)) 'convert long containing unsigned short
       
        'Event 1
        pEntry.event1_days = EventDays_ByteToULong(Buffer, &H4S) 'Event 1 days bytes 4 & 5
        pEntry.event1_hours = Buffer(&H6S) 'Event 1 hours byte 6
        pEntry.event1_minutes = Buffer(&H7S) 'Event 1 minutes byte 7
       
        'Event 2
        pEntry.event2_days = EventDays_ByteToULong(Buffer, &H8S) 'Event 2 days bytes 8 & 9
        pEntry.event2_hours = Buffer(&HAS) 'Event 2 hours byte 10
        pEntry.event2_minutes = Buffer(&HBS) 'Event 2 minutes byte 11
       
        'Event 1 & 2 seconds
        pEntry.event1_seconds = Buffer(&HCS) 'Event 1 seconds byte 12
        pEntry.event2_seconds = Buffer(&HDS) 'Event 1 seconds byte 13
       
        pEntry.mac(0) = Buffer(&HES) 'MAC Address byte 14 (6 bytes)
        pEntry.mac(1) = Buffer(&HFS)
        pEntry.mac(2) = Buffer(&H10S)
        pEntry.mac(3) = Buffer(&H11S)
        pEntry.mac(4) = Buffer(&H12S)
        pEntry.mac(5) = Buffer(&H13S)
       
        pEntry.ip(0) = Buffer(&H14S) 'IP Address byte 20 (4 bytes)
        pEntry.ip(1) = Buffer(&H15S)
        pEntry.ip(2) = Buffer(&H16S)
        pEntry.ip(3) = Buffer(&H17S)
       
        pEntry.subnet(0) = Buffer(&H18S) 'Subnet Mask byte 24 (4 bytes)
        pEntry.subnet(1) = Buffer(&H19S)
        pEntry.subnet(2) = Buffer(&H1AS)
        pEntry.subnet(3) = Buffer(&H1BS)
       
        pEntry.gateway(0) = Buffer(&H1CS) 'Default Gateway byte 28 (4 bytes)
        pEntry.gateway(1) = Buffer(&H1DS)
        pEntry.gateway(2) = Buffer(&H1ES)
        pEntry.gateway(3) = Buffer(&H1FS)
       
        idxBuffer = &H20s 'start of variable Length Null-Terminated strings byte 32
        'buffer index incremented to start of next string
        pEntry.str_type = GetNetFinderString(Buffer, idxBuffer) 'get embedded system
        pEntry.str_description = GetNetFinderString(Buffer, idxBuffer) 'get description
       
        ' Copy the Event1+2 text descripion without the ": "
        pEntry.str_ev1 = GetNetFinderString(Buffer, idxBuffer) 'get event 1 description
        pEntry.str_ev2 = GetNetFinderString(Buffer, idxBuffer) 'get event 2 description
       
        ' Add ": " to str_ev1 and str_ev2
        pEntry.str_ev1 = pEntry.str_ev1 & ": "
        pEntry.str_ev2 = pEntry.str_ev2 & ": "
       
        pEntry.ccConnectRGB = AlertByteToCOLORREF(pEntry.alert_level)
        pEntry.HasData = True
        pEntry.SockAddr = SockAddr_ByteToULong(Buffer, &H14S) 'IP Address (starts at byte 20)
        pEntry.str_display = EntryToStr(pEntry) ' Convert Entry infor to string
        'MsgBox pEntry.str_display
    End Sub
   
    Private Function GetNetFinderString(ByRef Buffer() As Byte, ByRef Index As Integer) As String
        Dim idxCh As Integer
        Dim strCh As String
        strCh = ""
        idxCh = Index
        Do
            If (Buffer(idxCh) > 0) And (idxCh < UBound(Buffer)) Then
                strCh = strCh & Chr(Buffer(idxCh))
                idxCh = idxCh + 1
            Else
                idxCh = idxCh + 1 'start of next string
                Exit Do
            End If
        Loop
        Index = idxCh 'update index to next position
        GetNetFinderString = strCh
    End Function

    Private Function GetSerialNumber(ByVal str_type As String) As Integer
        Dim lSerNum As Integer
        Dim lPos As Integer
        Dim strSerNum As String = ""
   
        lSerNum = 0
        If (Len(str_type) > 0) Then
            lPos = InStr(1, str_type, "S/N", CompareMethod.Text)
            If (lPos > 0) Then
                If ((lPos + 3) < Len(str_type)) Then
                    lSerNum = CLng(Math.Truncate(Convert.ToDouble(Mid(str_type, lPos+3))))
                End If
            End If
        End If
        GetSerialNumber = lSerNum
    End Function

    Private Function EventDays_ByteToULong(ByRef Buffer() As Byte, ByRef Index As Integer) As UInt32
        Dim lngDays As Integer
        lngDays = CInt(Buffer(Index)) * 256
        lngDays = lngDays + CInt(Buffer(Index + 1))
        EventDays_ByteToULong = lngDays
    End Function
   
    Private Function SockAddr_ByteToULong(ByRef Buffer() As Byte, ByRef Index As Integer) As UInt32
        Dim idxIp As Integer
        Dim byteIp(3) As Byte
        Dim strIp As String = ""
        For idxIp = 0 To 3
            byteIp(idxIp) = Buffer(Index + idxIp)
        Next
        strIp = IpAddrConvert(byteIp)
        SockAddr_ByteToULong = IpAddrConvert(strIp)
    End Function
   
    Private Function inc(ByRef varIndex As Object) As Integer
        varIndex = varIndex + 1
        inc = CInt(varIndex)
    End Function
   
    Public Function EntryToStr(ByRef pEntry As NETDISPLAY_ENTRY) As String
        Dim cstrAlertLevel As String = ""
        Dim strDeviceType As String = ""
        Dim cstrIpAddress As String = ""
        Dim cstrEvent1 As String = ""
        Dim cstrEvent2 As String = ""
        Dim cstrAdditionalDesc As String = ""
        Dim cstrMacAddress As String = ""
        Dim cstrEntry As String = ""
        Dim temp_str As String = ""
        Dim temp_str2 As String = ""
        Dim idxVal As Integer
        Dim strSep As String = ""
       
        cstrAlertLevel = "Connection: "
        Select Case (pEntry.alert_level) ' Alert Level
            Case &H0s '0 = Interface is open (unconnected)
                cstrAlertLevel = cstrAlertLevel & "Interface is open (unconnected)"
            Case &H1s '1 = Interface is connected (sharing is allowed)
                cstrAlertLevel = cstrAlertLevel & "Interface is connected (sharing is allowed)"
            Case &H2s '2 = Interface is connected (sharing is not allowed)
                cstrAlertLevel = cstrAlertLevel & "Interface is connected (sharing is not allowed)"
            Case &H3s '3 = Interface is locked
                cstrAlertLevel = cstrAlertLevel & "Interface is locked"
            Case &H4s '4 = Interface is unavailable because USB is connected
                cstrAlertLevel = cstrAlertLevel & "Interface is unavailable because USB is connected"
            Case Else
                cstrAlertLevel = cstrAlertLevel & "Interface configuration unknown"
        End Select
        strDeviceType = pEntry.str_type ' Device Name/Type
        ' IP Address String
        cstrIpAddress = "IP Address: " & IpAddrConvert(pEntry.SockAddr)
        cstrAdditionalDesc = pEntry.str_description ' Additional Description
        ' MacAddress
        cstrMacAddress = "MAC Address: "
        strSep = ""
        For idxVal = 0 To 5
            cstrMacAddress = cstrMacAddress & strSep & FmtHex(pEntry.mac(idxVal), 2)
            strSep = "-"
        Next
        ' Event1 Time
        temp_str = pEntry.str_ev1 & " "
        temp_str2 = FormatNetFinderTime(pEntry, 1)
        cstrEvent1 = temp_str + temp_str2
        ' Event2 Time
        temp_str = pEntry.str_ev2 & " "
        temp_str2 = FormatNetFinderTime(pEntry, 2)
        cstrEvent2 = temp_str + temp_str2
        cstrEntry = strDeviceType & vbCrLf
        cstrEntry = cstrEntry & cstrAlertLevel & vbCrLf
        cstrEntry = cstrEntry & cstrIpAddress & vbCrLf
        cstrEntry = cstrEntry & cstrAdditionalDesc & vbCrLf
        cstrEntry = cstrEntry & cstrMacAddress & vbCrLf
        cstrEntry = cstrEntry & cstrEvent1 & vbCrLf
        cstrEntry = cstrEntry & cstrEvent2 & vbCrLf
        EntryToStr = cstrEntry
    End Function
   
    Public Function EntryToStrRS232(ByRef pEntry As NETDISPLAY_ENTRY, ByRef strPort As String) As String
        Dim cstrAlertLevel As String = ""
        Dim strDeviceType As String = ""
        Dim cstrIpAddress As String = ""
        Dim cstrEvent1 As String = ""
        Dim cstrEvent2 As String = ""
        Dim cstrAdditionalDesc As String = ""
        Dim cstrMacAddress As String = ""
        Dim cstrEntry As String = ""
        Dim temp_str As String = ""
        Dim temp_str2 As String = ""
        Dim idxVal As Integer
        Dim strSep As String = ""
       
        cstrAlertLevel = "Connection: "
        cstrAlertLevel = cstrAlertLevel & strPort
        strDeviceType = pEntry.str_type ' Device Name/Type
        ' IP Address String
        cstrIpAddress = "IP Address: " & IpAddrConvert(pEntry.SockAddr)
        cstrAdditionalDesc = pEntry.str_description ' Additional Description
        ' MacAddress
        cstrMacAddress = "MAC Address: "
        strSep = ""
        For idxVal = 0 To 5
            cstrMacAddress = cstrMacAddress & strSep & FmtHex(pEntry.mac(idxVal), 2)
            strSep = "-"
        Next
        ' Event1 Time
        temp_str = pEntry.str_ev1 & " "
        temp_str2 = FormatNetFinderTime(pEntry, 1)
        cstrEvent1 = temp_str + temp_str2
        ' Event2 Time
        temp_str = pEntry.str_ev2 & " "
        temp_str2 = FormatNetFinderTime(pEntry, 2)
        cstrEvent2 = temp_str + temp_str2
       
        cstrEntry = strDeviceType & vbCrLf
        cstrEntry = cstrEntry & cstrAlertLevel & vbCrLf
        'cstrEntry = cstrEntry + cstrIpAddress & vbCrLf
        cstrEntry = cstrEntry & cstrAdditionalDesc & vbCrLf
        'cstrEntry = cstrEntry + cstrMacAddress & vbCrLf
        cstrEntry = cstrEntry + cstrEvent1 & vbCrLf
        'cstrEntry = cstrEntry + cstrEvent2 & vbCrLf
        EntryToStrRS232 = cstrEntry
    End Function
   
    Public Function EntryToStrUSB(ByRef pEntry As NETDISPLAY_ENTRY, ByRef strPort As String) As String
        Dim cstrAlertLevel As String = ""
        Dim strDeviceType As String = ""
        Dim cstrIpAddress As String = ""
        Dim cstrEvent1 As String = ""
        Dim cstrEvent2 As String = ""
        Dim cstrAdditionalDesc As String = ""
        Dim cstrMacAddress As String = ""
        Dim cstrEntry As String = ""
        Dim temp_str As String = ""
        Dim temp_str2 As String = ""
        Dim idxVal As Integer
        Dim strSep As String = ""
       
        cstrAlertLevel = "Connection: "
        cstrAlertLevel = cstrAlertLevel & strPort
        strDeviceType = pEntry.str_type ' Device Name/Type
        ' IP Address String
        cstrIpAddress = "IP Address: " & IpAddrConvert(pEntry.SockAddr)
        cstrAdditionalDesc = pEntry.str_description ' Additional Description
        ' MacAddress
        cstrMacAddress = "MAC Address: "
        strSep = ""
        For idxVal = 0 To 5
            cstrMacAddress = cstrMacAddress & strSep & FmtHex(pEntry.mac(idxVal), 2)
            strSep = "-"
        Next
        ' Event1 Time
        temp_str = pEntry.str_ev1 & " "
        temp_str2 = FormatNetFinderTime(pEntry, 1)
        cstrEvent1 = temp_str + temp_str2
        ' Event2 Time
        temp_str = pEntry.str_ev2 & " "
        temp_str2 = FormatNetFinderTime(pEntry, 2)
        cstrEvent2 = temp_str + temp_str2
       
        cstrEntry = strDeviceType & vbCrLf
        cstrEntry = cstrEntry & cstrAlertLevel & vbCrLf
        'cstrEntry = cstrEntry + cstrIpAddress & vbCrLf
        cstrEntry = cstrEntry & cstrAdditionalDesc & vbCrLf
        'cstrEntry = cstrEntry + cstrMacAddress & vbCrLf
        cstrEntry = cstrEntry + cstrEvent1 & vbCrLf
        'cstrEntry = cstrEntry + cstrEvent2 & vbCrLf
        EntryToStrUSB = cstrEntry
    End Function
   
    Private Function FormatNetFinderTime(ByRef pEntry As NETDISPLAY_ENTRY, ByRef NetFinderEvent As Byte) As String
        Dim event_days As Integer
        Dim event_hours As Byte
        Dim event_minutes As Byte
        Dim event_seconds As Byte
        Dim strTime(3) As String
        Dim idxValue As Integer
        Dim idxStart As Integer
        Dim strEventTime As String

        If (NetFinderEvent = 1) Then
            event_days = pEntry.event1_days
            event_hours = pEntry.event1_hours
            event_minutes = pEntry.event1_minutes
            event_seconds = pEntry.event1_seconds
        ElseIf (NetFinderEvent = 2) Then
            event_days = pEntry.event2_days
            event_hours = pEntry.event2_hours
            event_minutes = pEntry.event2_minutes
            event_seconds = pEntry.event2_seconds
        Else
            FormatNetFinderTime = ""
            Exit Function
        End If

        strTime(0) = CStr(event_days) & " days,"
        strTime(1) = CStr(event_hours) & " hours,"
        strTime(2) = CStr(event_minutes) & " minutes,"
        strTime(3) = CStr(event_seconds) & " seconds"

        idxStart = 0
        If (CBool(event_days)) Then
            idxStart = 0
        ElseIf (CBool(event_hours)) Then
            idxStart = 1
        ElseIf (CBool(event_minutes)) Then
            idxStart = 2
        Else
            idxStart = 3
        End If

        strEventTime = ""
        For idxValue = idxStart To 3
            strEventTime = strEventTime & strTime(idxValue)
        Next
        FormatNetFinderTime = strEventTime
    End Function

End Module
