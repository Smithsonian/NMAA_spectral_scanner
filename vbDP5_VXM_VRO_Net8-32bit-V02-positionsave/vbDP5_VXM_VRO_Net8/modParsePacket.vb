Option Strict Off
Option Explicit On
Module modParsePacket

    Public Const USB_DiagDataDelayMS As Integer = 2500

    Public ACK_Received As Boolean
    Public Timeout_flag As Boolean

    Public Structure DiagDataType
        <VBFixedArray(12)> Dim ADC_V() As Single
        <VBFixedArray(4)> Dim PC5_V() As Single
        Dim PC5_PRESENT As Boolean
        Dim PC5_SerialNumber As Integer
        Dim Firmware As Byte
        Dim FPGA As Byte
        Dim SRAMTestPass As Boolean
        Dim SRAMTestData As Integer
        Dim TempOffset As Short
        Dim strTempRaw As String
        Dim strTempCal As String
        Dim PC5Initialized As Boolean
        Dim PC5DCAL As Single
        Dim IsPosHV As Boolean
        Dim Is8_5VPreAmp As Boolean
        Dim Sup9VOn As Boolean
        Dim PreAmpON As Boolean
        Dim HVOn As Boolean
        Dim TECOn As Boolean
        <VBFixedArray(192)> Dim DiagData() As Byte

        Public Sub Initialize()
            ReDim ADC_V(12)
            ReDim PC5_V(4)
            ReDim DiagData(192)
        End Sub
    End Structure

    Public Enum DP5_DPP_TYPES
        dppDP5
        dppPX5
        dppDP5G
        dppMCA8000D
        dppTB5
        dppDP5X
    End Enum

    Public Enum CommType
        commRS232 = 0
        commUSB = 1
        commSockets = 2
        commNone = 3
    End Enum

    Public Const preqProcessNone As Integer = &H0
    Public Const preqProcessStatus As Integer = &H1
    Public Const preqProcessSpectrum As Integer = &H2
    Public Const preqRequestScopeData As Integer = &H4
    Public Const preqProcessScopeData As Integer = &H8
    Public Const preqProcessTextData As Integer = &H10
    Public Const preqProcessScopeDataOverFlow As Integer = &H20
    Public Const preqProcessInetSettings As Integer = &H40
    Public Const preqProcessDiagData As Integer = &H80
    Public Const preqProcessHwDesc As Integer = &H100
    Public Const preqProcessCfgRead As Integer = &H200
    Public Const preqProcessNetFindRead As Integer = &H400
    '...
    Public Const preqProcessSCAData As Integer = &H2000
    Public Const preqProcessAck As Integer = &H4000
    Public Const preqProcessError As Integer = &H8000

    Public Structure DppStateType
        Dim Interface_Renamed As CommType
        Dim strPort As String
        Dim ReqProcess As Integer
        Dim HaveScopeData As Boolean
        Dim ScopeAutoRearm As Boolean
    End Structure

    Public DppState As DppStateType

    Public Sub ParsePacketStatus(ByRef P() As Byte, ByRef PIN As Packet_In)
        Dim X As Short

        Dim CSum As Integer
        CSum = 0
        PIN.Initialize()
        If P(0) = SYNC1_ Then
            If P(1) = SYNC2_ Then
                If P(4) < 128 Then ' LEN MSB < 128, i.e. LEN < 32768
                    PIN.LEN_Renamed = (P(4) * 256) + P(5)
                    'PIN.LEN = TwoByteToLong(P(4), P(5))
                    PIN.PID1 = P(2)
                    PIN.PID2 = P(3)
                    For X = 0 To PIN.LEN_Renamed + 5 ' add up all the bytes except checksum
                        CSum = CSum + P(X)
                    Next X
                    CSum = CSum + 256 * CInt(P(PIN.LEN_Renamed + 6)) + CInt(P(PIN.LEN_Renamed + 7))
                    PIN.CheckSum = CSum
                    If (CSum And &HFFFF) = 0 Then
                        PIN.STATUS = 0 ' packet is OK
                        If PIN.LEN_Renamed > 0 Then
                            For X = 6 To 5 + PIN.LEN_Renamed
                                PIN.DATA(X - 6) = P(X)
                            Next X
                            ' call processing routines based on PID?
                        End If
                    Else
                        PIN.STATUS = modDP5_Protocol.PID2_ACK_TYPE.PID2_ACK_CHECKSUM_ERROR ' checksum error
                    End If
                Else
                    PIN.STATUS = modDP5_Protocol.PID2_ACK_TYPE.PID2_ACK_LEN_ERROR ' length error
                End If
            Else
                PIN.STATUS = modDP5_Protocol.PID2_ACK_TYPE.PID2_ACK_SYNC_ERROR ' sync error
            End If
        Else
            PIN.STATUS = modDP5_Protocol.PID2_ACK_TYPE.PID2_ACK_SYNC_ERROR ' sync error
        End If
        '    Packet_Received = True  ' flag that packet was received, regardless of errors

    End Sub

    Public Function PID2_TextToString(ByRef strPacketSource As String, ByRef PID2 As Byte) As String
        PID2_TextToString = ""
        Select Case PID2
            Case modDP5_Protocol.PID2_ACK_TYPE.PID2_ACK_OK ' ACK OK
                'PID2_TextToString = strPacketSource & ": OK" & vbTab
                PID2_TextToString = ""
                ACK_Received = True
            Case modDP5_Protocol.PID2_ACK_TYPE.PID2_ACK_SYNC_ERROR
                PID2_TextToString = strPacketSource & ": Sync Error" & vbTab
            Case modDP5_Protocol.PID2_ACK_TYPE.PID2_ACK_PID_ERROR
                PID2_TextToString = strPacketSource & ": PID Error" & vbTab
            Case modDP5_Protocol.PID2_ACK_TYPE.PID2_ACK_LEN_ERROR
                PID2_TextToString = strPacketSource & ": Length Error" & vbTab
            Case modDP5_Protocol.PID2_ACK_TYPE.PID2_ACK_CHECKSUM_ERROR
                PID2_TextToString = strPacketSource & ": Checksum Error" & vbTab
            Case modDP5_Protocol.PID2_ACK_TYPE.PID2_ACK_BAD_PARAM
                PID2_TextToString = strPacketSource & ": Bad Parameter" & vbTab
                '                            TextLog "ACK: BAD PARAM '" + T$ + "'" + vbCrLf
            Case modDP5_Protocol.PID2_ACK_TYPE.PID2_ACK_BAD_HEX_REC
                PID2_TextToString = strPacketSource & ": Bad HEX Record" & vbTab
            Case modDP5_Protocol.PID2_ACK_TYPE.PID2_ACK_UNRECOG
                PID2_TextToString = strPacketSource & ": Unrecognized Command" & vbTab
        End Select
    End Function

    Public Function ParsePacket(ByRef P() As Byte, ByRef PIN As Packet_In) As Integer
        ParsePacket = preqProcessNone
        ParsePacketStatus(P, PIN)
        If (PIN.STATUS = modDP5_Protocol.PID2_ACK_TYPE.PID2_ACK_OK) Then ' no errors
            If ((PIN.PID1 = modDP5_Protocol.PID1_TYPE.PID1_RCV_STATUS) And (PIN.PID2 = modDP5_Protocol.PID2_REQ_STATUS_TYPE.PID2_SEND_DP4_STYLE_STATUS)) Then ' DP4-style status
                ParsePacket = preqProcessStatus
            ElseIf ((PIN.PID1 = modDP5_Protocol.PID1_TYPE.PID1_RCV_SPECTRUM) And ((PIN.PID2 >= modDP5_Protocol.PID2_RCV_SPECTRUM_TYPE.RCVPT_256_CHANNEL_SPECTRUM) And (PIN.PID2 <= modDP5_Protocol.PID2_RCV_SPECTRUM_TYPE.RCVPT_8192_CHANNEL_SPECTRUM_STATUS))) Then  ' spectrum / spectrum+status
                ParsePacket = preqProcessSpectrum
                'ElseIf ((PIN.PID1 = PID1_RCV_SCOPE_MISC) And (PIN.PID2 = PID2_SEND_SCOPE_DATA)) Then'scope data packet
                '    ParsePacket = preqProcessScopeData
                'ElseIf ((PIN.PID1 = PID1_RCV_SCOPE_MISC) And (PIN.PID2 = PID2_SEND_512_BYTE_MISC_DATA)) Then 'text data packet
                '    ParsePacket = preqProcessTextData
            ElseIf ((PIN.PID1 = modDP5_Protocol.PID1_TYPE.PID1_RCV_SCOPE_MISC) And (PIN.PID2 = modDP5_Protocol.PID2_RCV_SCOPE_MISC_TYPE.RCVPT_SCOPE_DATA)) Then  'scope data packet
                'ScopeOverFlow = FALSE
                ParsePacket = preqProcessScopeData
            ElseIf ((PIN.PID1 = modDP5_Protocol.PID1_TYPE.PID1_RCV_SCOPE_MISC) And (PIN.PID2 = modDP5_Protocol.PID2_RCV_SCOPE_MISC_TYPE.RCVPT_512_BYTE_MISC_DATA)) Then  'text data packet
                ParsePacket = preqProcessTextData
            ElseIf ((PIN.PID1 = modDP5_Protocol.PID1_TYPE.PID1_RCV_SCOPE_MISC) And (PIN.PID2 = modDP5_Protocol.PID2_RCV_SCOPE_MISC_TYPE.RCVPT_SCOPE_DATA_WITH_OVERFLOW)) Then  'scope data with overflow packet
                'ScopeOverFlow = TRUE
                ParsePacket = preqProcessScopeData
                'ParsePacket = preqProcessScopeDataOverFlow
            ElseIf ((PIN.PID1 = modDP5_Protocol.PID1_TYPE.PID1_RCV_SCOPE_MISC) And (PIN.PID2 = modDP5_Protocol.PID2_RCV_SCOPE_MISC_TYPE.RCVPT_ETHERNET_SETTINGS)) Then  'ethernet settings packet
                'ParsePacket = preqProcessInetSettings
            ElseIf ((PIN.PID1 = modDP5_Protocol.PID1_TYPE.PID1_RCV_SCOPE_MISC) And (PIN.PID2 = modDP5_Protocol.PID2_RCV_SCOPE_MISC_TYPE.RCVPT_DIAGNOSTIC_DATA)) Then  'diagnostic data  packet
                ParsePacket = preqProcessDiagData
                'ElseIf ((PIN.PID1 = PID1_RCV_SCOPE_MISC) And (PIN.PID2 = RCVPT_HARDWARE_DESCRIPTION)) Then 'hardware description packet
                'ParsePacket = preqProcessHwDesc
            ElseIf ((PIN.PID1 = modDP5_Protocol.PID1_TYPE.PID1_RCV_SCOPE_MISC) And (PIN.PID2 = modDP5_Protocol.PID2_RCV_SCOPE_MISC_TYPE.RCVPT_CONFIG_READBACK)) Then
                ParsePacket = preqProcessCfgRead
            ElseIf ((PIN.PID1 = modDP5_Protocol.PID1_TYPE.PID1_RCV_SCOPE_MISC) And (PIN.PID2 = modDP5_Protocol.PID2_RCV_SCOPE_MISC_TYPE.RCVPT_NETFINDER_READBACK)) Then
                ParsePacket = preqProcessNetFindRead
                'MsgBox "NetFinder!"
            ElseIf (PIN.PID1 = modDP5_Protocol.PID1_TYPE.PID1_ACK) Then
                ParsePacket = preqProcessAck
            Else
                PIN.STATUS = modDP5_Protocol.PID2_ACK_TYPE.PID2_ACK_PID_ERROR ' unknown PID
                ParsePacket = preqProcessError
            End If
        Else
            ParsePacket = preqProcessError
        End If

        'PIN.PID1 = 0    ' overwrite the PIDs so packet can't be erroneously processed again
        'PIN.PID2 = 0

    End Function

    'convert a 4 byte long word to a double
    'lwStart - starting index of longword
    'buffer - byte buffer
    Private Function LongWordToDouble(ByVal lwStart As Long, ByVal Buffer() As Byte) As Double
        Dim dblVal As Double
        Dim idxByte As Long
        Dim ByteMask As Double
        dblVal = 0
        For idxByte = 0 To 3      ' build 4 bytes (lwStart-lwStart+3) into double
            ByteMask = 2 ^ (8 * idxByte)
            dblVal = dblVal + (Buffer((lwStart + idxByte)) * ByteMask)
        Next
        LongWordToDouble = dblVal
    End Function

    Public Sub Process_Status(ByRef STATUS As Stat)
        Dim bDMCA_LiveTime As Boolean
        Dim uiFwBuild As Long
        uiFwBuild = 0
        bDMCA_LiveTime = False
        STATUS.DEVICE_ID = STATUS.RAW(39)       '0=dp5,1=px5,2=dp5g,3=mca8000d,4=tb5,dp5x
        Select Case STATUS.DEVICE_ID
            Case 0
                STATUS.strDeviceID = "DP5"
            Case 1
                STATUS.strDeviceID = "PX5"
            Case 2
                STATUS.strDeviceID = "DP5G"
            Case 3
                STATUS.strDeviceID = "MCA8000D"
            Case 4
                STATUS.strDeviceID = "TB5"
            Case 5
                STATUS.strDeviceID = "DP5X"
            Case Else
                STATUS.strDeviceID = "DP5" 'default
        End Select

        STATUS.FastCount = LongWordToDouble(0, STATUS.RAW)
        STATUS.SlowCount = LongWordToDouble(4, STATUS.RAW)
        STATUS.GP_COUNTER = LongWordToDouble(8, STATUS.RAW)
        STATUS.AccumulationTime = STATUS.RAW(12) * 0.001 + (STATUS.RAW(13) + STATUS.RAW(14) * 256.0# + STATUS.RAW(15) * 65536.0#) * 0.1
        STATUS.RealTime = LongWordToDouble(20, STATUS.RAW) * 0.001

        STATUS.Firmware = STATUS.RAW(24)
        STATUS.FPGA = STATUS.RAW(25)

        If (STATUS.Firmware > &H65) Then
            STATUS.Build = STATUS.RAW(37) And &HF   ' Build # added in FW6.06
        Else
            STATUS.Build = 0
        End If

        'Firmware Version:  6.07  Build:  0 has LiveTime and PREL
        'DEVICE_ID 0=DP5,1=PX5,2=DP5G,3=MCA8000D,4=TB5,5=DP5X
        If (STATUS.DEVICE_ID = DP5_DPP_TYPES.dppMCA8000D) Then
            If (STATUS.Firmware >= &H67) Then
                bDMCA_LiveTime = True
            End If
        End If

        If (bDMCA_LiveTime) Then
            STATUS.LiveTime = LongWordToDouble(16, STATUS.RAW) * 0.001
        Else
            STATUS.LiveTime = 0
        End If

        If (STATUS.RAW(29) < 128) Then
            STATUS.SerialNumber = CLng(LongWordToDouble(26, STATUS.RAW))
        Else
            STATUS.SerialNumber = -1
        End If

        If (STATUS.RAW(30) < 128) Then        ' not negative
            STATUS.HV = CDbl(STATUS.RAW(31) + (STATUS.RAW(30) * 256.0#)) * 0.5    ' 0.5V/count
        Else
            STATUS.HV = CDbl((STATUS.RAW(31) + (STATUS.RAW(30) * 256.0#)) - 65536.0#) * 0.5       ' 0.5V/count
        End If

        STATUS.DET_TEMP = CSng(STATUS.RAW(33) + (STATUS.RAW(32) And 15) * 256) * 0.1 - 273.16 ' 0.1K/count
        STATUS.DP5_TEMP = STATUS.RAW(34) - ((STATUS.RAW(34) And 128) * 2)

        STATUS.PresetRtDone = ((STATUS.RAW(35) And 128) = 128)

        'BYTE:35 BIT:D6
        ' = Preset LiveTime Done for MCA8000D
        ' = FAST Thresh locked for other dpp devices
        STATUS.PresetLtDone = False
        STATUS.AFAST_LOCKED = False
        If (bDMCA_LiveTime) Then       ' test for MCA8000D
            STATUS.PresetLtDone = ((STATUS.RAW(35) And 64) = 64)
        Else
            STATUS.AFAST_LOCKED = ((STATUS.RAW(35) And 64) = 64)
        End If
        STATUS.MCA_EN = ((STATUS.RAW(35) And 32) = 32)
        STATUS.PresetCountDone = ((STATUS.RAW(35) And 16) = 16)
        STATUS.SCOPE_DR = ((STATUS.RAW(35) And 4) = 4)
        STATUS.DP5_CONFIGURED = ((STATUS.RAW(35) And 2) = 2)

        STATUS.AOFFSET_LOCKED = ((STATUS.RAW(36) And 128) = 0)  ' 0=locked, 1=searching
        STATUS.MCS_DONE = ((STATUS.RAW(36) And 64) = 64)

        STATUS.b80MHzMode = ((STATUS.RAW(36) And 2) = 2)
        STATUS.bFPGAAutoClock = ((STATUS.RAW(36) And 1) = 1)

        STATUS.PC5_PRESENT = ((STATUS.RAW(38) And 128) = 128)
        If (STATUS.PC5_PRESENT) Then
            STATUS.PC5_HV_POL = ((STATUS.RAW(38) And 64) = 64)
            STATUS.PC5_8_5V = ((STATUS.RAW(38) And 32) = 32)
        Else
            STATUS.PC5_HV_POL = False
            STATUS.PC5_8_5V = False
        End If

        If (STATUS.Firmware >= &H65) Then   ' reboot flag added FW6.05
            If ((STATUS.RAW(36) And 32) = 32) Then
                STATUS.ReBootFlag = True
            Else
                STATUS.ReBootFlag = False
            End If
        Else
            STATUS.ReBootFlag = False
        End If

        STATUS.TEC_Voltage = ((CDbl(STATUS.RAW(40) And 15) * 256.0#) + CDbl(STATUS.RAW(41))) / 758.5
        STATUS.DPP_ECO = STATUS.RAW(49)
        STATUS.bScintHas80MHzOption = False
        STATUS.DPP_options = (STATUS.RAW(42) And 15)
        STATUS.HPGe_HV_INH = False
        STATUS.HPGe_HV_INH_POL = False
        STATUS.AU34_2 = False
        STATUS.isAscInstalled = False
        STATUS.isDP5_RevDxGains = False
        If (STATUS.DEVICE_ID = DP5_DPP_TYPES.dppPX5) Then
            If (STATUS.DPP_options = PX5_OPTIONS.PX5_OPTION_HPGe_HVPS) Then
                STATUS.HPGe_HV_INH = ((STATUS.RAW(42) And 32) = 32)
                STATUS.HPGe_HV_INH_POL = ((STATUS.RAW(42) And 16) = 16)
                If (STATUS.DPP_ECO = 1) Then
                    STATUS.isAscInstalled = True
                    STATUS.AU34_2 = ((STATUS.RAW(42) And 64) = 64)
                End If
            End If
        ElseIf ((STATUS.DEVICE_ID = DP5_DPP_TYPES.dppDP5G) Or (STATUS.DEVICE_ID = DP5_DPP_TYPES.dppTB5)) Then
            If ((STATUS.DPP_ECO = 1) Or (STATUS.DPP_ECO = 2)) Then
                STATUS.bScintHas80MHzOption = True
            End If
        ElseIf (STATUS.DEVICE_ID = DP5_DPP_TYPES.dppDP5) Then
            uiFwBuild = STATUS.Firmware
            uiFwBuild = uiFwBuild * CLng(256)   '  << 8
            uiFwBuild = uiFwBuild + STATUS.Build
            ' uiFwBuild - firmware with build for comparison
            uiFwBuild = uiFwBuild And &HFFFF
            If (uiFwBuild >= &H686) Then
                ' 0xFF Value indicates old Analog Gain Count of 16
                ' Values < 0xFF indicate new gain count of 24 and new board rev
                ' "DP5 G3 Configuration P" will not be used (==&HFF)
                If (STATUS.DPP_ECO < &HFF) Then
                    STATUS.isDP5_RevDxGains = True
                End If
            End If
        End If

        '    MsgBox DP5_Dx_OptionFlags(STATUS.DPP_ECO)

    End Sub

    'DP5 Revision and Configuration from ECO Byte
    '   D7-D6: 0-4, = DP5 Rev D-G
    '   D5-D4: minor rev, 0-3 (i.e. Rev D0, D1 etc.)
    '   D3-D0: Configuration 0 = A, 1=B, 2=C... 15=P.
    Public Function DP5_Dx_OptionFlags(ByVal DP5_Dx_Options As Byte) As String
        Dim strRev As String
        strRev = "Reported Rev/Config: "
        strRev = strRev & Chr(((DP5_Dx_Options \ 64) And &H3) + Asc("D"))       'D7D6
        strRev = strRev & Chr((DP5_Dx_Options And &H30) \ 16 + Asc("0")) + "-"  'D5D4
        strRev = strRev & Chr((DP5_Dx_Options And 15) + Asc("A"))               'D3D0
        DP5_Dx_OptionFlags = strRev
    End Function



    Sub Process_Diagnostics(ByRef PIN As Packet_In, ByRef dd As DiagDataType)
        Dim idxVal As Integer
        Dim strVal As String
        Dim DP5_ADC_Gain(10) As Single ' convert each ADC count to engineering units - values calculated in FORM.LOAD
        Dim PC5_ADC_Gain(3) As Single
        Dim PX5_ADC_Gain(11) As Single

        DP5_ADC_Gain(0) = 1# / 0.00286 ' 2.86mV/C
        DP5_ADC_Gain(1) = 1# ' Vdd mon (out-of-scale)
        DP5_ADC_Gain(2) = (30.1 + 20#) / 20# ' PWR
        DP5_ADC_Gain(3) = (13# + 20#) / 20# ' 3.3V
        DP5_ADC_Gain(4) = (4.99 + 20#) / 20# ' 2.5V
        DP5_ADC_Gain(5) = 1# ' 1.2V
        DP5_ADC_Gain(6) = (35.7 + 20#) / 20# ' 5.5V
        DP5_ADC_Gain(7) = (35.7 + 75#) / 35.7 ' -5.5V (this one is tricky)
        DP5_ADC_Gain(8) = 1# ' AN_IN
        DP5_ADC_Gain(9) = 1# ' VREF_IN

        PC5_ADC_Gain(0) = 500# ' HV: 1500V/3V
        PC5_ADC_Gain(1) = 100# ' TEC: 300K/3V
        PC5_ADC_Gain(2) = (20# + 10#) / 10# ' +8.5/5V

        'PX5_ADC_Gain(0) = (30.1 + 20#) / 20#       ' PWR
        PX5_ADC_Gain(0) = (69.8 + 20#) / 20# ' 9V (was originally PWR)
        PX5_ADC_Gain(1) = (13# + 20#) / 20# ' 3.3V
        PX5_ADC_Gain(2) = (4.99 + 20#) / 20# ' 2.5V
        PX5_ADC_Gain(3) = 1# ' 1.2V
        PX5_ADC_Gain(4) = (30.1 + 20#) / 20# ' 5V
        PX5_ADC_Gain(5) = (10.7 + 75#) / 10.7 ' -5V (this one is tricky)
        PX5_ADC_Gain(6) = (64.9 + 20#) / 20# ' +PA
        PX5_ADC_Gain(7) = (10.7 + 75#) / 10.7 ' -PA
        PX5_ADC_Gain(8) = (16# + 20#) / 20# ' +TEC
        PX5_ADC_Gain(9) = 500# ' HV: 1500V/3V
        PX5_ADC_Gain(10) = 100# ' TEC: 300K/3V
        PX5_ADC_Gain(11) = 1# / 0.00286 ' 2.86mV/C

        dd.Firmware = PIN.DATA(0)
        dd.FPGA = PIN.DATA(1)
        strVal = FmtHex(PIN.DATA(2), 2) & FmtHex(PIN.DATA(3), 2) & FmtHex(PIN.DATA(4), 2)
        dd.SRAMTestData = CInt("&H" & strVal)
        dd.SRAMTestPass = CBool(dd.SRAMTestData = &HFFFFFF)
        dd.TempOffset = PIN.DATA(180) + 256 * CShort(PIN.DATA(180) > 127) ' 8-bit signed value
        If (STATUS.DEVICE_ID = DP5_DPP_TYPES.dppDP5) Then
            For idxVal = 0 To 9
                dd.ADC_V(idxVal) = ((CShort(PIN.DATA(5 + idxVal * 2) And 3) * 256) + PIN.DATA(6 + idxVal * 2)) * 2.44 / 1024 * DP5_ADC_Gain(idxVal) ' convert counts to engineering units (C or V)
            Next
            dd.ADC_V(7) = dd.ADC_V(7) + dd.ADC_V(6) * (1 - DP5_ADC_Gain(7)) ' -5.5V is a function of +5.5V
            dd.strTempRaw = String.Format(dd.ADC_V(0) - 271.3, "###.0C")
            dd.strTempCal = String.Format(dd.ADC_V(0) - 280 + dd.TempOffset, "###.0C")
        ElseIf (STATUS.DEVICE_ID = DP5_DPP_TYPES.dppPX5) Then
            For idxVal = 0 To 10
                dd.ADC_V(idxVal) = ((CShort(PIN.DATA(5 + idxVal * 2) And 15) * 256) + PIN.DATA(6 + idxVal * 2)) * 3 / 4096 * PX5_ADC_Gain(idxVal) ' convert counts to engineering units (C or V)
            Next
            dd.ADC_V(11) = ((CShort(PIN.DATA(5 + 11 * 2) And 3) * 256) + PIN.DATA(6 + 11 * 2)) * 3 / 1024 * PX5_ADC_Gain(11) ' convert counts to engineering units (C or V)
            dd.ADC_V(5) = dd.ADC_V(5) - (3 * PX5_ADC_Gain(5)) + 3 ' -5V uses +3VR
            dd.ADC_V(7) = dd.ADC_V(7) - (3 * PX5_ADC_Gain(7)) + 3 ' -PA uses +3VR
            dd.strTempRaw = String.Format(dd.ADC_V(11) - 271.3, "###.0C")
            dd.strTempCal = String.Format(dd.ADC_V(11) - 280 + dd.TempOffset, "###.0C")
        End If

        dd.PC5_PRESENT = False ' assume no PC5, then check to see if there are any non-zero bytes
        For idxVal = 25 To 38
            If PIN.DATA(idxVal) > 0 Then
                dd.PC5_PRESENT = True
                Exit For
            End If
        Next

        If dd.PC5_PRESENT Then
            For idxVal = 0 To 2
                dd.PC5_V(idxVal) = ((CShort(PIN.DATA(25 + idxVal * 2) And 15) * 256) + PIN.DATA(26 + idxVal * 2)) * 3 / 4096 * PC5_ADC_Gain(idxVal) ' convert counts to engineering units (C or V)
            Next
            If PIN.DATA(34) < 128 Then
                dd.PC5_SerialNumber = CInt(PIN.DATA(31)) + CInt(PIN.DATA(32)) * 256 + CInt(PIN.DATA(33)) * (256 ^ 2) + CInt(PIN.DATA(34)) * (256 ^ 3)
            Else
                dd.PC5_SerialNumber = -1 ' no PC5 S/N
            End If
            If (PIN.DATA(35) = 255) And (PIN.DATA(36) = 255) Then
                dd.PC5Initialized = False
                dd.PC5DCAL = 0
            Else
                dd.PC5Initialized = True
                dd.PC5DCAL = CSng((CSng(PIN.DATA(35)) * 256 + CSng(PIN.DATA(36))) * 3 / 4096)
            End If
            dd.IsPosHV = CBool(PIN.DATA(37) And 128)
            dd.Is8_5VPreAmp = CBool(PIN.DATA(37) And 64)
            dd.Sup9VOn = CBool(PIN.DATA(38) And 8)
            dd.PreAmpON = CBool(PIN.DATA(38) And 4)
            dd.HVOn = CBool(PIN.DATA(38) And 2)
            dd.TECOn = CBool(PIN.DATA(38) And 1)
        Else
            For idxVal = 0 To 2
                dd.PC5_V(idxVal) = 0
            Next
            dd.PC5_SerialNumber = -1 ' no PC5 S/N
            dd.PC5Initialized = False
            dd.PC5DCAL = 0
            dd.IsPosHV = False
            dd.Is8_5VPreAmp = False
            dd.Sup9VOn = False
            dd.PreAmpON = False
            dd.HVOn = False
            dd.TECOn = False
        End If
        For idxVal = 0 To 191
            dd.DiagData(idxVal) = PIN.DATA(idxVal + 39)
        Next idxVal
    End Sub

    Public Function DiagnosticsToString(ByVal dd As DiagDataType) As String
        Dim idxVal As Long
        Dim strDiag As String

        strDiag = "Firmware: " & VersionToStr(dd.Firmware) & vbCrLf
        strDiag = strDiag & "FPGA: " & VersionToStr(dd.FPGA) & vbCrLf
        strDiag = strDiag & "SRAM Test: "
        If dd.SRAMTestPass Then
            strDiag = strDiag & "PASS" & vbCrLf
        Else
            strDiag = strDiag & "ERROR @ 0x" & FmtHex(dd.SRAMTestData, 6) & vbCrLf
        End If

        If (STATUS.DEVICE_ID = DP5_DPP_TYPES.dppDP5) Then
            strDiag = strDiag & "DP5 Temp (raw): " & dd.strTempRaw & vbCrLf
            strDiag = strDiag & "DP5 Temp (cal'd): " & dd.strTempCal & vbCrLf
            strDiag = strDiag & "PWR: " & String.Format(dd.ADC_V(2), "#.##0V") & vbCrLf
            strDiag = strDiag & "3.3V: " & String.Format(dd.ADC_V(3), "#.##0V") & vbCrLf
            strDiag = strDiag & "2.5V: " & String.Format(dd.ADC_V(4), "#.##0V") & vbCrLf
            strDiag = strDiag & "1.2V: " & String.Format(dd.ADC_V(5), "#.##0V") & vbCrLf
            strDiag = strDiag & "+5.5V: " & String.Format(dd.ADC_V(6), "#.##0V") & vbCrLf
            strDiag = strDiag & "-5.5V: " & String.Format(dd.ADC_V(7), "#.##0V") & vbCrLf
            strDiag = strDiag & "AN_IN: " & String.Format(dd.ADC_V(8), "#.##0V") & vbCrLf
            strDiag = strDiag & "VREF_IN: " & String.Format(dd.ADC_V(9), "#.##0V") & vbCrLf

            strDiag = strDiag & vbCrLf
            If dd.PC5_PRESENT Then
                strDiag = strDiag & "PC5: Present" & vbCrLf
                strDiag = strDiag & "HV: " & String.Format(dd.PC5_V(0), "####V") & vbCrLf
                strDiag = strDiag & "Detector Temp: " & String.Format(dd.PC5_V(1), "###.#K") & vbCrLf
                strDiag = strDiag & "+8.5/5V: " & String.Format(dd.PC5_V(2), "#.##0V") & vbCrLf
                If (dd.PC5_SerialNumber > -1) Then
                    strDiag = strDiag & "PC5 S/N:" + Str(dd.PC5_SerialNumber) & vbCrLf
                Else
                    strDiag = strDiag & "PC5 S/N: none" & vbCrLf
                End If
                If (dd.PC5Initialized) Then
                    strDiag = strDiag & "PC5 DCAL: " & String.Format(dd.PC5DCAL, "#.##0V") & vbCrLf
                Else
                    strDiag = strDiag & "PC5 DCAL: Uninitialized" & vbCrLf
                End If
                strDiag = strDiag & "PC5 Flavor: "
                strDiag = strDiag & IsAorB(dd.IsPosHV, "+HV, ", "-HV, ")
                strDiag = strDiag & IsAorB(dd.Is8_5VPreAmp, "8.5V preamp", "5V preamp") & vbCrLf
                strDiag = strDiag & "9V: " & OnOffStr(dd.Sup9VOn) & vbCrLf
                strDiag = strDiag & "Preamp: " & OnOffStr(dd.PreAmpON) & vbCrLf
                strDiag = strDiag & "HV: " & OnOffStr(dd.HVOn) & vbCrLf
                strDiag = strDiag & "TEC: " & OnOffStr(dd.TECOn) & vbCrLf
            Else
                strDiag = strDiag & "PC5: Not Present" & vbCrLf
            End If

        ElseIf (STATUS.DEVICE_ID = DP5_DPP_TYPES.dppPX5) Then
            strDiag = strDiag & "PX5 Temp (raw): " & dd.strTempRaw & vbCrLf
            strDiag = strDiag & "PX5 Temp (cal'd): " & dd.strTempCal & vbCrLf
            'strDiag = strDiag & "PWR: " & Format(dd.ADC_V(0), "#.##0V") & vbCrLf
            strDiag = strDiag & "9V: " & String.Format(dd.ADC_V(0), "#.##0V") & vbCrLf
            strDiag = strDiag & "3.3V: " & String.Format(dd.ADC_V(1), "#.##0V") & vbCrLf
            strDiag = strDiag & "2.5V: " & String.Format(dd.ADC_V(2), "#.##0V") & vbCrLf
            strDiag = strDiag & "1.2V: " & String.Format(dd.ADC_V(3), "#.##0V") & vbCrLf
            strDiag = strDiag & "+5V: " & String.Format(dd.ADC_V(4), "#.##0V") & vbCrLf
            strDiag = strDiag & "-5V: " & String.Format(dd.ADC_V(5), "#.##0V") & vbCrLf
            strDiag = strDiag & "+PA: " & String.Format(dd.ADC_V(6), "#.##0V") & vbCrLf
            strDiag = strDiag & "-PA: " & String.Format(dd.ADC_V(7), "#.##0V") & vbCrLf
            strDiag = strDiag & "TEC: " & String.Format(dd.ADC_V(8), "#.##0V") & vbCrLf
            strDiag = strDiag & "ABS(HV): " & String.Format(dd.ADC_V(9), "###0.0V") & vbCrLf
            strDiag = strDiag & "DET_TEMP: " & String.Format(dd.ADC_V(10), "##0.0K") & vbCrLf
        End If

        strDiag = strDiag & vbCrLf & "Diagnostic Data" & vbCrLf
        strDiag = strDiag & "---------------" & vbCrLf
        For idxVal = 0 To 191
            If idxVal Mod 8 = 0 Then strDiag = strDiag & FmtHex(idxVal, 2) & ":"
            strDiag = strDiag & FmtHex(dd.DiagData(idxVal), 2) + " "
            If idxVal Mod 8 = 7 Then strDiag = strDiag & vbCrLf
        Next
        DiagnosticsToString = strDiag
    End Function

    Public Function FmtHex(ByVal DecNum As Integer, ByVal HexDig As Integer) As String
        Dim strHex As String
        Dim lHex As Integer

        strHex = Hex(DecNum)
        lHex = Len(strHex)
        If (lHex < HexDig) Then
            FmtHex = New String("0", HexDig - lHex) & strHex
        Else
            FmtHex = strHex
        End If
    End Function

    Public Function VersionToStr(ByRef bVersion As Byte) As String
        VersionToStr = Trim(Str(CShort(bVersion And &HF0S) / 16)) & "." & String.Format(bVersion And &HFS, "00")
    End Function

    Public Function OnOffStr(ByRef bOnOff As Boolean) As String
        If (bOnOff) Then
            OnOffStr = "ON"
        Else
            OnOffStr = "OFF"
        End If
    End Function

    Public Function IsAorB(ByRef bIsA As Boolean, ByRef strA As String, ByRef strB As String) As String
        If (bIsA) Then
            IsAorB = strA
        Else
            IsAorB = strB
        End If
    End Function
End Module
