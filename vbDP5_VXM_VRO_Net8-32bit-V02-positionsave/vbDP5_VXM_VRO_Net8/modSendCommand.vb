Option Strict Off
Option Explicit On
Module modSendCommand
    
    Public Const DP4_PX4_OLD_CFG_SIZE As Short = 64
    Public Whitespace As String

    Public Function DP5_CMD(ByRef Buffer() As Byte, ByRef XmtCmd As modDP5_Protocol.TRANSMIT_PACKET_TYPE) As Boolean
        Dim bCmdFound As Boolean
        Dim iFileNum As Short
        Dim D As String = ""
        Dim bt As Byte
        Dim idxData As Short
        Dim POUT As New Packet_Out
        Dim lLen As Integer
        Dim cstrCfg As String = ""

        POUT.Initialize()
        bCmdFound = True
        POUT.PID1 = 0
        POUT.PID2 = 0
        POUT.LEN_Renamed = 0
        POUT.DATA(0) = 0
        Select Case XmtCmd
            ''REQUEST_PACKETS_TO_DP5
            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_STATUS
                POUT.PID1 = modDP5_Protocol.PID1_TYPE.PID1_REQ_STATUS
                POUT.PID2 = modDP5_Protocol.PID2_REQ_STATUS_TYPE.PID2_SEND_DP4_STYLE_STATUS
                'Case XMTPT_SEND_SPECTRUM
                'Case XMTPT_SEND_CLEAR_SPECTRUM
            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_SPECTRUM_STATUS
                POUT.PID1 = modDP5_Protocol.PID1_TYPE.PID1_REQ_SPECTRUM
                POUT.PID2 = modDP5_Protocol.PID2_REQ_SPECTRUM_TYPE.PID2_SEND_SPECTRUM_STATUS ' send spectrum & status
            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_CLEAR_SPECTRUM_STATUS
                POUT.PID1 = modDP5_Protocol.PID1_TYPE.PID1_REQ_SPECTRUM
                POUT.PID2 = modDP5_Protocol.PID2_REQ_SPECTRUM_TYPE.PID2_SEND_CLEAR_SPECTRUM_STATUS ' send & clear spectrum & status
                'Case XMTPT_BUFFER_SPECTRUM
                'Case XMTPT_BUFFER_CLEAR_SPECTRUM
                'Case XMTPT_SEND_BUFFER
                'Case XMTPT_SEND_DP4_STYLE_STATUS
                'Case XMTPT_SEND_CONFIG
            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_SCOPE_DATA
                POUT.PID1 = modDP5_Protocol.PID1_TYPE.PID1_REQ_SCOPE_MISC
                POUT.PID2 = modDP5_Protocol.PID2_REQ_SCOPE_MISC_TYPE.PID2_SEND_SCOPE_DATA
            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_512_BYTE_MISC_DATA
                POUT.PID1 = modDP5_Protocol.PID1_TYPE.PID1_REQ_SCOPE_MISC
                POUT.PID2 = modDP5_Protocol.PID2_REQ_SCOPE_MISC_TYPE.PID2_SEND_512_BYTE_MISC_DATA ' request misc data
                'Case XMTPT_SEND_SCOPE_DATA_REARM
                'Case XMTPT_SEND_ETHERNET_SETTINGS
            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_DIAGNOSTIC_DATA
                POUT.PID1 = modDP5_Protocol.PID1_TYPE.PID1_REQ_SCOPE_MISC
                POUT.PID2 = modDP5_Protocol.PID2_REQ_SCOPE_MISC_TYPE.PID2_SEND_DIAGNOSTIC_DATA ' Request Diagnostic Packet
                POUT.LEN_Renamed = 0
            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_NETFINDER_PACKET
                POUT.PID1 = modDP5_Protocol.PID1_TYPE.PID1_REQ_SCOPE_MISC
                POUT.PID2 = modDP5_Protocol.PID2_REQ_SCOPE_MISC_TYPE.PID2_SEND_NETFINDER_READBACK ' Request NetFinder Packet
                POUT.LEN_Renamed = 0
                'Case XMTPT_SEND_HARDWARE_DESCRIPTION
                'Case XMTPT_SEND_SCA
                'Case XMTPT_LATCH_SEND_SCA
                'Case XMTPT_LATCH_CLEAR_SEND_SCA
                'Case XMTPT_SEND_ROI_OR_FIXED_BLOCK
            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_PX4_STYLE_CONFIG_PACKET
                iFileNum = FreeFile()
                FileOpen(iFileNum, My.Application.Info.DirectoryPath & "\RAW.CFG", OpenMode.Input)
                For bt = 0 To 63
                    Input(iFileNum, D)
                    POUT.DATA(bt) = Val("&H" & D)
                Next
                FileClose(iFileNum)
                POUT.PID1 = modDP5_Protocol.PID1_TYPE.PID1_REQ_CONFIG
                POUT.PID2 = modDP5_Protocol.PID2_REQ_CONFIG_TYPE.PID2_PX4_STYLE_CONFIG_PACKET ' PX4-style config packet
                POUT.LEN_Renamed = DP4_PX4_OLD_CFG_SIZE

            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_CONFIG_PACKET_TO_HW
                cstrCfg = ""
                cstrCfg = s.HwCfgDP5Out

                If (s.SendCoarseFineGain) Then
                    cstrCfg = RemoveCmd("GAIN", cstrCfg)
                Else
                    cstrCfg = RemoveCmd("GAIA", cstrCfg)
                    cstrCfg = RemoveCmd("GAIF", cstrCfg)
                End If

                cstrCfg = RemoveCmdByDeviceType(cstrCfg, STATUS.PC5_PRESENT, STATUS.DEVICE_ID, STATUS.isDP5_RevDxGains, STATUS.DPP_ECO)

                lLen = Len(cstrCfg)
                If (lLen > 0) Then
                    For idxData = 1 To lLen
                        POUT.DATA(idxData - 1) = Asc(Mid(cstrCfg, idxData, 1))
                    Next
                End If
                POUT.PID1 = modDP5_Protocol.PID1_TYPE.PID1_REQ_CONFIG
                POUT.PID2 = modDP5_Protocol.PID2_REQ_CONFIG_TYPE.PID2_TEXT_CONFIG_PACKET ' text config packet
                POUT.LEN_Renamed = lLen
            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_CONFIG_PACKET_EX ' bypass any filters
                cstrCfg = ""
                cstrCfg = s.HwCfgDP5Out

                lLen = Len(cstrCfg)
                If (lLen > 0) Then
                    For idxData = 1 To lLen
                        POUT.DATA(idxData - 1) = Asc(Mid(cstrCfg, idxData, 1))
                    Next
                End If
                POUT.PID1 = modDP5_Protocol.PID1_TYPE.PID1_REQ_CONFIG
                POUT.PID2 = modDP5_Protocol.PID2_REQ_CONFIG_TYPE.PID2_TEXT_CONFIG_PACKET ' text config packet
                POUT.LEN_Renamed = lLen
            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_READ_CONFIG_PACKET
                cstrCfg = ""
                cstrCfg = s.HwRdBkDP5Out
                lLen = Len(cstrCfg)
                If (lLen > 0) Then
                    For idxData = 1 To lLen
                        POUT.DATA(idxData - 1) = Asc(Mid(cstrCfg, idxData, 1))
                    Next
                End If
                POUT.PID1 = modDP5_Protocol.PID1_TYPE.PID1_REQ_CONFIG
                POUT.PID2 = modDP5_Protocol.PID2_REQ_CONFIG_TYPE.PID2_CONFIG_READBACK_PACKET ' read config packet
                POUT.LEN_Renamed = lLen
            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_FULL_READ_CONFIG_PACKET
                cstrCfg = ""
                cstrCfg = CreateFullReadBackCmd()
                lLen = Len(cstrCfg)
                If (lLen > 0) Then
                    For idxData = 1 To lLen
                        POUT.DATA(idxData - 1) = Asc(Mid(cstrCfg, idxData, 1))
                    Next
                End If
                POUT.PID1 = modDP5_Protocol.PID1_TYPE.PID1_REQ_CONFIG
                POUT.PID2 = modDP5_Protocol.PID2_REQ_CONFIG_TYPE.PID2_CONFIG_READBACK_PACKET ' read config packet
                POUT.LEN_Renamed = lLen
            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_ERASE_FPGA_IMAGE
                POUT.PID1 = modDP5_Protocol.PID1_TYPE.PID1_REQ_FPGA_UC
                POUT.PID2 = modDP5_Protocol.PID2_REQ_FPGA_UC_TYPE.PID2_ERASE_FPGA_IMAGE
                POUT.LEN_Renamed = 2
                POUT.DATA(0) = &H12S
                POUT.DATA(1) = &H34S
                'Case XMTPT_UPLOAD_PACKET_FPGA
                'Case XMTPT_REINITIALIZE_FPGA
                'Case XMTPT_ERASE_UC_IMAGE_0
            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_ERASE_UC_IMAGE_1
                POUT.PID1 = modDP5_Protocol.PID1_TYPE.PID1_REQ_FPGA_UC
                POUT.PID2 = modDP5_Protocol.PID2_REQ_FPGA_UC_TYPE.PID2_ERASE_UC_IMAGE_1 ' erase image #1 (sector 5)
                POUT.LEN_Renamed = 2
                POUT.DATA(0) = &H12S
                POUT.DATA(1) = &H34S

                'Case XMTPT_ERASE_UC_IMAGE_2
                'Case XMTPT_UPLOAD_PACKET_UC
                'Case XMTPT_SWITCH_TO_UC_IMAGE_0
            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SWITCH_TO_UC_IMAGE_1
                POUT.PID1 = modDP5_Protocol.PID1_TYPE.PID1_REQ_FPGA_UC
                POUT.PID2 = modDP5_Protocol.PID2_REQ_FPGA_UC_TYPE.PID2_SWITCH_TO_UC_IMAGE_1 ' switch to uC image #1
                POUT.LEN_Renamed = 2
                POUT.DATA(0) = &HA5S ' uC FLASH unlock keys
                POUT.DATA(1) = &HF1S
                'Case XMTPT_SWITCH_TO_UC_IMAGE_2
                'Case XMTPT_UC_FPGA_CHECKSUMS

                ''VENDOR_REQUESTS_TO_DP5
                'Case XMTPT_CLEAR_SPECTRUM_BUFFER_A
            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_ENABLE_MCA_MCS
                POUT.PID1 = modDP5_Protocol.PID1_TYPE.PID1_VENDOR_REQ
                POUT.PID2 = modDP5_Protocol.PID2_VENDOR_REQ_TYPE.PID2_ENABLE_MCA_MCS
                POUT.LEN_Renamed = 0
            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_DISABLE_MCA_MCS
                POUT.PID1 = modDP5_Protocol.PID1_TYPE.PID1_VENDOR_REQ
                POUT.PID2 = modDP5_Protocol.PID2_VENDOR_REQ_TYPE.PID2_DISABLE_MCA_MCS
                POUT.LEN_Renamed = 0
            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_ARM_DIGITAL_OSCILLOSCOPE
                POUT.PID1 = modDP5_Protocol.PID1_TYPE.PID1_VENDOR_REQ
                POUT.PID2 = modDP5_Protocol.PID2_VENDOR_REQ_TYPE.PID2_ARM_DIGITAL_OSCILLOSCOPE ' arm trigger
                'Case XMTPT_AUTOSET_INPUT_OFFSET
                'Case XMTPT_AUTOSET_FAST_THRESHOLD
                'Case XMTPT_READ_IO3_0
                'Case XMTPT_WRITE_IO3_0
                'Case XMTPT_SET_DCAL
                'Case XMTPT_SET_PZ_CORRECTION_UC_TEMP_CAL
                'Case XMTPT_SET_PZ_CORRECTION_UC_TEMP_CAL
                'Case XMTPT_SET_BOOT_FLAGS
                'Case XMTPT_SET_HV_DP4_EMULATION
                'Case XMTPT_SET_TEC_DP4_EMULATION
                'Case XMTPT_SET_INPUT_OFFSET_DP4_EMULATION
                'Case XMTPT_SET_ADC_CAL_GAIN_OFFSET
                'Case XMTPT_SET_SPECTRUM_OFFSET
                'Case XMTPT_REQ_SCOPE_DATA_MISC_DATA_SCA_PACKETS
                'Case XMTPT_SET_SERIAL_NUMBER
                'Case XMTPT_CLEAR_GP_COUNTER
                'Case XMTPT_SWITCH_SUPPLIES
                'Case XMTPT_SEND_TEST_PACKET

            Case Else
                bCmdFound = False
        End Select

        If bCmdFound Then
            If (Not POUT_Buffer(POUT, Buffer)) Then
                bCmdFound = False
            End If
        End If
        DP5_CMD = bCmdFound
    End Function
    
    Public Function DP5_CMD_Data(ByRef Buffer() As Byte, ByRef XmtCmd As modDP5_Protocol.TRANSMIT_PACKET_TYPE, ByRef DataOut As Object) As Boolean
        Dim bCmdFound As Boolean
        Dim idxData As Short
        Dim idxMiscData As Short
        Dim POUT As New Packet_Out
        Dim PktLen As Short

        POUT.Initialize()
        bCmdFound = False
        POUT.PID1 = 0
        POUT.PID2 = 0
        POUT.DATA(0) = 0
        POUT.LEN_Renamed = 0
        Select Case XmtCmd
            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_WRITE_512_BYTE_MISC_DATA
                POUT.PID1 = modDP5_Protocol.PID1_TYPE.PID1_VENDOR_REQ
                POUT.PID2 = modDP5_Protocol.PID2_VENDOR_REQ_TYPE.PID2_WRITE_512_BYTE_MISC_DATA ' write misc data
                POUT.LEN_Renamed = 512
                For idxMiscData = 0 To 511
                    POUT.DATA(idxMiscData) = Asc(Mid(DataOut + New String(Chr(0), 512), idxMiscData + 1, 1))
                Next idxMiscData
                bCmdFound = True
                If (Not POUT_Buffer(POUT, Buffer)) Then
                    bCmdFound = False
                End If
            Case modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_TEST_PACKET
                If (Not IsNothing(DataOut)) Then
                    PktLen = Len(DataOut)
                    If (PktLen >= 8) Then
                        For idxData = 0 To PktLen - 1
                            Buffer(idxData) = CInt(DataOut(idxData))
                        Next
                        bCmdFound = True
                    End If
                End If
            Case Else
                bCmdFound = False
        End Select
        DP5_CMD_Data = bCmdFound
    End Function

    Public Function POUT_Buffer(ByRef POUT As Packet_Out, ByRef Buffer() As Byte) As Boolean
        Dim idxBuffer As Short
        Dim CS As Int32
        Dim CsArr As Byte()
        On Error GoTo POUT_BufferErr

        If (UBound(Buffer) < POUT.LEN_Renamed + 7) Then ReDim Buffer(POUT.LEN_Renamed + 7)
        Buffer(0) = SYNC1_
        Buffer(1) = SYNC2_
        Buffer(2) = POUT.PID1
        Buffer(3) = POUT.PID2
        Buffer(4) = (POUT.LEN_Renamed And &HFF00S) \ 256
        Buffer(5) = POUT.LEN_Renamed And &HFFS

        CS = SYNC1_ + SYNC2_ + POUT.PID1 + POUT.PID2 + (POUT.LEN_Renamed And &HFF00S) \ 256 + CShort(POUT.LEN_Renamed And &HFFS)

        If POUT.LEN_Renamed > 0 Then
            For idxBuffer = 0 To POUT.LEN_Renamed - 1
                Buffer(idxBuffer + 6) = POUT.DATA(idxBuffer)
                CS = CS + POUT.DATA(idxBuffer)
            Next idxBuffer
        End If
        CS = CInt(CS Xor &HFFFFS) + 1
        CsArr = BitConverter.GetBytes(CS)
        Buffer(POUT.LEN_Renamed + 6) = CsArr(1)
        Buffer(POUT.LEN_Renamed + 7) = CsArr(0)
        POUT_Buffer = True
        Exit Function
POUT_BufferErr:
        POUT_Buffer = False
    End Function
End Module
