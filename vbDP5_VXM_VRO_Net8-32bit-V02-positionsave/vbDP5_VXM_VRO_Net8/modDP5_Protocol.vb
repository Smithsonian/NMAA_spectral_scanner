Option Strict Off
Option Explicit On
Imports System.Text
Imports Microsoft.Win32.SafeHandles
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports VB = Microsoft.VisualBasic
Module modDP5_Protocol

	Public strCommStatus As String

	<DllImport("winmm.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
	Public Function timeGetTime() As Integer   'timeGetTimeVB
	End Function
	Public Packet_Received As Boolean
	Public PacketIn(24648) As Byte ' largest possible IN packet
	Public PIN As Packet_In
	Public UDP_offset As Short
	Public USB_Default_Timeout As Boolean
	Public RS232ComPortTest As Boolean
	Public USBDeviceTest As Boolean

	Public Enum TRANSMIT_PACKET_TYPE
		'REQUEST_PACKETS_TO_DP5
		XMTPT_SEND_STATUS
		XMTPT_SEND_SPECTRUM
		XMTPT_SEND_CLEAR_SPECTRUM
		XMTPT_SEND_SPECTRUM_STATUS
		XMTPT_SEND_CLEAR_SPECTRUM_STATUS
		XMTPT_BUFFER_SPECTRUM
		XMTPT_BUFFER_CLEAR_SPECTRUM
		XMTPT_SEND_BUFFER
		XMTPT_SEND_DP4_STYLE_STATUS
		XMTPT_SEND_PX4_CONFIG
		XMTPT_SEND_SCOPE_DATA
		XMTPT_SEND_512_BYTE_MISC_DATA
		XMTPT_SEND_SCOPE_DATA_REARM
		XMTPT_SEND_ETHERNET_SETTINGS
		XMTPT_SEND_DIAGNOSTIC_DATA
		XMTPT_SEND_NETFINDER_PACKET
		XMTPT_SEND_HARDWARE_DESCRIPTION
		XMTPT_SEND_SCA
		XMTPT_LATCH_SEND_SCA
		XMTPT_LATCH_CLEAR_SEND_SCA
		XMTPT_SEND_ROI_OR_FIXED_BLOCK
		XMTPT_PX4_STYLE_CONFIG_PACKET
		XMTPT_SEND_CONFIG_PACKET_EX
		XMTPT_SEND_CONFIG_PACKET_TO_HW
		XMTPT_READ_CONFIG_PACKET
		XMTPT_FULL_READ_CONFIG_PACKET
		XMTPT_ERASE_FPGA_IMAGE
		XMTPT_UPLOAD_PACKET_FPGA
		XMTPT_REINITIALIZE_FPGA
		XMTPT_ERASE_UC_IMAGE_0
		XMTPT_ERASE_UC_IMAGE_1
		XMTPT_ERASE_UC_IMAGE_2
		XMTPT_UPLOAD_PACKET_UC
		XMTPT_SWITCH_TO_UC_IMAGE_0
		XMTPT_SWITCH_TO_UC_IMAGE_1
		XMTPT_SWITCH_TO_UC_IMAGE_2
		XMTPT_UC_FPGA_CHECKSUMS
		'VENDOR_REQUESTS_TO_DP5
		XMTPT_CLEAR_SPECTRUM_BUFFER_A
		XMTPT_ENABLE_MCA_MCS
		XMTPT_DISABLE_MCA_MCS
		XMTPT_ARM_DIGITAL_OSCILLOSCOPE
		XMTPT_AUTOSET_INPUT_OFFSET
		XMTPT_AUTOSET_FAST_THRESHOLD
		XMTPT_READ_IO3_0
		XMTPT_WRITE_IO3_0
		XMTPT_WRITE_512_BYTE_MISC_DATA
		XMTPT_SET_DCAL
		XMTPT_SET_PZ_CORRECTION_UC_TEMP_CAL_PZ
		XMTPT_SET_PZ_CORRECTION_UC_TEMP_CAL_UC
		XMTPT_SET_BOOT_FLAGS
		XMTPT_SET_HV_DP4_EMULATION
		XMTPT_SET_TEC_DP4_EMULATION
		XMTPT_SET_INPUT_OFFSET_DP4_EMULATION
		XMTPT_SET_ADC_CAL_GAIN_OFFSET
		XMTPT_SET_SPECTRUM_OFFSET
		XMTPT_REQ_SCOPE_DATA_MISC_DATA_SCA_PACKETS
		XMTPT_SET_SERIAL_NUMBER
		XMTPT_CLEAR_GP_COUNTER
		XMTPT_SWITCH_SUPPLIES
		XMTPT_SEND_TEST_PACKET
	End Enum 'TRANSMIT_PACKET_TYPE

	'Public Enum RECEIVE_PACKET_TYPE
	'	 'RESPONSE_PACKETS_FROM_DP5
	'	RCVPT_DP4_STYLE_STATUS
	'	RCVPT_256_CHANNEL_SPECTRUM
	'	RCVPT_256_CHANNEL_SPECTRUM_STATUS
	'	RCVPT_512_CHANNEL_SPECTRUM
	'	RCVPT_512_CHANNEL_SPECTRUM_STATUS
	'	RCVPT_1024_CHANNEL_SPECTRUM
	'	RCVPT_1024_CHANNEL_SPECTRUM_STATUS
	'	RCVPT_2048_CHANNEL_SPECTRUM
	'	RCVPT_2048_CHANNEL_SPECTRUM_STATUS
	'	RCVPT_4096_CHANNEL_SPECTRUM
	'	RCVPT_4096_CHANNEL_SPECTRUM_STATUS
	'	RCVPT_8192_CHANNEL_SPECTRUM
	'	RCVPT_8192_CHANNEL_SPECTRUM_STATUS
	'	RCVPT_SCOPE_DATA
	'	RCVPT_512_BYTE_MISC_DATA
	'	RCVPT_IO3_0_STATE
	'End Enum 'RECEIVE_PACKET_TYPE

	Public Enum PID1_TYPE
		PID1_REQ_STATUS = &H1S
		PID1_REQ_SPECTRUM = &H2S
		PID1_REQ_SCOPE_MISC = &H3S
		PID1_REQ_SCA = &H4S
		PID1_REQ_CONFIG = &H20S
		PID1_REQ_FPGA_UC = &H30S
		PID1_RCV_STATUS = &H80S
		PID1_RCV_SPECTRUM = &H81S
		PID1_RCV_SCOPE_MISC = &H82S
		PID1_RCV_SCA = &H83S
		PID1_VENDOR_REQ = &HF0S
		PID1_ACK = &HFFS
	End Enum 'PID1_TYPE

	Public Enum PID2_REQ_STATUS_TYPE
		PID2_SEND_DP4_STYLE_STATUS = &H1S
	End Enum 'PID2_REQ_STATUS_TYPE

	Public Enum PID2_REQ_SPECTRUM_TYPE
		' REQUEST_PACKETS_TO_DP5
		PID2_SEND_SPECTRUM = &H1S
		PID2_SEND_CLEAR_SPECTRUM = &H2S
		PID2_SEND_SPECTRUM_STATUS = &H3S
		PID2_SEND_CLEAR_SPECTRUM_STATUS = &H4S
		'PID2_BUFFER_SPECTRUM
		'PID2_BUFFER_CLEAR_SPECTRUM
		'PID2_SEND_BUFFER
		'PID2_SEND_CONFIG
	End Enum 'PID2_REQ_SPECTRUM_TYPE

	Public Enum PID2_REQ_SCOPE_MISC_TYPE
		PID2_SEND_SCOPE_DATA = &H1S
		PID2_SEND_512_BYTE_MISC_DATA = &H2S
		PID2_SEND_SCOPE_DATA_REARM = &H3S
		PID2_SEND_ETHERNET_SETTINGS = &H4S
		PID2_SEND_DIAGNOSTIC_DATA = &H5S
		PID2_SEND_HARDWARE_DESCRIPTION = &H6S
		PID2_SEND_NETFINDER_READBACK = &H7S
	End Enum 'PID2_REQ_SCOPE_MISC_TYPE

	Public Enum PID2_REQ_SCA_TYPE
		PID2_SEND_SCA = &H1S
		PID2_LATCH_SEND_SCA = &H2S
		PID2_LATCH_CLEAR_SEND_SCA = &H3S
	End Enum 'PID2_REQ_SCA_TYPE

	Public Enum PID2_REQ_CONFIG_TYPE
		PID2_PX4_STYLE_CONFIG_PACKET = &H1S
		PID2_TEXT_CONFIG_PACKET = &H2S
		PID2_CONFIG_READBACK_PACKET = &H3S
	End Enum 'PID2_SEND_CONFIG_TYPE

	Public Enum PID2_REQ_FPGA_UC_TYPE
		PID2_ERASE_FPGA_IMAGE = &H1S
		PID2_UPLOAD_PACKET_FPGA = &H2S
		PID2_REINITIALIZE_FPGA = &H3S
		PID2_ERASE_UC_IMAGE_0 = &H4S
		PID2_ERASE_UC_IMAGE_1 = &H5S
		PID2_ERASE_UC_IMAGE_2 = &H6S
		PID2_UPLOAD_PACKET_UC = &H7S
		PID2_SWITCH_TO_UC_IMAGE_0 = &H8S
		PID2_SWITCH_TO_UC_IMAGE_1 = &H9S
		PID2_SWITCH_TO_UC_IMAGE_2 = &HAS
		'PID2_UC_FPGA_CHECKSUMS
	End Enum 'PID2_REQ_FPGA_UC_TYPE

	Public Enum PID2_VENDOR_REQ_TYPE
		PID2_CLEAR_SPECTRUM_BUFFER_A = &H1S
		PID2_ENABLE_MCA_MCS = &H2S
		PID2_DISABLE_MCA_MCS = &H3S
		PID2_ARM_DIGITAL_OSCILLOSCOPE = &H4S
		PID2_AUTOSET_INPUT_OFFSET = &H5S
		PID2_AUTOSET_FAST_THRESHOLD = &H6S
		PID2_READ_IO3_0 = &H7S
		PID2_WRITE_IO3_0 = &H8S
		PID2_WRITE_512_BYTE_MISC_DATA = &H9S
		PID2_SET_DCAL = &HAS
		PID2_SET_PZ_CORRECTION_UC_TEMP_CAL_PZ = &HBS
		PID2_SET_PZ_CORRECTION_UC_TEMP_CAL_UC = &HCS
		PID2_SET_BOOT_FLAGS = &HDS
		'PID2_SET_HV_DP4_EMULATION
		'PID2_SET_TEC_DP4_EMULATION
		'PID2_SET_INPUT_OFFSET_DP4_EMULATION
		PID2_SET_ADC_CAL_GAIN_OFFSET = &HES
		'PID2_SET_SPECTRUM_OFFSET
		'PID2_REQ_SCOPE_DATA_MISC_DATA_SCA_PACKETS
		PID2_SET_SERIAL_NUMBER = &HFS
		PID2_CLEAR_GP_COUNTER = &H10S
		PID2_SET_ETHERNET_SETTINGS = &H11S
		'PID2_SWITCH_SUPPLIES
		PID2_ETHERNET_ALLOW_SHAREING = &H20S
		PID2_ETHERNET_NO_SHARING = &H21S
		PID2_ETHERNET_LOCK_IP = &H22S
	End Enum 'PID2_VENDOR_REQ_TYPE

	Public Enum PID2_RCV_STATUS_TYPE
		RCVPT_DP4_STYLE_STATUS = &H1S
	End Enum 'PID2_RCV_STATUS_TYPE

	Public Enum PID2_RCV_SPECTRUM_TYPE
		RCVPT_256_CHANNEL_SPECTRUM = &H1S
		RCVPT_256_CHANNEL_SPECTRUM_STATUS = &H2S
		RCVPT_512_CHANNEL_SPECTRUM = &H3S
		RCVPT_512_CHANNEL_SPECTRUM_STATUS = &H4S
		RCVPT_1024_CHANNEL_SPECTRUM = &H5S
		RCVPT_1024_CHANNEL_SPECTRUM_STATUS = &H6S
		RCVPT_2048_CHANNEL_SPECTRUM = &H7S
		RCVPT_2048_CHANNEL_SPECTRUM_STATUS = &H8S
		RCVPT_4096_CHANNEL_SPECTRUM = &H9S
		RCVPT_4096_CHANNEL_SPECTRUM_STATUS = &HAS
		RCVPT_8192_CHANNEL_SPECTRUM = &HBS
		RCVPT_8192_CHANNEL_SPECTRUM_STATUS = &HCS
	End Enum 'PID2_RCV_SPECTRUM_TYPE

	Public Enum PID2_RCV_SCOPE_MISC_TYPE
		RCVPT_SCOPE_DATA = &H1S
		RCVPT_512_BYTE_MISC_DATA = &H2S
		RCVPT_SCOPE_DATA_WITH_OVERFLOW = &H3S
		RCVPT_ETHERNET_SETTINGS = &H4S
		RCVPT_DIAGNOSTIC_DATA = &H5S
		RCVPT_HARDWARE_DESCRIPTION = &H6S
		RCVPT_CONFIG_READBACK = &H7S
		RCVPT_NETFINDER_READBACK = &H8S
	End Enum 'PID2_REQ_SCOPE_MISC_TYPE

	Public Enum PID2_RCV_SCA_TYPE
		RCVPT_SCA = &H1S
	End Enum 'PID2_REQ_SCA_TYPE

	Public Enum PID2_ACK_TYPE 'PID1_ACK, PID2_ACK_TYPE
		PID2_ACK_OK = &H0S
		PID2_ACK_SYNC_ERROR = &H1S
		PID2_ACK_PID_ERROR = &H2S
		PID2_ACK_LEN_ERROR = &H3S
		PID2_ACK_CHECKSUM_ERROR = &H4S
		PID2_ACK_BAD_PARAM = &H5S
		PID2_ACK_BAD_HEX_REC = &H6S
		PID2_ACK_UNRECOG = &H7S
		PID2_ACK_FPGA_ERROR = &H8S
		PID2_ACK_CP2201_NOT_FOUND = &H9S
		PID2_ACK_SCOPE_DATA_NOT_AVAIL = &HAS
		PID2_ACK_PC5_NOT_PRESENT = &HBS
		PID2_ACK_OK_ETHERNET_SHARE_REQ = &HCS
		PID2_ACK_ETHERNET_BUSY = &HDS
	End Enum 'PID2_ACK_TYPE

	'Public Const PID1_ACK = &HFF


	Public Const ACK As Short = 0
	'Public Const STATUS = 1
	'Public Const SPECTRUM = 2
	Public Const SYNC1_ As Short = &HF5S
	Public Const SYNC2_ As Short = &HFAS

	Public Const RS232 As Short = 0
	Public Const USB As Short = 1
	Public Const ETHERNET As Short = 2

	Public isDppFound As Boolean
	Public isNetFinderReady As Boolean
	Public strNetFinder As String

	Public Const ETHERNET_TIMEOUT As Short = 1000 ' default timeout of 1000mS
	Public Const INTERNET_TIMEOUT As Short = 3000 ' 3 seconds if DP5 is on different subnet than this PC
	Public Const RS232_TIMEOUT As Short = 2500 ' default timeout of 2500mS (8K spectrum ~2.2s)
	Public Const USB_TIMEOUT As Short = 500 ' default timeout of 500mS

	Public Const ETHERNET_KEEPALIVE As Short = 1500 ' send a keep-alive after 1.5s of interface inactivity
	Public Const Retries As Short = 2 ' total of 3 attempts

	Public OutStr As String
	Public InPacketType As Short
	'Public PacketIn(24648) As Byte ' largest possible IN packet
	Public RS232HeaderReceived As Boolean
	Public RS232PacketPtr As Short
	'Public CurrentInterface As Byte
	Public PlotColor As Integer
	Public Scope(2047) As Byte ' make this a structure, with scope info?
	Public ScopeOverFlow As Boolean
	Public MiscData(512) As Byte
	Public CommError As Boolean

	Public Structure Packet_In
		Dim PID1 As Byte
		Dim PID2 As Byte
		Dim LEN_Renamed As Short ' signed, but data payload always less than 32768
		Dim STATUS As Byte
		<VBFixedArray(32768)> Dim DATA() As Byte
		Dim CheckSum As Integer

		Public Sub Initialize()
			ReDim DATA(32768)
		End Sub
	End Structure

	Public BufferOUT(520) As Byte

	Public Structure Packet_Out
		Dim PID1 As Byte
		Dim PID2 As Byte
		Dim LEN_Renamed As Short                    ' signed, but data payload always less than 32768
		Dim EXPECTEDRESPONSE As Byte
		<VBFixedArray(514)> Dim DATA() As Byte
		Public Sub Initialize()
			ReDim DATA(514)
		End Sub
	End Structure

	Public Structure Spec
		Dim DATA() As Integer ' this keeps total of static data under 64K VB limit
		Dim Channels As Short
	End Structure

	Public Enum PX5_OPTIONS
		PX5_OPTION_NONE = 0
		PX5_OPTION_HPGe_HVPS = 1
		PX5_OPTION_TEST_TEK = 4
		PX5_OPTION_TEST_MOX = 5
		PX5_OPTION_TEST_AMP = 6
		PX5_OPTION_TEST_MODE_1 = 8
		PX5_OPTION_TEST_MODE_2 = 9
	End Enum

	Public Structure Stat
		<VBFixedArray(64)> Dim RAW() As Byte
		Dim SerialNumber As Long
		Dim FastCount As Double
		Dim SlowCount As Double
		Dim FPGA As Byte
		Dim Firmware As Byte
		Dim Build As Byte
		Dim AccumulationTime As Double
		Dim HV As Double
		Dim DET_TEMP As Double
		Dim DP5_TEMP As Double
		Dim PX4 As Boolean
		Dim AFAST_LOCKED As Boolean
		Dim MCA_EN As Boolean
		Dim PresetCountDone As Boolean
		Dim PresetRtDone As Boolean
		Dim PresetLtDone As Boolean
		Dim SUPPLIES_ON As Boolean
		Dim SCOPE_DR As Boolean
		Dim DP5_CONFIGURED As Boolean
		Dim GP_COUNTER As Double
		Dim AOFFSET_LOCKED As Boolean
		Dim MCS_DONE As Boolean
		Dim DCAL As Double
		Dim PZCORR As Byte
		Dim UC_TEMP_OFFSET As Byte
		Dim AN_IN As Double
		Dim VREF_IN As Double
		Dim PC5_SerialNumber As Integer
		Dim PC5_PRESENT As Boolean
		Dim PC5_HV_POL As Boolean
		Dim PC5_8_5V As Boolean
		Dim LiveTime As Double
		Dim RealTime As Double
		Dim BOOT_FLAG_LSB As Byte
		Dim BOOT_FLAG_MSB As Byte
		Dim BOOT_HV As Double
		Dim BOOT_TEC As Double
		Dim BOOT_INPOFFSET As Double
		Dim ADC_GAIN_CAL As Double
		Dim ADC_OFFSET_CAL As Byte
		Dim SPECTRUM_OFFSET As Integer

		Dim b80MHzMode As Boolean
		Dim bFPGAAutoClock As Boolean
		Dim DEVICE_ID As Byte          ' 0=DP5, 1=PX5, 2=DP5G, 3=MCA8000D, 4=TB5, 5=DP5X
		Dim strDeviceID As String
		Dim ReBootFlag As Boolean      ' FW6.05 and later
		Dim DPP_options As Byte      ' 1=HPGe HVPS
		Dim HPGe_HV_INH As Boolean
		Dim HPGe_HV_INH_POL As Boolean
		Dim TEC_Voltage As Double
		Dim DPP_ECO As Byte
		Dim AU34_2 As Boolean
		Dim isAscInstalled As Boolean
		Dim isAscEnabled As Boolean
		Dim bScintHas80MHzOption As Boolean
		Dim isDP5_RevDxGains As Boolean

		Public Sub Initialize()
			ReDim RAW(64)
		End Sub
	End Structure

	Public STATUS As Stat
	Public SPECTRUM As Spec

	Public Function timeGetTimeVB() As Double
		Dim lRawTime As Integer
		Dim dblTemp As Double

		lRawTime = timeGetTime() 'timeGetTimeVB
		If (lRawTime < 0) Then 'convert to unsigned integer
			dblTemp = (CDbl(lRawTime) + 1.0) + 4294967295.0
		Else
			dblTemp = lRawTime
		End If
		timeGetTimeVB = dblTemp
	End Function

	Public Function msTimeExpired(ByRef curStartTimeMS As Decimal, ByRef curDelayTimeMS As Decimal) As Boolean
		Dim curNewTime As Decimal
		curNewTime = timeGetTimeVB()
		If (curStartTimeMS < 0) Then curStartTimeMS = curStartTimeMS + CDec(2 ^ 32)
		If ((curNewTime < 0) Or (curNewTime < curStartTimeMS)) Then
			curNewTime = curNewTime + CDec(2 ^ 32)
		End If
		msTimeExpired = CBool(curNewTime >= (curStartTimeMS + curDelayTimeMS))
	End Function

	Public Function msTimeStart() As Decimal
		msTimeStart = timeGetTimeVB()
	End Function

	Public Function msTimeDiff(ByRef curStartTime As Decimal) As Decimal
		Dim curNewTime As Decimal
		curNewTime = timeGetTimeVB()
		If (curNewTime < curStartTime) Then
			curStartTime = curStartTime
			curNewTime = curNewTime + (2 ^ 32)
			msTimeDiff = curNewTime - curStartTime
		Else
			msTimeDiff = curNewTime - curStartTime
		End If
	End Function

	Public Function msDelay(ByRef MSec As Integer) As Boolean
		Dim curStart As Decimal
		Dim TimeExpired As Boolean
		TimeExpired = False
		curStart = msTimeStart()
		Do
			System.Windows.Forms.Application.DoEvents()
			TimeExpired = msTimeExpired(curStart, CDec(MSec))
		Loop While (Not TimeExpired)
		msDelay = TimeExpired
	End Function

	Public Function SendPacketUSB(ByRef DppWinUSB As WinUsbDevice, ByRef Buffer() As Byte, ByRef PacketIn() As Byte) As Boolean
		Dim bytesWritten, bytesRead As Integer
		Dim success As Boolean
		Dim strLog As String = ""
		Dim PLen As Integer

		If ((Buffer(2) = PID1_TYPE.PID1_REQ_SCOPE_MISC) And Buffer(3) = PID2_REQ_SCOPE_MISC_TYPE.PID2_SEND_DIAGNOSTIC_DATA) Then
			' Request Diagnostic Packet delay 'USB_DiagDataDelayMS
			Call SetPipePolicy1(DppWinUSB, CByte(DppWinUSB.bulkInPipe), CShort(modWinUSB.PIPE_POLICY_TYPE.PIPE_TRANSFER_TIMEOUT), USB_DiagDataDelayMS)
			USB_Default_Timeout = False ' flag that non-default timeout is in use
		Else
			If USB_Default_Timeout = False Then  ' change pipetimeout to default if it isn't already set that way
				Call SetPipePolicy1(DppWinUSB, CByte(DppWinUSB.bulkInPipe), CShort(modWinUSB.PIPE_POLICY_TYPE.PIPE_TRANSFER_TIMEOUT), USB_TIMEOUT)
				USB_Default_Timeout = True ' flag that default timeout is in use
			End If
		End If

		PLen = (Buffer(4) * 256) + Buffer(5) + 7

		ACK_Received = False
		Packet_Received = False ' a response packet is always expected
		PIN.STATUS = &HFFS ' packet invalid - will be overwritten soon by response/ACK packet

		success = WinUsb_WritePipe(DppWinUSB.winUsbHandle, CByte(DppWinUSB.bulkOutPipe), Buffer(0), PLen + 1, bytesWritten, 0)

		If (success) Then
			success = WinUsb_ReadPipe(DppWinUSB.winUsbHandle, CByte(DppWinUSB.bulkInPipe), PacketIn(0), 32767, bytesRead, 0)
			strLog = strLog & "USB Req sent: PID1=" & Str(Buffer(2)) & ", PID2=" & Str(Buffer(3)) & ": OK" & vbCrLf
			If (success) Then
				strLog = strLog & "USB Response: PID1=" & Str(PacketIn(2)) & ", PID2=" & Str(PacketIn(3)) & ": OK (read " & Str(bytesRead) & ")" & vbCrLf
				SendPacketUSB = True
			Else
				strLog = strLog & "USB Response: FAIL (read " & Str(bytesRead) & ")" & vbCrLf
				SendPacketUSB = False
			End If
		Else
			strLog = strLog & "USB Req sent: FAIL" & vbCrLf
			SendPacketUSB = False
		End If
	End Function

End Module
