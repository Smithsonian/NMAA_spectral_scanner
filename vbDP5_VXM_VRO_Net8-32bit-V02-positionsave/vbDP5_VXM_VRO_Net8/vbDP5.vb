Option Strict Off
Option Explicit On
'Imports VB = Microsoft.VisualBasic 'no longer needed
Imports System.IO
Imports System.IO.Ports
Imports System.Text.Encoding
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports System.Text

'Start the main window. Allows better closing options.
Module Program
    Sub Main()
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)     'add encoding
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New frmDP5())
    End Sub
End Module

'Main form/control window
Public Class frmDP5
    Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Integer
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)



    Sub ErrorOut()
        Dim errorCode As Integer = Marshal.GetLastWin32Error()
        Console.WriteLine("Last Win32 Error: " & errorCode)
    End Sub



    'Set up Connection names and Equip Specs
    'Connect COM Ports based on .ini file (need to manually edit file)
    Dim iniPath As String = My.Application.Info.DirectoryPath & "\" & "COMSettings-VXMVRODSD" & ".ini"
    Dim VXMPort As String = CSIniReader.ReadIniValue("COMPorts", "VXMPort", iniPath)
    Dim VROXPort As String = CSIniReader.ReadIniValue("COMPorts", "VROXPort", iniPath)
    Dim VROYZPort As String = CSIniReader.ReadIniValue("COMPorts", "VROYZPort", iniPath)
    Dim DSDPort As String = CSIniReader.ReadIniValue("COMPorts", "DSDPort", iniPath)

    'Can revert to direct coding port if necessary
    'Dim VXMPort As String = "COM3"
    'Dim VROXPort As String = "COM7"
    'Dim VROYZPort As String = "COM8"
    'Dim DSDPort As String = "COM4"

    Dim SerialPortVXM As New System.IO.Ports.SerialPort()
    Dim SerialPortVROX As New System.IO.Ports.SerialPort()
    Dim SerialPortVROYZ As New System.IO.Ports.SerialPort()
    Dim SerialPortDSD As New System.IO.Ports.SerialPort()

    'Labjack - not currently in use. If added, need to add dll ref
    'Dim u3 As U3 'LabJack device 'not used yet

    'VXM variables
    Dim VXMxmmpstepres As Double = 40 'number of steps to move one mm - x belt
    Dim VXMyzmmpstepres As Double = 157.5 'number of steps to move one mm - y,z screw
    Dim VXMxmmpstepresmod As Integer = 10 'starting modifier steps for x belt. Will add or subtract during batch run.
    'Dim VXMxextra As Integer = 15 'ADDING EXTRA STEPS To MOVE X-AXIS ONLY DUE TO BELT DRIVE
    Dim VXMxstart As Integer = 0
    Dim VXMxend As Integer = 0
    Dim EXstart As Double = 0
    Dim EXend As Double = 0
    Dim VXMystart As Integer = 0
    Dim VXMyend As Integer = 0
    Dim VXMxendrest As Integer = 0
    Dim VXMyendrest As Integer = 0
    Dim VXMBatchTop = False
    Dim VXMBatchBottom = False
    Dim VXMBatchAllowed = False
    Dim VXMxwidth As Integer = VXMxend - VXMxstart 'x width of area in motor steps
    Dim VXMywidth As Integer = VXMyend - VXMystart 'y width of area in motor steps
    Dim VXMxwidthmm As Double 'x width of area in mm
    Dim VXMywidthmm As Double 'y width of area in mm
    Dim VXMxstepsize As Integer 'how many x steps to get to defined mm spacing
    Dim VXMystepsize As Integer 'how many y steps to get to defined mm spacing
    Dim VXMxiternum As Integer 'how many jumps to get across area
    Dim VXMyiternum As Integer
    Dim YMotorDir As Integer = 1
    Dim YMotorPause As Integer = 20 'ms after each y-step
    Dim XMotorPause As Integer = 500 'ms after each x-step
    'Can this replace timer2?
    'Dim stopWatch As New Stopwatch() 'needed for timer routine
    'Dim milliSec As Long = 0

    '********************************
    'DP5 funcitons from Amptek, only updated to operate in VB .net 8
    '********************************

    Private Sub cmdSelectCommunications_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSelectCommunications.Click
        'Debug.WriteLine("cmdSelectCommunications clicked")
        Dim bStatusDone As Boolean

        Timer2.Enabled = False
        bStatusDone = msDelay(500)  'give time for timer to finish
        s.CurrentInterface = USB
        NetfinderUnit = 0

        InitAllEntries(DppEntries)
        frmDP5_Connect.ShowDialog()

        If (s.isDppConnected) Then
            lblSelectCommunications.Text = "USB - " & "WinUSB"
            SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_STATUS)
            bStatusDone = msDelay(500)
            s.Dp5CmdList = New Collection
            MakeDp5CmdList((s.Dp5CmdList))
            EnableDppCmdControls(True)
            cmdSelectCommunications.BackColor = Color.LightGreen
        Else
            lblSelectCommunications.Text = "Select Communications"
            EnableDppCmdControls(False)
        End If
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        If (chkDeltaMode.CheckState = System.Windows.Forms.CheckState.Checked) Then
            SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_CLEAR_SPECTRUM_STATUS)
        Else
            SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_SPECTRUM_STATUS)
        End If
    End Sub

    Private Sub cmdClearData_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClearData.Click
        SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_CLEAR_SPECTRUM_STATUS)
    End Sub

    Private Sub cmdReadMiscData_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_512_BYTE_MISC_DATA)
    End Sub

    Private Sub cmdShowConfiguration_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShowConfiguration.Click
        If (s.isDppConnected) Then
            s.DisplayCfg = True
            SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_FULL_READ_CONFIG_PACKET)
        End If
    End Sub

    Private Sub cmdStartAcquisition_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdStartAcquisition.Click
        SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_ENABLE_MCA_MCS)
        Timer2.Enabled = True
    End Sub


    Private Sub cmdSaveSpectrum_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSaveSpectrum.Click
        Dim curStart As Decimal 'start time from system time in milliseconds
        Dim TimeExpired As Boolean 'time expired flag
        Dim curElapsed As Decimal 'Elapsed time from start time in milliseconds
        Dim strCfg As String = ""
        Dim strStatus As String = ""
        strStatus = ""

        If (s.isDppConnected) Then
            s.HwCfgReady = False 'clear cfg ready flag
            s.HwCfgDP5 = "" 'clear config readback string
            s.cstrRawCfgIn = ""
            TimeExpired = False
            curStart = msTimeStart()
            s.SpectrumCfg = True
            SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_FULL_READ_CONFIG_PACKET)
            Do  'wait until s.HwCfgReady or timeout
                System.Windows.Forms.Application.DoEvents()
                curElapsed = msTimeDiff(curStart)
                TimeExpired = msTimeExpired(curStart, 1000) '1000 milliseconds max wait
            Loop Until (TimeExpired Or s.HwCfgReady)

            'read cfg in s.HwCfgDP5
            If ((Len(s.HwCfgDP5) > 0) And s.HwCfgReady) Then
                strCfg = s.HwCfgDP5
                If (Microsoft.VisualBasic.Strings.Right(strCfg, 2) = vbCrLf) Then
                    strCfg = strCfg.Substring(0, Len(strCfg) - 2)
                End If
            End If
        End If

        If (STATUS.SerialNumber > 0) Then
            strStatus = ShowStatusValueStrings(STATUS)
            If (Microsoft.VisualBasic.Strings.Right(strStatus, 2) = vbCrLf) Then
                strStatus = strStatus.Substring(0, Len(strStatus) - 2)
            End If
        End If
        SaveSpectrum(SPECTRUM, strStatus, strCfg, txtSpectrumTag.Text, txtSpectrumDescription.Text)
    End Sub

    Private Sub cmdSingleUpdate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSingleUpdate.Click
        Timer2.Enabled = False
        cmdSingleUpdate.Enabled = False
        If (chkDeltaMode.CheckState = System.Windows.Forms.CheckState.Checked) Then
            SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_CLEAR_SPECTRUM_STATUS)
        Else
            SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_SPECTRUM_STATUS)
        End If
        cmdSingleUpdate.Enabled = True
    End Sub

    Public Sub cmdUSBDeviceTest()
        SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_NETFINDER_PACKET)
    End Sub


    Private Sub cmdStopAcquisition_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdStopAcquisition.Click
        SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_DISABLE_MCA_MCS)
    End Sub

    Private Sub Form_Initialize_Renamed()
        On Error Resume Next
        LoadLibrary("shell32.dll")
        InitCommonControls()
        On Error GoTo 0
    End Sub

    Private Sub frmDP5_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        LoadApplicationSettings(Me)

        InitAllEntries(NfPktEntries)

        CurrentUSBDevice = 1
        NumUSBDevices = 0

        s.CfgReadBack = False
        s.DisplayCfg = False

        s.SaveCfg = False
        s.SpectrumCfg = False

        s.Dp5CmdList = New Collection
        MakeDp5CmdList((s.Dp5CmdList))
        s.HwCfgDP5 = ""

        EnableDppCmdControls(False)
        USBDeviceTest = False
        lblVersion.BorderStyle = System.Windows.Forms.FormBorderStyle.None
        Whitespace = Chr(0) & Chr(9) & Chr(10) & Chr(11) & Chr(12) & Chr(13) & Chr(32)
        s.CurrentInterface = USB
        PlotColor = &HFF
        USBDevicePathName = ""
        USBDeviceConnected = False
        USBDeviceNotificationHandle = 0
        pnlStatus.Text = "DP5 test application ready..."
        lblVersion.Text = "DP5 SDK vbDP5 v" & Microsoft.VisualBasic.Strings.Right(Str(My.Application.Info.Version.Major), 1) & "." & Microsoft.VisualBasic.Strings.Right(Str(My.Application.Info.Version.Minor), 2) & "; with VXM, VRO NMAA v 2025-07"
    End Sub


    'Public Sub LoadCOMPorts()
    '    'Load COM ports from ini file
    '    Dim iniPath As String = My.Application.Info.DirectoryPath & "\" & "COMSettings-VXMVRODSD" & ".ini"
    '    Dim ports As New List(Of String)
    '    'Debug.WriteLine("Loading ports")
    '    Dim VXMPort = CSIniReader.ReadIniValue("COMPorts", "VXMPort", iniPath)
    '    Dim VROXPort = CSIniReader.ReadIniValue("COMPorts", "VROXPort", iniPath)
    '    Dim VROYZPort = CSIniReader.ReadIniValue("COMPorts", "VROYZPort", iniPath)
    '    Dim DSDPort = CSIniReader.ReadIniValue("COMPorts", "DSDPort", iniPath)
    'End Sub





    Public Sub CloseProgram()
        MessageBox.Show("Close Program")
        Try 'try to allows for exception handling
            If SerialPortVXM.IsOpen Then 'This is the port in the Form Design
                SerialPortVXM.Close()
                Exit Try
            End If
        Catch ex As Exception 'exception handling
            MessageBox.Show(ex.Message, "Error closing", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        'Dim frm As System.Windows.Forms.Form
        'For Each frm In My.Application.OpenForms
        ' frm.Close()
        'Next frm
        For Each openForm As Form In Application.OpenForms.Cast(Of Form)().ToList()
            openForm.Close()
        Next

        Application.Exit()
    End Sub

    Private Sub UpdateStatusDisplay(ByRef STATUS As Stat)
        Dim m As Single
        Dim H As Single
        Dim s As Single
        Dim strStatus As String = ""
        Dim strParam As String = ""
        Dim strValue As String = ""

        strParam = strParam & "Device Type: " & vbCrLf
        strValue = STATUS.strDeviceID
        strStatus = strStatus & strValue & vbCrLf

        strParam = strParam & "Serial Number: " & vbCrLf
        strValue = CStr(STATUS.SerialNumber)
        strStatus = strStatus & strValue & vbCrLf

        strParam = strParam & "Firmware: " & vbCrLf
        strValue = "v" & String.Format(Fix(STATUS.Firmware / 16), "0") & "." & String.Format(STATUS.Firmware And 15, "00")
        strStatus = strStatus & strValue & vbCrLf

        strParam = strParam & "FPGA: " & vbCrLf
        strValue = "v" & String.Format(Fix(STATUS.FPGA / 16), "0") & "." & String.Format(STATUS.FPGA And 15, "00")
        strStatus = strStatus & strValue & vbCrLf

        strParam = strParam & "Fast Count: " & vbCrLf
        strValue = String.Format(STATUS.FastCount, "#,###,###,##0")
        strStatus = strStatus & strValue & vbCrLf

        strParam = strParam & "SlowCount: " & vbCrLf
        strValue = String.Format(STATUS.SlowCount, "#,###,###,##0")
        strStatus = strStatus & strValue & vbCrLf

        strParam = strParam & "Accumulation Time: " & vbCrLf
        If STATUS.AccumulationTime < 1000 Then
            strValue = String.Format(STATUS.AccumulationTime, "##0.000s")
        Else
            H = Fix(STATUS.AccumulationTime / 3600)
            m = Fix((STATUS.AccumulationTime - (H * 3600)) / 60)
            s = STATUS.AccumulationTime - H * 3600 - m * 60
            strValue = String.Format(H, "###0h ") & String.Format(m, "#0m ") & String.Format(s, "#0.0s")
        End If
        strStatus = strStatus & strValue & vbCrLf
        strParam = strParam & "PC5 Serial Number: " & vbCrLf
        If STATUS.PC5_PRESENT Then
            strValue = Str(STATUS.PC5_SerialNumber)
        Else
            strValue = "N/A"
        End If
        strStatus = strStatus & strValue & vbCrLf

        If (STATUS.DEVICE_ID <> modParsePacket.DP5_DPP_TYPES.dppDP5G) Then
            strParam = strParam & "Detector HV: " & vbCrLf
            strValue = String.Format(STATUS.HV, "###0V") ' round to nearest volt
            strStatus = strStatus & strValue & vbCrLf

            strParam = strParam & "Detector Temp: " & vbCrLf
            'strValue = String.Format(STATUS.DET_TEMP, "##0°C") ' round to nearest degree
            strValue = String.Format("{0:0.00}°C", STATUS.DET_TEMP) 'round hundreths
            strStatus = strStatus & strValue & vbCrLf
        End If

        strParam = strParam & "DP5 Temp: " & vbCrLf
        strValue = Str(STATUS.DP5_TEMP) & "°C"
        strStatus = strStatus & strValue & vbCrLf

        strParam = strParam & "G.P. Counter: " & vbCrLf
        strValue = Str(STATUS.GP_COUNTER)
        strStatus = strStatus & strValue & vbCrLf

        If (STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppDP5) Then
            strParam = strParam & "PC5 Present: " & vbCrLf
            strValue = Str(STATUS.PC5_PRESENT)
            strStatus = strStatus & strValue & vbCrLf

            If STATUS.PC5_PRESENT Then
                strParam = strParam & "PC5 HV Polarity: " & vbCrLf
                If (STATUS.PC5_HV_POL) Then
                    strValue = "Positive"
                Else
                    strValue = "Negative"
                End If
                strStatus = strStatus & strValue & vbCrLf

                strParam = strParam & "PC5 Preamp: " & vbCrLf
                If (STATUS.PC5_8_5V) Then
                    strValue = "+/-8.5V"
                Else
                    strValue = "+/-5V"
                End If
                strStatus = strStatus & strValue & vbCrLf
            End If

            strParam = strParam & "DP5 Temp: " & vbCrLf
            strValue = Str(STATUS.DP5_TEMP) & "°C"
            strStatus = strStatus & strValue & vbCrLf
        ElseIf (STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppDP5G) Then
            strParam = strParam & "PC5G Present: " & vbCrLf
            strValue = Str(STATUS.PC5_PRESENT)
            strStatus = strStatus & strValue & vbCrLf
        ElseIf (STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppPX5) Then
            'STATUS.DPP_options = 1
            'STATUS.HPGe_HV_INH = True
            'STATUS.HPGe_HV_INH_POL = True
            If STATUS.DPP_options > 0 Then
                '===============PX5 Options==================
                strParam = strParam & "PX5 Options: " & vbCrLf
                If (STATUS.DPP_options = 1) Then
                    strValue = "HPGe HVPS"
                Else
                    strValue = "Unknown"
                End If
                strStatus = strStatus & strValue & vbCrLf

                '===============HPGe HVPS HV Status==================
                strParam = strParam & "HPGe HV: " & vbCrLf
                If STATUS.HPGe_HV_INH Then
                    strValue = "not inhibited"
                Else
                    strValue = "inhibited"
                End If
                strStatus = strStatus & strValue & vbCrLf

                '===============HPGe HVPS Inhibit Status==================
                strParam = strParam & "INH Polarity: " & vbCrLf
                If STATUS.HPGe_HV_INH_POL Then
                    strValue = "high"
                Else
                    strValue = "low"
                End If
                strStatus = strStatus & strValue & vbCrLf
            Else
                strParam = strParam & "PX5 Options: " & vbCrLf
                strValue = "No Options Installed"
                strStatus = strStatus & strValue & vbCrLf
            End If
        End If
        _lblStatusDisplay_0.Text = strParam
        _lblStatusDisplay_1.Text = strStatus
    End Sub


    Public Sub RemCallParsePacket(ByRef PacketIn() As Byte)
        DppState.ReqProcess = ParsePacket(PacketIn, PIN)
        ParsePacketEx(PIN, DppState)
    End Sub

    Private Sub ParsePacketEx(ByRef PIN As Packet_In, ByRef DppState As DppStateType)
        Select Case DppState.ReqProcess
            Case preqProcessStatus
                ProcessStatusEx(PIN, DppState)
            Case preqProcessSpectrum
                ProcessSpectrumEx(PIN, DppState)
                'Case preqProcessScopeData
                'ProcessScopeDataEx(PIN, DppState)
                'Case preqProcessTextData
                'ProcessTextDataEx(PIN, DppState)
                '====================================================================
                '        Case preqProcessScopeDataOverFlow
                '            ProcessScopeDataEx PIN, DppState
                '            'ProcessScopeDataOverFlowEx PIN, DppState
                '        Case preqProcessInetSettings
                '            ProcessInetSettingsEx PIN, DppState
            Case preqProcessDiagData
                ProcessDiagDataEx(PIN, DppState)
                '        Case preqProcessHwDesc
                '            ProcessHwDescEx PIN, DppState
            Case preqProcessCfgRead
                ProcessCfgReadEx(PIN, DppState)
            Case preqProcessNetFindRead
                ProcessNetFindReadEx(PIN, DppState)
                '        Case preqProcessSCAData
                '            ProcessSCADataEx PIN, DppState
                '====================================================================
            Case preqProcessAck
                ProcessAck(PIN.PID2)
            Case preqProcessError
                DisplayError(PIN, DppState)
            Case Else
                'do nothing
        End Select
    End Sub

    Private Sub ProcessNetFindReadEx(ByRef PIN As Packet_In, ByRef DppState As DppStateType)
        Dim Buffer(4095) As Byte '0-4095, 4096 bytes
        Dim strDisplay As String = ""

        Dim idxNfInfo As Short
        For idxNfInfo = 0 To (PIN.LEN_Renamed - 1)
            Buffer(idxNfInfo) = PIN.DATA(idxNfInfo)
        Next
        InitEntry(NfPktEntries(DppState.Interface_Renamed))
        AddEntry(NfPktEntries(DppState.Interface_Renamed), Buffer, 0)
        If (USBDeviceTest) Then
            USBDeviceTest = False
            strDisplay = EntryToStrUSB(NfPktEntries(DppState.Interface_Renamed), "USB Device " & CurrentUSBDevice & " of " & NumUSBDevices)

            'add check for vb.net 8 to see if text box exists
            Dim isVisible As Boolean = False
            If frmDP5_Connect IsNot Nothing AndAlso Not frmDP5_Connect.IsDisposed Then

                If frmDP5_Connect.InvokeRequired Then
                    frmDP5_Connect.Invoke(Sub()
                                              If frmDP5_Connect.Visible Then
                                                  frmDP5_Connect.rtbStatus.Text = strDisplay
                                                  frmDP5_Connect.lblDppFound.Visible = True
                                                  pnlStatus.Text = "USB NetFind Found" & vbCrLf
                                              End If
                                          End Sub)
                Else
                    If frmDP5_Connect.Visible Then
                        frmDP5_Connect.rtbStatus.Text = strDisplay
                        frmDP5_Connect.lblDppFound.Visible = True
                        pnlStatus.Text = "USB NetFind Found" & vbCrLf
                    End If
                End If
            End If


            '            frmDP5_Connect.rtbStatus.Text = strDisplay
            isDppFound = True
            'frmDP5_Connect.lblDppFound.Visible = True
            pnlStatus.Text = "USB NetFind Found" & vbCrLf
        End If

    End Sub

    Private Sub ProcessCfgReadEx(ByRef PIN As Packet_In, ByRef DppState As DppStateType)
        Dim cstrRawCfgIn As String
        Dim cstrCmdD As String
        Dim strCfg As String
        Dim cstrDisplayCfgOut As String
        Dim idxCfg As Integer
        Dim varCmd As Object
        Dim strHwCfgDP5 As String

        cstrRawCfgIn = ""
        strHwCfgDP5 = ""
        ' =============================
        ' === Create Raw Configuration Buffer From Hardware ===
        For idxCfg = 0 To PIN.LEN_Renamed - 1
            cstrRawCfgIn = cstrRawCfgIn & Chr(PIN.DATA(idxCfg))
            strHwCfgDP5 = strHwCfgDP5 & Chr(PIN.DATA(idxCfg))
            If (PIN.DATA(idxCfg) = Asc(";")) Then
                cstrRawCfgIn = cstrRawCfgIn & vbCrLf
            End If
        Next
        If (s.DisplayCfg) Then
            s.DisplayCfg = False
            strCfg = cstrRawCfgIn
            'MsgBox(Len(strHwCfgDP5))
            cstrDisplayCfgOut = cstrRawCfgIn
            For Each varCmd In s.Dp5CmdList
                cstrCmdD = CStr(varCmd)
                If (Len(cstrCmdD) > 0) Then
                    'cstrDisplayCfgOut = ReplaceCmdDesc(cstrCmdD, cstrDisplayCfgOut)
                    cstrDisplayCfgOut = AppendCmdDesc(cstrCmdD, cstrDisplayCfgOut)
                End If
            Next varCmd
            frmDppConfigDisplay.m_strMessage = cstrDisplayCfgOut
            'frmDppConfigDisplay.m_strDelimiter = ";"
            frmDppConfigDisplay.m_strDelimiter = vbCrLf
            frmDppConfigDisplay.m_strTitle = "DPP Configuration"
            frmDppConfigDisplay.ShowDialog()
        ElseIf (s.CfgReadBack) Then
            s.CfgReadBack = False
            s.HwCfgDP5 = strHwCfgDP5
            s.cstrRawCfgIn = cstrRawCfgIn
            s.HwCfgReady = True
        ElseIf (s.SpectrumCfg) Then
            s.SpectrumCfg = False
            cstrDisplayCfgOut = cstrRawCfgIn
            For Each varCmd In s.Dp5CmdList
                cstrCmdD = CStr(varCmd)
                If (Len(cstrCmdD) > 0) Then
                    cstrDisplayCfgOut = AppendCmdDesc(cstrCmdD, cstrDisplayCfgOut)
                End If
            Next varCmd
            s.HwCfgDP5 = cstrDisplayCfgOut
            s.cstrRawCfgIn = cstrRawCfgIn
            s.HwCfgReady = True
        End If

        s.DisplayCfg = False
        s.CfgReadBack = False
        s.SpectrumCfg = False
    End Sub

    Private Sub DisplayError(ByRef PIN As Packet_In, ByRef DppState As DppStateType)
        pnlStatus.Text = PID2_TextToString("Received packet", CByte(PIN.STATUS))
        ' bad PID, assigned by ParsePacket
        If (PIN.STATUS = modDP5_Protocol.PID2_ACK_TYPE.PID2_ACK_PID_ERROR) Then
            pnlStatus.Text = "Received packet: PID1=0x" & FmtHex(PIN.PID1, 2) & ", PID2=0x" & FmtHex(PIN.PID2, 2) & ", LEN=" & Str(PIN.LEN_Renamed) & vbCrLf
        End If
    End Sub

    Private Sub ProcessAck(ByRef PID2 As Byte)
        pnlStatus.Text = PID2_TextToString("ACK", PID2)
    End Sub

    Private Sub ProcessStatusEx(ByRef PIN As Packet_In, ByRef DppState As DppStateType)
        Dim X As Short
        STATUS.Initialize()
        For X = 0 To 63
            STATUS.RAW(X) = PIN.DATA(X)
        Next X
        Call Process_Status(STATUS)
        RequestScopeData(STATUS.SCOPE_DR)
        UpdateStatusDisplay(STATUS)
        pnlStatus.Text = "Status Received" & vbCrLf
    End Sub

    Private Function ShowStatusValueStrings(ByRef m_DP5_Status As Stat) As String
        Dim strConfig As String
        Dim strTemp As String

        strConfig = "Device Type: " & m_DP5_Status.strDeviceID & vbCrLf
        strTemp = "Serial Number: " & CStr(m_DP5_Status.SerialNumber) & vbCrLf 'SerialNumber
        strConfig = strConfig & strTemp
        strTemp = "Firmware: " & VersionToStr(m_DP5_Status.Firmware) & vbCrLf
        strConfig = strConfig & strTemp
        If (m_DP5_Status.Firmware > &H65) Then
            strTemp = "Build: " & VersionToStr(m_DP5_Status.Build) & vbCrLf
            strConfig = strConfig & strTemp
        End If
        strTemp = "FPGA: " & VersionToStr(m_DP5_Status.FPGA) & vbCrLf
        strConfig = strConfig & strTemp
        If (m_DP5_Status.DEVICE_ID <> DP5_DPP_TYPES.dppMCA8000D) Then
            strTemp = "Fast Count: " & CStr(CDbl(m_DP5_Status.FastCount)) & vbCrLf 'FastCount
            strConfig = strConfig & strTemp
        End If
        strTemp = "Slow Count: " & CStr(CDbl(m_DP5_Status.SlowCount)) & vbCrLf 'SlowCount
        strConfig = strConfig & strTemp
        strTemp = "GP Count: " & CStr(CDbl(m_DP5_Status.GP_COUNTER)) & vbCrLf 'GP Count
        strConfig = strConfig & strTemp
        If (m_DP5_Status.DEVICE_ID <> DP5_DPP_TYPES.dppMCA8000D) Then
            strTemp = "Accumulation Time: " & CStr(m_DP5_Status.AccumulationTime) & vbCrLf 'AccumulationTime
            strConfig = strConfig & strTemp
        End If
        strTemp = "Real Time: " & CStr(m_DP5_Status.RealTime) & vbCrLf       'RealTime
        strConfig = strConfig & strTemp

        If (m_DP5_Status.DEVICE_ID <> DP5_DPP_TYPES.dppMCA8000D) Then
            strTemp = "Live Time: " & CStr(m_DP5_Status.LiveTime) & vbCrLf   'LiveTime
            strConfig = strConfig & strTemp
        End If

        If ((m_DP5_Status.DEVICE_ID <> DP5_DPP_TYPES.dppDP5G) And (m_DP5_Status.DEVICE_ID <> DP5_DPP_TYPES.dppMCA8000D)) Then
            strTemp = "Detector Temp: " & CStr(CInt(m_DP5_Status.DET_TEMP)) & "K" & vbCrLf     'Detector Temp
            strConfig = strConfig & strTemp
            strTemp = "Detector HV: " & CStr(CInt(m_DP5_Status.HV)) & "V" & vbCrLf     'Detector HV
            strConfig = strConfig & strTemp
            strTemp = "Board Temp: " & CStr(CInt(m_DP5_Status.DP5_TEMP)) & "°C" & vbCrLf     'Board Temp
            strConfig = strConfig & strTemp
        ElseIf (m_DP5_Status.DEVICE_ID = DP5_DPP_TYPES.dppDP5G) Then        ' GAMMARAD5
            If (m_DP5_Status.DET_TEMP > 0) Then
                strTemp = "Detector Temp: " & CStr(CInt(m_DP5_Status.DET_TEMP)) & "K" & vbCrLf     'Detector Temp
                strConfig = strConfig & strTemp
            Else
                strConfig = strConfig & ""
            End If
            '	strTemp.Format("HV Set: %.0fV\r\n",m_DP5_Status.HV);
            strConfig = strConfig & strTemp
        ElseIf (m_DP5_Status.DEVICE_ID = DP5_DPP_TYPES.dppMCA8000D) Then        ' Digital MCA
            strTemp = "Board Temp: " & CStr(CInt(m_DP5_Status.DP5_TEMP)) & "°C" & vbCrLf     'Board Temp
            strConfig = strConfig & strTemp
        End If

        If (m_DP5_Status.DEVICE_ID = DP5_DPP_TYPES.dppPX5) Then
            strTemp = PX5_OptionsString(m_DP5_Status)
            strConfig = strConfig & strTemp
            strTemp = "TEC V: " & String.Format(m_DP5_Status.TEC_Voltage, "0.000") & "V" & vbCrLf   'LiveTime
            strConfig = strConfig & strTemp
        End If

        ShowStatusValueStrings = strConfig
    End Function

    Private Function PX5_OptionsString(ByRef m_DP5_Status As Stat) As String
        Dim strOptions As String = ""
        Dim strValue As String = ""

        If (m_DP5_Status.DEVICE_ID = DP5_DPP_TYPES.dppPX5) Then
            'm_DP5_Status.DPP_options = 1;
            'm_DP5_Status.HPGe_HV_INH = true;
            'm_DP5_Status.HPGe_HV_INH_POL = true;
            If (m_DP5_Status.DPP_options > 0) Then
                '===============PX5 Options==================
                strOptions = strOptions & "PX5 Options: "
                If ((m_DP5_Status.DPP_options And 1) = 1) Then
                    strOptions = strOptions & "HPGe HVPS" & vbCrLf
                Else
                    strOptions = strOptions & "Unknown" & vbCrLf
                End If
                '===============HPGe HVPS HV Status==================
                strOptions = strOptions & "HPGe HV: "
                If (m_DP5_Status.HPGe_HV_INH) Then
                    strOptions = strOptions & "not inhibited" & vbCrLf
                Else
                    strOptions = strOptions & "inhibited" & vbCrLf
                End If
                '===============HPGe HVPS Inhibit Status==================
                strOptions = strOptions & "INH Polarity: "
                If (m_DP5_Status.HPGe_HV_INH_POL) Then
                    strOptions = strOptions & "high" & vbCrLf
                Else
                    strOptions = strOptions & "low" & vbCrLf
                End If
            Else
                strOptions = strOptions & "PX5 Options: None" & vbCrLf  'strOptions += "No Options Installed"
            End If
        End If
        PX5_OptionsString = strOptions
    End Function

    Private Sub ProcessSpectrumEx(ByRef PIN As Packet_In, ByRef DppState As DppStateType)
        Dim X As Short

        SPECTRUM.Channels = 256 * (2 ^ (((PIN.PID2 - 1) And 14) \ 2))
        ReDim SPECTRUM.DATA(SPECTRUM.Channels)

        For X = 0 To SPECTRUM.Channels - 1
            SPECTRUM.DATA(X) = CInt(PIN.DATA(X * 3)) + CInt(PIN.DATA(X * 3 + 1)) * 256 + CInt(PIN.DATA(X * 3 + 2)) * 65536
        Next X

        If (PIN.PID2 And 1) = 0 Then ' spectrum + status
            For X = 0 To 63
                STATUS.RAW(X) = PIN.DATA(X + SPECTRUM.Channels * 3)
            Next X
            Call Process_Status(STATUS)
            RequestScopeData(STATUS.SCOPE_DR)
        End If
        Call Plot_Spectrum(picMCAPlot, SPECTRUM, CheckLog.Checked, CheckLine.Checked)
        cmdSaveSpectrum.Enabled = True
        UpdateStatusDisplay(STATUS)

    End Sub


    Private Sub ProcessDiagDataEx(ByRef PIN As Packet_In, ByRef DppState As DppStateType)
        Dim dd As New DiagDataType
        Dim strDiag As String = ""

        dd.Initialize()
        Call Process_Diagnostics(PIN, dd)
        strDiag = DiagnosticsToString(dd)
    End Sub


    Private Sub RequestScopeData(ByRef ScopeDataReady As Boolean)
        If ScopeDataReady Then SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_SCOPE_DATA)
    End Sub

    Private Sub SendCommand(ByRef XmtCmd As modDP5_Protocol.TRANSMIT_PACKET_TYPE)
        Dim HaveBuffer As Boolean
        Dim SentPkt As Boolean

        HaveBuffer = DP5_CMD(BufferOUT, XmtCmd)
        If (HaveBuffer) Then
            Select Case s.CurrentInterface
                Case USB
                    SentPkt = SendPacketUSB(DppWinUSB, BufferOUT, PacketIn)
                    If (s.CurrentInterface = USB) Then
                        If (SentPkt) Then
                            RemCallParsePacket(PacketIn)
                        Else
                            Timer2.Enabled = False
                        End If
                    End If
            End Select
        End If
    End Sub



    Private Sub EnableDppCmdControls(ByRef EnableCmd As Boolean)

        '---- Configuration -----------------------------
        cmdShowConfiguration.Enabled = EnableCmd

        '---- Acquisition -------------------------------
        cmdStartAcquisition.Enabled = EnableCmd
        cmdStopAcquisition.Enabled = EnableCmd
        cmdSingleUpdate.Enabled = EnableCmd
        cmdClearData.Enabled = EnableCmd

        '---- MCA Spectrum File -------------------------
        cmdSaveSpectrum.Enabled = EnableCmd

        '---- Bacth Spectrum File -------------------------
        cmdSetFolder.Enabled = EnableCmd
        'cmdSetTime.Enabled = EnableCmd 'not in use currently

        s.isDppConnected = EnableCmd
    End Sub

    Private Function CountDP5WinusbDevices2() As Integer
        Dim hKey As Integer
        Dim retCode As Integer
        Dim lRet As Integer
        Dim idxDP5 As Integer
        Dim DevicePath(MAXREGBUFFER - 1) As Byte
        Dim cbDevicePath As Integer
        Dim KeyName As String = ""
        Dim cbKeyName As Integer
        Dim ErrMsg As String
        Dim cbErrMsg As Integer
        Dim strDevicePath As String = ""
        Dim strMsg As String = ""
        Dim strValueName As String = ""
        Dim iRemNulls As Integer = 0
        Dim strTest As String = ""

        NumUSBDevices = 0
        CountDP5WinusbDevices2 = 0
        retCode = RegOpenKeyEx(HKEY_LOCAL_MACHINE, WinUSBService, 0, KEY_QUERY_VALUE, hKey)
        If (retCode <> ERROR_SUCCESS) Then
            CountDP5WinusbDevices2 = 0
            Exit Function
        End If
        KeyName = Space(MAX_PATH)
        ' Test ALL Keys (0,1,... are device paths, Count,NextInstance,(Default) have other info)
        For idxDP5 = 0 To (MAXDP5S + 3) - 1 'devs + 3 other keys
            cbKeyName = MAX_PATH
            cbDevicePath = MAXREGBUFFER
            retCode = RegEnumValue(hKey, idxDP5, KeyName, cbKeyName, 0, REG_NONE, DevicePath(0), cbDevicePath)
            If (retCode = ERROR_SUCCESS) Then
                strValueName = Trim(KeyName)
                strDevicePath = ByteArrayToString(DevicePath)
                If (Len(strValueName) = 0) Then
                    'do nothing
                ElseIf (StrComp(KeyName.Substring(0, 5), "Count", CompareMethod.Text) = 0) Then
                    'do nothing
                ElseIf (StrComp(KeyName.Substring(0, 12), "NextInstance", CompareMethod.Text) = 0) Then
                    'do nothing
                ElseIf (StrComp(strDevicePath.Substring(0, WinUsbDP5Size), WinUsbDP5, CompareMethod.Text) = 0) Then  ' DP5 device path found
                    iRemNulls = InStr(strValueName, Chr(0))
                    If (iRemNulls > 0) Then
                        strValueName = strValueName.Substring(0, iRemNulls - 1)
                    End If
                    strMsg = strMsg & "KeyName " & strValueName & " "
                    iRemNulls = InStr(strDevicePath, Chr(0))
                    If (iRemNulls > 0) Then
                        strDevicePath = strDevicePath.Substring(0, iRemNulls - 1)
                    End If
                    strMsg = strMsg & strDevicePath & vbCrLf
                    'TRACE("DP5 device [%d]: %s=%s\r\n", (NumUSBDevices + 1), KeyName, DevicePath);
                    NumUSBDevices = NumUSBDevices + 1
                End If
            ElseIf (retCode = ERROR_NO_MORE_ITEMS) Then  ' no more values to read
                strMsg = strMsg & CStr(idxDP5) & " : " & CStr((MAXDP5S + 3) - 1) & " ERROR_NO_MORE_ITEMS" & vbCrLf
                Exit For
            Else ' error reading values
                cbErrMsg = MAXERRBUFFER
                ErrMsg = Space(MAXERRBUFFER)
                lRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, retCode, 0, ErrMsg, cbErrMsg, 0)
                If (lRet > 0) Then ErrMsg = ErrMsg.Substring(0, lRet) Else ErrMsg = ""
                strMsg = strMsg & "ErrMsg " & ErrMsg & vbCrLf
                Exit For
            End If
        Next
        RegCloseKey(hKey)
        CountDP5WinusbDevices2 = NumUSBDevices '/* return number of devices */
    End Function

    'sends a specified configuration to dpp
    Public Sub SendConfigToDpp(ByVal strCfg As String)
        s.CfgReadBack = True
        s.HwCfgDP5Out = strCfg
        If (Len(strCfg) > 0) Then
            SendCommand(TRANSMIT_PACKET_TYPE.XMTPT_SEND_CONFIG_PACKET_EX)
        End If
    End Sub

    '********************************
    'NMAA functions for scanner, encoders and sync
    '********************************

    Private Function WaitReadData(WaitChar As String, WaitTime As Integer) As String
        If SerialPortVXM.IsOpen Then
            Dim returnStr As String = ""
            Try
                SerialPortVXM.ReadTimeout = WaitTime
                Do
                    Dim Incoming As Integer = SerialPortVXM.ReadChar() 'ASCII character
                    'System.Diagnostics.Trace.WriteLine(Incoming) 'debugging to view ASCII
                    If Incoming = 94 Then '^ character, ASCII 94
                        returnStr &= Chr(Incoming) 'convert to actual characters
                        Exit Try
                    Else
                        'Keep waiting for ^
                    End If
                    Thread.Sleep(1) 'need to prevent missing data
                Loop
            Catch ex As TimeoutException
                LabelStatus.Text = "Error: Serial Port read timed out."
                WaitReadData = "Error"
            End Try
            WaitReadData = returnStr
            'LabelVXMOut.Text = returnStr
        Else
            WaitReadData = "No Connection Error"
        End If
    End Function

    Private Function WriteReadData(WriteToEX As String, PortLoc As String, ClearBuffer As Boolean) As String
        Dim SerialPortTemp As New System.IO.Ports.SerialPort()
        Select Case PortLoc
            Case "VXM"
                SerialPortTemp = SerialPortVXM
            Case "VROX"
                SerialPortTemp = SerialPortVROX
            Case "VROYZ"
                SerialPortTemp = SerialPortVROYZ
            Case Else
                LabelStatus.Text = "WriteReadData Case error"
        End Select
        If SerialPortTemp.IsOpen Then
            If ClearBuffer Then 'If ClearBuffer is True, then the incoming buffer is cleared prior to sending command
                SerialPortTemp.DiscardInBuffer()
                Thread.Sleep(2)
            End If
            Dim returnStr As String = ""
            SerialPortTemp.WriteLine(WriteToEX) 'send command to Encoder
            Thread.Sleep(10) 'need a brief pause to send and then receive some data into the feed. 10ms seems lowest for consistency.
            Try
                SerialPortTemp.ReadTimeout = 200
                Dim readattempt As Integer = 0
                Do
                    Dim bytelength As Integer = SerialPortTemp.BytesToRead
                    'System.Diagnostics.Trace.WriteLine(bytelength) 'read out the number of bytes each Do. Debugging only
                    If bytelength > 0 Then
                        Dim Incoming As Integer = SerialPortTemp.ReadChar() 'ASCII character
                        'System.Diagnostics.Trace.WriteLine(Incoming) 'debugging to view ASCII
                        If Incoming = 13 Then 'Removing Carriage return, ASCII 13. Appears in select data
                        Else
                            returnStr &= Chr(Incoming) 'convert to actual characters
                        End If
                        Thread.Sleep(1) 'need to prevent missing data
                    Else
                        If readattempt < 5 Then
                            readattempt = readattempt + 1
                            Thread.Sleep(5) 'need to prevent missing data
                        Else
                            LabelStatus.Text = "No more bytes"
                            Exit Do 'Could be Exit Do or Try
                        End If
                    End If
                Loop
            Catch ex As TimeoutException
                LabelStatus.Text = "Error: Serial Port read timed out."
                WriteReadData = "Error"
            End Try
            WriteReadData = returnStr
        Else
            WriteReadData = "No Connection Error"
        End If
    End Function

    Private Sub cmdSetFolder_Click(sender As Object, e As EventArgs) Handles cmdSetFolder.Click
        Dim strPath As String = ""
        Dim strData As String = ""
        Dim strFName As String = ""
        Dim saveFileDialogOut As New SaveFileDialog()
        saveFileDialogOut.Filter = dlgMCA_Filter
        If saveFileDialogOut.ShowDialog = DialogResult.OK Then
            strPath = IO.Path.GetDirectoryName(saveFileDialogOut.FileName) 'get the directory/path name
            strFName = IO.Path.GetFileName(saveFileDialogOut.FileName) 'get filename, with .mca
            lblBatchDir.Text = strPath
            lblBatchBase.Text = strFName
            cmdBatchStart.Enabled = True
        End If
    End Sub

    Private Sub txtPresetTime_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPresetTime.KeyPress
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBoxVXMinc_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBoxVXMinc.KeyPress
        If Not Char.IsNumber(e.KeyChar) And Not Char.IsControl(e.KeyChar) And Not ".".Contains(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub ButtonDSD_Click(sender As Object, e As EventArgs) Handles ButtonDSD.Click
        Try 'try to allows for exception handling
            If SerialPortDSD.IsOpen Then 'This is the port in the Form Design
                LabelStatus.Text = "DSD Already Connected. Reverting to NC."
                Try
                    SerialPortDSD.Write("AT+CH1=0") 'set to off which will keep NC closed
                    ButtonDSD.BackColor = Color.LightGreen
                Catch ex As Exception 'exception handling
                    LabelStatus.Text = "DSD Send Error"
                End Try
                Exit Sub
            End If
            SerialPortDSD.PortName = DSDPort
            SerialPortDSD.Open()
            LabelStatus.Text = "DSD Connected"
            ButtonDSD.BackColor = Color.LightGreen
            Try
                SerialPortDSD.Write("AT+CH1=0") 'set to off which will keep NC closed
            Catch ex As Exception 'exception handling
                LabelStatus.Text = "DSD Send Error"
            End Try
        Catch ex As Exception 'exception handling
            LabelStatus.Text = "DSD Connection Error"
        End Try
    End Sub

    Private Sub ButtonVXMConnect_Click(sender As Object, e As EventArgs) Handles ButtonVXMConnect.Click
        Dim ConnectCheck As Integer = 0
        Try 'try to allows for exception handling
            If SerialPortVXM.IsOpen Then 'This is the port in the Form Design
                LabelStatus.Text = "VXM already Connected"
                ConnectCheck = ConnectCheck + 1 'check that all components connect
                Exit Try
            End If
            SerialPortVXM.PortName = VXMPort
            SerialPortVXM.Open()
            LabelStatus.Text = "VXM Connected"
            Try
                SerialPortVXM.WriteLine("F,PM-0,S1M100,S3M2000,setj1M800,A1M1,setjA1M3,setJA1M1,setjA2M10,setJA2M2,setjA3M10,setJA3M2,R") 'set slower jog, speed, and accel for motor 1 belt, plus set joystick conditions
                'may want to have this change during going to corner start and then slow down again
                '400 seems safe jog speed. Was able to run at 800 without issue, but no weight.
                Thread.Sleep(1)
                'allow Jog mode
                SerialPortVXM.WriteLine("Q")
            Catch ex As Exception 'exception handling
                LabelStatus.Text = "Send Error"
            End Try
            ConnectCheck = ConnectCheck + 1 'check that all components connect
        Catch ex As FileNotFoundException
            'MessageBox.Show("VXM Port not found error: " & ex.Message)
            LabelStatus.Text = "VXM Error"
        Catch ex As UnauthorizedAccessException
            'MessageBox.Show("VXM Access denied to port: " & ex.Message)
            LabelStatus.Text = "VXM Error"
        Catch ex As IOException
            'MessageBox.Show("VXM I/O error: " & ex.Message)
            LabelStatus.Text = "VXM Error"
        Catch ex As InvalidOperationException
            'MessageBox.Show("VXM Invalid operation: " & ex.Message)
            LabelStatus.Text = "VXM Error"
        Catch ex As Exception
            'MessageBox.Show("VXM Serial Unexpected error: " & ex.Message)
            LabelStatus.Text = "VXM Error"
        End Try
        'Connect To Encoder for X-axis
        Try 'try to allows for exception handling
            If SerialPortVROX.IsOpen Then 'This is the port in the Form Design
                LabelStatus.Text = LabelStatus.Text & ", EncodeX already Connected"
                ConnectCheck = ConnectCheck + 1 'check that all components connect
                Exit Try
            End If
            SerialPortVROX.PortName = VROXPort
            SerialPortVROX.Open()
            LabelStatus.Text = LabelStatus.Text & ", EncodeX Connected"
            ConnectCheck = ConnectCheck + 1 'check that all components connect
        Catch ex As Exception 'exception handling
            LabelStatus.Text = LabelStatus.Text & ", EX Connection Error"
        End Try
        'Connect to Encoder for Y-Axis
        Try 'try to allows for exception handling
            If SerialPortVROYZ.IsOpen Then 'This is the port in the Form Design
                LabelStatus.Text = LabelStatus.Text & ", EncodeYZ already Connected"
                ConnectCheck = ConnectCheck + 1 'check that all components connect
                Exit Try
            End If
            SerialPortVROYZ.PortName = VROYZPort
            SerialPortVROYZ.Open()
            LabelStatus.Text = LabelStatus.Text & ", EncodeYZ Connected"
            ConnectCheck = ConnectCheck + 1 'check that all components connect
        Catch ex As Exception 'exception handling
            LabelStatus.Text = LabelStatus.Text & ", EYZ Connection Error"
        End Try
        If ConnectCheck = 3 Then
            ButtonVXMConnect.BackColor = Color.LightGreen
        End If
    End Sub




    Private Sub ButtonVXMjog_Click(sender As Object, e As EventArgs) Handles ButtonVXMjog.Click
        If SerialPortVXM.IsOpen Then
            Try
                SerialPortVXM.WriteLine("Q")
            Catch ex As Exception 'exception handling
                LabelStatus.Text = "Send Jog Error"
            End Try
        End If
    End Sub

    Private Sub ButtonVXMZero_Click(sender As Object, e As EventArgs) Handles ButtonVXMZero.Click
        If SerialPortVXM.IsOpen Then
            Try
                SerialPortVXM.WriteLine("N")
                'Zeroing motor position means image corners must be redefined. 
                VXMBatchTop = False
                VXMBatchBottom = False
            Catch ex As Exception 'exception handling
                LabelStatus.Text = "Send Zero Error VXM"
            End Try
        End If
    End Sub

    Private Sub ButtonVXMStatus_Click(sender As Object, e As EventArgs) Handles ButtonVXMStatus.Click
        'Get Status of VXM: R is Ready, B is Busy. A ^ may be present if this happens after a move.
        Dim returnStr As String = WriteReadData("V", "VXM", False) 'False will allow all buffer to be seen
        LabelVXMOut.Text = returnStr
    End Sub

    Private Sub ButtonVROsZero_Click(sender As Object, e As EventArgs) Handles ButtonVROsZero.Click
        If SerialPortVROX.IsOpen Then
            Try
                SerialPortVROX.WriteLine("N")
            Catch ex As Exception 'exception handling
                LabelStatus.Text = "Send Zero Error VROX"
            End Try
        End If
        If SerialPortVROYZ.IsOpen Then
            Try
                SerialPortVROYZ.WriteLine("N")
            Catch ex As Exception 'exception handling
                LabelStatus.Text = "Send Zero Error VROYZ"
            End Try
        End If
    End Sub

    Private Sub ButtonXYZLoc_Click(sender As Object, e As EventArgs) Handles ButtonXYZLoc.Click
        'Get location of each VXM motor axis
        'This currently works OK, but may have digits out  of place. If the motor is moving it may have ^ appear.
        'This has carriage character which may need to be removed for other cases (saving text of x,y,z)
        'Carriage return is ASCII 13
        'First read motor 1 as X
        Dim vxmXloc As String = WriteReadData("X", "VXM", True) 'True will clear buffer
        Thread.Sleep(1) 'need to prevent missing data
        'Second read motor 3 as Y (VXM axis 3 is wired to physical Y axis, but in VXM language is Z)
        Dim vxmYloc As String = WriteReadData("Z", "VXM", False) 'False will allow all buffer to be seen
        Thread.Sleep(1) 'need to prevent missing data
        'Third read motor 2 as Z (VXM axis 2 is wired to physical Z axis, but in VXM language is Y)
        Dim vxmZloc As String = WriteReadData("Y", "VXM", False) 'False will allow all buffer to be seen
        Thread.Sleep(1) 'need to prevent missing data and not garble up text
        Dim LocStringOut As String = "VXM: X=" & vxmXloc & vbCrLf & "Y=" & vxmYloc & vbCrLf & "Z=" & vxmZloc
        LabelVXMOut.Text = LocStringOut
        'Get location of encoder X axis primary position
        Dim vroXloc As String = WriteReadData("X", "VROX", True) 'False will allow all buffer to be seen
        Thread.Sleep(1) 'need to prevent missing data
        Dim vroYloc As String = WriteReadData("X", "VROYZ", True) 'False will allow all buffer to be seen
        Thread.Sleep(1) 'need to prevent missing data
        Dim vroZloc As String = WriteReadData("Y", "VROYZ", True) 'False will allow all buffer to be seen
        Thread.Sleep(1) 'need to prevent missing data
        Dim returnStr As String = "VRO: x=" & vroXloc & vbCrLf & "y=" & vroYloc & vbCrLf & "z=" & vroZloc
        Thread.Sleep(1) 'need to prevent missing data and not garble up text
        LabelEX.Text = returnStr
    End Sub

    Private Sub ButtonVXMTopCorner_Click(sender As Object, e As EventArgs) Handles ButtonVXMTopCorner.Click
        'Get location of VXM motor axis x and y as start positions
        If SerialPortVXM.IsOpen And SerialPortVROX.IsOpen And SerialPortVROYZ.IsOpen Then
            Try
                'Get X and Y axes (coded as X and Z per VXM connection)
                Dim vxmXloc As String = WriteReadData("X", "VXM", True) 'True will clear buffer
                Thread.Sleep(1) 'need to prevent missing data
                Dim vxmYloc As String = WriteReadData("Z", "VXM", False) 'False will allow all buffer to be seen
                Thread.Sleep(1) 'need to prevent missing data
                VXMxstart = Val(vxmXloc)
                VXMystart = Val(vxmYloc)
                Dim LocStringOut As String = "X=" & vxmXloc & vbCrLf & "Y=" & vxmYloc
                LabelVXMOut.Text = LocStringOut
                Dim vroXloc As String = WriteReadData("X", "VROX", True) 'False will allow all buffer to be seen
                Thread.Sleep(1) 'need to prevent missing data
                Dim vroYloc As String = WriteReadData("X", "VROYZ", False) 'False will allow all buffer to be seen
                Dim returnStr As String = "VRO: x=" & vroXloc & vbCrLf & "y=" & vroYloc
                LabelTCVRO.Text = "x=" & vroXloc & "; y=" & vroYloc
                EXstart = Val(returnStr) 'this is not currently used.
                LabelEX.Text = returnStr
                VXMBatchTop = True
                ButtonVXMTopCorner.BackColor = Color.LightBlue
            Catch ex As Exception 'exception handling
                LabelStatus.Text = "Connection Error"
            End Try
        End If
    End Sub

    Private Sub ButtonVXMBottomCorner_Click(sender As Object, e As EventArgs) Handles ButtonVXMBottomCorner.Click
        If SerialPortVXM.IsOpen And SerialPortVROX.IsOpen And SerialPortVROYZ.IsOpen Then
            Try
                'Get location of VXM motor axis x and y as end positions
                'Get X and Y axes (coded as X and Z per VXM connection)
                Dim vxmXloc As String = WriteReadData("X", "VXM", True) 'True will clear buffer
                Thread.Sleep(1) 'need to prevent missing data
                Dim vxmYloc As String = WriteReadData("Z", "VXM", False) 'False will allow all buffer to be seen
                Thread.Sleep(1) 'need to prevent missing data
                VXMxend = Val(vxmXloc)
                VXMyend = Val(vxmYloc)
                Dim LocStringOut As String = "X=" & vxmXloc & vbCrLf & "Y=" & vxmYloc
                LabelVXMOut.Text = LocStringOut
                Dim vroXloc As String = WriteReadData("X", "VROX", True) 'False will allow all buffer to be seen
                Thread.Sleep(1) 'need to prevent missing data
                Dim vroYloc As String = WriteReadData("X", "VROYZ", False) 'False will allow all buffer to be seen
                Dim returnStr As String = "VRO: x=" & vroXloc & vbCrLf & "y=" & vroYloc
                LabelBCVRO.Text = "x=" & vroXloc & "; y=" & vroYloc
                EXend = Val(returnStr) 'this is not currently used.
                LabelEX.Text = returnStr
                VXMBatchBottom = True
                ButtonVXMBottomCorner.BackColor = Color.LightBlue
            Catch ex As Exception 'exception handling
                LabelStatus.Text = "Connection Error"
            End Try
        End If
    End Sub

    Private Sub ButtonClearCorners_Click(sender As Object, e As EventArgs) Handles ButtonClearCorners.Click
        VXMBatchTop = False
        ButtonVXMTopCorner.BackColor = Color.Gainsboro
        VXMBatchBottom = False
        ButtonVXMBottomCorner.BackColor = Color.Gainsboro
        LabelTCVRO.Text = ""
        LabelBCVRO.Text = ""
    End Sub

    Private Sub ButtonVXMCalcBatch_Click(sender As Object, e As EventArgs) Handles ButtonVXMCalcBatch.Click
        If VXMBatchTop And VXMBatchBottom Then
            VXMxwidth = VXMxend - VXMxstart 'x width of area in motor steps
            VXMywidth = VXMyend - VXMystart 'y width of area in motor steps
            VXMxwidthmm = Math.Round(VXMxwidth / VXMxmmpstepres, 2) 'x width of area in mm
            VXMywidthmm = Math.Round(VXMywidth / VXMyzmmpstepres, 2) 'y width of area in mm
            VXMxstepsize = Math.Floor(Val(TextBoxVXMinc.Text) * VXMxmmpstepres) 'how many x steps to get to defined mm spacing
            VXMystepsize = Math.Floor(Val(TextBoxVXMinc.Text) * VXMyzmmpstepres) 'how many y steps to get to defined mm spacing
            VXMxiternum = Math.Floor(VXMxwidth / VXMxstepsize) 'how many jumps to get across area
            VXMyiternum = Math.Floor(VXMywidth / VXMystepsize)
            Dim timecalc As Integer = Math.Round(Val(VXMxiternum * VXMyiternum * (System.Convert.ToInt32(txtPresetTime.Text) + 320) + (VXMxiternum * 2000)), 0)
            Dim timeinfo As TimeSpan = TimeSpan.FromMilliseconds(timecalc)
            LabelBacthTime.Text = "Batch Time Info." 'resets the text if new scan is run
            LabelVXMOut.Text = "Batch Calculations " & vbCrLf &
                "Step width X " & VXMxwidth & ", Y " & VXMywidth & vbCrLf &
                "mm width X " & VXMxwidthmm & ", Y " & VXMywidthmm & vbCrLf &
                "# X iter" & VXMxiternum & "# Y iter" & VXMyiternum & vbCrLf &
                "Totals iter" & Val(VXMxiternum * VXMyiternum) & vbCrLf &
                "Minimum total time " & Math.Round((timecalc / 60000), 0) & " min" & vbCrLf &
                "(" & timeinfo.Hours & " hours, " & timeinfo.Minutes & " min.)"
        Else
            LabelStatus.Text = "Error: Top and Bottom need to be set first."
        End If

    End Sub

    Private Sub cmdBatchStart_Click(sender As Object, e As EventArgs) Handles cmdBatchStart.Click
        Dim timeiter As Int32 = System.Convert.ToInt32(txtPresetTime.Text)
        Dim strBasename As String = lblBatchBase.Text
        Dim outfilename As String = "" 'mca direct save w/o dialog
        Dim encodexfilename As String = lblBatchDir.Text & "\" & strBasename.Substring(0, Len(strBasename) - 4) & "_xpos" & ".txt"
        Dim encodeinfofilename As String = lblBatchDir.Text & "\" & strBasename.Substring(0, Len(strBasename) - 4) & "_info" & ".txt"
        Dim batchcurStart As Decimal
        Dim batchElapse As Decimal
        Dim VXMxmovecheck As Integer = 0 'failed move counter. 0 = success
        Dim posarray(VXMxiternum - 1) As String 'list of position strings from encoder
        LabelBacthTime.Text = "Time estimate updates at x-axis move." 'timer starts after in position
        cmdBatchStop.Enabled = True
        If SerialPortVXM.IsOpen Then
            Try
                SerialPortVXM.DiscardInBuffer() 'clear incoming buffer
                'Go to start by absolute command
                SerialPortVXM.WriteLine("F,PM-0,IA1M" & Str(VXMxstart) & ",IA3M" & Str(VXMystart) & ",R")
                LabelStatus.Text = WaitReadData("^", 60000) 'give VXM 60 sec to get into position.

                batchcurStart = msTimeStart() 'start timer to monitor batch progress

                'begin info file
                Using infowriter As New StreamWriter(encodeinfofilename, False) ' False = overwrite if file exists
                    infowriter.WriteLine("Amptek XRF Scanning File")
                    infowriter.WriteLine("Data spacing mm: " & TextBoxVXMinc.Text)
                    infowriter.WriteLine("Acc Time ms: " & timeiter)
                    infowriter.WriteLine("Encoder top set mm: " & LabelTCVRO.Text)
                    infowriter.WriteLine("Encoder bottom set mm: " & LabelBCVRO.Text)
                End Using

                'Send set of move commands to create raster
                Dim MoveString As String = ""
                Dim XYIterCounter As Integer = 0
                Dim YIterSerp As Integer = 1
                Dim curStart As Decimal 'start time from system time in milliseconds
                Dim TimeExpired As Boolean 'time expired flag
                Dim curElapsed As Decimal 'Elapsed time from start time in milliseconds
                Dim strCfg As String = ""
                Dim strStatus As String = ""
                SendCommand(TRANSMIT_PACKET_TYPE.XMTPT_DISABLE_MCA_MCS) 'pause mca
                SendConfigToDpp("PRET=" & timeiter / 1000 & ";") 'send preset time
                SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_CLEAR_SPECTRUM_STATUS) 'clear any remaining data

                'Get actual Encoder start position
                Dim tempvroXloc As String = WriteReadData("X", "VROX", True) 'False will allow all buffer to be seen
                Thread.Sleep(1) 'need to prevent missing data
                Dim tempvroYloc As String = WriteReadData("X", "VROYZ", True) 'False will allow all buffer to be seen
                Thread.Sleep(1) 'need to prevent missing data
                Using infowriter As New StreamWriter(encodeinfofilename, True) ' True = append mode
                    infowriter.WriteLine("Encoder top actual mm: x=" & tempvroXloc & ": y=" & tempvroYloc)
                End Using


                For i As Integer = 1 To VXMxiternum
                    If i = 1 Then
                        Dim returnStr As String = WriteReadData("X", "VROX", True) 'False will allow all buffer to be seen
                        LabelEX.Text = returnStr
                        posarray(i - 1) = returnStr
                    End If
                    For j As Integer = 1 To VXMyiternum
                        If cmdBatchStop.Enabled Then 'Check if stop was pressed
                        Else
                            Exit Try
                        End If
                        'Move VXM unless on first spot of a row
                        If j = 1 Then
                        Else
                            MoveString = "F,PM-0,I3M" & Str(YMotorDir * VXMystepsize).Trim & ",R"
                            SerialPortVXM.WriteLine(MoveString)
                            LabelStatus.Text = WaitReadData("^", 2000) 'give VXM 2 sec to complete or timeout
                            YIterSerp = YIterSerp + YMotorDir
                            msDelay(YMotorPause) 'short delay after moving Y 
                        End If
                        XYIterCounter = XYIterCounter + 1 'spectra counter
                        LabelStatus.Text = "Spot " & Str(XYIterCounter).Trim & " of " & Str(VXMxiternum * VXMyiternum).Trim
                        'Acquire XRF
                        If (s.isDppConnected) Then
                            s.HwCfgReady = False 'clear cfg ready flag
                            s.HwCfgDP5 = "" 'clear config readback string
                            s.cstrRawCfgIn = ""
                            TimeExpired = False
                            curStart = msTimeStart()
                            s.SpectrumCfg = True
                            SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_FULL_READ_CONFIG_PACKET)
                            Do  'wait until s.HwCfgReady or timeout
                                System.Windows.Forms.Application.DoEvents()
                                curElapsed = msTimeDiff(curStart)
                                TimeExpired = msTimeExpired(curStart, 1000) '1000 milliseconds max wait
                            Loop Until (TimeExpired Or s.HwCfgReady)

                            'read cfg in s.HwCfgDP5
                            If ((Len(s.HwCfgDP5) > 0) And s.HwCfgReady) Then
                                strCfg = s.HwCfgDP5
                                If (Microsoft.VisualBasic.Strings.Right(strCfg, 2) = vbCrLf) Then
                                    strCfg = strCfg.Substring(0, Len(strCfg) - 2)
                                End If
                            End If
                        End If

                        If (STATUS.SerialNumber > 0) Then
                            strStatus = ShowStatusValueStrings(STATUS)
                            If (Microsoft.VisualBasic.Strings.Right(strStatus, 2) = vbCrLf) Then
                                strStatus = strStatus.Substring(0, Len(strStatus) - 2)
                            End If
                        End If

                        'naming with X######Y######
                        outfilename = lblBatchDir.Text & "\" & strBasename.Substring(0, Len(strBasename) - 4) & "_X" & i.ToString("D8") & "Y" & YIterSerp.ToString("D8") & ".mca"

                        SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_ENABLE_MCA_MCS)
                        msDelay(timeiter) 'pause for set ms. This is the integration time (real time)
                        SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_DISABLE_MCA_MCS)
                        SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_CLEAR_SPECTRUM_STATUS)
                        SaveSpectrumBacth(SPECTRUM, strStatus, strCfg, txtSpectrumTag.Text, txtSpectrumDescription.Text, outfilename)

                    Next
                    'Move X Motor 1 unless at end of last line
                    'INCLUDES extra steps to account for belt drive - tested well for 1 and 2 mm spacing
                    If i = VXMxiternum Then
                        LabelBacthTime.Text = Math.Round((batchElapse / 60000), 0) & " min elapsed."
                    Else
                        'MoveString = "F,PM-0,I1M" & (Str(VXMxstepsize).Trim + VXMxextra) & ",R"
                        MoveString = "F,PM-0,I1M" & (Str(VXMxstepsize + VXMxmmpstepresmod).Trim) & ",R"
                        SerialPortVXM.WriteLine(MoveString)
                        YMotorDir = YMotorDir * -1
                        batchElapse = msTimeDiff(batchcurStart)
                        Dim timeinfo As TimeSpan = TimeSpan.FromMilliseconds(batchElapse / i * (VXMxiternum - i))
                        LabelBacthTime.Text = Math.Round((batchElapse / 60000), 0) & " min elapsed." & vbCrLf &
                            "Est. " & Math.Round(((batchElapse / i * (VXMxiternum - i)) / 60000), 0) & " min remain." & vbCrLf &
                            "(" & timeinfo.Hours & " hours, " & timeinfo.Minutes & " min.)"
                        WaitReadData("^", 2000) 'wait until move is done
                        'Read encoderX position and keep in array
                        Dim returnStr As String = WriteReadData("X", "VROX", True) 'False will allow all buffer to be seen
                        LabelEX.Text = returnStr
                        posarray(i) = returnStr
                        'If fail to move 0.1 mm after 5 attempts, stop batch
                        If Val(posarray(i)) - Val(posarray(i - 1)) < 0.1 Then
                            VXMxmovecheck = VXMxmovecheck + 1
                            If VXMxmovecheck = 5 Then
                                LabelEX.Text = "Movement failed 5 times! Run aborted!"
                                cmdBatchStop.Enabled = False
                                msDelay(10)
                            End If
                        Else
                            VXMxmovecheck = 0
                        End If
                        'calculate move and adjust steps by units of 10.
                        If Val(posarray(i)) - Val(posarray(i - 1)) < (Val(TextBoxVXMinc.Text) - 0.25) Then
                            VXMxmmpstepresmod = VXMxmmpstepresmod + 10
                        ElseIf Val(posarray(i)) - Val(posarray(i - 1)) > (Val(TextBoxVXMinc.Text) + 0.25) Then
                            VXMxmmpstepresmod = VXMxmmpstepresmod - 10
                        End If
                        msDelay(XMotorPause) 'longer delay after moving X 
                    End If
                Next
            Catch ex As Exception 'exception handling
                LabelStatus.Text = "Connection Error"
            End Try
            'Turn off Xray by flipping DSD interlock
            If SerialPortDSD.IsOpen Then 'This is the port in the Form Design
                SerialPortDSD.Write("AT+CH1=1") 'set to on which will open NC - XRay off
                msDelay(500)
                SerialPortDSD.Write("AT+CH1=0") 'set to off which will keep NC closed - XRay allowed (but not on yet)
            End If
            YMotorDir = 1 'Reset y-motor direction, in case another scan is performed
            File.WriteAllLines(encodexfilename, posarray)
            'Get actual Encoder end position
            Dim vroXloc As String = WriteReadData("X", "VROX", True) 'False will allow all buffer to be seen
            Thread.Sleep(1) 'need to prevent missing data
            Dim vroYloc As String = WriteReadData("X", "VROYZ", True) 'False will allow all buffer to be seen
            Thread.Sleep(1) 'need to prevent missing data
            Using infowriter As New StreamWriter(encodeinfofilename, True) ' True = append mode
                infowriter.WriteLine("Encoder bottom actual mm: x=" & vroXloc & ": y=" & vroYloc)
            End Using
            SerialPortVXM.WriteLine("Q") 'allow Jog mode
            cmdBatchStop.Enabled = False
        End If
    End Sub

    Private Sub ButtonXRayOFF_Click(sender As Object, e As EventArgs) Handles ButtonXRayOFF.Click
        Try 'try to allows for exception handling
            If SerialPortDSD.IsOpen Then 'This is the port in the Form Design
                SerialPortDSD.Write("AT+CH1=1") 'set to on which will open NC
                ButtonDSD.BackColor = Color.Red
                Exit Sub
            End If
        Catch ex As Exception 'exception handling
            LabelStatus.Text = "DSD Connection Error"
        End Try
    End Sub

    Private Sub ClosePort(SerialPortName)
        Try 'try to allows for exception handling
            If SerialPortName.IsOpen Then 'This is the port in the Form Design
                SerialPortName.Close()
                Exit Try
            End If
        Catch ex As Exception 'exception handling
            MessageBox.Show(ex.Message, "Error closing serial port", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmDP5_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Timer2.Stop()
        SaveApplicationSettings(Me)
        CloseDeviceHandle(DppWinUSB)
        ClosePort(SerialPortVXM)
        ClosePort(SerialPortVROX)
        ClosePort(SerialPortVROYZ)
        ClosePort(SerialPortDSD)
        Application.Exit()
    End Sub


End Class
