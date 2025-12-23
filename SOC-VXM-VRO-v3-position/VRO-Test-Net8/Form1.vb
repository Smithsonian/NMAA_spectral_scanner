Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.IO
Imports System.IO.Ports
Imports System.Threading
Imports System.Reflection.Emit
Imports System.Runtime.Intrinsics
' Import the UD .NET wrapper object.  The dll ByReferenced is installed by the
' LabJackUD installer.
Imports LabJack.LabJackUD
Public Class Form1
    'Set up Connection names and Equip Specs
    'Set up Connection names and Equip Specs
    'Connect COM Ports based on .ini file (need to manually edit file)
    Dim iniPath As String = My.Application.Info.DirectoryPath & "\" & "COMSettings" & ".ini"
    Dim VXMPort As String = CSIniReader.ReadIniValue("COMPorts", "VXMPort", iniPath)
    Dim VROXPort As String = CSIniReader.ReadIniValue("COMPorts", "VROXPort", iniPath)
    Dim VROYZPort As String = CSIniReader.ReadIniValue("COMPorts", "VROYZPort", iniPath)

    'Can revert to direct coding port if necessary
    'Dim VXMPort As String = "COM4"
    'Dim VROXPort As String = "COM5"
    'Dim VROYZPort As String = "COM6"

    Dim SerialPortVXM As New System.IO.Ports.SerialPort()
    Dim SerialPortVROX As New System.IO.Ports.SerialPort()
    Dim SerialPortVROYZ As New System.IO.Ports.SerialPort()
    Dim u3 As U3 'LabJack device
    Dim VXMxmmpstepres As Double = 40 'number of steps to move one mm - x belt
    Dim VXMyzmmpstepres As Double = 157.5 'number of steps to move one mm - y,z screw
    Dim SOCpix As Integer = 1024 ' pixels in nonscan direction
    Dim voltOn As Double = 3.5 'set the on voltage. Should be 3.5 for SOC TTL. Could be lower for lasers. max 5.
    Dim TTLdelay As Integer = 20 'ms of TTL, 20 ms minimum for stability?
    'Set up other variables
    Dim VXMxstart As Integer = 0
    Dim VXMxend As Integer = 0
    Dim VXMystart As Integer = 0
    Dim VXMyend As Integer = 0
    Dim VXMBatchTop = False
    Dim VXMBatchBottom = False
    Dim VXMBatchAllowed = False
    Dim VXMxwidth As Integer  'x width of area in motor steps
    Dim VXMywidth As Integer 'y width of area in motor steps
    Dim VXMxwidthmm As Double 'x width of area in mm
    Dim VXMywidthmm As Double 'y width of area in mm
    Dim VXMxstepsize As Double 'how many x steps to get to defined mm spacing
    Dim VXMystepsize As Double 'how many y steps to get to defined mm spacing
    Dim VXMyspeed As Integer 'set the speed of y-axis movement
    Dim VXMxiternum As Integer 'how many jumps to get across area
    Dim VXMyiternum As Integer
    Dim YMotorDir As Integer = 1
    Dim VXMSOCpixmm As Double 'pixels per mm
    Dim VXMSOCfps As Double 'frame per sec - integration plus delay
    Dim VXMSOCmsmm As Double 'millisec per mm
    Dim stopWatch As New Stopwatch() 'needed for timer routine
    Dim milliSec As Long = 0





    Private Sub ButtonVRO2_Click(sender As Object, e As EventArgs) Handles ButtonVRO2.Click
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
            ConnectCheck = ConnectCheck + 1 'check that all components connect
        Catch ex As Exception 'exception handling
            LabelStatus.Text = "VXM Connection Error"
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
            ButtonVRO2.BackColor = Color.LightGreen
            stopWatch.Start()
        End If
    End Sub

    Public Function msDelay(ByRef MSec As Integer) As Boolean
        Dim curStart As Decimal
        Dim TimeExpired As Boolean
        TimeExpired = False
        curStart = stopWatch.ElapsedMilliseconds
        Do
            System.Windows.Forms.Application.DoEvents()
            TimeExpired = msTimeExpired(curStart, CDec(MSec))
        Loop While (Not TimeExpired)
        msDelay = TimeExpired
    End Function

    Public Function msTimeExpired(ByRef curStartTimeMS As Decimal, ByRef curDelayTimeMS As Decimal) As Boolean
        Dim curNewTime As Decimal
        curNewTime = stopWatch.ElapsedMilliseconds
        msTimeExpired = CBool(curNewTime >= (curStartTimeMS + curDelayTimeMS))
    End Function


    Private Sub ButtonLXYZ_Click(sender As Object, e As EventArgs) Handles ButtonLXYZ.Click
        'Get encoder X, Y, Z positions
        'First read motor 1 as X
        'Get location of encoder X axis primary position
        Dim vroXloc As String = WriteReadData("X", "VROX", True) 'False will allow all buffer to be seen
        Thread.Sleep(1) 'need to prevent missing data
        Dim vroYloc As String = WriteReadData("X", "VROYZ", True) 'False will allow all buffer to be seen
        Thread.Sleep(1) 'need to prevent missing data
        Dim vroZloc As String = WriteReadData("Y", "VROYZ", True) 'False will allow all buffer to be seen
        Thread.Sleep(1) 'need to prevent missing data
        Dim returnStr As String = "VRO: x=" & vroXloc & vbCrLf & "y=" & vroYloc & vbCrLf & "z=" & vroZloc
        Thread.Sleep(1) 'need to prevent missing data and not garble up text
        LabelInfo.Text = returnStr
    End Sub



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


    Private Sub ButtonOpenPorts_Click(sender As Object, e As EventArgs) Handles ButtonOpenPorts.Click
        Dim PortArray() As String
        PortArray = IO.Ports.SerialPort.GetPortNames
        LabelStatus.Text = String.Join(",", PortArray)
    End Sub

    Private Sub ButtonLJconnect_Click(sender As Object, e As EventArgs) Handles ButtonLJconnect.Click
        Try
            u3 = New U3(LJUD.CONNECTION.USB, "0", True) ' Connection through USB
            'THIS RESET LIKELY NOT NEEDED
            'Start by using the pin_configuration_reset IOType so that all
            'pin assignments are in the factory default condition.
            LJUD.ePut(u3.ljhandle, LJUD.IO.PIN_CONFIGURATION_RESET, 0, 0, 0)
            LJUD.ePut(u3.ljhandle, LJUD.IO.PUT_DAC, 1, 0, 0)   ' Set DAC1 To 0 volts.
            ButtonLJconnect.BackColor = Color.LightGreen
        Catch ex As LabJackUDException
            showErrorMessage(ex)
        End Try
    End Sub

    Sub showErrorMessage(ByVal err As LabJackUDException)
        MsgBox("Function returned LabJackUD Error #" &
            Str$(err.LJUDError) &
            "  " &
            err.ToString)
    End Sub

    Private Sub TextBoxCycles_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBoxCycles.KeyPress
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBoxIntms_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBoxIntms.KeyPress
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBoxFOV_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBoxFOV.KeyPress
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub ButtonVXMEdge_Click(sender As Object, e As EventArgs) Handles ButtonVXMEdge.Click
        'Get location of VXM motor axis x and y as start positions
        If SerialPortVXM.IsOpen Then
            Try
                Dim vxmXloc As String = WriteReadData("X", "VXM", True) 'True will clear buffer
                Thread.Sleep(1) 'need to prevent missing data
                'physical y-axis is motor 3 which is coded as z
                Dim vxmYloc As String = WriteReadData("Z", "VXM", False) 'False will allow all buffer to be seen
                Thread.Sleep(1) 'need to prevent missing data
                VXMxstart = Val(vxmXloc)
                VXMystart = Val(vxmYloc)
                Dim LocStringOut As String = "X=" & vxmXloc & vbCrLf & "Y=" & vxmYloc
                LabelInfo.Text = LocStringOut
                VXMBatchTop = True
                ButtonVXMEdge.BackColor = Color.LightGreen
                'Get encoder Y, first port on second encoder
                Dim returnStr As String = WriteReadData("X", "VROYZ", True) 'False will allow all buffer to be seen
                LabelYEdge.Text = returnStr & " mm"
            Catch ex As Exception 'exception handling
                LabelStatus.Text = "Connection Error"
                ButtonVXMEdge.BackColor = Color.WhiteSmoke
            End Try
        End If
    End Sub

    Private Sub ButtonCalculate_Click(sender As Object, e As EventArgs) Handles ButtonCalculate.Click
        VXMywidth = Math.Floor(VXMyzmmpstepres * Val(TextBoxFOV.Text) * (Val(TextBoxCycles.Text) / SOCpix)) 'y width of area in motor steps
        VXMSOCpixmm = Math.Round(SOCpix / Val(TextBoxFOV.Text), 3) 'pixels per mm
        VXMSOCfps = Math.Round(1000 / (TTLdelay + Val(TextBoxIntms.Text)), 2) 'pixels per mm
        VXMSOCmsmm = Math.Round(VXMSOCpixmm * (TTLdelay + Val(TextBoxIntms.Text)), 2)
        VXMywidthmm = Math.Round(Val(TextBoxFOV.Text) * (Val(TextBoxCycles.Text) / SOCpix), 2) 'y width of area in mm
        VXMyspeed = 1000 / (VXMSOCmsmm / VXMyzmmpstepres) 'steps per second scan speed
        LabelInfo.Text = "Batch Calculations " & vbCrLf &
                            "Pixels/mm goal " & VXMSOCpixmm & ", fps " & VXMSOCfps & vbCrLf &
                            "ms/mm goal " & VXMSOCmsmm & vbCrLf &
                            "Scan width mm" & VXMywidthmm & ", steps " & VXMywidth & vbCrLf &
                            "Scan speed step/s " & VXMyspeed & vbCrLf &
                            "Time est " & Math.Ceiling(VXMywidth / VXMyspeed) & " s."
        ButtonStart.Enabled = True
    End Sub

    Private Sub ButtonMYup_Click(sender As Object, e As EventArgs) Handles ButtonMYup.Click
        If SerialPortVXM.IsOpen Then
            Try
                SerialPortVXM.DiscardInBuffer() 'clear incoming buffer
                'go to default speed and move 10 % FOV
                Dim MoveString As String = ""
                MoveString = "F,PM-0,S3M2000,I3M" & Str(Math.Floor(VXMyzmmpstepres * Val(TextBoxFOV.Text) * 0.1)) & ",R"
                SerialPortVXM.WriteLine(MoveString)
                LabelStatus.Text = WaitReadData("^", 10000) 'give VXM 10 sec to get into position.
            Catch ex As Exception 'exception handling
                LabelStatus.Text = "Connection Error"
            End Try
        End If
    End Sub

    Private Sub ButtonMYdn_Click(sender As Object, e As EventArgs) Handles ButtonMYdn.Click
        If SerialPortVXM.IsOpen Then
            Try
                SerialPortVXM.DiscardInBuffer() 'clear incoming buffer
                'go to default speed and move 10 % FOV
                Dim MoveString As String = ""
                MoveString = "F,PM-0,S3M2000,I3M-" & Str(Math.Floor(VXMyzmmpstepres * Val(TextBoxFOV.Text) * 0.1)) & ",R"
                SerialPortVXM.WriteLine(MoveString)
                LabelStatus.Text = WaitReadData("^", 10000) 'give VXM 10 sec to get into position.
            Catch ex As Exception 'exception handling
                LabelStatus.Text = "Connection Error"
            End Try
        End If
    End Sub

    Private Sub ButtonMXup_Click(sender As Object, e As EventArgs) Handles ButtonMXup.Click
        If SerialPortVXM.IsOpen Then
            Try
                SerialPortVXM.DiscardInBuffer() 'clear incoming buffer
                'go to default speed and move 10 % FOV
                Dim MoveString As String = ""
                MoveString = "F,PM-0,I1M" & Str(Math.Floor(VXMxmmpstepres * Val(TextBoxFOV.Text) * 0.1)) & ",R"
                SerialPortVXM.WriteLine(MoveString)
                LabelStatus.Text = WaitReadData("^", 10000) 'give VXM 10 sec to get into position.
            Catch ex As Exception 'exception handling
                LabelStatus.Text = "Connection Error"
            End Try
        End If
    End Sub

    Private Sub ButtonMXdn_Click(sender As Object, e As EventArgs) Handles ButtonMXdn.Click
        If SerialPortVXM.IsOpen Then
            Try
                SerialPortVXM.DiscardInBuffer() 'clear incoming buffer
                'go to default speed and move 10 % FOV
                Dim MoveString As String = ""
                MoveString = "F,PM-0,I1M-" & Str(Math.Floor(VXMxmmpstepres * Val(TextBoxFOV.Text) * 0.1)) & ",R"
                SerialPortVXM.WriteLine(MoveString)
                LabelStatus.Text = WaitReadData("^", 10000) 'give VXM 10 sec to get into position.
            Catch ex As Exception 'exception handling
                LabelStatus.Text = "Connection Error"
            End Try
        End If
    End Sub

    Private Sub ButtonMXupfar_Click(sender As Object, e As EventArgs) Handles ButtonMXupfar.Click
        If SerialPortVXM.IsOpen Then
            Try
                SerialPortVXM.DiscardInBuffer() 'clear incoming buffer
                'go to default speed and move 10 % FOV
                Dim MoveString As String = ""
                MoveString = "F,PM-0,I1M" & Str(Math.Floor(VXMxmmpstepres * Val(TextBoxFOV.Text) * 0.8)) & ",R"
                SerialPortVXM.WriteLine(MoveString)
                LabelStatus.Text = WaitReadData("^", 10000) 'give VXM 10 sec to get into position.
            Catch ex As Exception 'exception handling
                LabelStatus.Text = "Connection Error"
            End Try
        End If
    End Sub

    Private Sub ButtonMYedge_Click(sender As Object, e As EventArgs) Handles ButtonMYedge.Click
        If SerialPortVXM.IsOpen And VXMBatchTop Then
            Try
                SerialPortVXM.DiscardInBuffer() 'clear incoming buffer
                'go to default speed and move 10 % FOV
                Dim MoveString As String = ""
                MoveString = "F,PM-0,S3M2000,I3AM" & Str(VXMystart) & ",R"
                SerialPortVXM.WriteLine(MoveString)
                LabelStatus.Text = WaitReadData("^", 30000) 'give VXM 30 sec to get into position.
            Catch ex As Exception 'exception handling
                LabelStatus.Text = "Connection Error"
            End Try
        End If
    End Sub

    Private Sub ButtonStart_Click(sender As Object, e As EventArgs) Handles ButtonStart.Click
        ButtonStop.Enabled = True
        msDelay(2)
        If SerialPortVXM.IsOpen Then
            Try
                SerialPortVXM.DiscardInBuffer() 'clear incoming buffer
                'Send set of move commands to move slowly sending TTL
                Dim MoveString As String = ""
                MoveString = "F,PM-0,S3M" & Str(VXMyspeed) & ",I3M" & Str(Math.Floor(VXMyzmmpstepres * Val(TextBoxFOV.Text) * (Math.Round(Val(TextBoxCycles.Text) / SOCpix)))) & ",R"
                SerialPortVXM.WriteLine(MoveString)
                For i As Integer = 1 To Str(TextBoxCycles.Text)
                    LabelStatus.Text = "Line " & Str(i)
                    If ButtonStop.Enabled Then ' Check if stop button pressed
                    Else
                        Exit Try
                    End If
                    LJUD.ePut(u3.ljhandle, LJUD.IO.PUT_DAC, 1, voltOn, 0)   ' Set DAC1 To high volts.
                    msDelay(TTLdelay) '10 ms too fast
                    LJUD.ePut(u3.ljhandle, LJUD.IO.PUT_DAC, 1, 0, 0)   ' Set DAC1 To 0 volts.
                    msDelay(Str(TextBoxIntms.Text))
                Next
            Catch ex As Exception 'exception handling
                LabelStatus.Text = "Connection Error"
            End Try
        End If
        SerialPortVXM.WriteLine("D")
        LJUD.ePut(u3.ljhandle, LJUD.IO.PUT_DAC, 1, 0, 0)   ' Set DAC1 To 0 volts.
        ButtonStop.Enabled = False
    End Sub

    Private Sub ButtonStop_Click(sender As Object, e As EventArgs) Handles ButtonStop.Click
        ButtonStop.Enabled = False
        msDelay(10)
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
        VXMBatchTop = False
        LabelYEdge.Text = ""
        ButtonVXMEdge.BackColor = Color.WhiteSmoke
        ButtonLXYZ.PerformClick()
    End Sub
End Class
