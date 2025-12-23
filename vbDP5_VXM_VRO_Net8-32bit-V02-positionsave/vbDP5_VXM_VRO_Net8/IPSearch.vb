Option Strict Off
Option Explicit On
Friend Class frmDP5_Connect
    Inherits System.Windows.Forms.Form
    Private Sub frmDP5_Connect_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        NetfinderUnit = 0
        InitAllEntries(DppEntries)
        LoadApplicationSettings(Me, True, False, False)
        'strCommStatus = ""
        strCommStatus = "USB"
        lblDppFound.BackColor = System.Drawing.ColorTranslator.FromOle(colorLightSteelBlue)
        lblDppFound.ForeColor = System.Drawing.ColorTranslator.FromOle(colorLightSlateGray)
        lblDppFound.Visible = False
        isDppFound = False
        cmdCountDevices_Click(cmdCountDevices, New System.EventArgs())
        s.CurrentInterface = USB
        cmdFindDevice_Click(CurrentUSBDevice, New System.EventArgs())
    End Sub

    Private Sub cmdCountDevices_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCountDevices.Click
        NumUSBDevices = CountDP5WinusbDevices()
        lblNumUSBDevices.Text = CStr(NumUSBDevices)
        If (NumUSBDevices > 1) Then
            If (CurrentUSBDevice > NumUSBDevices) Then
                CurrentUSBDevice = NumUSBDevices
                udSelectDevice.Value = CurrentUSBDevice
            End If
            udSelectDevice.Maximum = NumUSBDevices
            udSelectDevice.Enabled = True
        Else 'disable spin and device selection
            udSelectDevice.Enabled = False
            CurrentUSBDevice = 1
            udSelectDevice.Value = 1
            udSelectDevice.Maximum = 1
        End If
    End Sub



    Private Sub cmdCloseDevice_Click(sender As Object, e As EventArgs) Handles cmdCloseDevice.Click
        If (USBDeviceConnected) Then
            CloseDeviceHandle(DppWinUSB)
            USBDeviceConnected = False
        End If
    End Sub

    Public Sub cmdFindDevice_Click(sender As Object, e As EventArgs) Handles cmdFindDevice.Click
        Dim isDetected As Boolean
        On Error GoTo cmdFindDevice_ClickErr 'moving to later position
        cmdFindDevice.Enabled = False
        cmdCloseDevice_Click(cmdCloseDevice, New System.EventArgs())    'Uses DppWinUSB defined in DppWinUSB
        isDetected = OpenDevice(DppWinUSB, USBDeviceConnected, USBDevicePathName, CurrentUSBDevice - 1)
        cmdFindDevice.Enabled = True
        If (USBDeviceConnected) Then
            lblDppFound.Visible = True
            isDppFound = True
            s.CurrentInterface = USB
            USBDeviceTest = True
            frmDP5.cmdUSBDeviceTest()
        Else
            lblDppFound.Visible = False
            isDppFound = False
        End If
        'On Error GoTo cmdFindDevice_ClickErr 'moved here
        Exit Sub
cmdFindDevice_ClickErr:
        Call ProcessError(Err)
    End Sub


    '    Private Sub udSelectDevice_ValueChanged(sender As Object, e As EventArgs) Handles udSelectDevice.ValueChanged
    Private Sub udSelectDevice_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles udSelectDevice.ValueChanged
        CurrentUSBDevice = udSelectDevice.Value
        cmdFindDevice_Click(CurrentUSBDevice, New System.EventArgs())
    End Sub

    'Is this still needed?
    'Private Sub frmDP5_Connect_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
    '   System.Windows.Forms.Application.DoEvents()
    'End Sub


    Private Sub frmDP5_Connect_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        FinalizeDPPConnect()
    End Sub

    Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click
        FinalizeDPPConnect()
        Me.Close()
    End Sub

    Private Sub FinalizeDPPConnect()
        Dim strSelectedComm As String
        strSelectedComm = "Select Communications"
        Select Case s.CurrentInterface 'if found display...
            Case USB 'display WinUSB
                If (isDppFound) Then
                    strSelectedComm = "USB - " & "WinUSB"
                    s.isDppConnected = True
                End If
            Case Else
                s.isDppConnected = False
        End Select
        SaveApplicationSettings(Me, True, False, True)
    End Sub

End Class