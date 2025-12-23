<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmDP5_Connect
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        btnOK = New Button()
        cmdCountDevices = New Button()
        cmdFindDevice = New Button()
        cmdCloseDevice = New Button()
        Label1 = New Label()
        lblNumUSBDevices = New Label()
        udSelectDevice = New NumericUpDown()
        rtbStatus = New RichTextBox()
        lblDppFound = New Label()
        CType(udSelectDevice, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' btnOK
        ' 
        btnOK.Location = New Point(45, 119)
        btnOK.Name = "btnOK"
        btnOK.Size = New Size(75, 23)
        btnOK.TabIndex = 0
        btnOK.Text = "OK"
        btnOK.UseVisualStyleBackColor = True
        ' 
        ' cmdCountDevices
        ' 
        cmdCountDevices.Location = New Point(12, 27)
        cmdCountDevices.Name = "cmdCountDevices"
        cmdCountDevices.Size = New Size(95, 23)
        cmdCountDevices.TabIndex = 1
        cmdCountDevices.Text = "Count Devices"
        cmdCountDevices.UseVisualStyleBackColor = True
        ' 
        ' cmdFindDevice
        ' 
        cmdFindDevice.Location = New Point(12, 65)
        cmdFindDevice.Name = "cmdFindDevice"
        cmdFindDevice.Size = New Size(95, 23)
        cmdFindDevice.TabIndex = 2
        cmdFindDevice.Text = "Find Device"
        cmdFindDevice.UseVisualStyleBackColor = True
        ' 
        ' cmdCloseDevice
        ' 
        cmdCloseDevice.Location = New Point(335, 164)
        cmdCloseDevice.Name = "cmdCloseDevice"
        cmdCloseDevice.Size = New Size(88, 23)
        cmdCloseDevice.TabIndex = 3
        cmdCloseDevice.Text = "Close Device"
        cmdCloseDevice.UseVisualStyleBackColor = True
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(12, 9)
        Label1.Name = "Label1"
        Label1.Size = New Size(121, 15)
        Label1.TabIndex = 4
        Label1.Text = "USB Connection Only"
        ' 
        ' lblNumUSBDevices
        ' 
        lblNumUSBDevices.BorderStyle = BorderStyle.Fixed3D
        lblNumUSBDevices.Location = New Point(113, 27)
        lblNumUSBDevices.Name = "lblNumUSBDevices"
        lblNumUSBDevices.Size = New Size(25, 23)
        lblNumUSBDevices.TabIndex = 5
        lblNumUSBDevices.Text = "0"
        lblNumUSBDevices.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' udSelectDevice
        ' 
        udSelectDevice.Location = New Point(113, 65)
        udSelectDevice.Name = "udSelectDevice"
        udSelectDevice.Size = New Size(36, 23)
        udSelectDevice.TabIndex = 7
        udSelectDevice.Value = New Decimal(New Integer() {1, 0, 0, 0})
        ' 
        ' rtbStatus
        ' 
        rtbStatus.Location = New Point(169, 12)
        rtbStatus.Name = "rtbStatus"
        rtbStatus.Size = New Size(254, 146)
        rtbStatus.TabIndex = 8
        rtbStatus.Text = ""
        ' 
        ' lblDppFound
        ' 
        lblDppFound.BorderStyle = BorderStyle.Fixed3D
        lblDppFound.Location = New Point(32, 91)
        lblDppFound.Name = "lblDppFound"
        lblDppFound.Size = New Size(62, 23)
        lblDppFound.TabIndex = 9
        lblDppFound.Text = "FOUND"
        ' 
        ' frmDP5_Connect
        ' 
        AutoScaleDimensions = New SizeF(7.0F, 15.0F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(440, 196)
        Controls.Add(lblDppFound)
        Controls.Add(rtbStatus)
        Controls.Add(udSelectDevice)
        Controls.Add(lblNumUSBDevices)
        Controls.Add(Label1)
        Controls.Add(cmdCloseDevice)
        Controls.Add(cmdFindDevice)
        Controls.Add(cmdCountDevices)
        Controls.Add(btnOK)
        Name = "frmDP5_Connect"
        Text = "Search for USB DPP Devices"
        CType(udSelectDevice, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Public WithEvents btnOK As Button
    Friend WithEvents cmdCountDevices As Button
    Public WithEvents cmdFindDevice As Button
    Public WithEvents cmdCloseDevice As Button
    Friend WithEvents Label1 As Label
    Public WithEvents lblNumUSBDevices As Label
    Friend WithEvents udSelectDevice As NumericUpDown
    Public WithEvents rtbStatus As RichTextBox
    Friend WithEvents lblDppFound As Label
End Class
