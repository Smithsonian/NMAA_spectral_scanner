<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmDP5
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
        components = New ComponentModel.Container()
        cmdSelectCommunications = New Button()
        lblSelectCommunications = New Label()
        Timer2 = New Timer(components)
        pnlStatus = New StatusStrip()
        picMCAPlot = New PictureBox()
        lblVersion = New Label()
        cmdStartAcquisition = New Button()
        cmdStopAcquisition = New Button()
        cmdSingleUpdate = New Button()
        cmdClearData = New Button()
        chkDeltaMode = New CheckBox()
        CheckLog = New CheckBox()
        CheckLine = New CheckBox()
        _lblStatusDisplay_0 = New Label()
        _lblStatusDisplay_1 = New Label()
        cmdShowConfiguration = New Button()
        cmdSaveSpectrum = New Button()
        cmdSetFolder = New Button()
        lblBatchDir = New Label()
        lblBatchBase = New Label()
        txtPresetTime = New TextBox()
        lbl_ms = New Label()
        cmdBatchStart = New Button()
        cmdBatchStop = New Button()
        TextBoxVXMinc = New TextBox()
        lbl_mm = New Label()
        txtSpectrumTag = New TextBox()
        txtSpectrumDescription = New TextBox()
        lbl_tag = New Label()
        lbl_description = New Label()
        LabelBacthTime = New Label()
        LabelEX = New Label()
        LabelStatus = New Label()
        LabelVXMOut = New Label()
        GroupBoxConnections = New GroupBox()
        ButtonXRayOFF = New Button()
        ButtonDSD = New Button()
        ButtonVXMConnect = New Button()
        ButtonVXMjog = New Button()
        GroupBox1 = New GroupBox()
        GroupBox2 = New GroupBox()
        ButtonXYZLoc = New Button()
        ButtonVROsZero = New Button()
        ButtonVXMStatus = New Button()
        ButtonVXMZero = New Button()
        ButtonVXMTopCorner = New Button()
        ButtonVXMBottomCorner = New Button()
        ButtonVXMCalcBatch = New Button()
        GroupBox3 = New GroupBox()
        ButtonClearCorners = New Button()
        GroupBox4 = New GroupBox()
        LabelTCVRO = New Label()
        LabelBCVRO = New Label()
        CType(picMCAPlot, ComponentModel.ISupportInitialize).BeginInit()
        GroupBoxConnections.SuspendLayout()
        GroupBox1.SuspendLayout()
        GroupBox2.SuspendLayout()
        GroupBox3.SuspendLayout()
        GroupBox4.SuspendLayout()
        SuspendLayout()
        ' 
        ' cmdSelectCommunications
        ' 
        cmdSelectCommunications.BackColor = Color.LightCoral
        cmdSelectCommunications.Location = New Point(19, 22)
        cmdSelectCommunications.Name = "cmdSelectCommunications"
        cmdSelectCommunications.Size = New Size(75, 23)
        cmdSelectCommunications.TabIndex = 0
        cmdSelectCommunications.Text = "DP5 SSD"
        cmdSelectCommunications.UseVisualStyleBackColor = False
        ' 
        ' lblSelectCommunications
        ' 
        lblSelectCommunications.AutoSize = True
        lblSelectCommunications.Location = New Point(100, 26)
        lblSelectCommunications.Name = "lblSelectCommunications"
        lblSelectCommunications.Size = New Size(123, 15)
        lblSelectCommunications.TabIndex = 1
        lblSelectCommunications.Text = "DP5 Communications"
        ' 
        ' Timer2
        ' 
        Timer2.Interval = 1000
        ' 
        ' pnlStatus
        ' 
        pnlStatus.ImageScalingSize = New Size(24, 24)
        pnlStatus.Location = New Point(0, 614)
        pnlStatus.Name = "pnlStatus"
        pnlStatus.Size = New Size(972, 22)
        pnlStatus.TabIndex = 2
        pnlStatus.Text = "Status"
        ' 
        ' picMCAPlot
        ' 
        picMCAPlot.BackColor = Color.White
        picMCAPlot.Location = New Point(18, 12)
        picMCAPlot.Name = "picMCAPlot"
        picMCAPlot.Size = New Size(435, 250)
        picMCAPlot.TabIndex = 3
        picMCAPlot.TabStop = False
        ' 
        ' lblVersion
        ' 
        lblVersion.Location = New Point(575, 590)
        lblVersion.Name = "lblVersion"
        lblVersion.Size = New Size(388, 19)
        lblVersion.TabIndex = 4
        lblVersion.Text = "Version label"
        lblVersion.TextAlign = ContentAlignment.MiddleRight
        ' 
        ' cmdStartAcquisition
        ' 
        cmdStartAcquisition.Location = New Point(12, 25)
        cmdStartAcquisition.Name = "cmdStartAcquisition"
        cmdStartAcquisition.Size = New Size(94, 23)
        cmdStartAcquisition.TabIndex = 5
        cmdStartAcquisition.Text = "1 - Start"
        cmdStartAcquisition.UseVisualStyleBackColor = True
        ' 
        ' cmdStopAcquisition
        ' 
        cmdStopAcquisition.Location = New Point(12, 55)
        cmdStopAcquisition.Name = "cmdStopAcquisition"
        cmdStopAcquisition.Size = New Size(94, 23)
        cmdStopAcquisition.TabIndex = 6
        cmdStopAcquisition.Text = "1 - STOP"
        cmdStopAcquisition.UseVisualStyleBackColor = True
        ' 
        ' cmdSingleUpdate
        ' 
        cmdSingleUpdate.Location = New Point(12, 84)
        cmdSingleUpdate.Name = "cmdSingleUpdate"
        cmdSingleUpdate.Size = New Size(94, 23)
        cmdSingleUpdate.TabIndex = 7
        cmdSingleUpdate.Text = "Single Update"
        cmdSingleUpdate.UseVisualStyleBackColor = True
        ' 
        ' cmdClearData
        ' 
        cmdClearData.Location = New Point(12, 113)
        cmdClearData.Name = "cmdClearData"
        cmdClearData.Size = New Size(94, 23)
        cmdClearData.TabIndex = 8
        cmdClearData.Text = "Clear Data"
        cmdClearData.UseVisualStyleBackColor = True
        ' 
        ' chkDeltaMode
        ' 
        chkDeltaMode.AutoSize = True
        chkDeltaMode.Location = New Point(19, 142)
        chkDeltaMode.Name = "chkDeltaMode"
        chkDeltaMode.Size = New Size(87, 19)
        chkDeltaMode.TabIndex = 9
        chkDeltaMode.Text = "Delta Mode"
        chkDeltaMode.UseVisualStyleBackColor = True
        ' 
        ' CheckLog
        ' 
        CheckLog.AutoSize = True
        CheckLog.Checked = True
        CheckLog.CheckState = CheckState.Checked
        CheckLog.Location = New Point(459, 216)
        CheckLog.Name = "CheckLog"
        CheckLog.Size = New Size(58, 19)
        CheckLog.TabIndex = 10
        CheckLog.Text = "Linear"
        CheckLog.UseVisualStyleBackColor = True
        ' 
        ' CheckLine
        ' 
        CheckLine.AutoSize = True
        CheckLine.Location = New Point(459, 241)
        CheckLine.Name = "CheckLine"
        CheckLine.Size = New Size(48, 19)
        CheckLine.TabIndex = 11
        CheckLine.Text = "Line"
        CheckLine.UseVisualStyleBackColor = True
        ' 
        ' _lblStatusDisplay_0
        ' 
        _lblStatusDisplay_0.AutoSize = True
        _lblStatusDisplay_0.Location = New Point(6, 17)
        _lblStatusDisplay_0.Name = "_lblStatusDisplay_0"
        _lblStatusDisplay_0.Size = New Size(67, 15)
        _lblStatusDisplay_0.TabIndex = 12
        _lblStatusDisplay_0.Text = "dp5 status1"
        ' 
        ' _lblStatusDisplay_1
        ' 
        _lblStatusDisplay_1.Location = New Point(136, 17)
        _lblStatusDisplay_1.Name = "_lblStatusDisplay_1"
        _lblStatusDisplay_1.Size = New Size(120, 243)
        _lblStatusDisplay_1.TabIndex = 13
        _lblStatusDisplay_1.Text = "dp5 status2"
        ' 
        ' cmdShowConfiguration
        ' 
        cmdShowConfiguration.Location = New Point(582, 212)
        cmdShowConfiguration.Name = "cmdShowConfiguration"
        cmdShowConfiguration.Size = New Size(113, 23)
        cmdShowConfiguration.TabIndex = 14
        cmdShowConfiguration.Text = "Show dp5 config"
        cmdShowConfiguration.UseVisualStyleBackColor = True
        ' 
        ' cmdSaveSpectrum
        ' 
        cmdSaveSpectrum.Location = New Point(127, 26)
        cmdSaveSpectrum.Name = "cmdSaveSpectrum"
        cmdSaveSpectrum.Size = New Size(99, 23)
        cmdSaveSpectrum.TabIndex = 15
        cmdSaveSpectrum.Text = "Save Spectrum"
        cmdSaveSpectrum.UseVisualStyleBackColor = True
        ' 
        ' cmdSetFolder
        ' 
        cmdSetFolder.Location = New Point(12, 23)
        cmdSetFolder.Name = "cmdSetFolder"
        cmdSetFolder.Size = New Size(141, 23)
        cmdSetFolder.TabIndex = 16
        cmdSetFolder.Text = "Scan Name, Folder"
        cmdSetFolder.UseVisualStyleBackColor = True
        ' 
        ' lblBatchDir
        ' 
        lblBatchDir.Location = New Point(12, 121)
        lblBatchDir.Name = "lblBatchDir"
        lblBatchDir.Size = New Size(402, 29)
        lblBatchDir.TabIndex = 17
        lblBatchDir.Text = "Current Scan Directory"
        ' 
        ' lblBatchBase
        ' 
        lblBatchBase.Location = New Point(12, 153)
        lblBatchBase.Name = "lblBatchBase"
        lblBatchBase.Size = New Size(402, 45)
        lblBatchBase.TabIndex = 18
        lblBatchBase.Text = "Scan Base Name"
        ' 
        ' txtPresetTime
        ' 
        txtPresetTime.Location = New Point(12, 86)
        txtPresetTime.Name = "txtPresetTime"
        txtPresetTime.Size = New Size(61, 23)
        txtPresetTime.TabIndex = 19
        txtPresetTime.Text = "200"
        txtPresetTime.TextAlign = HorizontalAlignment.Right
        ' 
        ' lbl_ms
        ' 
        lbl_ms.AutoSize = True
        lbl_ms.Location = New Point(77, 91)
        lbl_ms.Name = "lbl_ms"
        lbl_ms.Size = New Size(72, 15)
        lbl_ms.TabIndex = 20
        lbl_ms.Text = "ms real time"
        ' 
        ' cmdBatchStart
        ' 
        cmdBatchStart.BackColor = Color.LightGreen
        cmdBatchStart.Enabled = False
        cmdBatchStart.Location = New Point(482, 108)
        cmdBatchStart.Name = "cmdBatchStart"
        cmdBatchStart.Size = New Size(75, 30)
        cmdBatchStart.TabIndex = 21
        cmdBatchStart.Text = "Start Scan"
        cmdBatchStart.UseVisualStyleBackColor = False
        ' 
        ' cmdBatchStop
        ' 
        cmdBatchStop.BackColor = Color.Salmon
        cmdBatchStop.Enabled = False
        cmdBatchStop.Location = New Point(482, 147)
        cmdBatchStop.Name = "cmdBatchStop"
        cmdBatchStop.Size = New Size(75, 30)
        cmdBatchStop.TabIndex = 22
        cmdBatchStop.Text = "Stop Scan"
        cmdBatchStop.UseVisualStyleBackColor = False
        ' 
        ' TextBoxVXMinc
        ' 
        TextBoxVXMinc.Location = New Point(12, 57)
        TextBoxVXMinc.Name = "TextBoxVXMinc"
        TextBoxVXMinc.Size = New Size(61, 23)
        TextBoxVXMinc.TabIndex = 23
        TextBoxVXMinc.Text = "1"
        TextBoxVXMinc.TextAlign = HorizontalAlignment.Right
        ' 
        ' lbl_mm
        ' 
        lbl_mm.AutoSize = True
        lbl_mm.Location = New Point(76, 65)
        lbl_mm.Name = "lbl_mm"
        lbl_mm.Size = New Size(73, 15)
        lbl_mm.TabIndex = 24
        lbl_mm.Text = "mm spacing"
        ' 
        ' txtSpectrumTag
        ' 
        txtSpectrumTag.Location = New Point(126, 73)
        txtSpectrumTag.Name = "txtSpectrumTag"
        txtSpectrumTag.Size = New Size(100, 23)
        txtSpectrumTag.TabIndex = 25
        txtSpectrumTag.Text = "live_data"
        ' 
        ' txtSpectrumDescription
        ' 
        txtSpectrumDescription.Location = New Point(126, 117)
        txtSpectrumDescription.Name = "txtSpectrumDescription"
        txtSpectrumDescription.Size = New Size(100, 23)
        txtSpectrumDescription.TabIndex = 26
        txtSpectrumDescription.Text = "spectrum data"
        ' 
        ' lbl_tag
        ' 
        lbl_tag.AutoSize = True
        lbl_tag.Location = New Point(159, 55)
        lbl_tag.Name = "lbl_tag"
        lbl_tag.Size = New Size(29, 15)
        lbl_tag.TabIndex = 27
        lbl_tag.Text = "Tag:"
        ' 
        ' lbl_description
        ' 
        lbl_description.AutoSize = True
        lbl_description.Location = New Point(142, 99)
        lbl_description.Name = "lbl_description"
        lbl_description.Size = New Size(70, 15)
        lbl_description.TabIndex = 28
        lbl_description.Text = "Description:"
        ' 
        ' LabelBacthTime
        ' 
        LabelBacthTime.Location = New Point(601, 518)
        LabelBacthTime.Name = "LabelBacthTime"
        LabelBacthTime.Size = New Size(352, 59)
        LabelBacthTime.TabIndex = 29
        LabelBacthTime.Text = "Batch Time Info"
        ' 
        ' LabelEX
        ' 
        LabelEX.Location = New Point(601, 446)
        LabelEX.Name = "LabelEX"
        LabelEX.Size = New Size(352, 63)
        LabelEX.TabIndex = 30
        LabelEX.Text = "Encoder Info"
        ' 
        ' LabelStatus
        ' 
        LabelStatus.Location = New Point(601, 291)
        LabelStatus.Name = "LabelStatus"
        LabelStatus.Size = New Size(338, 38)
        LabelStatus.TabIndex = 31
        LabelStatus.Text = "VXM VRO Status"
        ' 
        ' LabelVXMOut
        ' 
        LabelVXMOut.Location = New Point(601, 332)
        LabelVXMOut.Name = "LabelVXMOut"
        LabelVXMOut.Size = New Size(352, 114)
        LabelVXMOut.TabIndex = 32
        LabelVXMOut.Text = "VXM Out Info"
        ' 
        ' GroupBoxConnections
        ' 
        GroupBoxConnections.Controls.Add(ButtonXRayOFF)
        GroupBoxConnections.Controls.Add(ButtonDSD)
        GroupBoxConnections.Controls.Add(ButtonVXMConnect)
        GroupBoxConnections.Controls.Add(cmdSelectCommunications)
        GroupBoxConnections.Controls.Add(lblSelectCommunications)
        GroupBoxConnections.Location = New Point(20, 273)
        GroupBoxConnections.Name = "GroupBoxConnections"
        GroupBoxConnections.Size = New Size(295, 125)
        GroupBoxConnections.TabIndex = 33
        GroupBoxConnections.TabStop = False
        GroupBoxConnections.Text = "Connections"
        ' 
        ' ButtonXRayOFF
        ' 
        ButtonXRayOFF.BackColor = Color.Orchid
        ButtonXRayOFF.Location = New Point(218, 74)
        ButtonXRayOFF.Name = "ButtonXRayOFF"
        ButtonXRayOFF.Size = New Size(63, 38)
        ButtonXRayOFF.TabIndex = 4
        ButtonXRayOFF.Text = "X-Ray OFF"
        ButtonXRayOFF.UseVisualStyleBackColor = False
        ' 
        ' ButtonDSD
        ' 
        ButtonDSD.BackColor = Color.LightCoral
        ButtonDSD.Location = New Point(19, 89)
        ButtonDSD.Name = "ButtonDSD"
        ButtonDSD.Size = New Size(75, 23)
        ButtonDSD.TabIndex = 3
        ButtonDSD.Text = "DSD"
        ButtonDSD.UseVisualStyleBackColor = False
        ' 
        ' ButtonVXMConnect
        ' 
        ButtonVXMConnect.BackColor = Color.LightCoral
        ButtonVXMConnect.Location = New Point(19, 55)
        ButtonVXMConnect.Name = "ButtonVXMConnect"
        ButtonVXMConnect.Size = New Size(75, 23)
        ButtonVXMConnect.TabIndex = 2
        ButtonVXMConnect.Text = "VXM, VROs"
        ButtonVXMConnect.UseVisualStyleBackColor = False
        ' 
        ' ButtonVXMjog
        ' 
        ButtonVXMjog.Location = New Point(13, 22)
        ButtonVXMjog.Name = "ButtonVXMjog"
        ButtonVXMjog.Size = New Size(75, 23)
        ButtonVXMjog.TabIndex = 34
        ButtonVXMjog.Text = "Jog Enable"
        ButtonVXMjog.UseVisualStyleBackColor = True
        ' 
        ' GroupBox1
        ' 
        GroupBox1.Controls.Add(lbl_description)
        GroupBox1.Controls.Add(txtSpectrumTag)
        GroupBox1.Controls.Add(cmdStartAcquisition)
        GroupBox1.Controls.Add(lbl_tag)
        GroupBox1.Controls.Add(cmdStopAcquisition)
        GroupBox1.Controls.Add(cmdSingleUpdate)
        GroupBox1.Controls.Add(txtSpectrumDescription)
        GroupBox1.Controls.Add(cmdClearData)
        GroupBox1.Controls.Add(chkDeltaMode)
        GroupBox1.Controls.Add(cmdSaveSpectrum)
        GroupBox1.Location = New Point(459, 12)
        GroupBox1.Name = "GroupBox1"
        GroupBox1.Size = New Size(236, 183)
        GroupBox1.TabIndex = 35
        GroupBox1.TabStop = False
        GroupBox1.Text = "Single Spectra"
        ' 
        ' GroupBox2
        ' 
        GroupBox2.Controls.Add(ButtonXYZLoc)
        GroupBox2.Controls.Add(ButtonVROsZero)
        GroupBox2.Controls.Add(ButtonVXMStatus)
        GroupBox2.Controls.Add(ButtonVXMZero)
        GroupBox2.Controls.Add(ButtonVXMjog)
        GroupBox2.Location = New Point(323, 273)
        GroupBox2.Name = "GroupBox2"
        GroupBox2.Size = New Size(272, 125)
        GroupBox2.TabIndex = 36
        GroupBox2.TabStop = False
        GroupBox2.Text = "VXM, VROs"
        ' 
        ' ButtonXYZLoc
        ' 
        ButtonXYZLoc.Location = New Point(103, 89)
        ButtonXYZLoc.Name = "ButtonXYZLoc"
        ButtonXYZLoc.Size = New Size(75, 23)
        ButtonXYZLoc.TabIndex = 37
        ButtonXYZLoc.Text = "Position"
        ButtonXYZLoc.UseVisualStyleBackColor = True
        ' 
        ' ButtonVROsZero
        ' 
        ButtonVROsZero.Location = New Point(103, 51)
        ButtonVROsZero.Name = "ButtonVROsZero"
        ButtonVROsZero.Size = New Size(75, 23)
        ButtonVROsZero.TabIndex = 36
        ButtonVROsZero.Text = "VROs Zero"
        ButtonVROsZero.UseVisualStyleBackColor = True
        ' 
        ' ButtonVXMStatus
        ' 
        ButtonVXMStatus.Location = New Point(13, 55)
        ButtonVXMStatus.Name = "ButtonVXMStatus"
        ButtonVXMStatus.Size = New Size(75, 23)
        ButtonVXMStatus.TabIndex = 35
        ButtonVXMStatus.Text = "VXM Status"
        ButtonVXMStatus.UseVisualStyleBackColor = True
        ' 
        ' ButtonVXMZero
        ' 
        ButtonVXMZero.Location = New Point(103, 22)
        ButtonVXMZero.Name = "ButtonVXMZero"
        ButtonVXMZero.Size = New Size(75, 23)
        ButtonVXMZero.TabIndex = 0
        ButtonVXMZero.Text = "VXM Zero"
        ButtonVXMZero.UseVisualStyleBackColor = True
        ' 
        ' ButtonVXMTopCorner
        ' 
        ButtonVXMTopCorner.Location = New Point(198, 23)
        ButtonVXMTopCorner.Name = "ButtonVXMTopCorner"
        ButtonVXMTopCorner.Size = New Size(114, 23)
        ButtonVXMTopCorner.TabIndex = 37
        ButtonVXMTopCorner.Text = "Set Top Corner"
        ButtonVXMTopCorner.UseVisualStyleBackColor = True
        ' 
        ' ButtonVXMBottomCorner
        ' 
        ButtonVXMBottomCorner.Location = New Point(198, 52)
        ButtonVXMBottomCorner.Name = "ButtonVXMBottomCorner"
        ButtonVXMBottomCorner.Size = New Size(114, 23)
        ButtonVXMBottomCorner.TabIndex = 38
        ButtonVXMBottomCorner.Text = "Set Bottom Corner"
        ButtonVXMBottomCorner.UseVisualStyleBackColor = True
        ' 
        ' ButtonVXMCalcBatch
        ' 
        ButtonVXMCalcBatch.Location = New Point(220, 81)
        ButtonVXMCalcBatch.Name = "ButtonVXMCalcBatch"
        ButtonVXMCalcBatch.Size = New Size(75, 23)
        ButtonVXMCalcBatch.TabIndex = 39
        ButtonVXMCalcBatch.Text = "Calculate"
        ButtonVXMCalcBatch.UseVisualStyleBackColor = True
        ' 
        ' GroupBox3
        ' 
        GroupBox3.Controls.Add(LabelBCVRO)
        GroupBox3.Controls.Add(LabelTCVRO)
        GroupBox3.Controls.Add(ButtonClearCorners)
        GroupBox3.Controls.Add(ButtonVXMCalcBatch)
        GroupBox3.Controls.Add(ButtonVXMBottomCorner)
        GroupBox3.Controls.Add(ButtonVXMTopCorner)
        GroupBox3.Controls.Add(lbl_mm)
        GroupBox3.Controls.Add(TextBoxVXMinc)
        GroupBox3.Controls.Add(lbl_ms)
        GroupBox3.Controls.Add(txtPresetTime)
        GroupBox3.Controls.Add(cmdSetFolder)
        GroupBox3.Controls.Add(lblBatchBase)
        GroupBox3.Controls.Add(cmdBatchStart)
        GroupBox3.Controls.Add(lblBatchDir)
        GroupBox3.Controls.Add(cmdBatchStop)
        GroupBox3.Location = New Point(18, 410)
        GroupBox3.Name = "GroupBox3"
        GroupBox3.Size = New Size(577, 201)
        GroupBox3.TabIndex = 40
        GroupBox3.TabStop = False
        GroupBox3.Text = "Scan Settings"
        ' 
        ' ButtonClearCorners
        ' 
        ButtonClearCorners.Font = New Font("Segoe UI", 7.8F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        ButtonClearCorners.Location = New Point(471, 23)
        ButtonClearCorners.Name = "ButtonClearCorners"
        ButtonClearCorners.Size = New Size(86, 23)
        ButtonClearCorners.TabIndex = 40
        ButtonClearCorners.Text = "Clear Corners"
        ButtonClearCorners.UseVisualStyleBackColor = True
        ' 
        ' GroupBox4
        ' 
        GroupBox4.BackColor = SystemColors.GradientActiveCaption
        GroupBox4.Controls.Add(_lblStatusDisplay_0)
        GroupBox4.Controls.Add(_lblStatusDisplay_1)
        GroupBox4.Location = New Point(701, 12)
        GroupBox4.Name = "GroupBox4"
        GroupBox4.Size = New Size(266, 263)
        GroupBox4.TabIndex = 41
        GroupBox4.TabStop = False
        GroupBox4.Text = "DP5 Status"
        ' 
        ' LabelTCVRO
        ' 
        LabelTCVRO.AutoSize = True
        LabelTCVRO.Location = New Point(317, 26)
        LabelTCVRO.Name = "LabelTCVRO"
        LabelTCVRO.Size = New Size(47, 15)
        LabelTCVRO.TabIndex = 41
        LabelTCVRO.Text = "Top X Y"
        ' 
        ' LabelBCVRO
        ' 
        LabelBCVRO.AutoSize = True
        LabelBCVRO.Location = New Point(318, 56)
        LabelBCVRO.Name = "LabelBCVRO"
        LabelBCVRO.Size = New Size(67, 15)
        LabelBCVRO.TabIndex = 42
        LabelBCVRO.Text = "Bottom X Y"
        ' 
        ' frmDP5
        ' 
        AutoScaleDimensions = New SizeF(7.0F, 15.0F)
        AutoScaleMode = AutoScaleMode.Font
        AutoSize = True
        ClientSize = New Size(972, 636)
        Controls.Add(GroupBox4)
        Controls.Add(GroupBox3)
        Controls.Add(GroupBox2)
        Controls.Add(GroupBox1)
        Controls.Add(GroupBoxConnections)
        Controls.Add(LabelVXMOut)
        Controls.Add(LabelStatus)
        Controls.Add(LabelEX)
        Controls.Add(LabelBacthTime)
        Controls.Add(CheckLog)
        Controls.Add(CheckLine)
        Controls.Add(cmdShowConfiguration)
        Controls.Add(lblVersion)
        Controls.Add(picMCAPlot)
        Controls.Add(pnlStatus)
        Name = "frmDP5"
        StartPosition = FormStartPosition.CenterScreen
        Text = "DP5 VXM VRO"
        CType(picMCAPlot, ComponentModel.ISupportInitialize).EndInit()
        GroupBoxConnections.ResumeLayout(False)
        GroupBoxConnections.PerformLayout()
        GroupBox1.ResumeLayout(False)
        GroupBox1.PerformLayout()
        GroupBox2.ResumeLayout(False)
        GroupBox3.ResumeLayout(False)
        GroupBox3.PerformLayout()
        GroupBox4.ResumeLayout(False)
        GroupBox4.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Public WithEvents cmdSelectCommunications As Button
    Friend WithEvents lblSelectCommunications As Label
    Public WithEvents Timer2 As Timer
    Friend WithEvents pnlStatus As StatusStrip
    Public WithEvents picMCAPlot As PictureBox
    Public WithEvents lblVersion As Label
    Public WithEvents cmdStartAcquisition As Button
    Public WithEvents cmdStopAcquisition As Button
    Public WithEvents cmdSingleUpdate As Button
    Public WithEvents cmdClearData As Button
    Friend WithEvents chkDeltaMode As CheckBox
    Public WithEvents CheckLog As CheckBox
    Public WithEvents CheckLine As CheckBox
    Public WithEvents _lblStatusDisplay_0 As Label
    Friend WithEvents _lblStatusDisplay_1 As Label
    Public WithEvents cmdShowConfiguration As Button
    Public WithEvents cmdSaveSpectrum As Button
    Public WithEvents cmdSetFolder As Button
    Friend WithEvents lblBatchDir As Label
    Friend WithEvents lblBatchBase As Label
    Friend WithEvents txtPresetTime As TextBox
    Friend WithEvents lbl_ms As Label
    Friend WithEvents cmdBatchStart As Button
    Friend WithEvents cmdBatchStop As Button
    Friend WithEvents TextBoxVXMinc As TextBox
    Friend WithEvents lbl_mm As Label
    Friend WithEvents txtSpectrumTag As TextBox
    Friend WithEvents txtSpectrumDescription As TextBox
    Friend WithEvents lbl_tag As Label
    Friend WithEvents lbl_description As Label
    Friend WithEvents LabelBacthTime As Label
    Friend WithEvents LabelEX As Label
    Friend WithEvents LabelStatus As Label
    Friend WithEvents LabelVXMOut As Label
    Friend WithEvents GroupBoxConnections As GroupBox
    Friend WithEvents ButtonXRayOFF As Button
    Friend WithEvents ButtonDSD As Button
    Friend WithEvents ButtonVXMConnect As Button
    Friend WithEvents ButtonVXMjog As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents ButtonVXMZero As Button
    Friend WithEvents ButtonVXMStatus As Button
    Friend WithEvents ButtonVXMTopCorner As Button
    Friend WithEvents ButtonVXMBottomCorner As Button
    Friend WithEvents ButtonVXMCalcBatch As Button
    Friend WithEvents ButtonVROsZero As Button
    Friend WithEvents ButtonXYZLoc As Button
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents ButtonClearCorners As Button
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents LabelBCVRO As Label
    Friend WithEvents LabelTCVRO As Label

End Class
