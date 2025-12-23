<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
        ButtonVRO2 = New Button()
        LabelStatus = New Label()
        ButtonLXYZ = New Button()
        LabelInfo = New Label()
        ButtonOpenPorts = New Button()
        ButtonLJconnect = New Button()
        LabelEncoder = New Label()
        GroupBoxConnect = New GroupBox()
        TextBoxCycles = New TextBox()
        TextBoxIntms = New TextBox()
        TextBoxFOV = New TextBox()
        LabelIterations = New Label()
        Labelms = New Label()
        LabelFOV = New Label()
        ButtonCalculate = New Button()
        ButtonStart = New Button()
        ButtonStop = New Button()
        GroupBoxMove = New GroupBox()
        LabelMoveEdge = New Label()
        LabelMove80 = New Label()
        ButtonMYedge = New Button()
        ButtonMXupfar = New Button()
        LabelMove10 = New Label()
        ButtonMYdn = New Button()
        ButtonMXup = New Button()
        ButtonMXdn = New Button()
        ButtonMYup = New Button()
        Label1 = New Label()
        LabelYEdge = New Label()
        ButtonVXMEdge = New Button()
        ButtonVXMjog = New Button()
        ButtonVROsZero = New Button()
        GroupBoxConnect.SuspendLayout()
        GroupBoxMove.SuspendLayout()
        SuspendLayout()
        ' 
        ' ButtonVRO2
        ' 
        ButtonVRO2.BackColor = Color.LightCoral
        ButtonVRO2.Location = New Point(24, 22)
        ButtonVRO2.Name = "ButtonVRO2"
        ButtonVRO2.Size = New Size(75, 23)
        ButtonVRO2.TabIndex = 0
        ButtonVRO2.Text = "VXM VROs"
        ButtonVRO2.UseVisualStyleBackColor = False
        ' 
        ' LabelStatus
        ' 
        LabelStatus.Location = New Point(295, 34)
        LabelStatus.Name = "LabelStatus"
        LabelStatus.Size = New Size(219, 57)
        LabelStatus.TabIndex = 1
        LabelStatus.Text = "LabelStatus"
        ' 
        ' ButtonLXYZ
        ' 
        ButtonLXYZ.Location = New Point(544, 34)
        ButtonLXYZ.Name = "ButtonLXYZ"
        ButtonLXYZ.Size = New Size(75, 23)
        ButtonLXYZ.TabIndex = 2
        ButtonLXYZ.Text = "Position"
        ButtonLXYZ.UseVisualStyleBackColor = True
        ' 
        ' LabelInfo
        ' 
        LabelInfo.Location = New Point(295, 121)
        LabelInfo.Name = "LabelInfo"
        LabelInfo.Size = New Size(219, 91)
        LabelInfo.TabIndex = 3
        LabelInfo.Text = "Info"
        ' 
        ' ButtonOpenPorts
        ' 
        ButtonOpenPorts.Location = New Point(654, 12)
        ButtonOpenPorts.Name = "ButtonOpenPorts"
        ButtonOpenPorts.Size = New Size(75, 23)
        ButtonOpenPorts.TabIndex = 6
        ButtonOpenPorts.Text = "Find Ports"
        ButtonOpenPorts.UseVisualStyleBackColor = True
        ' 
        ' ButtonLJconnect
        ' 
        ButtonLJconnect.BackColor = Color.LightCoral
        ButtonLJconnect.Location = New Point(24, 51)
        ButtonLJconnect.Name = "ButtonLJconnect"
        ButtonLJconnect.Size = New Size(75, 23)
        ButtonLJconnect.TabIndex = 7
        ButtonLJconnect.Text = "LabJack"
        ButtonLJconnect.UseVisualStyleBackColor = False
        ' 
        ' LabelEncoder
        ' 
        LabelEncoder.AutoSize = True
        LabelEncoder.Location = New Point(544, 16)
        LabelEncoder.Name = "LabelEncoder"
        LabelEncoder.Size = New Size(84, 15)
        LabelEncoder.TabIndex = 8
        LabelEncoder.Text = "Encoder check"
        ' 
        ' GroupBoxConnect
        ' 
        GroupBoxConnect.Controls.Add(ButtonVRO2)
        GroupBoxConnect.Controls.Add(ButtonLJconnect)
        GroupBoxConnect.Location = New Point(12, 12)
        GroupBoxConnect.Name = "GroupBoxConnect"
        GroupBoxConnect.Size = New Size(122, 85)
        GroupBoxConnect.TabIndex = 9
        GroupBoxConnect.TabStop = False
        GroupBoxConnect.Text = "Connect"
        ' 
        ' TextBoxCycles
        ' 
        TextBoxCycles.Location = New Point(21, 132)
        TextBoxCycles.Name = "TextBoxCycles"
        TextBoxCycles.Size = New Size(100, 23)
        TextBoxCycles.TabIndex = 10
        TextBoxCycles.Text = "1024"
        ' 
        ' TextBoxIntms
        ' 
        TextBoxIntms.Location = New Point(21, 169)
        TextBoxIntms.Name = "TextBoxIntms"
        TextBoxIntms.Size = New Size(100, 23)
        TextBoxIntms.TabIndex = 11
        TextBoxIntms.Text = "100"
        ' 
        ' TextBoxFOV
        ' 
        TextBoxFOV.Location = New Point(21, 206)
        TextBoxFOV.Name = "TextBoxFOV"
        TextBoxFOV.Size = New Size(100, 23)
        TextBoxFOV.TabIndex = 12
        TextBoxFOV.Text = "196"
        ' 
        ' LabelIterations
        ' 
        LabelIterations.AutoSize = True
        LabelIterations.Location = New Point(127, 140)
        LabelIterations.Name = "LabelIterations"
        LabelIterations.Size = New Size(85, 15)
        LabelIterations.TabIndex = 13
        LabelIterations.Text = "iterations/lines"
        ' 
        ' Labelms
        ' 
        Labelms.AutoSize = True
        Labelms.Location = New Point(127, 177)
        Labelms.Name = "Labelms"
        Labelms.Size = New Size(84, 15)
        Labelms.TabIndex = 14
        Labelms.Text = "ms integration"
        ' 
        ' LabelFOV
        ' 
        LabelFOV.AutoSize = True
        LabelFOV.Location = New Point(127, 214)
        LabelFOV.Name = "LabelFOV"
        LabelFOV.Size = New Size(54, 15)
        LabelFOV.TabIndex = 15
        LabelFOV.Text = "mm FOV"
        ' 
        ' ButtonCalculate
        ' 
        ButtonCalculate.Location = New Point(36, 249)
        ButtonCalculate.Name = "ButtonCalculate"
        ButtonCalculate.Size = New Size(75, 23)
        ButtonCalculate.TabIndex = 16
        ButtonCalculate.Text = "Calculate"
        ButtonCalculate.UseVisualStyleBackColor = True
        ' 
        ' ButtonStart
        ' 
        ButtonStart.BackColor = Color.FromArgb(CByte(192), CByte(255), CByte(192))
        ButtonStart.Enabled = False
        ButtonStart.Location = New Point(49, 347)
        ButtonStart.Name = "ButtonStart"
        ButtonStart.Size = New Size(75, 23)
        ButtonStart.TabIndex = 17
        ButtonStart.Text = "Start Scan"
        ButtonStart.UseVisualStyleBackColor = False
        ' 
        ' ButtonStop
        ' 
        ButtonStop.BackColor = Color.FromArgb(CByte(255), CByte(128), CByte(128))
        ButtonStop.Enabled = False
        ButtonStop.Location = New Point(49, 387)
        ButtonStop.Name = "ButtonStop"
        ButtonStop.Size = New Size(75, 23)
        ButtonStop.TabIndex = 18
        ButtonStop.Text = "Stop Scan"
        ButtonStop.UseVisualStyleBackColor = False
        ' 
        ' GroupBoxMove
        ' 
        GroupBoxMove.Controls.Add(LabelMoveEdge)
        GroupBoxMove.Controls.Add(LabelMove80)
        GroupBoxMove.Controls.Add(ButtonMYedge)
        GroupBoxMove.Controls.Add(ButtonMXupfar)
        GroupBoxMove.Controls.Add(LabelMove10)
        GroupBoxMove.Controls.Add(ButtonMYdn)
        GroupBoxMove.Controls.Add(ButtonMXup)
        GroupBoxMove.Controls.Add(ButtonMXdn)
        GroupBoxMove.Controls.Add(ButtonMYup)
        GroupBoxMove.Location = New Point(295, 279)
        GroupBoxMove.Name = "GroupBoxMove"
        GroupBoxMove.Size = New Size(225, 146)
        GroupBoxMove.TabIndex = 20
        GroupBoxMove.TabStop = False
        GroupBoxMove.Text = "Movement"
        ' 
        ' LabelMoveEdge
        ' 
        LabelMoveEdge.AutoSize = True
        LabelMoveEdge.Location = New Point(133, 99)
        LabelMoveEdge.Name = "LabelMoveEdge"
        LabelMoveEdge.Size = New Size(33, 15)
        LabelMoveEdge.TabIndex = 25
        LabelMoveEdge.Text = "Edge"
        ' 
        ' LabelMove80
        ' 
        LabelMove80.AutoSize = True
        LabelMove80.Location = New Point(133, 33)
        LabelMove80.Name = "LabelMove80"
        LabelMove80.Size = New Size(29, 15)
        LabelMove80.TabIndex = 24
        LabelMove80.Text = "80%"
        ' 
        ' ButtonMYedge
        ' 
        ButtonMYedge.Location = New Point(168, 95)
        ButtonMYedge.Name = "ButtonMYedge"
        ButtonMYedge.Size = New Size(38, 23)
        ButtonMYedge.TabIndex = 23
        ButtonMYedge.Text = "Y |↓|"
        ButtonMYedge.UseVisualStyleBackColor = True
        ' 
        ' ButtonMXupfar
        ' 
        ButtonMXupfar.Location = New Point(168, 29)
        ButtonMXupfar.Name = "ButtonMXupfar"
        ButtonMXupfar.Size = New Size(38, 23)
        ButtonMXupfar.TabIndex = 22
        ButtonMXupfar.Text = "X≡>"
        ButtonMXupfar.UseVisualStyleBackColor = True
        ' 
        ' LabelMove10
        ' 
        LabelMove10.AutoSize = True
        LabelMove10.Location = New Point(53, 62)
        LabelMove10.Name = "LabelMove10"
        LabelMove10.Size = New Size(29, 15)
        LabelMove10.TabIndex = 21
        LabelMove10.Text = "10%"
        ' 
        ' ButtonMYdn
        ' 
        ButtonMYdn.Location = New Point(53, 86)
        ButtonMYdn.Name = "ButtonMYdn"
        ButtonMYdn.Size = New Size(33, 23)
        ButtonMYdn.TabIndex = 3
        ButtonMYdn.Text = "↓Y"
        ButtonMYdn.UseVisualStyleBackColor = True
        ' 
        ' ButtonMXup
        ' 
        ButtonMXup.Location = New Point(90, 58)
        ButtonMXup.Name = "ButtonMXup"
        ButtonMXup.Size = New Size(33, 23)
        ButtonMXup.TabIndex = 2
        ButtonMXup.Text = "X→"
        ButtonMXup.UseVisualStyleBackColor = True
        ' 
        ' ButtonMXdn
        ' 
        ButtonMXdn.Location = New Point(16, 58)
        ButtonMXdn.Name = "ButtonMXdn"
        ButtonMXdn.Size = New Size(33, 23)
        ButtonMXdn.TabIndex = 1
        ButtonMXdn.Text = "←X"
        ButtonMXdn.UseVisualStyleBackColor = True
        ' 
        ' ButtonMYup
        ' 
        ButtonMYup.Location = New Point(53, 29)
        ButtonMYup.Name = "ButtonMYup"
        ButtonMYup.Size = New Size(33, 23)
        ButtonMYup.TabIndex = 0
        ButtonMYup.Text = "↑Y"
        ButtonMYup.UseVisualStyleBackColor = True
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(532, 431)
        Label1.Name = "Label1"
        Label1.Size = New Size(214, 15)
        Label1.TabIndex = 21
        Label1.Text = "National Museum of Asian Art, 2025-08"
        ' 
        ' LabelYEdge
        ' 
        LabelYEdge.AutoSize = True
        LabelYEdge.Location = New Point(394, 253)
        LabelYEdge.Name = "LabelYEdge"
        LabelYEdge.Size = New Size(68, 15)
        LabelYEdge.TabIndex = 22
        LabelYEdge.Text = "Y edge mm"
        ' 
        ' ButtonVXMEdge
        ' 
        ButtonVXMEdge.Location = New Point(302, 249)
        ButtonVXMEdge.Name = "ButtonVXMEdge"
        ButtonVXMEdge.Size = New Size(75, 23)
        ButtonVXMEdge.TabIndex = 23
        ButtonVXMEdge.Text = "Set Y Edge"
        ButtonVXMEdge.UseVisualStyleBackColor = True
        ' 
        ' ButtonVXMjog
        ' 
        ButtonVXMjog.Location = New Point(532, 249)
        ButtonVXMjog.Name = "ButtonVXMjog"
        ButtonVXMjog.Size = New Size(75, 23)
        ButtonVXMjog.TabIndex = 24
        ButtonVXMjog.Text = "Jog Enable"
        ButtonVXMjog.UseVisualStyleBackColor = True
        ' 
        ' ButtonVROsZero
        ' 
        ButtonVROsZero.Location = New Point(654, 65)
        ButtonVROsZero.Name = "ButtonVROsZero"
        ButtonVROsZero.Size = New Size(75, 23)
        ButtonVROsZero.TabIndex = 25
        ButtonVROsZero.Text = "VRO Zero"
        ButtonVROsZero.UseVisualStyleBackColor = True
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(758, 450)
        Controls.Add(ButtonVROsZero)
        Controls.Add(ButtonVXMjog)
        Controls.Add(ButtonVXMEdge)
        Controls.Add(LabelYEdge)
        Controls.Add(Label1)
        Controls.Add(GroupBoxMove)
        Controls.Add(ButtonStop)
        Controls.Add(ButtonStart)
        Controls.Add(ButtonCalculate)
        Controls.Add(LabelFOV)
        Controls.Add(Labelms)
        Controls.Add(LabelIterations)
        Controls.Add(TextBoxFOV)
        Controls.Add(TextBoxIntms)
        Controls.Add(TextBoxCycles)
        Controls.Add(GroupBoxConnect)
        Controls.Add(LabelEncoder)
        Controls.Add(ButtonOpenPorts)
        Controls.Add(LabelInfo)
        Controls.Add(ButtonLXYZ)
        Controls.Add(LabelStatus)
        Name = "Form1"
        Text = "Form1"
        GroupBoxConnect.ResumeLayout(False)
        GroupBoxMove.ResumeLayout(False)
        GroupBoxMove.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents ButtonVRO2 As Button
    Friend WithEvents LabelStatus As Label
    Friend WithEvents ButtonLXYZ As Button
    Friend WithEvents LabelInfo As Label
    Friend WithEvents ButtonOpenPorts As Button
    Friend WithEvents ButtonLJconnect As Button
    Friend WithEvents LabelEncoder As Label
    Friend WithEvents GroupBoxConnect As GroupBox
    Friend WithEvents TextBoxCycles As TextBox
    Friend WithEvents TextBoxIntms As TextBox
    Friend WithEvents TextBoxFOV As TextBox
    Friend WithEvents LabelIterations As Label
    Friend WithEvents Labelms As Label
    Friend WithEvents LabelFOV As Label
    Friend WithEvents ButtonCalculate As Button
    Friend WithEvents ButtonStart As Button
    Friend WithEvents ButtonStop As Button
    Friend WithEvents GroupBoxMove As GroupBox
    Friend WithEvents ButtonMYdn As Button
    Friend WithEvents ButtonMXup As Button
    Friend WithEvents ButtonMXdn As Button
    Friend WithEvents ButtonMYup As Button
    Friend WithEvents LabelMove10 As Label
    Friend WithEvents ButtonMXupfar As Button
    Friend WithEvents LabelMove80 As Label
    Friend WithEvents ButtonMYedge As Button
    Friend WithEvents LabelMoveEdge As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents LabelYEdge As Label
    Friend WithEvents ButtonVXMEdge As Button
    Friend WithEvents ButtonVXMjog As Button
    Friend WithEvents ButtonVROsZero As Button

End Class
