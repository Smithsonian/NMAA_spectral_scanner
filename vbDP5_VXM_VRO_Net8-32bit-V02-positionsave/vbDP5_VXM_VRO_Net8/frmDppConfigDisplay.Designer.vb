<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDppConfigDisplay
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        _txtCfg_0 = New TextBox()
        _txtCfg_1 = New TextBox()
        cmdOK = New Button()
        SuspendLayout()
        ' 
        ' _txtCfg_0
        ' 
        _txtCfg_0.Location = New Point(16, 10)
        _txtCfg_0.Margin = New Padding(2)
        _txtCfg_0.Multiline = True
        _txtCfg_0.Name = "_txtCfg_0"
        _txtCfg_0.Size = New Size(314, 438)
        _txtCfg_0.TabIndex = 0
        ' 
        ' _txtCfg_1
        ' 
        _txtCfg_1.Location = New Point(345, 10)
        _txtCfg_1.Margin = New Padding(2)
        _txtCfg_1.Multiline = True
        _txtCfg_1.Name = "_txtCfg_1"
        _txtCfg_1.Size = New Size(303, 438)
        _txtCfg_1.TabIndex = 1
        ' 
        ' cmdOK
        ' 
        cmdOK.Location = New Point(294, 452)
        cmdOK.Margin = New Padding(2)
        cmdOK.Name = "cmdOK"
        cmdOK.Size = New Size(78, 20)
        cmdOK.TabIndex = 2
        cmdOK.Text = "OK"
        cmdOK.UseVisualStyleBackColor = True
        ' 
        ' frmDppConfigDisplay
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(676, 483)
        Controls.Add(cmdOK)
        Controls.Add(_txtCfg_1)
        Controls.Add(_txtCfg_0)
        Margin = New Padding(2)
        Name = "frmDppConfigDisplay"
        Text = "DPP Configuration"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents _txtCfg_0 As TextBox
    Friend WithEvents _txtCfg_1 As TextBox
    Friend WithEvents cmdOK As Button
End Class
