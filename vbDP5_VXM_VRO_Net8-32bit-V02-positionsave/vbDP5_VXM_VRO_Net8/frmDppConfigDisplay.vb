Public Class frmDppConfigDisplay
	Inherits System.Windows.Forms.Form

	Public m_strMessage As String
	Public m_strTitle As String
	Public m_strDelimiter As String

	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Me.Close()
	End Sub

	Private Sub frmDppConfigDisplay_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		_txtCfg_0.BackColor = Me.BackColor
		_txtCfg_1.BackColor = Me.BackColor
		If (Len(Trim(m_strTitle)) > 0) Then
			Me.Text = m_strTitle
		End If
		If (Len(Trim(m_strDelimiter)) = 0) Then m_strDelimiter = vbCrLf
		SetMessage(m_strMessage, m_strDelimiter)
	End Sub

	Private Sub SetMessage(ByRef cstrMsg As String, ByRef cstrDelimiter As String)
		Dim iHalf As Integer
		Dim iDelimiter As Integer ' \r\n ms win
		If (Len(cstrMsg) = 0) Then Exit Sub
		If (Len(cstrMsg) < 100) Then
			_txtCfg_0.Text = cstrMsg
			_txtCfg_1.Text = ""
		Else
			iHalf = CInt(Len(cstrMsg) / 2)
			iDelimiter = InStr(iHalf, cstrMsg, cstrDelimiter)
			If (iDelimiter >= iHalf) Then ' try to find next \r\n
				_txtCfg_0.Text = cstrMsg.Substring(0, iDelimiter + cstrDelimiter.Length() - 1)
				_txtCfg_1.Text = Mid(cstrMsg, iDelimiter + cstrDelimiter.Length())
			End If
		End If
	End Sub
End Class