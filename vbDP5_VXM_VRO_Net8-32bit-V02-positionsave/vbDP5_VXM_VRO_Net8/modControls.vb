Option Strict Off
Option Explicit On

Module modControls

    Public Const ccNone As Integer = 0
    Public Const ccFixedSingle As Integer = 1

    Public Const cdlOFNFileMustExist As Integer = 4096
    Public Const cdlOFNHideReadOnly As Integer = 4
    Public Const cdlOFNExplorer As Integer = 524288
    Public Const cdlOFNLongNames As Integer = 2097152
    Public Const cdlOFNOverwritePrompt As Integer = 2

    ' From X11Color.h : header file
    Public Const colorAmptekBlue As Integer = &HFCCE7A
    Public Const colorBlue As Integer = &HFF0000
    Public Const colorGreen As Integer = &H8000
    Public Const colorLightGray As Integer = &HD3D3D3
    Public Const colorLightSlateGray As Integer = &H998877
    Public Const colorLightSteelBlue As Integer = &HDEC4B0
    Public Const colorRed As Integer = &HFF
    Public Const colorWhite As Integer = &HFFFFFF
    Public Const colorYellow As Integer = &HFFFF
    Public Const colorSilver As Integer = &HC0C0C0

    'dialog box file filters for file extensions
    Public Const dlgTXT_Filter As String = "Text File (*.txt)|*.txt|All Files (*.*)|*.*"
    Public Const dlgCFG_Filter As String = "Amptek Config File (*.cfg)|*.cfg|All Files (*.*)|*.*"
    Public Const dlgCSV_Filter As String = "Comma Separated Value (*.csv)|*.csv|All Files (*.*)|*.*"
    Public Const dlgMCA_Filter As String = "Amptek Spectrum File (*.mca)|*.mca|All Files (*.*)|*.*"
    Public Const dlgXL_Filter As String = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"

    Private Const cmdArrowWidth As Short = 285
    Private Const WS_EX_RIGHT As Short = &H1000S ' Alignment use extended sytle
    Private Const GWL_EXSTYLE As Short = (-20)
    Private Const GWL_STYLE As Short = (-16)
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Integer, ByVal nIndex As Integer) As Integer
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer

    '================================================================================
    ' Common Dialog Box Functions
    '
    ' This module contains the common dialog box functions and filters
    '================================================================================

    '================================================================================
    '==== Example Usage ====
    '================================================================================
    '    If (dlgOpen(cmnDlg, dlgMCA_Filter)) Then
    '        FileToOpen = cmnDlg.Filename
    '    Else
    '        'a file was not selected
    '    End If
    '
    '    If (dlgSave(cmnDlg, dlgCSV_Filter)) Then
    '        FileToSave = cmnDlg.Filename
    '    Else
    '        'a file was not selected
    '    End If
    '================================================================================

    'creates an open common dialog, returns true is a file was selected, false otherwise
    Public Function dlgOpen(ByRef CmnDlgFIO As System.Windows.Forms.OpenFileDialog, ByRef strFileFilters As String) As Boolean
        dlgOpen = False
        CmnDlgFIO.Filter = strFileFilters
        CmnDlgFIO.FilterIndex = 1
        CmnDlgFIO.CheckFileExists = True
        CmnDlgFIO.ShowReadOnly = False
        CmnDlgFIO.FileName = GetDefaultFileWildcard(strFileFilters)
        If (CmnDlgFIO.ShowDialog() = DialogResult.OK) Then
            dlgOpen = True
        End If
    End Function

    'creates a save common dialog, returns true if the file was saved, false otherwise
    Public Function dlgSave(ByRef CmnDlgFIO As System.Windows.Forms.SaveFileDialog, ByRef strFileFilters As String) As Boolean
        dlgSave = False
        CmnDlgFIO.Filter = strFileFilters
        CmnDlgFIO.FilterIndex = 1
        CmnDlgFIO.OverwritePrompt = True
        CmnDlgFIO.FileName = GetDefaultFileWildcard(strFileFilters)
        If (CmnDlgFIO.ShowDialog() = DialogResult.OK) Then
            dlgSave = True
        End If
    End Function

    Private Function GetDefaultFileWildcard(ByRef strFilter As String) As String
        Dim extBegin As Short
        Dim extEnd As Short
        Dim strExt As String

        GetDefaultFileWildcard = ""
        'find starting bar, read until next bar or end of string
        extBegin = InStr(strFilter, "|")
        If (extBegin < 1) Then Exit Function
        extEnd = InStr(extBegin + 1, strFilter, "|")
        If ((extEnd - 1) < extBegin) Then
            extEnd = Len(strFilter)
        Else
            extEnd = extEnd - 1
        End If
        strExt = Trim(Mid(strFilter, extBegin + 1, extEnd - extBegin))
        If (Len(strExt) < 1) Then Exit Function
        GetDefaultFileWildcard = strExt
    End Function

End Module
