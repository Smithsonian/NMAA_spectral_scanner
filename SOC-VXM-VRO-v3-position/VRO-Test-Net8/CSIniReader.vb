Imports System.Runtime.InteropServices
Imports System.Text

Public Class CSIniReader
    <DllImport("kernel32", CharSet:=CharSet.Auto)>
    Private Shared Function GetPrivateProfileString(
        lpAppName As String,
        lpKeyName As String,
        lpDefault As String,
        lpReturnedString As StringBuilder,
        nSize As Integer,
        lpFileName As String) As Integer
    End Function

    Public Shared Function ReadIniValue(section As String, key As String, filePath As String) As String
        Dim sb As New StringBuilder(255)
        GetPrivateProfileString(section, key, "", sb, sb.Capacity, filePath)
        Return sb.ToString()
    End Function

End Class
