Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Module CommonModule
    Public CurrentDirectory As String = Directory.GetCurrentDirectory()
    <DllImport("kernel32")>
    Public Function GetPrivateProfileString(section As String, key As String, def As String, retVal As StringBuilder, size As Integer, filePath As String) As Integer
    End Function
    Public Function GetIniValue(section As String, key As String, filename As String, Optional defaultValue As String = "") As String
        Dim sb As New StringBuilder(500)
        If GetPrivateProfileString(section, key, defaultValue, sb, sb.Capacity, filename) > 0 Then
            Return sb.ToString
        Else
            Return defaultValue
        End If
    End Function
    Public Sub InformationAlert(Message As String)
        MsgBox(Message, MsgBoxStyle.Information)
    End Sub

    Public Sub ErrorAlert(Message As String)
        MsgBox(Message, MsgBoxStyle.Critical)
    End Sub
End Module
