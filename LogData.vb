Imports System.IO
Imports System.IO.File

Public Class LogData
    Function saveLog(ValueLog As String) As String
        Dim filename As String = Application.StartupPath & $"\Log\{Date.Today.ToString("yyMMdd")}.log"
        Dim sw As StreamWriter = AppendText(filename)
        Dim l As String = Now() & ": " & ValueLog
        sw.WriteLine(Now() & ": " & ValueLog)
        sw.Close()
        Return l
    End Function
End Class
