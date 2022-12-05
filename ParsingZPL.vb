Imports System.IO
Imports System.Text.RegularExpressions
Public Class ParsingZPL
    Public Sub Parsing(Loc As String)
        Dim dummyText As String = GetFileMaster(Loc)
        Dim pattern As String = "(?<=\@)(.*?)(?=\@)"
        'Dim rgx As Regex = New Regex()
        showMatch(dummyText, pattern)
        '        Dim match As String = Regex.Matches(dummyText, )

    End Sub

    Sub showMatch(ByVal text As String, ByVal expr As String)
        Console.WriteLine("The Expression: " + expr)
        Console.WriteLine("TEXT: " + text)
        Dim mc As MatchCollection = Regex.Matches(text, expr)
        Dim m As Match
        For Each m In mc
            Console.WriteLine(m)
        Next m
    End Sub

    Function GetFileMaster(namafile As String) As String
        Try
            Using reader As New StreamReader($"format\{namafile}")
                Dim a = String.Empty
                Dim b = String.Empty
                Do
                    a = reader.ReadLine
                    b = b & vbCrLf & a
                Loop Until a Is Nothing
                reader.Close()
                Return b
            End Using
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
            'msg box here ...
            Return Nothing
        End Try

    End Function
End Class
