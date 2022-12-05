Public Class Form2
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim p As New ParsingZPL
        p.Parsing(ComboBox1.Text)
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim files() As String = IO.Directory.GetFiles(".\format\")

        For Each file As String In files
            If Not file.Split("\"c)(2).IndexOf(".sjn") = -1 Then
                ComboBox1.Items.Add(file.Split("\"c)(2))
            End If

            ' Do work, example
            'Dim text As String = IO.File.ReadAllText(file)

        Next
    End Sub
End Class