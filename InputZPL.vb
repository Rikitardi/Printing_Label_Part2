Imports System.IO
Imports System.Text.RegularExpressions
Public Class inputZPL

    'Contoh Akses       "Dim ZPL As inputZPL = New inputZPL(CBmodel.Text)"
    'RegEx Data Label   "ZPL.regexData(Datalabel)"

    Private ReadOnly _namaFile As String
    Public Sub New(ByVal model As String)
        If model IsNot Nothing Then
            _namaFile = model
            Dim unused = _readFormatFile
        Else
            Throw New ArgumentNullException(NameOf(model))
        End If
    End Sub
    Private ReadOnly Property _readFormatFile As String
        Get
            Try
                Using reader As New StreamReader($"format\{_namaFile}.zpf")
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
        End Get
    End Property

    Function regexSoundbarFuck(ByVal SN_Data As String, ByVal model As String, x As Integer, y As Integer) As String
        Try
            Dim SerialBarcode1, SerialBarcode2, SerialBarcode3, SerialBarcode4, SN_PC, SN, PC As String

            SN_PC = SN_Data
            SN = SN_Data.Split("+"c)(0)
            PC = SN_Data.Split("+"c)(1)

            'JsonData = 

            Dim dataSeri1 As String = "@T1@"
            Dim dataSeri2 As String = "@T2@"
            Dim dataSeri3 As String = "@T3@"
            Dim dataSeri4 As String = "@Q1@"

            Dim dataOutput = Regex.Replace(_readFormatFile, dataSeri1, SerialBarcode1)
            dataOutput = Regex.Replace(dataOutput, dataSeri2, SerialBarcode2)
            dataOutput = Regex.Replace(dataOutput, dataSeri3, SerialBarcode3)
            dataOutput = Regex.Replace(dataOutput, dataSeri4, SerialBarcode4)

            Dim nilaiX As Integer = Convert.ToInt32(showMatch(dataOutput, "!(.*)!").Replace("!", ""))
            Dim nilaiY As Integer = Convert.ToInt32(showMatch(dataOutput, "#(.*)#").Replace("#", ""))
            Debug.WriteLine("nilai x : " + nilaiX.ToString)
            Debug.WriteLine("nilai y : " + nilaiY.ToString)

            Dim regexCode1 As String = "!"
            Dim str1 = nilaiX.ToString
            For i = 0 To str1.Length - 1
                Dim a As Char = Str(i)
                regexCode1 += "."
            Next
            regexCode1 += "!"

            Debug.WriteLine("regCode1 : " + regexCode1.ToString)
            nilaiX += x

            dataOutput = Regex.Replace(dataOutput, regexCode1, nilaiX.ToString)

            Dim regexCode2 As String = "#"
            Dim str2 = nilaiY.ToString
            For i = 0 To str2.Length - 1
                Dim a As Char = Str(i)
                regexCode2 += "."
            Next
            regexCode2 += "#"

            Debug.WriteLine("regCode2 : " + regexCode2.ToString)
            nilaiY += y
            dataOutput = Regex.Replace(dataOutput, regexCode2, nilaiY.ToString)
            Return dataOutput
        Catch ex As Exception
        End Try
    End Function

    Function regexSoundbar(ByVal SN As String, ByVal model As String, x As Integer, y As Integer, indexID() As String) As String
        Try
            Dim SerialBarcode1, SerialBarcode2, SerialBarcode3, dataOutput, SerialBarcode4, StringData As String
            Dim TName As String = ""
            Dim TData As String = ""
            StringData = ""
            dataOutput = ""

            SerialBarcode1 = model
            SerialBarcode2 = SN
            SerialBarcode3 = SN.Split("+"c)(0)
            SerialBarcode4 = SN.Split("+"c)(1)

            Console.WriteLine("TOTAL NUMBER : " + SerialBarcode1 + " " + SerialBarcode2 + " " + SerialBarcode3 + " " + SerialBarcode4)

            For k As Integer = 0 To indexID.Count - 1
                TName = "@" + indexID(k).Split(":")(0) + "@"
                Console.WriteLine("TNAME :" + TName)
                TData = indexID(k).Split(":")(1)
                Select Case TData
                    Case "B" 'Q1:FULL_SN#T1:SN#T2:PC#T3:BUYER
                        StringData = SerialBarcode1
                        Console.WriteLine("Excellent!")
                    Case "FULL_SN"
                        StringData = SerialBarcode2
                        Console.WriteLine("Well done")
                    Case "SN"
                        StringData = SerialBarcode3
                        Console.WriteLine("You passed")
                    Case "PC"
                        StringData = SerialBarcode4
                        Console.WriteLine("You passed Away")
                End Select
                If k = 0 Then
                    dataOutput = Regex.Replace(_readFormatFile, TName, StringData)
                Else
                    dataOutput = Regex.Replace(dataOutput, TName, StringData)
                End If
            Next


            'Dim dataSeri1 As String = "@T1@"
            'Dim dataSeri2 As String = "@T2@"
            'Dim dataSeri3 As String = "@T3@"
            'Dim dataSeri4 As String = "@Q1@"


            'dataOutput = Regex.Replace(_readFormatFile, dataSeri1, SerialBarcode1)
            'dataOutput = Regex.Replace(dataOutput, dataSeri2, SerialBarcode2)
            'dataOutput = Regex.Replace(dataOutput, dataSeri3, SerialBarcode3)
            'dataOutput = Regex.Replace(dataOutput, dataSeri4, SerialBarcode4)

            Debug.WriteLine("SUCCES")
            Dim nilaiX As Integer = Convert.ToInt32(showMatch(dataOutput, "!(.*)!").Replace("!", ""))
            Dim nilaiY As Integer = Convert.ToInt32(showMatch(dataOutput, "#(.*)#").Replace("#", ""))
            Debug.WriteLine("nilai x : " + nilaiX.ToString)
            Debug.WriteLine("nilai y : " + nilaiY.ToString)



            Dim regexCode1 As String = "!"
            Dim str1 = nilaiX.ToString
            For i = 0 To str1.Length - 1
                Dim a As Char = Str(i)
                regexCode1 += "."
            Next
            regexCode1 += "!"

            Debug.WriteLine("regCode1 : " + regexCode1.ToString)
            nilaiX += x

            dataOutput = Regex.Replace(dataOutput, regexCode1, nilaiX.ToString)

            Dim regexCode2 As String = "#"
            Dim str2 = nilaiY.ToString
            For i = 0 To str2.Length - 1
                Dim a As Char = Str(i)
                regexCode2 += "."
            Next
            regexCode2 += "#"

            Debug.WriteLine("regCode2 : " + regexCode2.ToString)
            nilaiY += y
            dataOutput = Regex.Replace(dataOutput, regexCode2, nilaiY.ToString)
            Return dataOutput
        Catch ex As Exception
        End Try
    End Function

    Function regexDataTest(x As Integer, y As Integer) As String
        Dim nilaiX As Integer = Convert.ToInt32(showMatch(_readFormatFile, "#(.*)#").Split(",")(0).Replace("#", ""))
        Dim nilaiY As Integer = Convert.ToInt32(showMatch(_readFormatFile, "#(.*)#").Split(",")(1).Replace("#", ""))

        Dim regexCode As String = "#"
        For Each a As Char In nilaiX.ToString & "," & nilaiY.ToString
            regexCode += "."
        Next

        regexCode += "#"

        Console.WriteLine(regexCode)
        Console.WriteLine($"Koordinat Akhir :{nilaiX.ToString & "," & nilaiY.ToString}")

        Return Regex.Replace(_readFormatFile, regexCode, nilaiX.ToString & "," & nilaiY.ToString)

    End Function

    Function reworkStringCode(code As String) As String
        Dim nilai As String = code.Substring(3, 9)
        Return nilai
    End Function

    Function showMatch(ByVal text As String, ByVal expr As String) As String
        Try
            Dim m As Match
            For Each m In Regex.Matches(text, expr)
                Return m.ToString
            Next m
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
End Class

