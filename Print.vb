Imports Zebra.Sdk.Printer.Discovery
Imports Zebra.Sdk.Printer
Imports Zebra.Sdk.Comm
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Linq ' need to add 

Public Class Print
    Function discoveryPrinter() As DiscoveredUsbPrinter
        Try
            Dim symbolUsb As DiscoveredUsbPrinter = Nothing

            Form1.ComboBoxP.Items.Clear()
            For Each usb As DiscoveredUsbPrinter In UsbDiscoverer.GetZebraUsbPrinters(New ZebraPrinterFilter())
                Form1.ComboBoxP.Items.Add(usb)
                Form1.ComboBoxP.SelectedIndex = 0

                symbolUsb = usb
            Next
            Return symbolUsb
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Sub discoveryPrinterDriver()
        Try
            Dim symbolUsb As DiscoveredPrinterDriver = Nothing

            Form1.ComboBoxP.Items.Clear()
            For Each driver As DiscoveredPrinterDriver In UsbDiscoverer.GetZebraDriverPrinters
                Console.WriteLine("Driver Zebra : " + driver.ToString)
                Form1.ComboBoxP.Items.Add(driver)
                Form1.ComboBoxP.SelectedIndex = 0
            Next
            '            Return symbolUsb
        Catch ex As Exception
            '           Return Nothing
        End Try
    End Sub

    Function GetTypeLabel(namaLabel As String) As String
        Dim AccesDb As AccesDb = New AccesDb()
        Dim DtQuery As DataTable
        Dim Query As String = $"SELECT KomponenLabel FROM TB_DesignLabel WHERE ID = '{namaLabel}'"
        DtQuery = AccesDb.QueryDLtable(Query)
        Dim re As String = DtQuery.Rows(0).Item(0)
        Return re
    End Function
    Public Sub actionPrint(SN As String, Model As String, namaFile As String, index As Integer, x As Integer, y As Integer, Printer As String)
        '5A9N322111700001+JV40230-00013
        Dim koneksi As New DriverPrinterConnection(Printer)
        If Printer = "" Then
            'Return
        Else
            Console.WriteLine("verbose2")
            koneksi.Open()
            If koneksi.Connected = False Then
                Return
            End If
        End If

        Dim typeLabel() As String = GetTypeLabel(namaFile).Split("#")

        Dim dataZPL As String = ""
        Dim dataZPLDummy As String
        Dim formatModel As inputZPL = New inputZPL(namaFile)
        Dim TempSn As String = ""
        Console.WriteLine("SN : " + SN)
        For i As Integer = 1 To index
            TempSn = CountSerial(SN, i)
            dataZPLDummy = formatModel.regexSoundbar(TempSn, Model, x, y, typeLabel)
            dataZPL += dataZPLDummy
        Next
        Debug.WriteLine(dataZPL)
        Try
            koneksi.Write(System.Text.Encoding.UTF8.GetBytes(dataZPL))
        Catch ex As Exception
            Console.WriteLine("Printer Error")
        End Try
        koneksi.Close()

    End Sub
    Function CountSerial(SN As String, i As String) As String
        '5A9N322111700001+JV40230-00013

        Console.WriteLine(SN)
        Dim firstSN As Integer = CInt(SN.Substring(10, 5))
        Dim Serial1 As String = SN.Substring(0, 10)
        Dim Serial2 As String = SN.Split("+"c)(1)
        Dim Count As Integer = 0
        Count = firstSN + (i - 1)

        Console.WriteLine("FULL SERIAL1: " + Serial1)
        Console.WriteLine("FULL SERIAL2: " + Serial2)
        Console.WriteLine("FULL SERIAL3: " + Count.ToString)

        Dim FullSerial As String = Serial1 + Count.ToString("D5") + "+" + Serial2
        Console.WriteLine("FULL SERIAL: " + FullSerial)
        Return FullSerial
    End Function

End Class
