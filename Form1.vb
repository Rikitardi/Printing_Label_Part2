Imports System.Drawing.Printing
Imports System.Management
Imports Microsoft.Win32
Public Class Form1

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles ButtonPrint.Click
        If RadToggleSwitchMode.Value = True Then
            SAMPLE_Print()
        Else
            SAMPLE_RePrint()
        End If

    End Sub

    Function GET_LabelType(NameModel As String) As String
        Dim query As String = $"SELECT Label, Buyer FROM TB_Model WHERE NamaModel = '{NameModel}'"
        Dim dt As DataTable = DOWNLOAD_Database(query)
        Console.WriteLine("LABEL : " + dt.Rows(0).Item(0))
        Dim re As String = dt.Rows(0).Item(0) + "#" + dt.Rows(0).Item(1)
        Return re
    End Function

    Function GET_LabelName(NameLabel As String) As String
        Dim query As String = $"SELECT LabelType FROM TB_Label WHERE NameLabel = '{NameLabel}'"
        Dim dt As DataTable = DOWNLOAD_Database(query)
        Console.WriteLine("LABEL : " + dt.Rows(0).Item(0))
        Dim re As String = dt.Rows(0).Item(0)
        Return re
    End Function

    Sub SAMPLE_SN_RePrint()
        Dim Model As String = "5A9N"
        Dim partNo As String = "JV51003-0002I"
        If TBqtyInput.Text = "" Then
            LabelStartSN.Text = "_______________"
            LabelEndSN.Text = "_______________"
            LabelPartCode.Text = "_____________"
            Return
        End If

        Dim initSN As Integer = TextBoxStartSN.Text
        Dim lastSN As Integer = 0
        Dim tanggal As Date = DateTimePicker1.Value
        Dim Qty As Integer = TBqtyInput.Text
        Dim dt As DataTable

        Dim SNFull() As String

        Dim SNSearch As String = Model + tanggal.ToString("yy") + tanggal.ToString("MM") + tanggal.ToString("dd")
        Dim PCBtype As String = ComboBoxLabelType.Text
        Dim Query As String = $"SELECT LastSerial FROM TB_LogSN WHERE SerialNumber LIKE '{SNSearch}%' AND LabelType = '{PCBtype}'"
        dt = DOWNLOAD_Database(Query)
        Console.WriteLine("ROWS : " + dt.Rows.Count.ToString)
        If dt.Rows.Count < 1 Then
            MessageBox.Show("Reprint Error!" + vbCrLf + "Serial Number Not Found", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        Else

            For Each row As DataRow In dt.Rows
                lastSN = row(0)
            Next

            Dim TempSN As Integer = initSN + Qty

            If TempSN > lastSN Then
                MessageBox.Show("Reprint Error!" + vbCrLf + "Serial Number Beyond Limit", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TBqtyInput.Text = ""
                Return
            End If

            lastSN = initSN
        End If

        Dim SP As SerialProcces = New SerialProcces
        SNFull = SP.GetSerial(Model, tanggal, lastSN, partNo, Qty).Split("#")

        LabelStartSN.Text = SNFull(0).Split("+")(0)
        LabelEndSN.Text = SNFull(1).Split("+")(0)
        LabelPartCode.Text = partNo
    End Sub

    Sub SAMPLE_RePrint()
        Dim TryPrint As New Print()
        Dim SN As String = LabelStartSN.Text + "+" + LabelPartCode.Text
        Dim lastSN As String = LabelEndSN.Text.Substring(0, 10)
        Dim Buyer As String = GET_LabelType(ComboBoxModel.Text).Split("#")(1)
        Dim SeriAkhir As Integer = CInt(LabelEndSN.Text.Substring(10, 5))
        Dim NameFile As String = GET_LabelName(ComboBoxLabelType.Text).Split("#")(0)
        Dim Index As Integer = TBqtyInput.Text
        Dim KoorX As Integer = NumericUpDownX.Value
        Dim KoorY As Integer = NumericUpDownY.Value
        Dim NamePrinter As String = ComboBoxP.Text

        TryPrint.discoveryPrinterDriver()
        TryPrint.actionPrint(SN, Buyer, NameFile, Index, KoorX, KoorY, NamePrinter)

        Dim DataLog As String = $"Re-Print {ComboBoxModel.Text} SN {LabelStartSN.Text} To {LabelEndSN.Text} QTY {TBqtyInput.Text} Ea"
        SAMPLE_log(DataLog)

        'Dim QueryUP As String = $"INSERT INTO TB_LogSN (SerialNumber,LastSerial) VALUES ('{lastSN}', {SeriAkhir})"
        'Console.WriteLine(QueryUP)
        'UPLOAD_Database(QueryUP)
    End Sub
    Sub SAMPLE_Print()
        Dim TryPrint As New Print()
        Dim SN As String = LabelStartSN.Text + "+" + LabelPartCode.Text
        Dim lastSN As String = LabelEndSN.Text.Substring(0, 10)
        Dim Buyer As String = GET_LabelType(ComboBoxModel.Text).Split("#")(1)
        Dim SeriAkhir As Integer = CInt(LabelEndSN.Text.Substring(10, 5))
        Dim NameFile As String = GET_LabelName(ComboBoxLabelType.Text).Split("#")(0)
        Dim Index As Integer = TBqtyInput.Text
        Dim KoorX As Integer = NumericUpDownX.Value
        Dim KoorY As Integer = NumericUpDownY.Value
        Dim NamePrinter As String = ComboBoxP.Text
        Dim labelType As String = ComboBoxLabelType.Text

        TryPrint.discoveryPrinterDriver()
        TryPrint.actionPrint(SN, Buyer, NameFile, Index, KoorX, KoorY, NamePrinter)

        Dim QueryUP As String = $"INSERT INTO TB_LogSN (SerialNumber,LabelType,LastSerial) VALUES ('{lastSN}','{labelType}', {SeriAkhir})"
        Console.WriteLine(QueryUP)
        UPLOAD_Database(QueryUP)
        Dim DataLog As String = $"Print {ComboBoxModel.Text} SN {LabelStartSN.Text} To {LabelEndSN.Text} QTY {TBqtyInput.Text} Ea"
        SAMPLE_log(DataLog)


    End Sub
    Function DOWNLOAD_Database(query As String) As DataTable
        Dim AccesDb As AccesDb = New AccesDb()
        Dim DtQuery As DataTable
        'Dim Query As String = TextBox1.Text
        DtQuery = AccesDb.QueryDLtable(query)
        Return DtQuery
    End Function

    Sub UPLOAD_Database(query As String)
        Dim AccesDb As AccesDb = New AccesDb()
        AccesDb.QueryUL(query)
    End Sub

    Sub SAMPLE_SN()
        Dim Model As String = "5A9N"
        Dim partNo As String = "JV51003-0002I"
        If TBqtyInput.Text = "" Then
            LabelStartSN.Text = "_______________"
            LabelEndSN.Text = "_______________"
            LabelPartCode.Text = "_____________"
            Return
        End If


        Dim lastSN As Integer = 0
        Dim tanggal As Date = DateTimePicker1.Value
        Dim Qty As Integer = TBqtyInput.Text
        Dim dt As DataTable

        Dim SNFull() As String

        Dim SNSearch As String = Model + tanggal.ToString("yy") + tanggal.ToString("MM") + tanggal.ToString("dd")
        Dim PCBtype As String = ComboBoxLabelType.Text
        Dim Query As String = $"SELECT LastSerial FROM TB_LogSN WHERE SerialNumber LIKE '{SNSearch}%' AND LabelType = '{PCBtype}'"
        dt = DOWNLOAD_Database(Query)
        Console.WriteLine("ROWS : " + dt.Rows.Count.ToString)
        If dt.Rows.Count < 1 Then
            Dim QueryUP As String = $"INSERT INTO TB_LogSN (SerialNumber,LabelType,LastSerial) VALUES ('{SNSearch}','{PCBtype}', 1)"
            Console.WriteLine(QueryUP)
            UPLOAD_Database(QueryUP)
        Else
            For Each row As DataRow In dt.Rows
                lastSN = row(0)
            Next
        End If

        Dim SP As SerialProcces = New SerialProcces
        SNFull = SP.GetSerial(Model, tanggal, lastSN, partNo, Qty).Split("#")

        LabelStartSN.Text = SNFull(0).Split("+")(0)
        LabelEndSN.Text = SNFull(1).Split("+")(0)
        LabelPartCode.Text = partNo
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitProg()
    End Sub

    Private Sub InitProg()
        Dim TryPrint As New Print()
        TryPrint.discoveryPrinterDriver()
        DateTimePicker1.Value = Date.Now
        TextBoxStartSN.Enabled = False
        GetModel()
        ReadyStart()
    End Sub
    Private Sub GetModel()
        Dim Query As String = $"SELECT DISTINCT NamaModel FROM TB_Model WHERE 1"
        Dim dt As DataTable
        dt = DOWNLOAD_Database(Query)
        For i As Integer = 0 To dt.Rows.Count - 1
            ComboBoxModel.Items.Add(dt.Rows(0).Item(i))
        Next
    End Sub

    Private Sub GetPCB()
        Dim Query As String = $"SELECT Label FROM TB_Model WHERE NamaModel = '{ComboBoxModel.Text}'"
        Dim dt As DataTable
        dt = DOWNLOAD_Database(Query)
        For i As Integer = 0 To dt.Rows.Count - 1
            ComboBoxLabelType.Items.Add(dt.Rows(i).Item(0))
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'SAMPLE_Database()
        'SAMPLE_SN_RePrint()
        'SAMPLE_Print()
        'SAMPLE_log()
        'UsbSystem()
    End Sub

    Private Sub RadToggleSwitch1_ValueChanged(sender As Object, e As EventArgs) Handles RadToggleSwitchMode.ValueChanged
        If RadToggleSwitchMode.Value = False Then
            Panel1.BackColor = Color.Yellow
            TextBoxStartSN.Enabled = True
        Else
            TextBoxStartSN.Enabled = False
            Panel1.BackColor = Color.White
            TextBoxStartSN.Text = ""
        End If
    End Sub

    Public Sub SAMPLE_log(Data As String)
        Dim logData As LogData = New LogData
        Dim logValue As String = logData.saveLog(Data)
        ListBox1.Items.Add(logValue)
    End Sub

    Public Sub ReadyStart()
        ButtonPrint.Enabled = False
        If ComboBoxModel.Text = "" Then
            Return
        End If

        If ComboBoxLabelType.Text = "" Then
            Return
        End If

        Console.WriteLine("Nilai : " + TBqtyInput.Text)

        If TBqtyInput.Text = "" Then
            Return
        End If

        If RadToggleSwitchMode.Value = False Then
            If TextBoxStartSN.Text = "" Then
                Return
            End If
        End If

        ButtonPrint.Enabled = True
    End Sub

    Private Sub ComboBoxModel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxModel.SelectedIndexChanged
        GetPCB()
        ReadyStart()
    End Sub

    Private Sub ComboBoxPcb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxLabelType.SelectedIndexChanged
        ReadyStart()
    End Sub

    Private Sub TBqtyInput_TextChanged(sender As Object, e As EventArgs) Handles TBqtyInput.TextChanged
        ReadyStart()
        If RadToggleSwitchMode.Value = True Then
            SAMPLE_SN()
        Else
            SAMPLE_SN_RePrint()
        End If
    End Sub

    Private Sub TBqtyInput_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TBqtyInput.KeyPress
        If e.KeyChar <> ChrW(Keys.Back) Then
            If Char.IsNumber(e.KeyChar) Then
            Else
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub TextBoxStartSN_TextChanged(sender As Object, e As EventArgs) Handles TextBoxStartSN.TextChanged
        ReadyStart()
    End Sub
End Class