Public Class SerialProcces
    Public Function GetSerial(ModelCode As String, Tanggal As Date, LastSn As Integer, PartNo As String, qty As Integer) As String
        Dim FullSnStart As String = ""
        Dim FullSnEnd As String = ""
        Dim FullSn As String = ""
        Dim SumSN As Integer = 0
        FullSnStart = ModelCode + Tanggal.ToString("yy") + Tanggal.ToString("MM") + Tanggal.ToString("dd") + LastSn.ToString("D5") + "+" + PartNo
        SumSN = LastSn + qty
        FullSnEnd = ModelCode + Tanggal.ToString("yy") + Tanggal.ToString("MM") + Tanggal.ToString("dd") + SumSN.ToString("D5") + "+" + PartNo
        FullSn = FullSnStart + "#" + FullSnEnd
        Return FullSn
    End Function

End Class
