Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlAutoFillType

Public Class Form63
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim DStartN As Date
    Dim DStartE As Date
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommander2.Connection = oConnection
                oCommander2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        DStartN = Me.TextBox1.Text & "/01/01"
        DStartE = DStartN.AddYears(1).AddDays(-1)
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Form63_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        Me.TextBox1.Text = Today.Year()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Label2.Text = "处理中"
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Label2.Text = "处理完毕"
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Payable_Report"
        SaveFileDialog1.DefaultExt = ".xlsx"
        Dim SON As DialogResult = SaveFileDialog1.ShowDialog()
        If SON = DialogResult.OK Then
            Dim SFN As String = SaveFileDialog1.FileName
            Ws.SaveAs(SFN, XlFileFormat.xlOpenXMLWorkbook)
        Else
            MsgBox("没有储存文件", MsgBoxStyle.Critical)
        End If
        xWorkBook.Saved = True
        xWorkBook.Close()
        xExcel.Quit()
        If oConnection.State = ConnectionState.Open Then
            Try
                oConnection.Close()
                Module1.KillExcelProcess(OldExcel)
                MsgBox("Finished")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        Ws.Name = "AP"
        AdjustExcelFormat()
        oCommand.CommandText = "SELECT distinct apa06,pmc03 FROM apa_file left join pmc_file on apa06 = pmc01 where apa02 between to_date('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += DStartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and apa41 = 'Y' and apa00 in ('11','12') order by apa06"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                oRng = Ws.Range("A" & LineZ, "A" & LineZ + 1)
                oRng.Merge()
                oRng = Ws.Range("B" & LineZ, "B" & LineZ + 1)
                oRng.Merge()
                Ws.Cells(LineZ, 3) = "入库请款"
                Ws.Cells(LineZ + 1, 3) = "杂项应付请款"
                'Ws.Cells(LineZ + 2, 3) = "厂商应付请款"
                Ws.Cells(LineZ, 16) = "=SUM(D" & LineZ & ":O" & LineZ
                Ws.Cells(LineZ + 1, 16) = "=SUM(D" & LineZ + 1 & ":O" & LineZ + 1
                'Ws.Cells(LineZ + 2, 16) = "=SUM(D" & LineZ + 2 & ":O" & LineZ + 2
                Ws.Cells(LineZ, 1) = oReader.Item("apa06")
                Ws.Cells(LineZ, 2) = oReader.Item("pmc03")
                GetPayment(oReader.Item("apa06"), 11)
                GetPayment(oReader.Item("apa06"), 12)
                'GetPayment(oReader.Item("apa06"), 15)
            End While
        End If
        oReader.Close()
        ' 加總
        oRng = Ws.Range("A" & LineZ, "C" & LineZ)
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(LineZ, 1) = "合计"
        Ws.Cells(LineZ, 4) = "=SUM(D3:D" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 4))
        oRng.AutoFill(Destination:=Ws.Range("D" & LineZ & ":P" & LineZ), Type:=xlFillDefault)
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "P1")
        oRng.EntireColumn.ColumnWidth = 12.5
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.ColumnWidth = 25
        oRng = Ws.Range("A1", "B1")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        oRng = Ws.Range("C1", "C2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        oRng.AutoFill(Destination:=Ws.Range("C1:P2"), Type:=xlFillDefault)
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormat = "@"
        Ws.Cells(1, 1) = "供应商"
        Ws.Cells(2, 1) = "供应厂商编号"
        Ws.Cells(2, 2) = "名称"
        Ws.Cells(1, 3) = "账款来源"
        For i As Int16 = 1 To 12
            If i < 10 Then
                Ws.Cells(1, 3 + i) = DStartN.Year & "-0" & i
            Else
                Ws.Cells(1, 3 + i) = DStartN.Year & "-" & i
            End If
        Next
        Ws.Cells(1, 16) = "合计"
        oRng = Ws.Range("D1", "P1")
        oRng.EntireColumn.NumberFormatLocal = "0.00"
        LineZ = 3
    End Sub
    Private Sub GetPayment(ByVal apa06 As String, apa00 As Decimal)
        oCommander2.CommandText = "select nvl(sum(t1),0) as t1,nvl(sum(t2),0) as t2,nvl(sum(t3),0) as t3,nvl(sum(t4),0) as t4,nvl(sum(t5),0) as t5,nvl(sum(t6),0) as t6,nvl(sum(t7),0) as t7,nvl(sum(t8),0) as t8,nvl(sum(t9),0) as t9,nvl(sum(t10),0) as t10,nvl(sum(t11),0) as t11,nvl(sum(t12),0) as t12 from ( "
        oCommander2.CommandText += "select (case when month(apa02) = 1 then sum(apa34) else 0 end)  as t1,(case when month(apa02) = 2 then sum(apa34) else 0 end)  as t2,"
        oCommander2.CommandText += "(case when month(apa02) = 3 then sum(apa34) else 0 end)  as t3,(case when month(apa02) = 4 then sum(apa34) else 0 end)  as t4,"
        oCommander2.CommandText += "(case when month(apa02) = 5 then sum(apa34) else 0 end)  as t5,(case when month(apa02) = 6 then sum(apa34) else 0 end)  as t6,"
        oCommander2.CommandText += "(case when month(apa02) = 7 then sum(apa34) else 0 end)  as t7,(case when month(apa02) = 8 then sum(apa34) else 0 end)  as t8,"
        oCommander2.CommandText += "(case when month(apa02) = 9 then sum(apa34) else 0 end)  as t9,(case when month(apa02) = 10 then sum(apa34) else 0 end)  as t10,"
        oCommander2.CommandText += "(case when month(apa02) = 11 then sum(apa34) else 0 end)  as t11,(case when month(apa02) = 12 then sum(apa34) else 0 end)  as t12 from apa_file where apa06 = '"
        oCommander2.CommandText += apa06 & "' and apa41 = 'Y' and apa02 between to_date('"
        oCommander2.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommander2.CommandText += DStartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and apa00 = " & apa00 & " group by month(apa02) )"
        oReader2 = oCommander2.ExecuteReader()
        If oReader2.HasRows Then
            oReader2.Read()
            If oReader2.Item("t1") <> 0 Then
                Ws.Cells(LineZ, 4) = oReader2.Item("t1")
            End If
            If oReader2.Item("t2") <> 0 Then
                Ws.Cells(LineZ, 5) = oReader2.Item("t2")
            End If
            If oReader2.Item("t3") <> 0 Then
                Ws.Cells(LineZ, 6) = oReader2.Item("t3")
            End If
            If oReader2.Item("t4") <> 0 Then
                Ws.Cells(LineZ, 7) = oReader2.Item("t4")
            End If
            If oReader2.Item("t5") <> 0 Then
                Ws.Cells(LineZ, 8) = oReader2.Item("t5")
            End If
            If oReader2.Item("t6") <> 0 Then
                Ws.Cells(LineZ, 9) = oReader2.Item("t6")
            End If
            If oReader2.Item("t7") <> 0 Then
                Ws.Cells(LineZ, 10) = oReader2.Item("t7")
            End If
            If oReader2.Item("t8") <> 0 Then
                Ws.Cells(LineZ, 11) = oReader2.Item("t8")
            End If
            If oReader2.Item("t9") <> 0 Then
                Ws.Cells(LineZ, 12) = oReader2.Item("t9")
            End If
            If oReader2.Item("t10") <> 0 Then
                Ws.Cells(LineZ, 13) = oReader2.Item("t10")
            End If
            If oReader2.Item("t11") <> 0 Then
                Ws.Cells(LineZ, 14) = oReader2.Item("t11")
            End If
            If oReader2.Item("t12") <> 0 Then
                Ws.Cells(LineZ, 15) = oReader2.Item("t12")
            End If
            LineZ += 1
        Else
            LineZ += 1
        End If
        oReader2.Close()
    End Sub
End Class