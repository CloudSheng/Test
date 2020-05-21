Public Class Form97
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form97_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        Me.ComboBox1.SelectedIndex = 0
        Me.ComboBox2.SelectedIndex = 0
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "WorkOrder_Status_Report"
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
        Ws.Name = "模板"
        Ws.Activate()
        AdjustExcelFormat()
        Select Case ComboBox2.SelectedIndex
            Case 0
                oCommand.CommandText = "select gem02,sfb81,sfb01,sfb05,sfb25,sfb36,sfb08,sfb081,sfb09,sfb12,sfb04,(sfb081-sfb09-sfb12) from sfb_file left join gem_file on sfb82 = gem01   where sfb87 = 'Y' and sfb081 - sfb09 - sfb12 <> 0 and sfb02 not in (7,8) and sfb39 = 1 and sfb81 <= to_date('"
                oCommand.CommandText += DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
                Select Case ComboBox1.SelectedIndex
                    Case 0
                        oCommand.CommandText += " AND sfb04 = 8 "
                    Case 1
                        oCommand.CommandText += " AND sfb04 <> 8 "
                End Select
            Case 1
                oCommand.CommandText = "select pmc02,sfb81,sfb01,sfb05,sfb25,sfb36,sfb08,sfb081,sfb09,sfb12,sfb04,(sfb081-sfb09-sfb12) from sfb_file left join pmc_file on sfb82 = pmc01  where sfb87 = 'Y' and sfb081 - sfb09 - sfb12 <> 0 and sfb02 in (7,8) and sfb39 = 1 and sfb81 <= to_date('"
                oCommand.CommandText += DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
                Select Case ComboBox1.SelectedIndex
                    Case 0
                        oCommand.CommandText += " AND sfb04 = 8 "
                    Case 1
                        oCommand.CommandText += " AND sfb04 <> 8 "
                End Select
        End Select
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = oReader.Item(i)
                Next
                LineZ += 1
            End While
        End If
        oReader.Close()

    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 26.5
        Ws.Cells(1, 1) = "部门"
        Ws.Cells(1, 2) = "工单开立日期"
        Ws.Cells(1, 3) = "工单号"
        Ws.Cells(1, 4) = "主件料号"
        Ws.Cells(1, 5) = "实际开工日"
        Ws.Cells(1, 6) = "工单发料结束日期"
        Ws.Cells(1, 7) = "生产数量"
        Ws.Cells(1, 8) = "已发料套数"
        Ws.Cells(1, 9) = "入库数量"
        Ws.Cells(1, 10) = "报废数量"
        Ws.Cells(1, 11) = "工单状态"
        Ws.Cells(1, 12) = "在制剩余套数"
        oRng = Ws.Range("G1", "I1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0_ "
        oRng = Ws.Range("L1", "L1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0_ "
        LineZ = 2
    End Sub
End Class