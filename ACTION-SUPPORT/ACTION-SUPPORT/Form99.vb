Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form99
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form99_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
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
        tYear = Me.NumericUpDown1.Value
        tMonth = Me.NumericUpDown2.Value
        'ExportToExcel()
        'SaveExcel()
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Label5.Text = "导出中"
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Label5.Text = "已完成"
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "工单上阶在制成本明细表"
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
        AdjustExcelFormat()
        oCommand.CommandText = "select ccg01,ccg02,ccg03,ccg04,sfb38,ccg11,ccg12,ccg12a,ccg12b,ccg12c,ccg12e,ccg12d,ccg20,"
        oCommand.CommandText += "ccg21,ccg22,ccg22a,ccg22b,ccg22c,ccg22e,ccg22d,ccg23,ccg23a,ccg23b,ccg23c,ccg23e,ccg23d,ccg31,ccg32,"
        oCommand.CommandText += "ccg32a,ccg32b,ccg32c,ccg32e,ccg32d,ccg41,ccg91,ccg92,ccg92a,ccg92b,ccg92c,ccg92e,ccg92d from ccg_file,sfb_file where ccg01 = sfb01 and ccg02 =" & tYear & " and ccg03 = " & tMonth
        If Not String.IsNullOrEmpty(Me.TextBox1.Text) Then
            oCommand.CommandText += " AND ccg04 like '" & Me.TextBox1.Text & "%'"
        End If
        If Not String.IsNullOrEmpty(Me.TextBox2.Text) Then
            oCommand.CommandText += " AND ccg01 like '" & Me.TextBox2.Text & "%'"
        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = oReader.Item(i)
                Next
                LineZ += 1
            End While
            ' 加總 
            Ws.Cells(LineZ, 4) = "合计"
            Ws.Cells(LineZ, 6) = "=SUM(F2:F" & LineZ - 1 & ")"
            oRng = Ws.Range(Ws.Cells(LineZ, 6), Ws.Cells(LineZ, 6))
            oRng.AutoFill(Destination:=Ws.Range("F" & LineZ & ":AO" & LineZ), Type:=xlFillDefault)
        End If
        oReader.Close()

    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 18.5
        Ws.Cells(1, 1) = "工单编号"
        Ws.Cells(1, 2) = "年度"
        Ws.Cells(1, 3) = "月份"
        Ws.Cells(1, 4) = "主件料号"
        oRng = Ws.Range("D4", "D4")
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(1, 5) = "成本会计结束日期"
        Ws.Cells(1, 6) = "上月结存数量"
        Ws.Cells(1, 7) = "上月结存合计金额"
        Ws.Cells(1, 8) = "上月结存材料金额"
        Ws.Cells(1, 9) = "上月结存人工金额"
        Ws.Cells(1, 10) = "上月结存制费一金额"
        Ws.Cells(1, 11) = "上月结存制费二金额"
        Ws.Cells(1, 12) = "上月结存加工金额"
        Ws.Cells(1, 13) = "本月投入工时"
        Ws.Cells(1, 14) = "本月投入原料数量"
        Ws.Cells(1, 15) = "本月投入原料合计金额"
        Ws.Cells(1, 16) = "本月投入材料金额"
        Ws.Cells(1, 17) = "本月投入人工金额"
        Ws.Cells(1, 18) = "本月投入制费一金额"
        Ws.Cells(1, 19) = "本月投入制费二金额"
        Ws.Cells(1, 20) = "本月投入加工金额"
        Ws.Cells(1, 21) = "本月投入半成品合计金额"
        Ws.Cells(1, 22) = "本月投入半成品材料金额"
        Ws.Cells(1, 23) = "本月投入半成品人工金额"
        Ws.Cells(1, 24) = "本月投入半成品制费一金额"
        Ws.Cells(1, 25) = "本月投入半成品制费二金额"
        Ws.Cells(1, 26) = "本月投入半成品加工金额"
        Ws.Cells(1, 27) = "本月完工入库数量"
        Ws.Cells(1, 28) = "本月完工入库合计金额"
        Ws.Cells(1, 29) = "本月完工入库材料金额"
        Ws.Cells(1, 30) = "本月完工入库人工金额"
        Ws.Cells(1, 31) = "本月完工入库制费一金额"
        Ws.Cells(1, 32) = "本月完工入库制费一金额"
        Ws.Cells(1, 33) = "本月完工入库加工金额"
        Ws.Cells(1, 34) = "累计报废数量"
        Ws.Cells(1, 35) = "月底结存数量"
        Ws.Cells(1, 36) = "月底结存合计金额"
        Ws.Cells(1, 37) = "月底结存材料金额"
        Ws.Cells(1, 38) = "月底结存人工金额"
        Ws.Cells(1, 39) = "月底结存制费一金额"
        Ws.Cells(1, 40) = "月底结存制费二金额"
        Ws.Cells(1, 41) = "月底结存加工金额"

        oRng = Ws.Range("F1", "AO1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.000000_ "
        LineZ = 2
    End Sub
End Class