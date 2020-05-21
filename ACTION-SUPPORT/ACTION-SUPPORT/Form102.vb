Public Class Form102
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand3 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim pYear As Int16 = 0
    Dim pMonth As Int16 = 0
    Dim Start2 As Date
    Dim End2 As Date
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form102_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
                oCommand2.Connection = oConnection
                oCommand2.CommandType = CommandType.Text
                oCommand3.Connection = oConnection
                oCommand3.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        tYear = Me.NumericUpDown1.Value
        tMonth = Me.NumericUpDown2.Value
        pYear = tYear
        pMonth = tMonth - 1
        If pMonth = 0 Then
            pYear = tYear - 1
            pMonth = 12
        End If
        Start2 = Convert.ToDateTime(tYear & "/" & tMonth & "/01")
        End2 = Start2.AddMonths(1).AddDays(-1)
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "标准成本明细表"
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
        oCommand.CommandText = "select stb02,stb03,stb01,ima02,ima08,ima25,ima06,stb04,stb05,stb06,stb06a,stb07,stb08,stb09,stb09a,(stb07+stb08+stb09+stb09a),nvl(ccc23,0) from stb_file "
        oCommand.CommandText += "left join ima_file on stb01 = ima01 left join ccc_file on stb01 = ccc01 AND STB02 = ccc02 and stb03 = ccc03 where stb02 = "
        oCommand.CommandText += tYear & " and stb03 = " & tMonth & "  "
        If Not String.IsNullOrEmpty(Me.TextBox1.Text) Then
            oCommand.CommandText += " AND stb01 like '" & Me.TextBox1.Text & "%'"
        End If
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
        oRng = Ws.Range("A1", "Q1")
        oRng.EntireColumn.AutoFit()
    End Sub

    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Cells(1, 1) = "年度"
        Ws.Cells(1, 2) = "月份"
        Ws.Cells(1, 3) = "料件编号"
        Ws.Cells(1, 4) = "品名"
        Ws.Cells(1, 5) = "来源码"
        Ws.Cells(1, 6) = "库存单位"
        Ws.Cells(1, 7) = "分群码"
        Ws.Cells(1, 8) = "直接材料(本阶投入)"
        Ws.Cells(1, 9) = "直接人工(本阶投入)"
        Ws.Cells(1, 10) = "制造费用(本阶投入)"
        Ws.Cells(1, 11) = "其他制造费用(本阶投入)"
        Ws.Cells(1, 12) = "累计直接材料成本(含本阶)"
        Ws.Cells(1, 13) = "累计直接人工(含本阶)"
        Ws.Cells(1, 14) = "累计间接制造费用(含本阶)"
        Ws.Cells(1, 15) = "累计其他制造费用(含本阶)"
        Ws.Cells(1, 16) = "标准总成本"
        Ws.Cells(1, 17) = "实际总成本"
        oRng = Ws.Range("H1", "Q1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00"
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        LineZ = 2
    End Sub
End Class