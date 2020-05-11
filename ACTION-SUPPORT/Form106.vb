Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form106
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

    Private Sub Form106_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        SaveFileDialog1.FileName = "工单下阶在制成本明细表"
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
        oCommand.CommandText = "select cch02,cch03,ccg04,ima02,ima25,cch01,cch04,cch05,cch11,cch21,cch31,cch91,cch12,cch22,cch32,cch92,cch311,ccg11,ccg21,ccg31,ccg91,(ccg11+ccg21+ccg31-ccg91) "
        oCommand.CommandText += "from CCH_FILE,ccg_file,ima_file WHERE cch01 =ccg01 and cch02 = ccg02 and cch03 = ccg03 and ccg04 = ima01 "
        oCommand.CommandText += "and cch02 =" & tYear & " and cch03 = " & tMonth
        If Not String.IsNullOrEmpty(Me.TextBox1.Text) Then
            oCommand.CommandText += " AND ccg04 like '%" & Me.TextBox1.Text & "%'"
        End If
        If Not String.IsNullOrEmpty(Me.TextBox2.Text) Then
            oCommand.CommandText += " AND ccg01 like '%" & Me.TextBox2.Text & "%'"
        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = oReader.Item(i)
                Next
                LineZ += 1
            End While
            '' 加總 
            'Ws.Cells(LineZ, 4) = "合计"
            'Ws.Cells(LineZ, 6) = "=SUM(F2:F" & LineZ - 1 & ")"
            'oRng = Ws.Range(Ws.Cells(LineZ, 6), Ws.Cells(LineZ, 6))
            'oRng.AutoFill(Destination:=Ws.Range("F" & LineZ & ":AJ" & LineZ), Type:=xlFillDefault)
        End If
        oReader.Close()

    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 18.5

        Ws.Cells(1, 1) = "年度"
        Ws.Cells(1, 2) = "月份"
        Ws.Cells(1, 3) = "主件料号"
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(1, 4) = "品名"
        Ws.Cells(1, 5) = "单位"
        Ws.Cells(1, 6) = "工单编号"
        Ws.Cells(1, 7) = "元件料号"
        oRng = Ws.Range("G1", "G1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(1, 8) = "元件类型"
        Ws.Cells(1, 9) = "上月结存数量"
        Ws.Cells(1, 10) = "本月投入数量"
        Ws.Cells(1, 11) = "本月转出数量"
        Ws.Cells(1, 12) = "本月结存数量"
        Ws.Cells(1, 13) = "上月结存金额"
        Ws.Cells(1, 14) = "本月投入金额"
        Ws.Cells(1, 15) = "本月转出金额"
        Ws.Cells(1, 16) = "本月结存金额"
        Ws.Cells(1, 17) = "本月报废数量"
        Ws.Cells(1, 18) = "主件上月结存数量"
        Ws.Cells(1, 19) = "主件本月投入数量"
        Ws.Cells(1, 20) = "主件本月转出数量"
        Ws.Cells(1, 21) = "主件本月结存数量"
        Ws.Cells(1, 22) = "主件本月报废数量"

        oRng = Ws.Range("I1", "V1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.000000_ "
        LineZ = 2
    End Sub
End Class