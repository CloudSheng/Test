Public Class Form191
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim Start2 As Date
    Dim CharC As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form191_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Start2 = DateTimePicker1.Value
        CharC = TextBox1.Text
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "单阶材料标准成本"
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
        oCommand.CommandText = "Select STB02,STB03,bmb01,bma06,ia.ima02,ia.ima021,ia.ima55,bmb02,bmb03,ib.ima08,ib.ima02,ib.ima021,bmb10,Round((bmb06/bmb07) * (1+bmb08/100), 8),ib.ima25,Round(1/bmb10_fac,2), stb07, Round(Round((bmb06/bmb07) * (1+bmb08/100), 8) * stb07 / Round(1/bmb10_fac,2), 4)"
        oCommand.CommandText += ",stb09a,Round(Round((bmb06/bmb07) * (1+bmb08/100), 8) * stb09a / Round(1/bmb10_fac,2), 4), Round(Round((bmb06/bmb07) * (1+bmb08/100), 8) * stb07 / Round(1/bmb10_fac,2), 4) + Round(Round((bmb06/bmb07) * (1+bmb08/100), 8) * stb09a / Round(1/bmb10_fac,2), 4) "
        oCommand.CommandText += "from bmb_file left join ima_file ia on bmb01 = ia.ima01 left join ima_file ib on bmb03 = ib.ima01 left join stb_file on bmb03 = stb01 and stb02 = " & tYear & " and stb03 = " & tMonth & " left join bma_file on bmb01 = bma01 and bmb29 = bma06 "
        oCommand.CommandText += " where bmb01 like '%" & CharC & "%' AND BMB04 <= to_date('" & Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and (bmb05 >= to_date('" & Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') or bmb05 is null) and bma10 = 2 and bmaacti = 'Y' order by bmb01,bmb02"
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
        oRng = Ws.Range("A1", "U1")
        oRng.EntireColumn.AutoFit()
    End Sub

    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Cells(1, 1) = "年度"
        Ws.Cells(1, 2) = "月份"
        Ws.Cells(1, 3) = "主件料号"
        Ws.Cells(1, 4) = "特性代码"
        Ws.Cells(1, 5) = "品名"
        Ws.Cells(1, 6) = "规格"
        Ws.Cells(1, 7) = "BOM表生产单位"
        Ws.Cells(1, 8) = "项次"
        Ws.Cells(1, 9) = "元件料号"
        Ws.Cells(1, 10) = "来源码"
        Ws.Cells(1, 11) = "品名"
        Ws.Cells(1, 12) = "规格"
        Ws.Cells(1, 13) = "BOM表单位"
        Ws.Cells(1, 14) = "BOM表实际用量"
        Ws.Cells(1, 15) = "库存单位"
        Ws.Cells(1, 16) = "单位换算率"
        Ws.Cells(1, 17) = "材料标准单价"
        Ws.Cells(1, 18) = "材料成本"
        Ws.Cells(1, 19) = "委外加工标准单价"
        Ws.Cells(1, 20) = "委外加工成本"
        Ws.Cells(1, 21) = "材料成本合计"
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        oRng = Ws.Range("I1", "I1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        LineZ = 2
    End Sub
End Class