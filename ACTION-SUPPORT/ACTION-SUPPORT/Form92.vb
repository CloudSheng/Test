Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form92
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
    Dim ExchangeRate1 As Decimal = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form92_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        'ExportToExcel()
        'SaveExcel()
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "DAC客户价格比价"
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
        Ws.Name = "DAC 客户价格比价"
        Ws.Activate()
        AdjustExcelFormat()
        oCommand.CommandText = "select tc_bud01,tc_bud02,tc_bud03,tc_bud04,tc_bud05,tc_bud06,tc_bud14,"
        oCommand.CommandText += "nvl(sum(tc_bud11),0) as t1,nvl(sum(tc_bud12),0) as t2,nvl(sum(tc_bud13),0) as t3,"
        oCommand.CommandText += "nvl(sum(a.ccc61 * -1),0) as  t4, nvl(sum(a.ccc63),0) as t5,nvl(sum(b.ccc61 * -1),0) as t6,nvl(sum(b.ccc63),0) as t7 "
        oCommand.CommandText += "from tc_bud_file left join ccc_file a on tc_bud02 = a.ccc02 and tc_bud03 = a.ccc03 and tc_bud04 = a.ccc01 "
        oCommand.CommandText += "left join ccc_file b on b.ccc02 = " & tYear - 1 & " and b.ccc03 = " & tMonth & " and tc_bud04 = b.ccc01 "
        oCommand.CommandText += "where tc_bud01 = '1' and tc_bud02 = " & tYear & " and tc_bud03 = " & tMonth & " group by tc_bud01,tc_bud02,tc_bud03,tc_bud04,tc_bud05,tc_bud06,tc_bud14 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select '',ccc02,ccc03,ccc01,'','','',0,0,0,nvl(sum(ccc61 * -1),0),nvl(sum(ccc63),0),0,0 "
        oCommand.CommandText += "from ccc_file where ccc02 = " & tYear & " and ccc03 = " & tMonth & " and ccc01 not in (select distinct tc_bud04 from tc_bud_file where tc_bud01 = '1' and tc_bud02 = "
        oCommand.CommandText += tYear & " and tc_bud03 = " & tMonth & ") and ccc61 <> 0  group by ccc01,ccc02,ccc03 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select ''," & tYear & "," & tMonth & ",ccc01,'','','',0,0,0,0,0,nvl(sum(ccc61 * -1),0),nvl(sum(ccc63),0) "
        oCommand.CommandText += "from ccc_file where ccc02 = " & tYear - 1 & " and ccc03 = " & tMonth & " and ccc01 not in (select distinct tc_bud04 from tc_bud_file where tc_bud01 = '1' and tc_bud02 = " & tYear & " and tc_bud03 = " & tMonth & ") "
        oCommand.CommandText += "and ccc01 not in (select distinct ccc01 from ccc_file where ccc02 = " & tYear & " and ccc03 = " & tMonth & " and ccc61 <> 0)  and ccc61 <> 0  group by ccc01,ccc02,ccc03"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                If IsDBNull(oReader.Item("tc_bud01")) Then
                    Ws.Cells(LineZ, 1) = oReader.Item("tc_bud01")
                Else
                    Ws.Cells(LineZ, 1) = "1:料号收入预算"
                End If
                Ws.Cells(LineZ, 2) = oReader.Item("tc_bud02")
                Ws.Cells(LineZ, 3) = oReader.Item("tc_bud03")
                Ws.Cells(LineZ, 4) = oReader.Item("tc_bud04")
                Ws.Cells(LineZ, 5) = oReader.Item("tc_bud05")
                Ws.Cells(LineZ, 6) = oReader.Item("tc_bud06")
                Ws.Cells(LineZ, 7) = oReader.Item("tc_bud14")
                Ws.Cells(LineZ, 8) = oReader.Item("t1")
                If IsDBNull(oReader.Item("tc_bud14")) Then
                    Ws.Cells(LineZ, 10) = oReader.Item("t3") / ExchangeRate1
                Else
                    Select Case oReader.Item("tc_bud14")
                        Case "EUR"
                            Ws.Cells(LineZ, 10) = oReader.Item("t3") * 1.05
                        Case "USD"
                            Ws.Cells(LineZ, 10) = oReader.Item("t3")
                        Case Else
                            Ws.Cells(LineZ, 10) = oReader.Item("t3") / ExchangeRate1
                    End Select
                End If
                If oReader.Item("t1") = 0 Then
                    Ws.Cells(LineZ, 9) = 0
                Else
                    Ws.Cells(LineZ, 9) = "=J" & LineZ & "/H" & LineZ
                End If
                Ws.Cells(LineZ, 11) = oReader.Item("t4")
                Ws.Cells(LineZ, 13) = oReader.Item("t5") / ExchangeRate1
                If oReader.Item("t4") = 0 Then
                    Ws.Cells(LineZ, 12) = 0
                Else
                    Ws.Cells(LineZ, 12) = "=M" & LineZ & "/K" & LineZ
                End If
                Ws.Cells(LineZ, 14) = "=K" & LineZ & "-H" & LineZ
                Ws.Cells(LineZ, 15) = "=L" & LineZ & "-I" & LineZ
                Ws.Cells(LineZ, 16) = "=M" & LineZ & "-J" & LineZ
                Ws.Cells(LineZ, 17) = oReader.Item("t6")
                Ws.Cells(LineZ, 19) = oReader.Item("t7") / ExchangeRate1
                If oReader.Item("t6") = 0 Then
                    Ws.Cells(LineZ, 18) = 0
                Else
                    Ws.Cells(LineZ, 18) = "=S" & LineZ & "/Q" & LineZ
                End If
                Ws.Cells(LineZ, 20) = "=K" & LineZ & "-Q" & LineZ
                Ws.Cells(LineZ, 21) = "=V" & LineZ & "-T" & LineZ
                Ws.Cells(LineZ, 22) = "=M" & LineZ & "-S" & LineZ
                LineZ += 1
            End While
        End If
        oReader.Close()
        oRng = Ws.Range("A5", Ws.Cells(LineZ - 1, 22))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("I7", Ws.Cells(LineZ - 1, 9))
        oRng.NumberFormatLocal = "#,##0.00_ "
        oRng = Ws.Range("L7", Ws.Cells(LineZ - 1, 12))
        oRng.NumberFormatLocal = "#,##0.00_ "
        oRng = Ws.Range("O7", Ws.Cells(LineZ - 1, 15))
        oRng.NumberFormatLocal = "#,##0.00_ "
        oRng = Ws.Range("R7", Ws.Cells(LineZ - 1, 18))
        oRng.NumberFormatLocal = "#,##0.00_ "
        oRng = Ws.Range("U7", Ws.Cells(LineZ - 1, 21))
        oRng.NumberFormatLocal = "#,##0.00_ "

        oRng = Ws.Range("J7", Ws.Cells(LineZ - 1, 10))
        oRng.NumberFormatLocal = "#,##0_ "
        oRng = Ws.Range("M7", Ws.Cells(LineZ - 1, 13))
        oRng.NumberFormatLocal = "#,##0_ "
        oRng = Ws.Range("P7", Ws.Cells(LineZ - 1, 16))
        oRng.NumberFormatLocal = "#,##0_ "
        oRng = Ws.Range("S7", Ws.Cells(LineZ - 1, 19))
        oRng.NumberFormatLocal = "#,##0_ "
        oRng = Ws.Range("V7", Ws.Cells(LineZ - 1, 22))
        oRng.NumberFormatLocal = "#,##0_ "
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 11.22
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 15
        oRng = Ws.Range("D1", "D1")
        oRng.EntireColumn.ColumnWidth = 29.11
        oRng.EntireColumn.NumberFormatLocal = "@"
        oRng = Ws.Range("E1", "E1")
        oRng.EntireColumn.ColumnWidth = 18.67
        Ws.Cells(1, 1) = "营运中心：Dongguan Action Composites LTD Co."
        Ws.Cells(2, 1) = "产品销售单价与数量比较表"
        Ws.Cells(3, 1) = "报表期间：" & tYear & "/" & tMonth
        Dim TYM1 As String = String.Empty
        If tMonth < 10 Then
            TYM1 = tYear & "0" & tMonth
        Else
            TYM1 = tYear & tMonth
        End If
        oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & TYM1 & "'"
        ExchangeRate1 = oCommand.ExecuteScalar()
        Ws.Cells(4, 1) = "汇率：" & ExchangeRate1
        oRng = Ws.Range("A5", "A6")
        oRng.Merge()
        oRng = Ws.Range("B5", "B6")
        oRng.Merge()
        oRng = Ws.Range("C5", "C6")
        oRng.Merge()
        oRng = Ws.Range("D5", "D6")
        oRng.Merge()
        oRng = Ws.Range("E5", "E6")
        oRng.Merge()
        oRng = Ws.Range("F5", "F6")
        oRng.Merge()
        oRng = Ws.Range("G5", "G6")
        oRng.Merge()
        oRng = Ws.Range("H5", "J5")
        oRng.Merge()
        oRng = Ws.Range("K5", "M5")
        oRng.Merge()
        oRng = Ws.Range("N5", "P5")
        oRng.Merge()
        oRng = Ws.Range("Q5", "S5")
        oRng.Merge()
        oRng = Ws.Range("T5", "V5")
        oRng.Merge()
        Ws.Cells(5, 1) = "类型"
        Ws.Cells(5, 2) = "年度"
        Ws.Cells(5, 3) = "月度"
        Ws.Cells(5, 4) = "料件编号"
        Ws.Cells(5, 5) = "客户简称"
        Ws.Cells(5, 6) = "业务代表"
        Ws.Cells(5, 7) = "Currency"
        Ws.Cells(5, 8) = "预算收入"
        Ws.Cells(6, 8) = "数量"
        Ws.Cells(6, 9) = "单价"
        Ws.Cells(6, 10) = "金额"
        Ws.Cells(5, 11) = "实际收入"
        Ws.Cells(6, 11) = "数量"
        Ws.Cells(6, 12) = "单价"
        Ws.Cells(6, 13) = "金额"
        Ws.Cells(5, 14) = "（实际-预算）=差异"
        Ws.Cells(6, 14) = "数量"
        Ws.Cells(6, 15) = "单价"
        Ws.Cells(6, 16) = "金额"
        Ws.Cells(5, 17) = "同期收入"
        Ws.Cells(6, 17) = "数量"
        Ws.Cells(6, 18) = "单价"
        Ws.Cells(6, 19) = "金额"
        Ws.Cells(5, 20) = "（实收-同期）=差异"
        Ws.Cells(6, 20) = "数量"
        Ws.Cells(6, 21) = "单价"
        Ws.Cells(6, 22) = "金额"
        oRng = Ws.Range("A5", "V6")
        oRng.HorizontalAlignment = xlCenter
        LineZ = 7
    End Sub
End Class