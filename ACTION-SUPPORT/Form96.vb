Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form96
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
    Dim DBC As String = String.Empty
    Dim LineZ As Integer = 0
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim Start2 As Date
    Dim End2 As Date
    Dim BankNo As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form96_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If IsNothing(Me.ComboBox2.SelectedItem) Then
            MsgBox("未选定银行编号")
            Return
        End If
        If IsNothing(Me.ComboBox1.SelectedItem) Then
            MsgBox("未选定营运中心")
            Return
        End If
        If Not IsNothing(ComboBox2.SelectedItem) Then
            BankNo = ComboBox2.SelectedItem.ToString()
            Dim stCount As Int16 = Strings.InStr(BankNo, "|")
            If stCount > 0 Then
                BankNo = Strings.Left(BankNo, stCount - 1)
            End If
        End If
        DBC = Me.ComboBox1.SelectedItem.ToString.ToLower()
        oConnection.ConnectionString = Module1.OpenOracleConnection(DBC)

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
        Start2 = Convert.ToDateTime(tYear & "/" & tMonth & "/01")
        End2 = Start2.AddMonths(1).AddDays(-1)
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
        SaveFileDialog1.FileName = "Banking_Balance_Sheet"
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
        ' 先抓企業餘額
        oCommand.CommandText = "select nvl(sum(nmp06),0) from nmp_file where nmp01 = '" & BankNo & "' and nmp02 = " & tYear & " and nmp03 = " & tMonth
        Dim CompanyBalance As Decimal = oCommand.ExecuteScalar()
        Ws.Cells(5, 4) = CompanyBalance
        ' 再抓 B00 銀行餘額
        oCommand.CommandText = "select nvl(sum(tc_bal04),0) from tc_bal_file where tc_bal05 = '" & BankNo & "' and tc_bal01 = to_date('" & End2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_bal02 = 'B00'"
        Dim BankBalance As Decimal = oCommand.ExecuteScalar()
        Ws.Cells(5, 8) = BankBalance
        ' 抓取 銀+
        Dim CompanyTrueBalance As Decimal = CompanyBalance
        oCommand.CommandText = "select tc_bal01,tc_bal02,tc_bal03,tc_bal04 from tc_bal_file,nmk_file where tc_bal05 = '" & BankNo & "' and tc_bal01 BETWEEN to_date('"
        oCommand.CommandText += Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += End2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_bal02 like 'C%' AND nmk03 = '+' and tc_bal02 = nmk01"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("tc_bal01")
                Ws.Cells(LineZ, 2) = oReader.Item("tc_bal02")
                Ws.Cells(LineZ, 3) = oReader.Item("tc_bal03")
                Ws.Cells(LineZ, 4) = oReader.Item("tc_bal04")
                CompanyTrueBalance += oReader.Item("tc_bal04")
                LineZ += 1
            End While
        Else
            LineZ += 1
        End If
        oReader.Close()
        ' 空行
        LineZ += 1
        ' 抓取銀 - 
        Ws.Cells(LineZ, 1) = "减：银行已付、企业未付款"
        Ws.Cells(LineZ, 5) = "减：企业已付、银行未付款"
        LineZ += 1
        Ws.Cells(LineZ, 1) = "日期"
        Ws.Cells(LineZ, 2) = "调节码"
        Ws.Cells(LineZ, 3) = "备注"
        Ws.Cells(LineZ, 4) = "金额"
        Ws.Cells(LineZ, 5) = "收支日期"
        Ws.Cells(LineZ, 6) = "收支单号"
        Ws.Cells(LineZ, 7) = "摘要"
        Ws.Cells(LineZ, 8) = "金额"
        LineZ += 1
        oCommand.CommandText = "select tc_bal01,tc_bal02,tc_bal03,tc_bal04 from tc_bal_file,nmk_file where tc_bal05 = '" & BankNo & "' and tc_bal01 BETWEEN to_date('"
        oCommand.CommandText += Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += End2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_bal02 like 'C%' AND nmk03 = '-' and tc_bal02 = nmk01"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("tc_bal01")
                Ws.Cells(LineZ, 2) = oReader.Item("tc_bal02")
                Ws.Cells(LineZ, 3) = oReader.Item("tc_bal03")
                Ws.Cells(LineZ, 4) = oReader.Item("tc_bal04")
                CompanyTrueBalance -= oReader.Item("tc_bal04")
                LineZ += 1
            End While
        Else
            LineZ += 1
        End If
        oReader.Close()

        ' 空行
        LineZ += 1

        Ws.Cells(LineZ, 1) = "调整后余额"
        Ws.Cells(LineZ, 4) = CompanyTrueBalance
        Ws.Cells(LineZ, 5) = "调整后余额"
        Ws.Cells(LineZ, 8) = BankBalance

        ' 劃線
        oRng = Ws.Range("A5", Ws.Cells(LineZ, 4))
        oRng.Borders(xlEdgeLeft).LineStyle = xlDouble
        oRng.Borders(xlEdgeTop).LineStyle = xlDouble
        oRng.Borders(xlEdgeBottom).LineStyle = xlDouble
        oRng.Borders(xlEdgeRight).LineStyle = xlDouble
        oRng = Ws.Range("E5", Ws.Cells(LineZ, 8))
        oRng.Borders(xlEdgeLeft).LineStyle = xlDouble
        oRng.Borders(xlEdgeTop).LineStyle = xlDouble
        oRng.Borders(xlEdgeBottom).LineStyle = xlDouble
        oRng.Borders(xlEdgeRight).LineStyle = xlDouble

        LineZ += 2
        Ws.Cells(LineZ, 1) = "批准: "
        Ws.Cells(LineZ, 5) = "审核:"
        Ws.Cells(LineZ, 8) = "制表:黎莉"
        Ws.Cells(LineZ + 2, 7) = "制表日期:  " & Now.Date()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox2.Items.Clear()
        If IsNothing(Me.ComboBox1.SelectedItem) Then
            Me.ComboBox2.Items.Clear()
        Else
            DBC = Me.ComboBox1.SelectedItem.ToString.ToLower()
            If oConnection.State <> ConnectionState.Closed Then
                oConnection.Close()
            End If
            oConnection.ConnectionString = Module1.OpenOracleConnection(DBC)
            If oConnection.State <> ConnectionState.Open Then
                Try
                    oConnection.Open()
                    oCommand.Connection = oConnection
                    oCommand.CommandType = CommandType.Text
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
            oCommand.CommandText = "select nma01,nma02 from nma_file where nmaacti = 'Y'"
            oReader = oCommand.ExecuteReader()
            If oReader.HasRows() Then
                While oReader.Read()
                    Me.ComboBox2.Items.Add(oReader.Item(0).ToString() & "|" & oReader.Item(1).ToString())
                End While
            End If
            oReader.Close()
            oConnection.Close()
        End If
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        xExcel.ActiveWindow.DisplayGridlines = False
        Ws.Columns.EntireColumn.ColumnWidth = 15.8
        Ws.Rows.EntireRow.RowHeight = 29.3

        oRng = Ws.Range("A1", "H1")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        oRng.Font.Size = 18
        oRng = Ws.Range("A2", "H2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        oRng.Font.Size = 18
        oRng = Ws.Range("A3", "H3")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Select Case DBC
            Case "actiontest"
                Ws.Cells(1, 1) = "Dongguan Action Composites LTD Co."
            Case "hkacttest"
                Ws.Cells(1, 1) = "Action Composite Technology Limited"
            Case "action_bvi"
                Ws.Cells(1, 1) = "Action Composites International Limited"
        End Select
        Ws.Cells(2, 1) = "银行存款余额调节表"
        Ws.Cells(3, 1) = End2.ToString("yyyy/MM/dd")
        oCommand.CommandText = "select nma03 from nma_file where nma01 = '" & BankNo & "'"
        Dim BankDesc As String = oCommand.ExecuteScalar()
        oCommand.CommandText = "select nma04 from nma_file where nma01 = '" & BankNo & "'"
        Dim BankAccNo As String = oCommand.ExecuteScalar()
        Ws.Cells(4, 1) = "银行名称:" & BankDesc & "-" & BankAccNo
        oCommand.CommandText = "select nma10 from nma_file where nma01 = '" & BankNo & "'"
        Dim Currency As String = oCommand.ExecuteScalar()
        Ws.Cells(4, 8) = "币种：" & Currency
        Ws.Cells(5, 1) = "企业账面余额:"
        Ws.Cells(5, 5) = "银行对账单余额:"
        Ws.Cells(7, 1) = "加：银行已收、企业未收款"
        Ws.Cells(7, 5) = "加：企业已收、银行未收款"
        Ws.Cells(8, 1) = "日期"
        Ws.Cells(8, 2) = "调节码"
        Ws.Cells(8, 3) = "备注"
        Ws.Cells(8, 4) = "金额"
        Ws.Cells(8, 5) = "收支日期"
        Ws.Cells(8, 6) = "收支单号"
        Ws.Cells(8, 7) = "摘要"
        Ws.Cells(8, 8) = "金额"
        oRng = Ws.Range("D4", "D4")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00_ "
        oRng = Ws.Range("H4", "H4")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00_ "
        LineZ = 9
    End Sub
End Class