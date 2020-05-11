Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form94
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
    Dim Start2 As Date
    Dim End2 As Date
    Dim ExchangeRate1 As Decimal = 0
    Dim ExchangeRate2 As Decimal = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form94_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Start2 = Convert.ToDateTime(tYear & "/" & tMonth & "/01")
        End2 = Convert.ToDateTime(tYear & "/" & tMonth & "/01").AddMonths(1).AddDays(-1)
        ExportToExcel()
        SaveExcel()
        'BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "客户销售报表"
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
        Ws.Name = "客户销售"
        Ws.Activate()
        AdjustExcelFormat()
        oCommand.CommandText = "select distinct occ02 from ( select occ02 from occ_file union all select tc_bud05 from tc_bud_file where tc_bud01 = '1' and  tc_bud02 = " & tYear & " and tc_bud03 = " & tMonth & ")  where occ02 not in ('Austria Action','Sabelt')"
        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            While oReader.Read()
                Dim occ01 As String = String.Empty
                Dim occud02 As String = String.Empty
                occ01 = Getocc01(oReader.Item("occ02"))
                Ws.Cells(LineZ, 1) = oReader.Item("occ02")
                Ws.Cells(LineZ, 3) = "USD"
                If Not String.IsNullOrWhiteSpace(occ01) Then
                    occud02 = Getoccud02(occ01)
                    Ws.Cells(LineZ, 2) = GetCustName(occ01)
                    Ws.Cells(LineZ, 4) = GetSalesAct(occud02, "USD")
                    Ws.Cells(LineZ, 6) = GetCostReal(occud02, "USD")
                    Ws.Cells(LineZ, 10) = "=(D" & LineZ & "-F" & LineZ & ")/D" & LineZ
                End If
                Ws.Cells(LineZ, 5) = GetTcBud13(oReader.Item("occ02"), "USD")
                Ws.Cells(LineZ, 7) = GetBudgetCost(oReader.Item("occ02"), "USD")
                Ws.Cells(LineZ, 8) = "=D" & LineZ & "-F" & LineZ
                Ws.Cells(LineZ, 9) = "=E" & LineZ & "-G" & LineZ
                Ws.Cells(LineZ, 11) = "=(E" & LineZ & "-G" & LineZ & ")/E" & LineZ
                ' 第二行
                Ws.Cells(LineZ + 1, 1) = oReader.Item("occ02")
                Ws.Cells(LineZ + 1, 3) = "EUR"
                If Not String.IsNullOrWhiteSpace(occ01) Then
                    Ws.Cells(LineZ + 1, 2) = GetCustName(occ01)
                    Ws.Cells(LineZ + 1, 4) = GetSalesAct(occud02, "EUR")
                    Ws.Cells(LineZ + 1, 6) = GetCostReal(occud02, "EUR")
                    Ws.Cells(LineZ + 1, 10) = "=(D" & LineZ + 1 & "-F" & LineZ + 1 & ")/D" & LineZ + 1
                End If
                Ws.Cells(LineZ + 1, 5) = GetTcBud13(oReader.Item("occ02"), "EUR")
                Ws.Cells(LineZ + 1, 7) = GetBudgetCost(oReader.Item("occ02"), "EUR")
                Ws.Cells(LineZ + 1, 8) = "=D" & LineZ + 1 & "-F" & LineZ + 1
                Ws.Cells(LineZ + 1, 9) = "=E" & LineZ + 1 & "-G" & LineZ + 1
                Ws.Cells(LineZ + 1, 11) = "=(E" & LineZ + 1 & "-G" & LineZ + 1 & ")/E" & LineZ + 1

                LineZ += 2
            End While
        End If
        oReader.Close()
        ' 加總
        Ws.Cells(LineZ, 2) = "Grand Total"
        Ws.Cells(LineZ, 4) = "=SUM(D5:D" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 5) = "=SUM(E5:E" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 6) = "=SUM(F5:F" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 7) = "=SUM(G5:G" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 8) = "=SUM(H5:H" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 9) = "=SUM(I5:I" & LineZ - 1 & ")"
        oRng = Ws.Range("A3", Ws.Cells(LineZ, 11))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous



        oRng = Ws.Range("D5", Ws.Cells(LineZ, 9))
        oRng.NumberFormatLocal = "#,##0.00_ "
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 17
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "B1")
        oRng.EntireColumn.ColumnWidth = 40.89
        Ws.Cells(1, 1) = "Year:" & tYear
        Ws.Cells(2, 1) = "Month:" & tMonth
        oRng = Ws.Range("A3", "K3")
        oRng.Interior.Color = Color.FromArgb(146, 208, 80)
        Ws.Cells(4, 1) = "账款客户"
        Ws.Cells(4, 2) = "收款客户"
        Ws.Cells(4, 3) = "Currency"
        Ws.Cells(4, 4) = "Sales-Actual"
        Ws.Cells(4, 5) = "sales-Budget"
        Ws.Cells(4, 6) = "Cost-Actual"
        Ws.Cells(4, 7) = "Cost-Standard"
        Ws.Cells(4, 8) = "Margin-Actual"
        Ws.Cells(4, 9) = "Margin-Budget"
        Ws.Cells(4, 10) = "GM%-Actual"
        Ws.Cells(4, 11) = "GM%-Budget"
        oRng = Ws.Range("J5", "K5")
        oRng.EntireColumn.NumberFormatLocal = "0%"

        Dim TYM1 As String = String.Empty
        If tMonth < 10 Then
            TYM1 = tYear & "0" & tMonth
        Else
            TYM1 = tYear & tMonth
        End If
        oCommand.CommandText = "select nvl(azj041,1) from azj_file where azj01 = 'USD' and azj02 = '" & TYM1 & "'"
        ExchangeRate1 = oCommand.ExecuteScalar()
        oCommand.CommandText = "select nvl(azj07,1) from azj_file where azj01 = 'USD' and azj02 = '" & TYM1 & "'"
        ExchangeRate2 = oCommand.ExecuteScalar()
        LineZ = 5
    End Sub
    Private Function Getocc01(ByVal occ02 As String)
        oCommand2.CommandText = "select occ01 from occ_file where occ02 = '" & occ02 & "'"
        Dim l_occ01 As String = oCommand2.ExecuteScalar()
        If IsDBNull(l_occ01) Then
            l_occ01 = ""
        End If
        Return l_occ01
    End Function
    Private Function Getoccud02(ByVal occ01 As String)
        oCommand2.CommandText = "select occud02 from occ_file where occ01 = '" & occ01 & "'"
        Dim l_occud02 As String = oCommand2.ExecuteScalar()
        Return l_occud02
    End Function
    Private Function GetSalesAct(ByVal occud02 As String, ByVal Currency As String)
        oCommand2.CommandText = "select nvl(sum(t1),0) as t1 from ( "
        oCommand2.CommandText += "select round(sum(ogb14 * oga24 /" & ExchangeRate1 & "),3) as t1 from oga_file,ogb_file where oga01 = ogb01 and ogapost = 'Y' "
        oCommand2.CommandText += "and oga02 between to_date('" & Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') AND to_date('"
        oCommand2.CommandText += End2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') AND substr(ogb04,4,2) = '"
        oCommand2.CommandText += occud02 & "' AND oga23 ='" & Currency & "' union all select round(sum(ohb14 * -1 * oha24 /" & ExchangeRate1 & "),3) from oha_file,ohb_file where oha01 = ohb01 and ohapost = 'Y' "
        oCommand2.CommandText += "and oha02 between to_date('" & Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') AND to_date('"
        oCommand2.CommandText += End2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') AND substr(ohb04,4,2) = '"
        oCommand2.CommandText += occud02 & "' AND oha23 = '" & Currency & "' )"
        Dim GS As Decimal = oCommand2.ExecuteScalar()
        Return GS
    End Function
    Private Function GetTcBud13(ByVal occ02 As String, ByVal Currency As String)
        oCommand2.CommandText = "select nvl(SUM(tc_bud13),0) from tc_bud_file where tc_bud01 = '1' and tc_bud05 = '"
        oCommand2.CommandText += occ02 & "' and tc_bud14 = '" & Currency & "' AND tc_bud02 = " & tYear & " and tc_bud03 = " & tMonth
        Dim TS As Decimal = oCommand2.ExecuteScalar()
        Return TS
    End Function
    Private Function GetCostReal(ByVal occud02 As String, ByVal Currency As String)

        oCommand2.CommandText = "select round(nvl(sum(t1),0),3) as t1 from ( "
        oCommand2.CommandText += "select sum(ccc23 * ogb12 * ogb15_fac /" & ExchangeRate2 & ") as t1 from oga_file,ogb_file,ccc_file where oga01 = ogb01 and ogapost = 'Y' "
        oCommand2.CommandText += "and oga02 between to_date('" & Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') AND to_date('"
        oCommand2.CommandText += End2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') AND substr(ogb04,4,2) = '"
        oCommand2.CommandText += occud02 & "' AND oga23 ='" & Currency & "' and ogb04 = ccc01 and ccc02 = " & tYear & " and ccc03 = " & tMonth
        oCommand2.CommandText += " union all "
        oCommand2.CommandText += "select sum(ccc23 * ohb12 * ohb15_fac * -1 / " & ExchangeRate2 & ") from oha_file,ohb_file,ccc_file where oha01 = ohb01 and ohapost = 'Y' "
        oCommand2.CommandText += "and oha02 between to_date('" & Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') AND to_date('"
        oCommand2.CommandText += End2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') AND substr(ohb04,4,2) = '"
        oCommand2.CommandText += occud02 & "' and oha23 = '" & Currency & "' and ohb04 = ccc01 and ccc02 = " & tYear & " and ccc03 = " & tMonth & " )"
        Dim CS As Decimal = oCommand2.ExecuteScalar()
        Return CS
    End Function
    Private Function GetBudgetCost(ByVal occ02 As String, ByVal Currency As String)
        oCommand2.CommandText = "select round(nvl(sum(tc_bud11 * (stb07 + stb08 + stb09) / " & ExchangeRate2 & "),0),3)   from tc_bud_file,stb_file where tc_bud01 = '1' and tc_bud05 = '"
        oCommand2.CommandText += occ02 & "' and tc_bud14 = '" & Currency & "' AND tc_bud02 = " & tYear & " and tc_bud03 = " & tMonth & " and tc_bud04 = stb01 and stb02 = tc_bud02 and stb03 = tc_bud03 "
        Dim BS As Decimal = oCommand2.ExecuteScalar()
        Return BS
    End Function
    Private Function GetCustName(ByVal occ01 As String)
        oCommand2.CommandText = "select  occ02 from oga_file,occ_file where oga03 = '" & occ01 & "' and oga18 = occ01"
        Dim GCS As String = oCommand2.ExecuteScalar()
        Return GCS
    End Function
End Class