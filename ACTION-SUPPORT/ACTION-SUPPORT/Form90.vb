Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form90
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

    Private Sub Form90_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Working_Capital_Report"
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
        Ws.Name = "Working Capital"
        Ws.Activate()
        AdjustExcelFormat()
        Ws.Cells(6, 4) = GetLastYearSameMonth("112201") / ExchangeRate1
        Ws.Cells(6, 5) = GetLastYearValue("112201") / ExchangeRate1
        Ws.Cells(6, 6) = GetThisYearByQuarter("112201", 1) / ExchangeRate1
        Ws.Cells(6, 7) = GetThisYearByQuarter("112201", 2) / ExchangeRate1
        Ws.Cells(6, 8) = GetThisYearByQuarter("112201", 3) / ExchangeRate1
        Ws.Cells(6, 9) = GetThisYearByQuarter("112201", 4) / ExchangeRate1
        Ws.Cells(6, 10) = GetThisYearSameMonth("112201") / ExchangeRate1

        Ws.Cells(7, 4) = GetLastYearSameMonth("112202") / ExchangeRate1
        Ws.Cells(7, 5) = GetLastYearValue("112202") / ExchangeRate1
        Ws.Cells(7, 6) = GetThisYearByQuarter("112202", 1) / ExchangeRate1
        Ws.Cells(7, 7) = GetThisYearByQuarter("112202", 2) / ExchangeRate1
        Ws.Cells(7, 8) = GetThisYearByQuarter("112202", 3) / ExchangeRate1
        Ws.Cells(7, 9) = GetThisYearByQuarter("112202", 4) / ExchangeRate1
        Ws.Cells(7, 10) = GetThisYearSameMonth("112202") / ExchangeRate1

        Ws.Cells(8, 4) = GetLastYearSameMonthF("AR") / ExchangeRate1
        Ws.Cells(8, 5) = GetLastYearLastMonthF("AR") / ExchangeRate1
        Ws.Cells(8, 6) = GetThisYearSMonthF("AR", 3) / ExchangeRate1
        Ws.Cells(8, 7) = GetThisYearSMonthF("AR", 6) / ExchangeRate1
        Ws.Cells(8, 8) = GetThisYearSMonthF("AR", 9) / ExchangeRate1
        Ws.Cells(8, 9) = GetThisYearSMonthF("AR", 12) / ExchangeRate1
        Ws.Cells(8, 10) = GetThisYearSameMonthF("AR") / ExchangeRate1

        Ws.Cells(9, 4) = GetLastYearSameMonth("122101") / ExchangeRate1
        Ws.Cells(9, 5) = GetLastYearValue("122101") / ExchangeRate1
        Ws.Cells(9, 6) = GetThisYearByQuarter("122101", 1) / ExchangeRate1
        Ws.Cells(9, 7) = GetThisYearByQuarter("122101", 2) / ExchangeRate1
        Ws.Cells(9, 8) = GetThisYearByQuarter("122101", 3) / ExchangeRate1
        Ws.Cells(9, 9) = GetThisYearByQuarter("122101", 4) / ExchangeRate1
        Ws.Cells(9, 10) = GetThisYearSameMonth("122101") / ExchangeRate1

        Ws.Cells(10, 4) = GetLastYearSameMonth("122102") / ExchangeRate1
        Ws.Cells(10, 5) = GetLastYearValue("122102") / ExchangeRate1
        Ws.Cells(10, 6) = GetThisYearByQuarter("122102", 1) / ExchangeRate1
        Ws.Cells(10, 7) = GetThisYearByQuarter("122102", 2) / ExchangeRate1
        Ws.Cells(10, 8) = GetThisYearByQuarter("122102", 3) / ExchangeRate1
        Ws.Cells(10, 9) = GetThisYearByQuarter("122102", 4) / ExchangeRate1
        Ws.Cells(10, 10) = GetThisYearSameMonth("122102") / ExchangeRate1

        Dim R1 As Decimal = GetLastYearSameMonth("1403", "1404") / ExchangeRate1
        Dim R2 As Decimal = GetLastYearSameMonth("1406", "1408") / ExchangeRate1
        Ws.Cells(14, 4) = R1 + R2
        Dim R1A As Decimal = GetLastYearValue("1403", "1404") / ExchangeRate1
        Dim R2A As Decimal = GetLastYearValue("1406", "1408") / ExchangeRate1
        Ws.Cells(14, 5) = R1A + R2A
        Dim R1B As Decimal = GetThisYearByQuarter("1403", "1404", 1) / ExchangeRate1
        Dim R2B As Decimal = GetThisYearByQuarter("1406", "1408", 1) / ExchangeRate1
        Ws.Cells(14, 6) = R1B + R2B
        Dim R1C As Decimal = GetThisYearByQuarter("1403", "1404", 2) / ExchangeRate1
        Dim R2C As Decimal = GetThisYearByQuarter("1406", "1408", 2) / ExchangeRate1
        Ws.Cells(14, 7) = R1C + R2C
        Dim R1D As Decimal = GetThisYearByQuarter("1403", "1404", 3) / ExchangeRate1
        Dim R2D As Decimal = GetThisYearByQuarter("1406", "1408", 3) / ExchangeRate1
        Ws.Cells(14, 8) = R1D + R2D
        Dim R1E As Decimal = GetThisYearByQuarter("1403", "1404", 4) / ExchangeRate1
        Dim R2E As Decimal = GetThisYearByQuarter("1406", "1408", 4) / ExchangeRate1
        Ws.Cells(14, 9) = R1E + R2E
        Dim R1F As Decimal = GetThisYearSameMonth("1403", "1404") / ExchangeRate1
        Dim R2F As Decimal = GetThisYearSameMonth("1406", "1408") / ExchangeRate1
        Ws.Cells(14, 10) = R1F + R2F

        Ws.Cells(15, 4) = GetLastYearSameMonth("1413") / ExchangeRate1
        Ws.Cells(15, 5) = GetLastYearValue("1413") / ExchangeRate1
        Ws.Cells(15, 6) = GetThisYearByQuarter("1413", 1) / ExchangeRate1
        Ws.Cells(15, 7) = GetThisYearByQuarter("1413", 2) / ExchangeRate1
        Ws.Cells(15, 8) = GetThisYearByQuarter("1413", 3) / ExchangeRate1
        Ws.Cells(15, 9) = GetThisYearByQuarter("1413", 4) / ExchangeRate1
        Ws.Cells(15, 10) = GetThisYearSameMonth("1413") / ExchangeRate1

        Ws.Cells(16, 4) = GetLastYearSameMonth("1409") / ExchangeRate1
        Ws.Cells(16, 5) = GetLastYearValue("1409") / ExchangeRate1
        Ws.Cells(16, 6) = GetThisYearByQuarter("1409", 1) / ExchangeRate1
        Ws.Cells(16, 7) = GetThisYearByQuarter("1409", 2) / ExchangeRate1
        Ws.Cells(16, 8) = GetThisYearByQuarter("1409", 3) / ExchangeRate1
        Ws.Cells(16, 9) = GetThisYearByQuarter("1409", 4) / ExchangeRate1
        Ws.Cells(16, 10) = GetThisYearSameMonth("1409") / ExchangeRate1

        Ws.Cells(17, 4) = GetLastYearSameMonth("500101", "500104", "D9999") / ExchangeRate1
        Ws.Cells(17, 5) = GetLastYearValue("500101", "500104", "D9999") / ExchangeRate1
        Ws.Cells(17, 6) = GetThisYearByQuarter("500101", "500104", 1, "D9999") / ExchangeRate1
        Ws.Cells(17, 7) = GetThisYearByQuarter("500101", "500104", 2, "D9999") / ExchangeRate1
        Ws.Cells(17, 8) = GetThisYearByQuarter("500101", "500104", 3, "D9999") / ExchangeRate1
        Ws.Cells(17, 9) = GetThisYearByQuarter("500101", "500104", 4, "D9999") / ExchangeRate1
        Ws.Cells(17, 10) = GetThisYearSameMonth("500101", "500104", "D9999") / ExchangeRate1

        Dim R3 As Decimal = GetLastYearSameMonth("1405") / ExchangeRate1
        Dim R4 As Decimal = GetLastYearSameMonth("1410") / ExchangeRate1
        Ws.Cells(18, 4) = R3 + R4
        Dim R3A As Decimal = GetLastYearValue("1405") / ExchangeRate1
        Dim R4A As Decimal = GetLastYearValue("1410") / ExchangeRate1
        Ws.Cells(18, 5) = R3A + R4A
        Dim R3B As Decimal = GetThisYearByQuarter("1405", 1) / ExchangeRate1
        Dim R4B As Decimal = GetThisYearByQuarter("1410", 1) / ExchangeRate1
        Ws.Cells(18, 6) = R3B + R4B
        Dim R3C As Decimal = GetThisYearByQuarter("1405", 2) / ExchangeRate1
        Dim R4C As Decimal = GetThisYearByQuarter("1410", 2) / ExchangeRate1
        Ws.Cells(18, 7) = R3C + R4C
        Dim R3D As Decimal = GetThisYearByQuarter("1405", 3) / ExchangeRate1
        Dim R4D As Decimal = GetThisYearByQuarter("1410", 3) / ExchangeRate1
        Ws.Cells(18, 8) = R3D + R4D
        Dim R3E As Decimal = GetThisYearByQuarter("1405", 4) / ExchangeRate1
        Dim R4E As Decimal = GetThisYearByQuarter("1410", 4) / ExchangeRate1
        Ws.Cells(18, 9) = R3E + R4E
        Dim R3F As Decimal = GetThisYearSameMonth("1405") / ExchangeRate1
        Dim R4F As Decimal = GetThisYearSameMonth("1410") / ExchangeRate1
        Ws.Cells(18, 10) = R3F + R4F

        Ws.Cells(19, 4) = GetLastYearSameMonth("1412") / ExchangeRate1
        Ws.Cells(19, 5) = GetLastYearValue("1412") / ExchangeRate1
        Ws.Cells(19, 6) = GetThisYearByQuarter("1412", 1) / ExchangeRate1
        Ws.Cells(19, 7) = GetThisYearByQuarter("1412", 2) / ExchangeRate1
        Ws.Cells(19, 8) = GetThisYearByQuarter("1412", 3) / ExchangeRate1
        Ws.Cells(19, 9) = GetThisYearByQuarter("1412", 4) / ExchangeRate1
        Ws.Cells(19, 10) = GetThisYearSameMonth("1412") / ExchangeRate1

        Dim R5 As Decimal = GetLastYearSameMonth("220201") / ExchangeRate1
        Dim R6 As Decimal = GetLastYearSameMonth("220203") / ExchangeRate1
        Ws.Cells(24, 4) = R5 + R6
        Dim R5A As Decimal = GetLastYearValue("220201") / ExchangeRate1
        Dim R6A As Decimal = GetLastYearValue("220203") / ExchangeRate1
        Ws.Cells(24, 5) = R5A + R6A
        Dim R5B As Decimal = GetThisYearByQuarter("220201", 1) / ExchangeRate1
        Dim R6B As Decimal = GetThisYearByQuarter("220203", 1) / ExchangeRate1
        Ws.Cells(24, 6) = R5B + R6B
        Dim R5C As Decimal = GetThisYearByQuarter("220201", 2) / ExchangeRate1
        Dim R6C As Decimal = GetThisYearByQuarter("220203", 2) / ExchangeRate1
        Ws.Cells(24, 7) = R5C + R6C
        Dim R5D As Decimal = GetThisYearByQuarter("220201", 3) / ExchangeRate1
        Dim R6D As Decimal = GetThisYearByQuarter("220203", 3) / ExchangeRate1
        Ws.Cells(24, 8) = R5D + R6D
        Dim R5E As Decimal = GetThisYearByQuarter("220201", 4) / ExchangeRate1
        Dim R6E As Decimal = GetThisYearByQuarter("220203", 4) / ExchangeRate1
        Ws.Cells(24, 9) = R5E + R6E
        Dim R5F As Decimal = GetThisYearSameMonth("220201") / ExchangeRate1
        Dim R6F As Decimal = GetThisYearSameMonth("220203") / ExchangeRate1
        Ws.Cells(24, 10) = R5F + R6F

        Ws.Cells(25, 4) = GetLastYearSameMonth("220202") / ExchangeRate1
        Ws.Cells(25, 5) = GetLastYearValue("220202") / ExchangeRate1
        Ws.Cells(25, 6) = GetThisYearByQuarter("220202", 1) / ExchangeRate1
        Ws.Cells(25, 7) = GetThisYearByQuarter("220202", 2) / ExchangeRate1
        Ws.Cells(25, 8) = GetThisYearByQuarter("220202", 3) / ExchangeRate1
        Ws.Cells(25, 9) = GetThisYearByQuarter("220202", 4) / ExchangeRate1
        Ws.Cells(25, 10) = GetThisYearSameMonth("220202") / ExchangeRate1

        Ws.Cells(26, 4) = GetLastYearSameMonth("224101") / ExchangeRate1
        Ws.Cells(26, 5) = GetLastYearValue("224101") / ExchangeRate1
        Ws.Cells(26, 6) = GetThisYearByQuarter("224101", 1) / ExchangeRate1
        Ws.Cells(26, 7) = GetThisYearByQuarter("224101", 2) / ExchangeRate1
        Ws.Cells(26, 8) = GetThisYearByQuarter("224101", 3) / ExchangeRate1
        Ws.Cells(26, 9) = GetThisYearByQuarter("224101", 4) / ExchangeRate1
        Ws.Cells(26, 10) = GetThisYearSameMonth("224101") * Decimal.MinusOne / ExchangeRate1

        Ws.Cells(27, 4) = GetLastYearSameMonth("224102") / ExchangeRate1
        Ws.Cells(27, 5) = GetLastYearValue("224102") / ExchangeRate1
        Ws.Cells(27, 6) = GetThisYearByQuarter("224102", 1) / ExchangeRate1
        Ws.Cells(27, 7) = GetThisYearByQuarter("224102", 2) / ExchangeRate1
        Ws.Cells(27, 8) = GetThisYearByQuarter("224102", 3) / ExchangeRate1
        Ws.Cells(27, 9) = GetThisYearByQuarter("224102", 4) / ExchangeRate1
        Ws.Cells(27, 10) = GetThisYearSameMonth("224102") / ExchangeRate1

        Ws.Cells(34, 4) = GetLastYearSameMonth("1231") / ExchangeRate1
        Ws.Cells(34, 5) = GetLastYearValue("1231") * Decimal.MinusOne / ExchangeRate1
        Ws.Cells(34, 6) = GetThisYearByQuarter("1231", 1) / ExchangeRate1
        Ws.Cells(34, 7) = GetThisYearByQuarter("1231", 2) / ExchangeRate1
        Ws.Cells(34, 8) = GetThisYearByQuarter("1231", 3) / ExchangeRate1
        Ws.Cells(34, 9) = GetThisYearByQuarter("1231", 4) / ExchangeRate1
        Ws.Cells(34, 10) = GetThisYearSameMonth("1231") / ExchangeRate1

        Ws.Cells(39, 4) = GetLastYearSameMonth("1471") / ExchangeRate1
        Ws.Cells(39, 5) = GetLastYearValue("1471") * Decimal.MinusOne / ExchangeRate1
        Ws.Cells(39, 6) = GetThisYearByQuarter("1471", 1) / ExchangeRate1
        Ws.Cells(39, 7) = GetThisYearByQuarter("1471", 2) / ExchangeRate1
        Ws.Cells(39, 8) = GetThisYearByQuarter("1471", 3) / ExchangeRate1
        Ws.Cells(39, 9) = GetThisYearByQuarter("1471", 4) / ExchangeRate1
        Ws.Cells(39, 10) = GetThisYearSameMonth("1471") / ExchangeRate1

        Ws.Cells(41, 4) = GetLastYearSameMonth("2211", "2232") / ExchangeRate1
        Ws.Cells(41, 5) = GetLastYearValue("2211", "2232") / ExchangeRate1
        Ws.Cells(41, 6) = GetThisYearByQuarter("2211", "2232", 1) / ExchangeRate1
        Ws.Cells(41, 7) = GetThisYearByQuarter("2211", "2232", 2) / ExchangeRate1
        Ws.Cells(41, 8) = GetThisYearByQuarter("2211", "2232", 3) / ExchangeRate1
        Ws.Cells(41, 9) = GetThisYearByQuarter("2211", "2232", 4) / ExchangeRate1
        Ws.Cells(41, 10) = GetThisYearSameMonth("2211", "2232") / ExchangeRate1

        Ws.Cells(43, 4) = GetLastYearSameMonthF("AP") * Decimal.MinusOne / ExchangeRate1
        Ws.Cells(43, 5) = GetLastYearLastMonthF("AP") * Decimal.MinusOne / ExchangeRate1
        Ws.Cells(43, 6) = GetThisYearSMonthF("AP", 3) * Decimal.MinusOne / ExchangeRate1
        Ws.Cells(43, 7) = GetThisYearSMonthF("AP", 6) * Decimal.MinusOne / ExchangeRate1
        Ws.Cells(43, 8) = GetThisYearSMonthF("AP", 9) * Decimal.MinusOne / ExchangeRate1
        Ws.Cells(43, 9) = GetThisYearSMonthF("AP", 12) * Decimal.MinusOne / ExchangeRate1
        Ws.Cells(43, 10) = GetThisYearSameMonthF("AP") * Decimal.MinusOne / ExchangeRate1
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        xExcel.ActiveWindow.DisplayGridlines = False
        Ws.Columns.EntireColumn.ColumnWidth = 16.22
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 5.78
        Ws.Cells(1, 1) = "营运中心：Dongguan Action Composites LTD Co."
        Dim TYM1 As String = String.Empty
        If tMonth < 10 Then
            TYM1 = tYear & "0" & tMonth
        Else
            TYM1 = tYear & tMonth
        End If
        oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & TYM1 & "'"
        ExchangeRate1 = oCommand.ExecuteScalar()
        Ws.Cells(2, 1) = "Exchange Rate：" & ExchangeRate1
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 55.89
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.ColumnWidth = 23
        oRng.EntireColumn.HorizontalAlignment = xlLeft
        oRng = Ws.Range("A3", "J3")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        oRng.Interior.Color = Color.FromArgb(146, 208, 80)
        oRng = Ws.Range("C4", "J4")
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(3, 1) = "Y" & tYear & " Working Capital in USD"
        Ws.Cells(4, 3) = "起迄科目"
        Ws.Cells(4, 4) = tYear - 1 & "/" & GetMonthEnglish(tMonth)
        Ws.Cells(4, 5) = "Y" & tYear - 1
        Ws.Cells(4, 6) = "Q1'" & tYear
        Ws.Cells(4, 7) = "Q2'" & tYear
        Ws.Cells(4, 8) = "Q3'" & tYear
        Ws.Cells(4, 9) = "Q4'" & tYear
        Ws.Cells(4, 10) = tYear & "/" & GetMonthEnglish(tMonth)
        oRng = Ws.Range("A4", "B4")
        oRng.Merge()
        ' 劃線
        oRng = Ws.Range("A3", "J4")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        Ws.Cells(5, 1) = "Account Receivables"
        Ws.Cells(6, 2) = "Trade A/R -3rd parties"
        Ws.Cells(6, 3) = "112201"
        Ws.Cells(7, 2) = "Trade A/R- Related parties"
        Ws.Cells(7, 3) = "112202"
        Ws.Cells(8, 2) = "Forex-revaluation -AR"
        Ws.Cells(9, 2) = "Other receivables"
        Ws.Cells(9, 3) = "122101"
        Ws.Cells(10, 2) = "Other receivables-Related parties"
        Ws.Cells(10, 3) = "122102"
        Ws.Cells(11, 2) = "Total"
        Ws.Cells(11, 4) = "=SUM(D6:D10)"
        Ws.Cells(11, 5) = "=SUM(E6:E10)"
        Ws.Cells(11, 6) = "=SUM(F6:F10)"
        Ws.Cells(11, 7) = "=SUM(G6:G10)"
        Ws.Cells(11, 8) = "=SUM(H6:H10)"
        Ws.Cells(11, 9) = "=SUM(I6:I10)"
        Ws.Cells(11, 10) = "=SUM(J6:J10)"
        ' 劃線
        oRng = Ws.Range("A5", "B12")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("C5", "C12")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("D5", "D12")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("E5", "E12")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("F5", "F12")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("G5", "G12")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("H5", "H12")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("I5", "I12")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("J5", "J12")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous

        Ws.Cells(13, 1) = "Inventory"
        Ws.Cells(14, 2) = "Raw material"
        Ws.Cells(14, 3) = "1403-1404 &1406-1408"
        Ws.Cells(15, 2) = "Packaging material"
        Ws.Cells(15, 3) = "1413"
        Ws.Cells(15, 2) = "Packaging material"
        Ws.Cells(15, 3) = "1413"
        Ws.Cells(16, 2) = "Semi finished oods"
        Ws.Cells(16, 3) = "1409"
        Ws.Cells(17, 2) = "WIP"
        Ws.Cells(17, 3) = "500101-500104"
        Ws.Cells(18, 2) = "Finished goods"
        Ws.Cells(18, 3) = "1405 & 1410"
        Ws.Cells(19, 2) = "Supply and sparepart"
        Ws.Cells(19, 3) = "1412"
        Ws.Cells(21, 2) = "Total"
        Ws.Cells(21, 4) = "=SUM(D14:D19)"
        Ws.Cells(21, 5) = "=SUM(E14:E19)"
        Ws.Cells(21, 6) = "=SUM(F14:F19)"
        Ws.Cells(21, 7) = "=SUM(G14:G19)"
        Ws.Cells(21, 8) = "=SUM(H14:H19)"
        Ws.Cells(21, 9) = "=SUM(I14:I19)"
        Ws.Cells(21, 10) = "=SUM(J14:J19)"
        ' 劃線
        oRng = Ws.Range("A13", "B22")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("C13", "C22")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("D13", "D22")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("E13", "E22")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("F13", "F22")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("G13", "G22")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("H13", "H22")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("I13", "I22")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("J13", "J22")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous

        Ws.Cells(23, 1) = "Account payable"
        Ws.Cells(24, 2) = "Trade account payable- 3rd parties"
        Ws.Cells(24, 3) = "220201 &220203"
        Ws.Cells(25, 2) = "Trade account payable- Related parties"
        Ws.Cells(25, 3) = "220202"
        Ws.Cells(26, 2) = "Other Payable"
        Ws.Cells(26, 3) = "224101"
        Ws.Cells(27, 2) = "Other Payable-related parties"
        Ws.Cells(27, 3) = "224102"
        Ws.Cells(28, 2) = "Total"
        Ws.Cells(28, 4) = "=SUM(D24:D27)"
        Ws.Cells(28, 5) = "=SUM(E24:E27)"
        Ws.Cells(28, 6) = "=SUM(F24:F27)"
        Ws.Cells(28, 7) = "=SUM(G24:G27)"
        Ws.Cells(28, 8) = "=SUM(H24:H27)"
        Ws.Cells(28, 9) = "=SUM(I24:I27)"
        Ws.Cells(28, 10) = "=SUM(J24:J27)"
        ' 劃線
        oRng = Ws.Range("A23", "B28")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("C23", "C28")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("D23", "D28")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("E23", "E28")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("F23", "F28")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("G23", "G28")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("H23", "H28")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("I23", "I28")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("J23", "J28")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous

        Ws.Cells(30, 1) = "Account Receivables"
        Ws.Cells(31, 2) = "Suspense A/R"
        Ws.Cells(34, 2) = "Allowance for doubtful account - 3rd parties"
        Ws.Cells(34, 3) = "1231"
        Ws.Cells(35, 2) = "Allowance for doubtful account - Related parties"
        Ws.Cells(36, 2) = "Allowance for doubtful (All A/R)"
        Ws.Cells(36, 4) = "=SUM(D34:D35)+D32"
        Ws.Cells(36, 5) = "=SUM(E34:E35)+E32"
        Ws.Cells(36, 6) = "=SUM(F34:F35)+F32"
        Ws.Cells(36, 7) = "=SUM(G34:G35)+G32"
        Ws.Cells(36, 8) = "=SUM(H34:H35)+H32"
        Ws.Cells(36, 9) = "=SUM(I34:I35)+I32"
        Ws.Cells(36, 10) = "=SUM(J34:J35)+J32"

        ' 劃線
        oRng = Ws.Range("A30", "B36")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("C30", "C36")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("D30", "D36")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("E30", "E36")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("F30", "F36")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("G30", "G36")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("H30", "H36")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("I30", "I36")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("J30", "J36")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous

        Ws.Cells(38, 1) = "Inventory"
        Ws.Cells(39, 2) = "Provision on obsolete inventory"
        Ws.Cells(39, 3) = "1471"

        ' 劃線
        oRng = Ws.Range("A37", "B39")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("C37", "C39")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("D37", "D39")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("E37", "E39")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("F37", "F39")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("G37", "G39")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("H37", "H39")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("I37", "I39")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("J37", "J39")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous

        Ws.Cells(40, 1) = "Account payable"
        Ws.Cells(41, 2) = "Non trade account payable- 3rd parties"
        Ws.Cells(41, 3) = "2211-2232"
        Ws.Cells(42, 2) = "NonTrade AP-related parties"
        Ws.Cells(43, 2) = "Forex revaluation - Non Trade A/P"
        Ws.Cells(44, 4) = "=SUM(D41:D43)"
        Ws.Cells(44, 5) = "=SUM(E41:E43)"
        Ws.Cells(44, 6) = "=SUM(F41:F43)"
        Ws.Cells(44, 7) = "=SUM(G41:G43)"
        Ws.Cells(44, 8) = "=SUM(H41:H43)"
        Ws.Cells(44, 9) = "=SUM(I41:I43)"
        Ws.Cells(44, 10) = "=SUM(J41:J43)"

        ' 劃線
        oRng = Ws.Range("A40", "B44")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("C40", "C44")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("D40", "D44")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("E40", "E44")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("F40", "F44")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("G40", "G44")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("H40", "H44")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("I40", "I44")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng = Ws.Range("J40", "J44")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous

        Ws.Cells(45, 1) = "Net Working Capital"
        Ws.Cells(45, 4) = "=D11+D21+D28+D36+D39+D44"
        Ws.Cells(45, 5) = "=E11+E21+E28+E36+E39+E44"
        Ws.Cells(45, 6) = "=F11+F21+F28+F36+F39+F44"
        Ws.Cells(45, 7) = "=G11+G21+G28+G36+G39+G44"
        Ws.Cells(45, 8) = "=H11+H21+H28+H36+H39+H44"
        Ws.Cells(45, 9) = "=I11+I21+I28+I36+I39+I44"
        Ws.Cells(45, 10) = "=J11+J21+J28+J36+J39+J44"

        oRng = Ws.Range("A45", "J45")
        oRng.Interior.Color = Color.Yellow
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous

        oRng = Ws.Range("D5", "J45")
        oRng.NumberFormatLocal = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
        'LineZ = 2
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
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
    Private Function GetLastYearSameMonth(ByVal aag01 As String)
        oCommand.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = " & tYear - 1 & " and aah03 <= " & tMonth
        Dim NPV As Decimal = oCommand.ExecuteScalar()
        Return NPV
    End Function
    Private Function GetLastYearValue(ByVal aag01 As String)
        oCommand.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = " & tYear - 1
        Dim NPV As Decimal = oCommand.ExecuteScalar()
        Return NPV
    End Function
    Private Function GetThisYearByQuarter(ByVal aag01 As String, ByVal Quarter As Integer)
        If Quarter > 4 Or Quarter < 1 Then
            Return 0
        End If
        Dim qmonth2 As Int16 = 0
        Select Case Quarter
            Case 1
                qmonth2 = 3
            Case 2
                qmonth2 = 6
            Case 3
                qmonth2 = 9
            Case 4
                qmonth2 = 12
        End Select
        oCommand.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = " & tYear & " and aah03 <= " & qmonth2
        Dim NPV As Decimal = oCommand.ExecuteScalar()
        Return NPV
    End Function
    Private Function GetThisYearSameMonth(ByVal aag01 As String)
        oCommand.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = " & tYear & " and aah03 <= " & tMonth
        Dim NPV As Decimal = oCommand.ExecuteScalar()
        Return NPV
    End Function
    Private Function GetLastYearSameMonthF(ByVal FType As String)
        oCommand.CommandText = "select nvl(sum(oox10),0) from oox_file  where oox01 = " & tYear - 1 & " and oox02 = " & tMonth & " and oox00 = '" & FType & "'"
        Dim NPV As Decimal = oCommand.ExecuteScalar()
        Return NPV
    End Function
    Private Function GetLastYearLastMonthF(ByVal FType As String)
        oCommand.CommandText = "select nvl(sum(oox10),0) from oox_file  where oox01 = " & tYear - 1 & " and oox02 = 12  and oox00 = '" & FType & "'"
        Dim NPV As Decimal = oCommand.ExecuteScalar()
        Return NPV
    End Function
    Private Function GetThisYearSMonthF(ByVal FType As String, ByVal SMonth As Int16)
        oCommand.CommandText = "select nvl(sum(oox10),0) from oox_file  where oox01 = " & tYear & " and oox02 = " & SMonth & "  and oox00 = '" & FType & "'"
        Dim NPV As Decimal = oCommand.ExecuteScalar()
        Return NPV
    End Function
    Private Function GetThisYearSameMonthF(ByVal FType As String)
        oCommand.CommandText = "select nvl(sum(oox10),0) from oox_file  where oox01 = " & tYear & " and oox02 = " & tMonth & " and oox00 = '" & FType & "'"
        Dim NPV As Decimal = oCommand.ExecuteScalar()
        Return NPV
    End Function
    Private Function GetLastYearSameMonth(ByVal aag01 As String, ByVal aag02 As String)
        oCommand.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 between '" & aag01 & "' and '" & aag02 & "' and aah02 = " & tYear - 1 & " and aah03 <= " & tMonth
        Dim NPV As Decimal = oCommand.ExecuteScalar()
        Return NPV
    End Function
    Private Function GetLastYearValue(ByVal aag01 As String, ByVal aag02 As String)
        oCommand.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 between '" & aag01 & "' and '" & aag02 & "' and aah02 = " & tYear - 1
        Dim NPV As Decimal = oCommand.ExecuteScalar()
        Return NPV
    End Function
    Private Function GetThisYearByQuarter(ByVal aag01 As String, ByVal aag02 As String, ByVal Quarter As Integer)
        If Quarter > 4 Or Quarter < 1 Then
            Return 0
        End If
        Dim qmonth2 As Int16 = 0
        Select Case Quarter
            Case 1
                qmonth2 = 3
            Case 2
                qmonth2 = 6
            Case 3
                qmonth2 = 9
            Case 4
                qmonth2 = 12
        End Select
        oCommand.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 between '" & aag01 & "' and '" & aag02 & "' and aah02 = " & tYear & " and aah03 <= " & qmonth2
        Dim NPV As Decimal = oCommand.ExecuteScalar()
        Return NPV
    End Function
    Private Function GetThisYearSameMonth(ByVal aag01 As String, ByVal aag02 As String)
        oCommand.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 between '" & aag01 & "' and '" & aag02 & "' and aah02 = " & tYear & " and aah03 <= " & tMonth
        Dim NPV As Decimal = oCommand.ExecuteScalar()
        Return NPV
    End Function
    Private Function GetLastYearSameMonth(ByVal aag01 As String, ByVal aag02 As String, ByVal departno As String)
        oCommand.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 between '" & aag01 & "' and '" & aag02 & "' and aao02 not like '" & departno & "' and aao03 = " & tYear - 1 & " and aao04 <= " & tMonth
        Dim NPV As Decimal = oCommand.ExecuteScalar()
        Return NPV
    End Function
    Private Function GetLastYearValue(ByVal aag01 As String, ByVal aag02 As String, ByVal departno As String)
        oCommand.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 between '" & aag01 & "' and '" & aag02 & "' and aao02 not like '" & departno & "' and aao03 = " & tYear - 1
        Dim NPV As Decimal = oCommand.ExecuteScalar()
        Return NPV
    End Function
    Private Function GetThisYearByQuarter(ByVal aag01 As String, ByVal aag02 As String, ByVal Quarter As Integer, ByVal departno As String)
        If Quarter > 4 Or Quarter < 1 Then
            Return 0
        End If
        Dim qmonth2 As Int16 = 0
        Select Case Quarter
            Case 1
                qmonth2 = 3
            Case 2
                qmonth2 = 6
            Case 3
                qmonth2 = 9
            Case 4
                qmonth2 = 12
        End Select
        oCommand.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 between '" & aag01 & "' and '" & aag02 & "' and aao02 not like '" & departno & "' and aao03 = " & tYear & " and aao04 <= " & qmonth2
        Dim NPV As Decimal = oCommand.ExecuteScalar()
        Return NPV
    End Function
    Private Function GetThisYearSameMonth(ByVal aag01 As String, ByVal aag02 As String, ByVal departno As String)
        oCommand.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 between '" & aag01 & "' and '" & aag02 & "' and aao02 not like '" & departno & "' and aao03 = " & tYear & " and aao04 <= " & tMonth
        Dim NPV As Decimal = oCommand.ExecuteScalar()
        Return NPV
    End Function
End Class