Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form127
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
    Dim LineS1 As Int16 = 0
    Dim tYear As Int16 = 0
    Dim pYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim pMonth As Int16 = 0
    Dim lYear As Int16 = 0
    Dim tCurrency As String = String.Empty
    Dim ExchangeRate As Decimal = 0
    Dim ExchangeRate1 As Decimal = 0
    Dim gDatabase As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form127_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        If Today.Month < 10 Then
            TextBox1.Text = Today.Year & "0" & Today.Month
        Else
            TextBox1.Text = Today.Year & Today.Month
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If

        If TextBox1.Text.Length < 6 Then
            MsgBox("ERROR")
            Return
        End If
        gDatabase = Me.ComboBox2.SelectedItem.ToString()
        If String.IsNullOrEmpty(gDatabase) Then
            MsgBox("Database Error")
            Return
        End If
        Select Case gDatabase
            Case "DAC"
                oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
            Case "HAC"
                oConnection.ConnectionString = Module1.OpenOracleConnection("hkacttest")
            Case "BVI"
                oConnection.ConnectionString = Module1.OpenOracleConnection("action_bvi")
        End Select
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

        tYear = Strings.Left(Me.TextBox1.Text, 4)
        pYear = tYear - 1
        tMonth = Strings.Right(Me.TextBox1.Text, 2)
        pMonth = tMonth - 1
        If pMonth = 0 Then
            pMonth = 12
            lYear = tYear - 1
        Else
            lYear = tYear
        End If
        tCurrency = Me.ComboBox1.SelectedItem.ToString()
        If String.IsNullOrEmpty(tCurrency) Then
            MsgBox("Currency Error")
            Return
        End If
        ' 確認 ExchangeRate
        If tCurrency = "USD" And gDatabase = "DAC" Then
            Dim CS As String = String.Empty
            If tMonth < 10 Then
                CS = tYear & "0" & tMonth
            Else
                CS = tYear & tMonth
            End If
            oCommand.CommandText = "SELECT nvl(AZJ041,0) FROM AZJ_FILE WHERE AZJ01  = 'USD' AND AZJ02 = '" & CS & "'"
            ExchangeRate = oCommand.ExecuteScalar()
            If ExchangeRate = 0 Then
                ExchangeRate = 1
            End If
            ExchangeRate1 = 6.3
        Else
            ExchangeRate = 1
            ExchangeRate1 = 1
        End If

        BackgroundWorker1.RunWorkerAsync()
        'ExportToExcel()
        'SaveExcel()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        AdjustExcelFormat1()
        DoInputData("'4103','160201','160202','160203','160204','180201','180202','180203','180204','180205','180206','180207','1522'", 1)
        LineZ += 1
        DoInputData("'1121','112201','112202','1123','1231','140301','140302','1404','1405','1406','1407','1408','1409','1410','1411','1412','1413','1471','500101','500102','500103'", 1, "'2201','220201','220202','220203','220204'", 1)
        LineZ += 1
        DoInputData("'1124','1131','1132','122101','122102'", 1)
        LineZ += 1
        DoInputData("'2203','2204','2241','2242','2701','2702'", 1)
        LineZ += 1
        DoInputData("'2211'", 1)
        LineZ += 1
        DoInputData("'2232','250201','250202','250203','250204'", 1)
        LineZ += 1
        DoInputData("'1811'", 0, "'2901'", 1)
        LineZ += 2

        DoInputData1("22210101", "222108", 1)
        LineZ += 4
        DoInputData1("190101", "190102", 0)
        LineZ += 1
        DoInputData1("660301", "660301", 0)
        LineZ += 1
        DoInputData("'151101','151102','1512','1521','1523','160101','160102','160103','160104','160105','1603','1604','160501','160502','160503','160504','1606','1607','1608','170101','1703','1711','1712','180101','180102','180103','180104','180105','180106'", 1)
        LineZ += 6

        DoInputData1("660311", "660311", 0)
        LineZ += 1
        DoInputData("'400101','400102','400103','400104','410101','410102','410103','410104','410105','410106','4105'", 1)
        LineZ += 2
        DoInputData("'2001','2501'", 1)
        LineZ += 7
        DoInputData2("100101", "101301", 0)

    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.Font.Name = "Arial"
        Ws.Name = gDatabase & "-" & tCurrency

        oRng = Ws.Range("A2", "A2")
        oRng.EntireColumn.Font.Bold = True
        oRng.EntireColumn.Font.Size = 10

        oRng = Ws.Range("B2", "B2")
        oRng.EntireColumn.ColumnWidth = 35

        oRng = Ws.Range("C1", "O1")
        oRng.EntireColumn.ColumnWidth = 9.89

        Ws.Cells(2, 1) = "Company Name：Dongguan Action Composite LTD. Co"
        Ws.Cells(3, 1) = "STATEMENT OF CASH FLOW"

        oRng = Ws.Range("A7", "C8")
        oRng.Merge()
        oRng.HorizontalAlignment = xlLeft
        oRng.VerticalAlignment = xlTop
        Ws.Cells(7, 1) = " (In RMB000's)"

        oRng = Ws.Range("D7", "F7")
        oRng.EntireRow.RowHeight = 27
        oRng.Merge()
        oRng.EntireRow.Font.Bold = True
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(7, 4) = "Actual"
        Ws.Cells(7, 7) = "Budget"
        oRng = Ws.Range("H7", "H7")
        oRng.WrapText = True
        Ws.Cells(7, 8) = "Variance Act VS Bgt"

        oRng = Ws.Range("I7", "J7")
        oRng.Merge()
        Ws.Cells(7, 9) = "Actual"
        Ws.Cells(7, 11) = "Budget"
        oRng = Ws.Range("L7", "L7")
        oRng.WrapText = True
        Ws.Cells(7, 12) = "Variance Act VS Bgt"
        Ws.Cells(7, 13) = "Actual"
        Ws.Cells(7, 14) = "Forecast"
        Ws.Cells(7, 15) = "Budget"

        If tMonth < 10 Then
            Ws.Cells(8, 4) = pYear & "/0" & tMonth
            Ws.Cells(8, 6) = tYear & "/0" & tMonth
            Ws.Cells(8, 7) = tYear & "/0" & tMonth
        Else
            Ws.Cells(8, 4) = pYear & "/" & tMonth
            Ws.Cells(8, 6) = tYear & "/" & tMonth
            Ws.Cells(8, 7) = tYear & "/" & tMonth
        End If
        If pMonth < 10 Then
            Ws.Cells(8, 5) = lYear & "/0" & pMonth
        Else
            Ws.Cells(8, 5) = lYear & "/" & pMonth
        End If
        Ws.Cells(8, 8) = "USD"
        Ws.Cells(8, 9) = "YTD " & pYear
        Ws.Cells(8, 10) = "YTD " & tYear
        Ws.Cells(8, 11) = "YTD " & tYear
        Ws.Cells(8, 12) = "USD"
        Ws.Cells(8, 13) = "Y" & pYear
        Ws.Cells(8, 14) = "Y" & tYear
        Ws.Cells(8, 15) = "Y" & tYear

        Ws.Cells(9, 1) = "Cash Flows from Operating Activities:"
        Ws.Cells(10, 1) = "   Adj EBITDA"
        Ws.Cells(10, 3) = "USD$k"
        Ws.Cells(10, 8) = "=F10-G10"
        oRng = Ws.Range("H10", "H10")
        oRng.AutoFill(Destination:=Ws.Range("H10", "H18"), Type:=xlFillDefault)
        Ws.Cells(10, 12) = "=J10-K10"
        oRng = Ws.Range("L10", "L10")
        oRng.AutoFill(Destination:=Ws.Range("L10", "L18"), Type:=xlFillDefault)

        Ws.Cells(11, 1) = "   Change in NWC"
        Ws.Cells(11, 3) = "USD$k"
        Ws.Cells(12, 1) = "   Change in other assets"
        Ws.Cells(12, 3) = "USD$k"
        Ws.Cells(13, 1) = "   Change in liabilities"
        Ws.Cells(13, 3) = "USD$k"
        Ws.Cells(14, 1) = "   Change in provisions"
        Ws.Cells(14, 3) = "USD$k"
        Ws.Cells(15, 1) = "   Change in securities"
        Ws.Cells(15, 3) = "USD$k"
        Ws.Cells(16, 1) = "   Change in deferred income"
        Ws.Cells(16, 3) = "USD$k"
        Ws.Cells(17, 1) = "   Extraordinary result"
        Ws.Cells(17, 3) = "USD$k"
        Ws.Cells(18, 1) = "   Taxes"
        Ws.Cells(18, 3) = "USD$k"
        Ws.Cells(19, 1) = "Cash Flow from Operating Activities"
        Ws.Cells(19, 3) = "USD$k"
        Ws.Cells(19, 4) = "=SUM(D10:D18)"
        oRng = Ws.Range("D19", "D19")
        oRng.AutoFill(Destination:=Ws.Range("D19", "O19"), Type:=xlFillDefault)

        Ws.Cells(21, 1) = "Cash Flows from Investing Activities:"
        Ws.Cells(22, 1) = "    Asset  sales "
        Ws.Cells(22, 3) = "USD$k"
        Ws.Cells(22, 8) = "=F22-G22"
        oRng = Ws.Range("H22", "H22")
        oRng.AutoFill(Destination:=Ws.Range("H22", "H24"), Type:=xlFillDefault)
        Ws.Cells(22, 12) = "=J22-K22"

        oRng = Ws.Range("L22", "L22")
        oRng.AutoFill(Destination:=Ws.Range("L22", "L24"), Type:=xlFillDefault)
        Ws.Cells(23, 1) = "    Interest income "
        Ws.Cells(23, 3) = "USD$k"
        Ws.Cells(24, 1) = "    Investments "
        Ws.Cells(24, 3) = "USD$k"
        Ws.Cells(25, 1) = "Cash Flows from Investing Activities"
        Ws.Cells(25, 3) = "USD$k"
        Ws.Cells(25, 4) = "=SUM(D22:D24)"
        oRng = Ws.Range("D25", "D25")
        oRng.AutoFill(Destination:=Ws.Range("D25", "O25"), Type:=xlFillDefault)

        Ws.Cells(27, 1) = "Free Cash Flow"
        Ws.Cells(27, 3) = "USD$k"
        Ws.Cells(27, 4) = "=D19+D25"
        oRng = Ws.Range("D27", "D27")
        oRng.AutoFill(Destination:=Ws.Range("D27", "O27"), Type:=xlFillDefault)

        Ws.Cells(29, 1) = "Cash Flows from Financing Activities:"
        Ws.Cells(30, 1) = "    paid interest"
        Ws.Cells(30, 3) = "USD$k"
        Ws.Cells(30, 8) = "=F30-G30"
        oRng = Ws.Range("H30", "H30")
        oRng.AutoFill(Destination:=Ws.Range("H30", "H35"), Type:=xlFillDefault)
        Ws.Cells(30, 12) = "=J30-K30"
        oRng = Ws.Range("L30", "L30")
        oRng.AutoFill(Destination:=Ws.Range("L30", "L35"), Type:=xlFillDefault)

        Ws.Cells(31, 1) = "    Changes in equity"
        Ws.Cells(31, 3) = "USD$k"
        Ws.Cells(32, 1) = "    Changes in revolver"
        Ws.Cells(32, 3) = "USD$k"
        Ws.Cells(33, 1) = "    Changes in bank loans"
        Ws.Cells(33, 3) = "USD$k"
        Ws.Cells(34, 1) = "    Changes in sale-and-lease back"
        Ws.Cells(34, 3) = "USD$k"
        Ws.Cells(35, 1) = "    Changes in shareholder loan"
        Ws.Cells(35, 3) = "USD$k"
        Ws.Cells(36, 1) = "Cash Flows from Financing Activities"
        Ws.Cells(36, 3) = "USD$k"
        Ws.Cells(36, 4) = "=SUM(D30:D35)"
        oRng = Ws.Range("D36", "D36")
        oRng.AutoFill(Destination:=Ws.Range("D36", "O36"), Type:=xlFillDefault)

        Ws.Cells(38, 1) = "Total Cash Flow"
        Ws.Cells(38, 3) = "USD$k"
        Ws.Cells(38, 4) = "=D27+D36"
        oRng = Ws.Range("D38", "D38")
        oRng.AutoFill(Destination:=Ws.Range("D38", "O38"), Type:=xlFillDefault)

        Ws.Cells(40, 1) = "Beginning Cash Balance"
        Ws.Cells(40, 3) = "USD$k"
        Ws.Cells(40, 6) = "=E41"
        Ws.Cells(40, 8) = "=F40-G40"
        Ws.Cells(40, 10) = "=F40"
        Ws.Cells(40, 12) = "=J40-K40"
        Ws.Cells(40, 13) = "=I40"
        Ws.Cells(40, 14) = "=O40"
        Ws.Cells(40, 15) = "=K40"

        Ws.Cells(41, 1) = "Ending Cash Balance"
        Ws.Cells(41, 3) = "USD$k"
        Ws.Cells(41, 4) = "=D38+D40"
        Ws.Cells(41, 5) = "=E38+E40"
        Ws.Cells(41, 6) = "=F38+F40"
        Ws.Cells(41, 7) = "=G38+G40"
        Ws.Cells(41, 8) = "=F41-G41"
        Ws.Cells(41, 9) = "=I38+I40"
        Ws.Cells(41, 10) = "=J38+J40"
        Ws.Cells(41, 11) = "=K38+K40"
        Ws.Cells(41, 12) = "=J41-K41"
        Ws.Cells(41, 13) = "=M38+M40"
        Ws.Cells(41, 14) = "=N38+N40"
        Ws.Cells(41, 15) = "=O38+O40"

        oRng = Ws.Range("A7", "O41")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("D9", "O41")
        oRng.NumberFormatLocal = "#,##0,"

        LineZ = 10
    End Sub
    Private Function GetAAH(ByVal aah01 As String, sYear As Int16, sMonth As Int16, ByVal sDirection As Int16)
        oCommand2.CommandText = "select nvl(sum(aah05 -aah04),0) from aah_file where aah01 in (" & aah01 & ") and aah02 = " & sYear & " and aah03 = " & sMonth
        Dim AAS As Decimal = 0
        If sDirection = 0 Then
            AAS = oCommand2.ExecuteScalar() * Decimal.MinusOne
        Else
            AAS = oCommand2.ExecuteScalar()
        End If
        Return AAS
    End Function
    Private Sub DoInputData(ByVal aah01 As String, ByVal sDirection As Int16)
        Ws.Cells(LineZ, 4) = GetAAH(aah01, pYear, tMonth, sDirection)
        Ws.Cells(LineZ, 5) = GetAAH(aah01, lYear, pMonth, sDirection)
        Ws.Cells(LineZ, 6) = GetAAH(aah01, tYear, tMonth, sDirection)
    End Sub
    Private Sub DoInputData(ByVal aah01_1 As String, ByVal sDirection1 As Int16, ByVal aah01_2 As String, sDirection2 As Int16)
        Ws.Cells(LineZ, 4) = GetAAH(aah01_1, pYear, tMonth, sDirection1) + GetAAH(aah01_2, pYear, tMonth, sDirection2)
        Ws.Cells(LineZ, 5) = GetAAH(aah01_1, lYear, pMonth, sDirection1) + GetAAH(aah01_2, lYear, pMonth, sDirection2)
        Ws.Cells(LineZ, 6) = GetAAH(aah01_1, tYear, tMonth, sDirection1) + GetAAH(aah01_2, tYear, tMonth, sDirection2)

    End Sub
    Private Sub DoInputData1(ByVal aah01_1 As String, ByVal aah01_2 As String, ByVal sDirection As Int16)
        Ws.Cells(LineZ, 4) = GetAAH1(aah01_1, aah01_2, pYear, tMonth, sDirection)
        Ws.Cells(LineZ, 5) = GetAAH1(aah01_1, aah01_2, lYear, pMonth, sDirection)
        Ws.Cells(LineZ, 6) = GetAAH1(aah01_1, aah01_2, tYear, tMonth, sDirection)
    End Sub
    Private Function GetAAH1(ByVal aah01_1 As String, ByVal aah01_2 As String, sYear As Int16, sMonth As Int16, ByVal sDirection As Int16)
        oCommand2.CommandText = "select nvl(sum(aah05 -aah04),0) from aah_file,aag_file where aah01=aag01 and aag07 in ('2','3') and aah01 between '" & aah01_1 & "' and '" & aah01_2 & "' and aah02 = " & sYear & " and aah03 = " & sMonth
        Dim AAS As Decimal = 0
        If sDirection = 0 Then
            AAS = oCommand2.ExecuteScalar() * Decimal.MinusOne
        Else
            AAS = oCommand2.ExecuteScalar()
        End If
        Return AAS
    End Function
    Private Sub DoInputData2(ByVal aah01_1 As String, ByVal aah01_2 As String, ByVal sDirection As Int16)
        Ws.Cells(LineZ, 4) = GetAAH2(aah01_1, aah01_2, pYear, tMonth, sDirection)
        Ws.Cells(LineZ, 5) = GetAAH2(aah01_1, aah01_2, lYear, pMonth, sDirection)
        Ws.Cells(LineZ, 6) = GetAAH2(aah01_1, aah01_2, tYear, tMonth, sDirection)
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "CashFlow_" & gDatabase
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
    Private Function GetAAH2(ByVal aah01_1 As String, ByVal aah01_2 As String, sYear As Int16, sMonth As Int16, ByVal sDirection As Int16)
        oCommand2.CommandText = "select nvl(sum(aah05 -aah04),0) from aah_file,aag_file where aah01=aag01 and aag07 in ('2','3') and aah01 between '" & aah01_1 & "' and '" & aah01_2 & "' and aah02 = " & sYear & " and aah03 < " & sMonth
        Dim AAS As Decimal = 0
        If sDirection = 0 Then
            AAS = oCommand2.ExecuteScalar() * Decimal.MinusOne
        Else
            AAS = oCommand2.ExecuteScalar()
        End If
        Return AAS
    End Function
End Class