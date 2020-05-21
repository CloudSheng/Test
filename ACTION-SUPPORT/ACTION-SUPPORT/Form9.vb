Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form9
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    'Dim oConnection As 
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim tYear As String = String.Empty
    Dim tMonth As String = String.Empty
    Dim pYear As String = String.Empty
    Dim pMonth As String = String.Empty
    Dim AValueMonth0 As Decimal = 0
    Dim BValueMonth0 As Decimal = 0
    Dim CValueMonth0 As Decimal = 0
    Dim DValueMonth1 As Decimal = 0
    Dim EValueMonth1 As Decimal = 0
    Dim DValueMonthN As Decimal = 0
    Dim EValueMonthN As Decimal = 0
    Dim ACAAP1 As Decimal = 0
    Dim ACAAPN As Decimal = 0
    Dim ACAAR1 As Decimal = 0
    Dim ACAARN As Decimal = 0
    Dim WnStart As Date
    Dim WnEnd As Date
    Dim MStart As Date
    'Dim Month1End As Date
    'Dim MonthNStart As Date
    'Dim MonthNEnd As Date
    Dim DS As Data.DataSet = New DataSet()
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form9_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")

        'If String.IsNullOrEmpty(TextBox1.Text) Then
        'TextBox1.Text = Module1.GetYearAndMonthString(Date.Today)
        'End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If Label2.Text <> "AP档案读入" Or Label3.Text <> "AR档案读入" Then
            MsgBox("档案有误")
            Return
        End If
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        tYear = DateTimePicker1.Value.Year
        tMonth = DateTimePicker1.Value.Month
        If tMonth = 1 Then
            pYear = tYear - 1
            pMonth = "12"
        Else
            pYear = tYear
            pMonth = tMonth - 1
            If Strings.Len(pMonth) = 1 Then
                pMonth = "0" & pMonth
            End If
        End If
        If Strings.Len(tMonth) = 1 Then
            tMonth = "0" & tMonth
        End If
        MStart = Convert.ToDateTime(tYear & "/" & tMonth & "/01")
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Dull_Report"
        SaveFileDialog1.DefaultExt = ".xls"
        Dim SON As DialogResult = SaveFileDialog1.ShowDialog()
        If SON = DialogResult.OK Then
            Dim SFN As String = SaveFileDialog1.FileName
            Ws.SaveAs(SFN, XlFileFormat.xlExcel12)
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
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        LineZ = 4
        oCommand.CommandText = "SELECT nvl(azj041,0) FROM azj_file WHERE azj01 = 'USD' AND azj02 = '" & tYear & tMonth & "'"
        Dim USDRate As Decimal = 0
        Try
            USDRate = oCommand.ExecuteScalar()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        If USDRate = 0 Then
            MsgBox("USD汇率出错")
            Return
        End If
        oCommand.CommandText = "SELECT nvl(azj041,0) FROM hkacttest.azj_file WHERE azj01 = 'EUR' AND azj02 = '" & tYear & tMonth & "'"
        Dim EURTOUSDRate As Decimal = 0
        Try
            EURTOUSDRate = oCommand.ExecuteScalar()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        If EURTOUSDRate = 0 Then
            MsgBox("EUR汇率出错")
            Return
        End If
        oCommand.CommandText = "SELECT NVL(SUM(NMP09),0) FROM NMP_FILE WHERE nmp02 = '" _
                               & pYear & "' and nmp03 = '" & pMonth & "'"
        AValueMonth0 = oCommand.ExecuteScalar() ' / USDRate
        oCommand.CommandText = "select nvl(sum(case when npk01 = '1' then npk09 "
        oCommand.CommandText += " when npk01 = '2' then npk09 * -1 end),0)  as t1 "
        oCommand.CommandText += " from nmg_file,npk_file where nmg01 between to_date('"
        oCommand.CommandText += MStart.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and nmgconf = 'Y' and nmg00 = npk00 "
        AValueMonth0 += oCommand.ExecuteScalar()
        AValueMonth0 = AValueMonth0 / USDRate
        Ws.Cells(5, 2) = AValueMonth0
        Ws.Cells(9, 2) = "=SUM(B5:B8)"
        'oCommand.CommandText = "SELECT NVL(SUM(NMP09),0) FROM hkacttest.NMP_FILE WHERE nmp02 = '" _
        '                       & pYear & "' and nmp03 = '" & pMonth & "'"
        'BValueMonth0 = oCommand.ExecuteScalar()
        'Ws.Cells(6, 2) = BValueMonth0
        'TO -DO ACA BANK
        'CValueMonth0 = 0
        ' TO-DO END
        'Ws.Cells(7, 2) = CValueMonth0
        'Month1End = Convert.ToDateTime(tYear & "/" & tMonth & "/01").AddMonths(1).AddDays(Decimal.MinusOne)
        If Not IsDBNull(DS.Tables(1).Compute("sum(balance_amount)", "due_date <= '" & DateTimePicker1.Value.ToString() & "'")) Then
            ACAAR1 = DS.Tables(1).Compute("sum(balance_amount)", "due_date <= '" & DateTimePicker1.Value.ToString() & "'")
        Else
            ACAAR1 = 0
        End If
        Ws.Cells(11, 2) = ACAAR1 / EURTOUSDRate
        oCommand.CommandText = "select sum(t1) from ( " _
                             & " SELECT nvl(sum(apc13),0) as t1 FROM APA_FILE,APC_FILE WHERE APA01 = APC01 AND APC13 > 0 AND APA41  = 'Y' and apa00 in (11,15,16) " _
                             & " and apa12 <= to_date('" & DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') " _
                             & " union all " _
                             & "SELECT nvl(sum(rvw05),0) FROM rvw_file WHERE ta_rvw04 is null and ta_rvw01 <= to_date('" & DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') )"

        DValueMonth1 = oCommand.ExecuteScalar()
        Ws.Cells(14, 2) = DValueMonth1 / USDRate
        oCommand.CommandText = "SELECT nvl(sum(apc13),0) as t1 FROM APA_FILE,APC_FILE WHERE APA01 = APC01 AND APC13 > 0 AND APA41  = 'Y' and apa00 = 12 and apa12 <= to_date('" & DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        EValueMonth1 = oCommand.ExecuteScalar()
        Ws.Cells(15, 2) = EValueMonth1 / USDRate
        Ws.Cells(17, 2) = "=SUM(B14:B16)"
        'TO -DO ACA EUR/ USD
        If Not IsDBNull(DS.Tables(0).Compute("sum(balance_amount)", "due_date <= '" & DateTimePicker1.Value.ToString() & "'")) Then
            ACAAP1 = DS.Tables(0).Compute("sum(balance_amount)", "due_date <= '" & DateTimePicker1.Value.ToString() & "'")
        Else
            ACAAP1 = 0
        End If
        Ws.Cells(18, 2) = ACAAP1 / EURTOUSDRate
        'Ws.Cells(15, 2) = 0
        'Ws.Cells(16, 2) = ACAAP1 / EURTOUSDRate
        'Ws.Cells(17, 2) = "=B5+B6+B7+B9-B11-B12-B14-B15"
        Ws.Cells(20, 2) = "=SUM(B18:B19)"
        Ws.Cells(21, 2) = "=B9+B11-B17-B20"
        'TO -DO ACA AP
        ' 詢環處理
        For j As Int16 = 0 To 14 Step 1
            Dim ExcelColumn As String = String.Empty
            Select Case j
                Case 0
                    ExcelColumn = "B"
                Case 1
                    ExcelColumn = "C"
                Case 2
                    ExcelColumn = "D"
                Case 3
                    ExcelColumn = "E"
                Case 4
                    ExcelColumn = "F"
                Case 5
                    ExcelColumn = "G"
                Case 6
                    ExcelColumn = "H"
                Case 7
                    ExcelColumn = "I"
                Case 8
                    ExcelColumn = "J"
                Case 9
                    ExcelColumn = "K"
                Case 10
                    ExcelColumn = "L"
                Case 11
                    ExcelColumn = "M"
                Case 12
                    ExcelColumn = "N"
                Case 13
                    ExcelColumn = "O"
                Case 14
                    ExcelColumn = "P"
            End Select
            WnStart = DateTimePicker1.Value.AddDays(1 + 7 * j)
            WnEnd = WnStart.AddDays(6)

            'MonthNStart = Convert.ToDateTime(tYear & "/" & tMonth & "/01").AddMonths(j)
            'MonthNEnd = MonthNStart.AddMonths(1).AddDays(Decimal.MinusOne)
            Ws.Cells(5, 3 + j) = "=" & ExcelColumn & "5+" & ExcelColumn & "10-" & ExcelColumn & "17"
            'Ws.Cells(6, 2 + j) = "=B6"
            'Ws.Cells(7, 2 + j) = "=" & ExcelColumn & "7+" & ExcelColumn & "9-" & ExcelColumn & "16"
            If Not IsDBNull(DS.Tables(1).Compute("sum(balance_amount)", "due_date <= '" & WnEnd.ToString() & "' and due_date >= '" & WnStart.ToString() & "'")) Then
                ACAARN = DS.Tables(1).Compute("sum(balance_amount)", "due_date <= '" & WnEnd.ToString() & "' and due_date >= '" & WnStart.ToString() & "'")
            Else
                ACAARN = 0
            End If
            Ws.Cells(11, 3 + j) = ACAARN / EURTOUSDRate
            oCommand.CommandText = "select sum(t1) from ( " _
                             & " SELECT nvl(sum(apc13),0) as t1 FROM APA_FILE,APC_FILE WHERE APA01 = APC01 AND APC13 > 0 AND APA41  = 'Y' and apa00 in (11,15,16) " _
                             & " and apa12 between to_date('" & WnStart.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & WnEnd.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') " _
                             & " union all " _
                             & "SELECT nvl(sum(rvw05),0) FROM rvw_file WHERE ta_rvw04 is null and ta_rvw01 between to_date('" & WnStart.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & WnEnd.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') )"
            DValueMonthN = oCommand.ExecuteScalar()
            Ws.Cells(14, 3 + j) = DValueMonthN / USDRate
            oCommand.CommandText = "SELECT nvl(sum(apc13),0) as t1 FROM APA_FILE,APC_FILE WHERE APA01 = APC01 AND APC13 > 0 AND APA41  = 'Y' and apa00 = 12 and apa12 between to_date('" & WnStart.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & WnEnd.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
            EValueMonthN = oCommand.ExecuteScalar()
            Ws.Cells(15, 3 + j) = EValueMonthN / USDRate
            'Ws.Cells(13, 2 + j) = (DValueMonthN + EValueMonthN) / USDRate
            If Not IsDBNull(DS.Tables(0).Compute("sum(balance_amount)", "due_date <= '" & WnEnd.ToString() & "' and due_date >= '" & WnStart.ToString() & "'")) Then
                ACAAPN = DS.Tables(0).Compute("sum(balance_amount)", "due_date <= '" & WnEnd.ToString() & "' and due_date >= '" & WnStart.ToString() & "'")
            Else
                ACAAPN = 0
            End If

            ' TO-DO EUR/USD
            Ws.Cells(18, 2 + j) = ACAAPN / EURTOUSDRate
            'Ws.Cells(15, 2 + j) = 0
            'Ws.Cells(16, 2 + j) = ACAAPN / EURTOUSDRate
        Next
        oRng = Ws.Range("B9:B9")
        oRng.AutoFill(Destination:=Ws.Range("B9:Q9"), Type:=xlFillDefault)
        oRng = Ws.Range("B17:B17")
        oRng.AutoFill(Destination:=Ws.Range("B17:Q17"), Type:=xlFillDefault)
        oRng = Ws.Range("B21:B21")
        oRng.AutoFill(Destination:=Ws.Range("B21:Q21"), Type:=xlFillDefault)
        oRng = Ws.Range("B20:B20")
        oRng.AutoFill(Destination:=Ws.Range("B20:Q20"), Type:=xlFillDefault)
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 20
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 26.88
        oRng = Ws.Range("A1", "G1")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A2", "G2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A3", "G3")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        'Ws.Cells(1, 1) = "东莞艾可讯复合材料有限公司"
        Ws.Cells(1, 1) = "Action Composites"
        Ws.Cells(2, 1) = "Month Cash Flow Forecast"
        Ws.Cells(4, 1) = "Begin Bank Balance"
        Ws.Cells(5, 1) = "DG AC"
        Ws.Cells(6, 1) = "HK AC"
        Ws.Cells(7, 1) = "ACA AC"
        Ws.Cells(8, 1) = "BVI AC"
        Ws.Cells(9, 1) = "Beginning Group Bank balance"
        Ws.Cells(10, 1) = "HK AR"
        Ws.Cells(11, 1) = "AR-ACA 应收帐款"
        Ws.Cells(13, 1) = "AP"
        Ws.Cells(14, 1) = "DG-AP 应付帐款"
        Ws.Cells(15, 1) = "DG -Other payable杂项应付"
        Ws.Cells(16, 1) = "DG  monthly fixed expenditure 固定支出"
        Ws.Cells(17, 1) = "Total-DG payables"
        Ws.Cells(18, 1) = "ACA -AP应付帐款 "
        Ws.Cells(19, 1) = "ACA monthly fixed expenditure 固定支出"
        Ws.Cells(20, 1) = "Total-ACA"
        Ws.Cells(21, 1) = "Net Cash Flow"
        For i As Int16 = 0 To 15 Step 1
            'Dim Ct As Int16 = tMonth + i
            'If Ct > 12 Then
            'Ct = Ct - 12
            'End If
            'Ws.Cells(4, i + 2) = Module1.GetMonthEnglish(Ct)
            Ws.Cells(4, 2 + i) = "W" & i
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Try
                Excelconn.Open()
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            Dim ExcelString = "SELECT due_date,balance_amount FROM [Sheet1$]"
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)

            Try
                ExcelAdapater.Fill(DS, "table1")
                Me.Label2.Text = "AP档案读入"
            Catch ex As Exception
                MsgBox(ex.Message())
                Me.Label2.Text = "AP档案读入失败"
                Return
            End Try
            Dim ExcelString1 = "SELECT due_date,balance_amount FROM [Sheet2$]"
            Dim ExcelAdapater1 As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString1, Excelconn)

            Try
                ExcelAdapater1.Fill(DS, "table2")
                Me.Label3.Text = "AR档案读入"
            Catch ex As Exception
                MsgBox(ex.Message())
                Me.Label3.Text = "AR档案读入失败"
                Return
            End Try

            'Dim PS1 As Decimal = DS.Tables(0).Compute("sum(balance_amount)", "due_date <= '2015/08/31' and due_date >= '2015/08/01'")
            'MsgBox(PS1)
        End If
    End Sub
End Class