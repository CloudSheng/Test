Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel.XlChartType
Public Class Form54
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    'Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim CC As Integer = 0
    Dim DStartN As Date
    Dim LineZ As Integer = 0
    Dim LineX As Integer = 0
    Dim DW1 As Integer = 0
    Dim DW2 As Integer = 0
    Dim DW3 As Integer = 0
    Dim DW4 As Integer = 0
    Dim DS As Data.DataSet = New DataSet()
    Dim TimeS1 As DateTime
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString = "SELECT * FROM [011~25SG$]"
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)
            Try
                ExcelAdapater.Fill(DS, "ACAAR")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try

            ExcelString = "SELECT * FROM [011~25SG$],[ACA_customer_list$] where [011~25SG$].customer_nr = [ACA_customer_list$].NO "
            ExcelAdapater.SelectCommand.CommandText = ExcelString
            Try
                ExcelAdapater.Fill(DS, "table2")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            Excelconn.Close()
        End If
        If IsDBNull(DS.Tables("ACAAR")) Then
            Label1.Text = "读取失败"
            Me.GroupBox2.Enabled = False
        Else
            Label1.Text = "已读入"
            Me.GroupBox2.Enabled = True
        End If
    End Sub

    Private Sub Form54_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.GroupBox2.Enabled = False
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
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
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "ACA_overdue_AR_aging_report_" & TimeS1.ToString("yyyyMMdd")
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

    End Sub
    Private Sub ExportToExcel()
        TimeS1 = Me.DateTimePicker1.Value.ToString("yyyy/MM/dd")
        'DS.Clear()
        DStartN = Today()
        '先訂位
        'DW1 = GetAzn02(DStartN)
        'DW2 = GetAzn05(DStartN)
        DW1 = GetAzn02(TimeS1)
        DW2 = GetAzn05(TimeS1)
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        xWorkBook.Sheets.Add()  '第四頁
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        Ws.Name = "ACAAR"
        AdjustExcelFormat()
        For i As Integer = 0 To DS.Tables("ACAAR").Rows.Count - 1 Step 1
            If IsDBNull(DS.Tables("ACAAR").Rows(i).Item(0)) Then
                Continue For
            End If
            If DS.Tables("ACAAR").Rows(i).Item("due_date") > TimeS1 Then
                Continue For
            End If
            Ws.Cells(LineZ, 1) = DS.Tables("ACAAR").Rows(i).Item(0)
            Ws.Cells(LineZ, 2) = DS.Tables("ACAAR").Rows(i).Item(1)
            Ws.Cells(LineZ, 3) = DS.Tables("ACAAR").Rows(i).Item(2)
            Ws.Cells(LineZ, 4) = DS.Tables("ACAAR").Rows(i).Item(3)
            Ws.Cells(LineZ, 5) = DS.Tables("ACAAR").Rows(i).Item(4)
            Ws.Cells(LineZ, 6) = DS.Tables("ACAAR").Rows(i).Item(5)
            Ws.Cells(LineZ, 7) = DS.Tables("ACAAR").Rows(i).Item(6)
            Ws.Cells(LineZ, 8) = DS.Tables("ACAAR").Rows(i).Item(7)
            'Dim ER As Decimal = 0
            If Not IsDBNull(DS.Tables("ACAAR").Rows(i).Item(5)) Then
                'ER = GetExchangeRate(DS.Tables("ACAAR").Rows(i).Item(5))
            Else
                LineZ += 1
                Continue For
            End If
            If DS.Tables("ACAAR").Rows(i).Item(5) <= TimeS1 Then
                'Ws.Cells(LineZ, 9) = DS.Tables("ACAAR").Rows(i).Item(7) * ER
                Ws.Cells(LineZ, 9) = DS.Tables("ACAAR").Rows(i).Item(7)
            Else
                Setwk(DS.Tables("ACAAR").Rows(i).Item(5))
                If DW1 = DW3 Then
                    'Ws.Cells(LineZ, 9 + DW4 - DW2) = DS.Tables("ACAAR").Rows(i).Item(7) * ER
                    Ws.Cells(LineZ, 9 + DW4 - DW2) = DS.Tables("ACAAR").Rows(i).Item(7)
                Else
                    'Ws.Cells(LineZ, 9 + 52 + DW4 - DW2) = DS.Tables("ACAAR").Rows(i).Item(7) * ER
                    Ws.Cells(LineZ, 9 + 52 + DW4 - DW2) = DS.Tables("ACAAR").Rows(i).Item(7)
                End If
            End If
            LineZ += 1
        Next
        ' 加總
        Ws.Cells(LineZ, 1) = "Total_forecast_AR_EUR"
        Ws.Cells(LineZ, 9) = "=SUM(I2:I" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9))
        oRng.AutoFill(Destination:=Ws.Range("I" & LineZ & ":DB" & LineZ), Type:=xlFillDefault)

        ' 加總美元
        Dim ER As Decimal = 0
        'ER = GetExchangeRate(DStartN)
        ER = GetExchangeRate(TimeS1)
        Ws.Cells(LineZ + 1, 1) = "Total_forecast_AR_USD"
        Ws.Cells(LineZ + 1, 9) = "=I$" & LineZ & "*" & ER
        oRng = Ws.Range(Ws.Cells(LineZ + 1, 9), Ws.Cells(LineZ + 1, 9))
        oRng.EntireRow.NumberFormatLocal = "_-""US$""* #,##0.00_ ;_-""US$""* -#,##0.00 ;_-""US$""* ""-""??_ ;_-@_ "
        oRng.AutoFill(Destination:=Ws.Range("I" & LineZ + 1 & ":DB" & LineZ + 1), Type:=xlFillDefault)

        ' 第二頁
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        Ws.Name = "Overdue AR detail"
        AdjustExcelFormat1()

        For i As Integer = 0 To DS.Tables("table2").Rows.Count - 1 Step 1
            If IsDBNull(DS.Tables("table2").Rows(i).Item(0)) Then
                Continue For
            End If
            If DS.Tables("table2").Rows(i).Item("due_date") > TimeS1 Then
                Continue For
            End If
            Ws.Cells(LineZ, 1) = DS.Tables("table2").Rows(i).Item("customer_nr")
            Ws.Cells(LineZ, 2) = DS.Tables("table2").Rows(i).Item("Text")
            Ws.Cells(LineZ, 3) = DS.Tables("table2").Rows(i).Item("invoice_nr")
            Ws.Cells(LineZ, 4) = DS.Tables("table2").Rows(i).Item("invoice_date")
            Ws.Cells(LineZ, 5) = DS.Tables("table2").Rows(i).Item("due_date")
            Ws.Cells(LineZ, 6) = DS.Tables("table2").Rows(i).Item("due_balance")
            Ws.Cells(LineZ, 8) = DS.Tables("table2").Rows(i).Item("customer")
            Ws.Cells(LineZ, 9) = DS.Tables("table2").Rows(i).Item("payment_terms_day")
            Dim DD2 As Decimal = DateDiff(DateInterval.Day, Convert.ToDateTime(DS.Tables("table2").Rows(i).Item("due_date")), TimeS1)
            Ws.Cells(LineZ, 10) = DD2
            Select Case DD2
                Case Is < 0
                    Ws.Cells(LineZ, 11) = DS.Tables("table2").Rows(i).Item("due_balance")
                Case 0 To 29
                    'oRng = Ws.Range(Ws.Cells(LineZ, 10), Ws.Cells(LineZ, 10))
                    'oRng.Interior.Color = Color.Red
                    Ws.Cells(LineZ, 12) = DS.Tables("table2").Rows(i).Item("due_balance")
                Case 30 To 59
                    oRng = Ws.Range(Ws.Cells(LineZ, 10), Ws.Cells(LineZ, 10))
                    oRng.Interior.Color = Color.LightGreen
                    Ws.Cells(LineZ, 13) = DS.Tables("table2").Rows(i).Item("due_balance")
                Case 60 To 89
                    oRng = Ws.Range(Ws.Cells(LineZ, 10), Ws.Cells(LineZ, 10))
                    oRng.Interior.Color = Color.Yellow
                    Ws.Cells(LineZ, 14) = DS.Tables("table2").Rows(i).Item("due_balance")
                Case 90 To 119
                    oRng = Ws.Range(Ws.Cells(LineZ, 10), Ws.Cells(LineZ, 10))
                    oRng.Interior.Color = Color.Orange
                    Ws.Cells(LineZ, 15) = DS.Tables("table2").Rows(i).Item("due_balance")
                Case Is > 120
                    oRng = Ws.Range(Ws.Cells(LineZ, 10), Ws.Cells(LineZ, 10))
                    oRng.Interior.Color = Color.Red
                    Ws.Cells(LineZ, 16) = DS.Tables("table2").Rows(i).Item("due_balance")
            End Select
            LineZ += 1
        Next
        ' 加總
        Ws.Cells(LineZ, 5) = "Total AR amount_EUR"
        Ws.Cells(LineZ, 6) = "=SUM(F2:F" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 11) = "=SUM(K2:K" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 12) = "=SUM(L2:L" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 13) = "=SUM(M2:M" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 14) = "=SUM(N2:N" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 15) = "=SUM(O2:O" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 16) = "=SUM(P2:P" & LineZ - 1 & ")"
        ' 加總美元
        Ws.Cells(LineZ + 1, 5) = "Total AR amount_USD"
        Ws.Cells(LineZ + 1, 6) = "=F$" & LineZ & "*" & ER
        oRng = Ws.Range(Ws.Cells(LineZ + 1, 6), Ws.Cells(LineZ + 1, 6))
        oRng.EntireRow.NumberFormatLocal = "_-""US$""* #,##0.00_ ;_-""US$""* -#,##0.00 ;_-""US$""* ""-""??_ ;_-@_ "
        Ws.Cells(LineZ + 1, 11) = "=K$" & LineZ & "*" & ER
        Ws.Cells(LineZ + 1, 12) = "=L$" & LineZ & "*" & ER
        Ws.Cells(LineZ + 1, 13) = "=M$" & LineZ & "*" & ER
        Ws.Cells(LineZ + 1, 14) = "=N$" & LineZ & "*" & ER
        Ws.Cells(LineZ + 1, 15) = "=O$" & LineZ & "*" & ER
        Ws.Cells(LineZ + 1, 16) = "=P$" & LineZ & "*" & ER
        oRng = Ws.Range(Ws.Cells(LineZ + 1, 11), Ws.Cells(LineZ + 1, 16))
        oRng.EntireRow.NumberFormatLocal = "_-""US$""* #,##0.00_ ;_-""US$""* -#,##0.00 ;_-""US$""* ""-""??_ ;_-@_ "

        ' 第三頁
        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        Ws.Name = "Overdue AR overview"
        AdjustExcelFormat2()
        Ws.Cells(2, 2) = "='Overdue AR detail'!P" & LineZ + 1
        Ws.Cells(3, 2) = "='Overdue AR detail'!O" & LineZ + 1
        Ws.Cells(4, 2) = "='Overdue AR detail'!N" & LineZ + 1
        Ws.Cells(5, 2) = "='Overdue AR detail'!M" & LineZ + 1
        Ws.Cells(6, 2) = "='Overdue AR detail'!L" & LineZ + 1
        Ws.Cells(7, 2) = "='Overdue AR detail'!K" & LineZ + 1
        Ws.Cells(8, 2) = "=SUM(B2:B6)"

        Dim XShape As Microsoft.Office.Interop.Excel.Shape = Ws.Shapes.AddChart(xlLineMarkers, 330, 30, 550, 250)
        XShape.Chart.ChartType = xlLineMarkers
        XShape.Chart.SetSourceData(Ws.Range("A1", "B7"))
        XShape.Chart.ChartTitle.Text = "overdue amount by days"
        'XShape.Chart.Name = "Chart1"

        ' 處理下方資料
        oCommand.CommandText = "CREATE Table AgingTemp (customer varchar2(255), aamount number(15,2), iDate date)"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        For i As Integer = 0 To DS.Tables("table2").Rows.Count - 1 Step 1
            If IsDBNull(DS.Tables("table2").Rows(i).Item(0)) Then
                Continue For
            End If
            If DS.Tables("table2").Rows(i).Item("due_date") > TimeS1 Then
                Continue For
            End If
            oCommand.CommandText = "INSERT INTO AgingTemp Values ('" & DS.Tables("table2").Rows(i).Item("customer")
            oCommand.CommandText += "'," & DS.Tables("table2").Rows(i).Item("due_balance") * ER & ",to_date('" & DS.Tables("table2").Rows(i).Item("due_date") & "','yyyy/mm/dd'))"
            Try
                oCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        Next
        oCommand.CommandText = "SELECT sum(aamount) as t1,customer FROM AgingTemp WHERE 1 =1 group by customer Order by customer"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineX, 1) = oReader.Item("customer")
                Ws.Cells(LineX, 2) = oReader.Item("t1")
                LineX += 1
            End While
        End If
        oReader.Close()

        ' 加總

        Ws.Cells(LineX, 1) = "Total overdue AR_USD"
        Ws.Cells(LineX, 2) = "=SUM(B10:B" & LineX - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineX, 1), Ws.Cells(LineX, 2))
        oRng.Interior.Color = Color.Yellow
        ' 作圖
        Dim XShape2 As Microsoft.Office.Interop.Excel.Shape = Ws.Shapes.AddChart(xlColumnClustered, 330, 300, 700, 400)
        XShape2.Chart.ChartType = xlColumnClustered
        XShape2.Chart.SetSourceData(Ws.Range("A9", Ws.Cells(LineX - 1, 2)))
        XShape2.Chart.ChartTitle.Text = "overdue amount by customer"

        ' 第四頁   20160829
        Ws = xWorkBook.Sheets(4)
        Ws.Activate()
        Ws.Name = "Customer VS Overdue days"
        AdjustExcelFormat3()
        oCommand.CommandText = "select customer,sum(t0) as t0,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5 from ("
        oCommand.CommandText += "select customer,"
        oCommand.CommandText += "(Case when to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') - idate < 0 then aamount end) as t0,"
        oCommand.CommandText += "(Case when to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') - idate >= 0 and to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') - idate < 30 then aamount end) as t1,"
        oCommand.CommandText += "(Case when to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') - idate >= 30 and to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') - idate < 60 then aamount end) as t2,"
        oCommand.CommandText += "(Case when to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') - idate >= 60 and to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') - idate < 90 then aamount end) as t3,"
        oCommand.CommandText += "(Case when to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') - idate >= 90 and to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') - idate < 120 then aamount end) as t4,"
        oCommand.CommandText += "(Case when to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') - idate >= 120 then aamount end) as t5 from AgingTemp ) group by customer"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineX, 1) = oReader.Item("Customer")
                Ws.Cells(LineX, 2) = oReader.Item("t0")
                Ws.Cells(LineX, 3) = oReader.Item("t1")
                Ws.Cells(LineX, 4) = oReader.Item("t2")
                Ws.Cells(LineX, 5) = oReader.Item("t3")
                Ws.Cells(LineX, 6) = oReader.Item("t4")
                Ws.Cells(LineX, 7) = oReader.Item("t5")
                LineX += 1
            End While
        End If
        oReader.Close()
        ' 加總
        Ws.Cells(LineX, 1) = "Total AR_USD"
        Ws.Cells(LineX, 2) = "=SUM(B2:B" & LineX - 1 & ")"
        Ws.Cells(LineX, 3) = "=SUM(C2:C" & LineX - 1 & ")"
        Ws.Cells(LineX, 4) = "=SUM(D2:D" & LineX - 1 & ")"
        Ws.Cells(LineX, 5) = "=SUM(E2:E" & LineX - 1 & ")"
        Ws.Cells(LineX, 6) = "=SUM(F2:F" & LineX - 1 & ")"
        Ws.Cells(LineX, 7) = "=SUM(G2:G" & LineX - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineX, 1), Ws.Cells(LineX, 7))
        oRng.Interior.Color = Color.Yellow
        ' 關閉臨時檔

        oCommand.CommandText = "DROP TABLE AgingTemp"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        oConnection.Close()
    End Sub
    Private Function GetAzn02(ByVal eDate As Date)
        Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander2.Connection = oConnection
        oCommander2.CommandType = CommandType.Text
        oCommander2.CommandText = "SELECT azn02 FROM azn_file where azn01 = TO_DATE('" & eDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        Dim ADW1 As Integer = oCommander2.ExecuteScalar()
        Return ADW1
    End Function
    Private Function GetAzn05(ByVal eDate As Date)
        Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander2.Connection = oConnection
        oCommander2.CommandType = CommandType.Text
        oCommander2.CommandText = "SELECT azn05 FROM azn_file where azn01 = TO_DATE('" & eDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        Dim ADW2 As Integer = oCommander2.ExecuteScalar()
        Return ADW2
    End Function
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "H1")
        oRng.EntireColumn.ColumnWidth = 20
        Ws.Cells(1, 1) = "customer_nr"
        Ws.Cells(1, 2) = "Text"
        Ws.Cells(1, 3) = "BS"
        Ws.Cells(1, 4) = "Invoice_nr"
        Ws.Cells(1, 5) = "invoice_date"
        Ws.Cells(1, 6) = "due_date"
        Ws.Cells(1, 7) = "amount"
        Ws.Cells(1, 8) = "due_balance"
        oCommand.CommandText = "SELECT distinct azn02,azn05 FROM AZN_FILE WHERE AZN01 >= TO_DATE('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') order by azn02,azn05"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            CC = 0
            While oReader.Read()
                Ws.Cells(1, 9 + CC) = oReader.Item("azn02") & "W" & oReader.Item("azn05")
                CC += 1
            End While
        End If
        oRng = Ws.Range("G1", "DB1")
        oRng.EntireColumn.NumberFormatLocal = "_-[$€-2] * #,##0.00_-;-[$€-2] * #,##0.00_-;_-[$€-2] * ""-""??_-;_-@_-"
        oReader.Close()
        LineZ = 2
    End Sub
    Private Function GetExchangeRate(ByVal eDate As Date)
        Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander2.Connection = oConnection
        oCommander2.CommandType = CommandType.Text
        Dim MM As String = String.Empty
        MM = eDate.Month
        If Strings.Len(MM) = 1 Then
            MM = "0" & MM
        End If
        MM = eDate.Year & MM
        oCommander2.CommandText = "select azj04 from azj_file where azj01 = 'EUR' AND azj02 = '" & MM & "'"
        Dim EUR As Decimal = oCommander2.ExecuteScalar()
        If IsDBNull(EUR) Or EUR = 0 Then
            Dim SX As String = String.Empty
            SX = DStartN.Month()
            If Strings.Len(SX) = 1 Then
                SX = "0" & SX
            End If
            SX = DStartN.Year & SX
            oCommander2.CommandText = "select azj04 from azj_file where azj01 = 'EUR' AND azj02 = '" & SX & "'"
            EUR = oCommander2.ExecuteScalar()
            If IsDBNull(EUR) Or EUR = 0 Then
                EUR = 1
            End If
        End If
        oCommander2.CommandText = "select azj04 from azj_file where azj01 = 'USD' AND azj02 = '" & MM & "'"
        Dim USD As Decimal = oCommander2.ExecuteScalar()
        If IsDBNull(USD) Or USD = 0 Then
            Dim SX As String = String.Empty
            SX = DStartN.Month()
            If Strings.Len(SX) = 1 Then
                SX = "0" & SX
            End If
            SX = DStartN.Year & SX
            oCommander2.CommandText = "select azj04 from azj_file where azj01 = 'USD' AND azj02 = '" & SX & "'"
            USD = oCommander2.ExecuteScalar()
            If IsDBNull(USD) Or USD = 0 Then
                USD = 1
            End If
        End If
        Dim RateR As Decimal = EUR / USD
        Return RateR
    End Function
    Private Sub Setwk(ByVal eDate As Date)
        Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander2.Connection = oConnection
        oCommander2.CommandType = CommandType.Text
        oCommander2.CommandText = "SELECT azn02 FROM azn_file where azn01 = TO_DATE('" & eDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        DW3 = 0
        DW3 = oCommander2.ExecuteScalar()
        oCommander2.CommandText = "SELECT azn05 FROM azn_file where azn01 = TO_DATE('" & eDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        DW4 = 0
        DW4 = oCommander2.ExecuteScalar()
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "P1")
        oRng.EntireColumn.ColumnWidth = 22.22
        Ws.Cells(1, 1) = "customer_nr"
        Ws.Cells(1, 2) = "Text"
        Ws.Cells(1, 3) = "Invoice_nr"
        Ws.Cells(1, 4) = "invoice_date"
        Ws.Cells(1, 5) = "due_date"
        Ws.Cells(1, 6) = "due_balance"
        oRng = Ws.Range("F1", "F1")
        oRng.EntireColumn.NumberFormatLocal = "_-[$€-2] * #,##0.00_-;-[$€-2] * #,##0.00_-;_-[$€-2] * ""-""??_-;_-@_-"
        Ws.Cells(1, 7) = "note"
        Ws.Cells(1, 8) = "customer"
        Ws.Cells(1, 9) = "payment terms_days"
        Ws.Cells(1, 10) = "how many overdue days"
        Ws.Cells(1, 11) = "not overdue"
        Ws.Cells(1, 12) = "overdue less then 30 days"
        Ws.Cells(1, 13) = "overdue over 30 days"
        oRng = Ws.Range("M1", "M1")
        oRng.Interior.Color = Color.LightGreen
        Ws.Cells(1, 14) = "overdue over 60 days"
        oRng = Ws.Range("N1", "N1")
        oRng.Interior.Color = Color.Yellow
        Ws.Cells(1, 15) = "overdue over 90 days"
        oRng = Ws.Range("O1", "O1")
        oRng.Interior.Color = Color.Orange
        Ws.Cells(1, 16) = "overdue over 120 days"
        oRng = Ws.Range("P1", "P1")
        oRng.Interior.Color = Color.Red
        oRng = Ws.Range("K1", "P1")
        oRng.EntireColumn.NumberFormatLocal = "_-[$€-2] * #,##0.00_-;-[$€-2] * #,##0.00_-;_-[$€-2] * ""-""??_-;_-@_-"
        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 37.33
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 19.11
        oRng.EntireColumn.NumberFormatLocal = "_-""US$""* #,##0.00_ ;_-""US$""* -#,##0.00 ;_-""US$""* ""-""??_ ;_-@_ "
        Ws.Cells(1, 1) = "overdue days"
        Ws.Cells(1, 2) = "sum amount"
        Ws.Cells(2, 1) = "overdue over 120 days"
        Ws.Cells(3, 1) = "overdue over 90 days"
        Ws.Cells(4, 1) = "overdue over 60 days"
        Ws.Cells(5, 1) = "overdue over 30 days"
        Ws.Cells(6, 1) = "overdue less than 30 days"
        Ws.Cells(7, 1) = "not overdue"
        Ws.Cells(8, 1) = "Total overdue AR_USD"
        oRng = Ws.Range("A8", "B8")
        oRng.Interior.Color = Color.Yellow
        Ws.Cells(10, 1) = "customer"
        Ws.Cells(10, 2) = "overdue amount"
        LineX = 11
    End Sub
    Private Sub AdjustExcelFormat3()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 37.33
        oRng = Ws.Range("B1", "G1")
        oRng.EntireColumn.ColumnWidth = 19.11
        oRng.EntireColumn.NumberFormatLocal = "_-""US$""* #,##0.00_ ;_-""US$""* -#,##0.00 ;_-""US$""* ""-""??_ ;_-@_ "
        'Ws.Cells(1, 1) = "overdue days"
        'Ws.Cells(1, 2) = "sum amount"
        Ws.Cells(1, 7) = "overdue over 120 days"
        Ws.Cells(1, 6) = "overdue over 90 days"
        Ws.Cells(1, 5) = "overdue over 60 days"
        Ws.Cells(1, 4) = "overdue over 30 days"
        Ws.Cells(1, 3) = "overdue less than 30 days"
        Ws.Cells(1, 2) = "not overdue"
        'Ws.Cells(7, 1) = "Total overdue AR_USD"
        'oRng = Ws.Range("A7", "B7")
        'oRng.Interior.Color = Color.Yellow
        'Ws.Cells(9, 1) = "customer"
        'Ws.Cells(9, 2) = "overdue amount"
        LineX = 2
    End Sub
End Class