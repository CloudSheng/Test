Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form37
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim Vi As Int16 = 0
    Dim Vj As Integer = 0
    Dim LineZ As Integer = 0
    Dim LineX As Integer = 0
    Dim DS As Data.DataSet = New DataSet()
    Dim DStartN As Date
    Dim CC As Integer = 0
    Dim PaNext As Integer = 0
    Dim DW1 As Integer = 0
    Dim DW2 As Integer = 0
    Dim DW3 As Integer = 0
    Dim DW4 As Integer = 0
    Dim SC As Decimal = 0
    Dim DRate As Decimal = 0
    Dim DDate As Date
    Dim SC2 As Decimal = 0
    Dim YP As Decimal = 0
    Dim WY As String = String.Empty
    Dim WM As String = String.Empty
    Dim PaymentWord As String = String.Empty
    Dim PayPlusDay As Integer = 0
    Dim PaymentDay As Date
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form37_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.GroupBox2.Enabled = False
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES;IMEX=1'"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString = "SELECT project_no,payment_term,Zaefrina_only_Estimate_return_date,Quotation_Red_confirmedWithFinance,supplier FROM [2016$] WHERE purchase_Y_or_Blank IS NULL"
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)
            Try
                ExcelAdapater.Fill(DS, "tables1")
            Catch ex As Exception
                MsgBox(ex.Message())
                Label1.Text = "读取失败"
                Me.GroupBox2.Enabled = False
                Return
            End Try
            Label1.Text = "已读入"
            Me.GroupBox2.Enabled = True
        End If
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
        DStartN = Today()
        Label1.Text = "执行中"
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Label1.Text = "已完成"
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Mold_AP_aging_Report"
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
        '先訂位
        DW1 = GetAzn02(DStartN)
        DW2 = GetAzn05(DStartN)
        For i As Integer = 0 To DS.Tables("tables1").Rows.Count - 1 Step 1
            PaymentWord = String.Empty
            If IsDBNull(DS.Tables("tables1").Rows(i).Item("project_no")) Then
                'LineZ += 1
                'Label2.Text = LineZ
                Continue For
            End If
            Ws.Cells(LineZ, 1) = LineZ - 1
            Ws.Cells(LineZ, 2) = DS.Tables("tables1").Rows(i).Item("project_no")
            If Not IsDBNull(DS.Tables("tables1").Rows(i).Item("payment_term")) Then
                Dim AAA As String = String.Empty
                AAA = DS.Tables("tables1").Rows(i).Item("payment_term")
                If Strings.Len(AAA) = 1 Then
                    AAA = "0" & AAA
                End If
                PaymentWord = Getpma02(AAA)
                PayPlusDay = Getpma08(AAA)
                Ws.Cells(LineZ, 3) = PaymentWord
            End If
            Ws.Cells(LineZ, 5) = DS.Tables("tables1").Rows(i).Item("Supplier")
            Ws.Cells(LineZ, 6) = DS.Tables("tables1").Rows(i).Item("Quotation_Red_confirmedWithFinance")
            Ws.Cells(LineZ, 7) = DS.Tables("tables1").Rows(i).Item("Zaefrina_only_Estimate_return_date")
            If IsDBNull(DS.Tables("tables1").Rows(i).Item("payment_term")) Or IsDBNull(DS.Tables("tables1").Rows(i).Item("Zaefrina_only_Estimate_return_date")) Or String.IsNullOrEmpty(PaymentWord) Then
                LineZ += 1
                Label3.Text = LineZ
                Continue For
            End If
            ' 底下的表示資料集全, 可以開始算週數

            PaymentDay = DS.Tables("tables1").Rows(i).Item("Zaefrina_only_Estimate_return_date")
            PaymentDay = Convert.ToDateTime(PaymentDay.Year & "/" & PaymentDay.Month & "/01")
            PaymentDay = PaymentDay.AddMonths(1).AddDays(-1)
            PaymentDay = PaymentDay.AddDays(PayPlusDay)
            ' 處理匯率
            WY = PaymentDay.Year
            WM = PaymentDay.Month
            If Strings.Len(WM) = 1 Then
                WM = "0" & WM
            End If
            WY = WY & WM
            Dim WL As Decimal = GetAzj04(WY)
            YP = DS.Tables("tables1").Rows(i).Item("Quotation_Red_confirmedWithFinance") / WL
            ' 放入指定欄位
            If PaymentDay <= DStartN Then
                Ws.Cells(LineZ, 8) = YP
            Else
                DW3 = GetAzn02(PaymentDay)
                DW4 = GetAzn05(PaymentDay)
                If DW3 = DW1 Then  '同年度, 處理週
                    Ws.Cells(LineZ, 8 + (DW4 - DW2)) = YP
                Else  '跨年度
                    Dim MAXWK As Integer = GetMaxAzn05(DW1)
                    Ws.Cells(LineZ, 8 + (MAXWK - DW2 + DW4)) = YP
                End If
            End If
            LineZ += 1
            Label3.Text = LineZ
        Next
        ' 加總
        Ws.Cells(LineZ, 1) = "合计"
        Ws.Cells(LineZ, 8) = "=SUM(H2:H" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 8), Ws.Cells(LineZ, 8))
        oRng.AutoFill(Destination:=Ws.Range("H" & LineZ & ":DB" & LineZ), Type:=xlFillDefault)

        ' 第二頁
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustExcelFormat1()
        '先訂位
        DW1 = DStartN.Year
        DW2 = DStartN.Month
        For i As Integer = 0 To DS.Tables("tables1").Rows.Count - 1 Step 1
            PaymentWord = String.Empty
            If IsDBNull(DS.Tables("tables1").Rows(i).Item("project_no")) Then
                'LineZ += 1
                'Label2.Text = LineZ
                Continue For
            End If
            Ws.Cells(LineZ, 1) = LineZ - 1
            Ws.Cells(LineZ, 2) = DS.Tables("tables1").Rows(i).Item("project_no")
            If Not IsDBNull(DS.Tables("tables1").Rows(i).Item("payment_term")) Then
                Dim AAA As String = String.Empty
                AAA = DS.Tables("tables1").Rows(i).Item("payment_term")
                If Strings.Len(AAA) = 1 Then
                    AAA = "0" & AAA
                End If
                PaymentWord = Getpma02(AAA)
                PayPlusDay = Getpma08(AAA)
                'PaymentWord = Getpma02(DS.Tables("tables1").Rows(i).Item("payment_term"))
                Ws.Cells(LineZ, 3) = PaymentWord
            End If
            Ws.Cells(LineZ, 5) = DS.Tables("tables1").Rows(i).Item("Supplier")
            Ws.Cells(LineZ, 6) = DS.Tables("tables1").Rows(i).Item("Quotation_Red_confirmedWithFinance")
            Ws.Cells(LineZ, 7) = DS.Tables("tables1").Rows(i).Item("Zaefrina_only_Estimate_return_date")
            If IsDBNull(DS.Tables("tables1").Rows(i).Item("payment_term")) Or IsDBNull(DS.Tables("tables1").Rows(i).Item("Zaefrina_only_Estimate_return_date")) Or String.IsNullOrEmpty(PaymentWord) Then
                LineZ += 1
                Label3.Text = LineZ + PaNext
                Continue For
            End If
            ' 底下的表示資料集全, 可以開始算月數
            'PayPlusDay = Getpma08(DS.Tables("tables1").Rows(i).Item("payment_term"))
            PaymentDay = DS.Tables("tables1").Rows(i).Item("Zaefrina_only_Estimate_return_date")
            PaymentDay = Convert.ToDateTime(PaymentDay.Year & "/" & PaymentDay.Month & "/01")
            PaymentDay = PaymentDay.AddMonths(1).AddDays(-1)
            PaymentDay = PaymentDay.AddDays(PayPlusDay)
            ' 處理匯率
            WY = PaymentDay.Year
            WM = PaymentDay.Month
            If Strings.Len(WM) = 1 Then
                WM = "0" & WM
            End If
            WY = WY & WM
            Dim WL As Decimal = GetAzj04(WY)
            YP = DS.Tables("tables1").Rows(i).Item("Quotation_Red_confirmedWithFinance") / WL
            ' 放入指定欄位
            If PaymentDay <= DStartN Then
                Ws.Cells(LineZ, 8) = YP
            Else
                DW3 = PaymentDay.Year
                DW4 = PaymentDay.Month
                If DW3 = DW1 Then  '同年度, 處理月
                    Ws.Cells(LineZ, 8 + (DW4 - DW2)) = YP
                Else  '跨年度
                    Dim MAXWK As Integer = GetMaxAzn05(DW1)
                    Ws.Cells(LineZ, 8 + (MAXWK - DW2 + DW4)) = YP
                End If
            End If
            LineZ += 1
            Label3.Text = LineZ + PaNext
        Next
        ' 加總
        Ws.Cells(LineZ, 1) = "合计"
        Ws.Cells(LineZ, 8) = "=SUM(H2:H" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 8), Ws.Cells(LineZ, 8))
        oRng.AutoFill(Destination:=Ws.Range("H" & LineZ & ":DB" & LineZ), Type:=xlFillDefault)
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "weekly sum"
        'Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 20
        Ws.Cells(1, 1) = "項次"
        Ws.Cells(1, 2) = "型号"
        Ws.Cells(1, 3) = "付款条件"
        Ws.Cells(1, 4) = "是否采购"
        Ws.Cells(1, 5) = "供货商"
        Ws.Cells(1, 6) = "报价"
        Ws.Cells(1, 7) = "预计回厂时间"
        oCommand.CommandText = "SELECT distinct azn02,azn05 FROM AZN_FILE WHERE AZN01 >= TO_DATE('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') order by azn02,azn05"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            CC = 0
            While oReader.Read()
                Ws.Cells(1, 8 + CC) = oReader.Item("azn02") & "W" & oReader.Item("azn05")
                CC += 1
            End While
        End If
        oReader.Close()
        LineZ = 2
    End Sub
    Private Function Getpma02(ByVal pma01 As String)
        Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander2.Connection = oConnection
        oCommander2.CommandType = CommandType.Text
        oCommander2.CommandText = "SELECT pma02 FROM pma_file where pma01 = '" & pma01 & "'"
        Dim ADW1 As String = oCommander2.ExecuteScalar()
        Return ADW1
    End Function
    Private Function Getpma08(ByVal pma01 As String)
        Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander2.Connection = oConnection
        oCommander2.CommandType = CommandType.Text
        oCommander2.CommandText = "SELECT pma08 FROM pma_file where pma01 = '" & pma01 & "'"
        Dim ADW1 As Integer = oCommander2.ExecuteScalar()
        Return ADW1
    End Function
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
    Private Function GetAzj04(ByVal MM As String)
        Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander2.Connection = oConnection
        oCommander2.CommandType = CommandType.Text
        oCommander2.CommandText = "select azj04 from azj_file where azj01 = 'USD' AND azj02 = '" & MM & "'"
        Dim MK As Integer = oCommander2.ExecuteScalar()
        If IsDBNull(MK) Or MK = 0 Then
            Dim SX As String = String.Empty
            SX = DStartN.Month()
            If Strings.Len(SX) = 1 Then
                SX = "0" & SX
            End If
            SX = DStartN.Year & SX
            oCommander2.CommandText = "select azj04 from azj_file where azj01 = 'USD' AND azj02 = '" & SX & "'"
            MK = oCommander2.ExecuteScalar()
            If IsDBNull(MK) Or MK = 0 Then
                MK = 1
            End If
        End If
        Return MK
    End Function
    Private Function GetMaxAzn05(ByVal azn02 As Integer)
        Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander2.Connection = oConnection
        oCommander2.CommandType = CommandType.Text
        oCommander2.CommandText = "select max(azn05) from azn_file where azn02 = " & azn02
        Dim MK As Integer = oCommander2.ExecuteScalar()
        Return MK
    End Function
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "monthly sum"
        'Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 20
        Ws.Cells(1, 1) = "項次"
        Ws.Cells(1, 2) = "型号"
        Ws.Cells(1, 3) = "付款条件"
        Ws.Cells(1, 4) = "是否采购"
        Ws.Cells(1, 5) = "供货商"
        Ws.Cells(1, 6) = "报价"
        Ws.Cells(1, 7) = "预计回厂时间"
        oCommand.CommandText = "SELECT distinct azn02,azn04 FROM AZN_FILE WHERE AZN01 >= TO_DATE('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') order by azn02,azn04"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            CC = 0
            While oReader.Read()
                Ws.Cells(1, 8 + CC) = oReader.Item("azn02") & "M" & oReader.Item("azn04")
                CC += 1
            End While
        End If
        oReader.Close()
        PaNext = LineZ
        LineZ = 2
    End Sub
End Class