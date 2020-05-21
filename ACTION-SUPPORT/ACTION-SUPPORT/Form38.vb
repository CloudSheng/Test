Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form38
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim TYear As String = String.Empty
    Dim TMonth As String = String.Empty
    Dim PYear As String = String.Empty
    Dim PMonth As String = String.Empty
    Dim DStartN As Date
    Dim DstartE As Date
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form38_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If Now.Month < 10 Then
            TextBox1.Text = Now.Year & "0" & Now.Month
        Else
            TextBox1.Text = Now.Year & Now.Month
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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
        TYear = Strings.Left(TextBox1.Text, 4)
        TMonth = Strings.Right(TextBox1.Text, 2)
        If TMonth > 12 Or TMonth < 1 Then
            MsgBox("Error Month Data")
            Return
        End If
        If TMonth = 1 Then
            PYear = TYear - 1
            PMonth = 12
        Else
            PMonth = TMonth - 1
        End If
        DStartN = Convert.ToDateTime(TYear & "/" & TMonth & "/01")
        DstartE = DStartN.AddMonths(1).AddDays(-1)
        BackgroundWorker1.RunWorkerAsync()

    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Month_Purchase_Report"
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
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "Monthly Purchase Report"
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 25
        Ws.Cells(1, 1) = "日期 Date"
        Ws.Cells(1, 2) = "申请日期"
        Ws.Cells(1, 3) = "申购的需求日期"
        Ws.Cells(1, 4) = "收料日期"
        Ws.Cells(1, 5) = "料号 P/N"
        oRng = Ws.Range("E2", "E2")
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(1, 6) = "品名 Product Name"
        Ws.Cells(1, 7) = "规格 Specfication"
        Ws.Cells(1, 8) = "供应商 Vendor"
        Ws.Cells(1, 9) = "首次采购金额 First time purchase amount"
        Ws.Cells(1, 10) = "该次采购金额 Price"
        Ws.Cells(1, 11) = "差异 Diff(%)"
        oRng = Ws.Range("K2", "K2")
        oRng.EntireColumn.NumberFormatLocal = "0.00%"
        Ws.Cells(1, 12) = "币别 Currency"
        Ws.Cells(1, 13) = "该次采购量 Pruchase amount"
        Ws.Cells(1, 14) = "单位 Unit"
        Ws.Cells(1, 15) = "金额（amount）"
        Ws.Cells(1, 16) = "该次采购总金额 The total amount of purchase"
        Ws.Cells(1, 17) = "月采购金额 Monthly purchase amount"
        Ws.Cells(1, 18) = "前置期(天) L/T (days)"
        Ws.Cells(1, 19) = "实际周期(天) Actual delivery date"
        Ws.Cells(1, 20) = "帐期(天) Payment term"
        Ws.Cells(1, 21) = "採購員 Purchaser"
        Ws.Cells(1, 22) = "申请需求日与收料日差异天数"
        LineZ = 2
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        AdjustExcelFormat1()
        oCommand.CommandText = "select pmm04,pmk04,pml34,pmn04,pmn041,ima021,pmc03,pmn31t,pmm22,pmn20,pmn07,pmn88t,ima48,pma02,gen02,pmn01,pmn02,pmn34 from pmm_file,pmn_file,ima_file,pmc_file,pma_file,gen_file,pmk_file,pml_file where pmm01 = pmn01 and pmm04 between to_date('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmn24 = pml01 and pmn25 = pml02 and pmk01 = pml01 and pmm18 = 'Y' and pmn04 in (select distinct bmb03 from bmb_file) and pmn04 = ima01 and pmm09 = pmc01 and pmm20 = pma01 and pmm12 = gen01 and pmm02 <> 'SUB' "
        oCommand.CommandText += "order by pmn04"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("pmm04")
                Ws.Cells(LineZ, 2) = oReader.Item("pmk04")
                Dim DemandDate As Date = oReader.Item("pml34")
                Ws.Cells(LineZ, 3) = DemandDate
                Dim ReceiveDate As String = GetReturnDate1(oReader.Item("pmn01"), oReader.Item("pmn02"))
                Ws.Cells(LineZ, 4) = ReceiveDate
                Ws.Cells(LineZ, 5) = oReader.Item("pmn04")
                Ws.Cells(LineZ, 6) = oReader.Item("pmn041")
                Ws.Cells(LineZ, 7) = oReader.Item("ima021")
                Ws.Cells(LineZ, 8) = oReader.Item("pmc03")
                Dim FP As Decimal = GetFirstPrice(oReader.Item("pmn04"))
                Ws.Cells(LineZ, 9) = FP
                Ws.Cells(LineZ, 10) = oReader.Item("pmn31t")
                If FP <> 0 Then
                    Ws.Cells(LineZ, 11) = (oReader.Item("pmn31t") - FP) / FP
                Else
                    Ws.Cells(LineZ, 11) = "N/A"
                End If
                Ws.Cells(LineZ, 12) = oReader.Item("pmm22")
                Ws.Cells(LineZ, 13) = oReader.Item("pmn20")
                Ws.Cells(LineZ, 14) = oReader.Item("pmn07")
                Dim AP As Decimal = GetAveragePrice(oReader.Item("pmn04"))
                Ws.Cells(LineZ, 15) = AP
                Ws.Cells(LineZ, 16) = oReader.Item("pmn88t")
                Ws.Cells(LineZ, 18) = oReader.Item("ima48")
                Ws.Cells(LineZ, 19) = GetReturnDate(oReader.Item("pmn01"), oReader.Item("pmn02"), oReader.Item("pmn34"))
                Ws.Cells(LineZ, 20) = oReader.Item("pma02")
                Ws.Cells(LineZ, 21) = oReader.Item("gen02")
                If ReceiveDate = "分批收货" Or ReceiveDate = "未收货" Then
                    Ws.Cells(LineZ, 22) = ReceiveDate
                Else
                    Ws.Cells(LineZ, 22) = DateDiff(DateInterval.Day, DemandDate, Convert.ToDateTime(ReceiveDate))
                End If
                LineZ += 1
            End While
        End If
        oReader.Close()
    End Sub
    Private Function GetFirstPrice(ByVal pmn04 As String)
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "select nvl(pmn31t,0) from pmm_file,pmn_file where pmm01 = pmn01 and pmn04 = '"
        oCommander99.CommandText += pmn04 & "' and pmm18 = 'Y' and rownum = 1 order by pmm04 "
        Dim FP As Decimal = oCommander99.ExecuteScalar()
        Return FP
    End Function
    Private Function GetAveragePrice(ByVal pmn04 As String)
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "select Round((SUM(pmn88t)/sum(pmn20)),4) as t1 from pmm_file,pmn_file where pmm01 = pmn01 and pmn04 = '"
        oCommander99.CommandText += pmn04 & "' and pmm18 = 'Y' and pmm04 between to_date('"
        oCommander99.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommander99.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        Dim AP As Decimal = oCommander99.ExecuteScalar()
        Return AP
    End Function
    Private Function GetReturnDate(ByVal pmn01 As String, ByVal pmn02 As Integer, ByVal pmn34 As Date)
        Dim CRA As String = String.Empty
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "select count(*) from rva_file,rvb_file where rva01 = rvb01 and rvaconf = 'Y' and rvb04 = '"
        oCommander99.CommandText += pmn01 & "' and rvb03 = " & pmn02
        Dim CR As Decimal = oCommander99.ExecuteScalar()
        If CR > 1 Then
            CRA = "分批收货"
        ElseIf CR = 0 Then
            CRA = "未收货"
        Else
            oCommander99.CommandText = "select rva06 from rva_file,rvb_file where rva01 = rvb01 and rvaconf = 'Y' and rvb04 = '"
            oCommander99.CommandText += pmn01 & "' and rvb03 = " & pmn02
            Dim CRB As Date = oCommander99.ExecuteScalar()
            CRA = DateDiff(DateInterval.Day, pmn34, CRB)
        End If
        Return CRA
    End Function
    Private Function GetReturnDate1(ByVal pmn01 As String, ByVal pmn02 As Integer)
        Dim CRA As String = String.Empty
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "select count(*) from rva_file,rvb_file where rva01 = rvb01 and rvaconf = 'Y' and rvb04 = '"
        oCommander99.CommandText += pmn01 & "' and rvb03 = " & pmn02
        Dim CR As Decimal = oCommander99.ExecuteScalar()
        If CR > 1 Then
            CRA = "分批收货"
        ElseIf CR = 0 Then
            CRA = "未收货"
        Else
            oCommander99.CommandText = "select rva06 from rva_file,rvb_file where rva01 = rvb01 and rvaconf = 'Y' and rvb04 = '"
            oCommander99.CommandText += pmn01 & "' and rvb03 = " & pmn02
            CRA = oCommander99.ExecuteScalar()
            'CRA = DateDiff(DateInterval.Day, CRB, pmn34)
        End If
        Return CRA
    End Function
End Class