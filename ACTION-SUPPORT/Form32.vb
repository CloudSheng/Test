Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form32
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim DStartN As Date
    Dim DstartE As Date
    Dim TYear As String = String.Empty
    Dim TMonth As String = String.Empty
    Dim LineZ As Integer = 0
    Dim CC As Integer = 0
    Dim DW1 As Integer = 0
    Dim DW2 As Integer = 0
    Dim DW3 As Integer = 0
    Dim DW4 As Integer = 0
    Dim YP As Decimal = 0
    Dim WY As String = String.Empty
    Dim WM As String = String.Empty
    Dim PaNext As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form32_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
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
        SaveFileDialog1.FileName = "DAC_AP_aging_Report"
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
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "weekly sum"
        'Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 20
        Ws.Cells(1, 1) = "类型category"
        Ws.Cells(1, 2) = "供应厂商编号supplier Nr."
        Ws.Cells(1, 3) = "全称supplier"
        Ws.Cells(1, 4) = "付款方式payment term"
        Ws.Cells(1, 5) = "付款方式说明description of payment term"
        Ws.Cells(1, 6) = "币种currency"
        Ws.Cells(1, 7) = "月底重评价汇率exchange rate"
        Ws.Cells(1, 8) = "应付帐款(本币)AP(RMB)"
        Ws.Cells(1, 9) = "小计金额(原币)subtotal(original currency)"
        Ws.Cells(1, 10) = "暂估金额(原币)temporary estimation(original currency)"
        oCommand.CommandText = "SELECT distinct azn02,azn05 FROM AZN_FILE WHERE AZN01 >= TO_DATE('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') order by azn02,azn05"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            CC = 0
            While oReader.Read()
                Ws.Cells(1, 11 + CC) = oReader.Item("azn02") & "W" & oReader.Item("azn05")
                CC += 1
            End While
        End If
        oReader.Close()
        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "monthly sum"
        'Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 20
        Ws.Cells(1, 1) = "类型category"
        Ws.Cells(1, 2) = "供应厂商编号supplier Nr."
        Ws.Cells(1, 3) = "全称supplier"
        Ws.Cells(1, 4) = "付款方式payment term"
        Ws.Cells(1, 5) = "付款方式说明description of payment term"
        Ws.Cells(1, 6) = "币种currency"
        Ws.Cells(1, 7) = "月底重评价汇率exchange rate"
        Ws.Cells(1, 8) = "应付帐款(本币)AP(RMB)"
        Ws.Cells(1, 9) = "小计金额(原币)subtotal(original currency)"
        Ws.Cells(1, 10) = "暂估金额(原币)temporary estimation(original currency)"
        oCommand.CommandText = "SELECT distinct azn02,azn04 FROM AZN_FILE WHERE AZN01 >= TO_DATE('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') order by azn02,azn04"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            CC = 0
            While oReader.Read()
                Dim MX As String = oReader.Item("azn04")
                If Strings.Len(MX) = 1 Then
                    MX = "0" & MX
                End If
                Ws.Cells(1, 11 + CC) = oReader.Item("azn02") & MX & "USD"
                Ws.Cells(1, 12 + CC) = oReader.Item("azn02") & MX & "RMB"
                CC += 2
            End While
        End If
        oReader.Close()
        PaNext = LineZ
        LineZ = 2
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        'Ws.Name = "Week"
        AdjustExcelFormat()
        '先訂位
        'oCommand.CommandText = "SELECT azn02 FROM azn_file where azn01 = TO_DATE('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        DW1 = GetAzn02(DStartN)
        'oCommand.CommandText = "SELECT azn05 FROM azn_file where azn01 = TO_DATE('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        DW2 = GetAzn05(DStartN)

        oCommand.CommandText = "select apa06,apa07,apa11,pma02,apa13,apa72,apa00,(case when apa00 in (21,22,23,24,25) then apa34 *(-1) else apa34 end ) as apa34,"
        oCommand.CommandText += "(case when apa00 in (21,22,23,24,25) then apa34f *(-1) else apa34f end ) as apa34f,apa12,(case when apa00 in (21,22,23,24,25) then apc13 *(-1) else apc13 end ) as t1 from apa_file,apc_file,pma_file "
        oCommand.CommandText += "where apa11 = pma01 and apa01 = apc01 and apc13 > 0 and apa41 = 'Y' "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("apa06")
                Ws.Cells(LineZ, 3) = oReader.Item("apa07")
                Ws.Cells(LineZ, 4) = oReader.Item("apa11")
                Ws.Cells(LineZ, 5) = oReader.Item("pma02")
                Ws.Cells(LineZ, 6) = oReader.Item("apa13")
                Ws.Cells(LineZ, 7) = oReader.Item("apa72")
                Ws.Cells(LineZ, 8) = oReader.Item("apa34")
                If oReader.Item("apa00").ToString = "16" Then
                    Ws.Cells(LineZ, 10) = oReader.Item("apa34f")
                Else
                    Ws.Cells(LineZ, 9) = oReader.Item("apa34f")
                End If
                ' 處理匯率
                If oReader.Item("apa13") = "USD" Then
                    YP = oReader.Item("apa34f")
                Else
                    WY = Convert.ToDateTime(oReader.Item("apa12")).Year
                    WM = Convert.ToDateTime(oReader.Item("apa12")).Month
                    If Strings.Len(WM) = 1 Then
                        WM = "0" & WM
                    End If
                    WY = WY & WM
                    Dim WL As Decimal = GetAzj04(WY)
                    YP = oReader.Item("t1") / WL
                End If
                If oReader.Item("apa12") <= DStartN Then
                    Ws.Cells(LineZ, 11) = YP
                Else
                    DW3 = GetAzn02(oReader.Item("apa12"))
                    DW4 = GetAzn05(oReader.Item("apa12"))
                    If DW3 = DW1 Then  '同年度, 處理週
                        Ws.Cells(LineZ, 11 + (DW4 - DW2)) = YP
                    Else  '跨年度
                        Dim MAXWK As Integer = GetMaxAzn05(DW1)
                        Ws.Cells(LineZ, 11 + (MAXWK - DW2 + DW4)) = YP
                    End If
                End If
                LineZ += 1
                Label2.Text = LineZ
            End While
        End If
        oReader.Close()
        ' 換行
        Ws.Cells(LineZ, 1) = "已入库未请款"
        LineZ += 1
        ' 已入庫未請款
        oCommand.CommandText = "select rvu04,pmc03,pmc17,pma02,pmm22,'','',(rvv39 * pmm42) as t1,(rvv39) as t2,rvu03,'',pma08 from rvv_file,rvu_file,pmc_file,pma_file,pmm_file where rvv01 = rvu01 and rvv23 < rvv17 and rvuacti = 'Y' and rvu04 = pmc01 and pmc17 = pma01 and rvv36 = pmm01 and rvv39 > 0"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("rvu04")
                Ws.Cells(LineZ, 3) = oReader.Item("pmc03")
                Ws.Cells(LineZ, 4) = oReader.Item("pmc17")
                Ws.Cells(LineZ, 5) = oReader.Item("pma02")
                Ws.Cells(LineZ, 6) = oReader.Item("pmm22")
                'Ws.Cells(LineZ, 7) = oReader.Item("apa72")
                Ws.Cells(LineZ, 8) = oReader.Item("t1")
                Ws.Cells(LineZ, 9) = oReader.Item("t2")
                Dim PD As Date = oReader.Item("rvu03")
                PD = Convert.ToDateTime(Year(PD) & "/" & Month(PD) & "/01")
                PD = PD.AddMonths(1).AddDays(-1)
                PD = PD.AddDays(oReader.Item("pma08"))
                ' 處理匯率
                If oReader.Item("pmm22") = "USD" Then
                    YP = oReader.Item("t2")
                Else
                    WY = PD.Year
                    WM = PD.Month
                    If Strings.Len(WM) = 1 Then
                        WM = "0" & WM
                    End If
                    WY = WY & WM
                    Dim WL As Decimal = GetAzj04(WY)
                    YP = oReader.Item("t1") / WL
                End If
                If PD <= DStartN Then
                    Ws.Cells(LineZ, 11) = YP
                Else
                    DW3 = GetAzn02(PD)
                    DW4 = GetAzn05(PD)
                    If DW3 = DW1 Then  '同年度, 處理週
                        Ws.Cells(LineZ, 11 + (DW4 - DW2)) = YP
                    Else  '跨年度
                        Dim MAXWK As Integer = GetMaxAzn05(DW1)
                        Ws.Cells(LineZ, 11 + (MAXWK - DW2 + DW4)) = YP
                    End If
                End If
                LineZ += 1
                Label2.Text = LineZ
            End While
        End If
        oReader.Close()
        ' 換行
        Ws.Cells(LineZ, 1) = "已采购未入库"
        LineZ += 1
        ' 已採購未入庫
        oCommand.CommandText = "select pmm09,pmc03,pmm20,pma02,pmm22,'','',(pmn88*pmm42) as t1,pmn88 as t2,pmn35,'',pma08 from pmm_file,pmn_file,pmc_file,pma_file where pmm01 = pmn01 and pmm18 = 'Y' and pmm25 in (1,2) "
        oCommand.CommandText += "and pmn16 in (1,2) and pmn53 < pmn20 and pmm09 = pmc01 and pmm20 = pma01 and pmn88 > 0"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("pmm09")
                Ws.Cells(LineZ, 3) = oReader.Item("pmc03")
                Ws.Cells(LineZ, 4) = oReader.Item("pmm20")
                Ws.Cells(LineZ, 5) = oReader.Item("pma02")
                Ws.Cells(LineZ, 6) = oReader.Item("pmm22")
                'Ws.Cells(LineZ, 7) = oReader.Item("apa72")
                Ws.Cells(LineZ, 8) = oReader.Item("t1")
                Ws.Cells(LineZ, 9) = oReader.Item("t2")
                Dim PD As Date = oReader.Item("pmn35")
                PD = Convert.ToDateTime(Year(PD) & "/" & Month(PD) & "/01")
                PD = PD.AddMonths(1).AddDays(-1)
                PD = PD.AddDays(oReader.Item("pma08"))
                ' 處理匯率
                If oReader.Item("pmm22") = "USD" Then
                    YP = oReader.Item("t2")
                Else
                    WY = PD.Year
                    WM = PD.Month
                    If Strings.Len(WM) = 1 Then
                        WM = "0" & WM
                    End If
                    WY = WY & WM
                    Dim WL As Decimal = GetAzj04(WY)
                    YP = oReader.Item("t1") / WL
                End If
                If PD <= DStartN Then
                    Ws.Cells(LineZ, 11) = YP
                Else
                    DW3 = GetAzn02(PD)
                    DW4 = GetAzn05(PD)
                    If DW3 = DW1 Then  '同年度, 處理週
                        Ws.Cells(LineZ, 11 + (DW4 - DW2)) = YP
                    Else  '跨年度
                        Dim MAXWK As Integer = GetMaxAzn05(DW1)
                        Ws.Cells(LineZ, 11 + (MAXWK - DW2 + DW4)) = YP
                    End If
                End If
                LineZ += 1
                Label2.Text = LineZ
            End While
        End If
        oReader.Close()
        ' 加總
        Ws.Cells(LineZ, 1) = "合计"
        Ws.Cells(LineZ, 11) = "=SUM(K2:K" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 11), Ws.Cells(LineZ, 11))
        oRng.AutoFill(Destination:=Ws.Range("K" & LineZ & ":DB" & LineZ), Type:=xlFillDefault)


        ' 第二頁
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustExcelFormat1()
        'Ws.Name = "Month"
        '先訂位
        DW1 = DStartN.Year
        DW2 = DStartN.Month
        oCommand.CommandText = "select apa06,apa07,apa11,pma02,apa13,apa72,apa00,(case when apa00 in (21,22,23,24,25) then apa34 *(-1) else apa34 end ) as apa34,"
        oCommand.CommandText += "(case when apa00 in (21,22,23,24,25) then apa34f *(-1) else apa34f end ) as apa34f,apa12,(case when apa00 in (21,22,23,24,25) then apc13 *(-1) else apc13 end ) as t1 from apa_file,apc_file,pma_file "
        oCommand.CommandText += "where apa11 = pma01 and apa01 = apc01 and apc13 > 0 and apa41 = 'Y' "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("apa06")
                Ws.Cells(LineZ, 3) = oReader.Item("apa07")
                Ws.Cells(LineZ, 4) = oReader.Item("apa11")
                Ws.Cells(LineZ, 5) = oReader.Item("pma02")
                Ws.Cells(LineZ, 6) = oReader.Item("apa13")
                Ws.Cells(LineZ, 7) = oReader.Item("apa72")
                Ws.Cells(LineZ, 8) = oReader.Item("apa34")
                If oReader.Item("apa00").ToString = "16" Then
                    Ws.Cells(LineZ, 10) = oReader.Item("apa34f")
                Else
                    Ws.Cells(LineZ, 9) = oReader.Item("apa34f")
                End If
                ' 處理匯率
                If oReader.Item("apa13") = "USD" Then
                    YP = oReader.Item("apa34f")
                Else
                    WY = Convert.ToDateTime(oReader.Item("apa12")).Year
                    WM = Convert.ToDateTime(oReader.Item("apa12")).Month
                    If Strings.Len(WM) = 1 Then
                        WM = "0" & WM
                    End If
                    WY = WY & WM
                    Dim WL As Decimal = GetAzj04(WY)
                    YP = oReader.Item("t1") / WL
                End If
                If oReader.Item("apa12") <= DStartN Then
                    Ws.Cells(LineZ, 11) = YP
                    Ws.Cells(LineZ, 12) = oReader.Item("t1")
                Else
                    DW3 = Convert.ToDateTime(oReader.Item("apa12")).Year
                    DW4 = Convert.ToDateTime(oReader.Item("apa12")).Month
                    If DW3 = DW1 Then  '同年度, 處理月
                        Ws.Cells(LineZ, 11 + (DW4 - DW2) * 2) = YP
                        Ws.Cells(LineZ, 12 + (DW4 - DW2) * 2) = oReader.Item("t1")
                    Else  '跨年度
                        Dim MAXWK As Integer = 12
                        Ws.Cells(LineZ, 11 + (MAXWK - DW2 + DW4) * 2) = YP
                        Ws.Cells(LineZ, 12 + (MAXWK - DW2 + DW4) * 2) = oReader.Item("t1")
                    End If
                End If
                LineZ += 1
                Label2.Text = LineZ + PaNext
            End While
        End If
        oReader.Close()
        ' 換行
        Ws.Cells(LineZ, 1) = "已入库未请款"
        LineZ += 1
        ' 已入庫未請款
        oCommand.CommandText = "select rvu04,pmc03,pmc17,pma02,pmm22,'','',(rvv39 * pmm42) as t1,(rvv39) as t2,rvu03,'',pma08 from rvv_file,rvu_file,pmc_file,pma_file,pmm_file where rvv01 = rvu01 and rvv23 < rvv17 and rvuacti = 'Y' and rvu04 = pmc01 and pmc17 = pma01 and rvv36 = pmm01 and rvv39 > 0"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("rvu04")
                Ws.Cells(LineZ, 3) = oReader.Item("pmc03")
                Ws.Cells(LineZ, 4) = oReader.Item("pmc17")
                Ws.Cells(LineZ, 5) = oReader.Item("pma02")
                Ws.Cells(LineZ, 6) = oReader.Item("pmm22")
                'Ws.Cells(LineZ, 7) = oReader.Item("apa72")
                Ws.Cells(LineZ, 8) = oReader.Item("t1")
                Ws.Cells(LineZ, 9) = oReader.Item("t2")
                Dim PD As Date = oReader.Item("rvu03")
                PD = Convert.ToDateTime(Year(PD) & "/" & Month(PD) & "/01")
                PD = PD.AddMonths(1).AddDays(-1)
                PD = PD.AddDays(oReader.Item("pma08"))
                ' 處理匯率
                If oReader.Item("pmm22") = "USD" Then
                    YP = oReader.Item("t2")
                Else
                    WY = PD.Year
                    WM = PD.Month
                    If Strings.Len(WM) = 1 Then
                        WM = "0" & WM
                    End If
                    WY = WY & WM
                    Dim WL As Decimal = GetAzj04(WY)
                    YP = oReader.Item("t1") / WL
                End If
                If PD <= DStartN Then
                    Ws.Cells(LineZ, 11) = YP
                    Ws.Cells(LineZ, 12) = oReader.Item("t1")
                Else
                    DW3 = PD.Year
                    DW4 = PD.Month
                    If DW3 = DW1 Then  '同年度, 處理月
                        Ws.Cells(LineZ, 11 + (DW4 - DW2) * 2) = YP
                        Ws.Cells(LineZ, 12 + (DW4 - DW2) * 2) = oReader.Item("t1")
                    Else  '跨年度
                        Dim MAXWK As Integer = 12
                        Ws.Cells(LineZ, 11 + (MAXWK - DW2 + DW4) * 2) = YP
                        Ws.Cells(LineZ, 12 + (MAXWK - DW2 + DW4) * 2) = oReader.Item("t1")
                    End If
                End If
                LineZ += 1
                Label2.Text = LineZ + PaNext
            End While
        End If
        oReader.Close()
        ' 換行
        Ws.Cells(LineZ, 1) = "已采购未入库"
        LineZ += 1
        ' 已採購未入庫
        oCommand.CommandText = "select pmm09,pmc03,pmm20,pma02,pmm22,'','',(pmn88*pmm42) as t1,pmn88 as t2,pmn35,'',pma08 from pmm_file,pmn_file,pmc_file,pma_file where pmm01 = pmn01 and pmm18 = 'Y' and pmm25 in (1,2) "
        oCommand.CommandText += "and pmn16 in (1,2) and pmn53 < pmn20 and pmm09 = pmc01 and pmm20 = pma01 and pmn88 > 0"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("pmm09")
                Ws.Cells(LineZ, 3) = oReader.Item("pmc03")
                Ws.Cells(LineZ, 4) = oReader.Item("pmm20")
                Ws.Cells(LineZ, 5) = oReader.Item("pma02")
                Ws.Cells(LineZ, 6) = oReader.Item("pmm22")
                'Ws.Cells(LineZ, 7) = oReader.Item("apa72")
                Ws.Cells(LineZ, 8) = oReader.Item("t1")
                Ws.Cells(LineZ, 9) = oReader.Item("t2")
                Dim PD As Date = oReader.Item("pmn35")
                PD = Convert.ToDateTime(Year(PD) & "/" & Month(PD) & "/01")
                PD = PD.AddMonths(1).AddDays(-1)
                PD = PD.AddDays(oReader.Item("pma08"))
                ' 處理匯率
                If oReader.Item("pmm22") = "USD" Then
                    YP = oReader.Item("t2")
                Else
                    WY = PD.Year
                    WM = PD.Month
                    If Strings.Len(WM) = 1 Then
                        WM = "0" & WM
                    End If
                    WY = WY & WM
                    Dim WL As Decimal = GetAzj04(WY)
                    YP = oReader.Item("t1") / WL
                End If
                If PD <= DStartN Then
                    Ws.Cells(LineZ, 11) = YP
                    Ws.Cells(LineZ, 12) = oReader.Item("t1")
                Else
                    DW3 = PD.Year
                    DW4 = PD.Month
                    If DW3 = DW1 Then  '同年度, 處理週
                        Ws.Cells(LineZ, 11 + (DW4 - DW2) * 2) = YP
                        Ws.Cells(LineZ, 12 + (DW4 - DW2) * 2) = oReader.Item("t1")
                    Else  '跨年度
                        Dim MAXWK As Integer = 12
                        Ws.Cells(LineZ, 11 + (MAXWK - DW2 + DW4) * 2) = YP
                        Ws.Cells(LineZ, 12 + (MAXWK - DW2 + DW4) * 2) = oReader.Item("t1")
                    End If
                End If
                LineZ += 1
                Label2.Text = LineZ + PaNext
            End While
        End If
        oReader.Close()
        ' 加總
        Ws.Cells(LineZ, 1) = "合计"
        Ws.Cells(LineZ, 11) = "=SUM(K2:K" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 11), Ws.Cells(LineZ, 11))
        oRng.AutoFill(Destination:=Ws.Range("K" & LineZ & ":DB" & LineZ), Type:=xlFillDefault)

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
    Private Function GetMaxAzn05(ByVal azn02 As Integer)
        Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander2.Connection = oConnection
        oCommander2.CommandType = CommandType.Text
        oCommander2.CommandText = "select max(azn05) from azn_file where azn02 = " & azn02
        Dim MK As Integer = oCommander2.ExecuteScalar()
        Return MK
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
End Class