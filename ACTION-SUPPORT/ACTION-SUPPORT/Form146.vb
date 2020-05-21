Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form146
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim tWeek As Int16 = 0
    Dim nYear As Int16 = 0
    Dim nWeek As Int16 = 0
    Dim tDate As Date
    Dim AzjString As String = String.Empty
    Dim LineZ As Integer = 0

    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form146_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        
        tDate = Me.DateTimePicker1.Value
        tYear = tDate.Year
        tMonth = tDate.Month
        If tMonth < 10 Then
            AzjString = tYear & "0" & tMonth
        Else
            AzjString = tYear & tMonth
        End If
        oCommand.CommandText = "SELECT azn05 From azn_file where azn01 = to_date('" & tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        tWeek = oCommand.ExecuteScalar()

        nYear = tYear
        nWeek = tWeek

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
        SaveFileDialog1.FileName = "應收帳款预测表"
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
        xWorkBook.Sheets.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Name = "cxmt808"
        Ws.Activate()
        AdjustExcelFormat()
        oCommand.CommandText = "select tc_cif_01,tc_cif_02,oeb04, gea02, occ02,tc_cif_03,tc_cif_04,tc_cif_05,azn05,azn02, tc_prl06,"
        oCommand.CommandText += "(case when tc_prl06 = 'USD' then 1 else azj041 end) as t1, tc_prl03 from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
        oCommand.CommandText += "left join oea_file on oeb01 = oea01 left join occ_file on oea04 = occ01 left join gea_file on occ20 = gea01 left join azn_file on tc_cif_05 = azn01 "
        oCommand.CommandText += "left join tc_prl_file on tc_prl01 = oeb04 and tc_prl02 > tc_cif_05 left join hkacttest.azj_file on azj01 = 'EUR' and azj02 = '"
        oCommand.CommandText += AzjString & "' where tc_cif_05 >= to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and oeb70 <> 'Y' and tc_cif_01 not like 'FC%' and tc_prl02 = (select min(tc_prl02) from tc_prl_file where tc_prl01 = oeb04 and tc_prl02 > tc_cif_05) "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 2) = oReader.Item(i)
                Next
                ' 右方處理
                Ws.Cells(LineZ, 15) = "=H" & LineZ & "*N" & LineZ
                Ws.Cells(LineZ, 16) = "=N" & LineZ & "*M" & LineZ
                Ws.Cells(LineZ, 17) = "=H" & LineZ & "*P" & LineZ
                LineZ += 1
            End While

            ' 加入 格式
            oRng = Ws.Range("H4", Ws.Cells(LineZ - 1, 8))
            oRng.EntireColumn.NumberFormat = "#,##0_ ;[Red]-#,##0 "
            oRng = Ws.Range("N4", Ws.Cells(LineZ - 1, 17))
            oRng.EntireColumn.NumberFormat = "#,##0_ ;[Red]-#,##0 "
            oRng = Ws.Range("M4", Ws.Cells(LineZ - 1, 13))
            oRng.EntireColumn.NumberFormat = "#,##0.00_ ;[Red]-#,##0 "
            ' 劃線
            oRng = Ws.Range("B3", Ws.Cells(LineZ - 1, 17))
            oRng.Borders(xlEdgeLeft).LineStyle = xlNone
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlDouble
            oRng.Borders(xlEdgeRight).LineStyle = xlNone
            oRng.Borders(xlInsideHorizontal).LineStyle = xlNone
            oRng.Borders(xlInsideVertical).LineStyle = xlNone
        End If
        oReader.Close()
        ' C 到 最後一行 作自動判斷
        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 17))
        oRng.EntireColumn.AutoFit()
        ' 加入 邏輯說明
        LineZ += 1
        Ws.Cells(LineZ, 2) = "报表逻辑备注："
        Ws.Cells(LineZ + 1, 2) = "1.交货数量、交货日期、周别取之于cxmt808中大于等于报表查询日期的未结案销售订单资料"
        Ws.Cells(LineZ + 2, 2) = "2.产品售价取之于cxmt809中截止日期大于等于I栏交货日期对应的售价。如果符合条件的截止日期对应的售价有几个，取最小截止日期对应的售价"
        Ws.Cells(LineZ + 3, 2) = "3.如果币别为USD汇率为1，如果币别为EUR汇率以HK端欧元兑美元的汇率为准"

        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ + 3, 2))
        oRng.HorizontalAlignment = xlLeft
        ' 凍結
        oRng = Ws.Range("F4", "F4")
        oRng.Select()
        xExcel.ActiveWindow.FreezePanes = True

        ' 第二頁
        Ws = xWorkBook.Sheets(2)
        Ws.Name = "cxmt811"
        Ws.Activate()
        AdjustExcelFormat1()
        oCommand.CommandText = "select tc_cif_01,tc_cif_02,ta_opd14, gea02, occ02,tc_cif_03,tc_cif_04,tc_cif_05,azn05,azn02, tc_prl06,"
        oCommand.CommandText += "(case when tc_prl06 = 'USD' then 1 else azj041 end) as t1, tc_prl03 from tc_cif_file left join opd_file on tc_cif_01 = opd01 and tc_cif_02 = opd05 "
        oCommand.CommandText += "left join opc_file on opd01 = opc01 left join occ_file on opc02 = occ01 left join gea_file on occ20 = gea01 left join azn_file on tc_cif_05 = azn01 "
        oCommand.CommandText += "left join tc_prl_file on tc_prl01 = ta_opd14 and tc_prl02 > tc_cif_05 left join hkacttest.azj_file on azj01 = 'EUR' and azj02 = '"
        oCommand.CommandText += AzjString & "' where tc_cif_05 >= to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_01 like 'FC%' and tc_prl02 = (select min(tc_prl02) from tc_prl_file where tc_prl01 = ta_opd14 and tc_prl02 > tc_cif_05) "
        oCommand.CommandText += "and tc_cif_05 > (select nvl(max(tc_cif_05),to_date('2010/01/01','yyyy/mm/dd')) as t1 from tc_cif_file C left join oeb_file D on tc_cif_01 = oeb01 "
        oCommand.CommandText += "and tc_cif_02 = oeb03 where tc_cif_01 not like 'FC%' and oeb70 <> 'Y'  and d.oeb04 = ta_opd14)"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 2) = oReader.Item(i)
                Next
                ' 右方處理
                Ws.Cells(LineZ, 15) = "=G" & LineZ & "*M" & LineZ
                Ws.Cells(LineZ, 16) = "=N" & LineZ & "*M" & LineZ
                Ws.Cells(LineZ, 17) = "=G" & LineZ & "*O" & LineZ
                LineZ += 1
            End While

            ' 加入 格式
            oRng = Ws.Range("H4", Ws.Cells(LineZ - 1, 8))
            oRng.EntireColumn.NumberFormat = "#,##0_ ;[Red]-#,##0 "
            oRng = Ws.Range("N4", Ws.Cells(LineZ - 1, 17))
            oRng.EntireColumn.NumberFormat = "#,##0_ ;[Red]-#,##0 "
            oRng = Ws.Range("M4", Ws.Cells(LineZ - 1, 13))
            oRng.EntireColumn.NumberFormat = "#,##0.00_ ;[Red]-#,##0 "
            ' 劃線
            oRng = Ws.Range("B3", Ws.Cells(LineZ - 1, 17))
            oRng.Borders(xlEdgeLeft).LineStyle = xlNone
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlDouble
            oRng.Borders(xlEdgeRight).LineStyle = xlNone
            oRng.Borders(xlInsideHorizontal).LineStyle = xlNone
            oRng.Borders(xlInsideVertical).LineStyle = xlNone
        End If
        oReader.Close()
        ' C 到 最後一行 作自動判斷
        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 17))
        oRng.EntireColumn.AutoFit()
        ' 加入 邏輯說明
        LineZ += 1
        Ws.Cells(LineZ, 2) = "报表逻辑备注："
        Ws.Cells(LineZ + 1, 2) = "1.交货数量、交货日期、周别取之于cxmt811中大于等于报表查询日期的未结案预测单单号资料"
        Ws.Cells(LineZ + 2, 2) = "2.产品售价取之于cxmt809中截止日期大于等于I栏交货日期对应的售价。如果符合条件的截止日期对应的售价有几个，取最小截止日期对应的售价"
        Ws.Cells(LineZ + 3, 2) = "3.如果币别为USD汇率为1，如果币别为EUR汇率以HK端欧元兑美元的汇率为准"

        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ + 3, 2))
        oRng.HorizontalAlignment = xlLeft
        ' 凍結
        oRng = Ws.Range("F4", "F4")
        oRng.Select()
        xExcel.ActiveWindow.FreezePanes = True

        ' 第三頁
        Ws = xWorkBook.Sheets(3)
        Ws.Name = "USA Japan"
        Ws.Activate()
        AdjustExcelFormat2()

        GetData("EUROPE")
        GetData("USA")
        GetData("JAPAN")

        ' C 到 最後一行 作自動判斷
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 74))
        oRng.EntireColumn.AutoFit()

        ' 加入 格式
        oRng = Ws.Range("C6", "C9")
        oRng.EntireRow.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        oRng = Ws.Range("C13", "C13")
        oRng.EntireRow.NumberFormat = "#,##0_ ;[Red]-#,##0 "

        ' 凍結
        oRng = Ws.Range("C5", "C5")
        oRng.Select()
        xExcel.ActiveWindow.FreezePanes = True

    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Rows.RowHeight = 18
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        
        oRng = Ws.Range("B2", "B2")
        oRng.HorizontalAlignment = xlLeft

        oRng = Ws.Range("A2", "A3")
        oRng.EntireRow.Font.Bold = True
        oRng.Font.Size = 16
        oRng.Font.Bold = True

        Ws.Cells(2, 2) = "Company Name：Dongguan Action Composite LTD. Co"
        Ws.Cells(3, 2) = "订单单号"
        Ws.Cells(3, 3) = "项次"
        Ws.Cells(3, 4) = "产品编号"
        oRng = Ws.Range("D5", "D5")
        oRng.EntireColumn.NumberFormat = "@"
        Ws.Cells(3, 5) = "按产品分区域"
        Ws.Cells(3, 6) = "按订单送货客户"
        Ws.Cells(3, 7) = "项次"
        Ws.Cells(3, 8) = "数量"
        Ws.Cells(3, 9) = "交货日期"
        Ws.Cells(3, 10) = "周别"
        Ws.Cells(3, 11) = "年度"
        Ws.Cells(3, 12) = "原币"
        Ws.Cells(3, 13) = "转USD汇率"
        Ws.Cells(3, 14) = "cxmt809产品售价（原币）"
        Ws.Cells(3, 15) = "销售金额（原币）"
        Ws.Cells(3, 16) = "cxmt809产品售价（USD）"
        Ws.Cells(3, 17) = "销售金额（USD）"

        LineZ = 4
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Rows.RowHeight = 18
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter

        oRng = Ws.Range("B2", "B2")
        oRng.HorizontalAlignment = xlLeft

        oRng = Ws.Range("A2", "A3")
        oRng.EntireRow.Font.Bold = True
        oRng.Font.Size = 16
        oRng.Font.Bold = True

        Ws.Cells(2, 2) = "Company Name：Dongguan Action Composite LTD. Co"
        Ws.Cells(3, 2) = "预测单单号"
        Ws.Cells(3, 3) = "项次"
        Ws.Cells(3, 4) = "产品编号"
        oRng = Ws.Range("D5", "D5")
        oRng.EntireColumn.NumberFormat = "@"
        Ws.Cells(3, 5) = "按产品分区域"
        Ws.Cells(3, 6) = "按订单送货客户"
        Ws.Cells(3, 7) = "项次"
        Ws.Cells(3, 8) = "数量"
        Ws.Cells(3, 9) = "交货日期"
        Ws.Cells(3, 10) = "周别"
        Ws.Cells(3, 11) = "年度"
        Ws.Cells(3, 12) = "原币"
        Ws.Cells(3, 13) = "转USD汇率"
        Ws.Cells(3, 14) = "cxmt809产品售价（原币）"
        Ws.Cells(3, 15) = "销售金额（原币）"
        Ws.Cells(3, 16) = "cxmt809产品售价（USD）"
        Ws.Cells(3, 17) = "销售金额（USD）"

        LineZ = 4
    End Sub
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Rows.RowHeight = 18
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter

        oRng = Ws.Range("B2", "B3")
        oRng.HorizontalAlignment = xlLeft

        oRng = Ws.Range("A2", "A5")
        oRng.EntireRow.Font.Bold = True
        oRng.Font.Size = 16
        oRng.Font.Bold = True

        oRng = Ws.Range("A11", "A12")
        oRng.EntireRow.Font.Bold = True
        oRng.Font.Size = 16
        oRng.Font.Bold = True

        Ws.Cells(2, 2) = "Company Name：Dongguan Action Composite LTD. Co"
        Ws.Cells(3, 2) = "Currency:USD"
        Ws.Cells(4, 2) = "year"
        Ws.Cells(5, 2) = "Week"
        Ws.Cells(6, 2) = "Europe"
        Ws.Cells(7, 2) = "USA"
        Ws.Cells(8, 2) = "Japan"
        Ws.Cells(9, 2) = "Total"

        Ws.Cells(11, 2) = "year"
        Ws.Cells(12, 2) = "Week"
        Ws.Cells(13, 2) = "USA+Japan"

        For i As Int16 = 0 To 71 Step 1
            nWeek += 1
            oCommand.CommandText = "select count(*) from azn_file where azn02 = " & nYear & " and azn05 = " & nWeek
            Dim Hasweek As Int16 = oCommand.ExecuteScalar()
            If Hasweek <= 0 Then

                nYear += 1
                oCommand.CommandText = "select nvl(sum(t1),0) from ( select azn02,max(azn05) as t1 from azn_file where azn02 between "
                oCommand.CommandText += tYear & " and " & nYear - 1 & " group by azn02 ) "
                Dim TotalWeek As Int16 = oCommand.ExecuteScalar()
                nWeek = nWeek - TotalWeek
                If nWeek <= 0 Then
                    nWeek = 1
                End If
            End If
            Ws.Cells(4, 3 + i) = nYear & "年"
            Ws.Cells(5, 3 + i) = nWeek
            Ws.Cells(11, 3 + i) = nYear & "年"
            Ws.Cells(12, 3 + i) = nWeek
        Next

        Ws.Cells(9, 3) = "=SUM(C6:C8)"
        oRng = Ws.Range("C9", "C9")
        oRng.AutoFill(Destination:=Ws.Range("C9", Ws.Cells(9, 74)), Type:=xlFillDefault)

        Ws.Cells(13, 3) = "=C7+C8"
        oRng = Ws.Range("C13", "C13")
        oRng.AutoFill(Destination:=Ws.Range("C13", Ws.Cells(13, 74)), Type:=xlFillDefault)

        Ws.Cells(15, 2) = "报表逻辑："
        Ws.Cells(16, 2) = "1.按年度和周别汇总cxmt808和cxmt811的USD计划和预测出货金额"

        oRng = Ws.Range("B15", "B16")
        oRng.HorizontalAlignment = xlLeft
        'oRng = Ws.Range("B2", Ws.Cells(2, TotalWeek + 9))
        'oRng.Merge()
        'oRng.Font.Size = 16
        'Ws.Cells(2, 2) = "Call off shipping amount by week"

        LineZ = 6
    End Sub
    Private Sub GetData(ByVal gea02 As String)
        nYear = tYear
        nWeek = tWeek
        oCommand.CommandText = "select "
        For i As Int16 = 1 To 72 Step 1
            oCommand.CommandText += "sum(c" & i & ") as c" & i & ","
        Next
        oCommand.CommandText += "1 from ( select "
        For i As Int16 = 0 To 71 Step 1
            nWeek += 1
            oCommand2.CommandText = "select count(*) from azn_file where azn02 = " & nYear & " and azn05 = " & nWeek
            Dim Hasweek As Int16 = oCommand2.ExecuteScalar()
            If Hasweek <= 0 Then

                nYear += 1
                oCommand2.CommandText = "select nvl(sum(t1),0) from ( select azn02,max(azn05) as t1 from azn_file where azn02 between "
                oCommand2.CommandText += tYear & " and " & nYear - 1 & " group by azn02 ) "
                Dim TotalWeek As Int16 = oCommand2.ExecuteScalar()
                nWeek = nWeek - TotalWeek
                If nWeek <= 0 Then
                    nWeek = 1
                End If
            End If
            oCommand.CommandText += "(case when azn02 = " & nYear & " and azn05 = " & nWeek & " then (tc_cif_04 * t1 * tc_prl03) else 0 end) as c" & i + 1 & ","
        Next
        oCommand.CommandText += "1 from ( select gea02,tc_cif_04,azn05,azn02,(case when tc_prl06 = 'USD' then 1 else azj041 end) as t1, tc_prl03 from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
        oCommand.CommandText += "left join oea_file on oeb01 = oea01 left join occ_file on oea04 = occ01 left join gea_file on occ20 = gea01 left join azn_file on tc_cif_05 = azn01 "
        oCommand.CommandText += "left join tc_prl_file on tc_prl01 = oeb04 and tc_prl02 > tc_cif_05 left join hkacttest.azj_file on azj01 = 'EUR' and azj02 = '"
        oCommand.CommandText += AzjString & "' where tc_cif_05 >= to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and oeb70 <> 'Y' and tc_cif_01 not like 'FC%' and tc_prl02 = (select min(tc_prl02) from tc_prl_file where tc_prl01 = oeb04 and tc_prl02 > tc_cif_05) and gea02 = '"
        oCommand.CommandText += gea02 & "' ) "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select "
        nYear = tYear
        nWeek = tWeek
        For i As Int16 = 0 To 71 Step 1
            nWeek += 1
            oCommand2.CommandText = "select count(*) from azn_file where azn02 = " & nYear & " and azn05 = " & nWeek
            Dim Hasweek As Int16 = oCommand2.ExecuteScalar()
            If Hasweek <= 0 Then

                nYear += 1
                oCommand2.CommandText = "select nvl(sum(t1),0) from ( select azn02,max(azn05) as t1 from azn_file where azn02 between "
                oCommand2.CommandText += tYear & " and " & nYear - 1 & " group by azn02 ) "
                Dim TotalWeek As Int16 = oCommand2.ExecuteScalar()
                nWeek = nWeek - TotalWeek
                If nWeek <= 0 Then
                    nWeek = 1
                End If
            End If
            oCommand.CommandText += "(case when azn02 = " & nYear & " and azn05 = " & nWeek & " then (tc_cif_04 * t1 * tc_prl03) else 0 end) as c" & i & ","
        Next
        oCommand.CommandText += "1 from ( select gea02,tc_cif_04,azn05,azn02,(case when tc_prl06 = 'USD' then 1 else azj041 end) as t1, tc_prl03 from tc_cif_file left join opd_file on tc_cif_01 = opd01 and tc_cif_02 = opd05 "
        oCommand.CommandText += "left join opc_file on opd01 = opc01 left join occ_file on opc02 = occ01 left join gea_file on occ20 = gea01 left join azn_file on tc_cif_05 = azn01  left join tc_prl_file on tc_prl01 = ta_opd14 and tc_prl02 > tc_cif_05 left join hkacttest.azj_file on azj01 = 'EUR' and azj02 = '"
        oCommand.CommandText += AzjString & "' where tc_cif_05 >= to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_01 like 'FC%' and tc_prl02 = (select min(tc_prl02) from tc_prl_file where tc_prl01 = ta_opd14 and tc_prl02 > tc_cif_05) "
        oCommand.CommandText += "and tc_cif_05 > (select nvl(max(tc_cif_05),to_date('2010/01/01','yyyy/mm/dd')) as t1 from tc_cif_file C left join oeb_file D on tc_cif_01 = oeb01 "
        oCommand.CommandText += "and tc_cif_02 = oeb03 where tc_cif_01 not like 'FC%' and oeb70 <> 'Y'  and d.oeb04 = ta_opd14)  and gea02 = '"
        oCommand.CommandText += gea02 & "' ) )"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 2 Step 1
                    Ws.Cells(LineZ, 3 + i) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()
        LineZ += 1
    End Sub
End Class