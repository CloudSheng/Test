Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form144
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
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim tWeek As Int16 = 0
    Dim pYear As Int16 = 0
    Dim pMonth As Int16 = 0
    Dim tDate As Date
    Dim tDate1 As Date
    Dim tDate2 As Date '關帳日期後一日
    Dim tDate3 As Date
    Dim LineZ As Integer = 0
    Dim TotalWeek As Int16 = 0
    Dim MaxWeek As Int16 = 0
    Dim LMonth As Int16 = 0
    Dim aMonth As Int16 = 0
    Dim Start1 As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form144_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
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
        Start1 = TextBox1.Text
        If Len(Start1) <> 6 Then
            MsgBox("月份资料为6码")
            Return
        End If
        tDate = Me.DateTimePicker1.Value
        tYear = tDate.Year
        tMonth = tDate.Month
        pMonth = Convert.ToInt16(Strings.Right(TextBox1.Text, 2))
        pYear = Convert.ToInt16(Strings.Left(TextBox1.Text, 4))
        'If pMonth = 0 Then
        'pMonth = 12
        'pYear = tYear - 1
        'End If

        oCommand.CommandText = "SELECT azn05 From azn_file where azn01 = to_date('" & tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        tWeek = oCommand.ExecuteScalar()
        tDate1 = Convert.ToDateTime(tYear & "/12/31")
        oCommand.CommandText = "SELECT azn05 From azn_file where azn01 = to_date('" & tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        MaxWeek = oCommand.ExecuteScalar()
        ' 20181009
        oCommand.CommandText = "SELECT aaa07 from aaa_file"
        tDate2 = oCommand.ExecuteScalar()
        aMonth = tDate2.Month
        tDate2 = tDate2.AddDays(1)
        tDate3 = tDate.AddDays(-1)
        If tDate2.Month < tMonth Then
            LMonth = tDate2.Month
        Else
            LMonth = tMonth
        End If
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
        SaveFileDialog1.FileName = "Rolling_Forecast"
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
        Ws.Name = "call off shipping qty by w"
        Ws.Activate()
        AdjustExcelFormat()

        oCommand.CommandText = "select tqa02,oeb04,ima02,ima021,gea02,oeb05"
        For i As Int16 = tWeek To MaxWeek Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        oCommand.CommandText += " from ( "
        oCommand.CommandText += "select tqa02,oeb04,ima02,ima021,gea02,oeb05"
        For i As Int16 = tWeek To MaxWeek Step 1
            oCommand.CommandText += ",(case when azn05 = " & i & " then sum(tc_cif_04) else 0 end ) as t" & i
        Next
        oCommand.CommandText += " from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
        oCommand.CommandText += "left join ima_file on oeb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = 2 "
        oCommand.CommandText += "left join oea_file on oeb01 = oea01 left join occ_file on oea04 = occ01 left join gea_file on occ20 = gea01 "
        oCommand.CommandText += "left join azn_file on tc_cif_05 = azn01 where tc_cif_05 >= to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_05 <= to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and oeb70 <> 'Y' and tc_cif_01 not like 'FC%' group by tqa02,oeb04,ima02,ima021,gea02,oeb05,azn05 "
        oCommand.CommandText += ") group by tqa02,oeb04,ima02,ima021,gea02,oeb05"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    If i > 5 Then
                        Ws.Cells(LineZ, i + 3) = oReader.Item(i)
                    Else
                        Ws.Cells(LineZ, i + 2) = oReader.Item(i)
                    End If            
                Next
                ' 右方加總
                Ws.Cells(LineZ, 9 + TotalWeek).FormulaR1C1 = "=SUM(RC[-" & TotalWeek & "]:RC[-1])"
                LineZ += 1
            End While
            ' 下方加總
            Ws.Cells(LineZ, 8) = "Total"
            Ws.Cells(LineZ, 9) = "=SUM(I6:I" & LineZ - 1 & ")"
            ' 複制
            oRng = Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9))
            oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9 + TotalWeek)), Type:=xlFillDefault)

            ' 加入 格式
            oRng = Ws.Range("I6", Ws.Cells(LineZ, 9 + TotalWeek))
            oRng.EntireColumn.NumberFormat = "#,##0_ ;[Red]-#,##0 "

            ' 劃線
            oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 9 + TotalWeek))
            oRng.Borders(xlEdgeLeft).LineStyle = xlNone
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlDouble
            oRng.Borders(xlEdgeRight).LineStyle = xlNone
            oRng.Borders(xlInsideHorizontal).LineStyle = xlNone
            oRng.Borders(xlInsideVertical).LineStyle = xlNone
        End If
        oReader.Close()
        ' C 到 最後一行 作自動判斷
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 9))
        oRng.EntireColumn.AutoFit()
        ' 加入 邏輯說明
        LineZ += 2
        Ws.Cells(LineZ, 2) = "报表逻辑备注："
        Ws.Cells(LineZ + 1, 2) = "1.第4行显示的年度为报表当年"
        Ws.Cells(LineZ + 2, 2) = "2.第5行显示的周别为报表起始日期对应的周别至报表当年最后一个周别，报表起始日期之前的周别无需显示"
        Ws.Cells(LineZ + 3, 2) = "3.如果报表起始日期对应周别中有两部分日期组成：一部分日期小于报表起始日期，另外一部分日期大于等于报表起始日期。此份报表只需要抓取日期大于等于报表起始日期对应的资料"
        Ws.Cells(LineZ + 4, 2) = "4.如果报表当年最后一周对应的日期有跨年的情况，此份报表最后一个周别只需要抓取当年日期对应的资料。"
        Ws.Cells(LineZ + 5, 2) = "5.介于第3点和第4点之间的周别对应的日期，只需要按周别抓取资料即可，无需区别跨月的情况"
        Ws.Cells(LineZ + 6, 2) = "6.把cxmt808（订单项次多角期输入）中周别栏位（azn05）对应数量栏位（tc_cif_04）的数量按周别汇总。如果cxmt808（订单项次多角期输入）中订单单号栏位（oeb01）显示的订单已经无效了，则需要排除该订单相关资料"
        Ws.Cells(LineZ + 7, 2) = "7.如果报表第一周对应的交货日期栏位（tc_cif_05）既有部分交货日期小于报表起始日期又有部分交货日期大于等于报表起始日期，此时只需要汇总交货日期大于等于报表起始日期对应的的数量"
        Ws.Cells(LineZ + 8, 2) = "8.如果报表最后一周有跨年的情况，需要把报表次年的数量排除在外"
        Ws.Cells(LineZ + 9, 2) = "9.根据产品编号在axmi121（产品主档维护作业）中抓取品牌代码，再用品牌代码在atmi402（基础编号代码维护作业）抓取客户名称"

        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ + 9, 2))
        oRng.HorizontalAlignment = xlLeft
        ' 凍結
        oRng = Ws.Range("D6", "D6")
        oRng.Select()
        xExcel.ActiveWindow.FreezePanes = True

        Ws = xWorkBook.Sheets(2)
        Ws.Name = "call off shipping amount  by w"
        Ws.Activate()
        AdjustExcelFormat1()

        oCommand.CommandText = "select tqa02,oeb04,ima02,ima021,gea02,oeb05"
        For i As Int16 = tWeek To MaxWeek Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        oCommand.CommandText += " from ( "
        oCommand.CommandText += "select tqa02,oeb04,ima02,ima021,gea02,oeb05"
        For i As Int16 = tWeek To MaxWeek Step 1
            oCommand.CommandText += ",(case when azn05 = " & i & " then sum(tc_cif_04 * tc_prl03 * (case when tc_prl06 = 'USD' THEN 6.3 when tc_prl06 = 'EUR' then 7.56 else 0 end) ) else 0 end ) as t" & i
        Next
        oCommand.CommandText += " from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
        oCommand.CommandText += "left join ima_file on oeb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = 2 "
        oCommand.CommandText += "left join oea_file on oeb01 = oea01 left join occ_file on oea04 = occ01 left join gea_file on occ20 = gea01 "
        oCommand.CommandText += "left join azn_file on tc_cif_05 = azn01 left join tc_prl_file on tc_prl01 = oeb04 and tc_prl02 > tc_cif_05 where tc_cif_05 >= to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_05 <= to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and oeb70 <> 'Y' and tc_cif_01 not like 'FC%' and tc_prl02 = (select min(tc_prl02) from tc_prl_file where tc_prl01 = oeb04 and tc_prl02 > tc_cif_05) group by tqa02,oeb04,ima02,ima021,gea02,oeb05,azn05 "
        oCommand.CommandText += ") group by tqa02,oeb04,ima02,ima021,gea02,oeb05"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    If i > 5 Then
                        Ws.Cells(LineZ, i + 3) = oReader.Item(i)
                    Else
                        Ws.Cells(LineZ, i + 2) = oReader.Item(i)
                    End If
                Next
                ' 右方加總
                Ws.Cells(LineZ, 9 + TotalWeek).FormulaR1C1 = "=SUM(RC[-" & TotalWeek & "]:RC[-1])"
                LineZ += 1
            End While
            ' 下方加總
            Ws.Cells(LineZ, 8) = "Total"
            Ws.Cells(LineZ, 9) = "=SUM(I7:I" & LineZ - 1 & ")"
            ' 複制
            oRng = Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9))
            oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9 + TotalWeek)), Type:=xlFillDefault)

            ' 加入 格式
            oRng = Ws.Range("I7", Ws.Cells(LineZ, 9 + TotalWeek))
            oRng.EntireColumn.NumberFormat = "#,##0_ ;[Red]-#,##0 "

            ' 劃線
            oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 9 + TotalWeek))
            oRng.Borders(xlEdgeLeft).LineStyle = xlNone
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlDouble
            oRng.Borders(xlEdgeRight).LineStyle = xlNone
            oRng.Borders(xlInsideHorizontal).LineStyle = xlNone
            oRng.Borders(xlInsideVertical).LineStyle = xlNone
        End If
        oReader.Close()
        ' C 到 最後一行 作自動判斷
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 9))
        oRng.EntireColumn.AutoFit()
        ' 加入 邏輯說明
        LineZ += 2
        Ws.Cells(LineZ, 2) = "报表逻辑备注："
        Ws.Cells(LineZ + 1, 2) = "1.产品计划销售金额=产品数量（cxmt808中数量栏位（tc_cif_04）显示的数量）*产品售价（cxmt809中价格栏位（tc_prl03）的价格）"
        Ws.Cells(LineZ + 2, 2) = "2.第5行显示的年度为报表当年"
        Ws.Cells(LineZ + 3, 2) = "3.第6行显示的周别为报表起始日期对应的周别至报表当年最后一个周别，报表起始日期之前的周别无需显示"
        Ws.Cells(LineZ + 4, 2) = "4.如果报表起始日期对应周别中有两部分日期组成：一部分日期小于报表起始日期，另外一部分日期大于等于报表起始日期。此份报表只需要抓取日期大于等于报表起始日期对应的资料"
        Ws.Cells(LineZ + 5, 2) = "5.如果报表当年最后一周对应的日期有跨年的情况，此份报表最后一个周别只需要抓取当年日期对应的资料。"
        Ws.Cells(LineZ + 6, 2) = "6.介于第3点和第4点之间的周别对应的日期，只需要按周别抓取资料即可，无需区别跨月的情况"
        Ws.Cells(LineZ + 7, 2) = "7.根据产品编号及其在cxmt808中维护的交货日期（tc_cif_05），抓取该产品编号在cxmt809中最近一笔大于等于交货日期（tc_cif_05）的截止日期（tc_prl02）对应的价格（tc_prl03）。产品取价百分比为100%，cxmt808中交货日期（tc_cif_05）小于报表起始日期排除在外"
        Ws.Cells(LineZ + 8, 2) = "8.用截止日期（tc_prl02）对应的价格（tc_prl03）*交货日期（tc_cif_05）对应的数量（tc_cif_04）算出产品的计划销售金额，然后在按周别汇总每周的计划销售金额"
        Ws.Cells(LineZ + 9, 2) = "9.如果产品售价币别为USD,产品的销售金额按6.3汇率转RMB；如果产品售价币别为EUR,产品的销售金额按7.56汇率(即6.3*1.2)转RMB"
        Ws.Cells(LineZ + 10, 2) = "10.如果cxmt808（订单项次多角期输入）中订单单号（oeb01）栏位显示的订单已经无效了，则需要排除该订单相关资料"
        Ws.Cells(LineZ + 11, 2) = "11.如果报表第一周对应的交货日期栏位（tc_cif_05）既有部分交货日期小于报表起始日期又有部分交货日期大于等于报表起始日期，此时只需要汇总交货日期大于等于报表起始日期对应的的计划销售金额"
        Ws.Cells(LineZ + 12, 2) = "12.如果报表最后一周有跨年的情况，需要把报表次年的计划销售金额排除在外"
        Ws.Cells(LineZ + 13, 2) = "13.根据产品编号在axmi121（产品主档维护作业）中抓取品牌代码，再用品牌代码在atmi402（基础编号代码维护作业）抓取客户名称"

        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ + 13, 2))
        oRng.HorizontalAlignment = xlLeft
        ' 凍結
        oRng = Ws.Range("D6", "D6")
        oRng.Select()
        xExcel.ActiveWindow.FreezePanes = True

        '第3頁
        Ws = xWorkBook.Sheets(3)
        Ws.Name = "call off shipping by date"
        Ws.Activate()
        AdjustExcelFormat2()

        oCommand.CommandText = "select tqa02,oeb04,ima02,tc_cif_01,gea02,oeb05,tc_cif_05,tc_cif_04,azn05,tc_prl03,tc_prl06 from tc_cif_file "
        oCommand.CommandText += "left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 left join ima_file on oeb04 = ima01 "
        oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = 2 left join oea_file on oeb01 = oea01 "
        oCommand.CommandText += "left join occ_file on oea04 = occ01 left join gea_file on occ20 = gea01 left join azn_file on tc_cif_05 = azn01 "
        oCommand.CommandText += "left join tc_prl_file on tc_prl01 = oeb04 and tc_prl02 > tc_cif_05 where tc_cif_05 >= to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_05 <= to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_01 not like 'FC%' and oeb70 <> 'Y' and tc_prl02 = (select min(tc_prl02) from tc_prl_file where tc_prl01 = oeb04 and tc_prl02 > tc_cif_05) "

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 2) = oReader.Item(i)
                Next
                LineZ += 1
            End While
            ' 下方加總
            Ws.Cells(LineZ, 8) = "Total"
            Ws.Cells(LineZ, 9) = "=SUM(I6:I" & LineZ - 1 & ")"

            ' 加入 格式
            oRng = Ws.Range("I6", "I6")
            oRng.EntireColumn.NumberFormat = "#,##0_ ;[Red]-#,##0 "
            oRng = Ws.Range("K6", "K6")
            oRng.EntireColumn.NumberFormat = "#,##0_ ;[Red]-#,##0 "

            ' 劃線
            oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 12))
            oRng.Borders(xlEdgeLeft).LineStyle = xlNone
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlDouble
            oRng.Borders(xlEdgeRight).LineStyle = xlNone
            oRng.Borders(xlInsideHorizontal).LineStyle = xlNone
            oRng.Borders(xlInsideVertical).LineStyle = xlNone
        End If
        oReader.Close()
        ' C 到 最後一行 作自動判斷
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 9))
        oRng.EntireColumn.AutoFit()
        ' 加入 邏輯說明
        LineZ += 2
        Ws.Cells(LineZ, 2) = "报表逻辑备注："
        Ws.Cells(LineZ + 1, 2) = "1.把cxmt808（订单项次多角期输入）中交货日期（tc_cif_05）属于报表起始日期~当年12/31日期间的产品编号抓取出来"
        Ws.Cells(LineZ + 2, 2) = "2.订单号、交货日期、交货数量、周别抓取cxmt808（订单项次多角期输入）程式中订单号（oeb01）、交货日期（tc_cif_05）、交货数量（tc_cif_04）、周别栏位（azn05）的信息"
        Ws.Cells(LineZ + 3, 2) = "3.产品售价、币别抓取cxmt809（料件价格表）程式中价格（tc_prl03）、币别栏位（tc_prl06）的信息，产品取价百分比为100%"
        Ws.Cells(LineZ + 4, 2) = "4.同一个产品编号有不同订单需要分行显示"
        Ws.Cells(LineZ + 5, 2) = "5.同一个产品编号有不同交货日期也需要分行显示"
        Ws.Cells(LineZ + 6, 2) = "6.如果cxmt808（订单项次多角期输入）中订单单号（oeb01）栏位显示的订单已经无效了，则需要排除该订单相关资料"
        Ws.Cells(LineZ + 7, 2) = "7.产品售价取数原则：根据产品编号及其交货日期（tc_cif_05），抓取该产品编号在cxmt809中最近一笔大于等于交货日期（tc_cif_05）的截止日期（tc_prl02）对应的价格（tc_prl03）。cxmt808中交货日期（tc_cif_05）小于报表起始日期排除在外"
        Ws.Cells(LineZ + 8, 2) = "8.根据产品编号在axmi121（产品主档维护作业）中抓取品牌代码，再用品牌代码在atmi402（基础编号代码维护作业）抓取客户名称"

        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ + 8, 2))
        oRng.HorizontalAlignment = xlLeft
        ' 凍結
        oRng = Ws.Range("E6", "E6")
        oRng.Select()
        xExcel.ActiveWindow.FreezePanes = True

        ' 第4頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(4)
        Ws.Name = "forecast shipping qty by w "
        Ws.Activate()
        AdjustExcelFormat3()

        oCommand.CommandText = "select tqa02,ta_opd14,ima02,ima021,gea02,ima31"
        For i As Int16 = tWeek To MaxWeek Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        oCommand.CommandText += " from ( "
        oCommand.CommandText += "select tqa02,B.ta_opd14,ima02,ima021,gea02,ima31"
        For i As Int16 = tWeek To MaxWeek Step 1
            oCommand.CommandText += ",(case when azn05 = " & i & " then tc_cif_04 else 0 end ) as t" & i
        Next
        oCommand.CommandText += " from tc_cif_file A left join opd_file B on tc_cif_01 = opd01 and tc_cif_02 = opd05 "
        oCommand.CommandText += "left join ima_file on ta_opd14 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = 2 "
        oCommand.CommandText += "left join opc_file on opd01 = opc01 left join occ_file on opc02 = occ01 left join gea_file on occ20 = gea01 "
        oCommand.CommandText += "left join azn_file on tc_cif_05 = azn01 where tc_cif_05 >= to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_05 <= to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_01 like 'FC%' and tc_opc00 <> 'Y' and tc_cif_05 > (select nvl(max(tc_cif_05),to_date('2010/01/01','yyyy/mm/dd')) as t1 from tc_cif_file C left join oeb_file D on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
        oCommand.CommandText += "where tc_cif_01 not like 'FC%' and oeb70 <> 'Y'  and d.oeb04 = b.ta_opd14 )  "
        oCommand.CommandText += ") group by tqa02,ta_opd14,ima02,ima021,gea02,ima31"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    If i > 5 Then
                        Ws.Cells(LineZ, i + 3) = oReader.Item(i)
                    Else
                        Ws.Cells(LineZ, i + 2) = oReader.Item(i)
                    End If
                Next
                ' 右方加總
                Ws.Cells(LineZ, 9 + TotalWeek).FormulaR1C1 = "=SUM(RC[-" & TotalWeek & "]:RC[-1])"
                LineZ += 1
            End While
            ' 下方加總
            Ws.Cells(LineZ, 8) = "Total"
            Ws.Cells(LineZ, 9) = "=SUM(I6:I" & LineZ - 1 & ")"
            ' 複制
            oRng = Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9))
            oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9 + TotalWeek)), Type:=xlFillDefault)

            ' 加入 格式
            oRng = Ws.Range("I6", Ws.Cells(LineZ, 9 + TotalWeek))
            oRng.EntireColumn.NumberFormat = "#,##0_ ;[Red]-#,##0 "

            ' 劃線
            oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 9 + TotalWeek))
            oRng.Borders(xlEdgeLeft).LineStyle = xlNone
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlDouble
            oRng.Borders(xlEdgeRight).LineStyle = xlNone
            oRng.Borders(xlInsideHorizontal).LineStyle = xlNone
            oRng.Borders(xlInsideVertical).LineStyle = xlNone
        End If
        oReader.Close()
        ' C 到 最後一行 作自動判斷
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 9))
        oRng.EntireColumn.AutoFit()
        ' 加入 邏輯說明
        LineZ += 2
        Ws.Cells(LineZ, 2) = "报表逻辑备注："
        Ws.Cells(LineZ + 1, 2) = "1.第4行显示的年度为报表当年"
        Ws.Cells(LineZ + 2, 2) = "2.第5行显示的周别为报表起始日期对应的周别至报表当年最后一个周别，报表起始日期之前的周别无需显示"
        Ws.Cells(LineZ + 3, 2) = "3.如果报表起始日期对应周别中有两部分日期组成：一部分日期小于报表起始日期，另外一部分日期大于等于报表起始日期。此份报表只需要抓取日期大于等于报表起始日期对应的资料"
        Ws.Cells(LineZ + 4, 2) = "4.如果报表当年最后一周对应的日期有跨年的情况，此份报表最后一个周别只需要抓取当年日期对应的资料。"
        Ws.Cells(LineZ + 5, 2) = "5.介于第3点和第4点之间的周别对应的日期，只需要按周别抓取资料即可，无需区别跨月的情况"
        Ws.Cells(LineZ + 6, 2) = "6.如果同一款产品编号在cxmt808（订单项次多角期输入）和cxmt811（销售预测单项次多交期输入）中都有符合以上条件的资料，则此份报表需要以该产品编号在cxmt808（订单项次多角期输入）中最后一笔交货日期（（tc_cif_05））的次日开始统计预测数量。产品编号在cxmt811中交期小于等于cxmt808中交货日期的资料无需要统计"
        Ws.Cells(LineZ + 7, 2) = "7.如果cxmt811（销售预测单项次多交期输入）中订单单号栏位（oeb01）显示的订单已经无效了，则需要排除该订单相关资料"
        Ws.Cells(LineZ + 8, 2) = "8.把cxmt811（销售预测单项次多交期输入）中符合以上7点条件的产品编号按周别栏位（azn05）对应数量栏位（tc_cif_04）的汇总数量"
        Ws.Cells(LineZ + 9, 2) = "9.如果报表第一周对应的交期栏位（tc_cif_05）既有部分交货日期小于报表起始日期又有部分交期大于等于报表起始日期，此时只需要汇总交期大于等于报表起始日期对应的的数量"
        Ws.Cells(LineZ + 10, 2) = "10.如果报表最后一周有跨年的情况，需要把报表次年的数量排除在外"
        Ws.Cells(LineZ + 11, 2) = "11.同一款产品编号需要把cxmt811中交期小于等于该产品编号在cxmt808中交货日期的"
        Ws.Cells(LineZ + 12, 2) = "12.根据产品编号在axmi121（产品主档维护作业）中抓取品牌代码，再用品牌代码在atmi402（基础编号代码维护作业）抓取客户名称"

        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ + 12, 2))
        oRng.HorizontalAlignment = xlLeft
        ' 凍結
        oRng = Ws.Range("D6", "D6")
        oRng.Select()
        xExcel.ActiveWindow.FreezePanes = True

        ' 第五頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(5)
        Ws.Name = "forecast shipping amount by w"
        Ws.Activate()
        AdjustExcelFormat4()

        oCommand.CommandText = "select tqa02,ta_opd14,ima02,ima021,gea02,ima31"
        For i As Int16 = tWeek To MaxWeek Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        oCommand.CommandText += " from ( "
        oCommand.CommandText += "select tqa02,B.ta_opd14,ima02,ima021,gea02,ima31"
        For i As Int16 = tWeek To MaxWeek Step 1
            oCommand.CommandText += ",(case when azn05 = " & i & " then (tc_cif_04 * tc_prl03 * (case when tc_prl06 = 'USD' THEN 6.3 when tc_prl06 = 'EUR' then 7.56 else 0 end) ) else 0 end ) as t" & i
        Next
        oCommand.CommandText += " from tc_cif_file A left join opd_file B on tc_cif_01 = opd01 and tc_cif_02 = opd05 "
        oCommand.CommandText += "left join ima_file on ta_opd14 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = 2 "
        oCommand.CommandText += "left join opc_file on opd01 = opc01 left join occ_file on opc02 = occ01 left join gea_file on occ20 = gea01 "
        oCommand.CommandText += "left join azn_file on tc_cif_05 = azn01 left join tc_prl_file on tc_prl01 = ta_opd14 and tc_prl02 > tc_cif_05 where tc_cif_05 >= to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_05 <= to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_01 like 'FC%' and tc_opc00 <> 'Y' and tc_cif_05 > (select nvl(max(tc_cif_05),to_date('2010/01/01','yyyy/mm/dd')) as t1 from tc_cif_file C left join oeb_file D on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
        oCommand.CommandText += "where tc_cif_01 not like 'FC%' and oeb70 <> 'Y'  and d.oeb04 = b.ta_opd14 ) and tc_prl02 = (select min(tc_prl02) from tc_prl_file where tc_prl01 = ta_opd14 and tc_prl02 > tc_cif_05 ) "
        oCommand.CommandText += ") group by tqa02,ta_opd14,ima02,ima021,gea02,ima31"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    If i > 5 Then
                        Ws.Cells(LineZ, i + 3) = oReader.Item(i)
                    Else
                        Ws.Cells(LineZ, i + 2) = oReader.Item(i)
                    End If
                Next
                ' 右方加總
                Ws.Cells(LineZ, 9 + TotalWeek).FormulaR1C1 = "=SUM(RC[-" & TotalWeek & "]:RC[-1])"
                LineZ += 1
            End While
            ' 下方加總
            Ws.Cells(LineZ, 8) = "Total"
            Ws.Cells(LineZ, 9) = "=SUM(I7:I" & LineZ - 1 & ")"
            ' 複制
            oRng = Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9))
            oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9 + TotalWeek)), Type:=xlFillDefault)

            ' 加入 格式
            oRng = Ws.Range("I7", Ws.Cells(LineZ, 9 + TotalWeek))
            oRng.EntireColumn.NumberFormat = "#,##0_ ;[Red]-#,##0 "

            ' 劃線
            oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 9 + TotalWeek))
            oRng.Borders(xlEdgeLeft).LineStyle = xlNone
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlDouble
            oRng.Borders(xlEdgeRight).LineStyle = xlNone
            oRng.Borders(xlInsideHorizontal).LineStyle = xlNone
            oRng.Borders(xlInsideVertical).LineStyle = xlNone
        End If
        oReader.Close()
        ' C 到 最後一行 作自動判斷
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 9))
        oRng.EntireColumn.AutoFit()
        ' 加入 邏輯說明
        LineZ += 2
        Ws.Cells(LineZ, 2) = "报表逻辑备注："
        Ws.Cells(LineZ + 1, 2) = "1.产品预测销售金额=产品数量（cxmt811中数量栏位（tc_cif_04）显示的数量）*产品售价（cxmt809中价格栏位（tc_prl03）的价格）"
        Ws.Cells(LineZ + 2, 2) = "1.第5行显示的年度为报表当年"
        Ws.Cells(LineZ + 3, 2) = "2.第6行显示的周别为报表起始日期对应的周别至报表当年最后一个周别，报表起始日期之前的周别无需显示"
        Ws.Cells(LineZ + 4, 2) = "3.如果报表起始日期对应周别中有两部分日期组成：一部分日期小于报表起始日期，另外一部分日期大于等于报表起始日期。此份报表只需要抓取日期大于等于报表起始日期对应的资料"
        Ws.Cells(LineZ + 5, 2) = "4.如果报表当年最后一周对应的日期有跨年的情况，此份报表最后一个周别只需要抓取当年日期对应的资料。"
        Ws.Cells(LineZ + 6, 2) = "5.介于第3点和第4点之间的周别对应的日期，只需要按周别抓取资料即可，无需区别跨月的情况"
        Ws.Cells(LineZ + 7, 2) = "6.如果同一款产品编号在cxmt808（订单项次多角期输入）和cxmt811（销售预测单项次多交期输入）中都有符合以上条件的资料，则此份报表需要以该产品编号在cxmt808（订单项次多角期输入）中最后一笔交货日期（（tc_cif_05））的次日开始统计预测数量。产品编号在cxmt811中交期小于等于cxmt808中交货日期的资料无需要统计，且cxmt808（订单项次多角期输入）中订单单号栏位（oeb01）显示的订单已经无效了，需要排除在外"
        Ws.Cells(LineZ + 8, 2) = "7.根据第6点的要求找出产品编号在cxmt811（销售预测单项次多交期输入）中维护的交期（tc_cif_05），并该抓取到的交货日期（tc_cif_05）抓取该产品编号在cxmt809中最近一笔大于等于交期（tc_cif_05）的截止日期（tc_prl02）对应的价格（tc_prl03）。产品取价百分比为100%，cxmt811中交期（tc_cif_05）小于报表起始日期排除在外"
        Ws.Cells(LineZ + 9, 2) = "8.如果cxmt811（销售预测单项次多交期输入）中订单单号栏位（oeb01）显示的订单已经无效了，则需要排除该订单相关资料"
        Ws.Cells(LineZ + 10, 2) = "9.如果产品售价币别为USD,产品的销售金额按6.3汇率转RMB；如果产品售价币别为EUR,产品的销售金额按7.56汇率(即6.3*1.2)转RMB"
        Ws.Cells(LineZ + 11, 2) = "10.用截止日期（tc_prl02）对应的价格（tc_prl03）*交货日期（tc_cif_05）对应的数量（tc_cif_04）算出产品的计划销售金额，然后在按周别汇总每周的计划销售金额"
        Ws.Cells(LineZ + 12, 2) = "11.如果报表第一周对应的交期栏位（tc_cif_05）既有部分交货日期小于报表起始日期又有部分交期大于等于报表起始日期，此时只需要汇总交期大于等于报表起始日期对应的的数量"
        Ws.Cells(LineZ + 13, 2) = "12.如果报表最后一周有跨年的情况，需要把报表次年的数量排除在外"
        Ws.Cells(LineZ + 14, 2) = "13.根据产品编号在axmi121（产品主档维护作业）中抓取品牌代码，再用品牌代码在atmi402（基础编号代码维护作业）抓取客户名称"


        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ + 14, 2))
        oRng.HorizontalAlignment = xlLeft
        ' 凍結
        oRng = Ws.Range("D6", "D6")
        oRng.Select()
        xExcel.ActiveWindow.FreezePanes = True

        ' 第六頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(6)
        Ws.Name = "forecast shipping by date"
        Ws.Activate()
        AdjustExcelFormat5()

        oCommand.CommandText = "select tqa02,B.ta_opd14,ima02,opd01,gea02,ima31,tc_cif_05,tc_cif_04,azn05,tc_prl03,tc_prl06 "
        oCommand.CommandText += "from tc_cif_file A left join opd_file B on tc_cif_01 = opd01 and tc_cif_02 = opd05 left join ima_file on ta_opd14 = ima01 "
        oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = 2 left join opc_file on opd01 = opc01 left join occ_file on opc02 = occ01 "
        oCommand.CommandText += "left join gea_file on occ20 = gea01 left join azn_file on tc_cif_05 = azn01 left join tc_prl_file on tc_prl01 = ta_opd14 and tc_prl02 > tc_cif_05 where tc_cif_05 >= to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_05 <= to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_01 like 'FC%' and tc_opc00 <> 'Y' and tc_cif_05 > (select nvl(max(tc_cif_05),to_date('2010/01/01','yyyy/mm/dd')) as t1 from tc_cif_file C left join oeb_file D on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
        oCommand.CommandText += "where tc_cif_01 not like 'FC%' and oeb70 <> 'Y'  and d.oeb04 = b.ta_opd14 ) and tc_prl02 = (select min(tc_prl02) from tc_prl_file where tc_prl01 = ta_opd14 and tc_prl02 > tc_cif_05 ) "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 2) = oReader.Item(i)
                Next
                LineZ += 1
            End While
            ' 下方加總
            Ws.Cells(LineZ, 8) = "Total"
            Ws.Cells(LineZ, 9) = "=SUM(I6:I" & LineZ - 1 & ")"

            ' 加入 格式
            oRng = Ws.Range("I6", "I6")
            oRng.EntireColumn.NumberFormat = "#,##0_ ;[Red]-#,##0 "
            oRng = Ws.Range("K6", "K6")
            oRng.EntireColumn.NumberFormat = "#,##0_ ;[Red]-#,##0 "

            ' 劃線
            oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 12))
            oRng.Borders(xlEdgeLeft).LineStyle = xlNone
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlDouble
            oRng.Borders(xlEdgeRight).LineStyle = xlNone
            oRng.Borders(xlInsideHorizontal).LineStyle = xlNone
            oRng.Borders(xlInsideVertical).LineStyle = xlNone
        End If
        oReader.Close()
        ' C 到 最後一行 作自動判斷
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 9))
        oRng.EntireColumn.AutoFit()
        ' 加入 邏輯說明
        LineZ += 2
        Ws.Cells(LineZ, 2) = "报表逻辑备注："
        Ws.Cells(LineZ + 1, 2) = "1.把cxmt808（订单项次多角期输入）中交货日期（tc_cif_05）属于报表起始日期~当年12/31日期间的产品编号抓取出来"
        Ws.Cells(LineZ + 2, 2) = "2.订单号、交货日期、交货数量、周别抓取cxmt808（订单项次多角期输入）程式中订单号（oeb01）、交货日期（tc_cif_05）、交货数量（tc_cif_04）、周别栏位（azn05）的信息"
        Ws.Cells(LineZ + 3, 2) = "3.产品售价、币别抓取cxmt809（料件价格表）程式中价格（tc_prl03）、币别栏位（tc_prl06）的信息，产品取价百分比为100%"
        Ws.Cells(LineZ + 4, 2) = "4.同一个产品编号有不同订单需要分行显示"
        Ws.Cells(LineZ + 5, 2) = "5.同一个产品编号有不同交货日期也需要分行显示"
        Ws.Cells(LineZ + 6, 2) = "6.如果cxmt808（订单项次多角期输入）中订单单号（oeb01）栏位显示的订单已经无效了，则需要排除该订单相关资料"
        Ws.Cells(LineZ + 7, 2) = "7.产品售价取数原则：根据产品编号及其交货日期（tc_cif_05），抓取该产品编号在cxmt809中最近一笔大于等于交货日期（tc_cif_05）的截止日期（tc_prl02）对应的价格（tc_prl03）。cxmt808中交货日期（tc_cif_05）小于报表起始日期排除在外"
        Ws.Cells(LineZ + 8, 2) = "8.根据产品编号在axmi121（产品主档维护作业）中抓取品牌代码，再用品牌代码在atmi402（基础编号代码维护作业）抓取客户名称"

        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ + 8, 2))
        oRng.HorizontalAlignment = xlLeft
        ' 凍結
        oRng = Ws.Range("E6", "E6")
        oRng.Select()
        xExcel.ActiveWindow.FreezePanes = True

        ' 第七頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(7)
        Ws.Name = "actual delivery"
        Ws.Activate()
        AdjustExcelFormat6()

        oCommand.CommandText = "select ogbplant,oga01,oga02,oga04,occ02,ogaud02,ogb31,ogb03,ogb04,ogb06,ima021,ogb12,ogb05,oga23,oga24,"
        oCommand.CommandText += "ogbud08,ogbud09,ogb13,ogb14,round(ogb14 * oga24, 2),ogbud05,ima1005,tqa02,'','','','成本仓' from oga_file left join ogb_file on oga01 = ogb01 "
        oCommand.CommandText += "left join occ_file on oga04 = occ01 left join ima_file on ogb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' where ogapost = 'Y' and oga02 between to_date('"
        oCommand.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += tDate3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ogb09 not in (select jce02 from jce_file) "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select ohbplant,oha01,oha02,oha04,occ02,ohb30,ohb33,ohb03,ohb04,ohb06,ima021,ohb12*-1,ohb05,oha23,oha24,"
        oCommand.CommandText += "0,0,ohb13,ohb14,round(ohb14 * oha24,2), ohbud05, ima1005,tqa02,'','','','成本仓' from oha_file left join ohb_file on oha01 = ohb01 "
        oCommand.CommandText += "left join occ_file on oha04 = occ01 left join ima_file on ohb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' where ohapost = 'Y' and oha02 between to_date('"
        oCommand.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += tDate3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ohb09 not in (select jce02 from jce_file)"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 2) = oReader.Item(i)
                Next
                LineZ += 1
                End While
                ' 下方加總
            Ws.Cells(LineZ, 12) = "Total"
            Ws.Cells(LineZ, 13) = "=SUM(M7:M" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 20) = "=SUM(T7:T" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 21) = "=SUM(U7:U" & LineZ - 1 & ")"

                ' 加入 格式
            oRng = Ws.Range("Q7", "U7")
            oRng.EntireColumn.NumberFormat = "#,##0_ ;[Red]-#,##0 "

                ' 劃線
            oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 28))
            oRng.Borders(xlEdgeLeft).LineStyle = xlNone
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlDouble
            oRng.Borders(xlEdgeRight).LineStyle = xlNone
            oRng.Borders(xlInsideHorizontal).LineStyle = xlNone
            oRng.Borders(xlInsideVertical).LineStyle = xlNone
        End If
        oReader.Close()
            ' C 到 最後一行 作自動判斷
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 28))
        oRng.EntireColumn.AutoFit()
            ' 加入 邏輯說明
        LineZ += 2
        Ws.Cells(LineZ, 2) = "报表逻辑备注："
        Ws.Cells(LineZ + 1, 2) = "1.已审核已过账的多角出货单（axmt820a）、单边出货单（axmt620a）、多角销退单（axmt840）、单边销退单（axmt700）的出货金额和退货金额"
        Ws.Cells(LineZ + 2, 2) = "2.具体逻辑参看axmr620报表"
        Ws.Cells(LineZ + 3, 2) = "3.如果出货单或者销退单上的仓别属于非成本仓，则该数据不要显示出来"
        Ws.Cells(LineZ + 4, 2) = "4.取数的起始时间：aglp301中关账日期次日，取数的截止时间：为报表起始日期的前一日"

        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ + 4, 2))
        oRng.HorizontalAlignment = xlLeft
            ' 凍結
        oRng = Ws.Range("K7", "K7")
        oRng.Select()
        xExcel.ActiveWindow.FreezePanes = True

        ' 第八頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(8)
        'Ws.Name = "Cost by part"
        Ws.Name = "cost & Rolling forecast amount"
        Ws.Activate()
        AdjustExcelFormat7()
        Dim TotalMonth As Decimal = 13 - LMonth
        oCommand.CommandText = "SELECT tqa02,ogb04,ima02,ima021,ogb05,ccc23,c1,c2"
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",sum(a" & i & ") as a" & i
        Next
        oCommand.CommandText += " from ( select tqa02,ogb04,ima02,ima021,ogb05,ccc23,(case when ccc23 is null or ccc23 = 0 then (stb07 + stb08 + stb09a + stb09) else null end) as c1,(case when (ccc23 is null or ccc23 = 0) and ((stb07 + stb08 + stb09a + stb09) is null or (stb07 + stb08 + stb09a + stb09) = 0) then  (ogb13 * oga24) else null end) as c2"
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(oga02) = " & i & " then ogb12 else 0 end) as t" & i
        Next
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(oga02) = " & i & " then ogb14 * oga24 else 0 end) as a" & i
        Next
        oCommand.CommandText += " from oga_file left join ogb_file on oga01 = ogb01 left join ima_file on ogb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand.CommandText += "left join ccc_file on ccc01 = ogb04 and ccc02 = " & pYear & " and ccc03 = " & pMonth & " left join stb_file on stb01 = ogb04 and stb02 = " & pYear & " and stb03 = " & pMonth
        oCommand.CommandText += " where ogapost = 'Y' and oga02 between to_date('"
        oCommand.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += tDate3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ogb09 not in (select jce02 from jce_file) "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select tqa02,ohb04,ima02,ima021,ohb05,ccc23,(case when ccc23 is null then (stb07 + stb08 + stb09a + stb09) else null end) as c1,(case when (ccc23 is null or ccc23 = 0) and ((stb07 + stb08 + stb09a + stb09) is null or (stb07 + stb08 + stb09a + stb09) = 0) then  (ohb13 * oha24) else null end) as c2"
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(oha02) = " & i & " then ohb12 * -1 else 0 end) as t" & i
        Next
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(oha02) = " & i & " then ohb14 * oha24 *-1 end) as a" & i
        Next
        oCommand.CommandText += " from oha_file left join ohb_file on oha01 = ohb01 left join ima_file on ohb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand.CommandText += "left join ccc_file on ccc01 = ohb04 and ccc02 = " & pYear & " and ccc03 = " & pMonth & " left join stb_file on stb01 = ohb04 and stb02 = " & pYear & " and stb03 = " & pMonth
        oCommand.CommandText += " where ohapost = 'Y' and oha02 between to_date('"
        oCommand.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += tDate3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ohb09 not in (select jce02 from jce_file) "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select tqa02,oeb04,ima02,ima021,oeb05,ccc23,(case when ccc23 is null or ccc23 = 0 then (stb07 + stb08 + stb09a + stb09) else null end) as c1,(case when (ccc23 is null or ccc23 = 0) and ((stb07 + stb08 + stb09a + stb09) is null or (stb07 + stb08 + stb09a + stb09) = 0) then  tc_prl03 * (case when tc_prl06 = 'USD' THEN 6.3 when tc_prl06 = 'EUR' then 7.56 else 0 end) else null end) as c2"
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(tc_cif_05) = " & i & " then tc_cif_04 else 0 end) as t" & i
        Next
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(tc_cif_05) = " & i & " then tc_cif_04 * tc_prl03 * (case when tc_prl06 = 'USD' then 6.3 when tc_prl06 = 'EUR' then 7.56 else 0 end)  else 0 end) as a" & i
        Next
        oCommand.CommandText += " from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 left join ima_file on oeb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand.CommandText += "left join ccc_file on ccc01 = oeb04 and ccc02 = " & pYear & " and ccc03 = " & pMonth & " left join stb_file on stb01 = oeb04 and stb02 = " & pYear & " and stb03 = " & pMonth
        oCommand.CommandText += " left join tc_prl_file on tc_prl01 = oeb04 and tc_prl02 > tc_cif_05 where tc_cif_05 >= to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_05 <= to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_prl02 = (select min(tc_prl02) from tc_prl_file where tc_prl01 = oeb04 and tc_prl02 > tc_cif_05) "
        oCommand.CommandText += "and oeb70 <> 'Y' and tc_cif_01 not like 'FC%' "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select tqa02,B.ta_opd14,ima02,ima021,ima31,ccc23,(case when ccc23 is null or ccc23 = 0 then (stb07 + stb08 + stb09a + stb09) else null end) as c1,(case when (ccc23 is null or ccc23 = 0) and ((stb07 + stb08 + stb09a + stb09) is null or (stb07 + stb08 + stb09a + stb09) = 0) then tc_prl03 * (case when tc_prl06 = 'USD' THEN 6.3 when tc_prl06 = 'EUR' then 7.56 else 0 end) else null end) as c2"
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(tc_cif_05) = " & i & " then tc_cif_04 else 0 end) as t" & i
        Next
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(tc_cif_05) = " & i & " then tc_cif_04 * tc_prl03 * (case when tc_prl06 = 'USD' then 6.3 when tc_prl06 = 'EUR' then 7.56 else 0 end)  else 0 end) as a" & i
        Next
        oCommand.CommandText += " from tc_cif_file A left join opd_file B on tc_cif_01 = opd01 and tc_cif_02 = opd05 left join ima_file on ta_opd14 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand.CommandText += "left join opc_file on opd01 = opc01 left join ccc_file on ccc01 = ta_opd14 and ccc02 = " & pYear & " and ccc03 = " & pMonth & " left join stb_file on stb01 = ta_opd14 and stb02 = " & tYear & " and stb03 = " & tMonth
        oCommand.CommandText += " left join tc_prl_file on tc_prl01 = ta_opd14 and tc_prl02 > tc_cif_05 where tc_cif_05 >= to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_05 <= to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_01 like 'FC%' and tc_opc00 <> 'Y' "
        oCommand.CommandText += "and tc_cif_05 > (select nvl(max(tc_cif_05),to_date('2010/01/01','yyyy/mm/dd')) as t1 from tc_cif_file C left join oeb_file D on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
        oCommand.CommandText += "where tc_cif_01 not like 'FC%' and oeb70 <> 'Y'  and d.oeb04 = b.ta_opd14 ) and tc_prl02 = (select min(tc_prl02) from tc_prl_file where tc_prl01 = ta_opd14 and tc_prl02 > tc_cif_05 ) "
        oCommand.CommandText += ") group by tqa02,ogb04,ima02,ima021,ogb05,ccc23,c1,c2"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("tqa02")
                Ws.Cells(LineZ, 3) = oReader.Item("ogb04")
                Ws.Cells(LineZ, 4) = oReader.Item("ima02")
                Ws.Cells(LineZ, 5) = oReader.Item("ima021")
                Ws.Cells(LineZ, 6) = oReader.Item("ogb05")
                If IsDBNull(oReader.Item("ccc23")) Then
                    If IsDBNull(oReader.Item("c1")) Then
                        Ws.Cells(LineZ, 7) = oReader.Item("c2")
                        Ws.Cells(LineZ, 8) = "产品RMB售价"
                    Else
                        If oReader.Item("c1") = 0 Then
                            Ws.Cells(LineZ, 7) = oReader.Item("c2")
                            Ws.Cells(LineZ, 8) = "产品RMB售价"
                        Else
                            Ws.Cells(LineZ, 7) = oReader.Item("c1")
                            Ws.Cells(LineZ, 8) = "标准单位成本"
                        End If

                    End If
                Else
                    If oReader.Item("ccc23") = 0 Then
                        If IsDBNull(oReader.Item("c1")) Then
                            Ws.Cells(LineZ, 7) = oReader.Item("c2")
                            Ws.Cells(LineZ, 8) = "产品RMB售价"
                        Else
                            If oReader.Item("c1") = 0 Then
                                Ws.Cells(LineZ, 7) = oReader.Item("c2")
                                Ws.Cells(LineZ, 8) = "产品RMB售价"
                            Else
                                Ws.Cells(LineZ, 7) = oReader.Item("c1")
                                Ws.Cells(LineZ, 8) = "标准单位成本"
                            End If
                        End If
                    Else
                        Ws.Cells(LineZ, 7) = oReader.Item("ccc23")
                        Ws.Cells(LineZ, 8) = "实际单位成本"
                    End If
                End If
                'Ws.Cells(LineZ, 8) = oReader.Item("ccc23") & "/" & oReader.Item("c1") & "/" & oReader.Item("c2")

                For i As Int16 = 8 To 7 + TotalMonth Step 1
                    Ws.Cells(LineZ, i + 1) = "=" & oReader.Item(i) & "*G" & LineZ
                Next
                For i As Int16 = 8 + TotalMonth To 7 + 2 * TotalMonth Step 1
                    Ws.Cells(LineZ, i + 2) = oReader.Item(i)
                Next
                ' 右方加總
                Ws.Cells(LineZ, 22 - LMonth).FormulaR1C1 = "=SUM(RC[-" & 13 - LMonth & "]:RC[-1])"
                Ws.Cells(LineZ, 10 + 2 * TotalMonth).FormulaR1C1 = "=SUM(RC[-" & TotalMonth & "]:RC[-1])"
                Ws.Cells(LineZ, 11 + 2 * TotalMonth).FormulaR1C1 = "=(RC[-1]-RC[-" & 2 + TotalMonth & "])"
                LineZ += 1

            End While

                        ' 下方加總
            Ws.Cells(LineZ, 8) = "Total"
            Ws.Cells(LineZ, 9) = "=SUM(I7:I" & LineZ - 1 & ")"
            ' 複制
            oRng = Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9))
            oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 11 + 2 * TotalMonth)), Type:=xlFillDefault)

            ' 加入 格式
            oRng = Ws.Range("G7", Ws.Cells(LineZ, 7))
            oRng.NumberFormat = "#,##0_ ;[Red]-#,##0.00 "
            oRng = Ws.Range("I7", Ws.Cells(LineZ, 11 + 2 * TotalMonth))
            oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "

                        ' 劃線
            oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 11 + 2 * TotalMonth))
            oRng.Borders(xlEdgeLeft).LineStyle = xlNone
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlDouble
            oRng.Borders(xlEdgeRight).LineStyle = xlNone
            oRng.Borders(xlInsideHorizontal).LineStyle = xlNone
            oRng.Borders(xlInsideVertical).LineStyle = xlNone
                    End If
        oReader.Close()
                    ' C 到 最後一行 作自動判斷
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 11 + 2 * TotalMonth))
        oRng.EntireColumn.AutoFit()
                    ' 加入 邏輯說明
        LineZ += 2
        Ws.Cells(LineZ, 2) = "Cost by part报表逻辑备注："
        Ws.Cells(LineZ + 1, 2) = "1.aglp301中关账日期次日至报表起始日期前一日的销售数量等于已审核已过账的多角出货单（axmt820a）、单边出货单（axmt620a）、多角销退单（axmt840）、单边销退单（axmt700）中累计销售数量"
        Ws.Cells(LineZ + 2, 2) = "2.报表起始日期至报表当年12/31日的销售数量等于cxmt808（订单项次多角期输入）中大于等于报表起始日期的交货日期（tc_cif_05）对应的数量（tc_cif_04），cxmt808（订单项次多角期输入）中订单单号栏位（oeb01）显示的订单已经无效了，需要排除在外。"
        Ws.Cells(LineZ + 3, 2) = "3.同时加上cxmt811（销售预测单项次多交期输入）中大于等于报表起始日期的交期（tc_cif_05）对应的数量（tc_cif_04）。如果同一款产品编号在cxmt808（订单项次多角期输入）和cxmt811（销售预测单项次多交期输入）中都有符合以上条件的资料，则此份报表需要以该产品编号在cxmt808（订单项次多角期输入）中最后一笔交货日期（（tc_cif_05））的次日开始统计预测数量。产品编号在cxmt811中交期小于等于cxmt808中交货日期的资料无需要统计"
        Ws.Cells(LineZ + 4, 2) = "3.报表当月之前的月份无需显示出来"
        Ws.Cells(LineZ + 5, 2) = "5.单位成本为报表月份前一个月实际单位成本；如果前一个月无实际单位成本，则取前一个月的标准单位成本；如果连标准单位成本也没有，则直接取产品的RMB售价。同时在备注栏位显示实际单位成本/标准单位成本/产品RMB售价。红色字体部分改成如果标准单位成本为零，则直接取产品的RMB售价（产品有不同的RMB售价需要分行显示）。"
        Ws.Cells(LineZ + 6, 2) = "6.如果产品售价币别为USD,产品的销售金额按6.3汇率转RMB；如果产品售价币别为EUR,产品的销售金额按7.56汇率(即6.3*1.2)转RMB"
        Ws.Cells(LineZ + 7, 2) = "7.成本金额=单位成本*对应月份的销售数量"
        Ws.Cells(LineZ + 8, 2) = "8.根据产品编号在axmi121（产品主档维护作业）中抓取品牌代码，再用品牌代码在atmi402（基础编号代码维护作业）抓取客户名称"
        Ws.Cells(LineZ + 9, 2) = "Rolling forecast amount by part报表逻辑备注："
        Ws.Cells(LineZ + 10, 2) = "1.aglp301中关账日期次日至报表起始日期前一日的销售金额等于已审核已过账的多角出货单（axmt820a）、单边出货单（axmt620a）、多角销退单（axmt840）、单边销退单（axmt700）的出货金额和退货金额"
        Ws.Cells(LineZ + 11, 2) = "2.aglp301中关账日期次日至报表起始日期前一日的汇率：产品售价币别为USD,产品的销售金额按aooi060中银行中介汇率（每月汇率维护作业）转RMB；产品售价币别为EUR,产品的销售金额按aooi060中银行中介汇率（每月汇率维护作业）转RMB"
        Ws.Cells(LineZ + 12, 2) = "3.报表起始日期至报表当年12/31日的销售数量等于cxmt808（订单项次多角期输入）和cxmt811（销售预测单项次多交期输入）中对应日期的计划数量*cxmt809（料件价格表）对应的截止日期的产品售价，然后在按月份分别汇总销售金额"
        Ws.Cells(LineZ + 13, 2) = "4.cxmt808（订单项次多角期输入）订单单号oeb01中的订单已经无效的及cxmt811（销售预测单项次多交期输入）预测单单号opd01中预测订单已经无效的排除在外"
        Ws.Cells(LineZ + 14, 2) = "5.如果同一款产品编号在cxmt808（订单项次多角期输入）和cxmt811（销售预测单项次多交期输入）中都有符合以上条件的资料，则此份报表需要以该产品编号在cxmt808（订单项次多角期输入）中最后一笔交货日期（（tc_cif_05））的次日开始统计预测数量"
        Ws.Cells(LineZ + 15, 2) = "6.报表起始日期至报表当年12/31日的汇率：如果产品售价币别为USD,产品的销售金额按6.3汇率转RMB；如果产品售价币别为EUR,产品的销售金额按7.56汇率(即6.3*1.2)转RMB"
        Ws.Cells(LineZ + 16, 2) = "7.相同产品编号只需要显示一行"
        Ws.Cells(LineZ + 17, 2) = "8.根据产品编号在axmi121（产品主档维护作业）中抓取品牌代码，再用品牌代码在atmi402（基础编号代码维护作业）抓取客户名称"
        Ws.Cells(LineZ + 18, 2) = "9.各产品编号各月累计销售金额-各月累计销售成本"
        Ws.Cells(LineZ + 19, 2) = "10.销售金额请参照每月销售金额报表逻辑"
        Ws.Cells(LineZ + 20, 2) = "11.销售成本请参照cost by part报表逻辑"

        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ + 20, 2))
        oRng.HorizontalAlignment = xlLeft
                    ' 凍結
        oRng = Ws.Range("D7", "D7")
        oRng.Select()
        xExcel.ActiveWindow.FreezePanes = True

        '' 第九頁
        'Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        'Ws = xWorkBook.Sheets(9)
        'Ws.Name = "rolling forecast amount by part"
        'Ws.Activate()
        'AdjustExcelFormat8()

        'oCommand.CommandText = "SELECT tqa02,ogb04,ima02,ima021,ogb05,ccc23,c1,c2"
        'For i As Int16 = LMonth To 12 Step 1
        '    oCommand.CommandText += ",sum(t" & i & ") as t" & i
        'Next
        'oCommand.CommandText += " from ( select tqa02,ogb04,ima02,ima021,ogb05,ccc23,(case when ccc23 is null then (stb07 + stb08 + stb09a + stb09) else null end) as c1,(case when ccc23 is null and (stb07 + stb08 + stb09a + stb09) is null then  (ogb13 * oga24) else null end) as c2"
        'For i As Int16 = LMonth To 12 Step 1
        '    oCommand.CommandText += ",(case when month(oga02) = " & i & " then ogb12 else 0 end) as t" & i
        'Next
        'oCommand.CommandText += " from oga_file left join ogb_file on oga01 = ogb01 left join ima_file on ogb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        'oCommand.CommandText += "left join ccc_file on ccc01 = ogb04 and ccc02 = " & pYear & " and ccc03 = " & pMonth & " left join stb_file on stb01 = ogb04 and stb02 = " & pYear & " and stb03 = " & pMonth
        'oCommand.CommandText += " where ogapost = 'Y' and oga02 between to_date('"
        'oCommand.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        'oCommand.CommandText += tDate3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ogb09 not in (select jce02 from jce_file) "
        'oCommand.CommandText += "union all "
        'oCommand.CommandText += "select tqa02,ohb04,ima02,ima021,ohb05,ccc23,(case when ccc23 is null then (stb07 + stb08 + stb09a + stb09) else null end) as c1,(case when ccc23 is null and (stb07 + stb08 + stb09a + stb09) is null then  (ohb13 * oha24) else null end) as c2"
        'For i As Int16 = LMonth To 12 Step 1
        '    oCommand.CommandText += ",(case when month(oha02) = " & i & " then ohb12 * -1 else 0 end) as t" & i
        'Next
        'oCommand.CommandText += " from oha_file left join ohb_file on oha01 = ohb01 left join ima_file on ohb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        'oCommand.CommandText += "left join ccc_file on ccc01 = ohb04 and ccc02 = " & pYear & " and ccc03 = " & pMonth & " left join stb_file on stb01 = ohb04 and stb02 = " & pYear & " and stb03 = " & pMonth
        'oCommand.CommandText += " where ohapost = 'Y' and oha02 between to_date('"
        'oCommand.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        'oCommand.CommandText += tDate3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ohb09 not in (select jce02 from jce_file) "
        'oCommand.CommandText += "union all "
        'oCommand.CommandText += "select tqa02,oeb04,ima02,ima021,oeb05,ccc23,(case when ccc23 is null then (stb07 + stb08 + stb09a + stb09) else null end) as c1,(case when ccc23 is null and (stb07 + stb08 + stb09a + stb09) is null then  tc_prl03 * (case when tc_prl06 = 'USD' THEN 6.3 when tc_prl06 = 'EUR' then 7.56 else 0 end) else null end) as c2"
        'For i As Int16 = LMonth To 12 Step 1
        '    oCommand.CommandText += ",(case when month(tc_cif_05) = " & i & " then tc_cif_04 else 0 end) as t" & i
        'Next
        'oCommand.CommandText += " from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 left join ima_file on oeb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        'oCommand.CommandText += "left join ccc_file on ccc01 = oeb04 and ccc02 = " & pYear & " and ccc03 = " & pMonth & " left join stb_file on stb01 = oeb04 and stb02 = " & pYear & " and stb03 = " & pMonth
        'oCommand.CommandText += " left join tc_prl_file on tc_prl01 = oeb04 and tc_prl02 > tc_cif_05 where tc_cif_05 >= to_date('"
        'oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_05 <= to_date('"
        'oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_prl02 = (select min(tc_prl02) from tc_prl_file where tc_prl01 = oeb04 and tc_prl02 > tc_cif_05) "
        'oCommand.CommandText += "and oeb70 <> 'Y' and tc_cif_01 not like 'FC%' "
        'oCommand.CommandText += "union all "
        'oCommand.CommandText += "select tqa02,B.ta_opd14,ima02,ima021,ima31,ccc23,(case when ccc23 is null then (stb07 + stb08 + stb09a + stb09) else null end) as c1,(case when ccc23 is null and (stb07 + stb08 + stb09a + stb09) is null then  tc_prl03 * (case when tc_prl06 = 'USD' THEN 6.3 when tc_prl06 = 'EUR' then 7.56 else 0 end) else null end) as c2"
        'For i As Int16 = LMonth To 12 Step 1
        '    oCommand.CommandText += ",(case when month(tc_cif_05) = " & i & " then tc_cif_04 else 0 end) as t" & i
        'Next
        'oCommand.CommandText += " from tc_cif_file A left join opd_file B on tc_cif_01 = opd01 and tc_cif_02 = opd05 left join ima_file on ta_opd14 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        'oCommand.CommandText += "left join opc_file on opd01 = opc01 left join ccc_file on ccc01 = ta_opd14 and ccc02 = " & pYear & " and ccc03 = " & pMonth & " left join stb_file on stb01 = ta_opd14 and stb02 = " & tYear & " and stb03 = " & tMonth
        'oCommand.CommandText += " left join tc_prl_file on tc_prl01 = ta_opd14 and tc_prl02 > tc_cif_05 where tc_cif_05 >= to_date('"
        'oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_05 <= to_date('"
        'oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_01 like 'FC%' and tc_opc00 <> 'Y' "
        'oCommand.CommandText += "and tc_cif_05 > (select nvl(max(tc_cif_05),to_date('2010/01/01','yyyy/mm/dd')) as t1 from tc_cif_file C left join oeb_file D on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
        'oCommand.CommandText += "where tc_cif_01 not like 'FC%' and oeb70 <> 'Y'  and d.oeb04 = b.ta_opd14 ) and tc_prl02 = (select min(tc_prl02) from tc_prl_file where tc_prl01 = ta_opd14 and tc_prl02 > tc_cif_05 ) "
        'oCommand.CommandText += ") group by tqa02,ogb04,ima02,ima021,ogb05,ccc23,c1,c2"

        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        Ws.Cells(LineZ, 2) = oReader.Item("tqa02")
        '        Ws.Cells(LineZ, 3) = oReader.Item("ogb04")
        '        Ws.Cells(LineZ, 4) = oReader.Item("ima02")
        '        Ws.Cells(LineZ, 5) = oReader.Item("ima021")
        '        Ws.Cells(LineZ, 6) = oReader.Item("ogb05")
        '        If IsDBNull(oReader.Item("ccc23")) Then
        '            If IsDBNull(oReader.Item("c1")) Then
        '                Ws.Cells(LineZ, 7) = oReader.Item("c2")
        '                Ws.Cells(LineZ, 8) = "产品RMB售价"
        '            Else
        '                Ws.Cells(LineZ, 7) = oReader.Item("c1")
        '                Ws.Cells(LineZ, 8) = "标准单位成本"
        '            End If
        '        Else
        '            Ws.Cells(LineZ, 7) = oReader.Item("ccc23")
        '            Ws.Cells(LineZ, 8) = "实际单位成本"
        '        End If
        '        'Ws.Cells(LineZ, 8) = oReader.Item("ccc23") & "/" & oReader.Item("c1") & "/" & oReader.Item("c2")

        '        For i As Int16 = 8 To oReader.FieldCount - 1 Step 1
        '            Ws.Cells(LineZ, i + 1) = "=" & oReader.Item(i) & "*G" & LineZ
        '        Next
        '        ' 右方加總
        '        Ws.Cells(LineZ, 22 - LMonth).FormulaR1C1 = "=SUM(RC[-" & 13 - LMonth & "]:RC[-1])"
        '        LineZ += 1
        '    End While

        '    ' 下方加總
        '    Ws.Cells(LineZ, 8) = "Total"
        '    Ws.Cells(LineZ, 9) = "=SUM(I7:I" & LineZ - 1 & ")"
        '    ' 複制
        '    oRng = Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9))
        '    oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 22 - LMonth)), Type:=xlFillDefault)

        '    ' 加入 格式
        '    oRng = Ws.Range("I7", Ws.Cells(LineZ, 22 - LMonth))
        '    oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "

        '    ' 劃線
        '    oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 22 - LMonth))
        '    oRng.Borders(xlEdgeLeft).LineStyle = xlNone
        '    oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        '    oRng.Borders(xlEdgeBottom).LineStyle = xlDouble
        '    oRng.Borders(xlEdgeRight).LineStyle = xlNone
        '    oRng.Borders(xlInsideHorizontal).LineStyle = xlNone
        '    oRng.Borders(xlInsideVertical).LineStyle = xlNone
        'End If
        'oReader.Close()
        '' C 到 最後一行 作自動判斷
        'oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 22 - LMonth))
        'oRng.EntireColumn.AutoFit()
        '' 加入 邏輯說明
        'LineZ += 2
        'Ws.Cells(LineZ, 2) = "报表逻辑备注："
        'Ws.Cells(LineZ + 1, 2) = "1.aglp301中关账日期次日至报表起始日期前一日的销售数量等于已审核已过账的多角出货单（axmt820a）、单边出货单（axmt620a）、多角销退单（axmt840）、单边销退单（axmt700）中累计销售数量"
        'Ws.Cells(LineZ + 2, 2) = "2.报表起始日期至报表当年12/31日的销售数量等于cxmt808（订单项次多角期输入）中大于等于报表起始日期的交货日期（tc_cif_05）对应的数量（tc_cif_04），cxmt808（订单项次多角期输入）中订单单号栏位（oeb01）显示的订单已经无效了，需要排除在外。"
        'Ws.Cells(LineZ + 3, 2) = "3.同时加上cxmt811（销售预测单项次多交期输入）中大于等于报表起始日期的交期（tc_cif_05）对应的数量（tc_cif_04）。如果同一款产品编号在cxmt808（订单项次多角期输入）和cxmt811（销售预测单项次多交期输入）中都有符合以上条件的资料，则此份报表需要以该产品编号在cxmt808（订单项次多角期输入）中最后一笔交货日期（（tc_cif_05））的次日开始统计预测数量。产品编号在cxmt811中交期小于等于cxmt808中交货日期的资料无需要统计"
        'Ws.Cells(LineZ + 4, 2) = "3.报表当月之前的月份无需显示出来"
        'Ws.Cells(LineZ + 5, 2) = "5.单位成本为报表月份前一个月实际单位成本；如果前一个月无实际单位成本，则取前一个月的标准单位成本；如果连标准单位成本也没有，则直接取产品的RMB售价。同时在备注栏位显示实际单位成本/标准单位成本/产品RMB售价"
        'Ws.Cells(LineZ + 6, 2) = "6.如果产品售价币别为USD,产品的销售金额按6.3汇率转RMB；如果产品售价币别为EUR,产品的销售金额按7.56汇率(即6.3*1.2)转RMB"
        'Ws.Cells(LineZ + 7, 2) = "7.成本金额=单位成本*对应月份的销售数量"
        'Ws.Cells(LineZ + 8, 2) = "8.根据产品编号在axmi121（产品主档维护作业）中抓取品牌代码，再用品牌代码在atmi402（基础编号代码维护作业）抓取客户名称"

        'oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ + 8, 2))
        'oRng.HorizontalAlignment = xlLeft
        '' 凍結
        'oRng = Ws.Range("D7", "D7")
        'oRng.Select()
        'xExcel.ActiveWindow.FreezePanes = True

        ' 第九頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(9)
        Ws.Name = "margin by customer"
        Ws.Activate()
        AdjustExcelFormat8()
        oCommand.CommandText = "select tqa02,sum(h1) as h1 from ( "
        oCommand.CommandText += "select tqa02,sum("
        For i As Int16 = LMonth To 12 Step 1
                oCommand.CommandText += "+a" & i
        Next
        oCommand.CommandText += ") - (sum(case when ccc23 is not null AND ccc23 <> 0 then ccc23 when c1 is not null AND c1 <> 0 then c1 when c2 is not null then c2 end) * ("
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += "+t" & i
        Next
        oCommand.CommandText += ")) as h1 from ( "
        oCommand.CommandText += "SELECT tqa02,ogb04,ima02,ima021,ogb05,ccc23,c1,c2"
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",sum(a" & i & ") as a" & i
        Next
        oCommand.CommandText += " from ( select tqa02,ogb04,ima02,ima021,ogb05,ccc23,(case when ccc23 is null or ccc23 = 0 then (stb07 + stb08 + stb09a + stb09) else null end) as c1,(case when (ccc23 is null or ccc23 = 0) and ((stb07 + stb08 + stb09a + stb09) is null or (stb07 + stb08 + stb09a + stb09) = 0) then  (ogb13 * oga24) else null end) as c2"
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(oga02) = " & i & " then ogb12 else 0 end) as t" & i
        Next
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(oga02) = " & i & " then ogb14 * oga24 else 0 end) as a" & i
        Next
        oCommand.CommandText += " from oga_file left join ogb_file on oga01 = ogb01 left join ima_file on ogb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand.CommandText += "left join ccc_file on ccc01 = ogb04 and ccc02 = " & pYear & " and ccc03 = " & pMonth & " left join stb_file on stb01 = ogb04 and stb02 = " & pYear & " and stb03 = " & pMonth
        oCommand.CommandText += " where ogapost = 'Y' and oga02 between to_date('"
        oCommand.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += tDate3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ogb09 not in (select jce02 from jce_file) "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select tqa02,ohb04,ima02,ima021,ohb05,ccc23,(case when ccc23 is null or ccc23 = 0 then (stb07 + stb08 + stb09a + stb09) else null end) as c1,(case when (ccc23 is null or ccc23 = 0) and ((stb07 + stb08 + stb09a + stb09) is null or (stb07 + stb08 + stb09a + stb09) = 0) then  (ohb13 * oha24) else null end) as c2"
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(oha02) = " & i & " then ohb12 * -1 else 0 end) as t" & i
        Next
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(oha02) = " & i & " then ohb14 * oha24 *-1 end) as a" & i
        Next
        oCommand.CommandText += " from oha_file left join ohb_file on oha01 = ohb01 left join ima_file on ohb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand.CommandText += "left join ccc_file on ccc01 = ohb04 and ccc02 = " & pYear & " and ccc03 = " & pMonth & " left join stb_file on stb01 = ohb04 and stb02 = " & pYear & " and stb03 = " & pMonth
        oCommand.CommandText += " where ohapost = 'Y' and oha02 between to_date('"
        oCommand.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += tDate3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ohb09 not in (select jce02 from jce_file) "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select tqa02,oeb04,ima02,ima021,oeb05,ccc23,(case when ccc23 is null or ccc23 = 0 then (stb07 + stb08 + stb09a + stb09) else null end) as c1,(case when (ccc23 is null or ccc23 = 0) and ((stb07 + stb08 + stb09a + stb09) is null or (stb07 + stb08 + stb09a + stb09) = 0) then  tc_prl03 * (case when tc_prl06 = 'USD' THEN 6.3 when tc_prl06 = 'EUR' then 7.56 else 0 end) else null end) as c2"
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(tc_cif_05) = " & i & " then tc_cif_04 else 0 end) as t" & i
        Next
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(tc_cif_05) = " & i & " then tc_cif_04 * tc_prl03 * (case when tc_prl06 = 'USD' then 6.3 when tc_prl06 = 'EUR' then 7.56 else 0 end)  else 0 end) as a" & i
        Next
        oCommand.CommandText += " from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 left join ima_file on oeb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand.CommandText += "left join ccc_file on ccc01 = oeb04 and ccc02 = " & pYear & " and ccc03 = " & pMonth & " left join stb_file on stb01 = oeb04 and stb02 = " & pYear & " and stb03 = " & pMonth
        oCommand.CommandText += " left join tc_prl_file on tc_prl01 = oeb04 and tc_prl02 > tc_cif_05 where tc_cif_05 >= to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_05 <= to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_prl02 = (select min(tc_prl02) from tc_prl_file where tc_prl01 = oeb04 and tc_prl02 > tc_cif_05) "
        oCommand.CommandText += "and oeb70 <> 'Y' and tc_cif_01 not like 'FC%' "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select tqa02,B.ta_opd14,ima02,ima021,ima31,ccc23,(case when ccc23 is null or ccc23 = 0 then (stb07 + stb08 + stb09a + stb09) else null end) as c1,(case when (ccc23 is null or ccc23 = 0) and ((stb07 + stb08 + stb09a + stb09) is null or (stb07 + stb08 + stb09a + stb09) = 0) then  tc_prl03 * (case when tc_prl06 = 'USD' THEN 6.3 when tc_prl06 = 'EUR' then 7.56 else 0 end) else null end) as c2"
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(tc_cif_05) = " & i & " then tc_cif_04 else 0 end) as t" & i
        Next
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(tc_cif_05) = " & i & " then tc_cif_04 * tc_prl03 * (case when tc_prl06 = 'USD' then 6.3 when tc_prl06 = 'EUR' then 7.56 else 0 end)  else 0 end) as a" & i
        Next
        oCommand.CommandText += " from tc_cif_file A left join opd_file B on tc_cif_01 = opd01 and tc_cif_02 = opd05 left join ima_file on ta_opd14 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand.CommandText += "left join opc_file on opd01 = opc01 left join ccc_file on ccc01 = ta_opd14 and ccc02 = " & pYear & " and ccc03 = " & pMonth & " left join stb_file on stb01 = ta_opd14 and stb02 = " & tYear & " and stb03 = " & tMonth
        oCommand.CommandText += " left join tc_prl_file on tc_prl01 = ta_opd14 and tc_prl02 > tc_cif_05 where tc_cif_05 >= to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_05 <= to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_01 like 'FC%' and tc_opc00 <> 'Y' "
        oCommand.CommandText += "and tc_cif_05 > (select nvl(max(tc_cif_05),to_date('2010/01/01','yyyy/mm/dd')) as t1 from tc_cif_file C left join oeb_file D on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
        oCommand.CommandText += "where tc_cif_01 not like 'FC%' and oeb70 <> 'Y'  and d.oeb04 = b.ta_opd14 ) and tc_prl02 = (select min(tc_prl02) from tc_prl_file where tc_prl01 = ta_opd14 and tc_prl02 > tc_cif_05 ) "
        oCommand.CommandText += ") group by tqa02,ogb04,ima02,ima021,ogb05,ccc23,c1,c2 ) group by tqa02"
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",t" & i
        Next
        oCommand.CommandText += " ) group by tqa02 order by h1 desc"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("tqa02")
                Ws.Cells(LineZ, 3) = oReader.Item("h1")
                LineZ += 1

            End While

                ' 下方加總
            Ws.Cells(LineZ, 2) = "Total"
            Ws.Cells(LineZ, 3) = "=SUM(C7:C" & LineZ - 1 & ")"

                ' 加入 格式
            oRng = Ws.Range("C7", Ws.Cells(LineZ, 3))
            oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "

                ' 劃線
            oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 3))
            oRng.Borders(xlEdgeLeft).LineStyle = xlNone
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlDouble
            oRng.Borders(xlEdgeRight).LineStyle = xlNone
            oRng.Borders(xlInsideHorizontal).LineStyle = xlNone
            oRng.Borders(xlInsideVertical).LineStyle = xlNone
            End If
        oReader.Close()
            ' C 到 最後一行 作自動判斷
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
        oRng.EntireColumn.AutoFit()
            ' 加入 邏輯說明
        LineZ += 2
        Ws.Cells(LineZ, 2) = "报表逻辑备注："
        Ws.Cells(LineZ + 1, 2) = "1.客户各月累计销售金额-各月累计销售成本"
        Ws.Cells(LineZ + 2, 2) = "2.销售金额请参照每月销售金额报表逻辑，然后按客户汇总销售金额"
        Ws.Cells(LineZ + 3, 2) = "3.销售成本请参照cost by part报表逻辑，然后按客户汇总销售成本"
        Ws.Cells(LineZ + 4, 2) = "4.根据产品编号在axmi121（产品主档维护作业）中抓取品牌代码，再用品牌代码在atmi402（基础编号代码维护作业）抓取客户名称"

        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ + 4, 2))
        oRng.HorizontalAlignment = xlLeft
            ' 凍結
        'oRng = Ws.Range("D7", "D7")
        'oRng.Select()
        'xExcel.ActiveWindow.FreezePanes = True

        ' 第十頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(10)
        Ws.Name = "Budget"
        Ws.Activate()
        AdjustExcelFormat9()
        DoInputData2("600101", "600102", 1)
        LineZ += 1
        DoInputData2("6051", "6099", 1)
        LineZ += 9
        DoInputData2("640101", "6403", 1)
        LineZ += 4
        DoInputData2("660101", "660199", 1)
        LineZ += 1
        DoInputData2("660201", "660299", 1)
        LineZ += 1
        DoInputData2("660301", "660303", 1)
        LineZ += 1
        DoInputData2("660401", "660499", 1)
        LineZ += 4
        DoInputData2("6301", "6301", 1)
        LineZ += 2
        DoInputData2("6711", "6711", 1)
        LineZ += 4
        DoInputData2("6801", "6801", 1)

        ' C 到 最後一行 作自動判斷
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 35))
        oRng.EntireColumn.AutoFit()
        ' 加入 邏輯說明
        LineZ += 4
        Ws.Cells(LineZ, 2) = "报表逻辑备注："
        Ws.Cells(LineZ + 1, 2) = "1.所有资料都是抓取cgli600中维护的预算资料"
        Ws.Cells(LineZ + 2, 2) = "2.请参考外挂程式-ERP报表-总账-IS报表中预算利润表逻辑"
        Ws.Cells(LineZ + 3, 2) = "3.所有的%栏位公式需要加iferror函数"

        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ + 3, 2))
        oRng.HorizontalAlignment = xlLeft
        ' 凍結
        oRng = Ws.Range("D6", "D6")
        oRng.Select()
        xExcel.ActiveWindow.FreezePanes = True

        ' 第十一頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(11)
        Ws.Name = "A+F"
        Ws.Activate()
        AdjustExcelFormat9()
        DoInputData("600101", "600102", 1)
        LineZ += 1
        DoInputData("6051", "6099", 1)
        LineZ += 9
        DoInputData("640101", "6403", 0)
        LineZ += 4
        DoInputData("660101", "660199", 0)
        LineZ += 1
        DoInputData("660201", "660299", 0)
        LineZ += 1
        DoInputData("660301", "660303", 0)

        LineZ += 1
        DoInputData("660401", "660499", 0)
        LineZ += 1
        DoInputData("6701", "6701", 0)
        LineZ += 3
        DoInputData("6301", "6301", 1)
        LineZ += 2
        DoInputData("6711", "6711", 0)
        LineZ += 4
        DoInputData("6801", "6801", 0)

        oCommand.CommandText = "select "
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += "sum(a" & i & ") as a" & i & ","
        Next
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += "sum(b" & i & ") as b" & i & ","
        Next
        oCommand.CommandText += "1 from ( select "
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += "sum(a" & i & ") as a" & i & ","
        Next
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += "sum(case when ccc23 is not null AND ccc23 <> 0 then ccc23 when c1 is not null AND c1 <> 0 then c1 when c2 is not null then c2 end) * (t" & i & ") as b" & i & ","
        Next
        
        oCommand.CommandText += "1 from ( "
        oCommand.CommandText += "SELECT tqa02,ogb04,ima02,ima021,ogb05,ccc23,c1,c2"
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",sum(a" & i & ") as a" & i
        Next
        oCommand.CommandText += " from ( select tqa02,ogb04,ima02,ima021,ogb05,ccc23,(case when ccc23 is null or ccc23 = 0 then (stb07 + stb08 + stb09a + stb09) else null end) as c1,(case when (ccc23 is null or ccc23 = 0) and ((stb07 + stb08 + stb09a + stb09) is null or (stb07 + stb08 + stb09a + stb09) = 0) then  (ogb13 * oga24) else null end) as c2"
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(oga02) = " & i & " then ogb12 else 0 end) as t" & i
        Next
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(oga02) = " & i & " then ogb14 * oga24 else 0 end) as a" & i
        Next
        oCommand.CommandText += " from oga_file left join ogb_file on oga01 = ogb01 left join ima_file on ogb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand.CommandText += "left join ccc_file on ccc01 = ogb04 and ccc02 = " & pYear & " and ccc03 = " & pMonth & " left join stb_file on stb01 = ogb04 and stb02 = " & pYear & " and stb03 = " & pMonth
        oCommand.CommandText += " where ogapost = 'Y' and oga02 between to_date('"
        oCommand.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += tDate3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ogb09 not in (select jce02 from jce_file) "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select tqa02,ohb04,ima02,ima021,ohb05,ccc23,(case when ccc23 is null or ccc23 = 0 then (stb07 + stb08 + stb09a + stb09) else null end) as c1,(case when (ccc23 is null or ccc23 = 0) and ((stb07 + stb08 + stb09a + stb09) is null or (stb07 + stb08 + stb09a + stb09) = 0) then  (ohb13 * oha24) else null end) as c2"
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(oha02) = " & i & " then ohb12 * -1 else 0 end) as t" & i
        Next
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(oha02) = " & i & " then ohb14 * oha24 *-1 end) as a" & i
        Next
        oCommand.CommandText += " from oha_file left join ohb_file on oha01 = ohb01 left join ima_file on ohb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand.CommandText += "left join ccc_file on ccc01 = ohb04 and ccc02 = " & pYear & " and ccc03 = " & pMonth & " left join stb_file on stb01 = ohb04 and stb02 = " & pYear & " and stb03 = " & pMonth
        oCommand.CommandText += " where ohapost = 'Y' and oha02 between to_date('"
        oCommand.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += tDate3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ohb09 not in (select jce02 from jce_file) "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select tqa02,oeb04,ima02,ima021,oeb05,ccc23,(case when ccc23 is null or ccc23 = 0 then (stb07 + stb08 + stb09a + stb09) else null end) as c1,(case when (ccc23 is null or ccc23 = 0) and ((stb07 + stb08 + stb09a + stb09) is null or (stb07 + stb08 + stb09a + stb09) = 0) then  tc_prl03 * (case when tc_prl06 = 'USD' THEN 6.3 when tc_prl06 = 'EUR' then 7.56 else 0 end) else null end) as c2"
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(tc_cif_05) = " & i & " then tc_cif_04 else 0 end) as t" & i
        Next
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(tc_cif_05) = " & i & " then tc_cif_04 * tc_prl03 * (case when tc_prl06 = 'USD' then 6.3 when tc_prl06 = 'EUR' then 7.56 else 0 end)  else 0 end) as a" & i
        Next
        oCommand.CommandText += " from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 left join ima_file on oeb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand.CommandText += "left join ccc_file on ccc01 = oeb04 and ccc02 = " & pYear & " and ccc03 = " & pMonth & " left join stb_file on stb01 = oeb04 and stb02 = " & pYear & " and stb03 = " & pMonth
        oCommand.CommandText += " left join tc_prl_file on tc_prl01 = oeb04 and tc_prl02 > tc_cif_05 where tc_cif_05 >= to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_05 <= to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_prl02 = (select min(tc_prl02) from tc_prl_file where tc_prl01 = oeb04 and tc_prl02 > tc_cif_05) "
        oCommand.CommandText += "and oeb70 <> 'Y' and tc_cif_01 not like 'FC%' "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select tqa02,B.ta_opd14,ima02,ima021,ima31,ccc23,(case when ccc23 is null or ccc23 = 0 then (stb07 + stb08 + stb09a + stb09) else null end) as c1,(case when (ccc23 is null or ccc23 = 0) and ((stb07 + stb08 + stb09a + stb09) is null or (stb07 + stb08 + stb09a + stb09) = 0) then  tc_prl03 * (case when tc_prl06 = 'USD' THEN 6.3 when tc_prl06 = 'EUR' then 7.56 else 0 end) else null end) as c2"
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(tc_cif_05) = " & i & " then tc_cif_04 else 0 end) as t" & i
        Next
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += ",(case when month(tc_cif_05) = " & i & " then tc_cif_04 * tc_prl03 * (case when tc_prl06 = 'USD' then 6.3 when tc_prl06 = 'EUR' then 7.56 else 0 end)  else 0 end) as a" & i
        Next
        oCommand.CommandText += " from tc_cif_file A left join opd_file B on tc_cif_01 = opd01 and tc_cif_02 = opd05 left join ima_file on ta_opd14 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand.CommandText += "left join opc_file on opd01 = opc01 left join ccc_file on ccc01 = ta_opd14 and ccc02 = " & pYear & " and ccc03 = " & pMonth & " left join stb_file on stb01 = ta_opd14 and stb02 = " & tYear & " and stb03 = " & tMonth
        oCommand.CommandText += " left join tc_prl_file on tc_prl01 = ta_opd14 and tc_prl02 > tc_cif_05 where tc_cif_05 >= to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_05 <= to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_01 like 'FC%' and tc_opc00 <> 'Y' "
        oCommand.CommandText += "and tc_cif_05 > (select nvl(max(tc_cif_05),to_date('2010/01/01','yyyy/mm/dd')) as t1 from tc_cif_file C left join oeb_file D on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
        oCommand.CommandText += "where tc_cif_01 not like 'FC%' and oeb70 <> 'Y'  and d.oeb04 = b.ta_opd14 ) and tc_prl02 = (select min(tc_prl02) from tc_prl_file where tc_prl01 = ta_opd14 and tc_prl02 > tc_cif_05 ) "
        oCommand.CommandText += ") group by tqa02,ogb04,ima02,ima021,ogb05,ccc23,c1,c2 ) group by "
        For i As Int16 = LMonth To 12 Step 1
            oCommand.CommandText += "t" & i & ","
        Next
        oCommand.CommandText += "1 )"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 2 Step 1
                    If i < TotalMonth Then
                        Ws.Cells(6, 2 * (i + LMonth) + 2) = oReader.Item(i)
                    Else
                        Ws.Cells(16, 2 * (i + LMonth - TotalMonth) + 2) = oReader.Item(i)
                    End If
                Next
            End While
        End If
        oReader.Close()

        ' C 到 最後一行 作自動判斷
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 35))
        oRng.EntireColumn.AutoFit()
        ' 加入 邏輯說明
        LineZ += 4
        Ws.Cells(LineZ, 2) = "报表逻辑备注："
        Ws.Cells(LineZ + 1, 2) = "1.aglp301中关账日期之前的月份（含关账日期当月），各项数据抓取每月实际发生的金额。请参考外挂程式-ERP报表-总账-IS报表中实际利润表逻辑"
        Ws.Cells(LineZ + 2, 2) = "2.aglp301中关账日期次日至报表起始日期前一日的计划和预测销售金额抓取已审核已过账的多角出货单（axmt820a）、单边出货单（axmt620a）、多角销退单（axmt840）、单边销退单（axmt700）的出货金额和退货金额"
        Ws.Cells(LineZ + 3, 2) = "3.aglp301中关账日期次日至报表起始日期前一日的汇率：产品售价币别为USD,产品的销售金额按aooi060中银行中介汇率（每月汇率维护作业）转RMB；产品售价币别为EUR,产品的销售金额按aooi060中银行中介汇率（每月汇率维护作业）转RMB"
        Ws.Cells(LineZ + 4, 2) = "4.aglp301中关账日期次日至报表起始日期前一日的计划和预测销售成本等于已审核已过账的多角出货单（axmt820a）、单边出货单（axmt620a）、多角销退单（axmt840）、单边销退单（axmt700）的出货数量和退货数量*单位成本"
        Ws.Cells(LineZ + 5, 2) = "5.报表起始日期至报表当年12/31日的计划和预测销售金额等于cxmt808（订单项次多角期输入）和cxmt811（销售预测单项次多交期输入）中对应日期的计划数量*cxmt809（料件价格表）对应的截止日期的产品售价，然后在按月份分别汇总销售金额"
        Ws.Cells(LineZ + 6, 2) = "6.报表起始日期至报表当年12/31日的汇率：如果产品售价币别为USD,产品的销售金额按6.3汇率转RMB；如果产品售价币别为EUR,产品的销售金额按7.56汇率(即6.3*1.2)转RMB"
        Ws.Cells(LineZ + 7, 2) = "7.报表起始日期至报表当年12/31日的计划和预测销售成本等于cxmt808（订单项次多角期输入）和cxmt811（销售预测单项次多交期输入）中对应日期的计划数量*单位成本"
        Ws.Cells(LineZ + 8, 2) = "8.单位成本为报表月份前一个月实际单位成本；如果前一个月无实际单位成本，则取前一个月的标准单位成本；如果连标准单位成本也没有，则直接取产品的RMB售价。如果产品售价币别为USD,产品的销售金额按6.3汇率转RMB；如果产品售价币别为EUR,产品的销售金额按7.56汇率(即6.3*1.2)转RMB"
        Ws.Cells(LineZ + 9, 2) = "9.所有的%栏位公式需要加iferror函数"
        Ws.Cells(LineZ + 10, 2) = "10.aglp301中关账日期的次月至本年度12月的营业费用、管理费用、财务费用、研发费用、资产减值损失、营业外收入、营业外支出、所得税栏位均为空值"
        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ + 10, 2))
        oRng.HorizontalAlignment = xlLeft
        ' 凍結
        oRng = Ws.Range("D6", "D6")
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
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 4
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 8


        oRng = Ws.Range("B3", "B3")
        oRng.HorizontalAlignment = xlLeft

        oRng = Ws.Range("A2", "A5")
        oRng.EntireRow.Font.Bold = True

        Ws.Cells(3, 2) = "Company Name：Dongguan Action Composite LTD. Co"
        Ws.Cells(4, 2) = "Customer"
        Ws.Cells(5, 2) = "客户"
        Ws.Cells(4, 3) = "Part Number"
        Ws.Cells(5, 3) = "产品编号"
        oRng = Ws.Range("C5", "C5")
        oRng.EntireColumn.NumberFormat = "@"
        Ws.Cells(4, 4) = "Part Description"
        Ws.Cells(5, 4) = "品名"
        Ws.Cells(4, 5) = "Spec."
        Ws.Cells(5, 5) = "规格"
        Ws.Cells(4, 6) = "Area"
        Ws.Cells(5, 6) = "区域"
        Ws.Cells(4, 7) = "Uint"
        Ws.Cells(5, 7) = "销售单位"
        Ws.Cells(4, 8) = "Year"
        Ws.Cells(5, 8) = "Weekly"

        oCommand.CommandText = "select distinct azn05 from azn_file where azn01 >= to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and azn02 = " & tYear & " order by azn05"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            TotalWeek = 0
            While oReader.Read()
                TotalWeek += 1
                Ws.Cells(4, TotalWeek + 8) = tYear
                Ws.Cells(5, TotalWeek + 8) = oReader.Item("azn05")
            End While
        End If
        oReader.Close()
        Ws.Cells(4, TotalWeek + 9) = tYear
        Ws.Cells(5, TotalWeek + 9) = "YTD"

        oRng = Ws.Range("B2", Ws.Cells(2, TotalWeek + 9))
        oRng.Merge()
        oRng.Font.Size = 16
        Ws.Cells(2, 2) = "Call off shipping qty by week"

        LineZ = 6
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Rows.RowHeight = 18
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 4
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 8


        oRng = Ws.Range("B3", "B4")
        oRng.HorizontalAlignment = xlLeft

        oRng = Ws.Range("A2", "A6")
        oRng.EntireRow.Font.Bold = True

        Ws.Cells(3, 2) = "Company Name：Dongguan Action Composite LTD. Co"
        Ws.Cells(4, 2) = "Currency:RMB"
        Ws.Cells(5, 2) = "Customer"
        Ws.Cells(6, 2) = "客户"
        Ws.Cells(5, 3) = "Part Number"
        Ws.Cells(6, 3) = "产品编号"
        oRng = Ws.Range("C5", "C5")
        oRng.EntireColumn.NumberFormat = "@"
        Ws.Cells(5, 4) = "Part Description"
        Ws.Cells(6, 4) = "品名"
        Ws.Cells(5, 5) = "Spec."
        Ws.Cells(6, 5) = "规格"
        Ws.Cells(5, 6) = "Area"
        Ws.Cells(6, 6) = "区域"
        Ws.Cells(5, 7) = "Uint"
        Ws.Cells(6, 7) = "销售单位"
        Ws.Cells(5, 8) = "Year"
        Ws.Cells(6, 8) = "Weekly"


        For i As Int16 = 1 To MaxWeek - tWeek + 1 Step 1
            Ws.Cells(5, i + 8) = tYear
            Ws.Cells(6, i + 8) = tWeek + i - 1
        Next
        Ws.Cells(5, TotalWeek + 9) = tYear
        Ws.Cells(6, TotalWeek + 9) = "YTD"

        oRng = Ws.Range("B2", Ws.Cells(2, TotalWeek + 9))
        oRng.Merge()
        oRng.Font.Size = 16
        Ws.Cells(2, 2) = "Call off shipping amount by week"

        LineZ = 7
    End Sub
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Rows.RowHeight = 18
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 4
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 8


        oRng = Ws.Range("B3", "B3")
        oRng.HorizontalAlignment = xlLeft

        oRng = Ws.Range("A2", "A5")
        oRng.EntireRow.Font.Bold = True

        Ws.Cells(3, 2) = "Company Name：Dongguan Action Composite LTD. Co"
        Ws.Cells(4, 2) = "Customer"
        Ws.Cells(5, 2) = "客户"
        Ws.Cells(4, 3) = "Part Number"
        Ws.Cells(5, 3) = "产品编号"
        oRng = Ws.Range("C5", "C5")
        oRng.EntireColumn.NumberFormat = "@"
        Ws.Cells(4, 4) = "Part Description"
        Ws.Cells(5, 4) = "品名"
        Ws.Cells(4, 5) = "Order No."
        Ws.Cells(5, 5) = "订单号"
        Ws.Cells(4, 6) = "Area"
        Ws.Cells(5, 6) = "区域"
        Ws.Cells(4, 7) = "Uint"
        Ws.Cells(5, 7) = "销售单位"
        Ws.Cells(4, 8) = "Shipment Date"
        Ws.Cells(5, 8) = "交货日期"
        Ws.Cells(4, 9) = "Shipment Qty"
        Ws.Cells(5, 9) = "交货数量"
        Ws.Cells(4, 10) = "Weekly"
        Ws.Cells(5, 10) = "周别"
        Ws.Cells(4, 11) = "Unit Price"
        Ws.Cells(5, 11) = "产品售价"
        Ws.Cells(4, 12) = "Currency"
        Ws.Cells(5, 12) = "币别"

        oRng = Ws.Range("B2", "L2")
        oRng.Merge()
        oRng.Font.Size = 16
        Ws.Cells(2, 2) = "Call off shipping by date"

        LineZ = 6
    End Sub
    Private Sub AdjustExcelFormat3()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Rows.RowHeight = 18
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 4
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 8


        oRng = Ws.Range("B3", "B3")
        oRng.HorizontalAlignment = xlLeft

        oRng = Ws.Range("A2", "A5")
        oRng.EntireRow.Font.Bold = True

        Ws.Cells(3, 2) = "Company Name：Dongguan Action Composite LTD. Co"
        Ws.Cells(4, 2) = "Customer"
        Ws.Cells(5, 2) = "客户"
        Ws.Cells(4, 3) = "Part Number"
        Ws.Cells(5, 3) = "产品编号"
        oRng = Ws.Range("C5", "C5")
        oRng.EntireColumn.NumberFormat = "@"
        Ws.Cells(4, 4) = "Part Description"
        Ws.Cells(5, 4) = "品名"
        Ws.Cells(4, 5) = "Spec."
        Ws.Cells(5, 5) = "规格"
        Ws.Cells(4, 6) = "Area"
        Ws.Cells(5, 6) = "区域"
        Ws.Cells(4, 7) = "Uint"
        Ws.Cells(5, 7) = "销售单位"
        Ws.Cells(4, 8) = "Year"
        Ws.Cells(5, 8) = "Weekly"

        oCommand.CommandText = "select distinct azn05 from azn_file where azn01 >= to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and azn02 = " & tYear & " order by azn05"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            TotalWeek = 0
            While oReader.Read()
                TotalWeek += 1
                Ws.Cells(4, TotalWeek + 8) = tYear
                Ws.Cells(5, TotalWeek + 8) = oReader.Item("azn05")
            End While
        End If
        oReader.Close()
        Ws.Cells(4, TotalWeek + 9) = tYear
        Ws.Cells(5, TotalWeek + 9) = "YTD"

        oRng = Ws.Range("B2", Ws.Cells(2, TotalWeek + 9))
        oRng.Merge()
        oRng.Font.Size = 16
        Ws.Cells(2, 2) = "Forecast shipping qty by week"

        LineZ = 6
    End Sub
    Private Sub AdjustExcelFormat4()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Rows.RowHeight = 18
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 4
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 8


        oRng = Ws.Range("B3", "B4")
        oRng.HorizontalAlignment = xlLeft

        oRng = Ws.Range("A2", "A6")
        oRng.EntireRow.Font.Bold = True

        Ws.Cells(3, 2) = "Company Name：Dongguan Action Composite LTD. Co"
        Ws.Cells(4, 2) = "Currency:RMB"
        Ws.Cells(5, 2) = "Customer"
        Ws.Cells(6, 2) = "客户"
        Ws.Cells(5, 3) = "Part Number"
        Ws.Cells(6, 3) = "产品编号"
        oRng = Ws.Range("C5", "C5")
        oRng.EntireColumn.NumberFormat = "@"
        Ws.Cells(5, 4) = "Part Description"
        Ws.Cells(6, 4) = "品名"
        Ws.Cells(5, 5) = "Spec."
        Ws.Cells(6, 5) = "规格"
        Ws.Cells(5, 6) = "Area"
        Ws.Cells(6, 6) = "区域"
        Ws.Cells(5, 7) = "Uint"
        Ws.Cells(6, 7) = "销售单位"
        Ws.Cells(5, 8) = "Year"
        Ws.Cells(6, 8) = "Weekly"


        For i As Int16 = 1 To MaxWeek - tWeek + 1 Step 1
            Ws.Cells(5, i + 8) = tYear
            Ws.Cells(6, i + 8) = tWeek + i - 1
        Next
        Ws.Cells(5, TotalWeek + 9) = tYear
        Ws.Cells(6, TotalWeek + 9) = "YTD"

        oRng = Ws.Range("B2", Ws.Cells(2, TotalWeek + 9))
        oRng.Merge()
        oRng.Font.Size = 16
        Ws.Cells(2, 2) = "Forecast shipping amount by week"

        LineZ = 7
    End Sub
    Private Sub AdjustExcelFormat5()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Rows.RowHeight = 18
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 4
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 8


        oRng = Ws.Range("B3", "B3")
        oRng.HorizontalAlignment = xlLeft

        oRng = Ws.Range("A2", "A5")
        oRng.EntireRow.Font.Bold = True

        Ws.Cells(3, 2) = "Company Name：Dongguan Action Composite LTD. Co"
        Ws.Cells(4, 2) = "Customer"
        Ws.Cells(5, 2) = "客户"
        Ws.Cells(4, 3) = "Part Number"
        Ws.Cells(5, 3) = "产品编号"
        oRng = Ws.Range("C5", "C5")
        oRng.EntireColumn.NumberFormat = "@"
        Ws.Cells(4, 4) = "Part Description"
        Ws.Cells(5, 4) = "品名"
        Ws.Cells(4, 5) = "Order No."
        Ws.Cells(5, 5) = "订单号"
        Ws.Cells(4, 6) = "Area"
        Ws.Cells(5, 6) = "区域"
        Ws.Cells(4, 7) = "Uint"
        Ws.Cells(5, 7) = "销售单位"
        Ws.Cells(4, 8) = "Shipment Date"
        Ws.Cells(5, 8) = "交货日期"
        Ws.Cells(4, 9) = "Shipment Qty"
        Ws.Cells(5, 9) = "交货数量"
        Ws.Cells(4, 10) = "Weekly"
        Ws.Cells(5, 10) = "周别"
        Ws.Cells(4, 11) = "Unit Price"
        Ws.Cells(5, 11) = "产品售价"
        Ws.Cells(4, 12) = "Currency"
        Ws.Cells(5, 12) = "币别"

        oRng = Ws.Range("B2", "L2")
        oRng.Merge()
        oRng.Font.Size = 16
        Ws.Cells(2, 2) = "Forecast shipping by date"

        LineZ = 6
    End Sub
    Private Sub AdjustExcelFormat6()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Rows.RowHeight = 18
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 4
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 8


        oRng = Ws.Range("B3", "B4")
        oRng.HorizontalAlignment = xlLeft

        oRng = Ws.Range("A2", "A6")
        oRng.EntireRow.Font.Bold = True

        Ws.Cells(3, 2) = "Company Name：Dongguan Action Composite LTD. Co"
        Ws.Cells(4, 2) = "Currency:RMB"
        Ws.Cells(5, 2) = "Operation Center"
        Ws.Cells(6, 2) = "营运中心"
        Ws.Cells(5, 3) = "Delivery Note No."
        Ws.Cells(6, 3) = "出货单号"
        Ws.Cells(5, 4) = "Shipment Date"
        Ws.Cells(6, 4) = "出货日期"
        Ws.Cells(5, 5) = "Customer Code"
        Ws.Cells(6, 5) = "送货客戶"
        Ws.Cells(5, 6) = "Customer Short Name"
        Ws.Cells(6, 6) = "客户简称"
        Ws.Cells(5, 7) = "AC Invoice No."
        Ws.Cells(6, 7) = "內部发票号"
        Ws.Cells(5, 8) = "AC Sales Order No."
        Ws.Cells(6, 8) = "订单单号"
        Ws.Cells(5, 9) = "Position No."
        Ws.Cells(6, 9) = "项次"
        Ws.Cells(5, 10) = "Part Name"
        Ws.Cells(6, 10) = "产品编号"
        oRng = Ws.Range("J5", "J5")
        oRng.EntireColumn.NumberFormat = "@"
        Ws.Cells(5, 11) = "Part Description"
        Ws.Cells(6, 11) = "品名"
        Ws.Cells(5, 12) = "Spec."
        Ws.Cells(6, 12) = "规格"
        Ws.Cells(5, 13) = "Shipping Qty"
        Ws.Cells(6, 13) = "出货数量"
        Ws.Cells(5, 14) = "Unit"
        Ws.Cells(6, 14) = "单位"
        Ws.Cells(5, 15) = "Currency"
        Ws.Cells(6, 15) = "币别"
        Ws.Cells(5, 16) = "Exchange Rate"
        Ws.Cells(6, 16) = "汇率"
        Ws.Cells(5, 17) = "Unit Price"
        Ws.Cells(6, 17) = "产品售价"
        Ws.Cells(5, 18) = "Tooling Prce"
        Ws.Cells(6, 18) = "模具加价"
        Ws.Cells(5, 19) = "Contract price"
        Ws.Cells(6, 19) = "合同单价"
        Ws.Cells(5, 20) = "Contract Amount (Original Currency)"
        Ws.Cells(6, 20) = "销售金额(原币)"
        Ws.Cells(5, 21) = "Amount (Domestic Currency)"
        Ws.Cells(6, 21) = "销售金额(本币)"
        Ws.Cells(5, 22) = "Customer PO No."
        Ws.Cells(6, 22) = "客戶订单号"
        Ws.Cells(5, 23) = "Brand No."
        Ws.Cells(6, 23) = "品牌代号"
        Ws.Cells(5, 24) = "Brand"
        Ws.Cells(6, 24) = "品牌说明"
        Ws.Cells(5, 25) = "Area"
        Ws.Cells(6, 25) = "区域"
        Ws.Cells(5, 26) = "Country"
        Ws.Cells(6, 26) = "国家"
        Ws.Cells(5, 27) = "District"
        Ws.Cells(6, 27) = "地区"
        Ws.Cells(5, 28) = "Warehouse Properties"
        Ws.Cells(6, 28) = "仓别属性"

        oRng = Ws.Range("B2", "AB2")
        oRng.Merge()
        oRng.Font.Size = 16
        Ws.Cells(2, 2) = "Forecast shipping by date"

        LineZ = 7
    End Sub
    Private Sub AdjustExcelFormat7()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Rows.RowHeight = 18
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 4
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 8


        oRng = Ws.Range("B3", "B4")
        oRng.HorizontalAlignment = xlLeft

        oRng = Ws.Range("A2", "A6")
        oRng.EntireRow.Font.Bold = True

        Ws.Cells(3, 2) = "Company Name：Dongguan Action Composite LTD. Co"
        Ws.Cells(4, 2) = "Currency:RMB"
        Ws.Cells(5, 2) = "Customer"
        Ws.Cells(6, 2) = "客户"
        Ws.Cells(5, 3) = "Part Number"
        Ws.Cells(6, 3) = "产品编号"
        oRng = Ws.Range("C5", "C5")
        oRng.EntireColumn.NumberFormat = "@"
        Ws.Cells(5, 4) = "Part Description"
        Ws.Cells(6, 4) = "品名"
        Ws.Cells(5, 5) = "Spec."
        Ws.Cells(6, 5) = "规格"
        Ws.Cells(5, 6) = "Uint"
        Ws.Cells(6, 6) = "单位"
        Ws.Cells(5, 7) = "Unit Cost"
        Ws.Cells(6, 7) = "单位成本"
        Ws.Cells(5, 8) = "Remarks"
        Ws.Cells(6, 8) = "备注"
        Dim TotalMonth As Decimal = 13 - LMonth
        For i As Int16 = 1 To 13 - LMonth Step 1
            Ws.Cells(5, i + 8) = "Cost"
            Ws.Cells(6, i + 8) = tYear & "/" & LMonth + i - 1 & "/01"
        Next
        Ws.Cells(5, 22 - LMonth) = "Cost"
        Ws.Cells(6, 22 - LMonth) = "YTD" & tYear
        For i As Int16 = 1 To 13 - LMonth Step 1
            Ws.Cells(5, 22 + i - LMonth) = "Amount"
            Ws.Cells(6, 22 + i - LMonth) = tYear & "/" & LMonth + i - 1 & "/01"
        Next
        Ws.Cells(5, 10 + 2 * TotalMonth) = "Amount"
        Ws.Cells(6, 10 + 2 * TotalMonth) = "YTD" & tYear
        Ws.Cells(5, 11 + 2 * TotalMonth) = "Margin"
        Ws.Cells(6, 11 + 2 * TotalMonth) = "毛利金额"
        '格式
        oRng = Ws.Range(Ws.Cells(6, 9), Ws.Cells(6, 21 - LMonth))
        oRng.NumberFormatLocal = "mmm-yy"
        oRng = Ws.Range(Ws.Cells(6, 23 - LMonth), Ws.Cells(6, 9 + 2 * TotalMonth))
        oRng.NumberFormatLocal = "mmm-yy"

        oRng = Ws.Range("B2", Ws.Cells(2, 11 + 2 * TotalMonth))
        oRng.Merge()
        oRng.Font.Size = 16
        Ws.Cells(2, 2) = "Cost by part & Rolling forecast amount by part"

        LineZ = 7
    End Sub
    Private Sub AdjustExcelFormat8()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Rows.RowHeight = 18
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 4
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 8


        oRng = Ws.Range("B3", "B4")
        oRng.HorizontalAlignment = xlLeft

        oRng = Ws.Range("A2", "A6")
        oRng.EntireRow.Font.Bold = True

        Ws.Cells(3, 2) = "Company Name：Dongguan Action Composite LTD. Co"
        Ws.Cells(4, 2) = "Currency:RMB"
        Ws.Cells(5, 2) = "Customer"
        Ws.Cells(6, 2) = "客户"
        Ws.Cells(5, 3) = "Margin"
        Ws.Cells(6, 3) = "毛利金额"

        oRng = Ws.Range("B2", "F2")
        oRng.Merge()
        oRng.Font.Size = 16
        Ws.Cells(2, 2) = "Margin by customer"

        LineZ = 7
    End Sub
    Private Sub AdjustExcelFormat9()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Rows.RowHeight = 18
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 4
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 8


        oRng = Ws.Range("B3", "B4")
        oRng.HorizontalAlignment = xlLeft

        oRng = Ws.Range("A2", "A6")
        oRng.EntireRow.Font.Bold = True

        Ws.Cells(3, 2) = "Company Name：Dongguan Action Composite LTD. Co"
        Ws.Cells(4, 2) = "Currency:RMB"
        Ws.Cells(5, 2) = "Account_Chinese"
        Ws.Cells(5, 3) = "Account_English"
        For i As Int16 = 1 To 12 Step 1
            If i < 10 Then
                Ws.Cells(5, 2 + 2 * i) = tYear & "/0" & i
            Else
                Ws.Cells(5, 2 + 2 * i) = tYear & "/" & i
            End If
            Ws.Cells(5, 3 + 2 * i) = "%"
        Next
        Ws.Cells(5, 28) = "Total"
        Ws.Cells(5, 29) = "%"

        oRng = Ws.Range("B2", "z2")
        oRng.Merge()
        oRng.Font.Size = 16
        Ws.Cells(2, 2) = "Income Statement"
        ' 添加格式
        oRng = Ws.Range("D6", "D35")
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        oRng = Ws.Range("F6", "F35")
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        oRng = Ws.Range("H6", "H35")
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        oRng = Ws.Range("J6", "J35")
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        oRng = Ws.Range("L6", "L35")
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        oRng = Ws.Range("N6", "N35")
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        oRng = Ws.Range("P6", "P35")
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        oRng = Ws.Range("R6", "R35")
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        oRng = Ws.Range("T6", "T35")
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        oRng = Ws.Range("V6", "V35")
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        oRng = Ws.Range("X6", "X35")
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        oRng = Ws.Range("Z6", "Z35")
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        oRng = Ws.Range("AB6", "AB35")
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "

        oRng = Ws.Range("E6", "E35")
        oRng.NumberFormat = "0.00%"
        oRng = Ws.Range("G6", "G35")
        oRng.NumberFormat = "0.00%"
        oRng = Ws.Range("I6", "I35")
        oRng.NumberFormat = "0.00%"
        oRng = Ws.Range("K6", "K35")
        oRng.NumberFormat = "0.00%"
        oRng = Ws.Range("M6", "M35")
        oRng.NumberFormat = "0.00%"
        oRng = Ws.Range("O6", "O35")
        oRng.NumberFormat = "0.00%"
        oRng = Ws.Range("Q6", "Q35")
        oRng.NumberFormat = "0.00%"
        oRng = Ws.Range("S6", "S35")
        oRng.NumberFormat = "0.00%"
        oRng = Ws.Range("U6", "U35")
        oRng.NumberFormat = "0.00%"
        oRng = Ws.Range("W6", "W35")
        oRng.NumberFormat = "0.00%"
        oRng = Ws.Range("Y6", "Y35")
        oRng.NumberFormat = "0.00%"
        oRng = Ws.Range("AA6", "AA35")
        oRng.NumberFormat = "0.00%"
        oRng = Ws.Range("AC6", "AC35")
        oRng.NumberFormat = "0.00%"
        ' Total 橫向加總
        Ws.Cells(6, 28) = "=D6+F6+H6+J6+L6+N6+P6+R6+T6+V6+X6+Z6"
        Ws.Cells(7, 28) = "=D7+F7+H7+J7+L7+N7+P7+R7+T7+V7+X7+Z7"
        Ws.Cells(8, 28) = "=D8+F8+H8+J8+L8+N8+P8+R8+T8+V8+X8+Z8"
        Ws.Cells(9, 28) = "=D9+F9+H9+J9+L9+N9+P9+R9+T9+V9+X9+Z9"

        Ws.Cells(12, 28) = "=D12+F12+H12+J12+L12+N12+P12+R12+T12+V12+X12+Z12"
        Ws.Cells(13, 28) = "=D13+F13+H13+J13+L13+N13+P13+R13+T13+V13+X13+Z13"
        Ws.Cells(14, 28) = "=D14+F14+H14+J14+L14+N14+P14+R14+T14+V14+X14+Z14"
        Ws.Cells(15, 28) = "=D15+F15+H15+J15+L15+N15+P15+R15+T15+V15+X15+Z15"
        Ws.Cells(16, 28) = "=D16+F16+H16+J16+L16+N16+P16+R16+T16+V16+X16+Z16"

        Ws.Cells(20, 28) = "=D20+F20+H20+J20+L20+N20+P20+R20+T20+V20+X20+Z20"
        Ws.Cells(21, 28) = "=D21+F21+H21+J21+L21+N21+P21+R21+T21+V21+X21+Z21"
        Ws.Cells(22, 28) = "=D22+F22+H22+J22+L22+N22+P22+R22+T22+V22+X22+Z22"
        Ws.Cells(23, 28) = "=D23+F23+H23+J23+L23+N23+P23+R23+T23+V23+X23+Z23"

        Ws.Cells(27, 28) = "=D27+F27+H27+J27+L27+N27+P27+R27+T27+V27+X27+Z27"
        Ws.Cells(29, 28) = "=D29+F29+H29+J29+L29+N29+P29+R29+T29+V29+X29+Z29"
        Ws.Cells(33, 28) = "=D33+F33+H33+J33+L33+N33+P33+R33+T33+V33+X33+Z33"

        ' 添加公式 和 copy 
        Ws.Cells(6, 5) = "=IFERROR(D6/D$10,0)"
        oRng = Ws.Range("E6", "E6")
        oRng.AutoFill(Destination:=Ws.Range("E6", "E35"), Type:=xlFillDefault)
        Ws.Cells(6, 7) = "=IFERROR(F6/F$10,0)"
        oRng = Ws.Range("G6", "G6")
        oRng.AutoFill(Destination:=Ws.Range("G6", "G35"), Type:=xlFillDefault)
        Ws.Cells(6, 9) = "=IFERROR(H6/H$10,0)"
        oRng = Ws.Range("I6", "I6")
        oRng.AutoFill(Destination:=Ws.Range("I6", "I35"), Type:=xlFillDefault)
        Ws.Cells(6, 11) = "=IFERROR(J6/J$10,0)"
        oRng = Ws.Range("K6", "K6")
        oRng.AutoFill(Destination:=Ws.Range("K6", "K35"), Type:=xlFillDefault)
        Ws.Cells(6, 13) = "=IFERROR(L6/L$10,0)"
        oRng = Ws.Range("M6", "M6")
        oRng.AutoFill(Destination:=Ws.Range("M6", "M35"), Type:=xlFillDefault)
        Ws.Cells(6, 15) = "=IFERROR(N6/N$10,0)"
        oRng = Ws.Range("O6", "O6")
        oRng.AutoFill(Destination:=Ws.Range("O6", "O35"), Type:=xlFillDefault)
        Ws.Cells(6, 17) = "=IFERROR(P6/P$10,0)"
        oRng = Ws.Range("Q6", "Q6")
        oRng.AutoFill(Destination:=Ws.Range("Q6", "Q35"), Type:=xlFillDefault)
        Ws.Cells(6, 19) = "=IFERROR(R6/R$10,0)"
        oRng = Ws.Range("S6", "S6")
        oRng.AutoFill(Destination:=Ws.Range("S6", "S35"), Type:=xlFillDefault)
        Ws.Cells(6, 21) = "=IFERROR(T6/T$10,0)"
        oRng = Ws.Range("U6", "U6")
        oRng.AutoFill(Destination:=Ws.Range("U6", "U35"), Type:=xlFillDefault)
        Ws.Cells(6, 23) = "=IFERROR(V6/V$10,0)"
        oRng = Ws.Range("W6", "W6")
        oRng.AutoFill(Destination:=Ws.Range("W6", "W35"), Type:=xlFillDefault)
        Ws.Cells(6, 25) = "=IFERROR(X6/X$10,0)"
        oRng = Ws.Range("Y6", "Y6")
        oRng.AutoFill(Destination:=Ws.Range("Y6", "Y35"), Type:=xlFillDefault)
        Ws.Cells(6, 27) = "=IFERROR(Z6/Z$10,0)"
        oRng = Ws.Range("AA6", "AA6")
        oRng.AutoFill(Destination:=Ws.Range("AA6", "AA35"), Type:=xlFillDefault)
        Ws.Cells(6, 29) = "=IFERROR(AB6/AB$10,0)"
        oRng = Ws.Range("AC6", "AC6")
        oRng.AutoFill(Destination:=Ws.Range("AC6", "AC35"), Type:=xlFillDefault)

        ' 格式2 橫
        Ws.Cells(10, 4) = "=SUM(D6:D9)"
        oRng = Ws.Range("D10", "E10")
        oRng.AutoFill(Destination:=Ws.Range("D10", "AB10"), Type:=xlFillDefault)
        Ws.Cells(17, 4) = "=D10-D16"
        oRng = Ws.Range("D17", "E17")
        oRng.AutoFill(Destination:=Ws.Range("D17", "AB17"), Type:=xlFillDefault)
        Ws.Cells(19, 4) = "=SUM(D20:D24)"
        oRng = Ws.Range("D19", "E19")
        oRng.AutoFill(Destination:=Ws.Range("D19", "AB19"), Type:=xlFillDefault)
        Ws.Cells(25, 4) = "=D17-D19"
        oRng = Ws.Range("D25", "E25")
        oRng.AutoFill(Destination:=Ws.Range("D25", "AB25"), Type:=xlFillDefault)
        Ws.Cells(31, 4) = "=D25+D27-D29"
        oRng = Ws.Range("D31", "E31")
        oRng.AutoFill(Destination:=Ws.Range("D31", "AB31"), Type:=xlFillDefault)
        Ws.Cells(35, 4) = "=D31-D33"
        oRng = Ws.Range("D35", "E35")
        oRng.AutoFill(Destination:=Ws.Range("D35", "AB35"), Type:=xlFillDefault)

        ' 添加文字
        Ws.Cells(6, 2) = "销售收入-产品"
        Ws.Cells(6, 3) = "Sales Revenue-Product"
        Ws.Cells(7, 2) = "销售收入-模具"
        Ws.Cells(7, 3) = "Sales Revenue-Tooling"
        Ws.Cells(8, 2) = "销货退回"
        Ws.Cells(8, 3) = "Sales Returns"
        Ws.Cells(9, 2) = "销售折让"
        Ws.Cells(9, 3) = "Sales Discounts"
        Ws.Cells(10, 2) = "销售收入"
        Ws.Cells(10, 3) = "Revenue"
        oRng = Ws.Range("B10", "C10")
        oRng.Font.Underline = True
        oRng.Font.Bold = True
        oRng = Ws.Range("B10", "AC10")
        oRng.Interior.Color = Color.FromArgb(220, 230, 241)

        Ws.Cells(12, 2) = "生产成本"
        Ws.Cells(12, 3) = "Product cost"
        Ws.Cells(13, 2) = "直接材料"
        Ws.Cells(13, 3) = "Materials"
        Ws.Cells(14, 2) = "直接人工"
        Ws.Cells(14, 3) = "Direct Labor"
        Ws.Cells(15, 2) = "制造费用"
        Ws.Cells(15, 3) = "Manufacturing Overhead"
        Ws.Cells(16, 2) = "销售成本"
        Ws.Cells(16, 3) = "Cost of goods sold"
        Ws.Cells(17, 2) = "毛利润"
        Ws.Cells(17, 3) = "Gross Margin"
        oRng = Ws.Range("B16", "C17")
        oRng.Font.Underline = True
        oRng.Font.Bold = True
        oRng = Ws.Range("B16", "AC17")
        oRng.Interior.Color = Color.FromArgb(220, 230, 241)

        Ws.Cells(19, 2) = "营业成本"
        Ws.Cells(19, 3) = "Operating Expenses"
        oRng = Ws.Range("B19", "C19")
        oRng.Font.Underline = True
        oRng.Font.Bold = True
        oRng = Ws.Range("B19", "AC19")
        oRng.Interior.Color = Color.FromArgb(220, 230, 241)

        Ws.Cells(20, 2) = "营业费用"
        Ws.Cells(20, 3) = "Selling Expenses"
        Ws.Cells(21, 2) = "管理费用"
        Ws.Cells(21, 3) = "General and Administration Exp"
        Ws.Cells(22, 2) = "财务费用"
        Ws.Cells(22, 3) = "Financial Rev/Exp"
        Ws.Cells(23, 2) = "研发费用"
        Ws.Cells(23, 3) = "R&D Exp"
        Ws.Cells(24, 2) = "资产减值损失"
        Ws.Cells(24, 3) = "Assets Devaluation"

        Ws.Cells(25, 2) = "营业利润"
        Ws.Cells(25, 3) = "Income from Operations"
        oRng = Ws.Range("B25", "C25")
        oRng.Font.Underline = True
        oRng.Font.Bold = True
        oRng = Ws.Range("B25", "AC25")
        oRng.Interior.Color = Color.FromArgb(220, 230, 241)

        Ws.Cells(27, 2) = "营业外收入"
        Ws.Cells(27, 3) = "Non-Operating Revenue"
        Ws.Cells(29, 2) = "营业外支出"
        Ws.Cells(29, 3) = "Non-Operating Expenses"
        Ws.Cells(31, 2) = "利润总额"
        Ws.Cells(31, 3) = "Profit Before Tax"
        oRng = Ws.Range("B31", "C31")
        oRng.Font.Underline = True
        oRng.Font.Bold = True
        oRng = Ws.Range("B31", "AC31")
        oRng.Interior.Color = Color.FromArgb(220, 230, 241)

        Ws.Cells(33, 2) = "所得税"
        Ws.Cells(33, 3) = "Income Tax"

        Ws.Cells(35, 2) = "净利润"
        Ws.Cells(35, 3) = "Net Profti/Loss After Tax"
        oRng = Ws.Range("B35", "C35")
        oRng.Font.Underline = True
        oRng.Font.Bold = True
        oRng = Ws.Range("B35", "AC35")
        oRng.Interior.Color = Color.FromArgb(220, 230, 241)

        ' 劃線
        oRng = Ws.Range("B5", "AC35")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        LineZ = 6
    End Sub
    Private Sub DoInputData2(ByVal ACC1 As String, ByVal ACC2 As String, ByVal ACC3 As Int16)

        oCommand.CommandText = "select "
        For i As Int16 = 1 To 12 Step 1
            oCommand.CommandText += "nvl(sum(t" & i & "),0) as t" & i & ","
        Next
        oCommand.CommandText += "1 from ( select "
        For i As Int16 = 1 To 12 Step 1
            oCommand.CommandText += "(case when tc_bud03 = " & i & " then (tc_bud13) else 0 end) as t" & i & ","
        Next
        oCommand.CommandText += "1 from tc_bud_file where tc_bud02 = " & tYear & " and tc_bud01 = 2 and tc_bud07 between '" & ACC1 & "' AND '" & ACC2 & "')"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                For i As Int16 = 1 To 12 Step 1
                    If ACC3 = 0 Then
                        Ws.Cells(LineZ, 2 * i + 2) = (oReader.Item(i - 1) * Decimal.MinusOne)
                    Else
                        Ws.Cells(LineZ, 2 * i + 2) = oReader.Item(i - 1)
                    End If

                Next
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub DoInputData(ByVal ACC1 As String, ByVal ACC2 As String, ByVal ACC3 As Int16)

        oCommand.CommandText = "select "
        For i As Int16 = 1 To aMonth Step 1
            oCommand.CommandText += "nvl(sum(t" & i & "),0) as t" & i & ","
        Next
        oCommand.CommandText += "1 from ( select "
        For i As Int16 = 1 To aMonth Step 1
            oCommand.CommandText += "(case when aah03 = " & i & " then (aah05 - aah04) else 0 end ) as t" & i & ","
        Next
        oCommand.CommandText += "1 from aah_file,aag_file where aah01 = aag01 and aag07 in ('2','3') and aah02 = " & tYear & " and aah03 <= " & tMonth & " and aah01 between '" & ACC1 & "' and '" & ACC2 & "' ) "

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                For i As Int16 = 1 To aMonth Step 1
                    If ACC3 = 0 Then
                        Ws.Cells(LineZ, 2 * i + 2) = (oReader.Item(i - 1) * Decimal.MinusOne)
                    Else
                        Ws.Cells(LineZ, 2 * i + 2) = oReader.Item(i - 1)
                    End If

                Next
            End While
        End If
        oReader.Close()
    End Sub
End Class