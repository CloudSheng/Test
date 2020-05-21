Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.XlContainsOperator
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlChartType
Imports Microsoft.Office.Core.MsoChartElementType

Public Class Form132
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
    Dim tMonth As Int16 = 0
    Dim pYear As Int16 = 0
    Dim TimeS1 As Date
    Dim TimeS2 As Date
    Dim TimeS3 As Date
    Dim TimeS4 As Date
    Dim gDatabase As String = String.Empty
    Dim gCurrency As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form132_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        'oConnection.ConnectionString = Module1.OpenOracleConnection("hkacttest")
        Me.TextBox1.Text = Today.Year
        Me.TextBox2.Text = Today.Month
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        gDatabase = Me.ComboBox1.SelectedItem.ToString()
        Select Case gDatabase
            Case "DAC"
                oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
                gCurrency = "RMB"
            Case "HAC"
                oConnection.ConnectionString = Module1.OpenOracleConnection("hkacttest")
                gCurrency = "USD"
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
        tYear = Me.TextBox1.Text
        tMonth = Me.TextBox2.Text
        pYear = tYear - 1
        TimeS1 = Convert.ToDateTime(tYear & "/01/01")
        TimeS2 = TimeS1.AddYears(1).AddDays(-1)
        TimeS3 = TimeS1.AddYears(-1)
        TimeS4 = TimeS2.AddYears(-1)
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
        oCommand.CommandText = "select tqa02,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,"
        oCommand.CommandText += "sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( "
        oCommand.CommandText += "select tqa02,(case when month(oga02) = 1 then ogb14 * oga24 else 0 end ) as t1,(case when month(oga02) = 2 then ogb14 * oga24 else 0 end ) as t2,"
        oCommand.CommandText += "(case when month(oga02) = 3 then ogb14 * oga24 else 0 end ) as t3,(case when month(oga02) = 4 then ogb14 * oga24 else 0 end ) as t4,"
        oCommand.CommandText += "(case when month(oga02) = 5 then ogb14 * oga24 else 0 end ) as t5,(case when month(oga02) = 6 then ogb14 * oga24 else 0 end ) as t6,"
        oCommand.CommandText += "(case when month(oga02) = 7 then ogb14 * oga24 else 0 end ) as t7,(case when month(oga02) = 8 then ogb14 * oga24 else 0 end ) as t8,"
        oCommand.CommandText += "(case when month(oga02) = 9 then ogb14 * oga24 else 0 end ) as t9,(case when month(oga02) = 10 then ogb14 * oga24 else 0 end ) as t10,"
        oCommand.CommandText += "(case when month(oga02) = 11 then ogb14 * oga24 else 0 end ) as t11,(case when month(oga02) = 12 then ogb14 * oga24 else 0 end ) as t12 from oga_file "
        oCommand.CommandText += "left join ogb_file on oga01 = ogb01 left join ima_file on ogb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand.CommandText += "where ogapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oga02) = " & tYear & " and month(oga02) <= " & tMonth & " "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select tqa02,(case when month(oha02) = 1 then ohb14 * oha24 * -1 else 0 end ) as t1,"
        oCommand.CommandText += "(case when month(oha02) = 2 then ohb14 * oha24 * -1 else 0 end ) as t2,(case when month(oha02) = 3 then ohb14 * oha24 * -1 else 0 end ) as t3,"
        oCommand.CommandText += "(case when month(oha02) = 4 then ohb14 * oha24 * -1 else 0 end ) as t4,(case when month(oha02) = 5 then ohb14 * oha24 * -1 else 0 end ) as t5,"
        oCommand.CommandText += "(case when month(oha02) = 6 then ohb14 * oha24 * -1 else 0 end ) as t6,(case when month(oha02) = 7 then ohb14 * oha24 * -1 else 0 end ) as t7,"
        oCommand.CommandText += "(case when month(oha02) = 8 then ohb14 * oha24 * -1 else 0 end ) as t8,(case when month(oha02) = 9 then ohb14 * oha24 * -1 else 0 end ) as t9,"
        oCommand.CommandText += "(case when month(oha02) = 10 then ohb14 * oha24 * -1 else 0 end ) as t10,(case when month(oha02) = 11 then ohb14 * oha24 * -1 else 0 end ) as t11,"
        oCommand.CommandText += "(case when month(oha02) = 12 then ohb14 * oha24 * -1 else 0 end ) as t12  from oha_file left join ohb_file on oha01 = ohb01  left join ima_file on ohb04 = ima01 "
        oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' where ohapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oha02) = " & tYear & " and month(oha02) <= " & tMonth & " ) group by tqa02 order by tqa02"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 1 To tMonth + 1 Step 1
                    Ws.Cells(LineZ, i) = oReader.Item(i - 1)
                Next
                Ws.Cells(LineZ, 14) = "=SUM(B" & LineZ & ":M" & LineZ & ")"
                LineZ += 1
            End While
        End If
        oReader.Close()
        AdjustExcelFormat2()
        LineZ += 1

        '200321 add by Brady
        'If gDatabase = "DAC" Then
        '    '190319 add by Brady
        '    'oCommand.CommandText = "select tc_bud03,nvl(round(sum(case when tc_bud14 = 'USD' then tc_bud13 * 6.3 else tc_bud13 * 7.56 end ),2),0) as t1 from tc_bud_file where tc_bud02 = 2018 and tc_bud01 = 1 group by tc_bud03 order by tc_bud03 "
        '    oCommand.CommandText = "select tc_bud03,nvl(round(sum(case when tc_bud14 = 'USD' then tc_bud13 * 6.85 else tc_bud13 * 8.22 end ),2),0) as t1 from tc_bud_file where tc_bud02 = 2019 and tc_bud01 = 1 group by tc_bud03 order by tc_bud03 "
        '    '190319 add by Brady end
        'Else
        '    oCommand.CommandText = "select tc_bud03,nvl(round(sum(case when tc_bud14 = 'EUR' then tc_bud13 * 1.2 else tc_bud13 end ),2),0) as t1 from tc_bud_file where tc_bud02 = 2019 and tc_bud01 = 1 group by tc_bud03 order by tc_bud03"
        'End If
        If gDatabase = "DAC" Then
            '190319 add by Brady
            'oCommand.CommandText = "select tc_bud03,nvl(round(sum(case when tc_bud14 = 'USD' then tc_bud13 * 6.3 else tc_bud13 * 7.56 end ),2),0) as t1 from tc_bud_file where tc_bud02 = 2018 and tc_bud01 = 1 group by tc_bud03 order by tc_bud03 "
            oCommand.CommandText = "select tc_bud03,nvl(round(sum(case when tc_bud14 = 'USD' then tc_bud13 * 6.85 else tc_bud13 * 8.22 end ),2),0) as t1 from tc_bud_file where tc_bud02 = " & tYear & " and tc_bud01 = 1 group by tc_bud03 order by tc_bud03 "
            '190319 add by Brady end
        Else
            oCommand.CommandText = "select tc_bud03,nvl(round(sum(case when tc_bud14 = 'EUR' then tc_bud13 * 1.2 else tc_bud13 end ),2),0) as t1 from tc_bud_file where tc_bud02 = " & tYear & " and tc_bud01 = 1 group by tc_bud03 order by tc_bud03"
        End If
        '200321 add by Brady END

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                'For i As Int16 = 1 To 12 Step 1
                Dim SSM As Decimal = oReader.Item(0)
                Ws.Cells(LineZ, SSM + 1) = oReader.Item(1)
                'Next
                'Dim tColumn As String = String.Empty
                'Select Case tMonth
                '    Case 1
                '        tColumn = "B"
                '    Case 2
                '        tColumn = "C"
                '    Case 3
                '        tColumn = "D"
                '    Case 4
                '        tColumn = "E"
                '    Case 5
                '        tColumn = "F"
                '    Case 6
                '        tColumn = "G"
                '    Case 7
                '        tColumn = "H"
                '    Case 8
                '        tColumn = "I"
                '    Case 9
                '        tColumn = "J"
                '    Case 10
                '        tColumn = "K"
                '    Case 11
                '        tColumn = "L"
                '    Case 12
                '        tColumn = "M"
                '    Case Else
                '        tColumn = "M"
                'End Select
                'Ws.Cells(LineZ, 14) = "=SUM(B" & LineZ & ":" & tColumn & LineZ & ")"
                'LineZ += 1
            End While
        End If
        oReader.Close()
        Ws.Cells(LineZ, 14) = "=SUM(B" & LineZ & ":M" & LineZ & ")"
        LineZ += 1

        Ws.Cells(LineZ, 2) = "=B" & LineZ - 2 & "-B" & LineZ - 1
        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 2))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, tMonth + 1)), Type:=xlFillDefault)
        Ws.Cells(LineZ, 14) = "=SUM(B" & LineZ & ":M" & LineZ & ")"
        LineZ += 1
        'Ws.Cells(LineZ, 2) = "=B" & LineZ - 1 & "/B" & LineZ - 2
        'Ws.Cells(LineZ, 2) = "=IF(B" & LineZ - 2 & "="""",0,B" & LineZ - 1 & "/B" & LineZ - 2 & ")"
        Ws.Cells(LineZ, 2) = "=IF(B" & LineZ - 1 & "=0,0,IF(B" & LineZ - 2 & "=0,1,B" & LineZ - 1 & "/B" & LineZ - 2 & "))"
        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 2))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 13)), Type:=xlFillDefault)
        ' 添加 負數為紅色 20180531
        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 14))
        oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
        oRng.FormatConditions(1).FONT.COLOR = Color.Red

        Ws.Cells(LineZ, 14) = "=N" & LineZ - 1 & "/N" & LineZ - 2
        LineZ += 1
        Ws.Cells(LineZ, 2) = "=B" & LineZ - 2
        Ws.Cells(LineZ, 3) = "=B" & LineZ & "+C" & LineZ - 2
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 13)), Type:=xlFillDefault)
        oRng = Ws.Range("B3", "O3")
        'oRng.EntireColumn.AutoFit()
        oRng = Ws.Range("A3", Ws.Cells(LineZ, 15))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        ' 第二頁

        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustExcelFormat3()
        oCommand.CommandText = "select distinct tqa02 from ( Select tqa02 from oga_file left join ogb_file on oga01 = ogb01 "
        oCommand.CommandText += "left join ima_file on ogb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand.CommandText += "where ogapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and (year(oga02) = " & pYear & " or ( year(oga02) = " & tYear & " and month(oga02) <= " & tMonth & ") ) "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "Select tqa02 from oha_file left join ohb_file on oha01 = ohb01  left join ima_file on ohb04 = ima01 "
        oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' where ohapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and (year(oha02) = " & pYear & " or ( year(oha02) = " & tYear & " and month(oha02) <= " & tMonth & ") ) ) order by tqa02"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                If Not String.IsNullOrEmpty(oReader.Item("tqa02").ToString()) Then
                    DOINPutData(oReader.Item("tqa02"), tYear, tMonth)
                    DOINPutDataA(oReader.Item("tqa02"), tYear, 12)
                    oCommand2.CommandText = "select nvl(gea02,'NA') from occ_file left join gea_file on occ20 = gea01 where occacti = 'Y' and occ02 = '" & oReader.Item("tqa02") & "'"
                    Dim C1 As String = oCommand2.ExecuteScalar()
                    AdjustExcelFormat4(C1)
                End If
            End While
        End If
        oReader.Close()

        ' ADD by cloud 20190424 增加資料, 和圖
        LineZ += 2
        AdjustExcelFormat11()
        LineZ += 1
        Ws.Cells(LineZ, 3) = "Europe total Budget"
        GeaDataBudget("Europe", tYear, 12)
        Ws.Cells(LineZ, 3) = "Europe total Sales"
        GeaDataSales("Europe", tYear, tMonth)
        Ws.Cells(LineZ, 3) = "Europe total Variance amount"
        Ws.Cells(LineZ + 1, 3) = "Europe total Variance %"
        AdjustExcelFormat12()
        Ws.Cells(LineZ, 3) = "USA/Japan total Budget"
        GeaDataBudget("USA/Japan", tYear, 12)
        Ws.Cells(LineZ, 3) = "USA/Japan total Sales"
        GeaDataSales("USA/Japan", tYear, tMonth)
        Ws.Cells(LineZ, 3) = "USA/Japan  total Variance amount"
        Ws.Cells(LineZ + 1, 3) = "USA/Japan total Variance %"
        AdjustExcelFormat12()
        ' 劃圖 - 1
        Dim XC As Microsoft.Office.Interop.Excel.Shape = Ws.Shapes.AddChart2(227, xlLineMarkers, 180, 1300, 900, 220)
        oRng = Ws.Range(Ws.Cells(LineZ - 9, 3), Ws.Cells(LineZ - 7, tMonth + 3))
        Dim XD As Chart = XC.Chart
        XD.SetSourceData(oRng)

        XD.SetElement(msoElementChartTitleAboveChart)
        XD.ChartTitle.Text = "YTD Europe Sales VS.Target"
        XD.SetElement(msoElementDataTableWithLegendKeys)
        XD.SetElement(msoElementLegendNone)

        ' 劃圖 - 2
        Dim XA As Microsoft.Office.Interop.Excel.Shape = Ws.Shapes.AddChart2(227, xlLineMarkers, 180, 1540, 900, 220)
        'oRng = Ws.Range(Ws.Range(Ws.Cells(LineZ - 9, 3), Ws.Cells(LineZ - 9, tMonth + 3)), Ws.Range(Ws.Cells(LineZ - 4, 3), Ws.Cells(LineZ - 3, tMonth + 3)))
        oRng = Ws.Range("C" & LineZ - 9 & ":G" & LineZ - 9 & ",C" & LineZ - 4 & ":G" & LineZ - 3)

        Dim XB As Chart = XA.Chart
        XB.SetSourceData(oRng)

        XB.SetElement(msoElementChartTitleAboveChart)
        XB.ChartTitle.Text = "YTD USA/Japan Sales VS.Target"
        XB.SetElement(msoElementDataTableWithLegendKeys)
        XB.SetElement(msoElementLegendNone)

        ' 第四頁 ->改第3頁

        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        AdjustExcelFormat7()
        oCommand.CommandText = "select distinct tqa02 from ( Select tqa02 from oga_file left join ogb_file on oga01 = ogb01 "
        oCommand.CommandText += "left join ima_file on ogb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand.CommandText += "where ogapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and (year(oga02) = " & pYear & " or ( year(oga02) = " & tYear & " and month(oga02) <= " & tMonth & ") ) "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "Select tqa02 from oha_file left join ohb_file on oha01 = ohb01  left join ima_file on ohb04 = ima01 "
        oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' where ohapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and (year(oha02) = " & pYear & " or ( year(oha02) = " & tYear & " and month(oha02) <= " & tMonth & ") ) ) order by tqa02"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                If Not String.IsNullOrEmpty(oReader.Item("tqa02").ToString()) Then
                    DOINPutData1(oReader.Item("tqa02"), tYear, tMonth)
                    DOINPutData1A(oReader.Item("tqa02"), tYear, 12)
                    oCommand2.CommandText = "select nvl(gea02,'NA') from occ_file left join gea_file on occ20 = gea01 where occacti = 'Y' and occ02 = '" & oReader.Item("tqa02") & "'"
                    Dim C1 As String = oCommand2.ExecuteScalar()
                    AdjustExcelFormat8(C1)
                End If
            End While
        End If
        oReader.Close()

        ' ADD by cloud 20190425 增加資料, 和圖
        LineZ += 2
        AdjustExcelFormat11()
        LineZ += 1
        Ws.Cells(LineZ, 3) = "Europe total Budget"
        GeaDataBudget1("Europe", tYear, 12)
        Ws.Cells(LineZ, 3) = "Europe total Sales"
        GeaDataSales1("Europe", tYear, tMonth)
        Ws.Cells(LineZ, 3) = "Europe total Variance amount"
        Ws.Cells(LineZ + 1, 3) = "Europe total Variance %"
        AdjustExcelFormat12()
        Ws.Cells(LineZ, 3) = "USA/Japan total Budget"
        GeaDataBudget1("USA/Japan", tYear, 12)
        Ws.Cells(LineZ, 3) = "USA/Japan total Sales"
        GeaDataSales1("USA/Japan", tYear, tMonth)
        Ws.Cells(LineZ, 3) = "USA/Japan  total Variance amount"
        Ws.Cells(LineZ + 1, 3) = "USA/Japan total Variance %"
        AdjustExcelFormat12()

        ' 劃圖 - 1
        Dim XE As Microsoft.Office.Interop.Excel.Shape = Ws.Shapes.AddChart2(227, xlLineMarkers, 180, 1300, 900, 220)
        oRng = Ws.Range(Ws.Cells(LineZ - 9, 3), Ws.Cells(LineZ - 7, tMonth + 3))
        Dim XF As Chart = XE.Chart
        XF.SetSourceData(oRng)

        XF.SetElement(msoElementChartTitleAboveChart)
        XF.ChartTitle.Text = "YTD Europe Sales VS.Target"
        XF.SetElement(msoElementDataTableWithLegendKeys)
        XF.SetElement(msoElementLegendNone)

        ' 劃圖 - 2
        Dim XG As Microsoft.Office.Interop.Excel.Shape = Ws.Shapes.AddChart2(227, xlLineMarkers, 180, 1540, 900, 220)
        'oRng = Ws.Range(Ws.Range(Ws.Cells(LineZ - 9, 3), Ws.Cells(LineZ - 9, tMonth + 3)), Ws.Range(Ws.Cells(LineZ - 4, 3), Ws.Cells(LineZ - 3, tMonth + 3)))
        oRng = Ws.Range("C" & LineZ - 9 & ":G" & LineZ - 9 & ",C" & LineZ - 4 & ":G" & LineZ - 3)

        Dim XH As Chart = XG.Chart
        XH.SetSourceData(oRng)

        XH.SetElement(msoElementChartTitleAboveChart)
        XH.ChartTitle.Text = "YTD USA/Japan Sales VS.Target"
        XH.SetElement(msoElementDataTableWithLegendKeys)
        XH.SetElement(msoElementLegendNone)


        ' 第三頁 改第四頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(4)
        Ws.Activate()
        AdjustExcelFormat5()
        oCommand.CommandText = "select tqa02,(t1-t2) as c1, (case when t2 <> 0 then round((t1-t2)/t2, 4) else 0 end) as c2 from ( "
        oCommand.CommandText += "select tqa02,sum(t1) as t1,sum(t2) as t2 from ( select tqa02,sum(ogb14 * oga24) as t1,0 as t2 from oga_file "
        oCommand.CommandText += "left join ogb_file on oga01 = ogb01 left join ima_file on ogb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand.CommandText += "where ogapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oga02) = " & tYear & " and month(oga02) <= " & tMonth & " group by tqa02 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select tc_bud05,0, "
        If gDatabase = "DAC" Then
            '190319 add by Brady
            'oCommand.CommandText += "sum(case when tc_bud14 = 'USD' then tc_bud13 * 6.3 else tc_bud13 * 7.56 end ) "
            oCommand.CommandText += "sum(case when tc_bud14 = 'USD' then tc_bud13 * 6.85 else tc_bud13 * 8.22 end ) "
            '190319 add by Brady end
        Else
            oCommand.CommandText += "sum(case when tc_bud14 = 'EUR' then tc_bud13 * 1.2 else tc_bud13 end ) "
        End If
        oCommand.CommandText += " from tc_bud_file where tc_bud01 = 1 and tc_bud02 = " & tYear & " and tc_bud03 <= " & tMonth
        oCommand.CommandText += "  group by tc_bud05 ) group by tqa02  ) order by tqa02"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                For i As Int16 = 1 To 3 Step 1
                    Ws.Cells(LineZ, i) = oReader.Item(i - 1)
                Next
                LineZ += 1
            End While
            ' 上色
            oRng = Ws.Range("A4", Ws.Cells(LineZ - 1, 1))
            oRng.Interior.Color = Color.FromArgb(250, 191, 143)
            ' 格式
            oRng = Ws.Range("B4", Ws.Cells(LineZ - 1, 2))
            oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
            oRng = Ws.Range("C4", Ws.Cells(LineZ - 1, 3))
            oRng.NumberFormatLocal = "0%"
            oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
            oRng.FormatConditions(1).FONT.COLOR = Color.Red
            '劃線
            oRng = Ws.Range("A4", Ws.Cells(LineZ - 1, 3))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        oReader.Close()

        AdjustExcelFormat6()
        ' 定錨
        LineS1 = LineZ
        oCommand.CommandText = "select tqa02,(t1-t13) as c1,(case when t13 <> 0 then round((t1-t13)/t13,4) else 0 end) as c2,"
        oCommand.CommandText += "(t2-t14) as c3,(case when t14 <> 0 then round((t2-t14)/t14,4) else 0 end) as c4,(t3-t15) as c5,(case when t15 <> 0 then round((t3-t15)/t15,4) else 0 end) as c6,"
        oCommand.CommandText += "(t4-t16) as c7,(case when t16 <> 0 then round((t4-t16)/t16,4) else 0 end) as c8,(t5-t17) as c9,(case when t17 <> 0 then round((t5-t17)/t17,4) else 0 end) as c10,"
        oCommand.CommandText += "(t6-t18) as c11,(case when t18 <> 0 then round((t6-t18)/t18,4) else 0 end) as c12,(t7-t19) as c13,(case when t19 <> 0 then round((t7-t19)/t19,4) else 0 end) as c14,"
        oCommand.CommandText += "(t8-t20) as c15,(case when t20 <> 0 then round((t8-t20)/t20,4) else 0 end) as c16,(t9-t21) as c17,(case when t21 <> 0 then round((t9-t21)/t21,4) else 0 end) as c18,"
        oCommand.CommandText += "(t10-t22) as c19,(case when t22 <> 0 then round((t10-t22)/t22,4) else 0 end) as c20,(t11-t23) as c21,(case when t23 <> 0 then round((t11-t23)/t23,4) else 0 end) as c22,"
        oCommand.CommandText += "(t12-t24) as c23,(case when t24 <> 0 then round((t12-t24)/t24,4) else 0 end) as c24 from ( select tqa02,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,"
        oCommand.CommandText += "sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12,sum(t13) as t13,sum(t14) as t14,sum(t15) as t15,"
        oCommand.CommandText += "sum(t16) as t16,sum(t17) as t17,sum(t18) as t18,sum(t19) as t19,sum(t20) as t20,sum(t21) as t21,sum(t22) as t22,sum(t23) as t23,sum(t24) as t24 from ( "
        oCommand.CommandText += "select tqa02,(case when month(oga02) = 1 then ogb14 * oga24 else 0 end ) as t1,(case when month(oga02) = 2 then ogb14 * oga24 else 0 end ) as t2,"
        oCommand.CommandText += "(case when month(oga02) = 3 then ogb14 * oga24 else 0 end ) as t3,(case when month(oga02) = 4 then ogb14 * oga24 else 0 end ) as t4,"
        oCommand.CommandText += "(case when month(oga02) = 5 then ogb14 * oga24 else 0 end ) as t5,(case when month(oga02) = 6 then ogb14 * oga24 else 0 end ) as t6,"
        oCommand.CommandText += "(case when month(oga02) = 7 then ogb14 * oga24 else 0 end ) as t7,(case when month(oga02) = 8 then ogb14 * oga24 else 0 end ) as t8,"
        oCommand.CommandText += "(case when month(oga02) = 9 then ogb14 * oga24 else 0 end ) as t9,(case when month(oga02) = 10 then ogb14 * oga24 else 0 end ) as t10,"
        oCommand.CommandText += "(case when month(oga02) = 11 then ogb14 * oga24 else 0 end ) as t11,(case when month(oga02) = 12 then ogb14 * oga24 else 0 end ) as t12,0 as t13,0 as t14,0 as t15,0 as t16,0 as t17,0 as t18,0 as t19,0 as t20,0 as t21,0 as t22,0 as t23,0 as t24 from oga_file "
        oCommand.CommandText += "left join ogb_file on oga01 = ogb01 left join ima_file on ogb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand.CommandText += "where ogapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oga02) = " & tYear & " and month(oga02) <= " & tMonth & " "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select tc_bud05,0,0,0,0,0,0,0,0,0,0,0,0,(case when tc_bud03 = 1 then d1 else 0 end) ,(case when tc_bud03 = 2 then d1 else 0 end)"
        oCommand.CommandText += ",(case when tc_bud03 = 3 then d1 else 0 end),(case when tc_bud03 = 4 then d1 else 0 end),(case when tc_bud03 = 5 then d1 else 0 end)"
        oCommand.CommandText += ",(case when tc_bud03 = 6 then d1 else 0 end),(case when tc_bud03 = 7 then d1 else 0 end),(case when tc_bud03 = 8 then d1 else 0 end)"
        oCommand.CommandText += ",(case when tc_bud03 = 9 then d1 else 0 end),(case when tc_bud03 = 10 then d1 else 0 end),(case when tc_bud03 = 11 then d1 else 0 end)"
        oCommand.CommandText += ",(case when tc_bud03 = 12 then d1 else 0 end) from ( select tc_bud05,tc_bud03,(case when tc_bud14 = 'USD' then "
        If gDatabase = "DAC" Then
            '190319 add by Brady
            'oCommand.CommandText += "tc_bud13 * 6.3 else tc_bud13 * 7.56"
            oCommand.CommandText += "tc_bud13 * 6.85 else tc_bud13 * 8.22"
            '190319 add by Brady end
        Else
            oCommand.CommandText += "tc_bud13 else tc_bud13 * 1.2"
        End If
        oCommand.CommandText += " end ) as d1 from tc_bud_file where tc_bud01 =1 and tc_bud02 = " & tYear & " and tc_bud03 <= " & tMonth
        oCommand.CommandText += ") ) group by tqa02 ) order by tqa02"
        oReader = oCommand.ExecuteReader()

        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 1 To (2 * tMonth) + 1
                    Ws.Cells(LineZ, i) = oReader.Item(i - 1)
                Next
                LineZ += 1
            End While
            ' 上色
            oRng = Ws.Range(Ws.Cells(LineS1, 1), Ws.Cells(LineZ - 1, 1))
            oRng.Interior.Color = Color.FromArgb(250, 191, 143)
            oRng = Ws.Range(Ws.Cells(LineS1, 2), Ws.Cells(LineZ - 1, 3))
            oRng.Interior.Color = Color.LightGreen
            oRng = Ws.Range(Ws.Cells(LineS1, 4), Ws.Cells(LineZ - 1, 5))
            oRng.Interior.Color = Color.FromArgb(250, 191, 143)
            oRng = Ws.Range(Ws.Cells(LineS1, 6), Ws.Cells(LineZ - 1, 7))
            oRng.Interior.Color = Color.LightGreen
            oRng = Ws.Range(Ws.Cells(LineS1, 8), Ws.Cells(LineZ - 1, 9))
            oRng.Interior.Color = Color.FromArgb(250, 191, 143)
            oRng = Ws.Range(Ws.Cells(LineS1, 10), Ws.Cells(LineZ - 1, 11))
            oRng.Interior.Color = Color.LightGreen
            oRng = Ws.Range(Ws.Cells(LineS1, 12), Ws.Cells(LineZ - 1, 13))
            oRng.Interior.Color = Color.FromArgb(250, 191, 143)
            oRng = Ws.Range(Ws.Cells(LineS1, 14), Ws.Cells(LineZ - 1, 15))
            oRng.Interior.Color = Color.LightGreen
            oRng = Ws.Range(Ws.Cells(LineS1, 16), Ws.Cells(LineZ - 1, 17))
            oRng.Interior.Color = Color.FromArgb(250, 191, 143)
            oRng = Ws.Range(Ws.Cells(LineS1, 18), Ws.Cells(LineZ - 1, 19))
            oRng.Interior.Color = Color.LightGreen
            oRng = Ws.Range(Ws.Cells(LineS1, 20), Ws.Cells(LineZ - 1, 21))
            oRng.Interior.Color = Color.FromArgb(250, 191, 143)
            oRng = Ws.Range(Ws.Cells(LineS1, 22), Ws.Cells(LineZ - 1, 23))
            oRng.Interior.Color = Color.LightGreen
            oRng = Ws.Range(Ws.Cells(LineS1, 24), Ws.Cells(LineZ - 1, 25))
            oRng.Interior.Color = Color.FromArgb(250, 191, 143)
            ' 格式
            oRng = Ws.Range(Ws.Cells(LineS1, 2), Ws.Cells(LineZ - 1, 2))
            oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
            oRng = Ws.Range(Ws.Cells(LineS1, 3), Ws.Cells(LineZ - 1, 3))
            oRng.NumberFormatLocal = "0%"
            ' 添加 負數為紅色 20180531
            oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
            oRng.FormatConditions(1).FONT.COLOR = Color.Red
            oRng = Ws.Range(Ws.Cells(LineS1, 4), Ws.Cells(LineZ - 1, 4))
            oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
            oRng = Ws.Range(Ws.Cells(LineS1, 5), Ws.Cells(LineZ - 1, 5))
            oRng.NumberFormatLocal = "0%"
            oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
            oRng.FormatConditions(1).FONT.COLOR = Color.Red
            oRng = Ws.Range(Ws.Cells(LineS1, 6), Ws.Cells(LineZ - 1, 6))
            oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
            oRng = Ws.Range(Ws.Cells(LineS1, 7), Ws.Cells(LineZ - 1, 7))
            oRng.NumberFormatLocal = "0%"
            oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
            oRng.FormatConditions(1).FONT.COLOR = Color.Red
            oRng = Ws.Range(Ws.Cells(LineS1, 8), Ws.Cells(LineZ - 1, 8))
            oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
            oRng = Ws.Range(Ws.Cells(LineS1, 9), Ws.Cells(LineZ - 1, 9))
            oRng.NumberFormatLocal = "0%"
            oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
            oRng.FormatConditions(1).FONT.COLOR = Color.Red
            oRng = Ws.Range(Ws.Cells(LineS1, 10), Ws.Cells(LineZ - 1, 10))
            oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
            oRng = Ws.Range(Ws.Cells(LineS1, 11), Ws.Cells(LineZ - 1, 11))
            oRng.NumberFormatLocal = "0%"
            oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
            oRng.FormatConditions(1).FONT.COLOR = Color.Red
            oRng = Ws.Range(Ws.Cells(LineS1, 12), Ws.Cells(LineZ - 1, 12))
            oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
            oRng = Ws.Range(Ws.Cells(LineS1, 13), Ws.Cells(LineZ - 1, 13))
            oRng.NumberFormatLocal = "0%"
            oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
            oRng.FormatConditions(1).FONT.COLOR = Color.Red
            oRng = Ws.Range(Ws.Cells(LineS1, 14), Ws.Cells(LineZ - 1, 14))
            oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
            oRng = Ws.Range(Ws.Cells(LineS1, 15), Ws.Cells(LineZ - 1, 15))
            oRng.NumberFormatLocal = "0%"
            oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
            oRng.FormatConditions(1).FONT.COLOR = Color.Red
            oRng = Ws.Range(Ws.Cells(LineS1, 16), Ws.Cells(LineZ - 1, 16))
            oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
            oRng = Ws.Range(Ws.Cells(LineS1, 17), Ws.Cells(LineZ - 1, 17))
            oRng.NumberFormatLocal = "0%"
            oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
            oRng.FormatConditions(1).FONT.COLOR = Color.Red
            oRng = Ws.Range(Ws.Cells(LineS1, 18), Ws.Cells(LineZ - 1, 18))
            oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
            oRng = Ws.Range(Ws.Cells(LineS1, 19), Ws.Cells(LineZ - 1, 19))
            oRng.NumberFormatLocal = "0%"
            oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
            oRng.FormatConditions(1).FONT.COLOR = Color.Red
            oRng = Ws.Range(Ws.Cells(LineS1, 20), Ws.Cells(LineZ - 1, 20))
            oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
            oRng = Ws.Range(Ws.Cells(LineS1, 21), Ws.Cells(LineZ - 1, 21))
            oRng.NumberFormatLocal = "0%"
            oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
            oRng.FormatConditions(1).FONT.COLOR = Color.Red
            oRng = Ws.Range(Ws.Cells(LineS1, 22), Ws.Cells(LineZ - 1, 22))
            oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
            oRng = Ws.Range(Ws.Cells(LineS1, 23), Ws.Cells(LineZ - 1, 23))
            oRng.NumberFormatLocal = "0%"
            oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
            oRng.FormatConditions(1).FONT.COLOR = Color.Red
            oRng = Ws.Range(Ws.Cells(LineS1, 24), Ws.Cells(LineZ - 1, 24))
            oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
            oRng = Ws.Range(Ws.Cells(LineS1, 25), Ws.Cells(LineZ - 1, 25))
            oRng.NumberFormatLocal = "0%"
            oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
            oRng.FormatConditions(1).FONT.COLOR = Color.Red
            '劃線
            oRng = Ws.Range(Ws.Cells(LineS1, 1), Ws.Cells(LineZ - 1, 25))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        oReader.Close()



        ' 明細, 很多頁
        oCommand.CommandText = "select distinct tqa02 from ( Select tqa02 from oga_file left join ogb_file on oga01 = ogb01 "
        oCommand.CommandText += "left join ima_file on ogb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand.CommandText += "where ogapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and (year(oga02) = " & pYear & " or (year(oga02) >= " & pYear & " and month(oga02) <= " & tMonth & ") ) "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "Select tqa02 from oha_file left join ohb_file on oha01 = ohb01  left join ima_file on ohb04 = ima01 "
        oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' where ohapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and (year(oha02) = " & pYear & " or ( year(oha02) >= " & pYear & " and month(oha02) <= " & tMonth & ") ) ) order by tqa02"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                If Not String.IsNullOrEmpty(oReader.Item("tqa02").ToString()) Then
                    Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                    Ws = xWorkBook.Sheets(xWorkBook.Sheets.Count)
                    Ws.Activate()
                    AdjustExcelFormat9(oReader.Item("tqa02"))
                    DOINPutData2(oReader.Item("tqa02"), tYear, tMonth)
                    LineZ += 2
                    AdjustExcelFormat10(oReader.Item("tqa02"))
                    DOINPutData2A(oReader.Item("tqa02"), tYear, tMonth)

                End If
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = gDatabase & "_Sales compare with budget_Summary.xlsx"
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
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "Summary"
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 11
        Ws.Rows.RowHeight = 14
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 32.5
        oRng.EntireColumn.Font.Bold = True
        If gDatabase = "DAC" Then
            Ws.Cells(1, 1) = "Company Name：Dongguan Action Composites LTD Co."
        Else
            Ws.Cells(1, 1) = "Company Name：ACTION COMPOSITE TECHNOLOGY LIMITED"
        End If

        Ws.Cells(2, 1) = "Revenue: month-on-month basis"
        Ws.Cells(2, 2) = "Currency：" & gCurrency
        Ws.Cells(3, 1) = "Customer"
        For i As Int16 = 1 To 12 Step 1
            If i < 10 Then
                Ws.Cells(3, i + 1) = tYear & "/0" & i
            Else
                Ws.Cells(3, i + 1) = tYear & "/" & i
            End If
        Next
        oRng = Ws.Range("B3", "O3")
        oRng.HorizontalAlignment = xlcenter
        oRng.VerticalAlignment = xlBottom
        oRng.EntireColumn.ColumnWidth = 15
        Ws.Cells(3, 14) = "YTD " & tYear
        Ws.Cells(3, 15) = "% by customer"

        oRng = Ws.Range("A1", "A2")
        oRng.EntireRow.Font.Bold = True
        oRng = Ws.Range("N3", "O3")
        oRng.Interior.Color = Color.Yellow

        LineZ = 4
    End Sub
    Private Sub AdjustExcelFormat2()
        Ws.Cells(LineZ, 1) = "Revenue of Y" & tYear & "（Actual）"
        Ws.Cells(LineZ + 1, 1) = "Revenue of Y" & tYear & "（Budget）"
        Ws.Cells(LineZ + 2, 1) = "Variance amount"
        Ws.Cells(LineZ + 3, 1) = "Variance %"
        Ws.Cells(LineZ + 4, 1) = "YTD Variance amount"
        Ws.Cells(LineZ, 2) = "=SUM(B4:B" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 2))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 13)), Type:=xlFillDefault)
        Ws.Cells(LineZ, 14) = "=SUM(B" & LineZ & ":M" & LineZ & ")"
        Ws.Cells(LineZ, 15) = "=N" & LineZ & "/$N$" & LineZ
        For i As Int16 = 4 To LineZ - 1 Step 1
            Ws.Cells(i, 15) = "=N" & i & "/$N$" & LineZ
        Next
        ' 上色
        oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 13))
        oRng.Interior.Color = Color.FromArgb(146, 205, 220)
        oRng = Ws.Range(Ws.Cells(LineZ, 14), Ws.Cells(LineZ + 4, 15))
        oRng.Interior.Color = Color.Yellow
        oRng = Ws.Range(Ws.Cells(LineZ + 1, 1), Ws.Cells(LineZ + 1, 13))
        oRng.Interior.Color = Color.Yellow
        oRng = Ws.Range(Ws.Cells(LineZ + 1, 1), Ws.Cells(LineZ + 1, 1))
        oRng.Font.Color = Color.Red
        'oRng = Ws.Range(Ws.Cells(LineZ + 4, 1), Ws.Cells(LineZ + 4, 13))
        'oRng.Font.Color = Color.Red
        '格式
        oRng = Ws.Range("B3", Ws.Cells(LineZ + 2, 14))
        oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        oRng = Ws.Range(Ws.Cells(LineZ + 4, 2), Ws.Cells(LineZ + 4, 14))
        oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        oRng = Ws.Range(Ws.Cells(LineZ + 3, 2), Ws.Cells(LineZ + 3, 14))
        oRng.NumberFormatLocal = "0%"
        oRng = Ws.Range("O3", Ws.Cells(LineZ, 15))
        oRng.NumberFormatLocal = "0%"
    End Sub
    Private Sub AdjustExcelFormat3()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "Rev. Variance % by customer"
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 11
        Ws.Rows.RowHeight = 14
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 17.75
        oRng.EntireColumn.Font.Bold = True
        If gDatabase = "DAC" Then
            Ws.Cells(1, 1) = "Company Name：Dongguan Action Composites LTD Co."
        Else
            Ws.Cells(1, 1) = "Company Name：ACTION COMPOSITE TECHNOLOGY LIMITED"
        End If
        Ws.Cells(2, 1) = "Currency：" & gCurrency
        Ws.Cells(3, 1) = "Classification"
        Ws.Cells(3, 2) = "Customer"
        Ws.Cells(3, 3) = "Year"
        For i As Int16 = 1 To 12 Step 1
            'If i < 10 Then
            'Ws.Cells(3, i + 2) = tYear & "/0" & i
            'Else
            Dim TempString As String = String.Empty
            Select Case i
                Case 1
                    TempString = "Jan."
                Case 2
                    TempString = "Feb."
                Case 3
                    TempString = "Mar."
                Case 4
                    TempString = "Apr."
                Case 5
                    TempString = "May."
                Case 6
                    TempString = "Jun."
                Case 7
                    TempString = "Jul."
                Case 8
                    TempString = "Aug."
                Case 9
                    TempString = "Sept."
                Case 10
                    TempString = "Oct."
                Case 11
                    TempString = "Nov."
                Case 12
                    TempString = "Dec."
            End Select
            Ws.Cells(3, i + 3) = TempString

            'End If
        Next
        Ws.Cells(3, 16) = "YTD"
        oRng = Ws.Range("C3", "P3")
        oRng.HorizontalAlignment = xlCenter
        oRng.VerticalAlignment = xlBottom
        oRng = Ws.Range("B3", "C3")
        oRng.EntireColumn.ColumnWidth = 8.38
        oRng = Ws.Range("D3", "P3")
        oRng.EntireColumn.ColumnWidth = 15


        oRng = Ws.Range("B2", "B2")
        oRng.EntireRow.Font.Bold = True
        oRng = Ws.Range("P3", "P3")
        oRng.Interior.Color = Color.FromArgb(250, 191, 143)

        '劃線
        oRng = Ws.Range("A3", "P3")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        LineZ = 4
    End Sub
    Private Sub DOINPutData(ByVal tqa02 As String, ByVal sYear As Decimal, ByVal iTerm As Decimal)
        oCommand2.CommandText = "select nvl(gea02,'NA') from occ_file left join gea_file on occ20 = gea01 where occacti = 'Y' and occ02 = '" & tqa02 & "'"
        Dim C1 As String = oCommand2.ExecuteScalar()

        oCommand2.CommandText = "select tqa02,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,"
        oCommand2.CommandText += "sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( "
        oCommand2.CommandText += "select tqa02,(case when month(oga02) = 1 then ogb14 * oga24 else 0 end ) as t1,(case when month(oga02) = 2 then ogb14 * oga24 else 0 end ) as t2,"
        oCommand2.CommandText += "(case when month(oga02) = 3 then ogb14 * oga24 else 0 end ) as t3,(case when month(oga02) = 4 then ogb14 * oga24 else 0 end ) as t4,"
        oCommand2.CommandText += "(case when month(oga02) = 5 then ogb14 * oga24 else 0 end ) as t5,(case when month(oga02) = 6 then ogb14 * oga24 else 0 end ) as t6,"
        oCommand2.CommandText += "(case when month(oga02) = 7 then ogb14 * oga24 else 0 end ) as t7,(case when month(oga02) = 8 then ogb14 * oga24 else 0 end ) as t8,"
        oCommand2.CommandText += "(case when month(oga02) = 9 then ogb14 * oga24 else 0 end ) as t9,(case when month(oga02) = 10 then ogb14 * oga24 else 0 end ) as t10,"
        oCommand2.CommandText += "(case when month(oga02) = 11 then ogb14 * oga24 else 0 end ) as t11,(case when month(oga02) = 12 then ogb14 * oga24 else 0 end ) as t12 from oga_file "
        oCommand2.CommandText += "left join ogb_file on oga01 = ogb01 left join ima_file on ogb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand2.CommandText += "where ogapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and tqa02 = '" & tqa02 & "' and year(oga02) = " & sYear & " and month(oga02) <= " & iTerm & " "
        oCommand2.CommandText += "union all "
        oCommand2.CommandText += "select tqa02,(case when month(oha02) = 1 then ohb14 * oha24 * -1 else 0 end ) as t1,"
        oCommand2.CommandText += "(case when month(oha02) = 2 then ohb14 * oha24 * -1 else 0 end ) as t2,(case when month(oha02) = 3 then ohb14 * oha24 * -1 else 0 end ) as t3,"
        oCommand2.CommandText += "(case when month(oha02) = 4 then ohb14 * oha24 * -1 else 0 end ) as t4,(case when month(oha02) = 5 then ohb14 * oha24 * -1 else 0 end ) as t5,"
        oCommand2.CommandText += "(case when month(oha02) = 6 then ohb14 * oha24 * -1 else 0 end ) as t6,(case when month(oha02) = 7 then ohb14 * oha24 * -1 else 0 end ) as t7,"
        oCommand2.CommandText += "(case when month(oha02) = 8 then ohb14 * oha24 * -1 else 0 end ) as t8,(case when month(oha02) = 9 then ohb14 * oha24 * -1 else 0 end ) as t9,"
        oCommand2.CommandText += "(case when month(oha02) = 10 then ohb14 * oha24 * -1 else 0 end ) as t10,(case when month(oha02) = 11 then ohb14 * oha24 * -1 else 0 end ) as t11,"
        oCommand2.CommandText += "(case when month(oha02) = 12 then ohb14 * oha24 * -1 else 0 end ) as t12  from oha_file left join ohb_file on oha01 = ohb01  left join ima_file on ohb04 = ima01 "
        oCommand2.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' where ohapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and tqa02 = '" & tqa02 & "' and year(oha02) = " & sYear & " and month(oha02) <= " & iTerm & " ) group by tqa02 order by tqa02"
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                Ws.Cells(LineZ, 1) = C1
                Ws.Cells(LineZ, 2) = oReader2.Item("tqa02")
                'Ws.Cells(LineZ, 2) = sYear
                Ws.Cells(LineZ, 3) = "Actual"
                oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
                oRng.HorizontalAlignment = xlCenter
                oRng.VerticalAlignment = xlBottom
                For i As Int16 = 1 To iTerm Step 1
                    Ws.Cells(LineZ, i + 3) = oReader2.Item(i)
                Next
                Dim tColumn As String = String.Empty
                Select Case tMonth
                    Case 1
                        tColumn = "D"
                    Case 2
                        tColumn = "E"
                    Case 3
                        tColumn = "F"
                    Case 4
                        tColumn = "G"
                    Case 5
                        tColumn = "H"
                    Case 6
                        tColumn = "I"
                    Case 7
                        tColumn = "J"
                    Case 8
                        tColumn = "K"
                    Case 9
                        tColumn = "L"
                    Case 10
                        tColumn = "M"
                    Case 11
                        tColumn = "N"
                    Case 12
                        tColumn = "O"
                    Case Else
                        tColumn = "O"
                End Select
                Ws.Cells(LineZ, 16) = "=SUM(D" & LineZ & ":" & tColumn & LineZ & ")"
                LineZ += 1
            End While
        Else
            Ws.Cells(LineZ, 1) = C1
            Ws.Cells(LineZ, 2) = tqa02
            'Ws.Cells(LineZ, 2) = sYear
            Ws.Cells(LineZ, 3) = "Actual"
            oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
            oRng.HorizontalAlignment = xlCenter
            oRng.VerticalAlignment = xlBottom
            For i As Int16 = 1 To iTerm Step 1
                Ws.Cells(LineZ, i + 3) = 0
            Next

            LineZ += 1
        End If
        oReader2.Close()
    End Sub
    Private Sub DOINPutDataA(ByVal tqa02 As String, ByVal sYear As Decimal, ByVal iTerm As Decimal)
        oCommand2.CommandText = "select nvl(gea02,'NA') from occ_file left join gea_file on occ20 = gea01 where occacti = 'Y' and occ02 = '" & tqa02 & "'"
        Dim C1 As String = oCommand2.ExecuteScalar()

        If gDatabase = "DAC" Then
            oCommand2.CommandText = "select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,"
            oCommand2.CommandText += "sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( "
            oCommand2.CommandText += "select (case when tc_bud03 = 1 then t1 else 0 end ) as t1,(case when tc_bud03 = 2 then t1 else 0 end ) as t2,"
            oCommand2.CommandText += "(case when tc_bud03 = 3 then t1 else 0 end ) as t3,(case when tc_bud03 = 4 then t1 else 0 end ) as t4,"
            oCommand2.CommandText += "(case when tc_bud03 = 5 then t1 else 0 end ) as t5,(case when tc_bud03 = 6 then t1 else 0 end ) as t6,"
            oCommand2.CommandText += "(case when tc_bud03 = 7 then t1 else 0 end ) as t7,(case when tc_bud03 = 8 then t1 else 0 end ) as t8,"
            oCommand2.CommandText += "(case when tc_bud03 = 9 then t1 else 0 end ) as t9,(case when tc_bud03 = 10 then t1 else 0 end ) as t10,"
            oCommand2.CommandText += "(case when tc_bud03 = 11 then t1 else 0 end ) as t11,(case when tc_bud03 = 12 then t1 else 0 end ) as t12 "
            '190319 add by Brady
            'oCommand2.CommandText += "from ( select tc_bud03,(case when tc_bud14 = 'USD' then tc_bud13 * 6.3 else tc_bud13 * 7.56 end ) as t1 "
            oCommand2.CommandText += "from ( select tc_bud03,(case when tc_bud14 = 'USD' then tc_bud13 * 6.85 else tc_bud13 * 8.22 end ) as t1 "
            '190319 add by Brady END
            oCommand2.CommandText += "from tc_bud_file where tc_bud02 = " & sYear & " and tc_bud01 = 1 AND tc_bud05 = '" & tqa02 & "' )  )"
        Else
            oCommand2.CommandText = "select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,"
            oCommand2.CommandText += "sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( "
            oCommand2.CommandText += "select (case when tc_bud03 = 1 then t1 else 0 end ) as t1,(case when tc_bud03 = 2 then t1 else 0 end ) as t2,"
            oCommand2.CommandText += "(case when tc_bud03 = 3 then t1 else 0 end ) as t3,(case when tc_bud03 = 4 then t1 else 0 end ) as t4,"
            oCommand2.CommandText += "(case when tc_bud03 = 5 then t1 else 0 end ) as t5,(case when tc_bud03 = 6 then t1 else 0 end ) as t6,"
            oCommand2.CommandText += "(case when tc_bud03 = 7 then t1 else 0 end ) as t7,(case when tc_bud03 = 8 then t1 else 0 end ) as t8,"
            oCommand2.CommandText += "(case when tc_bud03 = 9 then t1 else 0 end ) as t9,(case when tc_bud03 = 10 then t1 else 0 end ) as t10,"
            oCommand2.CommandText += "(case when tc_bud03 = 11 then t1 else 0 end ) as t11,(case when tc_bud03 = 12 then t1 else 0 end ) as t12 "
            oCommand2.CommandText += "from ( select tc_bud03,(case when tc_bud14 = 'EUR' then tc_bud13 * 1.2 else tc_bud13 end ) as t1 "
            oCommand2.CommandText += "from tc_bud_file where tc_bud02 = " & sYear & " and tc_bud01 = 1 AND tc_bud05 = '" & tqa02 & "' )  )"
        End If

        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                Ws.Cells(LineZ, 1) = C1
                'Ws.Cells(LineZ, 1) = oReader2.Item("tqa02")
                Ws.Cells(LineZ, 2) = tqa02
                'Ws.Cells(LineZ, 2) = sYear
                Ws.Cells(LineZ, 3) = "Budget"
                oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
                oRng.HorizontalAlignment = xlCenter
                oRng.VerticalAlignment = xlBottom
                For i As Int16 = 1 To iTerm Step 1
                    Ws.Cells(LineZ, i + 3) = oReader2.Item(i - 1)
                Next
                Dim tColumn As String = String.Empty
                Select Case tMonth
                    Case 1
                        tColumn = "D"
                    Case 2
                        tColumn = "E"
                    Case 3
                        tColumn = "F"
                    Case 4
                        tColumn = "G"
                    Case 5
                        tColumn = "H"
                    Case 6
                        tColumn = "I"
                    Case 7
                        tColumn = "J"
                    Case 8
                        tColumn = "K"
                    Case 9
                        tColumn = "L"
                    Case 10
                        tColumn = "M"
                    Case 11
                        tColumn = "N"
                    Case 12
                        tColumn = "O"
                    Case Else
                        tColumn = "O"
                End Select
                Ws.Cells(LineZ, 16) = "=SUM(D" & LineZ & ":" & tColumn & LineZ & ")"
                LineZ += 1
            End While
        Else
            Ws.Cells(LineZ, 1) = C1
            Ws.Cells(LineZ, 2) = tqa02
            'Ws.Cells(LineZ, 2) = sYear
            Ws.Cells(LineZ, 3) = "Budget"
            oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
            oRng.HorizontalAlignment = xlCenter
            oRng.VerticalAlignment = xlBottom
            For i As Int16 = 1 To iTerm Step 1
                Ws.Cells(LineZ, i + 3) = 0
            Next

            LineZ += 1
        End If
        oReader2.Close()
    End Sub
    Private Sub AdjustExcelFormat4(ByVal Capt As String)
        Ws.Cells(LineZ, 1) = Capt
        Ws.Cells(LineZ + 1, 1) = Capt
        Ws.Cells(LineZ, 2) = "Variance amount"
        Ws.Cells(LineZ + 1, 2) = "Variance %"
        Ws.Cells(LineZ, 4) = "=D" & LineZ - 2 & "-D" & LineZ - 1
        oRng = Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 4))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, tMonth + 3)), Type:=xlFillDefault)
        Ws.Cells(LineZ, 16) = "=SUM(D" & LineZ & ":O" & LineZ & ")"
        'Ws.Cells(LineZ + 1, 3) = "=C" & LineZ & "/C" & LineZ - 1
        'Ws.Cells(LineZ + 1, 3) = "=IF(C" & LineZ - 1 & "="""",0,C" & LineZ & "/C" & LineZ - 1 & ")"
        Ws.Cells(LineZ + 1, 4) = "=IF(D" & LineZ & "=0,0,IF(D" & LineZ - 1 & "=0,1,D" & LineZ & "/D" & LineZ - 1 & "))"
        oRng = Ws.Range(Ws.Cells(LineZ + 1, 4), Ws.Cells(LineZ + 1, 4))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ + 1, 4), Ws.Cells(LineZ + 1, tMonth + 3)), Type:=xlFillDefault)
        Ws.Cells(LineZ + 1, 16) = "=IF(P" & LineZ & "=0,0,IF(P" & LineZ - 1 & "=0,1,P" & LineZ & "/P" & LineZ - 1 & "))"
        ' 上色
        oRng = Ws.Range(Ws.Cells(LineZ - 2, 16), Ws.Cells(LineZ + 1, 16))
        oRng.Interior.Color = Color.FromArgb(250, 191, 143)
        'oRng = Ws.Range(Ws.Cells(LineZ + 2, 1), Ws.Cells(LineZ + 2, 16))
        'oRng.Merge()
        'oRng.Interior.Color = Color.Yellow

        '格式
        oRng = Ws.Range(Ws.Cells(LineZ - 2, 4), Ws.Cells(LineZ, 16))
        oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        oRng = Ws.Range(Ws.Cells(LineZ + 1, 4), Ws.Cells(LineZ + 1, 16))
        oRng.NumberFormatLocal = "0%"
        ' 添加 負數為紅色 20180531
        oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
        oRng.FormatConditions(1).FONT.COLOR = Color.Red

        '劃線
        oRng = Ws.Range(Ws.Cells(LineZ - 2, 1), Ws.Cells(LineZ + 1, 16))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        LineZ += 2
    End Sub
    Private Sub AdjustExcelFormat5()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "turnover by customer"
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 11
        Ws.Rows.RowHeight = 14
        oRng = Ws.Range("A3", "C3")
        oRng.EntireColumn.ColumnWidth = 19.25
        oRng.EntireRow.Font.Bold = True
        oRng.Interior.Color = Color.LightGreen
        oRng = Ws.Range("B3", "C3")
        oRng.HorizontalAlignment = xlCenter
        oRng.VerticalAlignment = xlBottom

        oRng = Ws.Range("B3", "Y3")
        oRng.EntireColumn.ColumnWidth = 15

        oRng = Ws.Range("B3", "C3")
        oRng.EntireRow.RowHeight = 41.4
        oRng.WrapText = True

        If gDatabase = "DAC" Then
            Ws.Cells(1, 1) = "Company Name：Dongguan Action Composites LTD Co."
        Else
            Ws.Cells(1, 1) = "Company Name：ACTION COMPOSITE TECHNOLOGY LIMITED"
        End If
        Ws.Cells(2, 1) = "Currency: " & gCurrency
        Ws.Cells(3, 1) = "Customer Variance amount"
        Ws.Cells(3, 2) = "YTD " & tYear & " Variance amount"
        Ws.Cells(3, 3) = "YTD " & tYear & " Variance %"

        '劃線
        oRng = Ws.Range("A3", "C3")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        LineZ = 4
    End Sub
    Private Sub AdjustExcelFormat6()
        LineZ += 1
        oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 1))
        oRng.Interior.Color = Color.Yellow
        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 3))
        oRng.Interior.Color = Color.LightGreen
        oRng = Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 5))
        oRng.Interior.Color = Color.Yellow
        oRng = Ws.Range(Ws.Cells(LineZ, 6), Ws.Cells(LineZ, 7))
        oRng.Interior.Color = Color.LightGreen
        oRng = Ws.Range(Ws.Cells(LineZ, 8), Ws.Cells(LineZ, 9))
        oRng.Interior.Color = Color.Yellow
        oRng = Ws.Range(Ws.Cells(LineZ, 10), Ws.Cells(LineZ, 11))
        oRng.Interior.Color = Color.LightGreen
        oRng = Ws.Range(Ws.Cells(LineZ, 12), Ws.Cells(LineZ, 13))
        oRng.Interior.Color = Color.Yellow
        oRng = Ws.Range(Ws.Cells(LineZ, 14), Ws.Cells(LineZ, 15))
        oRng.Interior.Color = Color.LightGreen
        oRng = Ws.Range(Ws.Cells(LineZ, 16), Ws.Cells(LineZ, 17))
        oRng.Interior.Color = Color.Yellow
        oRng = Ws.Range(Ws.Cells(LineZ, 18), Ws.Cells(LineZ, 19))
        oRng.Interior.Color = Color.LightGreen
        oRng = Ws.Range(Ws.Cells(LineZ, 20), Ws.Cells(LineZ, 21))
        oRng.Interior.Color = Color.Yellow
        oRng = Ws.Range(Ws.Cells(LineZ, 22), Ws.Cells(LineZ, 23))
        oRng.Interior.Color = Color.LightGreen
        oRng = Ws.Range(Ws.Cells(LineZ, 24), Ws.Cells(LineZ, 25))
        oRng.Interior.Color = Color.Yellow

        Ws.Cells(LineZ, 1) = "Revenue:Year-on-year"
        Ws.Cells(LineZ, 2) = tYear & "/01"
        Ws.Cells(LineZ, 3) = tYear & "/01 %"
        Ws.Cells(LineZ, 4) = tYear & "/02"
        Ws.Cells(LineZ, 5) = tYear & "/02 %"
        Ws.Cells(LineZ, 6) = tYear & "/03"
        Ws.Cells(LineZ, 7) = tYear & "/03 %"
        Ws.Cells(LineZ, 8) = tYear & "/04"
        Ws.Cells(LineZ, 9) = tYear & "/04 %"
        Ws.Cells(LineZ, 10) = tYear & "/05"
        Ws.Cells(LineZ, 11) = tYear & "/05 %"
        Ws.Cells(LineZ, 12) = tYear & "/06"
        Ws.Cells(LineZ, 13) = tYear & "/06 %"
        Ws.Cells(LineZ, 14) = tYear & "/07"
        Ws.Cells(LineZ, 15) = tYear & "/07 %"
        Ws.Cells(LineZ, 16) = tYear & "/08"
        Ws.Cells(LineZ, 17) = tYear & "/08 %"
        Ws.Cells(LineZ, 18) = tYear & "/09"
        Ws.Cells(LineZ, 19) = tYear & "/09 %"
        Ws.Cells(LineZ, 20) = tYear & "/10"
        Ws.Cells(LineZ, 21) = tYear & "/10 %"
        Ws.Cells(LineZ, 22) = tYear & "/11"
        Ws.Cells(LineZ, 23) = tYear & "/11 %"
        Ws.Cells(LineZ, 24) = tYear & "/12"
        Ws.Cells(LineZ, 25) = tYear & "/12 %"

        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 25))
        oRng.HorizontalAlignment = xlCenter
        oRng.VerticalAlignment = xlBottom

        '劃線
        oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 25))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        LineZ += 1
    End Sub
    Private Sub AdjustExcelFormat7()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "Qty. Variance % by customer"
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 11
        Ws.Rows.RowHeight = 14
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 17.75
        oRng.EntireColumn.Font.Bold = True
        If gDatabase = "DAC" Then
            Ws.Cells(1, 1) = "Company Name：Dongguan Action Composites LTD Co."
        Else
            Ws.Cells(1, 1) = "Company Name：ACTION COMPOSITE TECHNOLOGY LIMITED"
        End If
        Ws.Cells(2, 1) = "Classification"
        Ws.Cells(2, 2) = "Customer"
        Ws.Cells(2, 3) = "Year"
        For i As Int16 = 1 To 12 Step 1
            'If i < 10 Then
            'Ws.Cells(2, i + 2) = tYear & "/0" & i
            'Else
            'Ws.Cells(2, i + 2) = tYear & "/" & i
            'End If
            Dim TempString As String = String.Empty
            Select Case i
                Case 1
                    TempString = "Jan."
                Case 2
                    TempString = "Feb."
                Case 3
                    TempString = "Mar."
                Case 4
                    TempString = "Apr."
                Case 5
                    TempString = "May."
                Case 6
                    TempString = "Jun."
                Case 7
                    TempString = "Jul."
                Case 8
                    TempString = "Aug."
                Case 9
                    TempString = "Sept."
                Case 10
                    TempString = "Oct."
                Case 11
                    TempString = "Nov."
                Case 12
                    TempString = "Dec."
            End Select
            Ws.Cells(2, i + 3) = TempString
        Next
        oRng = Ws.Range("B2", "P2")
        oRng.HorizontalAlignment = xlCenter
        oRng.VerticalAlignment = xlBottom
        oRng = Ws.Range("C3", "C3")
        oRng.EntireColumn.ColumnWidth = 8.38
        oRng = Ws.Range("D3", "P3")
        oRng.EntireColumn.ColumnWidth = 15


        Ws.Cells(2, 16) = "YTD"

        oRng = Ws.Range("A2", "A2")
        oRng.EntireRow.Font.Bold = True
        oRng = Ws.Range("P2", "P2")
        oRng.Interior.Color = Color.FromArgb(250, 191, 143)

        '劃線
        oRng = Ws.Range("A2", "P2")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        LineZ = 3
    End Sub
    Private Sub DOINPutData1(ByVal tqa02 As String, ByVal sYear As Decimal, ByVal iTerm As Int16)
        oCommand2.CommandText = "select nvl(gea02,'NA') from occ_file left join gea_file on occ20 = gea01 where occacti = 'Y' and occ02 = '" & tqa02 & "'"
        Dim C1 As String = oCommand2.ExecuteScalar()

        oCommand2.CommandText = "select tqa02,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,"
        oCommand2.CommandText += "sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( "
        oCommand2.CommandText += "select tqa02,(case when month(oga02) = 1 then ogb12 else 0 end ) as t1,(case when month(oga02) = 2 then ogb12 else 0 end ) as t2,"
        oCommand2.CommandText += "(case when month(oga02) = 3 then ogb12 else 0 end ) as t3,(case when month(oga02) = 4 then ogb12 else 0 end ) as t4,"
        oCommand2.CommandText += "(case when month(oga02) = 5 then ogb12 else 0 end ) as t5,(case when month(oga02) = 6 then ogb12 else 0 end ) as t6,"
        oCommand2.CommandText += "(case when month(oga02) = 7 then ogb12 else 0 end ) as t7,(case when month(oga02) = 8 then ogb12 else 0 end ) as t8,"
        oCommand2.CommandText += "(case when month(oga02) = 9 then ogb12 else 0 end ) as t9,(case when month(oga02) = 10 then ogb12 else 0 end ) as t10,"
        oCommand2.CommandText += "(case when month(oga02) = 11 then ogb12 else 0 end ) as t11,(case when month(oga02) = 12 then ogb12 else 0 end ) as t12 from oga_file "
        oCommand2.CommandText += "left join ogb_file on oga01 = ogb01 left join ima_file on ogb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand2.CommandText += "where ogapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and tqa02 = '" & tqa02 & "' and year(oga02) = " & sYear & " and month(oga02) <= " & iTerm & " "
        oCommand2.CommandText += "union all "
        oCommand2.CommandText += "select tqa02,(case when month(oha02) = 1 then ohb12 * -1 else 0 end ) as t1,"
        oCommand2.CommandText += "(case when month(oha02) = 2 then ohb12 * -1 else 0 end ) as t2,(case when month(oha02) = 3 then ohb12 * -1 else 0 end ) as t3,"
        oCommand2.CommandText += "(case when month(oha02) = 4 then ohb12 * -1 else 0 end ) as t4,(case when month(oha02) = 5 then ohb12 * -1 else 0 end ) as t5,"
        oCommand2.CommandText += "(case when month(oha02) = 6 then ohb12 * -1 else 0 end ) as t6,(case when month(oha02) = 7 then ohb12 * -1 else 0 end ) as t7,"
        oCommand2.CommandText += "(case when month(oha02) = 8 then ohb12 * -1 else 0 end ) as t8,(case when month(oha02) = 9 then ohb12 * -1 else 0 end ) as t9,"
        oCommand2.CommandText += "(case when month(oha02) = 10 then ohb12 * -1 else 0 end ) as t10,(case when month(oha02) = 11 then ohb12 * -1 else 0 end ) as t11,"
        oCommand2.CommandText += "(case when month(oha02) = 12 then ohb12 * -1 else 0 end ) as t12  from oha_file left join ohb_file on oha01 = ohb01  left join ima_file on ohb04 = ima01 "
        oCommand2.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' where ohapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and tqa02 = '" & tqa02 & "' and year(oha02) = " & sYear & " and month(oha02) <= " & iTerm & " ) group by tqa02 order by tqa02"
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                Ws.Cells(LineZ, 1) = C1
                Ws.Cells(LineZ, 2) = tqa02
                Ws.Cells(LineZ, 3) = "Actual"
                oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
                oRng.HorizontalAlignment = xlCenter
                oRng.VerticalAlignment = xlBottom

                For i As Int16 = 1 To iTerm Step 1
                    Ws.Cells(LineZ, i + 3) = oReader2.Item(i)
                Next
                Dim tColumn As String = String.Empty
                Select Case tMonth
                    Case 1
                        tColumn = "D"
                    Case 2
                        tColumn = "E"
                    Case 3
                        tColumn = "F"
                    Case 4
                        tColumn = "G"
                    Case 5
                        tColumn = "H"
                    Case 6
                        tColumn = "I"
                    Case 7
                        tColumn = "J"
                    Case 8
                        tColumn = "K"
                    Case 9
                        tColumn = "L"
                    Case 10
                        tColumn = "M"
                    Case 11
                        tColumn = "N"
                    Case 12
                        tColumn = "O"
                    Case Else
                        tColumn = "O"
                End Select
                Ws.Cells(LineZ, 16) = "=SUM(D" & LineZ & ":" & tColumn & LineZ & ")"
                LineZ += 1
            End While
        Else
            Ws.Cells(LineZ, 1) = C1
            Ws.Cells(LineZ, 2) = tqa02
            Ws.Cells(LineZ, 3) = sYear
            oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
            oRng.HorizontalAlignment = xlCenter
            oRng.VerticalAlignment = xlBottom

            For i As Int16 = 1 To iTerm Step 1
                Ws.Cells(LineZ, i + 3) = 0
            Next
            LineZ += 1
        End If
        oReader2.Close()
    End Sub
    Private Sub AdjustExcelFormat8(ByVal Capt As String)
        Ws.Cells(LineZ, 1) = Capt
        Ws.Cells(LineZ + 1, 1) = Capt
        Ws.Cells(LineZ, 2) = "Variance Qty"
        Ws.Cells(LineZ + 1, 2) = "Variance %"
        Ws.Cells(LineZ, 4) = "=D" & LineZ - 2 & "-D" & LineZ - 1
        oRng = Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 4))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, tMonth + 3)), Type:=xlFillDefault)
        Ws.Cells(LineZ, 16) = "=SUM(D" & LineZ & ":O" & LineZ & ")"
        'Ws.Cells(LineZ + 1, 3) = "=C" & LineZ & "/C" & LineZ - 1
        'Ws.Cells(LineZ + 1, 3) = "=IF(C" & LineZ - 1 & "="""",0,C" & LineZ & "/C" & LineZ - 1 & ")"
        Ws.Cells(LineZ + 1, 4) = "=IF(D" & LineZ & "=0,0,IF(D" & LineZ - 1 & "=0,1,D" & LineZ & "/D" & LineZ - 1 & "))"
        oRng = Ws.Range(Ws.Cells(LineZ + 1, 4), Ws.Cells(LineZ + 1, 4))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ + 1, 4), Ws.Cells(LineZ + 1, tMonth + 3)), Type:=xlFillDefault)
        Ws.Cells(LineZ + 1, 16) = "=IF(P" & LineZ & "=0,0,IF(P" & LineZ - 1 & "=0,1,P" & LineZ & "/P" & LineZ - 1 & "))"
        ' 上色
        oRng = Ws.Range(Ws.Cells(LineZ - 2, 16), Ws.Cells(LineZ + 1, 16))
        oRng.Interior.Color = Color.FromArgb(250, 191, 143)
        'oRng = Ws.Range(Ws.Cells(LineZ + 2, 1), Ws.Cells(LineZ + 2, 15))
        'oRng.Merge()
        'oRng.Interior.Color = Color.Yellow

        '格式
        oRng = Ws.Range(Ws.Cells(LineZ - 2, 4), Ws.Cells(LineZ, 16))
        oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        oRng = Ws.Range(Ws.Cells(LineZ + 1, 4), Ws.Cells(LineZ + 1, 16))
        oRng.NumberFormatLocal = "0%"
        oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
        oRng.FormatConditions(1).FONT.COLOR = Color.Red

        '劃線
        oRng = Ws.Range(Ws.Cells(LineZ - 2, 1), Ws.Cells(LineZ + 1, 16))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        LineZ += 2
    End Sub
    Private Sub AdjustExcelFormat9(ByVal tqa02 As String)
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = tqa02
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 11
        Ws.Rows.RowHeight = 14
        oRng = Ws.Range("A1", "C1")
        'oRng.EntireColumn.ColumnWidth =
        oRng.EntireColumn.Font.Bold = True
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 32.5
        oRng = Ws.Range("B3", "B3")
        oRng.EntireColumn.ColumnWidth = 45
        oRng = Ws.Range("C3", "C3")
        oRng.EntireColumn.ColumnWidth = 25
        oRng = Ws.Range("D3", "D3")
        oRng.EntireColumn.ColumnWidth = 8
        oRng = Ws.Range("E3", "AD3")
        oRng.EntireColumn.ColumnWidth = 15

        If gDatabase = "DAC" Then
            Ws.Cells(1, 1) = "Company Name：Dongguan Action Composites LTD Co."
        Else
            Ws.Cells(1, 1) = "Company Name：ACTION COMPOSITE TECHNOLOGY LIMITED"
        End If
        Ws.Cells(2, 1) = "Customer:" & tqa02
        Ws.Cells(2, 2) = "Currency：" & gCurrency
        Ws.Cells(3, 1) = "Revenue of Y" & tYear & "（Amount）"
        Ws.Cells(4, 1) = "Part Name"
        Ws.Cells(4, 2) = "Part Description"
        Ws.Cells(4, 3) = "Spec."
        Ws.Cells(4, 4) = "Uint"
        For i As Int16 = 1 To 12 Step 1
            oRng = Ws.Range(Ws.Cells(3, 3 + i * 2), Ws.Cells(3, 4 + i * 2))
            oRng.Merge()
            oRng.HorizontalAlignment = xlCenter
            If i < 10 Then
                Ws.Cells(3, 3 + 2 * i) = tYear & "/0" & i
                'Ws.Cells(3, 4 + 2 * i) = tYear & "/0" & i
            Else
                Ws.Cells(3, 3 + 2 * i) = tYear & "/" & i
                'Ws.Cells(3, 4 + 2 * i) = tYear & "/" & i
            End If
            Ws.Cells(4, 3 + 2 * i) = "Actual"
            Ws.Cells(4, 4 + 2 * i) = "Budget"
        Next
        oRng = Ws.Range("AC3", "AF3")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(3, 29) = "YTD" & tYear
        'Ws.Cells(3, 30) = "YTD" & tYear
        Ws.Cells(4, 29) = "Actual"
        Ws.Cells(4, 30) = "Budget"
        Ws.Cells(4, 31) = "difference"
        Ws.Cells(4, 32) = "%"

        oRng = Ws.Range("D3", "AF4")
        oRng.HorizontalAlignment = xlCenter
        oRng.VerticalAlignment = xlBottom
        oRng = Ws.Range("D3", "D3")
        oRng.EntireColumn.HorizontalAlignment = xlCenter
        oRng.EntireColumn.VerticalAlignment = xlBottom

        oRng = Ws.Range("A2", "B2")
        oRng.Interior.Color = Color.Yellow

        '劃線
        oRng = Ws.Range("A3", "AF4")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        LineZ = 5
    End Sub
    Private Sub DOINPutData2(ByVal tqa02 As String, ByVal sYear As Decimal, iTerm As Int16)
        LineS1 = LineZ
        oCommand2.CommandText = "select ogb04,ima02,ima021,ogb05,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,"
        oCommand2.CommandText += "sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12,sum(t13) as t13,sum(t14) as t14,sum(t15) as t15,sum(t16) as t16,sum(t17) as t17,"
        oCommand2.CommandText += "sum(t18) as t18,sum(t19) as t19,sum(t20) as t20,sum(t21) as t21,sum(t22) as t22,sum(t23) as t23,sum(t24) as t24 from ( "
        oCommand2.CommandText += "select ogb04,ima02,ima021,ogb05,(case when month(oga02) = 1 then ogb14 * oga24 else 0 end ) as t1,0 as t2,(case when month(oga02) = 2 then ogb14 * oga24 else 0 end ) as t3,0 as t4,"
        oCommand2.CommandText += "(case when month(oga02) = 3 then ogb14 * oga24 else 0 end ) as t5,0 as t6,(case when month(oga02) = 4 then ogb14 * oga24 else 0 end ) as t7,0 as t8,"
        oCommand2.CommandText += "(case when month(oga02) = 5 then ogb14 * oga24 else 0 end ) as t9,0 as t10,(case when month(oga02) = 6 then ogb14 * oga24 else 0 end ) as t11,0 as t12,"
        oCommand2.CommandText += "(case when month(oga02) = 7 then ogb14 * oga24 else 0 end ) as t13,0 as t14,(case when month(oga02) = 8 then ogb14 * oga24 else 0 end ) as t15,0 as t16,"
        oCommand2.CommandText += "(case when month(oga02) = 9 then ogb14 * oga24 else 0 end ) as t17,0 as t18,(case when month(oga02) = 10 then ogb14 * oga24 else 0 end ) as t19,0 as t20,"
        oCommand2.CommandText += "(case when month(oga02) = 11 then ogb14 * oga24 else 0 end ) as t21,0 as t22,(case when month(oga02) = 12 then ogb14 * oga24 else 0 end ) as t23,0 as t24 "
        oCommand2.CommandText += "from oga_file left join ogb_file on oga01 = ogb01 left join ima_file on ogb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand2.CommandText += "where ogapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and tqa02 = '" & tqa02 & "' and year(oga02) = " & sYear & " and month(oga02) <= " & iTerm
        oCommand2.CommandText += " union all "
        oCommand2.CommandText += "select tc_bud04,ima02,ima021,ima25,0,(case when tc_bud03 = 1 then c1 else 0 end),0,(case when tc_bud03 = 2 then c1 else 0 end) "
        oCommand2.CommandText += ",0,(case when tc_bud03 = 3 then c1 else 0 end),0,(case when tc_bud03 = 4 then c1 else 0 end),0,(case when tc_bud03 = 5 then c1 else 0 end)"
        oCommand2.CommandText += ",0,(case when tc_bud03 = 6 then c1 else 0 end),0,(case when tc_bud03 = 7 then c1 else 0 end),0,(case when tc_bud03 = 8 then c1 else 0 end)"
        oCommand2.CommandText += ",0,(case when tc_bud03 = 9 then c1 else 0 end),0,(case when tc_bud03 = 10 then c1 else 0 end),0,(case when tc_bud03 = 11 then c1 else 0 end)"
        oCommand2.CommandText += ",0,(case when tc_bud03 = 12 then c1 else 0 end) from ( "
        If gDatabase = "DAC" Then
            '190319 add by Brady
            'oCommand2.CommandText += "select tc_bud03,tc_bud04,(case when tc_bud14 = 'USD' then tc_bud13 * 6.3 else tc_bud13 * 7.56 end ) as c1 from tc_bud_file "
            oCommand2.CommandText += "select tc_bud03,tc_bud04,(case when tc_bud14 = 'USD' then tc_bud13 * 6.85 else tc_bud13 * 8.22 end ) as c1 from tc_bud_file "
            '190319 add by Brady END
        Else
            oCommand2.CommandText += "select tc_bud03,tc_bud04,(case when tc_bud14 = 'EUR' then tc_bud13 * 1.2 else tc_bud13 end ) as c1 from tc_bud_file "
        End If
        oCommand2.CommandText += "where tc_bud01 = 1 and tc_bud02 = " & sYear & " and tc_bud05 = '" & tqa02 & "' ) AA left join ima_file on AA.tc_bud04 = ima01 ) group by ogb04,ima02,ima021,ogb05 order by ogb04"
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                For i As Int16 = 1 To 28 Step 1
                    Ws.Cells(LineZ, i) = oReader2.Item(i - 1)
                Next
                Ws.Cells(LineZ, 29) = "=E" & LineZ & "+G" & LineZ & "+I" & LineZ & "+K" & LineZ & "+M" & LineZ & "+O" & LineZ & "+Q" & LineZ & "+S" & LineZ & "+U" & LineZ & "+W" & LineZ & "+Y" & LineZ & "+AA" & LineZ
                Ws.Cells(LineZ, 30) = "=F" & LineZ & "+H" & LineZ & "+J" & LineZ & "+L" & LineZ & "+N" & LineZ & "+P" & LineZ & "+R" & LineZ & "+T" & LineZ & "+V" & LineZ & "+X" & LineZ & "+Z" & LineZ & "+AB" & LineZ
                Ws.Cells(LineZ, 31) = "=AC" & LineZ & "-AD" & LineZ
                'Ws.Cells(LineZ, 32) = "=AC" & LineZ & "/AD" & LineZ
                Ws.Cells(LineZ, 32) = "=IFERROR(AC" & LineZ & "/AD" & LineZ & ",0)"
                LineZ += 1
                End While
        Else
            LineZ += 1
            End If
        oReader2.Close()

            ' 加總
        Ws.Cells(LineZ, 1) = "Total"
        Ws.Cells(LineZ, 5) = "=SUM(E" & LineS1 & ":E" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 5), Ws.Cells(LineZ, 5))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 5), Ws.Cells(LineZ, 30)), Type:=xlFillDefault)
        Ws.Cells(LineZ, 31) = "=AC" & LineZ & "-AD" & LineZ
        'Ws.Cells(LineZ, 32) = "=AC" & LineZ & "/AD" & LineZ
        Ws.Cells(LineZ, 32) = "=IFERROR(AC" & LineZ & "/AD" & LineZ & ",0)"
            ' 格式
        oRng = Ws.Range(Ws.Cells(LineS1, 5), Ws.Cells(LineZ, 31))
        oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        oRng = Ws.Range("AF1", "AF1")
        oRng.EntireColumn.NumberFormat = "0%"
            ' 劃線
        oRng = Ws.Range(Ws.Cells(LineS1, 1), Ws.Cells(LineZ, 32))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

    End Sub
    Private Sub AdjustExcelFormat10(ByVal tqa02 As String)
        Ws.Cells(LineZ, 1) = "Customer:" & tqa02
        'Ws.Cells(LineZ, 2) = "Currency：" & gCurrency
        oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 2))
        oRng.Interior.Color = Color.Yellow
        LineZ += 1

        Ws.Cells(LineZ, 1) = "Revenue of Y" & tYear & "（Qty）"
        Ws.Cells(LineZ + 1, 1) = "Part Name"
        Ws.Cells(LineZ + 1, 2) = "Part Description"
        Ws.Cells(LineZ + 1, 3) = "Spec."
        Ws.Cells(LineZ + 1, 4) = "Uint"
        For i As Int16 = 1 To 12 Step 1
            oRng = Ws.Range(Ws.Cells(LineZ, 3 + i * 2), Ws.Cells(LineZ, 4 + i * 2))
            oRng.Merge()
            If i < 10 Then
                Ws.Cells(LineZ, 3 + 2 * i) = tYear & "/0" & i
                'Ws.Cells(LineZ, 4 + 2 * i) = pYear & "/0" & i
            Else
                Ws.Cells(LineZ, 3 + 2 * i) = tYear & "/" & i
                'Ws.Cells(LineZ, 4 + 2 * i) = pYear & "/" & i
            End If
            Ws.Cells(LineZ + 1, 3 + 2 * i) = "Actual"
            Ws.Cells(LineZ + 1, 4 + 2 * i) = "Budget"
        Next
        oRng = Ws.Range(Ws.Cells(LineZ, 5), Ws.Cells(LineZ + 1, 32))
        oRng.HorizontalAlignment = xlCenter
        oRng.VerticalAlignment = xlBottom

        Ws.Cells(LineZ, 29) = "YTD" & pYear
        'Ws.Cells(LineZ, 30) = "YTD" & pYear
        Ws.Cells(LineZ + 1, 29) = "Actual"
        Ws.Cells(LineZ + 1, 30) = "Budget"
        Ws.Cells(LineZ + 1, 31) = "difference"
        Ws.Cells(LineZ + 1, 32) = "%"

        '劃線
        oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ + 1, 32))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        LineZ += 2
    End Sub
    Private Sub DOINPutData1A(ByVal tqa02 As String, ByVal sYear As Decimal, ByVal iTerm As Int16)
        oCommand2.CommandText = "select nvl(gea02,'NA') from occ_file left join gea_file on occ20 = gea01 where occ02 = '" & tqa02 & "'"
        Dim C1 As String = oCommand2.ExecuteScalar()

        oCommand2.CommandText = "select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,"
        oCommand2.CommandText += "sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( "
        oCommand2.CommandText += "select (case when tc_bud03 = 1 then t1 else 0 end ) as t1,(case when tc_bud03 = 2 then t1 else 0 end ) as t2,"
        oCommand2.CommandText += "(case when tc_bud03 = 3 then t1 else 0 end ) as t3,(case when tc_bud03 = 4 then t1 else 0 end ) as t4,"
        oCommand2.CommandText += "(case when tc_bud03 = 5 then t1 else 0 end ) as t5,(case when tc_bud03 = 6 then t1 else 0 end ) as t6,"
        oCommand2.CommandText += "(case when tc_bud03 = 7 then t1 else 0 end ) as t7,(case when tc_bud03 = 8 then t1 else 0 end ) as t8,"
        oCommand2.CommandText += "(case when tc_bud03 = 9 then t1 else 0 end ) as t9,(case when tc_bud03 = 10 then t1 else 0 end ) as t10,"
        oCommand2.CommandText += "(case when tc_bud03 = 11 then t1 else 0 end ) as t11,(case when tc_bud03 = 12 then t1 else 0 end ) as t12 "
        oCommand2.CommandText += "from ( select tc_bud03,tc_bud11 as t1 "
        oCommand2.CommandText += "from tc_bud_file where tc_bud02 = " & sYear & " and tc_bud01 = 1 AND tc_bud05 = '" & tqa02 & "' )  )"
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                Ws.Cells(LineZ, 1) = C1
                Ws.Cells(LineZ, 2) = tqa02
                'Ws.Cells(LineZ, 2) = sYear
                Ws.Cells(LineZ, 3) = "Budget"
                oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
                oRng.HorizontalAlignment = xlCenter
                oRng.VerticalAlignment = xlBottom

                For i As Int16 = 1 To iTerm Step 1
                    Ws.Cells(LineZ, i + 3) = oReader2.Item(i - 1)
                Next
                Dim tColumn As String = String.Empty
                Select Case tMonth
                    Case 1
                        tColumn = "D"
                    Case 2
                        tColumn = "E"
                    Case 3
                        tColumn = "F"
                    Case 4
                        tColumn = "G"
                    Case 5
                        tColumn = "H"
                    Case 6
                        tColumn = "I"
                    Case 7
                        tColumn = "J"
                    Case 8
                        tColumn = "K"
                    Case 9
                        tColumn = "L"
                    Case 10
                        tColumn = "M"
                    Case 11
                        tColumn = "N"
                    Case 12
                        tColumn = "O"
                    Case Else
                        tColumn = "O"
                End Select
                Ws.Cells(LineZ, 16) = "=SUM(D" & LineZ & ":" & tColumn & LineZ & ")"
                LineZ += 1
            End While
        Else
            Ws.Cells(LineZ, 1) = C1
            Ws.Cells(LineZ, 2) = tqa02
            'Ws.Cells(LineZ, 2) = sYear
            Ws.Cells(LineZ, 3) = "Budget"
            oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
            oRng.HorizontalAlignment = xlCenter
            oRng.VerticalAlignment = xlBottom

            For i As Int16 = 1 To iTerm Step 1
                Ws.Cells(LineZ, i + 3) = 0
            Next
            LineZ += 1
        End If
        oReader2.Close()
    End Sub
    Private Sub DOINPutData2A(ByVal tqa02 As String, ByVal sYear As Decimal, iTerm As Int16)
        LineS1 = LineZ
        oCommand2.CommandText = "select ogb04,ima02,ima021,ogb05,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,"
        oCommand2.CommandText += "sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12,sum(t13) as t13,sum(t14) as t14,sum(t15) as t15,sum(t16) as t16,sum(t17) as t17,"
        oCommand2.CommandText += "sum(t18) as t18,sum(t19) as t19,sum(t20) as t20,sum(t21) as t21,sum(t22) as t22,sum(t23) as t23,sum(t24) as t24 from ( "
        oCommand2.CommandText += "select ogb04,ima02,ima021,ogb05,(case when month(oga02) = 1 then ogb12 else 0 end ) as t1,0 as t2,(case when month(oga02) = 2 then ogb12 else 0 end ) as t3,0 as t4,"
        oCommand2.CommandText += "(case when month(oga02) = 3 then ogb12 else 0 end ) as t5,0 as t6,(case when month(oga02) = 4 then ogb12 else 0 end ) as t7,0 as t8,"
        oCommand2.CommandText += "(case when month(oga02) = 5 then ogb12 else 0 end ) as t9,0 as t10,(case when month(oga02) = 6 then ogb12 else 0 end ) as t11,0 as t12,"
        oCommand2.CommandText += "(case when month(oga02) = 7 then ogb12 else 0 end ) as t13,0 as t14,(case when month(oga02) = 8 then ogb12 else 0 end ) as t15,0 as t16,"
        oCommand2.CommandText += "(case when month(oga02) = 9 then ogb12 else 0 end ) as t17,0 as t18,(case when month(oga02) = 10 then ogb12 else 0 end ) as t19,0 as t20,"
        oCommand2.CommandText += "(case when month(oga02) = 11 then ogb12 else 0 end ) as t21,0 as t22,(case when month(oga02) = 12 then ogb12 else 0 end ) as t23,0 as t24 "
        oCommand2.CommandText += "from oga_file left join ogb_file on oga01 = ogb01 left join ima_file on ogb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand2.CommandText += "where ogapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and tqa02 = '" & tqa02 & "' and year(oga02) = " & sYear & " and month(oga02) <= " & iTerm
        oCommand2.CommandText += " union all "
        oCommand2.CommandText += "select tc_bud04,ima02,ima021,ima25,0,(case when tc_bud03 = 1 then c1 else 0 end),0,(case when tc_bud03 = 2 then c1 else 0 end) "
        oCommand2.CommandText += ",0,(case when tc_bud03 = 3 then c1 else 0 end),0,(case when tc_bud03 = 4 then c1 else 0 end),0,(case when tc_bud03 = 5 then c1 else 0 end)"
        oCommand2.CommandText += ",0,(case when tc_bud03 = 6 then c1 else 0 end),0,(case when tc_bud03 = 7 then c1 else 0 end),0,(case when tc_bud03 = 8 then c1 else 0 end)"
        oCommand2.CommandText += ",0,(case when tc_bud03 = 9 then c1 else 0 end),0,(case when tc_bud03 = 10 then c1 else 0 end),0,(case when tc_bud03 = 11 then c1 else 0 end)"
        oCommand2.CommandText += ",0,(case when tc_bud03 = 12 then c1 else 0 end) from ( "
        oCommand2.CommandText += "select tc_bud03,tc_bud04,tc_bud11 as c1 from tc_bud_file "
        oCommand2.CommandText += "where tc_bud01 = 1 and tc_bud02 = " & sYear & " and tc_bud05 = '" & tqa02 & "' ) AA left join ima_file on AA.tc_bud04 = ima01 ) group by ogb04,ima02,ima021,ogb05 order by ogb04"
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                For i As Int16 = 1 To 28 Step 1
                    Ws.Cells(LineZ, i) = oReader2.Item(i - 1)
                Next
                Ws.Cells(LineZ, 29) = "=E" & LineZ & "+G" & LineZ & "+I" & LineZ & "+K" & LineZ & "+M" & LineZ & "+O" & LineZ & "+Q" & LineZ & "+S" & LineZ & "+U" & LineZ & "+W" & LineZ & "+Y" & LineZ & "+AA" & LineZ
                Ws.Cells(LineZ, 30) = "=F" & LineZ & "+H" & LineZ & "+J" & LineZ & "+L" & LineZ & "+N" & LineZ & "+P" & LineZ & "+R" & LineZ & "+T" & LineZ & "+V" & LineZ & "+X" & LineZ & "+Z" & LineZ & "+AB" & LineZ
                Ws.Cells(LineZ, 31) = "=AC" & LineZ & "-AD" & LineZ
                'Ws.Cells(LineZ, 32) = "=AC" & LineZ & "/AD" & LineZ
                Ws.Cells(LineZ, 32) = "=IFERROR(AC" & LineZ & "/AD" & LineZ & ",0)"
                LineZ += 1
            End While
        Else
            LineZ += 1
        End If
        oReader2.Close()

        ' 加總
        Ws.Cells(LineZ, 1) = "Total"
        Ws.Cells(LineZ, 5) = "=SUM(E" & LineS1 & ":E" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 5), Ws.Cells(LineZ, 5))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 5), Ws.Cells(LineZ, 30)), Type:=xlFillDefault)
        Ws.Cells(LineZ, 31) = "=AC" & LineZ & "-AD" & LineZ
        'Ws.Cells(LineZ, 32) = "=AC" & LineZ & "/AD" & LineZ
        Ws.Cells(LineZ, 32) = "=IFERROR(AC" & LineZ & "/AD" & LineZ & ",0)"
        ' 格式
        oRng = Ws.Range(Ws.Cells(LineS1, 5), Ws.Cells(LineZ, 31))
        oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        ' 劃線
        oRng = Ws.Range(Ws.Cells(LineS1, 1), Ws.Cells(LineZ, 32))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

    End Sub
    Private Sub AdjustExcelFormat11()
        For i As Int16 = 1 To 12 Step 1
            Dim TempString As String = String.Empty
            Select Case i
                Case 1
                    TempString = "Jan."
                Case 2
                    TempString = "Feb."
                Case 3
                    TempString = "Mar."
                Case 4
                    TempString = "Apr."
                Case 5
                    TempString = "May."
                Case 6
                    TempString = "Jun."
                Case 7
                    TempString = "Jul."
                Case 8
                    TempString = "Aug."
                Case 9
                    TempString = "Sept."
                Case 10
                    TempString = "Oct."
                Case 11
                    TempString = "Nov."
                Case 12
                    TempString = "Dec."
            End Select
            Ws.Cells(LineZ, i + 3) = TempString

            'End If
        Next
        Ws.Cells(LineZ, 16) = "YTD"

        '劃線
        oRng = Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 16))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
    End Sub
    Private Sub GeaDataBudget(ByVal l_gea02 As String, ByVal sYear As Int16, ByVal iTerm As Decimal)
        If gDatabase = "DAC" Then
            oCommand2.CommandText = "select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,"
            oCommand2.CommandText += "sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( "
            oCommand2.CommandText += "select (case when tc_bud03 = 1 then t1 else 0 end ) as t1,(case when tc_bud03 = 2 then t1 else 0 end ) as t2,"
            oCommand2.CommandText += "(case when tc_bud03 = 3 then t1 else 0 end ) as t3,(case when tc_bud03 = 4 then t1 else 0 end ) as t4,"
            oCommand2.CommandText += "(case when tc_bud03 = 5 then t1 else 0 end ) as t5,(case when tc_bud03 = 6 then t1 else 0 end ) as t6,"
            oCommand2.CommandText += "(case when tc_bud03 = 7 then t1 else 0 end ) as t7,(case when tc_bud03 = 8 then t1 else 0 end ) as t8,"
            oCommand2.CommandText += "(case when tc_bud03 = 9 then t1 else 0 end ) as t9,(case when tc_bud03 = 10 then t1 else 0 end ) as t10,"
            oCommand2.CommandText += "(case when tc_bud03 = 11 then t1 else 0 end ) as t11,(case when tc_bud03 = 12 then t1 else 0 end ) as t12 "
            oCommand2.CommandText += "from ( select tc_bud03,(case when tc_bud14 = 'USD' then tc_bud13 * 6.85 else tc_bud13 * 8.22 end ) as t1 "
            oCommand2.CommandText += "from tc_bud_file left join occ_file on tc_bud05 = occ02 and occacti = 'Y' left join gea_file on occ20 = gea01 where tc_bud02 = " & sYear
            oCommand2.CommandText += " and tc_bud01 = 1 and gea02 = '" & l_gea02 & "' )  )"
        Else
            oCommand2.CommandText = "select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,"
            oCommand2.CommandText += "sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( "
            oCommand2.CommandText += "select (case when tc_bud03 = 1 then t1 else 0 end ) as t1,(case when tc_bud03 = 2 then t1 else 0 end ) as t2,"
            oCommand2.CommandText += "(case when tc_bud03 = 3 then t1 else 0 end ) as t3,(case when tc_bud03 = 4 then t1 else 0 end ) as t4,"
            oCommand2.CommandText += "(case when tc_bud03 = 5 then t1 else 0 end ) as t5,(case when tc_bud03 = 6 then t1 else 0 end ) as t6,"
            oCommand2.CommandText += "(case when tc_bud03 = 7 then t1 else 0 end ) as t7,(case when tc_bud03 = 8 then t1 else 0 end ) as t8,"
            oCommand2.CommandText += "(case when tc_bud03 = 9 then t1 else 0 end ) as t9,(case when tc_bud03 = 10 then t1 else 0 end ) as t10,"
            oCommand2.CommandText += "(case when tc_bud03 = 11 then t1 else 0 end ) as t11,(case when tc_bud03 = 12 then t1 else 0 end ) as t12 "
            oCommand2.CommandText += "from ( select tc_bud03,(case when tc_bud14 = 'EUR' then tc_bud13 * 1.2 else tc_bud13 end ) as t1 "
            oCommand2.CommandText += "from tc_bud_file left join occ_file on tc_bud05 = occ02 and occacti = 'Y' left join gea_file on occ20 = gea01 where tc_bud02 = " & sYear
            oCommand2.CommandText += " and tc_bud01 = 1 and gea02 = '" & l_gea02 & "' )  )"
        End If
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                For i As Int16 = 1 To iTerm Step 1
                    Ws.Cells(LineZ, i + 3) = oReader2.Item(i - 1)
                Next
                Dim tColumn As String = String.Empty
                Select Case tMonth
                    Case 1
                        tColumn = "D"
                    Case 2
                        tColumn = "E"
                    Case 3
                        tColumn = "F"
                    Case 4
                        tColumn = "G"
                    Case 5
                        tColumn = "H"
                    Case 6
                        tColumn = "I"
                    Case 7
                        tColumn = "J"
                    Case 8
                        tColumn = "K"
                    Case 9
                        tColumn = "L"
                    Case 10
                        tColumn = "M"
                    Case 11
                        tColumn = "N"
                    Case 12
                        tColumn = "O"
                    Case Else
                        tColumn = "O"
                End Select
                Ws.Cells(LineZ, 16) = "=SUM(D" & LineZ & ":" & tColumn & LineZ & ")"
                LineZ += 1
            End While
        Else
            For i As Int16 = 1 To iTerm Step 1
                Ws.Cells(LineZ, i + 3) = 0
            Next

            LineZ += 1
        End If
        oReader2.Close()
    End Sub
    Private Sub GeaDataSales(ByVal l_gea02 As String, ByVal syear As Int16, ByVal iTerm As Decimal)
        oCommand2.CommandText = "select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,"
        oCommand2.CommandText += "sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( "
        oCommand2.CommandText += "select (case when month(oga02) = 1 then ogb14 * oga24 else 0 end ) as t1,(case when month(oga02) = 2 then ogb14 * oga24 else 0 end ) as t2,"
        oCommand2.CommandText += "(case when month(oga02) = 3 then ogb14 * oga24 else 0 end ) as t3,(case when month(oga02) = 4 then ogb14 * oga24 else 0 end ) as t4,"
        oCommand2.CommandText += "(case when month(oga02) = 5 then ogb14 * oga24 else 0 end ) as t5,(case when month(oga02) = 6 then ogb14 * oga24 else 0 end ) as t6,"
        oCommand2.CommandText += "(case when month(oga02) = 7 then ogb14 * oga24 else 0 end ) as t7,(case when month(oga02) = 8 then ogb14 * oga24 else 0 end ) as t8,"
        oCommand2.CommandText += "(case when month(oga02) = 9 then ogb14 * oga24 else 0 end ) as t9,(case when month(oga02) = 10 then ogb14 * oga24 else 0 end ) as t10,"
        oCommand2.CommandText += "(case when month(oga02) = 11 then ogb14 * oga24 else 0 end ) as t11,(case when month(oga02) = 12 then ogb14 * oga24 else 0 end ) as t12 from oga_file "
        oCommand2.CommandText += "left join ogb_file on oga01 = ogb01 left join ima_file on ogb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand2.CommandText += "left join occ_file on tqa02 = occ02 and occacti = 'Y' left join gea_file on occ20 = gea01 "
        oCommand2.CommandText += "where ogapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and gea02 = '" & l_gea02 & "' and year(oga02) = " & syear & " and month(oga02) <= " & iTerm & " "
        oCommand2.CommandText += "union all "
        oCommand2.CommandText += "select (case when month(oha02) = 1 then ohb14 * oha24 * -1 else 0 end ) as t1,"
        oCommand2.CommandText += "(case when month(oha02) = 2 then ohb14 * oha24 * -1 else 0 end ) as t2,(case when month(oha02) = 3 then ohb14 * oha24 * -1 else 0 end ) as t3,"
        oCommand2.CommandText += "(case when month(oha02) = 4 then ohb14 * oha24 * -1 else 0 end ) as t4,(case when month(oha02) = 5 then ohb14 * oha24 * -1 else 0 end ) as t5,"
        oCommand2.CommandText += "(case when month(oha02) = 6 then ohb14 * oha24 * -1 else 0 end ) as t6,(case when month(oha02) = 7 then ohb14 * oha24 * -1 else 0 end ) as t7,"
        oCommand2.CommandText += "(case when month(oha02) = 8 then ohb14 * oha24 * -1 else 0 end ) as t8,(case when month(oha02) = 9 then ohb14 * oha24 * -1 else 0 end ) as t9,"
        oCommand2.CommandText += "(case when month(oha02) = 10 then ohb14 * oha24 * -1 else 0 end ) as t10,(case when month(oha02) = 11 then ohb14 * oha24 * -1 else 0 end ) as t11,"
        oCommand2.CommandText += "(case when month(oha02) = 12 then ohb14 * oha24 * -1 else 0 end ) as t12  from oha_file left join ohb_file on oha01 = ohb01  left join ima_file on ohb04 = ima01 "
        oCommand2.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join occ_file on tqa02 = occ02 and occacti = 'Y' left join gea_file on occ20 = gea01 where ohapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and gea02  = '" & l_gea02 & "' and year(oha02) = " & syear & " and month(oha02) <= " & iTerm & " ) "
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                For i As Int16 = 1 To iTerm Step 1
                    Ws.Cells(LineZ, i + 3) = oReader2.Item(i - 1)
                Next
                Dim tColumn As String = String.Empty
                Select Case tMonth
                    Case 1
                        tColumn = "D"
                    Case 2
                        tColumn = "E"
                    Case 3
                        tColumn = "F"
                    Case 4
                        tColumn = "G"
                    Case 5
                        tColumn = "H"
                    Case 6
                        tColumn = "I"
                    Case 7
                        tColumn = "J"
                    Case 8
                        tColumn = "K"
                    Case 9
                        tColumn = "L"
                    Case 10
                        tColumn = "M"
                    Case 11
                        tColumn = "N"
                    Case 12
                        tColumn = "O"
                    Case Else
                        tColumn = "O"
                End Select
                Ws.Cells(LineZ, 16) = "=SUM(D" & LineZ & ":" & tColumn & LineZ & ")"
                LineZ += 1
            End While
        Else
            For i As Int16 = 1 To iTerm Step 1
                Ws.Cells(LineZ, i + 3) = 0
            Next

            LineZ += 1
        End If
        oReader2.Close()
    End Sub
    Private Sub AdjustExcelFormat12()

        Ws.Cells(LineZ, 4) = "=D" & LineZ - 1 & "-D" & LineZ - 2
        oRng = Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 4))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, tMonth + 3)), Type:=xlFillDefault)
        Ws.Cells(LineZ, 16) = "=SUM(D" & LineZ & ":O" & LineZ & ")"
        'Ws.Cells(LineZ + 1, 3) = "=C" & LineZ & "/C" & LineZ - 1
        'Ws.Cells(LineZ + 1, 3) = "=IF(C" & LineZ - 1 & "="""",0,C" & LineZ & "/C" & LineZ - 1 & ")"
        Ws.Cells(LineZ + 1, 4) = "=IF(D" & LineZ & "=0,0,IF(D" & LineZ - 2 & "=0,1,D" & LineZ & "/D" & LineZ - 2 & "))"
        oRng = Ws.Range(Ws.Cells(LineZ + 1, 4), Ws.Cells(LineZ + 1, 4))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ + 1, 4), Ws.Cells(LineZ + 1, tMonth + 3)), Type:=xlFillDefault)
        Ws.Cells(LineZ + 1, 16) = "=IF(P" & LineZ & "=0,0,IF(P" & LineZ - 2 & "=0,1,P" & LineZ & "/P" & LineZ - 2 & "))"
        ' 上色
        oRng = Ws.Range(Ws.Cells(LineZ - 2, 16), Ws.Cells(LineZ + 1, 16))
        oRng.Interior.Color = Color.FromArgb(250, 191, 143)

        '格式
        oRng = Ws.Range(Ws.Cells(LineZ - 2, 4), Ws.Cells(LineZ, 16))
        oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        oRng = Ws.Range(Ws.Cells(LineZ + 1, 4), Ws.Cells(LineZ + 1, 16))
        oRng.NumberFormatLocal = "0%"
        ' 添加 負數為紅色 20180531
        oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
        oRng.FormatConditions(1).FONT.COLOR = Color.Red

        '劃線
        oRng = Ws.Range(Ws.Cells(LineZ - 2, 3), Ws.Cells(LineZ + 1, 16))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        LineZ += 2
    End Sub
    Private Sub GeaDataBudget1(ByVal l_gea02 As String, ByVal sYear As Int16, ByVal iTerm As Decimal)

        oCommand2.CommandText = "select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,"
        oCommand2.CommandText += "sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( "
        oCommand2.CommandText += "select (case when tc_bud03 = 1 then t1 else 0 end ) as t1,(case when tc_bud03 = 2 then t1 else 0 end ) as t2,"
        oCommand2.CommandText += "(case when tc_bud03 = 3 then t1 else 0 end ) as t3,(case when tc_bud03 = 4 then t1 else 0 end ) as t4,"
        oCommand2.CommandText += "(case when tc_bud03 = 5 then t1 else 0 end ) as t5,(case when tc_bud03 = 6 then t1 else 0 end ) as t6,"
        oCommand2.CommandText += "(case when tc_bud03 = 7 then t1 else 0 end ) as t7,(case when tc_bud03 = 8 then t1 else 0 end ) as t8,"
        oCommand2.CommandText += "(case when tc_bud03 = 9 then t1 else 0 end ) as t9,(case when tc_bud03 = 10 then t1 else 0 end ) as t10,"
        oCommand2.CommandText += "(case when tc_bud03 = 11 then t1 else 0 end ) as t11,(case when tc_bud03 = 12 then t1 else 0 end ) as t12 "
        oCommand2.CommandText += "from ( select tc_bud03,tc_bud11 as t1 "
        oCommand2.CommandText += "from tc_bud_file left join occ_file on tc_bud05 = occ02 and occacti = 'Y' left join gea_file on occ20 = gea01 where tc_bud02 = " & sYear
        oCommand2.CommandText += " and tc_bud01 = 1 and gea02 = '" & l_gea02 & "' )  )"

        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                For i As Int16 = 1 To iTerm Step 1
                    Ws.Cells(LineZ, i + 3) = oReader2.Item(i - 1)
                Next
                Dim tColumn As String = String.Empty
                Select Case tMonth
                    Case 1
                        tColumn = "D"
                    Case 2
                        tColumn = "E"
                    Case 3
                        tColumn = "F"
                    Case 4
                        tColumn = "G"
                    Case 5
                        tColumn = "H"
                    Case 6
                        tColumn = "I"
                    Case 7
                        tColumn = "J"
                    Case 8
                        tColumn = "K"
                    Case 9
                        tColumn = "L"
                    Case 10
                        tColumn = "M"
                    Case 11
                        tColumn = "N"
                    Case 12
                        tColumn = "O"
                    Case Else
                        tColumn = "O"
                End Select
                Ws.Cells(LineZ, 16) = "=SUM(D" & LineZ & ":" & tColumn & LineZ & ")"
                LineZ += 1
            End While
        Else
            For i As Int16 = 1 To iTerm Step 1
                Ws.Cells(LineZ, i + 3) = 0
            Next

            LineZ += 1
        End If
        oReader2.Close()
    End Sub
    Private Sub GeaDataSales1(ByVal l_gea02 As String, ByVal syear As Int16, ByVal iTerm As Decimal)
        oCommand2.CommandText = "select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,"
        oCommand2.CommandText += "sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( "
        oCommand2.CommandText += "select (case when month(oga02) = 1 then ogb12 else 0 end ) as t1,(case when month(oga02) = 2 then ogb12 else 0 end ) as t2,"
        oCommand2.CommandText += "(case when month(oga02) = 3 then ogb12 else 0 end ) as t3,(case when month(oga02) = 4 then ogb12 else 0 end ) as t4,"
        oCommand2.CommandText += "(case when month(oga02) = 5 then ogb12 else 0 end ) as t5,(case when month(oga02) = 6 then ogb12 else 0 end ) as t6,"
        oCommand2.CommandText += "(case when month(oga02) = 7 then ogb12 else 0 end ) as t7,(case when month(oga02) = 8 then ogb12 else 0 end ) as t8,"
        oCommand2.CommandText += "(case when month(oga02) = 9 then ogb12 else 0 end ) as t9,(case when month(oga02) = 10 then ogb12 else 0 end ) as t10,"
        oCommand2.CommandText += "(case when month(oga02) = 11 then ogb12 else 0 end ) as t11,(case when month(oga02) = 12 then ogb12 else 0 end ) as t12 from oga_file "
        oCommand2.CommandText += "left join ogb_file on oga01 = ogb01 left join ima_file on ogb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
        oCommand2.CommandText += "left join occ_file on tqa02 = occ02 and occacti = 'Y' left join gea_file on occ20 = gea01 "
        oCommand2.CommandText += "where ogapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and gea02 = '" & l_gea02 & "' and year(oga02) = " & syear & " and month(oga02) <= " & iTerm & " "
        oCommand2.CommandText += "union all "
        oCommand2.CommandText += "select (case when month(oha02) = 1 then ohb12 * -1 else 0 end ) as t1,"
        oCommand2.CommandText += "(case when month(oha02) = 2 then ohb12 * -1 else 0 end ) as t2,(case when month(oha02) = 3 then ohb12 * -1 else 0 end ) as t3,"
        oCommand2.CommandText += "(case when month(oha02) = 4 then ohb12 * -1 else 0 end ) as t4,(case when month(oha02) = 5 then ohb12 * -1 else 0 end ) as t5,"
        oCommand2.CommandText += "(case when month(oha02) = 6 then ohb12 * -1 else 0 end ) as t6,(case when month(oha02) = 7 then ohb12 * -1 else 0 end ) as t7,"
        oCommand2.CommandText += "(case when month(oha02) = 8 then ohb12 * -1 else 0 end ) as t8,(case when month(oha02) = 9 then ohb12 * -1 else 0 end ) as t9,"
        oCommand2.CommandText += "(case when month(oha02) = 10 then ohb12 * -1 else 0 end ) as t10,(case when month(oha02) = 11 then ohb12 * -1 else 0 end ) as t11,"
        oCommand2.CommandText += "(case when month(oha02) = 12 then ohb12 * -1 else 0 end ) as t12  from oha_file left join ohb_file on oha01 = ohb01  left join ima_file on ohb04 = ima01 "
        oCommand2.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join occ_file on tqa02 = occ02 and occacti = 'Y' left join gea_file on occ20 = gea01 where ohapost = 'Y' and ima06 = '103' and ima01 not like 'S%' and gea02  = '" & l_gea02 & "' and year(oha02) = " & syear & " and month(oha02) <= " & iTerm & " ) "
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                For i As Int16 = 1 To iTerm Step 1
                    Ws.Cells(LineZ, i + 3) = oReader2.Item(i - 1)
                Next
                Dim tColumn As String = String.Empty
                Select Case tMonth
                    Case 1
                        tColumn = "D"
                    Case 2
                        tColumn = "E"
                    Case 3
                        tColumn = "F"
                    Case 4
                        tColumn = "G"
                    Case 5
                        tColumn = "H"
                    Case 6
                        tColumn = "I"
                    Case 7
                        tColumn = "J"
                    Case 8
                        tColumn = "K"
                    Case 9
                        tColumn = "L"
                    Case 10
                        tColumn = "M"
                    Case 11
                        tColumn = "N"
                    Case 12
                        tColumn = "O"
                    Case Else
                        tColumn = "O"
                End Select
                Ws.Cells(LineZ, 16) = "=SUM(D" & LineZ & ":" & tColumn & LineZ & ")"
                LineZ += 1
            End While
        Else
            For i As Int16 = 1 To iTerm Step 1
                Ws.Cells(LineZ, i + 3) = 0
            Next

            LineZ += 1
        End If
        oReader2.Close()
    End Sub
End Class