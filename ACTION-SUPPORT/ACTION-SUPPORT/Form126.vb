Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.XlChartType
Imports Microsoft.Office.Core.MsoChartElementType
Imports Microsoft.Office.Core.MsoTriState
Public Class Form126
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
    Dim gDataBase As String = String.Empty
    Dim lMonth As Int16 = 0
    Dim DNP As String = String.Empty
    Dim tDate As Date
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form126_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        gDataBase = Me.ComboBox1.SelectedItem.ToString()
        Select Case gDataBase
            Case "DAC"
                oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
            Case "HAC"
                oConnection.ConnectionString = Module1.OpenOracleConnection("hkacttest")
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
        tDate = Convert.ToDateTime(tYear & "/" & tMonth & "/01")
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

        ' 第一頁
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        AdjustExcelFormat1()
        For i As Int16 = 1 To tMonth Step 1
            Ws.Cells(4, 2 + i) = GetD146103(0, i)
            Ws.Cells(5, 2 + i) = GetD146103USD(i)
            Ws.Cells(6, 2 + i) = GetExpense(i)
            Ws.Cells(7, 2 + i) = Gettc_ccj(i)
        Next

        ' 圖1
        Dim YE As Excel.Chart = Ws.Shapes.AddChart(xlLine, 15, 120, 850, 200).Chart
        oRng = Ws.Range("B3:N3,B8:N8")
        YE.SetSourceData(oRng, Microsoft.Office.Interop.Excel.XlRowCol.xlRows)
        YE.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        YE.ApplyLayout(5)
        YE.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Delete()

        ' 第二頁
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustExcelFormat2()

        For i As Int16 = 1 To tMonth Step 1
            Ws.Cells(3, 2 + i) = GetSales(i, "USD")
            Ws.Cells(4, 2 + i) = GetSalesRMB(i, "USD")
            Ws.Cells(5, 2 + i) = GetSales(i, "EUR")
            Ws.Cells(6, 2 + i) = GetSalesRMB(i, "EUR")
        Next
        ' 圖1
        Dim YF As Excel.Chart = Ws.Shapes.AddChart(xlLine, 15, 120, 600, 200).Chart
        oRng = Ws.Range("B2:N2,B8:N8")
        YF.SetSourceData(oRng, Microsoft.Office.Interop.Excel.XlRowCol.xlRows)
        'YF.SeriesCollection.NewSeries
        'YF.SeriesCollection(1).Name = "='Business Split'!$B$8"
        'YF.SeriesCollection(1).Values = "='Business Split'!$C$8:$N$8"
        YF.SeriesCollection.NewSeries()
        YF.SeriesCollection(2).Name = "='Business Split'!$B$9"
        YF.SeriesCollection(2).Values = "='Business Split'!$C$9:$N$9"
        YF.SeriesCollection(2).XValues = "='Business Split'!$C$2:$N$2"

        'oRng = Ws.Range("B2:N2,B8:N9")
        'YF.SetSourceData(oRng, Microsoft.Office.Interop.Excel.XlRowCol.xlRows)
        'YF.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        YF.ApplyLayout(5)
        YF.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Delete()
        YF.ChartTitle.Text = "USD/EURO Business Split"

        ' 圖2
        Dim YG As Excel.Chart = Ws.Shapes.AddChart(xlPie, 630, 120, 500, 200).Chart
        oRng = Ws.Range("B8:B9,O8:O9")
        YG.SetSourceData(oRng, Microsoft.Office.Interop.Excel.XlRowCol.xlColumns)
        YG.SeriesCollection(1).ApplyDataLabels()
        YG.SetElement(msoElementChartTitleAboveChart)
        YG.ChartTitle.Text = "YTD " & tYear
        'YF.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        'YF.ApplyLayout(5)
        'YF.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Delete()

        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        AdjustExcelFormat3()
        ' 美元匯率
        oCommand.CommandText = "select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "sum(t" & i & ") as t" & i & ","
        Next
        oCommand.CommandText += "1 from ( select "
        For i As Int16 = 1 To tMonth Step 1
            If i < 10 Then
                oCommand.CommandText += "(CASE WHEN SUBSTR(AZJ02,5,2) = '0" & i & "' THEN AZJ041 END) AS t" & i & ","
            Else
                oCommand.CommandText += "(CASE WHEN SUBSTR(AZJ02,5,2) = '" & i & "' THEN AZJ041 END) AS t" & i & ","
            End If
        Next
        oCommand.CommandText += "1 from azj_file where azj01 = 'USD' AND AZJ02 LIKE '" & tYear & "%' ORDER BY AZJ02 )"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 1 To tMonth Step 1
                    Ws.Cells(3, 2 + i) = oReader.Item(i - 1)
                Next
            End While
        End If
        oReader.Close()
        ' 歐元匯率
        oCommand.CommandText = "select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "sum(t" & i & ") as t" & i & ","
        Next
        oCommand.CommandText += "1 from ( select "
        For i As Int16 = 1 To tMonth Step 1
            If i < 10 Then
                oCommand.CommandText += "(CASE WHEN SUBSTR(AZJ02,5,2) = '0" & i & "' THEN AZJ041 END) AS t" & i & ","
            Else
                oCommand.CommandText += "(CASE WHEN SUBSTR(AZJ02,5,2) = '" & i & "' THEN AZJ041 END) AS t" & i & ","
            End If
        Next
        oCommand.CommandText += "1 from azj_file where azj01 = 'EUR' AND AZJ02 LIKE '" & tYear & "%' ORDER BY AZJ02 )"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 1 To tMonth Step 1
                    Ws.Cells(4, 2 + i) = oReader.Item(i - 1)
                Next
            End While
        End If
        oReader.Close()

        Ws.Cells(5, 3) = "=ROUND(C4/C3,4)"
        oRng = Ws.Range("C5", "C5")
        oRng.AutoFill(Destination:=Ws.Range("C5", Ws.Cells(5, 2 + tMonth)), Type:=xlFillDefault)

        ' 圖1
        Dim YB As Excel.Chart = Ws.Shapes.AddChart(xlLine, 15, 80, 650, 150).Chart
        oRng = Ws.Range("B2", "N3")
        YB.SetSourceData(oRng, Microsoft.Office.Interop.Excel.XlRowCol.xlRows)
        YB.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(0, 176, 240)
        YB.ApplyLayout(5)
        YB.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Delete()
        ' 圖2
        Dim YC As Excel.Chart = Ws.Shapes.AddChart(xlLine, 15, 240, 650, 150).Chart
        oRng = Ws.Range("B2:N2,B4:N4")
        YC.SetSourceData(oRng, Microsoft.Office.Interop.Excel.XlRowCol.xlRows)
        YC.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        YC.ApplyLayout(5)
        YC.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Delete()
        ' 圖3
        Dim YD As Excel.Chart = Ws.Shapes.AddChart(xlLine, 15, 400, 650, 150).Chart
        oRng = Ws.Range("B2:N2,B5:N5")
        YD.SetSourceData(oRng, Microsoft.Office.Interop.Excel.XlRowCol.xlRows)
        YD.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(146, 208, 80)
        YD.ApplyLayout(5)
        YD.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Delete()


        ' 第四頁 20180623
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(4)
        Ws.Activate()
        Ws.Name = "Selling Exp 汇总"
        AdjustExcelFormat4()
        oCommand.CommandText = "select aag01,aag02 from aag_file where aag01 like '6601%' and aag07 = 2 order by aag01"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("aag01")
                Ws.Cells(LineZ, 2) = oReader.Item("aag02")
                Ws.Cells(LineZ, 3) = Decimal.Round(GetLastYearSameMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 4) = Decimal.Round(GetLastMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 5) = Decimal.Round(GetThisYearSameMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 6) = GetThisYearSameMonthBudget(oReader.Item("aag01").ToString())
                Ws.Cells(LineZ, 7) = "=E" & LineZ & "-F" & LineZ
                Ws.Cells(LineZ, 8) = "=E" & LineZ & "-C" & LineZ
                Ws.Cells(LineZ, 9) = "=E" & LineZ & "-D" & LineZ
                Ws.Cells(LineZ, 10) = Decimal.Round(GetLastYearBeforeMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 11) = Decimal.Round(GetThisYearBeforeMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 12) = GetThisYearBeforeMonthBudget(oReader.Item("aag01").ToString())
                Ws.Cells(LineZ, 13) = "=K" & LineZ & "-L" & LineZ
                Ws.Cells(LineZ, 14) = "=K" & LineZ & "-J" & LineZ
                Ws.Cells(LineZ, 15) = Decimal.Round(GetLastYearNoMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 16) = "=Q" & LineZ & "-L" & LineZ & "+K" & LineZ
                Ws.Cells(LineZ, 17) = GetThisYearBudget(oReader.Item("aag01").ToString())
                LineZ += 1
            End While
        End If
        oReader.Close()
        Ws.Cells(LineZ, 2) = "Total Selling Exp"
        Ws.Cells(LineZ, 3) = "=SUM(C7:C" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 17)), Type:=xlFillDefault)
        ' 劃線
        oRng = Ws.Range("A7", Ws.Cells(LineZ, 17))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "

        '第五頁 
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(5)
        Ws.Name = "Selling Exp 部门明细"
        Ws.Activate()
        AdjustExcelFormat5()
        oCommand3.CommandText = "select distinct aao02 from ( select aao02 from aao_file where aao01 like '6601%' union all select tc_bud08 from tc_bud_file where tc_bud07 like '6601%' ) "
        oReader2 = oCommand3.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                DNP = oReader2.Item("aao02")
                oCommand.CommandText = "select aag01,aag02 from aag_file where aag01 like '6601%' and aag07 = 2 order by aag01"
                oReader = oCommand.ExecuteReader()
                If oReader.HasRows() Then
                    While oReader.Read()
                        Ws.Cells(LineZ, 1) = oReader.Item("aag01")
                        Ws.Cells(LineZ, 2) = oReader.Item("aag02")
                        Ws.Cells(LineZ, 3) = GetDepartNmae(DNP)
                        Ws.Cells(LineZ, 4) = Decimal.Round(GetLastYearSameMonth(oReader.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 5) = Decimal.Round(GetLastMonth(oReader.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearSameMonth(oReader.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 7) = GetThisYearSameMonthBudget(oReader.Item("aag01").ToString(), DNP)
                        Ws.Cells(LineZ, 8) = "=F" & LineZ & "-G" & LineZ
                        Ws.Cells(LineZ, 9) = "=F" & LineZ & "-D" & LineZ
                        Ws.Cells(LineZ, 10) = "=F" & LineZ & "-E" & LineZ
                        Ws.Cells(LineZ, 11) = Decimal.Round(GetLastYearBeforeMonth(oReader.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 12) = Decimal.Round(GetThisYearBeforeMonth(oReader.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 13) = GetThisYearBeforeMonthBudget(oReader.Item("aag01").ToString(), DNP)
                        Ws.Cells(LineZ, 14) = "=L" & LineZ & "-M" & LineZ
                        Ws.Cells(LineZ, 15) = "=L" & LineZ & "-K" & LineZ
                        Ws.Cells(LineZ, 16) = Decimal.Round(GetLastYearNoMonth(oReader.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 17) = "=R" & LineZ & "-M" & LineZ & "+L" & LineZ
                        Ws.Cells(LineZ, 18) = GetThisYearBudget(oReader.Item("aag01").ToString(), DNP)
                        LineZ += 1
                    End While
                End If
                oReader.Close()
            End While
        End If
        oReader2.Close()

        Ws.Cells(LineZ, 2) = "Total Selling Exp"
        Ws.Cells(LineZ, 4) = "=SUM(D7:D" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 4))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 18)), Type:=xlFillDefault)
        ' 劃線
        oRng = Ws.Range("A7", Ws.Cells(LineZ, 18))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "

        '第六頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(6)
        Ws.Name = "RD Exp 汇总"
        Ws.Activate()
        AdjustExcelFormat6()
        oCommand.CommandText = "select aag01,aag02 from aag_file where aag01 like '6604%' and aag07 = 2 order by aag01"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("aag01")
                Ws.Cells(LineZ, 2) = oReader.Item("aag02")
                Ws.Cells(LineZ, 3) = Decimal.Round(GetLastYearSameMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 4) = Decimal.Round(GetLastMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 5) = Decimal.Round(GetThisYearSameMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 6) = GetThisYearSameMonthBudget(oReader.Item("aag01").ToString())
                Ws.Cells(LineZ, 7) = "=E" & LineZ & "-F" & LineZ
                Ws.Cells(LineZ, 8) = "=E" & LineZ & "-C" & LineZ
                Ws.Cells(LineZ, 9) = "=E" & LineZ & "-D" & LineZ
                Ws.Cells(LineZ, 10) = Decimal.Round(GetLastYearBeforeMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 11) = Decimal.Round(GetThisYearBeforeMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 12) = GetThisYearBeforeMonthBudget(oReader.Item("aag01").ToString())
                Ws.Cells(LineZ, 13) = "=K" & LineZ & "-L" & LineZ
                Ws.Cells(LineZ, 14) = "=K" & LineZ & "-J" & LineZ
                Ws.Cells(LineZ, 15) = Decimal.Round(GetLastYearNoMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 16) = "=Q" & LineZ & "-L" & LineZ & "+K" & LineZ
                Ws.Cells(LineZ, 17) = GetThisYearBudget(oReader.Item("aag01").ToString())
                LineZ += 1
            End While
        End If
        oReader.Close()
        Ws.Cells(LineZ, 2) = "Total RD Exp"
        Ws.Cells(LineZ, 3) = "=SUM(C7:C" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 17)), Type:=xlFillDefault)
        ' 劃線
        oRng = Ws.Range("A7", Ws.Cells(LineZ, 17))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "

        '第七頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(7)
        Ws.Name = "RD Exp 部门明细"
        Ws.Activate()
        AdjustExcelFormat7()
        oCommand3.CommandText = "select distinct aao02 from ( select aao02 from aao_file where aao01 like '6604%' union all select tc_bud08 from tc_bud_file where tc_bud07 like '6604%' ) "
        oReader2 = oCommand3.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                DNP = oReader2.Item("aao02")
                oCommand.CommandText = "select aag01,aag02 from aag_file where aag01 like '6604%' and aag07 = 2 order by aag01"
                oReader = oCommand.ExecuteReader()
                If oReader.HasRows() Then
                    While oReader.Read()
                        Ws.Cells(LineZ, 1) = oReader.Item("aag01")
                        Ws.Cells(LineZ, 2) = oReader.Item("aag02")
                        Ws.Cells(LineZ, 3) = GetDepartNmae(DNP)
                        Ws.Cells(LineZ, 4) = Decimal.Round(GetLastYearSameMonth(oReader.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 5) = Decimal.Round(GetLastMonth(oReader.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearSameMonth(oReader.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 7) = GetThisYearSameMonthBudget(oReader.Item("aag01").ToString(), DNP)
                        Ws.Cells(LineZ, 8) = "=F" & LineZ & "-G" & LineZ
                        Ws.Cells(LineZ, 9) = "=F" & LineZ & "-D" & LineZ
                        Ws.Cells(LineZ, 10) = "=F" & LineZ & "-E" & LineZ
                        Ws.Cells(LineZ, 11) = Decimal.Round(GetLastYearBeforeMonth(oReader.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 12) = Decimal.Round(GetThisYearBeforeMonth(oReader.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 13) = GetThisYearBeforeMonthBudget(oReader.Item("aag01").ToString(), DNP)
                        Ws.Cells(LineZ, 14) = "=L" & LineZ & "-M" & LineZ
                        Ws.Cells(LineZ, 15) = "=L" & LineZ & "-K" & LineZ
                        Ws.Cells(LineZ, 16) = Decimal.Round(GetLastYearNoMonth(oReader.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 17) = "=R" & LineZ & "-M" & LineZ & "+L" & LineZ
                        Ws.Cells(LineZ, 18) = GetThisYearBudget(oReader.Item("aag01").ToString(), DNP)
                        LineZ += 1
                    End While
                End If
                oReader.Close()
            End While
        End If
        oReader2.Close()

        Ws.Cells(LineZ, 2) = "Total RD Exp"
        Ws.Cells(LineZ, 4) = "=SUM(D7:D" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 4))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 18)), Type:=xlFillDefault)
        ' 劃線
        oRng = Ws.Range("A7", Ws.Cells(LineZ, 18))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "

        '第八頁

        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(8)
        Ws.Name = "ADM Exp 汇总"
        Ws.Activate()
        AdjustExcelFormat8()
        oCommand.CommandText = "select aag01,aag02 from aag_file where aag01 like '6602%' and aag07 = 2 order by aag01"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("aag01")
                Ws.Cells(LineZ, 2) = oReader.Item("aag02")
                Ws.Cells(LineZ, 3) = Decimal.Round(GetLastYearSameMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 4) = Decimal.Round(GetLastMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 5) = Decimal.Round(GetThisYearSameMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 6) = GetThisYearSameMonthBudget(oReader.Item("aag01").ToString())
                Ws.Cells(LineZ, 7) = "=E" & LineZ & "-F" & LineZ
                Ws.Cells(LineZ, 8) = "=E" & LineZ & "-C" & LineZ
                Ws.Cells(LineZ, 9) = "=E" & LineZ & "-D" & LineZ
                Ws.Cells(LineZ, 10) = Decimal.Round(GetLastYearBeforeMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 11) = Decimal.Round(GetThisYearBeforeMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 12) = GetThisYearBeforeMonthBudget(oReader.Item("aag01").ToString())
                Ws.Cells(LineZ, 13) = "=K" & LineZ & "-L" & LineZ
                Ws.Cells(LineZ, 14) = "=K" & LineZ & "-J" & LineZ
                Ws.Cells(LineZ, 15) = Decimal.Round(GetLastYearNoMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 16) = "=Q" & LineZ & "-L" & LineZ & "+K" & LineZ
                Ws.Cells(LineZ, 17) = GetThisYearBudget(oReader.Item("aag01").ToString())
                LineZ += 1
            End While
        End If
        oReader.Close()
        Ws.Cells(LineZ, 2) = "Total ADM Exp"
        Ws.Cells(LineZ, 3) = "=SUM(C7:C" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 17)), Type:=xlFillDefault)
        ' 劃線
        oRng = Ws.Range("A7", Ws.Cells(LineZ, 17))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "

        '第九頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(9)
        Ws.Name = "ADM Exp 部门明细"
        Ws.Activate()
        AdjustExcelFormat9()
        oCommand3.CommandText = "select distinct aao02 from ( select aao02 from aao_file where aao01 like '6602%' union all select tc_bud08 from tc_bud_file where tc_bud07 like '6602%' ) "
        oReader2 = oCommand3.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                DNP = oReader2.Item("aao02")
                oCommand.CommandText = "select aag01,aag02 from aag_file where aag01 like '6602%' and aag07 = 2 order by aag01"
                oReader = oCommand.ExecuteReader()
                If oReader.HasRows() Then
                    While oReader.Read()
                        Ws.Cells(LineZ, 1) = oReader.Item("aag01")
                        Ws.Cells(LineZ, 2) = oReader.Item("aag02")
                        Ws.Cells(LineZ, 3) = GetDepartNmae(DNP)
                        Ws.Cells(LineZ, 4) = Decimal.Round(GetLastYearSameMonth(oReader.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 5) = Decimal.Round(GetLastMonth(oReader.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearSameMonth(oReader.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 7) = GetThisYearSameMonthBudget(oReader.Item("aag01").ToString(), DNP)
                        Ws.Cells(LineZ, 8) = "=F" & LineZ & "-G" & LineZ
                        Ws.Cells(LineZ, 9) = "=F" & LineZ & "-D" & LineZ
                        Ws.Cells(LineZ, 10) = "=F" & LineZ & "-E" & LineZ
                        Ws.Cells(LineZ, 11) = Decimal.Round(GetLastYearBeforeMonth(oReader.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 12) = Decimal.Round(GetThisYearBeforeMonth(oReader.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 13) = GetThisYearBeforeMonthBudget(oReader.Item("aag01").ToString(), DNP)
                        Ws.Cells(LineZ, 14) = "=L" & LineZ & "-M" & LineZ
                        Ws.Cells(LineZ, 15) = "=L" & LineZ & "-K" & LineZ
                        Ws.Cells(LineZ, 16) = Decimal.Round(GetLastYearNoMonth(oReader.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 17) = "=R" & LineZ & "-M" & LineZ & "+L" & LineZ
                        Ws.Cells(LineZ, 18) = GetThisYearBudget(oReader.Item("aag01").ToString(), DNP)
                        LineZ += 1
                    End While
                End If
                oReader.Close()
            End While
        End If
        oReader2.Close()

        Ws.Cells(LineZ, 2) = "Total ADM Exp"
        Ws.Cells(LineZ, 4) = "=SUM(D7:D" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 4))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 18)), Type:=xlFillDefault)
        ' 劃線
        oRng = Ws.Range("A7", Ws.Cells(LineZ, 18))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "

    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "accounting reports_" & gDataBase
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
        Ws.Name = "Index"
        Ws.Columns.HorizontalAlignment = xlCenter
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10
        oRng = Ws.Range("B1", "J2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 1.89
        oRng = Ws.Range("B2", "B2")
        oRng.EntireColumn.ColumnWidth = 18.11
        oRng.EntireColumn.HorizontalAlignment = xlCenter
        oRng.EntireRow.HorizontalAlignment = xlCenter
        oRng = Ws.Range("C1", "O1")
        oRng.EntireColumn.ColumnWidth = 9.89

        Ws.Cells(1, 2) = "Index" & "(Total FG Output/used hour from HR record)"
        Ws.Cells(2, 15) = "币别：USD"
        For i As Int16 = 1 To 12 Step 1
            Ws.Cells(3, 2 + i) = tYear & "-" & i
        Next
        Ws.Cells(3, 15) = "Total"
        Ws.Cells(4, 2) = "FG Qty"
        Ws.Cells(4, 15) = "=SUM(C4:N4)"
        Ws.Cells(5, 2) = "value output（USD）"
        Ws.Cells(5, 15) = "=SUM(C5:N5)"
        Ws.Cells(6, 2) = "Direct labor Cost"
        Ws.Cells(6, 15) = "=SUM(C6:N6)"
        Ws.Cells(7, 2) = "work hours"
        Ws.Cells(7, 15) = "=SUM(C7:N7)"
        Ws.Cells(8, 2) = "Output Index"
        Ws.Cells(8, 3) = "=C5/C7"
        oRng = Ws.Range("C8", "C8")
        oRng.AutoFill(Destination:=Ws.Range("C8", "O8"), Type:=xlFillDefault)

        ' 上色
        oRng = Ws.Range("B8", "O8")
        oRng.Interior.Color = Color.FromArgb(220, 230, 241)
        oRng = Ws.Range("O3", "O7")
        oRng.Interior.Color = Color.FromArgb(220, 230, 241)

        oRng = Ws.Range("B3", "O8")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("C4", "O7")
        oRng.NumberFormat = "#,##0"

        oRng = Ws.Range("C8", "O8")
        oRng.NumberFormat = "0.00"
        LineZ = 4
    End Sub
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "Business Split"
        Ws.Columns.HorizontalAlignment = xlCenter
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10
        oRng = Ws.Range("B1", "J1")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 1.89
        oRng = Ws.Range("B2", "B2")
        oRng.EntireColumn.ColumnWidth = 13.11
        oRng.EntireColumn.HorizontalAlignment = xlCenter
        oRng.EntireRow.HorizontalAlignment = xlCenter
        oRng = Ws.Range("C1", "O1")
        oRng.EntireColumn.ColumnWidth = 13.56

        Ws.Cells(1, 2) = "USD/EUR Business Split"
        Ws.Cells(2, 2) = "currency"
        For i As Int16 = 1 To 12 Step 1
            If i < 10 Then
                Ws.Cells(2, 2 + i) = tYear & "-0" & i
            Else
                Ws.Cells(2, 2 + i) = tYear & "-" & i
            End If
        Next
        oRng = Ws.Range("B3", "B4")
        oRng.Merge()
        oRng = Ws.Range("B5", "B6")
        oRng.Merge()
        Ws.Cells(2, 15) = "YTD " & tYear
        Ws.Cells(3, 2) = "USD"
        Ws.Cells(3, 15) = "=SUM(C3:N3)"
        Ws.Cells(4, 15) = "=SUM(C4:N4)"
        Ws.Cells(5, 15) = "=SUM(C5:N5)"
        Ws.Cells(6, 15) = "=SUM(C6:N6)"
        Ws.Cells(5, 2) = "EUR"
        Ws.Cells(7, 2) = "Total"
        Ws.Cells(7, 3) = "=C4+C6"
        oRng = Ws.Range("C7", "C7")
        oRng.AutoFill(Destination:=Ws.Range("C7", "N7"), Type:=xlFillDefault)
        Ws.Cells(7, 15) = "=SUM(C7:N7)"
        Ws.Cells(8, 2) = "USD %"
        Ws.Cells(8, 3) = "=IF(OR(C3="""",C5=""""),""0%"",C4/C7)"
        oRng = Ws.Range("C8", "C8")
        oRng.AutoFill(Destination:=Ws.Range("C8", "N8"), Type:=xlFillDefault)
        Ws.Cells(8, 15) = "=O4/O7"
        Ws.Cells(9, 2) = "EUR %"
        Ws.Cells(9, 3) = "=IF(OR(C3="""",C5=""""),""0%"",C6/C7)"
        oRng = Ws.Range("C9", "C9")
        oRng.AutoFill(Destination:=Ws.Range("C9", "N9"), Type:=xlFillDefault)
        Ws.Cells(9, 15) = "=O6/O7"
        ' 格式
        oRng = Ws.Range("C3", "O3")
        oRng.NumberFormatLocal = "$#,##0_);[红色]($#,##0)"
        oRng = Ws.Range("C4", "O4")
        oRng.NumberFormatLocal = "¥#,##0_);[红色](¥#,##0)"
        oRng = Ws.Range("C5", "O5")
        oRng.NumberFormatLocal = "[$€-x-euro2] #,##0;[红色][$€-x-euro2] #,##0"
        oRng = Ws.Range("C6", "O7")
        oRng.NumberFormatLocal = "¥#,##0_);[红色](¥#,##0)"
        oRng = Ws.Range("C8", "O9")
        oRng.NumberFormatLocal = "0%"

        oRng = Ws.Range("B2", "O9")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        'oRng = Ws.Range("C3", "N5")
        'oRng.NumberFormatLocal = "#,##0.0000_);[红色](#,##0.0000)"

        LineZ = 3
    End Sub
    Private Sub AdjustExcelFormat3()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "Exchange Rate Chart"
        Ws.Columns.HorizontalAlignment = xlCenter
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10
        oRng = Ws.Range("B1", "N1")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 1.89
        oRng = Ws.Range("B2", "B2")
        oRng.EntireColumn.ColumnWidth = 15
        oRng.EntireColumn.HorizontalAlignment = xlCenter
        oRng.EntireRow.HorizontalAlignment = xlCenter
        oRng = Ws.Range("C1", "N1")
        oRng.EntireColumn.ColumnWidth = 7.89

        Ws.Cells(1, 2) = tYear & "Exchange Rate Chart"
        Ws.Cells(2, 2) = "Exchange Rate"
        For i As Int16 = 1 To 12 Step 1
            If i < 10 Then
                Ws.Cells(2, 2 + i) = tYear & "-0" & i
            Else
                Ws.Cells(2, 2 + i) = tYear & "-" & i
            End If
        Next
        Ws.Cells(3, 2) = "YTD " & tYear
        Ws.Cells(3, 2) = "USD:RMB"
        Ws.Cells(4, 2) = "EUR:RMB"
        Ws.Cells(5, 2) = "EUR：USD"

        oRng = Ws.Range("B2", "N5")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("C3", "N5")
        oRng.NumberFormatLocal = "#,##0.0000_);[红色](#,##0.0000)"

        LineZ = 3
    End Sub
    Private Sub AdjustExcelFormat4()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 10.44
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 60
        oRng = Ws.Range("B3", "Q3")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        'oRng.Interior.Color = Color.FromArgb(169, 209, 141)
        Ws.Cells(3, 2) = "Selling Exp. By account"
        Ws.Cells(4, 2) = "USD"
        oRng = Ws.Range("B5", "B5")
        oRng.NumberFormatLocal = "mmm-yy"
        oRng.HorizontalAlignment = xlLeft
        Ws.Cells(5, 2) = tDate
        Select Case gDataBase
            Case "DAC"
                Ws.Cells(6, 2) = "Dongguan Action Composites LTD Co."
                Dim TYM1 As String = String.Empty
                If tMonth < 10 Then
                    TYM1 = tYear & "0" & tMonth
                Else
                    TYM1 = tYear & tMonth
                End If
                oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & TYM1 & "'"
                ExchangeRate1 = oCommand.ExecuteScalar()

            Case "HAC"
                Ws.Cells(6, 2) = "Action Composite Technology Limited"
                ExchangeRate1 = 1
            Case "action_bvi"
                Ws.Cells(6, 2) = "Action Composites International Limited"
                ExchangeRate1 = 1
        End Select
        oRng = Ws.Range("C4", "E5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 3) = "Actual"
        Ws.Cells(6, 3) = tDate.AddYears(-1)
        Ws.Cells(6, 4) = tDate.AddMonths(-1)
        Ws.Cells(6, 5) = tDate
        Ws.Cells(6, 6) = tDate
        oRng = Ws.Range("C6", "F6")
        oRng.NumberFormatLocal = "mmm-yy"
        oRng = Ws.Range("F4", "F5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 6) = "Budget"
        oRng = Ws.Range("G4", "I4")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 7) = "Variance" '& Chr(10) & "Act& Bud"
        Ws.Cells(5, 7) = "Act & But"
        Ws.Cells(5, 8) = "year-on-year"
        Ws.Cells(5, 9) = "Month-on-month"
        Ws.Cells(6, 7) = "USD"
        Ws.Cells(6, 8) = "USD"
        Ws.Cells(6, 9) = "USD"
        'oRng = Ws.Range("C4", "I6")
        'oRng.Interior.Color = Color.FromArgb(255, 218, 101)
        oRng = Ws.Range("J4", "K5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 10) = "Actual"
        Ws.Cells(6, 10) = "YTD " & pYear
        Ws.Cells(6, 11) = "YTD " & tYear
        oRng = Ws.Range("L4", "L5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 12) = "Budget"
        Ws.Cells(6, 12) = "YTD " & tYear
        oRng = Ws.Range("M4", "N4")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 13) = "Variance" '& Chr(10) & "Act& Bud"
        Ws.Cells(5, 13) = "Act & But"
        Ws.Cells(5, 14) = "year-on-year"
        Ws.Cells(6, 13) = "USD"
        Ws.Cells(6, 14) = "USD"
        'oRng = Ws.Range("J4", "N6")
        'oRng.Interior.Color = Color.FromArgb(156, 195, 230)
        oRng = Ws.Range("O4", "O5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 15) = "Actual"
        Ws.Cells(6, 15) = "Y" & pYear
        oRng = Ws.Range("P4", "P5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 16) = "Rollling" & Chr(10) & "Forecast"
        Ws.Cells(6, 16) = "Y" & tYear
        oRng = Ws.Range("Q4", "Q5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 17) = "Budget"
        Ws.Cells(6, 17) = "Y" & tYear
        'oRng = Ws.Range("O4", "Q6")
        'oRng.Interior.Color = Color.FromArgb(146, 208, 80)
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        ' 劃線
        oRng = Ws.Range("B3", "Q6")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("C6", "Q6")
        oRng.HorizontalAlignment = xlRight
        LineZ = 7
    End Sub
    Private Sub AdjustExcelFormat5()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 10.44
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 60
        oRng = Ws.Range("B3", "R3")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        'oRng.Interior.Color = Color.FromArgb(169, 209, 141)
        Ws.Cells(3, 2) = "Selling Exp. By account"
        Ws.Cells(4, 2) = "USD"
        oRng = Ws.Range("B5", "B5")
        oRng.NumberFormatLocal = "mmm-yy"
        oRng.HorizontalAlignment = xlLeft
        Ws.Cells(5, 2) = tDate
        Select Case gDataBase
            Case "DAC"
                Ws.Cells(6, 2) = "Dongguan Action Composites LTD Co."
            Case "HAC"
                Ws.Cells(6, 2) = "Action Composite Technology Limited"
            Case "action_bvi"
                Ws.Cells(6, 2) = "Action Composites International Limited"
        End Select
        oRng = Ws.Range("C4", "C6")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 3) = "Cost" & Chr(10) & "Center"
        oRng = Ws.Range("D4", "F5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 4) = "Actual"
        Ws.Cells(6, 4) = tDate.AddYears(-1)
        Ws.Cells(6, 5) = tDate.AddMonths(-1)
        Ws.Cells(6, 6) = tDate
        Ws.Cells(6, 7) = tDate
        oRng = Ws.Range("D6", "G6")
        oRng.NumberFormatLocal = "mmm-yy"
        oRng = Ws.Range("G4", "G5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 7) = "Budget"
        oRng = Ws.Range("H4", "J4")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 8) = "Variance" '& Chr(10) & "Act& Bud"
        Ws.Cells(5, 8) = "Act & But"
        Ws.Cells(5, 9) = "year-on-year"
        Ws.Cells(5, 10) = "Month-on-month"
        Ws.Cells(6, 8) = "USD"
        Ws.Cells(6, 9) = "USD"
        Ws.Cells(6, 10) = "USD"
        'oRng = Ws.Range("D4", "J6")
        'oRng.Interior.Color = Color.FromArgb(255, 218, 101)
        oRng = Ws.Range("K4", "L5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 11) = "Actual"
        Ws.Cells(6, 11) = "YTD " & pYear
        Ws.Cells(6, 12) = "YTD " & tYear
        oRng = Ws.Range("M4", "M5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 13) = "Budget"
        Ws.Cells(6, 13) = "YTD " & tYear
        oRng = Ws.Range("N4", "O4")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 14) = "Variance" '& Chr(10) & "Act& Bud"
        Ws.Cells(5, 14) = "Act & But"
        Ws.Cells(5, 15) = "year-on-year"
        Ws.Cells(6, 14) = "USD"
        Ws.Cells(6, 15) = "USD"
        'oRng = Ws.Range("K4", "O6")
        'oRng.Interior.Color = Color.FromArgb(156, 195, 230)
        oRng = Ws.Range("P4", "P5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 16) = "Actual"
        Ws.Cells(6, 16) = "Y" & pYear
        oRng = Ws.Range("Q4", "Q5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 17) = "Rollling" & Chr(10) & "Forecast"
        Ws.Cells(6, 17) = "Y" & tYear
        oRng = Ws.Range("R4", "R5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 18) = "Budget"
        Ws.Cells(6, 18) = tYear
        'oRng = Ws.Range("M4", "O6")
        'oRng.Interior.Color = Color.FromArgb(146, 208, 80)
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        ' 劃線
        oRng = Ws.Range("B3", "R6")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng = Ws.Range("D6", "R6")
        oRng.HorizontalAlignment = xlRight
        LineZ = 7
    End Sub
    Private Sub AdjustExcelFormat6()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 10.44
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 60
        oRng = Ws.Range("B3", "Q3")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        'oRng.Interior.Color = Color.FromArgb(169, 209, 141)
        Ws.Cells(3, 2) = "RD Exp. By account"
        Ws.Cells(4, 2) = "USD"
        oRng = Ws.Range("B5", "B5")
        oRng.NumberFormatLocal = "mmm-yy"
        oRng.HorizontalAlignment = xlLeft
        Ws.Cells(5, 2) = tDate
        Select Case gDataBase
            Case "DAC"
                Ws.Cells(6, 2) = "Dongguan Action Composites LTD Co."
                Dim TYM1 As String = String.Empty
                If tMonth < 10 Then
                    TYM1 = tYear & "0" & tMonth
                Else
                    TYM1 = tYear & tMonth
                End If
                oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & TYM1 & "'"
                ExchangeRate1 = oCommand.ExecuteScalar()

            Case "HAC"
                Ws.Cells(6, 2) = "Action Composite Technology Limited"
                ExchangeRate1 = 1
            Case "action_bvi"
                Ws.Cells(6, 2) = "Action Composites International Limited"
                ExchangeRate1 = 1
        End Select
        oRng = Ws.Range("C4", "E5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 3) = "Actual"
        Ws.Cells(6, 3) = tDate.AddYears(-1)
        Ws.Cells(6, 4) = tDate.AddMonths(-1)
        Ws.Cells(6, 5) = tDate
        Ws.Cells(6, 6) = tDate
        oRng = Ws.Range("C6", "F6")
        oRng.NumberFormatLocal = "mmm-yy"
        oRng = Ws.Range("F4", "F5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 6) = "Budget"
        oRng = Ws.Range("G4", "I4")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 7) = "Variance" '& Chr(10) & "Act& Bud"
        Ws.Cells(5, 7) = "Act & But"
        Ws.Cells(5, 8) = "year-on-year"
        Ws.Cells(5, 9) = "Month-on-month"
        Ws.Cells(6, 7) = "USD"
        Ws.Cells(6, 8) = "USD"
        Ws.Cells(6, 9) = "USD"
        'oRng = Ws.Range("C4", "I6")
        'oRng.Interior.Color = Color.FromArgb(255, 218, 101)
        oRng = Ws.Range("J4", "K5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 10) = "Actual"
        Ws.Cells(6, 10) = "YTD " & pYear
        Ws.Cells(6, 11) = "YTD " & tYear
        oRng = Ws.Range("L4", "L5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 12) = "Budget"
        Ws.Cells(6, 12) = "YTD " & tYear
        oRng = Ws.Range("M4", "N4")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 13) = "Variance" '& Chr(10) & "Act& Bud"
        Ws.Cells(5, 13) = "Act & But"
        Ws.Cells(5, 14) = "year-on-year"
        Ws.Cells(6, 13) = "USD"
        Ws.Cells(6, 14) = "USD"
        'oRng = Ws.Range("J4", "N6")
        'oRng.Interior.Color = Color.FromArgb(156, 195, 230)
        oRng = Ws.Range("O4", "O5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 15) = "Actual"
        Ws.Cells(6, 15) = "Y" & pYear
        oRng = Ws.Range("P4", "P5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 16) = "Rollling" & Chr(10) & "Forecast"
        Ws.Cells(6, 16) = "Y" & tYear
        oRng = Ws.Range("Q4", "Q5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 17) = "Budget"
        Ws.Cells(6, 17) = "Y" & tYear
        'oRng = Ws.Range("O4", "Q6")
        'oRng.Interior.Color = Color.FromArgb(146, 208, 80)
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        ' 劃線
        oRng = Ws.Range("B3", "Q6")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("C6", "Q6")
        oRng.HorizontalAlignment = xlRight
        LineZ = 7
    End Sub
    Private Sub AdjustExcelFormat7()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 10.44
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 60
        oRng = Ws.Range("B3", "R3")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        'oRng.Interior.Color = Color.FromArgb(169, 209, 141)
        Ws.Cells(3, 2) = "RD Exp. By account"
        Ws.Cells(4, 2) = "USD"
        oRng = Ws.Range("B5", "B5")
        oRng.NumberFormatLocal = "mmm-yy"
        oRng.HorizontalAlignment = xlLeft
        Ws.Cells(5, 2) = tDate
        Select Case gDataBase
            Case "DAC"
                Ws.Cells(6, 2) = "Dongguan Action Composites LTD Co."
            Case "HAC"
                Ws.Cells(6, 2) = "Action Composite Technology Limited"
            Case "action_bvi"
                Ws.Cells(6, 2) = "Action Composites International Limited"
        End Select
        oRng = Ws.Range("C4", "C6")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 3) = "Cost" & Chr(10) & "Center"
        oRng = Ws.Range("D4", "F5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 4) = "Actual"
        Ws.Cells(6, 4) = tDate.AddYears(-1)
        Ws.Cells(6, 5) = tDate.AddMonths(-1)
        Ws.Cells(6, 6) = tDate
        Ws.Cells(6, 7) = tDate
        oRng = Ws.Range("D6", "G6")
        oRng.NumberFormatLocal = "mmm-yy"
        oRng = Ws.Range("G4", "G5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 7) = "Budget"
        oRng = Ws.Range("H4", "J4")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 8) = "Variance" '& Chr(10) & "Act& Bud"
        Ws.Cells(5, 8) = "Act & But"
        Ws.Cells(5, 9) = "year-on-year"
        Ws.Cells(5, 10) = "Month-on-month"
        Ws.Cells(6, 8) = "USD"
        Ws.Cells(6, 9) = "USD"
        Ws.Cells(6, 10) = "USD"
        'oRng = Ws.Range("D4", "J6")
        'oRng.Interior.Color = Color.FromArgb(255, 218, 101)
        oRng = Ws.Range("K4", "L5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 11) = "Actual"
        Ws.Cells(6, 11) = "YTD " & pYear
        Ws.Cells(6, 12) = "YTD " & tYear
        oRng = Ws.Range("M4", "M5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 13) = "Budget"
        Ws.Cells(6, 13) = "YTD " & tYear
        oRng = Ws.Range("N4", "O4")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 14) = "Variance" '& Chr(10) & "Act& Bud"
        Ws.Cells(5, 14) = "Act & But"
        Ws.Cells(5, 15) = "year-on-year"
        Ws.Cells(6, 14) = "USD"
        Ws.Cells(6, 15) = "USD"
        'oRng = Ws.Range("K4", "O6")
        'oRng.Interior.Color = Color.FromArgb(156, 195, 230)
        oRng = Ws.Range("P4", "P5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 16) = "Actual"
        Ws.Cells(6, 16) = "Y" & pYear
        oRng = Ws.Range("Q4", "Q5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 17) = "Rollling" & Chr(10) & "Forecast"
        Ws.Cells(6, 17) = "Y" & tYear
        oRng = Ws.Range("R4", "R5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 18) = "Budget"
        Ws.Cells(6, 18) = tYear
        'oRng = Ws.Range("M4", "O6")
        'oRng.Interior.Color = Color.FromArgb(146, 208, 80)
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        ' 劃線
        oRng = Ws.Range("B3", "R6")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng = Ws.Range("D6", "R6")
        oRng.HorizontalAlignment = xlRight
        LineZ = 7
    End Sub
    Private Sub AdjustExcelFormat8()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 10.44
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 60
        oRng = Ws.Range("B3", "Q3")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        'oRng.Interior.Color = Color.FromArgb(169, 209, 141)
        Ws.Cells(3, 2) = "ADM Exp. By account"
        Ws.Cells(4, 2) = "USD"
        oRng = Ws.Range("B5", "B5")
        oRng.NumberFormatLocal = "mmm-yy"
        oRng.HorizontalAlignment = xlLeft
        Ws.Cells(5, 2) = tDate
        Select Case gDataBase
            Case "DAC"
                Ws.Cells(6, 2) = "Dongguan Action Composites LTD Co."
                Dim TYM1 As String = String.Empty
                If tMonth < 10 Then
                    TYM1 = tYear & "0" & tMonth
                Else
                    TYM1 = tYear & tMonth
                End If
                oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & TYM1 & "'"
                ExchangeRate1 = oCommand.ExecuteScalar()

            Case "HAC"
                Ws.Cells(6, 2) = "Action Composite Technology Limited"
                ExchangeRate1 = 1
            Case "action_bvi"
                Ws.Cells(6, 2) = "Action Composites International Limited"
                ExchangeRate1 = 1
        End Select
        oRng = Ws.Range("C4", "E5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 3) = "Actual"
        Ws.Cells(6, 3) = tDate.AddYears(-1)
        Ws.Cells(6, 4) = tDate.AddMonths(-1)
        Ws.Cells(6, 5) = tDate
        Ws.Cells(6, 6) = tDate
        oRng = Ws.Range("C6", "F6")
        oRng.NumberFormatLocal = "mmm-yy"
        oRng = Ws.Range("F4", "F5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 6) = "Budget"
        oRng = Ws.Range("G4", "I4")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 7) = "Variance" '& Chr(10) & "Act& Bud"
        Ws.Cells(5, 7) = "Act & But"
        Ws.Cells(5, 8) = "year-on-year"
        Ws.Cells(5, 9) = "Month-on-month"
        Ws.Cells(6, 7) = "USD"
        Ws.Cells(6, 8) = "USD"
        Ws.Cells(6, 9) = "USD"
        'oRng = Ws.Range("C4", "I6")
        'oRng.Interior.Color = Color.FromArgb(255, 218, 101)
        oRng = Ws.Range("J4", "K5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 10) = "Actual"
        Ws.Cells(6, 10) = "YTD " & pYear
        Ws.Cells(6, 11) = "YTD " & tYear
        oRng = Ws.Range("L4", "L5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 12) = "Budget"
        Ws.Cells(6, 12) = "YTD " & tYear
        oRng = Ws.Range("M4", "N4")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 13) = "Variance" '& Chr(10) & "Act& Bud"
        Ws.Cells(5, 13) = "Act & But"
        Ws.Cells(5, 14) = "year-on-year"
        Ws.Cells(6, 13) = "USD"
        Ws.Cells(6, 14) = "USD"
        'oRng = Ws.Range("J4", "N6")
        'oRng.Interior.Color = Color.FromArgb(156, 195, 230)
        oRng = Ws.Range("O4", "O5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 15) = "Actual"
        Ws.Cells(6, 15) = "Y" & pYear
        oRng = Ws.Range("P4", "P5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 16) = "Rollling" & Chr(10) & "Forecast"
        Ws.Cells(6, 16) = "Y" & tYear
        oRng = Ws.Range("Q4", "Q5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 17) = "Budget"
        Ws.Cells(6, 17) = "Y" & tYear
        'oRng = Ws.Range("O4", "Q6")
        'oRng.Interior.Color = Color.FromArgb(146, 208, 80)
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        ' 劃線
        oRng = Ws.Range("B3", "Q6")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("C6", "Q6")
        oRng.HorizontalAlignment = xlRight
        LineZ = 7
    End Sub
    Private Sub AdjustExcelFormat9()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 10.44
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 60
        oRng = Ws.Range("B3", "R3")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        'oRng.Interior.Color = Color.FromArgb(169, 209, 141)
        Ws.Cells(3, 2) = "ADM Exp. By account"
        Ws.Cells(4, 2) = "USD"
        oRng = Ws.Range("B5", "B5")
        oRng.NumberFormatLocal = "mmm-yy"
        oRng.HorizontalAlignment = xlLeft
        Ws.Cells(5, 2) = tDate
        Select Case gDataBase
            Case "DAC"
                Ws.Cells(6, 2) = "Dongguan Action Composites LTD Co."
            Case "HAC"
                Ws.Cells(6, 2) = "Action Composite Technology Limited"
            Case "action_bvi"
                Ws.Cells(6, 2) = "Action Composites International Limited"
        End Select
        oRng = Ws.Range("C4", "C6")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 3) = "Cost" & Chr(10) & "Center"
        oRng = Ws.Range("D4", "F5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 4) = "Actual"
        Ws.Cells(6, 4) = tDate.AddYears(-1)
        Ws.Cells(6, 5) = tDate.AddMonths(-1)
        Ws.Cells(6, 6) = tDate
        Ws.Cells(6, 7) = tDate
        oRng = Ws.Range("D6", "G6")
        oRng.NumberFormatLocal = "mmm-yy"
        oRng = Ws.Range("G4", "G5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 7) = "Budget"
        oRng = Ws.Range("H4", "J4")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 8) = "Variance" '& Chr(10) & "Act& Bud"
        Ws.Cells(5, 8) = "Act & But"
        Ws.Cells(5, 9) = "year-on-year"
        Ws.Cells(5, 10) = "Month-on-month"
        Ws.Cells(6, 8) = "USD"
        Ws.Cells(6, 9) = "USD"
        Ws.Cells(6, 10) = "USD"
        'oRng = Ws.Range("D4", "J6")
        'oRng.Interior.Color = Color.FromArgb(255, 218, 101)
        oRng = Ws.Range("K4", "L5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 11) = "Actual"
        Ws.Cells(6, 11) = "YTD " & pYear
        Ws.Cells(6, 12) = "YTD " & tYear
        oRng = Ws.Range("M4", "M5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 13) = "Budget"
        Ws.Cells(6, 13) = "YTD " & tYear
        oRng = Ws.Range("N4", "O4")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 14) = "Variance" '& Chr(10) & "Act& Bud"
        Ws.Cells(5, 14) = "Act & But"
        Ws.Cells(5, 15) = "year-on-year"
        Ws.Cells(6, 14) = "USD"
        Ws.Cells(6, 15) = "USD"
        'oRng = Ws.Range("K4", "O6")
        'oRng.Interior.Color = Color.FromArgb(156, 195, 230)
        oRng = Ws.Range("P4", "P5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 16) = "Actual"
        Ws.Cells(6, 16) = "Y" & pYear
        oRng = Ws.Range("Q4", "Q5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 17) = "Rollling" & Chr(10) & "Forecast"
        Ws.Cells(6, 17) = "Y" & tYear
        oRng = Ws.Range("R4", "R5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 18) = "Budget"
        Ws.Cells(6, 18) = tYear
        'oRng = Ws.Range("M4", "O6")
        'oRng.Interior.Color = Color.FromArgb(146, 208, 80)
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        ' 劃線
        oRng = Ws.Range("B3", "R6")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng = Ws.Range("D6", "R6")
        oRng.HorizontalAlignment = xlRight
        LineZ = 7
    End Sub
    Private Function GetD146103(ByVal eType As Int16, ByVal sMonth As Int16)
        Dim S1 As Date = Convert.ToDateTime(tYear & "/" & sMonth & "/01")
        Dim S2 As Date = S1.AddMonths(1).AddDays(-1)
        oCommand2.CommandText = "select nvl(sum(tlf10 * tlf12 * tlf907),0)  from tlf_file where tlf13 = 'aimt324' and tlf902 = 'D146103' and tlf06 between to_date('" & S1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & S2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        Dim SSA As Decimal = oCommand2.ExecuteScalar()
        Return SSA
    End Function
    Private Function GetD146103USD(ByVal sMonth As Int16)
        Dim S1 As Date = Convert.ToDateTime(tYear & "/" & sMonth & "/01")
        Dim S2 As Date = S1.AddMonths(1).AddDays(-1)
        oCommand2.CommandText = "select nvl(round(sum(tlf10 * tlf12 * (stb07+stb08+stb09+stb09a) /azj041),2),0) from tlf_file, stb_file, azj_file "
        oCommand2.CommandText += "where tlf01 = stb01 and azj01 ='USD' AND azj02 = stb02 || case when length(stb03) = 1 then '0' end || stb03 "
        oCommand2.CommandText += "and tlf13 = 'aimt324' and tlf902 = 'D146103' and tlf907 = 1 and tlf06 between to_date('" & S1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & S2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and year(tlf06) =stb02 and month(tlf06) = stb03"
        Dim ASS As Decimal = oCommand2.ExecuteScalar()
        Return ASS
    End Function
    Private Function Gettc_ccj(ByVal sMonth As Int16)
        oCommand2.CommandText = "select nvl(round(sum(tc_ccj04)/60,2),0) from tc_ccj_file "
        oCommand2.CommandText += "where tc_ccj01 = " & tYear & " and tc_ccj02 = " & sMonth
        Dim ASS As Decimal = oCommand2.ExecuteScalar()
        Return ASS
    End Function
    Private Function GetExpense(ByVal sMonth As Int16)
        oCommand2.CommandText = "select nvl(round(sum(aao05-aao06)/azj041,2),0) from aao_file, azj_file "
        oCommand2.CommandText += "where aao01 = '510121' and aao03 = " & tYear & " and aao02 <> 'D9999' and aao04 = " & sMonth & " and azj01 ='USD' AND azj02 = aao03 || case when length(aao04) = 1 then '0' end || aao04 group by azj041"
        Dim ASS As Decimal = oCommand2.ExecuteScalar()
        Return ASS
    End Function
    Private Function GetSales(ByVal sMonth As Int16, ByVal Tcurrency As String)
        oCommand2.CommandText = "SELECT sum(t1) from ( select nvl(sum(ogb14),0) as t1 from hkacttest.ogb_file,hkacttest.oga_file where ogb01 = oga01 and ogapost = 'Y' and year(oga02) = " & tYear & " and month(oga02) = " & sMonth & " and oga23 = '" & Tcurrency & "' "
        oCommand2.CommandText += "union all "
        oCommand2.CommandText += "select nvl(sum(ohb14 * -1),0) from hkacttest.ohb_file,hkacttest.oha_file where ohb01 = oha01 and ohapost = 'Y' and year(oha02) = " & tYear & " and month(oha02) = " & sMonth & " and oha23 = '" & Tcurrency & "' )"
        Dim Ass As Decimal = oCommand2.ExecuteScalar()
        Return Ass
    End Function
    Private Function GetSalesRMB(ByVal sMonth As Int16, ByVal Tcurrency As String)
        oCommand2.CommandText = "select sum(t1) from ( select nvl(sum(ogb14 * azj041),0) as t1 from hkacttest.ogb_file,hkacttest.oga_file,actiontest.azj_file where ogb01 = oga01 and ogapost = 'Y' and year(oga02) = " & tYear & " and month(oga02) = " & sMonth & " and oga23 = '" & Tcurrency & "' and azj01 = oga23 AND azj02 = year(oga02) || case when length(month(oga02)) = 1 then '0' end || month(oga02) "
        oCommand2.CommandText += "union all "
        oCommand2.CommandText += "select nvl(sum(ohb14 * -1 * azj041),0) from hkacttest.ohb_file,hkacttest.oha_file,actiontest.azj_file where ohb01 = oha01 and ohapost = 'Y' and year(oha02) = " & tYear & " and month(oha02) = " & sMonth & " and oha23 = '" & Tcurrency & "' and azj01 = oha23 AND azj02 = year(oha02) || case when length(month(oha02)) = 1 then '0' end || month(oha02) ) "
        Dim Ass As Decimal = oCommand2.ExecuteScalar()
        Return Ass
    End Function
    Private Function GetLastYearSameMonth(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
        oCommand2.CommandText += pYear & " and aah03 = " & tMonth
        Dim LYTM As Decimal = oCommand2.ExecuteScalar()
        Return LYTM
    End Function
    Private Function GetThisYearSameMonth(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
        oCommand2.CommandText += tYear & " and aah03 = " & tMonth
        Dim TYTM As Decimal = oCommand2.ExecuteScalar()
        Return TYTM
    End Function
    Private Function GetLastMonth(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
        oCommand2.CommandText += lYear & " and aah03 = " & lMonth
        Dim TYTM As Decimal = oCommand2.ExecuteScalar()
        Return TYTM
    End Function
    Private Function GetThisYearSameMonthBudget(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(tc_bud13),0) from tc_bud_file where tc_bud07 = '" & aag01 & "' and tc_bud02 = "
        oCommand2.CommandText += tYear & " and tc_bud03 = " & tMonth
        Dim TYTMB As Decimal = Decimal.Round(oCommand2.ExecuteScalar() / ExchangeRate1, 3)
        Return TYTMB
    End Function
    Private Function GetLastYearBeforeMonth(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
        oCommand2.CommandText += pYear & " and aah03 <= " & tMonth & " and aah03 > 0"
        Dim LYBM As Decimal = oCommand2.ExecuteScalar()
        Return LYBM
    End Function
    Private Function GetThisYearBeforeMonth(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
        oCommand2.CommandText += tYear & " and aah03 <= " & tMonth & " and aah03 > 0"
        Dim TYBM As Decimal = oCommand2.ExecuteScalar()
        Return TYBM
    End Function
    Private Function GetThisYearBeforeMonthBudget(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(tc_bud13),0) from tc_bud_file where tc_bud07 = '" & aag01 & "' and tc_bud02 = "
        oCommand2.CommandText += tYear & " and tc_bud03 <= " & tMonth
        Dim TYBMB As Decimal = Decimal.Round(oCommand2.ExecuteScalar() / ExchangeRate1, 3)
        Return TYBMB
    End Function
    Private Function GetLastYearNoMonth(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
        oCommand2.CommandText += pYear.ToString() & " and aah03 > 0"
        Dim TYNM As Decimal = oCommand2.ExecuteScalar()
        Return TYNM
    End Function
    Private Function GetThisYearBudget(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(tc_bud13),0) from tc_bud_file where tc_bud07 = '" & aag01 & "' and tc_bud02 = "
        oCommand2.CommandText += tYear.ToString()
        Dim TYTMB As Decimal = Decimal.Round(oCommand2.ExecuteScalar() / ExchangeRate1, 3)
        Return TYTMB
    End Function
    Private Function GetDepartNmae(ByVal gem01 As String)
        oCommand2.CommandText = "select gem02 from gem_file where gem01 = '" & gem01 & "'"
        Dim DN As String = oCommand2.ExecuteScalar()
        Return DN
    End Function
    Private Function GetLastYearSameMonth(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
        oCommand2.CommandText += pYear & " and aao04 = " & tMonth
        Dim LYTM As Decimal = oCommand2.ExecuteScalar()
        Return LYTM
    End Function
    Private Function GetThisYearSameMonth(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
        oCommand2.CommandText += tYear & " and aao04 = " & tMonth
        Dim TYTM As Decimal = oCommand2.ExecuteScalar()
        Return TYTM
    End Function
    Private Function GetLastMonth(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
        oCommand2.CommandText += lYear & " and aao04 = " & lMonth
        Dim TYTM As Decimal = oCommand2.ExecuteScalar()
        Return TYTM
    End Function
    Private Function GetThisYearSameMonthBudget(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(tc_bud13),0) from tc_bud_file where tc_bud07 = '" & aag01 & "' and tc_bud08 = '" & gem01 & "' and tc_bud02 = "
        oCommand2.CommandText += tYear & " and tc_bud03 = " & tMonth
        Dim TYTMB As Decimal = Decimal.Round(oCommand2.ExecuteScalar() / ExchangeRate1, 3)
        Return TYTMB
    End Function
    Private Function GetLastYearBeforeMonth(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
        oCommand2.CommandText += pYear & " and aao04 <= " & tMonth & " and aao04 > 0"
        Dim LYBM As Decimal = oCommand2.ExecuteScalar()
        Return LYBM
    End Function
    Private Function GetThisYearBeforeMonth(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
        oCommand2.CommandText += tYear & " and aao04 <= " & tMonth & " and aao04 > 0"
        Dim TYBM As Decimal = oCommand2.ExecuteScalar()
        Return TYBM
    End Function
    Private Function GetThisYearBeforeMonthBudget(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(tc_bud13),0) from tc_bud_file where tc_bud07 = '" & aag01 & "' and tc_bud08 = '" & gem01 & "' and tc_bud02 = "
        oCommand2.CommandText += tYear & " and tc_bud03 <= " & tMonth
        Dim TYBMB As Decimal = Decimal.Round(oCommand2.ExecuteScalar() / ExchangeRate1, 3)
        Return TYBMB
    End Function
    Private Function GetLastYearNoMonth(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
        oCommand2.CommandText += pYear.ToString() & " and aao04 > 0"
        Dim TYNM As Decimal = oCommand2.ExecuteScalar()
        Return TYNM
    End Function
    Private Function GetThisYearBudget(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(tc_bud13),0) from tc_bud_file where tc_bud07 = '" & aag01 & "' and tc_bud08 = '" & gem01 & "' and tc_bud02 = "
        oCommand2.CommandText += tYear.ToString()
        Dim TYTMB As Decimal = Decimal.Round(oCommand2.ExecuteScalar() / ExchangeRate1, 3)
        Return TYTMB
    End Function
End Class