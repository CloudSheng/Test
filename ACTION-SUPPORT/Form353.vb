Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop

Public Class Form353
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
    Dim eYear As Int16 = 0
    Dim eMonth As Int16 = 0
    Dim cMonth As Int16 = 0
    Dim m_cnt As Int16 = 0
    Dim pYear As Int16 = 0
    Dim tDate As Date
    Dim lYear As Int16 = 0
    Dim lMonth As Int16 = 0
    Dim DBC As String = String.Empty
    Dim LineZ As Integer = 0
    Dim DNP As String = String.Empty
    Dim ExchangeRate1 As Decimal = 1
    Dim ExchangeRate2 As Decimal = 1
    Dim ExchangeRate3 As Decimal = 1
    Dim ExchangeRate4 As Decimal = 1
    Dim ExchangeRate5 As Decimal = 1
    Dim ExchangeRate6 As Decimal = 1
    Dim ExchangeRate7 As Decimal = 1
    Dim ExchangeRate8 As Decimal = 1
    Dim ExchangeRate9 As Decimal = 1
    Dim ExchangeRate10 As Decimal = 1
    Dim ExchangeRate11 As Decimal = 1
    Dim ExchangeRate12 As Decimal = 1
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Dim SaveFileDialog1 As New SaveFileDialog
    Private Sub Form352_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'If Me.BackgroundWorker1.IsBusy() Then
        'MsgBox("处理中，请等待")
        'Return
        'End If        
        DBC = "hkacttest"
        oConnection.ConnectionString = Module1.OpenOracleConnection(DBC)
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
        tYear = Me.DateTimePicker1.Value.Year
        tMonth = Me.DateTimePicker1.Value.Month
        eYear = Me.DateTimePicker2.Value.Year
        eMonth = Me.DateTimePicker2.Value.Month

        If tYear <> eYear Then
            MsgBox("請輸入相同年度")
            oConnection.Close()
            Return
        End If

        If eMonth < tMonth Then
            MsgBox("年月區間輸入錯誤")
            oConnection.Close()
            Return
        End If

        m_cnt = eMonth - tMonth + 1

        'tMonth = 1
        'pYear = Me.DateTimePicker1.Value.AddYears(-1).Year
        tDate = Convert.ToDateTime(tYear & "/" & tMonth & "/01")
        'pYear = tDate.AddYears(-1).Year
        'lYear = Me.DateTimePicker1.Value.AddMonths(-1).Year
        'lMonth = Me.DateTimePicker1.Value.AddMonths(-1).Month

        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        ExportToExcel()
        oConnection.Close()

        DBC = "actiontest"
        oConnection.ConnectionString = Module1.OpenOracleConnection(DBC)
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
        ExportToExcel_1()

        SaveExcel()
    End Sub

    Private Sub ExportToExcel()
        ' 第一頁 (HAC)    
        Ws = xWorkBook.Sheets(1)
        Ws.Name = "HAC"
        Ws.Activate()
        AdjustExcelFormat()
        oCommand.CommandText = "select aag01,aag02 from aag_file where aag01 = '66013101' and aag07 = 2 order by aag01"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("aag01")
                Ws.Cells(LineZ, 2) = oReader.Item("aag02")
                oRng = Ws.Cells(LineZ, 2)
                oRng.WrapText = True
                oRng.ColumnWidth = 35

                'Ws.Cells(LineZ, 3) = Decimal.Round(GetThisYearMonth_1(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 4) = Decimal.Round(GetThisYearMonth_2(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 5) = Decimal.Round(GetThisYearMonth_3(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearMonth_4(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 7) = Decimal.Round(GetThisYearMonth_5(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 8) = Decimal.Round(GetThisYearMonth_6(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 9) = Decimal.Round(GetThisYearMonth_7(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 10) = Decimal.Round(GetThisYearMonth_8(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 11) = Decimal.Round(GetThisYearMonth_9(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 12) = Decimal.Round(GetThisYearMonth_10(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 13) = Decimal.Round(GetThisYearMonth_11(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 14) = Decimal.Round(GetThisYearMonth_12(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 15) = "=SUM(C" & LineZ & ":N" & LineZ & ")"

                For i = 1 To m_cnt
                    cMonth = tMonth + i - 1
                    Ws.Cells(LineZ, i + 2) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Next

                If m_cnt = 1 Then
                    Ws.Cells(LineZ, 4) = "=SUM(C" & LineZ & ":C" & LineZ & ")"
                End If
                If m_cnt = 2 Then
                    Ws.Cells(LineZ, 5) = "=SUM(C" & LineZ & ":D" & LineZ & ")"
                End If
                If m_cnt = 3 Then
                    Ws.Cells(LineZ, 6) = "=SUM(C" & LineZ & ":E" & LineZ & ")"
                End If
                If m_cnt = 4 Then
                    Ws.Cells(LineZ, 7) = "=SUM(C" & LineZ & ":F" & LineZ & ")"
                End If
                If m_cnt = 5 Then
                    Ws.Cells(LineZ, 8) = "=SUM(C" & LineZ & ":G" & LineZ & ")"
                End If
                If m_cnt = 6 Then
                    Ws.Cells(LineZ, 9) = "=SUM(C" & LineZ & ":H" & LineZ & ")"
                End If
                If m_cnt = 7 Then
                    Ws.Cells(LineZ, 10) = "=SUM(C" & LineZ & ":I" & LineZ & ")"
                End If
                If m_cnt = 8 Then
                    Ws.Cells(LineZ, 11) = "=SUM(C" & LineZ & ":J" & LineZ & ")"
                End If
                If m_cnt = 9 Then
                    Ws.Cells(LineZ, 12) = "=SUM(C" & LineZ & ":K" & LineZ & ")"
                End If
                If m_cnt = 10 Then
                    Ws.Cells(LineZ, 13) = "=SUM(C" & LineZ & ":L" & LineZ & ")"
                End If
                If m_cnt = 11 Then
                    Ws.Cells(LineZ, 14) = "=SUM(C" & LineZ & ":M" & LineZ & ")"
                End If
                If m_cnt = 12 Then
                    Ws.Cells(LineZ, 15) = "=SUM(C" & LineZ & ":N" & LineZ & ")"
                End If

                LineZ += 1
            End While
        End If
        oReader.Close()
        Ws.Cells(LineZ, 2) = "sut total in USD"
        Ws.Cells(LineZ, 3) = "=SUM(C7:C" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
        'oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 15)), Type:=xlFillDefault)
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, m_cnt + 3)), Type:=xlFillDefault)

        ' 劃線
        'oRng = Ws.Range("B7", Ws.Cells(LineZ, 15))
        If m_cnt = 1 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 4))
        End If
        If m_cnt = 2 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 5))
        End If
        If m_cnt = 3 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 6))
        End If
        If m_cnt = 4 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 7))
        End If
        If m_cnt = 5 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 8))
        End If
        If m_cnt = 6 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 9))
        End If
        If m_cnt = 7 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 10))
        End If
        If m_cnt = 8 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 11))
        End If
        If m_cnt = 9 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 12))
        End If
        If m_cnt = 10 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 13))
        End If
        If m_cnt = 11 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 14))
        End If
        If m_cnt = 12 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 15))
        End If
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "

        oRng = Ws.Range("A6", Ws.Cells(LineZ, 1))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        'oRng = Ws.Range("C6", "N6")
        If m_cnt = 1 Then
            oRng = Ws.Range("C6", "C6")
        End If
        If m_cnt = 2 Then
            oRng = Ws.Range("C6", "D6")
        End If
        If m_cnt = 3 Then
            oRng = Ws.Range("C6", "E6")
        End If
        If m_cnt = 4 Then
            oRng = Ws.Range("C6", "F6")
        End If
        If m_cnt = 5 Then
            oRng = Ws.Range("C6", "G6")
        End If
        If m_cnt = 6 Then
            oRng = Ws.Range("C6", "H6")
        End If
        If m_cnt = 7 Then
            oRng = Ws.Range("C6", "I6")
        End If
        If m_cnt = 8 Then
            oRng = Ws.Range("C6", "J6")
        End If
        If m_cnt = 9 Then
            oRng = Ws.Range("C6", "K6")
        End If
        If m_cnt = 10 Then
            oRng = Ws.Range("C6", "L6")
        End If
        If m_cnt = 11 Then
            oRng = Ws.Range("C6", "M6")
        End If
        If m_cnt = 12 Then
            oRng = Ws.Range("C6", "N6")
        End If
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        '參考Form114
        'LineS1 = LineZ
        'LineZ += 2
        'AdjustExcelFormat2()
        'oCommand.CommandText = "select oga03,oga032"
        'For i As Int16 = 1 To 53 Step 1
        'oCommand.CommandText += ",sum(t" & i & ") as t" & i
        'Next
        'oCommand.CommandText += " from ( select oga03,oga032"
        'For i As Int16 = 1 To 53 Step 1
        'oCommand.CommandText += ",(case when azn05 = " & i & " then ogb14t * oga24 else 0 end) as t" & i
        'Next
        'oCommand.CommandText += " from hkacttest.oga_file left join hkacttest.ogb_file on oga01 = ogb01 "
        'oCommand.CommandText += "left join azn_file on oga02 = azn01 where ogapost = 'Y' and oga04 <> 'D0003' and oga02 between to_date('"
        'oCommand.CommandText += Me.DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        'oCommand.CommandText += Me.DateTimePicker2.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ogb04 = 'AC0000000000' "
        'If Not String.IsNullOrEmpty(g_oga03) Then
        'oCommand.CommandText += " AND oga03 ='" & g_oga03 & "' "
        'End If
        ' 20180312 add oha ohb
        'oCommand.CommandText += "union all "
        'oCommand.CommandText += "select oha03,oha032"
        'For i As Int16 = 1 To 53 Step 1
        'oCommand.CommandText += ",(case when azn05 = " & i & " then ohb14t * oha24 * (-1) else 0 end) as t" & i
        'Next
        'oCommand.CommandText += " from hkacttest.oha_file left join hkacttest.ohb_file on oha01 = ohb01 "
        'oCommand.CommandText += "left join azn_file on oha02 = azn01 where ohapost = 'Y' and oha04 <> 'D0003' and oha02 between to_date('"
        'oCommand.CommandText += Me.DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        'oCommand.CommandText += Me.DateTimePicker2.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ohb04 = 'AC0000000000' "
        'If Not String.IsNullOrEmpty(g_oga03) Then
        'oCommand.CommandText += " AND oha03 ='" & g_oga03 & "' "
        'End If
        'oCommand.CommandText += " ) group by oga03,oga032 order by oga03"
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        'While oReader.Read()
        'For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
        'Ws.Cells(LineZ, i + 1) = oReader.Item(i)
        'Next
        'Ws.Cells(LineZ, 56) = "=SUM(C" & LineZ & ":BC" & LineZ & ")"
        'LineZ += 1
        'End While
        'oRng = Ws.Range(Ws.Cells(LineS1 + 4, 3), Ws.Cells(LineZ, 56))
        'oRng.NumberFormatLocal = "[$$-en-CA]#,##0.00;-[$$-en-CA]#,##0.00"
        'oRng = Ws.Range(Ws.Cells(LineS1 + 4, 56), Ws.Cells(LineZ, 56))
        'oRng.Interior.Color = Color.DarkGray
        'Ws.Cells(LineZ, 1) = "Total"
        'Ws.Cells(LineZ, 3) = "=SUM(C" & LineS1 + 4 & ":C" & LineZ - 1 & ")"
        'oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
        'oRng.Interior.Color = Color.DarkGray
        'oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 56)), Type:=xlFillDefault)
        'oRng = Ws.Range(Ws.Cells(LineS1 + 4, 1), Ws.Cells(LineZ, 56))
        'oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        'oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        'oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        'End If
        'oReader.Close()
        'LineS2 = LineZ
        ' 試作 圖表 20180315
        'Dim XA As Excel.Chart = Ws.Shapes.AddChart(xlColumnClustered, 50, 400, 2500, 600).Chart
        'oRng = Ws.Range("B5", Ws.Cells(LineS1 - 1, 55))
        'XA.SetSourceData(oRng)

        'XA.SetElement(msoElementChartTitleAboveChart)
        'XA.ChartTitle.Text = "Weekly HAC Part sales"
        'XA.SetElement(msoElementLegendNone)
        'XA.SetElement(msoElementDataTableWithLegendKeys)

    End Sub

    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 10.44
        Ws.Cells(4, 2) = tYear & " Premium Freight Cost"

        Ws.Cells(5, 2) = "HAC"
        ExchangeRate1 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "01'"
        'ExchangeRate1 = oCommand.ExecuteScalar()
        'If ExchangeRate1 = 0 Then ExchangeRate1 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "02'"
        'ExchangeRate2 = oCommand.ExecuteScalar()
        'If ExchangeRate2 = 0 Then ExchangeRate2 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "03'"
        'ExchangeRate3 = oCommand.ExecuteScalar()
        'If ExchangeRate3 = 0 Then ExchangeRate3 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "04'"
        'ExchangeRate4 = oCommand.ExecuteScalar()
        'If ExchangeRate4 = 0 Then ExchangeRate4 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "05'"
        'ExchangeRate5 = oCommand.ExecuteScalar()
        'If ExchangeRate5 = 0 Then ExchangeRate5 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "06'"
        'ExchangeRate6 = oCommand.ExecuteScalar()
        'If ExchangeRate6 = 0 Then ExchangeRate6 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "07'"
        'ExchangeRate7 = oCommand.ExecuteScalar()
        'If ExchangeRate7 = 0 Then ExchangeRate7 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "08'"
        'ExchangeRate8 = oCommand.ExecuteScalar()
        'If ExchangeRate8 = 0 Then ExchangeRate8 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "09'"
        'ExchangeRate9 = oCommand.ExecuteScalar()
        'If ExchangeRate9 = 0 Then ExchangeRate9 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "10'"
        'ExchangeRate10 = oCommand.ExecuteScalar()
        'If ExchangeRate10 = 0 Then ExchangeRate10 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "11'"
        'ExchangeRate11 = oCommand.ExecuteScalar()
        'If ExchangeRate11 = 0 Then ExchangeRate11 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "12'"
        'ExchangeRate12 = oCommand.ExecuteScalar()
        'If ExchangeRate12 = 0 Then ExchangeRate12 = 1

        Ws.Cells(6, 1) = "Account"
        Ws.Cells(6, 2) = "Month"
        'Ws.Cells(6, 3) = tYear & "-01"
        'Ws.Cells(6, 4) = tYear & "-02"
        'Ws.Cells(6, 5) = tYear & "-03"
        'Ws.Cells(6, 6) = tYear & "-04"
        'Ws.Cells(6, 7) = tYear & "-05"
        'Ws.Cells(6, 8) = tYear & "-06"
        'Ws.Cells(6, 9) = tYear & "-07"
        'Ws.Cells(6, 10) = tYear & "-08"
        'Ws.Cells(6, 11) = tYear & "-09"
        'Ws.Cells(6, 12) = tYear & "-10"
        'Ws.Cells(6, 13) = tYear & "-11"
        'Ws.Cells(6, 14) = tYear & "-12"
        'oRng = Ws.Range("C6", "N6")
        If m_cnt = 1 Then
            oRng = Ws.Range("C6", "C6")
        End If
        If m_cnt = 2 Then
            oRng = Ws.Range("C6", "D6")
        End If
        If m_cnt = 3 Then
            oRng = Ws.Range("C6", "E6")
        End If
        If m_cnt = 4 Then
            oRng = Ws.Range("C6", "F6")
        End If
        If m_cnt = 5 Then
            oRng = Ws.Range("C6", "G6")
        End If
        If m_cnt = 6 Then
            oRng = Ws.Range("C6", "H6")
        End If
        If m_cnt = 7 Then
            oRng = Ws.Range("C6", "I6")
        End If
        If m_cnt = 8 Then
            oRng = Ws.Range("C6", "J6")
        End If
        If m_cnt = 9 Then
            oRng = Ws.Range("C6", "K6")
        End If
        If m_cnt = 10 Then
            oRng = Ws.Range("C6", "L6")
        End If
        If m_cnt = 11 Then
            oRng = Ws.Range("C6", "M6")
        End If
        If m_cnt = 12 Then
            oRng = Ws.Range("C6", "N6")
        End If

        oRng.NumberFormatLocal = "mmm-yy"
        'Ws.Cells(6, 3) = tDate.AddMonths(0)
        'Ws.Cells(6, 4) = tDate.AddMonths(1)
        'Ws.Cells(6, 5) = tDate.AddMonths(2)
        'Ws.Cells(6, 6) = tDate.AddMonths(3)
        'Ws.Cells(6, 7) = tDate.AddMonths(4)
        'Ws.Cells(6, 8) = tDate.AddMonths(5)
        'Ws.Cells(6, 9) = tDate.AddMonths(6)
        'Ws.Cells(6, 10) = tDate.AddMonths(7)
        'Ws.Cells(6, 11) = tDate.AddMonths(8)
        'Ws.Cells(6, 12) = tDate.AddMonths(9)
        'Ws.Cells(6, 13) = tDate.AddMonths(10)
        'Ws.Cells(6, 14) = tDate.AddMonths(11)
        'Ws.Cells(6, 15) = "YTD " & tYear
        For i = 1 To m_cnt
            Ws.Cells(6, i + 2) = tDate.AddMonths(i - 1)
        Next
        Ws.Cells(6, m_cnt + 3) = "YTD " & tYear

        ' 劃線
        'oRng = Ws.Range("B6", "O6")
        If m_cnt = 1 Then
            oRng = Ws.Range("B6", "D6")
        End If
        If m_cnt = 2 Then
            oRng = Ws.Range("B6", "E6")
        End If
        If m_cnt = 3 Then
            oRng = Ws.Range("B6", "F6")
        End If
        If m_cnt = 4 Then
            oRng = Ws.Range("B6", "G6")
        End If
        If m_cnt = 5 Then
            oRng = Ws.Range("B6", "H6")
        End If
        If m_cnt = 6 Then
            oRng = Ws.Range("B6", "I6")
        End If
        If m_cnt = 7 Then
            oRng = Ws.Range("B6", "J6")
        End If
        If m_cnt = 8 Then
            oRng = Ws.Range("B6", "K6")
        End If
        If m_cnt = 9 Then
            oRng = Ws.Range("B6", "L6")
        End If
        If m_cnt = 10 Then
            oRng = Ws.Range("B6", "M6")
        End If
        If m_cnt = 11 Then
            oRng = Ws.Range("B6", "N6")
        End If
        If m_cnt = 12 Then
            oRng = Ws.Range("B6", "O6")
        End If
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        'oRng = Ws.Range("C6", "O6")
        If m_cnt = 1 Then
            oRng = Ws.Range("C6", "D6")
        End If
        If m_cnt = 2 Then
            oRng = Ws.Range("C6", "E6")
        End If
        If m_cnt = 3 Then
            oRng = Ws.Range("C6", "F6")
        End If
        If m_cnt = 4 Then
            oRng = Ws.Range("C6", "G6")
        End If
        If m_cnt = 5 Then
            oRng = Ws.Range("C6", "H6")
        End If
        If m_cnt = 6 Then
            oRng = Ws.Range("C6", "I6")
        End If
        If m_cnt = 7 Then
            oRng = Ws.Range("C6", "J6")
        End If
        If m_cnt = 8 Then
            oRng = Ws.Range("C6", "K6")
        End If
        If m_cnt = 9 Then
            oRng = Ws.Range("C6", "L6")
        End If
        If m_cnt = 10 Then
            oRng = Ws.Range("C6", "M6")
        End If
        If m_cnt = 11 Then
            oRng = Ws.Range("C6", "N6")
        End If
        If m_cnt = 12 Then
            oRng = Ws.Range("C6", "O6")
        End If
        oRng.HorizontalAlignment = xlRight
        LineZ = 7
    End Sub

    Private Sub ExportToExcel_1()
        ' 第二頁 (DAC)    
        Ws = xWorkBook.Sheets(2)
        Ws.Name = "DAC"
        Ws.Activate()
        AdjustExcelFormat1()
        oCommand.CommandText = "select aag01,aag02 from aag_file where aag01 = '66013101' and aag07 = 2 order by aag01"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("aag01")
                Ws.Cells(LineZ, 2) = oReader.Item("aag02")
                oRng = Ws.Cells(LineZ, 2)
                oRng.WrapText = True
                oRng.ColumnWidth = 35

                'Ws.Cells(LineZ, 3) = Decimal.Round(GetThisYearMonth_1(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 4) = Decimal.Round(GetThisYearMonth_2(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 5) = Decimal.Round(GetThisYearMonth_3(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearMonth_4(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 7) = Decimal.Round(GetThisYearMonth_5(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 8) = Decimal.Round(GetThisYearMonth_6(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 9) = Decimal.Round(GetThisYearMonth_7(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 10) = Decimal.Round(GetThisYearMonth_8(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 11) = Decimal.Round(GetThisYearMonth_9(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 12) = Decimal.Round(GetThisYearMonth_10(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 13) = Decimal.Round(GetThisYearMonth_11(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 14) = Decimal.Round(GetThisYearMonth_12(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 15) = "=SUM(C" & LineZ & ":N" & LineZ & ")"

                For i = 1 To m_cnt
                    cMonth = tMonth + i - 1
                    Ws.Cells(LineZ, i + 2) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Next

                If m_cnt = 1 Then
                    Ws.Cells(LineZ, 4) = "=SUM(C" & LineZ & ":C" & LineZ & ")"
                End If
                If m_cnt = 2 Then
                    Ws.Cells(LineZ, 5) = "=SUM(C" & LineZ & ":D" & LineZ & ")"
                End If
                If m_cnt = 3 Then
                    Ws.Cells(LineZ, 6) = "=SUM(C" & LineZ & ":E" & LineZ & ")"
                End If
                If m_cnt = 4 Then
                    Ws.Cells(LineZ, 7) = "=SUM(C" & LineZ & ":F" & LineZ & ")"
                End If
                If m_cnt = 5 Then
                    Ws.Cells(LineZ, 8) = "=SUM(C" & LineZ & ":G" & LineZ & ")"
                End If
                If m_cnt = 6 Then
                    Ws.Cells(LineZ, 9) = "=SUM(C" & LineZ & ":H" & LineZ & ")"
                End If
                If m_cnt = 7 Then
                    Ws.Cells(LineZ, 10) = "=SUM(C" & LineZ & ":I" & LineZ & ")"
                End If
                If m_cnt = 8 Then
                    Ws.Cells(LineZ, 11) = "=SUM(C" & LineZ & ":J" & LineZ & ")"
                End If
                If m_cnt = 9 Then
                    Ws.Cells(LineZ, 12) = "=SUM(C" & LineZ & ":K" & LineZ & ")"
                End If
                If m_cnt = 10 Then
                    Ws.Cells(LineZ, 13) = "=SUM(C" & LineZ & ":L" & LineZ & ")"
                End If
                If m_cnt = 11 Then
                    Ws.Cells(LineZ, 14) = "=SUM(C" & LineZ & ":M" & LineZ & ")"
                End If
                If m_cnt = 12 Then
                    Ws.Cells(LineZ, 15) = "=SUM(C" & LineZ & ":N" & LineZ & ")"
                End If
                LineZ += 1
            End While
        End If
        oReader.Close()
        Ws.Cells(LineZ, 2) = "sut total in RMB"
        Ws.Cells(LineZ, 3) = "=SUM(C7:C" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
        'oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 15)), Type:=xlFillDefault)
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, m_cnt + 3)), Type:=xlFillDefault)

        ' 劃線
        'oRng = Ws.Range("B7", Ws.Cells(LineZ, 15))
        If m_cnt = 1 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 4))
        End If
        If m_cnt = 2 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 5))
        End If
        If m_cnt = 3 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 6))
        End If
        If m_cnt = 4 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 7))
        End If
        If m_cnt = 5 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 8))
        End If
        If m_cnt = 6 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 9))
        End If
        If m_cnt = 7 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 10))
        End If
        If m_cnt = 8 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 11))
        End If
        If m_cnt = 9 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 12))
        End If
        If m_cnt = 10 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 13))
        End If
        If m_cnt = 11 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 14))
        End If
        If m_cnt = 12 Then
            oRng = Ws.Range("B7", Ws.Cells(LineZ, 15))
        End If
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "

        oRng = Ws.Range("A6", Ws.Cells(LineZ, 1))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        'oRng = Ws.Range("C6", "N6")
        If m_cnt = 1 Then
            oRng = Ws.Range("C6", "C6")
        End If
        If m_cnt = 2 Then
            oRng = Ws.Range("C6", "D6")
        End If
        If m_cnt = 3 Then
            oRng = Ws.Range("C6", "E6")
        End If
        If m_cnt = 4 Then
            oRng = Ws.Range("C6", "F6")
        End If
        If m_cnt = 5 Then
            oRng = Ws.Range("C6", "G6")
        End If
        If m_cnt = 6 Then
            oRng = Ws.Range("C6", "H6")
        End If
        If m_cnt = 7 Then
            oRng = Ws.Range("C6", "I6")
        End If
        If m_cnt = 8 Then
            oRng = Ws.Range("C6", "J6")
        End If
        If m_cnt = 9 Then
            oRng = Ws.Range("C6", "K6")
        End If
        If m_cnt = 10 Then
            oRng = Ws.Range("C6", "L6")
        End If
        If m_cnt = 11 Then
            oRng = Ws.Range("C6", "M6")
        End If
        If m_cnt = 12 Then
            oRng = Ws.Range("C6", "N6")
        End If
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        AdjustExcelFormat1_1()
        oCommand.CommandText = "select aag01,aag02 from aag_file where aag01 = '66013101' and aag07 = 2 order by aag01"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("aag01")
                Ws.Cells(LineZ, 2) = oReader.Item("aag02")
                oRng = Ws.Cells(LineZ, 2)
                oRng.WrapText = True
                oRng.ColumnWidth = 35

                'Ws.Cells(LineZ, 3) = Decimal.Round(GetThisYearMonth_1(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Ws.Cells(LineZ, 4) = Decimal.Round(GetThisYearMonth_2(oReader.Item("aag01").ToString()) / ExchangeRate2, 3)
                'Ws.Cells(LineZ, 5) = Decimal.Round(GetThisYearMonth_3(oReader.Item("aag01").ToString()) / ExchangeRate3, 3)
                'Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearMonth_4(oReader.Item("aag01").ToString()) / ExchangeRate4, 3)
                'Ws.Cells(LineZ, 7) = Decimal.Round(GetThisYearMonth_5(oReader.Item("aag01").ToString()) / ExchangeRate5, 3)
                'Ws.Cells(LineZ, 8) = Decimal.Round(GetThisYearMonth_6(oReader.Item("aag01").ToString()) / ExchangeRate6, 3)
                'Ws.Cells(LineZ, 9) = Decimal.Round(GetThisYearMonth_7(oReader.Item("aag01").ToString()) / ExchangeRate7, 3)
                'Ws.Cells(LineZ, 10) = Decimal.Round(GetThisYearMonth_8(oReader.Item("aag01").ToString()) / ExchangeRate8, 3)
                'Ws.Cells(LineZ, 11) = Decimal.Round(GetThisYearMonth_9(oReader.Item("aag01").ToString()) / ExchangeRate9, 3)
                'Ws.Cells(LineZ, 12) = Decimal.Round(GetThisYearMonth_10(oReader.Item("aag01").ToString()) / ExchangeRate10, 3)
                'Ws.Cells(LineZ, 13) = Decimal.Round(GetThisYearMonth_11(oReader.Item("aag01").ToString()) / ExchangeRate11, 3)
                'Ws.Cells(LineZ, 14) = Decimal.Round(GetThisYearMonth_12(oReader.Item("aag01").ToString()) / ExchangeRate12, 3)
                'Ws.Cells(LineZ, 15) = "=SUM(C" & LineZ & ":N" & LineZ & ")"

                If m_cnt = 1 Then
                    cMonth = tMonth
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 3) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    Ws.Cells(LineZ, 4) = "=SUM(C" & LineZ & ":C" & LineZ & ")"
                End If
                If m_cnt = 2 Then
                    cMonth = tMonth
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 3) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 1
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 4) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    Ws.Cells(LineZ, 5) = "=SUM(C" & LineZ & ":D" & LineZ & ")"
                End If
                If m_cnt = 3 Then
                    cMonth = tMonth
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 3) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 1
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 4) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 2
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 5) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    Ws.Cells(LineZ, 6) = "=SUM(C" & LineZ & ":E" & LineZ & ")"
                End If
                If m_cnt = 4 Then
                    cMonth = tMonth
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 3) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 1
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 4) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 2
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 5) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 3
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    Ws.Cells(LineZ, 7) = "=SUM(C" & LineZ & ":F" & LineZ & ")"
                End If
                If m_cnt = 5 Then
                    cMonth = tMonth
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 3) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 1
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 4) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 2
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 5) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 3
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 4
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 7) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    Ws.Cells(LineZ, 8) = "=SUM(C" & LineZ & ":G" & LineZ & ")"
                End If
                If m_cnt = 6 Then
                    cMonth = tMonth
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 3) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 1
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 4) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 2
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 5) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 3
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 4
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 7) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 5
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 8) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    Ws.Cells(LineZ, 9) = "=SUM(C" & LineZ & ":H" & LineZ & ")"
                End If
                If m_cnt = 7 Then
                    cMonth = tMonth
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 3) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 1
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 4) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 2
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 5) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 3
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 4
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 7) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 5
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 8) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 6
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 9) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    Ws.Cells(LineZ, 10) = "=SUM(C" & LineZ & ":I" & LineZ & ")"
                End If
                If m_cnt = 8 Then
                    cMonth = tMonth
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 3) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 1
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 4) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 2
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 5) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 3
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 4
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 7) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 5
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 8) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 6
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 9) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 7
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 10) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    Ws.Cells(LineZ, 11) = "=SUM(C" & LineZ & ":J" & LineZ & ")"
                End If
                If m_cnt = 9 Then
                    cMonth = tMonth
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 3) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 1
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 4) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 2
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 5) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 3
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 4
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 7) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 5
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 8) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 6
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 9) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 7
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 10) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 8
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 11) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    Ws.Cells(LineZ, 12) = "=SUM(C" & LineZ & ":K" & LineZ & ")"
                End If
                If m_cnt = 10 Then
                    cMonth = tMonth
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 3) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 1
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 4) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 2
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 5) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 3
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 4
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 7) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 5
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 8) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 6
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 9) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 7
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 10) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 8
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 11) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 9
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 12) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    Ws.Cells(LineZ, 13) = "=SUM(C" & LineZ & ":L" & LineZ & ")"
                End If
                If m_cnt = 11 Then
                    cMonth = tMonth
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 3) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 1
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 4) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 2
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 5) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 3
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 4
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 7) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 5
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 8) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 6
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 9) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 7
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 10) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 8
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 11) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 9
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 12) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 10
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 13) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    Ws.Cells(LineZ, 14) = "=SUM(C" & LineZ & ":M" & LineZ & ")"
                End If
                If m_cnt = 12 Then
                    cMonth = tMonth
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 3) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 1
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 4) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 2
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 5) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 3
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 4
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 7) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 5
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 8) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 6
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 9) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 7
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 10) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 8
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 11) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 9
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 12) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 10
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 13) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    cMonth = tMonth + 11
                    Find_ExchangeRate()
                    Ws.Cells(LineZ, 14) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                    Ws.Cells(LineZ, 15) = "=SUM(C" & LineZ & ":N" & LineZ & ")"
                End If
                LineZ += 1
            End While
        End If
        oReader.Close()
        Ws.Cells(LineZ, 2) = "sut total in USD"
        Ws.Cells(LineZ, 3) = "=SUM(C11:C" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
        'oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 15)), Type:=xlFillDefault)
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, m_cnt + 3)), Type:=xlFillDefault)

        ' 劃線
        'oRng = Ws.Range("B11", Ws.Cells(LineZ, 15))
        If m_cnt = 1 Then
            oRng = Ws.Range("B11", Ws.Cells(LineZ, 4))
        End If
        If m_cnt = 2 Then
            oRng = Ws.Range("B11", Ws.Cells(LineZ, 5))
        End If
        If m_cnt = 3 Then
            oRng = Ws.Range("B11", Ws.Cells(LineZ, 6))
        End If
        If m_cnt = 4 Then
            oRng = Ws.Range("B11", Ws.Cells(LineZ, 7))
        End If
        If m_cnt = 5 Then
            oRng = Ws.Range("B11", Ws.Cells(LineZ, 8))
        End If
        If m_cnt = 6 Then
            oRng = Ws.Range("B11", Ws.Cells(LineZ, 9))
        End If
        If m_cnt = 7 Then
            oRng = Ws.Range("B11", Ws.Cells(LineZ, 10))
        End If
        If m_cnt = 8 Then
            oRng = Ws.Range("B11", Ws.Cells(LineZ, 11))
        End If
        If m_cnt = 9 Then
            oRng = Ws.Range("B11", Ws.Cells(LineZ, 12))
        End If
        If m_cnt = 10 Then
            oRng = Ws.Range("B11", Ws.Cells(LineZ, 13))
        End If
        If m_cnt = 11 Then
            oRng = Ws.Range("B11", Ws.Cells(LineZ, 14))
        End If
        If m_cnt = 12 Then
            oRng = Ws.Range("B11", Ws.Cells(LineZ, 15))
        End If
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "

        oRng = Ws.Range("A10", Ws.Cells(LineZ, 1))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        'oRng = Ws.Range("C10", "N10")
        If m_cnt = 1 Then
            oRng = Ws.Range("C10", "C10")
        End If
        If m_cnt = 2 Then
            oRng = Ws.Range("C10", "D10")
        End If
        If m_cnt = 3 Then
            oRng = Ws.Range("C10", "E10")
        End If
        If m_cnt = 4 Then
            oRng = Ws.Range("C10", "F10")
        End If
        If m_cnt = 5 Then
            oRng = Ws.Range("C10", "G10")
        End If
        If m_cnt = 6 Then
            oRng = Ws.Range("C10", "H10")
        End If
        If m_cnt = 7 Then
            oRng = Ws.Range("C10", "I10")
        End If
        If m_cnt = 8 Then
            oRng = Ws.Range("C10", "J10")
        End If
        If m_cnt = 9 Then
            oRng = Ws.Range("C10", "K10")
        End If
        If m_cnt = 10 Then
            oRng = Ws.Range("C10", "L10")
        End If
        If m_cnt = 11 Then
            oRng = Ws.Range("C10", "M10")
        End If
        If m_cnt = 12 Then
            oRng = Ws.Range("C10", "N10")
        End If
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
    End Sub

    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 10.44
        Ws.Cells(4, 2) = tYear & " Premium Freight Cost"

        Ws.Cells(5, 2) = "DAC"
        ExchangeRate1 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "01'"
        'ExchangeRate1 = oCommand.ExecuteScalar()
        'If ExchangeRate1 = 0 Then ExchangeRate1 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "02'"
        'ExchangeRate2 = oCommand.ExecuteScalar()
        'If ExchangeRate2 = 0 Then ExchangeRate2 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "03'"
        'ExchangeRate3 = oCommand.ExecuteScalar()
        'If ExchangeRate3 = 0 Then ExchangeRate3 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "04'"
        'ExchangeRate4 = oCommand.ExecuteScalar()
        'If ExchangeRate4 = 0 Then ExchangeRate4 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "05'"
        'ExchangeRate5 = oCommand.ExecuteScalar()
        'If ExchangeRate5 = 0 Then ExchangeRate5 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "06'"
        'ExchangeRate6 = oCommand.ExecuteScalar()
        'If ExchangeRate6 = 0 Then ExchangeRate6 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "07'"
        'ExchangeRate7 = oCommand.ExecuteScalar()
        'If ExchangeRate7 = 0 Then ExchangeRate7 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "08'"
        'ExchangeRate8 = oCommand.ExecuteScalar()
        'If ExchangeRate8 = 0 Then ExchangeRate8 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "09'"
        'ExchangeRate9 = oCommand.ExecuteScalar()
        'If ExchangeRate9 = 0 Then ExchangeRate9 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "10'"
        'ExchangeRate10 = oCommand.ExecuteScalar()
        'If ExchangeRate10 = 0 Then ExchangeRate10 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "11'"
        'ExchangeRate11 = oCommand.ExecuteScalar()
        'If ExchangeRate11 = 0 Then ExchangeRate11 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'EUR' and azj02 = '" & tYear & "12'"
        'ExchangeRate12 = oCommand.ExecuteScalar()
        'If ExchangeRate12 = 0 Then ExchangeRate12 = 1

        Ws.Cells(6, 1) = "Account"
        Ws.Cells(6, 2) = "Month"
        'Ws.Cells(6, 3) = tYear & "-01"
        'Ws.Cells(6, 4) = tYear & "-02"
        'Ws.Cells(6, 5) = tYear & "-03"
        'Ws.Cells(6, 6) = tYear & "-04"
        'Ws.Cells(6, 7) = tYear & "-05"
        'Ws.Cells(6, 8) = tYear & "-06"
        'Ws.Cells(6, 9) = tYear & "-07"
        'Ws.Cells(6, 10) = tYear & "-08"
        'Ws.Cells(6, 11) = tYear & "-09"
        'Ws.Cells(6, 12) = tYear & "-10"
        'Ws.Cells(6, 13) = tYear & "-11"
        'Ws.Cells(6, 14) = tYear & "-12"
        'oRng = Ws.Range("C6", "N6")
        If m_cnt = 1 Then
            oRng = Ws.Range("C6", "C6")
        End If
        If m_cnt = 2 Then
            oRng = Ws.Range("C6", "D6")
        End If
        If m_cnt = 3 Then
            oRng = Ws.Range("C6", "E6")
        End If
        If m_cnt = 4 Then
            oRng = Ws.Range("C6", "F6")
        End If
        If m_cnt = 5 Then
            oRng = Ws.Range("C6", "G6")
        End If
        If m_cnt = 6 Then
            oRng = Ws.Range("C6", "H6")
        End If
        If m_cnt = 7 Then
            oRng = Ws.Range("C6", "I6")
        End If
        If m_cnt = 8 Then
            oRng = Ws.Range("C6", "J6")
        End If
        If m_cnt = 9 Then
            oRng = Ws.Range("C6", "K6")
        End If
        If m_cnt = 10 Then
            oRng = Ws.Range("C6", "L6")
        End If
        If m_cnt = 11 Then
            oRng = Ws.Range("C6", "M6")
        End If
        If m_cnt = 12 Then
            oRng = Ws.Range("C6", "N6")
        End If
        oRng.NumberFormatLocal = "mmm-yy"
        'Ws.Cells(6, 3) = tDate.AddMonths(0)
        'Ws.Cells(6, 4) = tDate.AddMonths(1)
        'Ws.Cells(6, 5) = tDate.AddMonths(2)
        'Ws.Cells(6, 6) = tDate.AddMonths(3)
        'Ws.Cells(6, 7) = tDate.AddMonths(4)
        'Ws.Cells(6, 8) = tDate.AddMonths(5)
        'Ws.Cells(6, 9) = tDate.AddMonths(6)
        'Ws.Cells(6, 10) = tDate.AddMonths(7)
        'Ws.Cells(6, 11) = tDate.AddMonths(8)
        'Ws.Cells(6, 12) = tDate.AddMonths(9)
        'Ws.Cells(6, 13) = tDate.AddMonths(10)
        'Ws.Cells(6, 14) = tDate.AddMonths(11)
        'Ws.Cells(6, 15) = "YTD " & tYear
        For i = 1 To m_cnt
            Ws.Cells(6, i + 2) = tDate.AddMonths(i - 1)
        Next
        Ws.Cells(6, m_cnt + 3) = "YTD " & tYear

        ' 劃線
        'oRng = Ws.Range("B6", "O6")
        If m_cnt = 1 Then
            oRng = Ws.Range("B6", "D6")
        End If
        If m_cnt = 2 Then
            oRng = Ws.Range("B6", "E6")
        End If
        If m_cnt = 3 Then
            oRng = Ws.Range("B6", "F6")
        End If
        If m_cnt = 4 Then
            oRng = Ws.Range("B6", "G6")
        End If
        If m_cnt = 5 Then
            oRng = Ws.Range("B6", "H6")
        End If
        If m_cnt = 6 Then
            oRng = Ws.Range("B6", "I6")
        End If
        If m_cnt = 7 Then
            oRng = Ws.Range("B6", "J6")
        End If
        If m_cnt = 8 Then
            oRng = Ws.Range("B6", "K6")
        End If
        If m_cnt = 9 Then
            oRng = Ws.Range("B6", "L6")
        End If
        If m_cnt = 10 Then
            oRng = Ws.Range("B6", "M6")
        End If
        If m_cnt = 11 Then
            oRng = Ws.Range("B6", "N6")
        End If
        If m_cnt = 12 Then
            oRng = Ws.Range("B6", "O6")
        End If
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        'oRng = Ws.Range("C6", "O6")
        If m_cnt = 1 Then
            oRng = Ws.Range("C6", "D6")
        End If
        If m_cnt = 2 Then
            oRng = Ws.Range("C6", "E6")
        End If
        If m_cnt = 3 Then
            oRng = Ws.Range("C6", "F6")
        End If
        If m_cnt = 4 Then
            oRng = Ws.Range("C6", "G6")
        End If
        If m_cnt = 5 Then
            oRng = Ws.Range("C6", "H6")
        End If
        If m_cnt = 6 Then
            oRng = Ws.Range("C6", "I6")
        End If
        If m_cnt = 7 Then
            oRng = Ws.Range("C6", "J6")
        End If
        If m_cnt = 8 Then
            oRng = Ws.Range("C6", "K6")
        End If
        If m_cnt = 9 Then
            oRng = Ws.Range("C6", "L6")
        End If
        If m_cnt = 10 Then
            oRng = Ws.Range("C6", "M6")
        End If
        If m_cnt = 11 Then
            oRng = Ws.Range("C6", "N6")
        End If
        If m_cnt = 12 Then
            oRng = Ws.Range("C6", "O6")
        End If
        oRng.HorizontalAlignment = xlRight
        LineZ = 7
    End Sub

    Private Sub Find_ExchangeRate()
        Dim TYM1 As String = String.Empty
        If cMonth < 10 Then
            TYM1 = tYear & "0" & cMonth
        Else
            TYM1 = tYear & cMonth
        End If
        oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & TYM1 & "'"
        ExchangeRate1 = oCommand.ExecuteScalar()
        If ExchangeRate1 = 0 Then ExchangeRate1 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & tYear & "01'"
        'ExchangeRate1 = oCommand.ExecuteScalar()
        'If ExchangeRate1 = 0 Then ExchangeRate1 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & tYear & "02'"
        'ExchangeRate2 = oCommand.ExecuteScalar()
        'If ExchangeRate2 = 0 Then ExchangeRate2 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & tYear & "03'"
        'ExchangeRate3 = oCommand.ExecuteScalar()
        'If ExchangeRate3 = 0 Then ExchangeRate3 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & tYear & "04'"
        'ExchangeRate4 = oCommand.ExecuteScalar()
        'If ExchangeRate4 = 0 Then ExchangeRate4 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & tYear & "05'"
        'ExchangeRate5 = oCommand.ExecuteScalar()
        'If ExchangeRate5 = 0 Then ExchangeRate5 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & tYear & "06'"
        'ExchangeRate6 = oCommand.ExecuteScalar()
        'If ExchangeRate6 = 0 Then ExchangeRate6 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & tYear & "07'"
        'ExchangeRate7 = oCommand.ExecuteScalar()
        'If ExchangeRate7 = 0 Then ExchangeRate7 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & tYear & "08'"
        'ExchangeRate8 = oCommand.ExecuteScalar()
        'If ExchangeRate8 = 0 Then ExchangeRate8 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & tYear & "09'"
        'ExchangeRate9 = oCommand.ExecuteScalar()
        'If ExchangeRate9 = 0 Then ExchangeRate9 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & tYear & "10'"
        'ExchangeRate10 = oCommand.ExecuteScalar()
        'If ExchangeRate10 = 0 Then ExchangeRate10 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & tYear & "11'"
        'ExchangeRate11 = oCommand.ExecuteScalar()
        'If ExchangeRate11 = 0 Then ExchangeRate11 = 1
        'oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & tYear & "12'"
        'ExchangeRate12 = oCommand.ExecuteScalar()
        'If ExchangeRate12 = 0 Then ExchangeRate12 = 1
    End Sub

    Private Sub AdjustExcelFormat1_1()
        Ws.Cells(10, 1) = "Account"
        Ws.Cells(10, 2) = "Month"
        'oRng = Ws.Range("C10", "N10")
        If m_cnt = 1 Then
            oRng = Ws.Range("C10", "C10")
        End If
        If m_cnt = 2 Then
            oRng = Ws.Range("C10", "D10")
        End If
        If m_cnt = 3 Then
            oRng = Ws.Range("C10", "E10")
        End If
        If m_cnt = 4 Then
            oRng = Ws.Range("C10", "F10")
        End If
        If m_cnt = 5 Then
            oRng = Ws.Range("C10", "G10")
        End If
        If m_cnt = 6 Then
            oRng = Ws.Range("C10", "H10")
        End If
        If m_cnt = 7 Then
            oRng = Ws.Range("C10", "I10")
        End If
        If m_cnt = 8 Then
            oRng = Ws.Range("C10", "J10")
        End If
        If m_cnt = 9 Then
            oRng = Ws.Range("C10", "K10")
        End If
        If m_cnt = 10 Then
            oRng = Ws.Range("C10", "L10")
        End If
        If m_cnt = 11 Then
            oRng = Ws.Range("C10", "M10")
        End If
        If m_cnt = 12 Then
            oRng = Ws.Range("C10", "N10")
        End If
        oRng.NumberFormatLocal = "mmm-yy"
        'Ws.Cells(10, 3) = tDate.AddMonths(0)
        'Ws.Cells(10, 4) = tDate.AddMonths(1)
        'Ws.Cells(10, 5) = tDate.AddMonths(2)
        'Ws.Cells(10, 6) = tDate.AddMonths(3)
        'Ws.Cells(10, 7) = tDate.AddMonths(4)
        'Ws.Cells(10, 8) = tDate.AddMonths(5)
        'Ws.Cells(10, 9) = tDate.AddMonths(6)
        'Ws.Cells(10, 10) = tDate.AddMonths(7)
        'Ws.Cells(10, 11) = tDate.AddMonths(8)
        'Ws.Cells(10, 12) = tDate.AddMonths(9)
        'Ws.Cells(10, 13) = tDate.AddMonths(10)
        'Ws.Cells(10, 14) = tDate.AddMonths(11)
        'Ws.Cells(10, 15) = "YTD " & tYear
        For i = 1 To m_cnt
            Ws.Cells(10, i + 2) = tDate.AddMonths(i - 1)
        Next
        Ws.Cells(10, m_cnt + 3) = "YTD " & tYear

        ' 劃線
        'oRng = Ws.Range("B10", "O10")
        If m_cnt = 1 Then
            oRng = Ws.Range("B10", "D10")
        End If
        If m_cnt = 2 Then
            oRng = Ws.Range("B10", "E10")
        End If
        If m_cnt = 3 Then
            oRng = Ws.Range("B10", "F10")
        End If
        If m_cnt = 4 Then
            oRng = Ws.Range("B10", "G10")
        End If
        If m_cnt = 5 Then
            oRng = Ws.Range("B10", "H10")
        End If
        If m_cnt = 6 Then
            oRng = Ws.Range("B10", "I10")
        End If
        If m_cnt = 7 Then
            oRng = Ws.Range("B10", "J10")
        End If
        If m_cnt = 8 Then
            oRng = Ws.Range("B10", "K10")
        End If
        If m_cnt = 9 Then
            oRng = Ws.Range("B10", "L10")
        End If
        If m_cnt = 10 Then
            oRng = Ws.Range("B10", "M10")
        End If
        If m_cnt = 11 Then
            oRng = Ws.Range("B10", "N10")
        End If
        If m_cnt = 12 Then
            oRng = Ws.Range("B10", "O10")
        End If
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        'oRng = Ws.Range("C10", "O10")
        If m_cnt = 1 Then
            oRng = Ws.Range("C10", "D10")
        End If
        If m_cnt = 2 Then
            oRng = Ws.Range("C10", "E10")
        End If
        If m_cnt = 3 Then
            oRng = Ws.Range("C10", "F10")
        End If
        If m_cnt = 4 Then
            oRng = Ws.Range("C10", "G10")
        End If
        If m_cnt = 5 Then
            oRng = Ws.Range("C10", "H10")
        End If
        If m_cnt = 6 Then
            oRng = Ws.Range("C10", "I10")
        End If
        If m_cnt = 7 Then
            oRng = Ws.Range("C10", "J10")
        End If
        If m_cnt = 8 Then
            oRng = Ws.Range("C10", "K10")
        End If
        If m_cnt = 9 Then
            oRng = Ws.Range("C10", "L10")
        End If
        If m_cnt = 10 Then
            oRng = Ws.Range("C10", "M10")
        End If
        If m_cnt = 11 Then
            oRng = Ws.Range("C10", "N10")
        End If
        If m_cnt = 12 Then
            oRng = Ws.Range("C10", "O10")
        End If
        oRng.HorizontalAlignment = xlRight
        LineZ = 11
    End Sub

    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "group premium freight"
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

    'Private Function GetThisYearMonth_1(ByVal aag01 As String)
    '    oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
    '    oCommand2.CommandText += tYear & " and aah03 = 1"
    '    Dim TYTM As Decimal = oCommand2.ExecuteScalar()
    '    Return TYTM
    'End Function
    'Private Function GetThisYearMonth_2(ByVal aag01 As String)
    '    oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
    '    oCommand2.CommandText += tYear & " and aah03 = 2"
    '    Dim TYTM As Decimal = oCommand2.ExecuteScalar()
    '    Return TYTM
    'End Function
    'Private Function GetThisYearMonth_3(ByVal aag01 As String)
    '    oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
    '    oCommand2.CommandText += tYear & " and aah03 = 3"
    '    Dim TYTM As Decimal = oCommand2.ExecuteScalar()
    '    Return TYTM
    'End Function
    'Private Function GetThisYearMonth_4(ByVal aag01 As String)
    '    oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
    '    oCommand2.CommandText += tYear & " and aah03 = 4"
    '    Dim TYTM As Decimal = oCommand2.ExecuteScalar()
    '    Return TYTM
    'End Function
    'Private Function GetThisYearMonth_5(ByVal aag01 As String)
    '    oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
    '    oCommand2.CommandText += tYear & " and aah03 = 5"
    '    Dim TYTM As Decimal = oCommand2.ExecuteScalar()
    '    Return TYTM
    'End Function
    'Private Function GetThisYearMonth_6(ByVal aag01 As String)
    '    oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
    '    oCommand2.CommandText += tYear & " and aah03 = 6"
    '    Dim TYTM As Decimal = oCommand2.ExecuteScalar()
    '    Return TYTM
    'End Function
    'Private Function GetThisYearMonth_7(ByVal aag01 As String)
    '    oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
    '    oCommand2.CommandText += tYear & " and aah03 = 7"
    '    Dim TYTM As Decimal = oCommand2.ExecuteScalar()
    '    Return TYTM
    'End Function
    'Private Function GetThisYearMonth_8(ByVal aag01 As String)
    '    oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
    '    oCommand2.CommandText += tYear & " and aah03 = 8"
    '    Dim TYTM As Decimal = oCommand2.ExecuteScalar()
    '    Return TYTM
    'End Function
    'Private Function GetThisYearMonth_9(ByVal aag01 As String)
    '    oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
    '    oCommand2.CommandText += tYear & " and aah03 = 9"
    '    Dim TYTM As Decimal = oCommand2.ExecuteScalar()
    '    Return TYTM
    'End Function
    'Private Function GetThisYearMonth_10(ByVal aag01 As String)
    '    oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
    '    oCommand2.CommandText += tYear & " and aah03 = 10"
    '    Dim TYTM As Decimal = oCommand2.ExecuteScalar()
    '    Return TYTM
    'End Function
    'Private Function GetThisYearMonth_11(ByVal aag01 As String)
    '    oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
    '    oCommand2.CommandText += tYear & " and aah03 = 11"
    '    Dim TYTM As Decimal = oCommand2.ExecuteScalar()
    '    Return TYTM
    'End Function
    'Private Function GetThisYearMonth_12(ByVal aag01 As String)
    '    oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
    '    oCommand2.CommandText += tYear & " and aah03 = 12"
    '    Dim TYTM As Decimal = oCommand2.ExecuteScalar()
    '    Return TYTM
    'End Function
    Private Function GetThisYearMonth(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
        oCommand2.CommandText += tYear & " and aah03 = " & cMonth
        Dim TYTM As Decimal = oCommand2.ExecuteScalar()
        Return TYTM
    End Function
End Class