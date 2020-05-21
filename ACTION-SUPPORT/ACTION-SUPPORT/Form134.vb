Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form134
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand3 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand4 As New Oracle.ManagedDataAccess.Client.OracleCommand        '200422 add by Brady
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader3 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader4 As Oracle.ManagedDataAccess.Client.OracleDataReader          '200422 add by Brady
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim pYear As Int16 = 0
    Dim tDate As Date
    Dim tDate_1 As Date                        '190515 add by Brady
    Dim lYear As Int16 = 0
    Dim lMonth As Int16 = 0
    Dim LineZ As Integer = 0
    Dim DNP As String = String.Empty
    Dim DP_BOSS As String = String.Empty       '190404 add by Brady
    Dim ExchangeRate1 As Decimal = 0
    ' 2018/09/3
    Dim Ldate1 As Date
    Dim Ldate2 As Date
    ' 2018/09/04
    Dim TP As Decimal = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Dim l_aah04_05 As Decimal = 0              '200403 add by Brady
    Private Sub Form134_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
                oCommand3.Connection = oConnection
                oCommand3.CommandType = CommandType.Text
                oCommand4.Connection = oConnection                '200422 add by Brady
                oCommand4.CommandType = CommandType.Text          '200422 add by Brady
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        tYear = Me.DateTimePicker1.Value.Year
        tMonth = Me.DateTimePicker1.Value.Month
        pYear = Me.DateTimePicker1.Value.AddYears(-1).Year
        tDate = Convert.ToDateTime(tYear & "/" & tMonth & "/01")
        tDate_1 = tDate.AddDays(-1)    '190515 add by Brady
        tDate = tDate.AddMonths(1)     '190425 add by Brady
        tDate = tDate.AddDays(-1)      '190425 add by Brady
        pYear = tDate.AddYears(-1).Year
        lYear = Me.DateTimePicker1.Value.AddMonths(-1).Year
        lMonth = Me.DateTimePicker1.Value.AddMonths(-1).Month
        ExchangeRate1 = 1
        ' 2018/09/03
        Ldate1 = Convert.ToDateTime(tYear & "/01/01")
        Ldate2 = Convert.ToDateTime(tYear & "/" & tMonth & "/01").AddMonths(1).AddDays(-1)
        ExportToExcel()
        ExportToExcel_2()                     '200417 add by Brady 單獨再增加CS另一個報表
        'SaveExcel()
        'BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
        ExportToExcel_2()                     '200417 add by Brady 單獨再增加CS另一個報表
    End Sub
    Private Sub ExportToExcel()
        '190404 add by Brady
        'oCommand.CommandText = "select gem01,gem02 from gem_file where gemacti = 'Y'"
        oCommand.CommandText = "select gem01,gem02,gem06 from gem_file where gemacti = 'Y'"
        '190404 add by Brady END
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                TP = 0
                DNP = oReader.Item("gem01")

                '190404 add by Brady
                If Not oReader.Item("gem06") Is DBNull.Value Then
                    DP_BOSS = oReader.Item("gem06")
                End If
                '190404 add by Brady END
                
                xExcel = New Microsoft.Office.Interop.Excel.Application
                xWorkBook = xExcel.Workbooks.Add()
                oCommand2.CommandText = "SELECT nvl(sum(aao05 + aao06),0) FROM aao_file where aao03 = " & tYear & " and aao04 = " & tMonth & " and aao01 like '6601%' and aao02 = '" & oReader.Item("gem01") & "'"
                Dim PY As Decimal = oCommand2.ExecuteScalar()
                'If PY > 0 Then     '191121 mark by Brady
                TP += 1
                Ws = xWorkBook.Sheets(TP)
                Ws.Activate()
                AdjustExcelFormat1("Selling Exp")

                oCommand2.CommandText = "select aag01,aag02 from aag_file where aag01 like '6601%' and aag07 = 2 order by aag01"
                oReader2 = oCommand2.ExecuteReader()
                If oReader2.HasRows() Then
                    While oReader2.Read()
                        Ws.Cells(LineZ, 1) = oReader2.Item("aag01")
                        Ws.Cells(LineZ, 2) = oReader2.Item("aag02")
                        Ws.Cells(LineZ, 3) = GetDepartNmae(DNP)
                        Ws.Cells(LineZ, 4) = Decimal.Round(GetLastYearSameMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 5) = Decimal.Round(GetLastMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearSameMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 7) = GetThisYearSameMonthBudget(oReader2.Item("aag01").ToString(), DNP)
                        Ws.Cells(LineZ, 8) = "=F" & LineZ & "-G" & LineZ
                        Ws.Cells(LineZ, 9) = "=F" & LineZ & "-D" & LineZ
                        Ws.Cells(LineZ, 10) = "=F" & LineZ & "-E" & LineZ
                        Ws.Cells(LineZ, 11) = Decimal.Round(GetLastYearBeforeMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 12) = Decimal.Round(GetThisYearBeforeMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 13) = GetThisYearBeforeMonthBudget(oReader2.Item("aag01").ToString(), DNP)
                        Ws.Cells(LineZ, 14) = "=L" & LineZ & "-M" & LineZ
                        Ws.Cells(LineZ, 15) = "=L" & LineZ & "-K" & LineZ
                        Ws.Cells(LineZ, 16) = Decimal.Round(GetLastYearNoMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 17) = "=R" & LineZ & "-M" & LineZ & "+L" & LineZ
                        Ws.Cells(LineZ, 18) = GetThisYearBudget(oReader2.Item("aag01").ToString(), DNP)
                        LineZ += 1

                    End While
                    Ws.Cells(LineZ, 2) = "Total Selling Exp"
                    Ws.Cells(LineZ, 4) = "=SUM(D7:D" & LineZ - 1 & ")"
                    oRng = Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 4))
                    oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 18)), Type:=xlFillDefault)
                End If
                oReader2.Close()

                ' 劃線
                oRng = Ws.Range("A3", Ws.Cells(LineZ, 18))
                oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
                oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
                oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
                oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
                oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
                oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
                'End If           '191121 mark by Brady


                ' 第二頁
                oCommand2.CommandText = "SELECT nvl(sum(aao05 + aao06),0) FROM aao_file where aao03 = " & tYear & " and aao04 = " & tMonth & " and aao01 like '6604%' and aao02 = '" & oReader.Item("gem01") & "'"
                Dim PY1 As Decimal = oCommand2.ExecuteScalar()
                'If PY1 > 0 Then  '191121 mark by Brady
                TP += 1
                Ws = xWorkBook.Sheets(TP)
                Ws.Activate()
                AdjustExcelFormat1("RD Exp")

                oCommand2.CommandText = "select aag01,aag02 from aag_file where aag01 like '6604%' and aag07 = 2 order by aag01"
                oReader2 = oCommand2.ExecuteReader()
                If oReader2.HasRows() Then
                    While oReader2.Read()
                        Ws.Cells(LineZ, 1) = oReader2.Item("aag01")
                        Ws.Cells(LineZ, 2) = oReader2.Item("aag02")
                        Ws.Cells(LineZ, 3) = GetDepartNmae(DNP)
                        Ws.Cells(LineZ, 4) = Decimal.Round(GetLastYearSameMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 5) = Decimal.Round(GetLastMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearSameMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 7) = GetThisYearSameMonthBudget(oReader2.Item("aag01").ToString(), DNP)
                        Ws.Cells(LineZ, 8) = "=F" & LineZ & "-G" & LineZ
                        Ws.Cells(LineZ, 9) = "=F" & LineZ & "-D" & LineZ
                        Ws.Cells(LineZ, 10) = "=F" & LineZ & "-E" & LineZ
                        Ws.Cells(LineZ, 11) = Decimal.Round(GetLastYearBeforeMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 12) = Decimal.Round(GetThisYearBeforeMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 13) = GetThisYearBeforeMonthBudget(oReader2.Item("aag01").ToString(), DNP)
                        Ws.Cells(LineZ, 14) = "=L" & LineZ & "-M" & LineZ
                        Ws.Cells(LineZ, 15) = "=L" & LineZ & "-K" & LineZ
                        Ws.Cells(LineZ, 16) = Decimal.Round(GetLastYearNoMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 17) = "=R" & LineZ & "-M" & LineZ & "+L" & LineZ
                        Ws.Cells(LineZ, 18) = GetThisYearBudget(oReader2.Item("aag01").ToString(), DNP)
                        LineZ += 1

                    End While
                    Ws.Cells(LineZ, 2) = "Total RD Exp"
                    Ws.Cells(LineZ, 4) = "=SUM(D7:D" & LineZ - 1 & ")"
                    oRng = Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 4))
                    oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 18)), Type:=xlFillDefault)
                End If
                oReader2.Close()

                ' 劃線
                oRng = Ws.Range("A3", Ws.Cells(LineZ, 18))
                oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
                oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
                oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
                oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
                oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
                oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
                'End If                '191121 mark by Brady

                ' 第三頁
                oCommand2.CommandText = "SELECT nvl(sum(aao05 + aao06),0) FROM aao_file where aao03 = " & tYear & " and aao04 = " & tMonth & " and aao01 like '5101%' and aao02 = '" & oReader.Item("gem01") & "'"
                Dim PY3 As Decimal = oCommand2.ExecuteScalar()
                'If PY3 > 0 Then       '190515 mark by Brady
                TP += 1
                Ws = xWorkBook.Sheets(TP)
                Ws.Activate()
                AdjustExcelFormat1("Overhead")

                oCommand2.CommandText = "select aag01,aag02 from aag_file where aag01 like '5101%' and aag07 = 2 order by aag01"
                oReader2 = oCommand2.ExecuteReader()
                If oReader2.HasRows() Then
                    While oReader2.Read()
                        Ws.Cells(LineZ, 1) = oReader2.Item("aag01")
                        Ws.Cells(LineZ, 2) = oReader2.Item("aag02")
                        Ws.Cells(LineZ, 3) = GetDepartNmae(DNP)
                        Ws.Cells(LineZ, 4) = Decimal.Round(GetLastYearSameMonthA(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 5) = Decimal.Round(GetLastMonthA(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearSameMonthA(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 7) = GetThisYearSameMonthBudget(oReader2.Item("aag01").ToString(), DNP)
                        Ws.Cells(LineZ, 8) = "=F" & LineZ & "-G" & LineZ
                        Ws.Cells(LineZ, 9) = "=F" & LineZ & "-D" & LineZ
                        Ws.Cells(LineZ, 10) = "=F" & LineZ & "-E" & LineZ
                        Ws.Cells(LineZ, 11) = Decimal.Round(GetLastYearBeforeMonthA(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 12) = Decimal.Round(GetThisYearBeforeMonthA(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 13) = GetThisYearBeforeMonthBudget(oReader2.Item("aag01").ToString(), DNP)
                        Ws.Cells(LineZ, 14) = "=L" & LineZ & "-M" & LineZ
                        Ws.Cells(LineZ, 15) = "=L" & LineZ & "-K" & LineZ
                        Ws.Cells(LineZ, 16) = Decimal.Round(GetLastYearNoMonthA(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 17) = "=R" & LineZ & "-M" & LineZ & "+L" & LineZ
                        Ws.Cells(LineZ, 18) = GetThisYearBudget(oReader2.Item("aag01").ToString(), DNP)
                        LineZ += 1
                    End While
                    Ws.Cells(LineZ, 2) = "Total Overhead"
                    Ws.Cells(LineZ, 4) = "=SUM(D7:D" & LineZ - 1 & ")"
                    oRng = Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 4))
                    oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 18)), Type:=xlFillDefault)
                End If
                oReader2.Close()
                ' 劃線
                oRng = Ws.Range("A3", Ws.Cells(LineZ, 18))
                oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
                oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
                oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
                oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
                oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
                oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
                'End If                '190515 mark by Brady


                ' 第四頁
                oCommand2.CommandText = "SELECT nvl(sum(aao05 + aao06),0) FROM aao_file where aao03 = " & tYear & " and aao04 = " & tMonth & " and aao01 like '6602%' and aao02 = '" & oReader.Item("gem01") & "'"
                Dim PY2 As Decimal = oCommand2.ExecuteScalar()
                'If PY2 > 0 Then      '191121 mark by Brady
                TP += 1
                If TP > 3 Then
                    Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                Else
                    Ws = xWorkBook.Sheets(TP)
                End If
                'Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                Ws.Activate()
                AdjustExcelFormat1("ADM Exp")

                oCommand2.CommandText = "select aag01,aag02 from aag_file where aag01 like '6602%' and aag07 = 2 order by aag01"
                oReader2 = oCommand2.ExecuteReader()
                If oReader2.HasRows() Then
                    While oReader2.Read()
                        Ws.Cells(LineZ, 1) = oReader2.Item("aag01")
                        Ws.Cells(LineZ, 2) = oReader2.Item("aag02")
                        Ws.Cells(LineZ, 3) = GetDepartNmae(DNP)
                        Ws.Cells(LineZ, 4) = Decimal.Round(GetLastYearSameMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 5) = Decimal.Round(GetLastMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearSameMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 7) = GetThisYearSameMonthBudget(oReader2.Item("aag01").ToString(), DNP)
                        Ws.Cells(LineZ, 8) = "=F" & LineZ & "-G" & LineZ
                        Ws.Cells(LineZ, 9) = "=F" & LineZ & "-D" & LineZ
                        Ws.Cells(LineZ, 10) = "=F" & LineZ & "-E" & LineZ
                        Ws.Cells(LineZ, 11) = Decimal.Round(GetLastYearBeforeMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 12) = Decimal.Round(GetThisYearBeforeMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 13) = GetThisYearBeforeMonthBudget(oReader2.Item("aag01").ToString(), DNP)
                        Ws.Cells(LineZ, 14) = "=L" & LineZ & "-M" & LineZ
                        Ws.Cells(LineZ, 15) = "=L" & LineZ & "-K" & LineZ
                        Ws.Cells(LineZ, 16) = Decimal.Round(GetLastYearNoMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                        Ws.Cells(LineZ, 17) = "=R" & LineZ & "-M" & LineZ & "+L" & LineZ
                        Ws.Cells(LineZ, 18) = GetThisYearBudget(oReader2.Item("aag01").ToString(), DNP)
                        LineZ += 1

                    End While
                    Ws.Cells(LineZ, 2) = "Total ADM Exp"
                    Ws.Cells(LineZ, 4) = "=SUM(D7:D" & LineZ - 1 & ")"
                    oRng = Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 4))
                    oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 18)), Type:=xlFillDefault)
                End If
                oReader2.Close()


                ' 劃線
                oRng = Ws.Range("A3", Ws.Cells(LineZ, 18))
                oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
                oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
                oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
                oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
                oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
                oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
                'End If           '191121 mark by Brady

                ' 200401 add by Brady
                ' 第五頁 
                TP += 1
                If TP > 3 Then
                    Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                Else
                    Ws = xWorkBook.Sheets(TP)
                End If
                Ws.Activate()
                AdjustExcelFormat3("KPI")
                Ws.Cells(3, 1) = "2020/01/01"
                Ws.Cells(4, 1) = "2020/02/01"
                Ws.Cells(5, 1) = "2020/03/01"
                Ws.Cells(6, 1) = "2020/04/01"
                Ws.Cells(7, 1) = "2020/05/01"
                Ws.Cells(8, 1) = "2020/06/01"
                Ws.Cells(9, 1) = "2020/07/01"
                Ws.Cells(10, 1) = "2020/08/01"
                Ws.Cells(11, 1) = "2020/09/01"
                Ws.Cells(12, 1) = "2020/10/01"
                Ws.Cells(13, 1) = "2020/11/01"
                Ws.Cells(14, 1) = "2020/12/01"
                For L As Int16 = 1 To tMonth Step 1
                    Select Case L
                        Case 1
                            oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                            oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 1 AND aag01 in ('600101') "
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                l_aah04_05 = oReader2.Item("aah04_05")
                                If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                                Ws.Cells(3, 2) = l_aah04_05
                            End If
                            oReader2.Close()

                            Ws.Cells(3, 3) = 16648075.15
                            Ws.Cells(3, 4) = "=B3/C3"

                            Dim l_abb07 As Decimal = 0
                            Dim t_abb07 As Decimal = 0
                            oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/01/01'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/01/31'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                            oCommand2.CommandText += DNP & "' order by abb03"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                While oReader2.Read()
                                    oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/01/01'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/01/31'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                                    oReader3 = oCommand3.ExecuteReader()
                                    If oReader3.HasRows() Then
                                        While oReader3.Read()
                                            If oReader3.Item("abb06") = 1 Then
                                                ' 借方
                                                l_abb07 = oReader3.Item("abb07")
                                                t_abb07 = t_abb07 + l_abb07
                                            Else
                                                '貸方
                                                l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                                t_abb07 = t_abb07 + l_abb07
                                            End If
                                        End While
                                    End If
                                    oReader3.Close()
                                End While
                                Ws.Cells(3, 5) = t_abb07
                            End If
                            oReader2.Close()

                            oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                            oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 1 AND tc_bud08 = '" & DNP & "'"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                Ws.Cells(3, 6) = oReader2.Item("s_tc_bud13")
                            End If
                            oReader2.Close()

                            Ws.Cells(3, 7) = "=E3-F3"
                            Ws.Cells(3, 8) = "=(E3-F3)/F3"
                            Ws.Cells(3, 9) = "=D3*F3"
                            Ws.Cells(3, 10) = "=E3-I3"
                            Ws.Cells(3, 11) = "=(E3-I3)/I3"
                        Case 2
                            oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                            oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 2 AND aag01 in ('600101') "
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                l_aah04_05 = oReader2.Item("aah04_05")
                                If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                                Ws.Cells(4, 2) = l_aah04_05
                            End If
                            oReader2.Close()

                            Ws.Cells(4, 3) = 18577805.4844597
                            Ws.Cells(4, 4) = "=B4/C4"

                            Dim l_abb07 As Decimal = 0
                            Dim t_abb07 As Decimal = 0
                            oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/02/01'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/02/29'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                            oCommand2.CommandText += DNP & "' order by abb03"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                While oReader2.Read()
                                    oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/02/01'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/02/29'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                                    oReader3 = oCommand3.ExecuteReader()
                                    If oReader3.HasRows() Then
                                        While oReader3.Read()
                                            If oReader3.Item("abb06") = 1 Then
                                                ' 借方
                                                l_abb07 = oReader3.Item("abb07")
                                                t_abb07 = t_abb07 + l_abb07
                                            Else
                                                '貸方
                                                l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                                t_abb07 = t_abb07 + l_abb07
                                            End If
                                        End While
                                    End If
                                    oReader3.Close()
                                End While
                                Ws.Cells(4, 5) = t_abb07
                            End If
                            oReader2.Close()

                            oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                            oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 2 AND tc_bud08 = '" & DNP & "'"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                Ws.Cells(4, 6) = oReader2.Item("s_tc_bud13")
                            End If
                            oReader2.Close()

                            Ws.Cells(4, 7) = "=E4-F4"
                            Ws.Cells(4, 8) = "=(E4-F4)/F4"
                            Ws.Cells(4, 9) = "=D4*F4"
                            Ws.Cells(4, 10) = "=E4-I4"
                            Ws.Cells(4, 11) = "=(E4-I4)/I4"
                        Case 3
                            oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                            oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 3 AND aag01 in ('600101') "
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                l_aah04_05 = oReader2.Item("aah04_05")
                                If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                                Ws.Cells(5, 2) = l_aah04_05
                            End If
                            oReader2.Close()

                            Ws.Cells(5, 3) = 16799152.1945558
                            Ws.Cells(5, 4) = "=B5/C5"

                            Dim l_abb07 As Decimal = 0
                            Dim t_abb07 As Decimal = 0
                            oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/03/01'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/03/31'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                            oCommand2.CommandText += DNP & "' order by abb03"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                While oReader2.Read()
                                    oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/03/01'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/03/31'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                                    oReader3 = oCommand3.ExecuteReader()
                                    If oReader3.HasRows() Then
                                        While oReader3.Read()
                                            If oReader3.Item("abb06") = 1 Then
                                                ' 借方
                                                l_abb07 = oReader3.Item("abb07")
                                                t_abb07 = t_abb07 + l_abb07
                                            Else
                                                '貸方
                                                l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                                t_abb07 = t_abb07 + l_abb07
                                            End If
                                        End While
                                    End If
                                    oReader3.Close()
                                End While
                                Ws.Cells(5, 5) = t_abb07
                            End If
                            oReader2.Close()

                            oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                            oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 3 AND tc_bud08 = '" & DNP & "'"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                Ws.Cells(5, 6) = oReader2.Item("s_tc_bud13")
                            End If
                            oReader2.Close()

                            Ws.Cells(5, 7) = "=E5-F5"
                            Ws.Cells(5, 8) = "=(E5-F5)/F5"
                            Ws.Cells(5, 9) = "=D5*F5"
                            Ws.Cells(5, 10) = "=E5-I5"
                            Ws.Cells(5, 11) = "=(E5-I5)/I5"
                        Case 4
                            oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                            oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 4 AND aag01 in ('600101') "
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                l_aah04_05 = oReader2.Item("aah04_05")
                                If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                                Ws.Cells(6, 2) = l_aah04_05
                            End If
                            oReader2.Close()

                            Ws.Cells(6, 3) = 16262202.8683268
                            Ws.Cells(6, 4) = "=B6/C6"

                            Dim l_abb07 As Decimal = 0
                            Dim t_abb07 As Decimal = 0
                            oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/04/01'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/04/30'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                            oCommand2.CommandText += DNP & "' order by abb03"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                While oReader2.Read()
                                    oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/04/01'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/04/30'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                                    oReader3 = oCommand3.ExecuteReader()
                                    If oReader3.HasRows() Then
                                        While oReader3.Read()
                                            If oReader3.Item("abb06") = 1 Then
                                                ' 借方
                                                l_abb07 = oReader3.Item("abb07")
                                                t_abb07 = t_abb07 + l_abb07
                                            Else
                                                '貸方
                                                l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                                t_abb07 = t_abb07 + l_abb07
                                            End If
                                        End While
                                    End If
                                    oReader3.Close()
                                End While
                                Ws.Cells(6, 5) = t_abb07
                            End If
                            oReader2.Close()

                            oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                            oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 4 AND tc_bud08 = '" & DNP & "'"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                Ws.Cells(6, 6) = oReader2.Item("s_tc_bud13")
                            End If
                            oReader2.Close()

                            Ws.Cells(6, 7) = "=E6-F6"
                            Ws.Cells(6, 8) = "=(E6-F6)/F6"
                            Ws.Cells(6, 9) = "=D6*F6"
                            Ws.Cells(6, 10) = "=E6-I6"
                            Ws.Cells(6, 11) = "=(E6-I6)/I6"
                        Case 5
                            oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                            oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 5 AND aag01 in ('600101') "
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                l_aah04_05 = oReader2.Item("aah04_05")
                                If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                                Ws.Cells(7, 2) = l_aah04_05
                            End If
                            oReader2.Close()

                            Ws.Cells(7, 3) = 18198819.4663555
                            Ws.Cells(7, 4) = "=B7/C7"

                            Dim l_abb07 As Decimal = 0
                            Dim t_abb07 As Decimal = 0
                            oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/05/01'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/05/31'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                            oCommand2.CommandText += DNP & "' order by abb03"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                While oReader2.Read()
                                    oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/05/01'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/05/31'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                                    oReader3 = oCommand3.ExecuteReader()
                                    If oReader3.HasRows() Then
                                        While oReader3.Read()
                                            If oReader3.Item("abb06") = 1 Then
                                                ' 借方
                                                l_abb07 = oReader3.Item("abb07")
                                                t_abb07 = t_abb07 + l_abb07
                                            Else
                                                '貸方
                                                l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                                t_abb07 = t_abb07 + l_abb07
                                            End If
                                        End While
                                    End If
                                    oReader3.Close()
                                End While
                                Ws.Cells(7, 5) = t_abb07
                            End If
                            oReader2.Close()

                            oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                            oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 5 AND tc_bud08 = '" & DNP & "'"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                Ws.Cells(7, 6) = oReader2.Item("s_tc_bud13")
                            End If
                            oReader2.Close()

                            Ws.Cells(7, 7) = "=E7-F7"
                            Ws.Cells(7, 8) = "=(E7-F7)/F7"
                            Ws.Cells(7, 9) = "=D7*F7"
                            Ws.Cells(7, 10) = "=E7-I7"
                            Ws.Cells(7, 11) = "=(E7-I7)/I7"
                        Case 6
                            oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                            oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 6 AND aag01 in ('600101') "
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                l_aah04_05 = oReader2.Item("aah04_05")
                                If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                                Ws.Cells(8, 2) = l_aah04_05
                            End If
                            oReader2.Close()

                            Ws.Cells(8, 3) = 11654412.3220121
                            Ws.Cells(8, 4) = "=B8/C8"

                            Dim l_abb07 As Decimal = 0
                            Dim t_abb07 As Decimal = 0
                            oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/06/01'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/06/30'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                            oCommand2.CommandText += DNP & "' order by abb03"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                While oReader2.Read()
                                    oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/06/01'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/06/30'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                                    oReader3 = oCommand3.ExecuteReader()
                                    If oReader3.HasRows() Then
                                        While oReader3.Read()
                                            If oReader3.Item("abb06") = 1 Then
                                                ' 借方
                                                l_abb07 = oReader3.Item("abb07")
                                                t_abb07 = t_abb07 + l_abb07
                                            Else
                                                '貸方
                                                l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                                t_abb07 = t_abb07 + l_abb07
                                            End If
                                        End While
                                    End If
                                    oReader3.Close()
                                End While
                                Ws.Cells(8, 5) = t_abb07
                            End If
                            oReader2.Close()

                            oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                            oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 6 AND tc_bud08 = '" & DNP & "'"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                Ws.Cells(8, 6) = oReader2.Item("s_tc_bud13")
                            End If
                            oReader2.Close()

                            Ws.Cells(8, 7) = "=E8-F8"
                            Ws.Cells(8, 8) = "=(E8-F8)/F8"
                            Ws.Cells(8, 9) = "=D8*F8"
                            Ws.Cells(8, 10) = "=E8-I8"
                            Ws.Cells(8, 11) = "=(E8-I8)/I8"
                        Case 7
                            oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                            oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 7 AND aag01 in ('600101') "
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                l_aah04_05 = oReader2.Item("aah04_05")
                                If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                                Ws.Cells(9, 2) = l_aah04_05
                            End If
                            oReader2.Close()

                            Ws.Cells(9, 3) = 12880421.7803447
                            Ws.Cells(9, 4) = "=B9/C9"

                            Dim l_abb07 As Decimal = 0
                            Dim t_abb07 As Decimal = 0
                            oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/07/01'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/07/31'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                            oCommand2.CommandText += DNP & "' order by abb03"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                While oReader2.Read()
                                    oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/07/01'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/07/31'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                                    oReader3 = oCommand3.ExecuteReader()
                                    If oReader3.HasRows() Then
                                        While oReader3.Read()
                                            If oReader3.Item("abb06") = 1 Then
                                                ' 借方
                                                l_abb07 = oReader3.Item("abb07")
                                                t_abb07 = t_abb07 + l_abb07
                                            Else
                                                '貸方
                                                l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                                t_abb07 = t_abb07 + l_abb07
                                            End If
                                        End While
                                    End If
                                    oReader3.Close()
                                End While
                                Ws.Cells(9, 5) = t_abb07
                            End If
                            oReader2.Close()

                            oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                            oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 7 AND tc_bud08 = '" & DNP & "'"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                Ws.Cells(9, 6) = oReader2.Item("s_tc_bud13")
                            End If
                            oReader2.Close()

                            Ws.Cells(9, 7) = "=E9-F9"
                            Ws.Cells(9, 8) = "=(E9-F9)/F9"
                            Ws.Cells(9, 9) = "=D9*F9"
                            Ws.Cells(9, 10) = "=E9-I9"
                            Ws.Cells(9, 11) = "=(E9-I9)/I9"
                        Case 8
                            oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                            oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 8 AND aag01 in ('600101') "
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                l_aah04_05 = oReader2.Item("aah04_05")
                                If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                                Ws.Cells(10, 2) = l_aah04_05
                            End If
                            oReader2.Close()

                            Ws.Cells(10, 3) = 18073323.7356463
                            Ws.Cells(10, 4) = "=B10/C10"

                            Dim l_abb07 As Decimal = 0
                            Dim t_abb07 As Decimal = 0
                            oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/08/01'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/08/31'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                            oCommand2.CommandText += DNP & "' order by abb03"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                While oReader2.Read()
                                    oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/08/01'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/08/31'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                                    oReader3 = oCommand3.ExecuteReader()
                                    If oReader3.HasRows() Then
                                        While oReader3.Read()
                                            If oReader3.Item("abb06") = 1 Then
                                                ' 借方
                                                l_abb07 = oReader3.Item("abb07")
                                                t_abb07 = t_abb07 + l_abb07
                                            Else
                                                '貸方
                                                l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                                t_abb07 = t_abb07 + l_abb07
                                            End If
                                        End While
                                    End If
                                    oReader3.Close()
                                End While
                                Ws.Cells(10, 5) = t_abb07
                            End If
                            oReader2.Close()

                            oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                            oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 8 AND tc_bud08 = '" & DNP & "'"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                Ws.Cells(10, 6) = oReader2.Item("s_tc_bud13")
                            End If
                            oReader2.Close()

                            Ws.Cells(10, 7) = "=E10-F10"
                            Ws.Cells(10, 8) = "=(E10-F10)/F10"
                            Ws.Cells(10, 9) = "=D10*F10"
                            Ws.Cells(10, 10) = "=E10-I10"
                            Ws.Cells(10, 11) = "=(E10-I10)/I10"
                        Case 9
                            oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                            oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 9 AND aag01 in ('600101') "
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                l_aah04_05 = oReader2.Item("aah04_05")
                                If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                                Ws.Cells(11, 2) = l_aah04_05
                            End If
                            oReader2.Close()

                            Ws.Cells(11, 3) = 13379917.3701231
                            Ws.Cells(11, 4) = "=B11/C11"

                            Dim l_abb07 As Decimal = 0
                            Dim t_abb07 As Decimal = 0
                            oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/09/01'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/09/30'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                            oCommand2.CommandText += DNP & "' order by abb03"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                While oReader2.Read()
                                    oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/09/01'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/09/30'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                                    oReader3 = oCommand3.ExecuteReader()
                                    If oReader3.HasRows() Then
                                        While oReader3.Read()
                                            If oReader3.Item("abb06") = 1 Then
                                                ' 借方
                                                l_abb07 = oReader3.Item("abb07")
                                                t_abb07 = t_abb07 + l_abb07
                                            Else
                                                '貸方
                                                l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                                t_abb07 = t_abb07 + l_abb07
                                            End If
                                        End While
                                    End If
                                    oReader3.Close()
                                End While
                                Ws.Cells(11, 5) = t_abb07
                            End If
                            oReader2.Close()

                            oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                            oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 9 AND tc_bud08 = '" & DNP & "'"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                Ws.Cells(11, 6) = oReader2.Item("s_tc_bud13")
                            End If
                            oReader2.Close()

                            Ws.Cells(11, 7) = "=E11-F11"
                            Ws.Cells(11, 8) = "=(E11-F11)/F11"
                            Ws.Cells(11, 9) = "=D11*F11"
                            Ws.Cells(11, 10) = "=E11-I11"
                            Ws.Cells(11, 11) = "=(E11-I11)/I11"
                        Case 10
                            oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                            oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 10 AND aag01 in ('600101') "
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                l_aah04_05 = oReader2.Item("aah04_05")
                                If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                                Ws.Cells(12, 2) = l_aah04_05
                            End If
                            oReader2.Close()

                            Ws.Cells(12, 3) = 14633836.7242984
                            Ws.Cells(12, 4) = "=B12/C12"

                            Dim l_abb07 As Decimal = 0
                            Dim t_abb07 As Decimal = 0
                            oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/10/01'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/10/31'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                            oCommand2.CommandText += DNP & "' order by abb03"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                While oReader2.Read()
                                    oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/10/01'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/10/31'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                                    oReader3 = oCommand3.ExecuteReader()
                                    If oReader3.HasRows() Then
                                        While oReader3.Read()
                                            If oReader3.Item("abb06") = 1 Then
                                                ' 借方
                                                l_abb07 = oReader3.Item("abb07")
                                                t_abb07 = t_abb07 + l_abb07
                                            Else
                                                '貸方
                                                l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                                t_abb07 = t_abb07 + l_abb07
                                            End If
                                        End While
                                    End If
                                    oReader3.Close()
                                End While
                                Ws.Cells(12, 5) = t_abb07
                            End If
                            oReader2.Close()

                            oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                            oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 10 AND tc_bud08 = '" & DNP & "'"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                Ws.Cells(12, 6) = oReader2.Item("s_tc_bud13")
                            End If
                            oReader2.Close()

                            Ws.Cells(12, 7) = "=E12-F12"
                            Ws.Cells(12, 8) = "=(E12-F12)/F12"
                            Ws.Cells(12, 9) = "=D12*F12"
                            Ws.Cells(12, 10) = "=E12-I12"
                            Ws.Cells(12, 11) = "=(E12-I12)/I12"
                        Case 11
                            oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                            oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 11 AND aag01 in ('600101') "
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                l_aah04_05 = oReader2.Item("aah04_05")
                                If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                                Ws.Cells(13, 2) = l_aah04_05
                            End If
                            oReader2.Close()

                            Ws.Cells(13, 3) = 9739716.54373806
                            Ws.Cells(13, 4) = "=B13/C13"

                            Dim l_abb07 As Decimal = 0
                            Dim t_abb07 As Decimal = 0
                            oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/11/01'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/11/30'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                            oCommand2.CommandText += DNP & "' order by abb03"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                While oReader2.Read()
                                    oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/11/01'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/11/30'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                                    oReader3 = oCommand3.ExecuteReader()
                                    If oReader3.HasRows() Then
                                        While oReader3.Read()
                                            If oReader3.Item("abb06") = 1 Then
                                                ' 借方
                                                l_abb07 = oReader3.Item("abb07")
                                                t_abb07 = t_abb07 + l_abb07
                                            Else
                                                '貸方
                                                l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                                t_abb07 = t_abb07 + l_abb07
                                            End If
                                        End While
                                    End If
                                    oReader3.Close()
                                End While
                                Ws.Cells(13, 5) = t_abb07
                            End If
                            oReader2.Close()

                            oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                            oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 11 AND tc_bud08 = '" & DNP & "'"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                Ws.Cells(13, 6) = oReader2.Item("s_tc_bud13")
                            End If
                            oReader2.Close()

                            Ws.Cells(13, 7) = "=E13-F13"
                            Ws.Cells(13, 8) = "=(E13-F13)/F13"
                            Ws.Cells(13, 9) = "=D13*F13"
                            Ws.Cells(13, 10) = "=E13-I13"
                            Ws.Cells(13, 11) = "=(E13-I13)/I13"
                        Case 12
                            oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                            oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 12 AND aag01 in ('600101') "
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                l_aah04_05 = oReader2.Item("aah04_05")
                                If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                                Ws.Cells(14, 2) = l_aah04_05
                            End If
                            oReader2.Close()

                            Ws.Cells(14, 3) = 11758282.582403
                            Ws.Cells(14, 4) = "=B14/C14"

                            Dim l_abb07 As Decimal = 0
                            Dim t_abb07 As Decimal = 0
                            oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/12/01'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/12/31'"
                            oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                            oCommand2.CommandText += DNP & "' order by abb03"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                While oReader2.Read()
                                    oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/12/01'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/12/31'"
                                    oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                                    oReader3 = oCommand3.ExecuteReader()
                                    If oReader3.HasRows() Then
                                        While oReader3.Read()
                                            If oReader3.Item("abb06") = 1 Then
                                                ' 借方
                                                l_abb07 = oReader3.Item("abb07")
                                                t_abb07 = t_abb07 + l_abb07
                                            Else
                                                '貸方
                                                l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                                t_abb07 = t_abb07 + l_abb07
                                            End If
                                        End While
                                    End If
                                    oReader3.Close()
                                End While
                                Ws.Cells(14, 5) = t_abb07
                            End If
                            oReader2.Close()

                            oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                            oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 12 AND tc_bud08 = '" & DNP & "'"
                            oReader2 = oCommand2.ExecuteReader()
                            If oReader2.HasRows() Then
                                oReader2.Read()
                                Ws.Cells(14, 6) = oReader2.Item("s_tc_bud13")
                            End If
                            oReader2.Close()

                            Ws.Cells(14, 7) = "=E14-F14"
                            Ws.Cells(14, 8) = "=(E14-F14)/F14"
                            Ws.Cells(14, 9) = "=D14*F14"
                            Ws.Cells(14, 10) = "=E14-I14"
                            Ws.Cells(14, 11) = "=(E14-I14)/I14"
                    End Select
                Next
                Ws.Cells(15, 2) = "=SUM(B3:B14)"
                Ws.Cells(15, 3) = "=SUM(C3:C14)"
                Ws.Cells(15, 4) = "=B15/C15"
                Ws.Cells(15, 5) = "=SUM(E3:E14)"
                Ws.Cells(15, 6) = "=SUM(F3:F14)"
                Ws.Cells(15, 7) = "=SUM(G3:G14)"
                Ws.Cells(15, 8) = "=G15/F15"
                Ws.Cells(15, 9) = "=SUM(I3:I14)"
                Ws.Cells(15, 10) = "=SUM(J3:J14)"
                Ws.Cells(15, 11) = "=J15/I15"
                ' 200401 add by Brady END

                ' 200401 Note By Brady 變成第六頁 
                ' 第五頁  -->改為按月區分 20191011
                For z As Int16 = Ldate1.Month To Ldate2.Month Step 1
                    TP += 1
                    If TP > 3 Then
                        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                    Else
                        Ws = xWorkBook.Sheets(TP)
                    End If
                    Ws.Activate()
                    AdjustExcelFormat2(z)
                    oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('"
                    oCommand2.CommandText += Ldate1.AddMonths(z - 1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                    oCommand2.CommandText += Ldate1.AddMonths(z).AddDays(-1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                    oCommand2.CommandText += DNP & "' order by abb03"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        While oReader2.Read()
                            '' 期初
                            'oCommand3.CommandText = "select nvl((aao05 -aao06),0) from aao_file where aao03 = " & tYear & " and aao04 = 0 and aao01 = '" & oReader2.Item("abb03") & "' and aao02 = '" & DNP & "'"
                            'Dim St1 As Decimal = oCommand3.ExecuteScalar()
                            'Dim St2 As Decimal = 0
                            'Dim St3 As Decimal = 0
                            'Ws.Cells(LineZ, 2) = "期初"
                            'Ws.Cells(LineZ, 10) = St1
                            'LineZ += 1
                            ' 之後
                            oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('"
                            oCommand3.CommandText += Ldate1.AddMonths(z - 1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                            oCommand3.CommandText += Ldate1.AddMonths(z).AddDays(-1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                            oReader3 = oCommand3.ExecuteReader()
                            If oReader3.HasRows() Then
                                While oReader3.Read()
                                    Ws.Cells(LineZ, 1) = DNP
                                    Ws.Cells(LineZ, 2) = oReader.Item("gem02")
                                    Ws.Cells(LineZ, 3) = oReader3.Item("aba02")
                                    Ws.Cells(LineZ, 4) = oReader3.Item("abb01")
                                    Ws.Cells(LineZ, 5) = oReader2.Item("abb03")
                                    Ws.Cells(LineZ, 6) = oReader3.Item("aag02")
                                    Ws.Cells(LineZ, 7) = oReader3.Item("abb04")
                                    If oReader3.Item("abb06") = 1 Then
                                        ' 借方
                                        'Ws.Cells(LineZ, 8) = oReader3.Item("abb07")
                                        Ws.Cells(LineZ, 8) = oReader3.Item("abb07")
                                        'St1 += oReader3.Item("abb07")
                                        'St2 += oReader3.Item("abb07")
                                    Else
                                        '貸方
                                        'Ws.Cells(LineZ, 9) = oReader3.Item("abb07")
                                        Ws.Cells(LineZ, 8) = oReader3.Item("abb07") * Decimal.MinusOne
                                        'St1 -= oReader3.Item("abb07")
                                        'St3 += oReader3.Item("abb07")
                                    End If
                                    'Ws.Cells(LineZ, 10) = St1
                                    LineZ += 1
                                End While
                                'Ws.Cells(LineZ, 3) = "合    计:   "
                                'Ws.Cells(LineZ, 8) = St2
                                'Ws.Cells(LineZ, 9) = St3
                                'LineZ += 1
                            End If
                            oReader3.Close()
                        End While
                        'Ws.Cells(LineZ, 3) = "部门合计:"
                        'Ws.Cells(LineZ, 8) = "=SUM(H3:H" & LineZ - 1 & ")"
                        'Ws.Cells(LineZ, 9) = "=SUM(I3:I" & LineZ - 1 & ")"
                    End If
                    oReader2.Close()
                    oRng = Ws.Range("A1", "G1")
                    oRng.EntireColumn.AutoFit()

                Next
                'SaveExcel(oReader.Item("gem02"))                         '200411 mark by Brady 
                SaveExcel(oReader.Item("gem02"), oReader.Item("gem01"))   '200411 add by Brady
            End While
        End If
        oReader.Close()
        'MsgBox("Finished")
    End Sub
    '200417 add by Brady 單獨再增加CS另一個報表
    Private Sub ExportToExcel_2()
        TP = 0
        DNP = "D0210"
        DP_BOSS = "Chris Ma"

        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        oCommand2.CommandText = "SELECT nvl(sum(aao05 + aao06),0) FROM aao_file where aao03 = " & tYear & " and aao04 = " & tMonth & " and aao01 like '6601%' and aao02 = 'D0210'"
        Dim PY As Decimal = oCommand2.ExecuteScalar()
        'If PY > 0 Then     
        TP += 1
        Ws = xWorkBook.Sheets(TP)
        Ws.Activate()
        AdjustExcelFormat1("Selling Exp")

        oCommand2.CommandText = "select aag01,aag02 from aag_file where aag01 like '6601%' and aag07 = 2 order by aag01"
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                Ws.Cells(LineZ, 1) = oReader2.Item("aag01")
                Ws.Cells(LineZ, 2) = oReader2.Item("aag02")
                Ws.Cells(LineZ, 3) = GetDepartNmae(DNP)
                Ws.Cells(LineZ, 4) = Decimal.Round(GetLastYearSameMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 5) = Decimal.Round(GetLastMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearSameMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3) - Decimal.Round(GetThisYearSameMonth_D0210(oReader2.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 7) = GetThisYearSameMonthBudget(oReader2.Item("aag01").ToString(), DNP)
                Ws.Cells(LineZ, 8) = "=F" & LineZ & "-G" & LineZ
                Ws.Cells(LineZ, 9) = "=F" & LineZ & "-D" & LineZ
                Ws.Cells(LineZ, 10) = "=F" & LineZ & "-E" & LineZ
                Ws.Cells(LineZ, 11) = Decimal.Round(GetLastYearBeforeMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 12) = Decimal.Round(GetThisYearBeforeMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3) - Decimal.Round(GetThisYearSameMonth_D0210(oReader2.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 13) = GetThisYearBeforeMonthBudget(oReader2.Item("aag01").ToString(), DNP)
                Ws.Cells(LineZ, 14) = "=L" & LineZ & "-M" & LineZ
                Ws.Cells(LineZ, 15) = "=L" & LineZ & "-K" & LineZ
                Ws.Cells(LineZ, 16) = Decimal.Round(GetLastYearNoMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 17) = "=R" & LineZ & "-M" & LineZ & "+L" & LineZ
                Ws.Cells(LineZ, 18) = GetThisYearBudget(oReader2.Item("aag01").ToString(), DNP)
                LineZ += 1

            End While
            Ws.Cells(LineZ, 2) = "Total Selling Exp"
            Ws.Cells(LineZ, 4) = "=SUM(D7:D" & LineZ - 1 & ")"
            oRng = Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 4))
            oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 18)), Type:=xlFillDefault)
        End If
        oReader2.Close()

        ' 劃線
        oRng = Ws.Range("A3", Ws.Cells(LineZ, 18))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        'End If           


        ' 第二頁
        oCommand2.CommandText = "SELECT nvl(sum(aao05 + aao06),0) FROM aao_file where aao03 = " & tYear & " and aao04 = " & tMonth & " and aao01 like '6604%' and aao02 = 'D0210'"
        Dim PY1 As Decimal = oCommand2.ExecuteScalar()
        'If PY1 > 0 Then  
        TP += 1
        Ws = xWorkBook.Sheets(TP)
        Ws.Activate()
        AdjustExcelFormat1("RD Exp")

        oCommand2.CommandText = "select aag01,aag02 from aag_file where aag01 like '6604%' and aag07 = 2 order by aag01"
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                Ws.Cells(LineZ, 1) = oReader2.Item("aag01")
                Ws.Cells(LineZ, 2) = oReader2.Item("aag02")
                Ws.Cells(LineZ, 3) = GetDepartNmae(DNP)
                Ws.Cells(LineZ, 4) = Decimal.Round(GetLastYearSameMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 5) = Decimal.Round(GetLastMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearSameMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3) - Decimal.Round(GetThisYearSameMonth_D0210(oReader2.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 7) = GetThisYearSameMonthBudget(oReader2.Item("aag01").ToString(), DNP)
                Ws.Cells(LineZ, 8) = "=F" & LineZ & "-G" & LineZ
                Ws.Cells(LineZ, 9) = "=F" & LineZ & "-D" & LineZ
                Ws.Cells(LineZ, 10) = "=F" & LineZ & "-E" & LineZ
                Ws.Cells(LineZ, 11) = Decimal.Round(GetLastYearBeforeMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 12) = Decimal.Round(GetThisYearBeforeMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 13) = GetThisYearBeforeMonthBudget(oReader2.Item("aag01").ToString(), DNP)
                Ws.Cells(LineZ, 14) = "=L" & LineZ & "-M" & LineZ
                Ws.Cells(LineZ, 15) = "=L" & LineZ & "-K" & LineZ
                Ws.Cells(LineZ, 16) = Decimal.Round(GetLastYearNoMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 17) = "=R" & LineZ & "-M" & LineZ & "+L" & LineZ
                Ws.Cells(LineZ, 18) = GetThisYearBudget(oReader2.Item("aag01").ToString(), DNP)
                LineZ += 1

            End While
            Ws.Cells(LineZ, 2) = "Total RD Exp"
            Ws.Cells(LineZ, 4) = "=SUM(D7:D" & LineZ - 1 & ")"
            oRng = Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 4))
            oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 18)), Type:=xlFillDefault)
        End If
        oReader2.Close()

        ' 劃線
        oRng = Ws.Range("A3", Ws.Cells(LineZ, 18))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        'End If                

        ' 第三頁
        oCommand2.CommandText = "SELECT nvl(sum(aao05 + aao06),0) FROM aao_file where aao03 = " & tYear & " and aao04 = " & tMonth & " and aao01 like '5101%' and aao02 = 'D0210'"
        Dim PY3 As Decimal = oCommand2.ExecuteScalar()
        'If PY3 > 0 Then       
        TP += 1
        Ws = xWorkBook.Sheets(TP)
        Ws.Activate()
        AdjustExcelFormat1("Overhead")

        oCommand2.CommandText = "select aag01,aag02 from aag_file where aag01 like '5101%' and aag07 = 2 order by aag01"
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                Ws.Cells(LineZ, 1) = oReader2.Item("aag01")
                Ws.Cells(LineZ, 2) = oReader2.Item("aag02")
                Ws.Cells(LineZ, 3) = GetDepartNmae(DNP)
                Ws.Cells(LineZ, 4) = Decimal.Round(GetLastYearSameMonthA(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 5) = Decimal.Round(GetLastMonthA(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearSameMonthA(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3) - Decimal.Round(GetThisYearSameMonth_D0210(oReader2.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 7) = GetThisYearSameMonthBudget(oReader2.Item("aag01").ToString(), DNP)
                Ws.Cells(LineZ, 8) = "=F" & LineZ & "-G" & LineZ
                Ws.Cells(LineZ, 9) = "=F" & LineZ & "-D" & LineZ
                Ws.Cells(LineZ, 10) = "=F" & LineZ & "-E" & LineZ
                Ws.Cells(LineZ, 11) = Decimal.Round(GetLastYearBeforeMonthA(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 12) = Decimal.Round(GetThisYearBeforeMonthA(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 13) = GetThisYearBeforeMonthBudget(oReader2.Item("aag01").ToString(), DNP)
                Ws.Cells(LineZ, 14) = "=L" & LineZ & "-M" & LineZ
                Ws.Cells(LineZ, 15) = "=L" & LineZ & "-K" & LineZ
                Ws.Cells(LineZ, 16) = Decimal.Round(GetLastYearNoMonthA(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 17) = "=R" & LineZ & "-M" & LineZ & "+L" & LineZ
                Ws.Cells(LineZ, 18) = GetThisYearBudget(oReader2.Item("aag01").ToString(), DNP)
                LineZ += 1
            End While
            Ws.Cells(LineZ, 2) = "Total Overhead"
            Ws.Cells(LineZ, 4) = "=SUM(D7:D" & LineZ - 1 & ")"
            oRng = Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 4))
            oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 18)), Type:=xlFillDefault)
        End If
        oReader2.Close()
        ' 劃線
        oRng = Ws.Range("A3", Ws.Cells(LineZ, 18))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        'End If                


        ' 第四頁
        oCommand2.CommandText = "SELECT nvl(sum(aao05 + aao06),0) FROM aao_file where aao03 = " & tYear & " and aao04 = " & tMonth & " and aao01 like '6602%' and aao02 = 'D0210'"
        Dim PY2 As Decimal = oCommand2.ExecuteScalar()
        'If PY2 > 0 Then      
        TP += 1
        If TP > 3 Then
            Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Else
            Ws = xWorkBook.Sheets(TP)
        End If
        Ws.Activate()
        AdjustExcelFormat1("ADM Exp")

        oCommand2.CommandText = "select aag01,aag02 from aag_file where aag01 like '6602%' and aag07 = 2 order by aag01"
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                Ws.Cells(LineZ, 1) = oReader2.Item("aag01")
                Ws.Cells(LineZ, 2) = oReader2.Item("aag02")
                Ws.Cells(LineZ, 3) = GetDepartNmae(DNP)
                Ws.Cells(LineZ, 4) = Decimal.Round(GetLastYearSameMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 5) = Decimal.Round(GetLastMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 6) = Decimal.Round(GetThisYearSameMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3) - Decimal.Round(GetThisYearSameMonth_D0210(oReader2.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 7) = GetThisYearSameMonthBudget(oReader2.Item("aag01").ToString(), DNP)
                Ws.Cells(LineZ, 8) = "=F" & LineZ & "-G" & LineZ
                Ws.Cells(LineZ, 9) = "=F" & LineZ & "-D" & LineZ
                Ws.Cells(LineZ, 10) = "=F" & LineZ & "-E" & LineZ
                Ws.Cells(LineZ, 11) = Decimal.Round(GetLastYearBeforeMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 12) = Decimal.Round(GetThisYearBeforeMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 13) = GetThisYearBeforeMonthBudget(oReader2.Item("aag01").ToString(), DNP)
                Ws.Cells(LineZ, 14) = "=L" & LineZ & "-M" & LineZ
                Ws.Cells(LineZ, 15) = "=L" & LineZ & "-K" & LineZ
                Ws.Cells(LineZ, 16) = Decimal.Round(GetLastYearNoMonth(oReader2.Item("aag01").ToString(), DNP) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 17) = "=R" & LineZ & "-M" & LineZ & "+L" & LineZ
                Ws.Cells(LineZ, 18) = GetThisYearBudget(oReader2.Item("aag01").ToString(), DNP)
                LineZ += 1

            End While
            Ws.Cells(LineZ, 2) = "Total ADM Exp"
            Ws.Cells(LineZ, 4) = "=SUM(D7:D" & LineZ - 1 & ")"
            oRng = Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 4))
            oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ, 18)), Type:=xlFillDefault)
        End If
        oReader2.Close()


        ' 劃線
        oRng = Ws.Range("A3", Ws.Cells(LineZ, 18))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        'End If           

        ' 200401 add by Brady
        ' 第五頁 
        TP += 1
        If TP > 3 Then
            Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Else
            Ws = xWorkBook.Sheets(TP)
        End If
        Ws.Activate()
        AdjustExcelFormat3("KPI")
        Ws.Cells(3, 1) = "2020/01/01"
        Ws.Cells(4, 1) = "2020/02/01"
        Ws.Cells(5, 1) = "2020/03/01"
        Ws.Cells(6, 1) = "2020/04/01"
        Ws.Cells(7, 1) = "2020/05/01"
        Ws.Cells(8, 1) = "2020/06/01"
        Ws.Cells(9, 1) = "2020/07/01"
        Ws.Cells(10, 1) = "2020/08/01"
        Ws.Cells(11, 1) = "2020/09/01"
        Ws.Cells(12, 1) = "2020/10/01"
        Ws.Cells(13, 1) = "2020/11/01"
        Ws.Cells(14, 1) = "2020/12/01"
        For L As Int16 = 1 To tMonth Step 1
            Select Case L
                Case 1
                    oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 1 AND aag01 in ('600101') "
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        l_aah04_05 = oReader2.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(3, 2) = l_aah04_05
                    End If
                    oReader2.Close()

                    Ws.Cells(3, 3) = 16648075.15
                    Ws.Cells(3, 4) = "=B3/C3"

                    Dim l_abb07 As Decimal = 0
                    Dim t_abb07 As Decimal = 0
                    Dim t_tc_exu06 As Decimal = 0           '200422 add by Brady
                    oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/01/01'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/01/31'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                    oCommand2.CommandText += DNP & "' order by abb03"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        While oReader2.Read()
                            oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/01/01'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/01/31'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                            oReader3 = oCommand3.ExecuteReader()
                            If oReader3.HasRows() Then
                                While oReader3.Read()
                                    If oReader3.Item("abb06") = 1 Then
                                        ' 借方
                                        l_abb07 = oReader3.Item("abb07")
                                        t_abb07 = t_abb07 + l_abb07
                                    Else
                                        '貸方
                                        l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                        t_abb07 = t_abb07 + l_abb07
                                    End If
                                End While
                            End If
                            oReader3.Close()
                        End While
                        '200422 add by Brady
                        'Ws.Cells(3, 5) = t_abb07
                        oCommand4.CommandText = "select nvl(sum(tc_exu06),0) as s_tc_exu06 from tc_exu_file where year(tc_exu01) = 2020 and month(tc_exu01) = 1 and tc_exu02 = 'D0210'  "
                        oReader4 = oCommand4.ExecuteReader()
                        If oReader4.HasRows() Then
                            oReader4.Read()
                            t_tc_exu06 = oReader4.Item("s_tc_exu06")
                        End If
                        oReader4.Close()
                        Ws.Cells(3, 5) = t_abb07 - t_tc_exu06
                        '200422 add by Brady END
                    End If
                    oReader2.Close()

                    oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                    oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 1 AND tc_bud08 = '" & DNP & "'"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        Ws.Cells(3, 6) = oReader2.Item("s_tc_bud13")
                    End If
                    oReader2.Close()

                    Ws.Cells(3, 7) = "=E3-F3"
                    Ws.Cells(3, 8) = "=(E3-F3)/F3"
                    Ws.Cells(3, 9) = "=D3*F3"
                    Ws.Cells(3, 10) = "=E3-I3"
                    Ws.Cells(3, 11) = "=(E3-I3)/I3"
                Case 2
                    oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 2 AND aag01 in ('600101') "
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        l_aah04_05 = oReader2.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(4, 2) = l_aah04_05
                    End If
                    oReader2.Close()

                    Ws.Cells(4, 3) = 18577805.4844597
                    Ws.Cells(4, 4) = "=B4/C4"

                    Dim l_abb07 As Decimal = 0
                    Dim t_abb07 As Decimal = 0
                    Dim t_tc_exu06 As Decimal = 0           '200422 add by Brady
                    oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/02/01'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/02/29'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                    oCommand2.CommandText += DNP & "' order by abb03"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        While oReader2.Read()
                            oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/02/01'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/02/29'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                            oReader3 = oCommand3.ExecuteReader()
                            If oReader3.HasRows() Then
                                While oReader3.Read()
                                    If oReader3.Item("abb06") = 1 Then
                                        ' 借方
                                        l_abb07 = oReader3.Item("abb07")
                                        t_abb07 = t_abb07 + l_abb07
                                    Else
                                        '貸方
                                        l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                        t_abb07 = t_abb07 + l_abb07
                                    End If
                                End While
                            End If
                            oReader3.Close()
                        End While
                        '200422 add by Brady
                        'Ws.Cells(4, 5) = t_abb07
                        oCommand4.CommandText = "select nvl(sum(tc_exu06),0) as s_tc_exu06 from tc_exu_file where year(tc_exu01) = 2020 and month(tc_exu01) = 2 and tc_exu02 = 'D0210'  "
                        oReader4 = oCommand4.ExecuteReader()
                        If oReader4.HasRows() Then
                            oReader4.Read()
                            t_tc_exu06 = oReader4.Item("s_tc_exu06")
                        End If
                        oReader4.Close()
                        Ws.Cells(4, 5) = t_abb07 - t_tc_exu06
                        '200422 add by Brady END
                    End If
                    oReader2.Close()

                    oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                    oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 2 AND tc_bud08 = '" & DNP & "'"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        Ws.Cells(4, 6) = oReader2.Item("s_tc_bud13")
                    End If
                    oReader2.Close()

                    Ws.Cells(4, 7) = "=E4-F4"
                    Ws.Cells(4, 8) = "=(E4-F4)/F4"
                    Ws.Cells(4, 9) = "=D4*F4"
                    Ws.Cells(4, 10) = "=E4-I4"
                    Ws.Cells(4, 11) = "=(E4-I4)/I4"
                Case 3
                    oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 3 AND aag01 in ('600101') "
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        l_aah04_05 = oReader2.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(5, 2) = l_aah04_05
                    End If
                    oReader2.Close()

                    Ws.Cells(5, 3) = 16799152.1945558
                    Ws.Cells(5, 4) = "=B5/C5"

                    Dim l_abb07 As Decimal = 0
                    Dim t_abb07 As Decimal = 0
                    Dim t_tc_exu06 As Decimal = 0           '200422 add by Brady
                    oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/03/01'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/03/31'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                    oCommand2.CommandText += DNP & "' order by abb03"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        While oReader2.Read()
                            oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/03/01'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/03/31'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                            oReader3 = oCommand3.ExecuteReader()
                            If oReader3.HasRows() Then
                                While oReader3.Read()
                                    If oReader3.Item("abb06") = 1 Then
                                        ' 借方
                                        l_abb07 = oReader3.Item("abb07")
                                        t_abb07 = t_abb07 + l_abb07
                                    Else
                                        '貸方
                                        l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                        t_abb07 = t_abb07 + l_abb07
                                    End If
                                End While
                            End If
                            oReader3.Close()
                        End While
                        '200422 add by Brady
                        'Ws.Cells(5, 5) = t_abb07
                        oCommand4.CommandText = "select nvl(sum(tc_exu06),0) as s_tc_exu06 from tc_exu_file where year(tc_exu01) = 2020 and month(tc_exu01) = 3 and tc_exu02 = 'D0210'  "
                        oReader4 = oCommand4.ExecuteReader()
                        If oReader4.HasRows() Then
                            oReader4.Read()
                            t_tc_exu06 = oReader4.Item("s_tc_exu06")
                        End If
                        oReader4.Close()
                        Ws.Cells(5, 5) = t_abb07 - t_tc_exu06
                        '200422 add by Brady END
                    End If
                    oReader2.Close()

                    oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                    oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 3 AND tc_bud08 = '" & DNP & "'"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        Ws.Cells(5, 6) = oReader2.Item("s_tc_bud13")
                    End If
                    oReader2.Close()

                    Ws.Cells(5, 7) = "=E5-F5"
                    Ws.Cells(5, 8) = "=(E5-F5)/F5"
                    Ws.Cells(5, 9) = "=D5*F5"
                    Ws.Cells(5, 10) = "=E5-I5"
                    Ws.Cells(5, 11) = "=(E5-I5)/I5"
                Case 4
                    oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 4 AND aag01 in ('600101') "
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        l_aah04_05 = oReader2.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(6, 2) = l_aah04_05
                    End If
                    oReader2.Close()

                    Ws.Cells(6, 3) = 16262202.8683268
                    Ws.Cells(6, 4) = "=B6/C6"

                    Dim l_abb07 As Decimal = 0
                    Dim t_abb07 As Decimal = 0
                    Dim t_tc_exu06 As Decimal = 0           '200422 add by Brady
                    oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/04/01'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/04/30'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                    oCommand2.CommandText += DNP & "' order by abb03"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        While oReader2.Read()
                            oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/04/01'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/04/30'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                            oReader3 = oCommand3.ExecuteReader()
                            If oReader3.HasRows() Then
                                While oReader3.Read()
                                    If oReader3.Item("abb06") = 1 Then
                                        ' 借方
                                        l_abb07 = oReader3.Item("abb07")
                                        t_abb07 = t_abb07 + l_abb07
                                    Else
                                        '貸方
                                        l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                        t_abb07 = t_abb07 + l_abb07
                                    End If
                                End While
                            End If
                            oReader3.Close()
                        End While
                        '200422 add by Brady
                        'Ws.Cells(6, 5) = t_abb07
                        oCommand4.CommandText = "select nvl(sum(tc_exu06),0) as s_tc_exu06 from tc_exu_file where year(tc_exu01) = 2020 and month(tc_exu01) = 4 and tc_exu02 = 'D0210'  "
                        oReader4 = oCommand4.ExecuteReader()
                        If oReader4.HasRows() Then
                            oReader4.Read()
                            t_tc_exu06 = oReader4.Item("s_tc_exu06")
                        End If
                        oReader4.Close()
                        Ws.Cells(6, 5) = t_abb07 - t_tc_exu06
                        '200422 add by Brady END
                    End If
                    oReader2.Close()

                    oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                    oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 4 AND tc_bud08 = '" & DNP & "'"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        Ws.Cells(6, 6) = oReader2.Item("s_tc_bud13")
                    End If
                    oReader2.Close()

                    Ws.Cells(6, 7) = "=E6-F6"
                    Ws.Cells(6, 8) = "=(E6-F6)/F6"
                    Ws.Cells(6, 9) = "=D6*F6"
                    Ws.Cells(6, 10) = "=E6-I6"
                    Ws.Cells(6, 11) = "=(E6-I6)/I6"
                Case 5
                    oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 5 AND aag01 in ('600101') "
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        l_aah04_05 = oReader2.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(7, 2) = l_aah04_05
                    End If
                    oReader2.Close()

                    Ws.Cells(7, 3) = 18198819.4663555
                    Ws.Cells(7, 4) = "=B7/C7"

                    Dim l_abb07 As Decimal = 0
                    Dim t_abb07 As Decimal = 0
                    Dim t_tc_exu06 As Decimal = 0           '200422 add by Brady
                    oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/05/01'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/05/31'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                    oCommand2.CommandText += DNP & "' order by abb03"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        While oReader2.Read()
                            oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/05/01'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/05/31'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                            oReader3 = oCommand3.ExecuteReader()
                            If oReader3.HasRows() Then
                                While oReader3.Read()
                                    If oReader3.Item("abb06") = 1 Then
                                        ' 借方
                                        l_abb07 = oReader3.Item("abb07")
                                        t_abb07 = t_abb07 + l_abb07
                                    Else
                                        '貸方
                                        l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                        t_abb07 = t_abb07 + l_abb07
                                    End If
                                End While
                            End If
                            oReader3.Close()
                        End While
                        '200422 add by Brady
                        'Ws.Cells(7, 5) = t_abb07
                        oCommand4.CommandText = "select nvl(sum(tc_exu06),0) as s_tc_exu06 from tc_exu_file where year(tc_exu01) = 2020 and month(tc_exu01) = 5 and tc_exu02 = 'D0210'  "
                        oReader4 = oCommand4.ExecuteReader()
                        If oReader4.HasRows() Then
                            oReader4.Read()
                            t_tc_exu06 = oReader4.Item("s_tc_exu06")
                        End If
                        oReader4.Close()
                        Ws.Cells(7, 5) = t_abb07 - t_tc_exu06
                        '200422 add by Brady END
                    End If
                    oReader2.Close()

                    oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                    oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 5 AND tc_bud08 = '" & DNP & "'"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        Ws.Cells(7, 6) = oReader2.Item("s_tc_bud13")
                    End If
                    oReader2.Close()

                    Ws.Cells(7, 7) = "=E7-F7"
                    Ws.Cells(7, 8) = "=(E7-F7)/F7"
                    Ws.Cells(7, 9) = "=D7*F7"
                    Ws.Cells(7, 10) = "=E7-I7"
                    Ws.Cells(7, 11) = "=(E7-I7)/I7"
                Case 6
                    oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 6 AND aag01 in ('600101') "
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        l_aah04_05 = oReader2.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(8, 2) = l_aah04_05
                    End If
                    oReader2.Close()

                    Ws.Cells(8, 3) = 11654412.3220121
                    Ws.Cells(8, 4) = "=B8/C8"

                    Dim l_abb07 As Decimal = 0
                    Dim t_abb07 As Decimal = 0
                    Dim t_tc_exu06 As Decimal = 0           '200422 add by Brady
                    oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/06/01'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/06/30'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                    oCommand2.CommandText += DNP & "' order by abb03"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        While oReader2.Read()
                            oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/06/01'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/06/30'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                            oReader3 = oCommand3.ExecuteReader()
                            If oReader3.HasRows() Then
                                While oReader3.Read()
                                    If oReader3.Item("abb06") = 1 Then
                                        ' 借方
                                        l_abb07 = oReader3.Item("abb07")
                                        t_abb07 = t_abb07 + l_abb07
                                    Else
                                        '貸方
                                        l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                        t_abb07 = t_abb07 + l_abb07
                                    End If
                                End While
                            End If
                            oReader3.Close()
                        End While
                        '200422 add by Brady
                        'Ws.Cells(8, 5) = t_abb07
                        oCommand4.CommandText = "select nvl(sum(tc_exu06),0) as s_tc_exu06 from tc_exu_file where year(tc_exu01) = 2020 and month(tc_exu01) = 6 and tc_exu02 = 'D0210'  "
                        oReader4 = oCommand4.ExecuteReader()
                        If oReader4.HasRows() Then
                            oReader4.Read()
                            t_tc_exu06 = oReader4.Item("s_tc_exu06")
                        End If
                        oReader4.Close()
                        Ws.Cells(8, 5) = t_abb07 - t_tc_exu06
                        '200422 add by Brady END
                    End If
                    oReader2.Close()

                    oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                    oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 6 AND tc_bud08 = '" & DNP & "'"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        Ws.Cells(8, 6) = oReader2.Item("s_tc_bud13")
                    End If
                    oReader2.Close()

                    Ws.Cells(8, 7) = "=E8-F8"
                    Ws.Cells(8, 8) = "=(E8-F8)/F8"
                    Ws.Cells(8, 9) = "=D8*F8"
                    Ws.Cells(8, 10) = "=E8-I8"
                    Ws.Cells(8, 11) = "=(E8-I8)/I8"
                Case 7
                    oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 7 AND aag01 in ('600101') "
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        l_aah04_05 = oReader2.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(9, 2) = l_aah04_05
                    End If
                    oReader2.Close()

                    Ws.Cells(9, 3) = 12880421.7803447
                    Ws.Cells(9, 4) = "=B9/C9"

                    Dim l_abb07 As Decimal = 0
                    Dim t_abb07 As Decimal = 0
                    Dim t_tc_exu06 As Decimal = 0           '200422 add by Brady
                    oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/07/01'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/07/31'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                    oCommand2.CommandText += DNP & "' order by abb03"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        While oReader2.Read()
                            oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/07/01'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/07/31'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                            oReader3 = oCommand3.ExecuteReader()
                            If oReader3.HasRows() Then
                                While oReader3.Read()
                                    If oReader3.Item("abb06") = 1 Then
                                        ' 借方
                                        l_abb07 = oReader3.Item("abb07")
                                        t_abb07 = t_abb07 + l_abb07
                                    Else
                                        '貸方
                                        l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                        t_abb07 = t_abb07 + l_abb07
                                    End If
                                End While
                            End If
                            oReader3.Close()
                        End While
                        '200422 add by Brady
                        'Ws.Cells(9, 5) = t_abb07
                        oCommand4.CommandText = "select nvl(sum(tc_exu06),0) as s_tc_exu06 from tc_exu_file where year(tc_exu01) = 2020 and month(tc_exu01) = 7 and tc_exu02 = 'D0210'  "
                        oReader4 = oCommand4.ExecuteReader()
                        If oReader4.HasRows() Then
                            oReader4.Read()
                            t_tc_exu06 = oReader4.Item("s_tc_exu06")
                        End If
                        oReader4.Close()
                        Ws.Cells(9, 5) = t_abb07 - t_tc_exu06
                        '200422 add by Brady END
                    End If
                    oReader2.Close()

                    oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                    oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 7 AND tc_bud08 = '" & DNP & "'"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        Ws.Cells(9, 6) = oReader2.Item("s_tc_bud13")
                    End If
                    oReader2.Close()

                    Ws.Cells(9, 7) = "=E9-F9"
                    Ws.Cells(9, 8) = "=(E9-F9)/F9"
                    Ws.Cells(9, 9) = "=D9*F9"
                    Ws.Cells(9, 10) = "=E9-I9"
                    Ws.Cells(9, 11) = "=(E9-I9)/I9"
                Case 8
                    oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 8 AND aag01 in ('600101') "
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        l_aah04_05 = oReader2.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(10, 2) = l_aah04_05
                    End If
                    oReader2.Close()

                    Ws.Cells(10, 3) = 18073323.7356463
                    Ws.Cells(10, 4) = "=B10/C10"

                    Dim l_abb07 As Decimal = 0
                    Dim t_abb07 As Decimal = 0
                    Dim t_tc_exu06 As Decimal = 0           '200422 add by Brady
                    oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/08/01'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/08/31'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                    oCommand2.CommandText += DNP & "' order by abb03"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        While oReader2.Read()
                            oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/08/01'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/08/31'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                            oReader3 = oCommand3.ExecuteReader()
                            If oReader3.HasRows() Then
                                While oReader3.Read()
                                    If oReader3.Item("abb06") = 1 Then
                                        ' 借方
                                        l_abb07 = oReader3.Item("abb07")
                                        t_abb07 = t_abb07 + l_abb07
                                    Else
                                        '貸方
                                        l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                        t_abb07 = t_abb07 + l_abb07
                                    End If
                                End While
                            End If
                            oReader3.Close()
                        End While
                        '200422 add by Brady
                        'Ws.Cells(10, 5) = t_abb07
                        oCommand4.CommandText = "select nvl(sum(tc_exu06),0) as s_tc_exu06 from tc_exu_file where year(tc_exu01) = 2020 and month(tc_exu01) = 8 and tc_exu02 = 'D0210'  "
                        oReader4 = oCommand4.ExecuteReader()
                        If oReader4.HasRows() Then
                            oReader4.Read()
                            t_tc_exu06 = oReader4.Item("s_tc_exu06")
                        End If
                        oReader4.Close()
                        Ws.Cells(10, 5) = t_abb07 - t_tc_exu06
                        '200422 add by Brady END
                    End If
                    oReader2.Close()

                    oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                    oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 8 AND tc_bud08 = '" & DNP & "'"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        Ws.Cells(10, 6) = oReader2.Item("s_tc_bud13")
                    End If
                    oReader2.Close()

                    Ws.Cells(10, 7) = "=E10-F10"
                    Ws.Cells(10, 8) = "=(E10-F10)/F10"
                    Ws.Cells(10, 9) = "=D10*F10"
                    Ws.Cells(10, 10) = "=E10-I10"
                    Ws.Cells(10, 11) = "=(E10-I10)/I10"
                Case 9
                    oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 9 AND aag01 in ('600101') "
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        l_aah04_05 = oReader2.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(11, 2) = l_aah04_05
                    End If
                    oReader2.Close()

                    Ws.Cells(11, 3) = 13379917.3701231
                    Ws.Cells(11, 4) = "=B11/C11"

                    Dim l_abb07 As Decimal = 0
                    Dim t_abb07 As Decimal = 0
                    Dim t_tc_exu06 As Decimal = 0           '200422 add by Brady
                    oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/09/01'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/09/30'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                    oCommand2.CommandText += DNP & "' order by abb03"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        While oReader2.Read()
                            oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/09/01'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/09/30'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                            oReader3 = oCommand3.ExecuteReader()
                            If oReader3.HasRows() Then
                                While oReader3.Read()
                                    If oReader3.Item("abb06") = 1 Then
                                        ' 借方
                                        l_abb07 = oReader3.Item("abb07")
                                        t_abb07 = t_abb07 + l_abb07
                                    Else
                                        '貸方
                                        l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                        t_abb07 = t_abb07 + l_abb07
                                    End If
                                End While
                            End If
                            oReader3.Close()
                        End While
                        '200422 add by Brady
                        'Ws.Cells(11, 5) = t_abb07
                        oCommand4.CommandText = "select nvl(sum(tc_exu06),0) as s_tc_exu06 from tc_exu_file where year(tc_exu01) = 2020 and month(tc_exu01) = 9 and tc_exu02 = 'D0210'  "
                        oReader4 = oCommand4.ExecuteReader()
                        If oReader4.HasRows() Then
                            oReader4.Read()
                            t_tc_exu06 = oReader4.Item("s_tc_exu06")
                        End If
                        oReader4.Close()
                        Ws.Cells(11, 5) = t_abb07 - t_tc_exu06
                        '200422 add by Brady END
                    End If
                    oReader2.Close()

                    oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                    oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 9 AND tc_bud08 = '" & DNP & "'"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        Ws.Cells(11, 6) = oReader2.Item("s_tc_bud13")
                    End If
                    oReader2.Close()

                    Ws.Cells(11, 7) = "=E11-F11"
                    Ws.Cells(11, 8) = "=(E11-F11)/F11"
                    Ws.Cells(11, 9) = "=D11*F11"
                    Ws.Cells(11, 10) = "=E11-I11"
                    Ws.Cells(11, 11) = "=(E11-I11)/I11"
                Case 10
                    oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 10 AND aag01 in ('600101') "
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        l_aah04_05 = oReader2.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(12, 2) = l_aah04_05
                    End If
                    oReader2.Close()

                    Ws.Cells(12, 3) = 14633836.7242984
                    Ws.Cells(12, 4) = "=B12/C12"

                    Dim l_abb07 As Decimal = 0
                    Dim t_abb07 As Decimal = 0
                    Dim t_tc_exu06 As Decimal = 0           '200422 add by Brady
                    oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/10/01'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/10/31'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                    oCommand2.CommandText += DNP & "' order by abb03"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        While oReader2.Read()
                            oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/10/01'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/10/31'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                            oReader3 = oCommand3.ExecuteReader()
                            If oReader3.HasRows() Then
                                While oReader3.Read()
                                    If oReader3.Item("abb06") = 1 Then
                                        ' 借方
                                        l_abb07 = oReader3.Item("abb07")
                                        t_abb07 = t_abb07 + l_abb07
                                    Else
                                        '貸方
                                        l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                        t_abb07 = t_abb07 + l_abb07
                                    End If
                                End While
                            End If
                            oReader3.Close()
                        End While
                        '200422 add by Brady
                        'Ws.Cells(12, 5) = t_abb07
                        oCommand4.CommandText = "select nvl(sum(tc_exu06),0) as s_tc_exu06 from tc_exu_file where year(tc_exu01) = 2020 and month(tc_exu01) = 10 and tc_exu02 = 'D0210'  "
                        oReader4 = oCommand4.ExecuteReader()
                        If oReader4.HasRows() Then
                            oReader4.Read()
                            t_tc_exu06 = oReader4.Item("s_tc_exu06")
                        End If
                        oReader4.Close()
                        Ws.Cells(12, 5) = t_abb07 - t_tc_exu06
                        '200422 add by Brady END
                    End If
                    oReader2.Close()

                    oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                    oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 10 AND tc_bud08 = '" & DNP & "'"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        Ws.Cells(12, 6) = oReader2.Item("s_tc_bud13")
                    End If
                    oReader2.Close()

                    Ws.Cells(12, 7) = "=E12-F12"
                    Ws.Cells(12, 8) = "=(E12-F12)/F12"
                    Ws.Cells(12, 9) = "=D12*F12"
                    Ws.Cells(12, 10) = "=E12-I12"
                    Ws.Cells(12, 11) = "=(E12-I12)/I12"
                Case 11
                    oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 11 AND aag01 in ('600101') "
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        l_aah04_05 = oReader2.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(13, 2) = l_aah04_05
                    End If
                    oReader2.Close()

                    Ws.Cells(13, 3) = 9739716.54373806
                    Ws.Cells(13, 4) = "=B13/C13"

                    Dim l_abb07 As Decimal = 0
                    Dim t_abb07 As Decimal = 0
                    Dim t_tc_exu06 As Decimal = 0           '200422 add by Brady
                    oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/11/01'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/11/30'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                    oCommand2.CommandText += DNP & "' order by abb03"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        While oReader2.Read()
                            oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/11/01'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/11/30'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                            oReader3 = oCommand3.ExecuteReader()
                            If oReader3.HasRows() Then
                                While oReader3.Read()
                                    If oReader3.Item("abb06") = 1 Then
                                        ' 借方
                                        l_abb07 = oReader3.Item("abb07")
                                        t_abb07 = t_abb07 + l_abb07
                                    Else
                                        '貸方
                                        l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                        t_abb07 = t_abb07 + l_abb07
                                    End If
                                End While
                            End If
                            oReader3.Close()
                        End While
                        '200422 add by Brady
                        'Ws.Cells(13, 5) = t_abb07
                        oCommand4.CommandText = "select nvl(sum(tc_exu06),0) as s_tc_exu06 from tc_exu_file where year(tc_exu01) = 2020 and month(tc_exu01) = 11 and tc_exu02 = 'D0210'  "
                        oReader4 = oCommand4.ExecuteReader()
                        If oReader4.HasRows() Then
                            oReader4.Read()
                            t_tc_exu06 = oReader4.Item("s_tc_exu06")
                        End If
                        oReader4.Close()
                        Ws.Cells(13, 5) = t_abb07 - t_tc_exu06
                        '200422 add by Brady END
                    End If
                    oReader2.Close()

                    oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                    oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 11 AND tc_bud08 = '" & DNP & "'"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        Ws.Cells(13, 6) = oReader2.Item("s_tc_bud13")
                    End If
                    oReader2.Close()

                    Ws.Cells(13, 7) = "=E13-F13"
                    Ws.Cells(13, 8) = "=(E13-F13)/F13"
                    Ws.Cells(13, 9) = "=D13*F13"
                    Ws.Cells(13, 10) = "=E13-I13"
                    Ws.Cells(13, 11) = "=(E13-I13)/I13"
                Case 12
                    oCommand2.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand2.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 12 AND aag01 in ('600101') "
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        l_aah04_05 = oReader2.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(14, 2) = l_aah04_05
                    End If
                    oReader2.Close()

                    Ws.Cells(14, 3) = 11758282.582403
                    Ws.Cells(14, 4) = "=B14/C14"

                    Dim l_abb07 As Decimal = 0
                    Dim t_abb07 As Decimal = 0
                    Dim t_tc_exu06 As Decimal = 0           '200422 add by Brady
                    oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('2020/12/01'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and to_date('2020/12/31'"
                    oCommand2.CommandText += ",'yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
                    oCommand2.CommandText += DNP & "' order by abb03"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        While oReader2.Read()
                            oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('2020/12/01'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and to_date('2020/12/31'"
                            oCommand3.CommandText += ",'yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                            oReader3 = oCommand3.ExecuteReader()
                            If oReader3.HasRows() Then
                                While oReader3.Read()
                                    If oReader3.Item("abb06") = 1 Then
                                        ' 借方
                                        l_abb07 = oReader3.Item("abb07")
                                        t_abb07 = t_abb07 + l_abb07
                                    Else
                                        '貸方
                                        l_abb07 = oReader3.Item("abb07") * Decimal.MinusOne
                                        t_abb07 = t_abb07 + l_abb07
                                    End If
                                End While
                            End If
                            oReader3.Close()
                        End While
                        '200422 add by Brady
                        'Ws.Cells(14, 5) = t_abb07
                        oCommand4.CommandText = "select nvl(sum(tc_exu06),0) as s_tc_exu06 from tc_exu_file where year(tc_exu01) = 2020 and month(tc_exu01) = 12 and tc_exu02 = 'D0210'  "
                        oReader4 = oCommand4.ExecuteReader()
                        If oReader4.HasRows() Then
                            oReader4.Read()
                            t_tc_exu06 = oReader4.Item("s_tc_exu06")
                        End If
                        oReader4.Close()
                        Ws.Cells(14, 5) = t_abb07 - t_tc_exu06
                        '200422 add by Brady END
                    End If
                    oReader2.Close()

                    oCommand2.CommandText = "  SELECT SUM(tc_bud13) as s_tc_bud13 FROM tc_bud_file "
                    oCommand2.CommandText += "  WHERE tc_bud01 = '2' AND tc_bud02 = " & tYear & " AND tc_bud03 = 12 AND tc_bud08 = '" & DNP & "'"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        Ws.Cells(14, 6) = oReader2.Item("s_tc_bud13")
                    End If
                    oReader2.Close()

                    Ws.Cells(14, 7) = "=E14-F14"
                    Ws.Cells(14, 8) = "=(E14-F14)/F14"
                    Ws.Cells(14, 9) = "=D14*F14"
                    Ws.Cells(14, 10) = "=E14-I14"
                    Ws.Cells(14, 11) = "=(E14-I14)/I14"
            End Select
        Next
        Ws.Cells(15, 2) = "=SUM(B3:B14)"
        Ws.Cells(15, 3) = "=SUM(C3:C14)"
        Ws.Cells(15, 4) = "=B15/C15"
        Ws.Cells(15, 5) = "=SUM(E3:E14)"
        Ws.Cells(15, 6) = "=SUM(F3:F14)"
        Ws.Cells(15, 7) = "=SUM(G3:G14)"
        Ws.Cells(15, 8) = "=G15/F15"
        Ws.Cells(15, 9) = "=SUM(I3:I14)"
        Ws.Cells(15, 10) = "=SUM(J3:J14)"
        Ws.Cells(15, 11) = "=J15/I15"
        ' 200401 add by Brady END

        ' 200401 Note By Brady 變成第六頁 
        ' 第五頁  -->改為按月區分 20191011
        For z As Int16 = Ldate1.Month To Ldate2.Month Step 1
            TP += 1
            If TP > 3 Then
                Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
            Else
                Ws = xWorkBook.Sheets(TP)
            End If
            Ws.Activate()
            AdjustExcelFormat2(z)
            oCommand2.CommandText = "select distinct abb03 from abb_file,aba_file where abb01 = aba01 and abapost = 'Y' and aba02 between to_date('"
            oCommand2.CommandText += Ldate1.AddMonths(z - 1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            oCommand2.CommandText += Ldate1.AddMonths(z).AddDays(-1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and (abb03 like '6601%' or abb03 like '6604%' or abb03 like '6602%' or abb03 like '5101%') and abb05 = '"
            oCommand2.CommandText += DNP & "' order by abb03"
            oReader2 = oCommand2.ExecuteReader()
            If oReader2.HasRows() Then
                While oReader2.Read()
                    oCommand3.CommandText = "select aba02,abb01,abb03,aag02,abb04,abb06,abb07 from abb_file,aba_file,aag_file where abb01 = aba01 and abb03 = aag01 and abapost = 'Y' and aba02 between to_date('"
                    oCommand3.CommandText += Ldate1.AddMonths(z - 1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                    oCommand3.CommandText += Ldate1.AddMonths(z).AddDays(-1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and abb03  = '" & oReader2.Item("abb03") & "' and abb05 = '" & DNP & "' order by aba02"
                    oReader3 = oCommand3.ExecuteReader()
                    If oReader3.HasRows() Then
                        While oReader3.Read()
                            Ws.Cells(LineZ, 1) = DNP
                            Ws.Cells(LineZ, 2) = "D0210"
                            Ws.Cells(LineZ, 3) = oReader3.Item("aba02")
                            Ws.Cells(LineZ, 4) = oReader3.Item("abb01")
                            Ws.Cells(LineZ, 5) = oReader2.Item("abb03")
                            Ws.Cells(LineZ, 6) = oReader3.Item("aag02")
                            Ws.Cells(LineZ, 7) = oReader3.Item("abb04")
                            If oReader3.Item("abb06") = 1 Then
                                ' 借方
                                Ws.Cells(LineZ, 8) = oReader3.Item("abb07")
                            Else
                                '貸方
                                Ws.Cells(LineZ, 8) = oReader3.Item("abb07") * Decimal.MinusOne
                            End If
                            LineZ += 1
                        End While
                    End If
                    oReader3.Close()
                End While
            End If
            oReader2.Close()
            oRng = Ws.Range("A1", "G1")
            oRng.EntireColumn.AutoFit()
        Next

        ' 第六頁 cgli602
        TP += 1
        If TP > 3 Then
            Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Else
            Ws = xWorkBook.Sheets(TP)
        End If
        Ws.Activate()
        AdjustExcelFormat4("Debit Note")
        oCommand2.CommandText = "select tc_exu01,tc_exu02,tc_exu03,aag02,tc_exu04,tc_exu05,tc_exu06,tc_exu07 from tc_exu_file,aag_file where tc_exu03 = aag01 "
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                Ws.Cells(LineZ, 1) = oReader2.Item("tc_exu01")
                Ws.Cells(LineZ, 2) = oReader2.Item("tc_exu02")
                Ws.Cells(LineZ, 3) = oReader2.Item("tc_exu03")
                Ws.Cells(LineZ, 4) = oReader2.Item("aag02")
                Ws.Cells(LineZ, 5) = oReader2.Item("tc_exu04")
                Ws.Cells(LineZ, 6) = oReader2.Item("tc_exu05")
                Ws.Cells(LineZ, 7) = oReader2.Item("tc_exu06")
                Ws.Cells(LineZ, 8) = oReader2.Item("tc_exu07")
                LineZ += 1
            End While
        End If
        oReader2.Close()

        SaveExcel_2("CS 客服", "D0210")
        MsgBox("Finished")
    End Sub
    '200417 add by Brady END
    Private Sub AdjustExcelFormat1(ByVal hh As String)
        '190404 add by Brady
        'xExcel.ActiveWindow.Zoom = 75
        'Ws.Name = hh
        'Ws.Columns.EntireColumn.ColumnWidth = 10.44
        'oRng = Ws.Range("B1", "B1")
        'oRng.EntireColumn.ColumnWidth = 60
        'oRng = Ws.Range("B3", "R3")
        'oRng.Merge()
        'oRng.HorizontalAlignment = xlCenter
        ''oRng.Interior.Color = Color.FromArgb(169, 209, 141)
        'Ws.Cells(3, 2) = hh & ". By account"
        'Ws.Cells(4, 2) = "RMB"
        'oRng = Ws.Range("B5", "B5")
        'oRng.NumberFormatLocal = "mmm-yy"
        'oRng.HorizontalAlignment = xlLeft
        'Ws.Cells(5, 2) = tDate
        'Ws.Cells(6, 2) = "Dongguan Action Composites LTD Co."
        'oRng = Ws.Range("C4", "C6")
        'oRng.Merge()
        'oRng.HorizontalAlignment = xlCenter
        'Ws.Cells(4, 3) = "Cost" & Chr(10) & "Center"
        'oRng = Ws.Range("D4", "F5")
        'oRng.Merge()
        'oRng.HorizontalAlignment = xlCenter
        'Ws.Cells(4, 4) = "Actual"
        'Ws.Cells(6, 4) = tDate.AddYears(-1)
        'Ws.Cells(6, 5) = tDate.AddMonths(-1)
        'Ws.Cells(6, 6) = tDate
        'Ws.Cells(6, 7) = tDate
        'oRng = Ws.Range("D6", "G6")
        'oRng.NumberFormatLocal = "mmm-yy"
        'oRng = Ws.Range("G4", "G5")
        'oRng.Merge()
        'oRng.HorizontalAlignment = xlCenter
        'Ws.Cells(4, 7) = "Budget"
        'oRng = Ws.Range("H4", "J4")
        'oRng.Merge()
        'oRng.HorizontalAlignment = xlCenter
        'Ws.Cells(4, 8) = "Variance" '& Chr(10) & "Act& Bud"
        'Ws.Cells(5, 8) = "Act & But"
        'Ws.Cells(5, 9) = "year-on-year"
        'Ws.Cells(5, 10) = "Month-on-month"
        'Ws.Cells(6, 8) = "RMB"
        'Ws.Cells(6, 9) = "RMB"
        'Ws.Cells(6, 10) = "RMB"
        ''oRng = Ws.Range("D4", "J6")
        ''oRng.Interior.Color = Color.FromArgb(255, 218, 101)
        'oRng = Ws.Range("K4", "L5")
        'oRng.Merge()
        'oRng.HorizontalAlignment = xlCenter
        'Ws.Cells(4, 11) = "Actual"
        'Ws.Cells(6, 11) = "YTD " & pYear
        'Ws.Cells(6, 12) = "YTD " & tYear
        'oRng = Ws.Range("M4", "M5")
        'oRng.Merge()
        'oRng.HorizontalAlignment = xlCenter
        'Ws.Cells(4, 13) = "Budget"
        'Ws.Cells(6, 13) = "YTD " & tYear
        'oRng = Ws.Range("N4", "O4")
        'oRng.Merge()
        'oRng.HorizontalAlignment = xlCenter
        'Ws.Cells(4, 14) = "Variance" '& Chr(10) & "Act& Bud"
        'Ws.Cells(5, 14) = "Act & But"
        'Ws.Cells(5, 15) = "year-on-year"
        'Ws.Cells(6, 14) = "RMB"
        'Ws.Cells(6, 15) = "RMB"
        ''oRng = Ws.Range("K4", "O6")
        ''oRng.Interior.Color = Color.FromArgb(156, 195, 230)
        'oRng = Ws.Range("P4", "P5")
        'oRng.Merge()
        'oRng.HorizontalAlignment = xlCenter
        'Ws.Cells(4, 16) = "Actual"
        'Ws.Cells(6, 16) = "Y" & pYear
        'oRng = Ws.Range("Q4", "Q5")
        'oRng.Merge()
        'oRng.HorizontalAlignment = xlCenter
        'Ws.Cells(4, 17) = "Rollling" & Chr(10) & "Forecast"
        'Ws.Cells(6, 17) = "Y" & tYear
        'oRng = Ws.Range("R4", "R5")
        'oRng.Merge()
        'oRng.HorizontalAlignment = xlCenter
        'Ws.Cells(4, 18) = "Budget"
        'Ws.Cells(6, 18) = tYear
        ''oRng = Ws.Range("M4", "O6")
        ''oRng.Interior.Color = Color.FromArgb(146, 208, 80)
        'oRng = Ws.Range("A1", "A1")
        'oRng.EntireColumn.NumberFormatLocal = "@"
        'oRng = Ws.Range("C1", "C1")
        'oRng.EntireColumn.NumberFormatLocal = "@"
        '' 劃線
        'oRng = Ws.Range("B3", "R6")
        'oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        'oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        'oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        ''oRng = Ws.Range("D6", "R6")
        ''oRng.HorizontalAlignment = xlRight
        'oRng = Ws.Range("C4", "R6")
        'oRng.HorizontalAlignment = xlCenter
        'oRng = Ws.Range("C1", "R1")
        'oRng.EntireColumn.ColumnWidth = 14
        'LineZ = 7
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = hh
        Ws.Columns.EntireColumn.ColumnWidth = 10.44
        Ws.Cells(1, 1) = "负责人："
        Ws.Cells(1, 2) = DP_BOSS
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 60
        oRng = Ws.Range("B3", "R3")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(3, 2) = hh & ". By account 制造费用-科目类别"
        Ws.Cells(4, 1) = "币别"
        Ws.Cells(4, 2) = "RMB"
        Ws.Cells(5, 1) = "期别"
        Ws.Cells(6, 1) = "科目代码"
        oRng = Ws.Range("B5", "B5")
        oRng.NumberFormatLocal = "mmm-yy"
        oRng.HorizontalAlignment = xlLeft
        Ws.Cells(5, 2) = tDate
        Ws.Cells(6, 2) = "Dongguan Action Composites LTD Co."
        oRng = Ws.Range("C4", "C6")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 3) = "Cost" & Chr(10) & "Center" & Chr(10) & "成本中心"
        oRng = Ws.Range("D4", "F5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 4) = "Actual 月实际发生金额"
        Ws.Cells(6, 4) = tDate.AddYears(-1) & Chr(10) & "去年同期"

        '190515 add by Brady
        'Ws.Cells(6, 5) = tDate.AddMonths(-1) & Chr(10) & "上期"
        Ws.Cells(6, 5) = tDate_1 & Chr(10) & "上期"
        '190515 add by Brady END

        Ws.Cells(6, 6) = tDate & Chr(10) & "当期"
        Ws.Cells(6, 7) = tDate & Chr(10) & "当期"
        oRng = Ws.Range("D6", "G6")
        oRng.NumberFormatLocal = "mmm-yy"
        oRng = Ws.Range("G4", "G5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 7) = "Budget 预算"
        oRng = Ws.Range("H4", "J4")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 8) = "Variance"
        Ws.Cells(5, 8) = "Act & But" & Chr(10) & "实际-预算"
        Ws.Cells(5, 9) = "year-on-year" & Chr(10) & "同期比较"
        Ws.Cells(5, 10) = "Month-on-month" & Chr(10) & "上期比较"
        Ws.Cells(6, 8) = "RMB"
        Ws.Cells(6, 9) = "RMB"
        Ws.Cells(6, 10) = "RMB"
        oRng = Ws.Range("K4", "L5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 11) = "Actual 累计发生金额"
        Ws.Cells(6, 11) = "YTD " & pYear & Chr(10) & "去年同期累计"
        Ws.Cells(6, 12) = "YTD " & tYear & Chr(10) & "本年同期累计"
        oRng = Ws.Range("M4", "M5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 13) = "Budget" & Chr(10) & "累计预算"
        Ws.Cells(6, 13) = "YTD " & tYear
        oRng = Ws.Range("N4", "O4")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 14) = "Variance 差异"
        Ws.Cells(5, 14) = "Act & But" & Chr(10) & "累计实际-累计预算"
        Ws.Cells(5, 15) = "year-on-year" & Chr(10) & "累计去年同期-累计本年同期"
        Ws.Cells(6, 14) = "RMB"
        Ws.Cells(6, 15) = "RMB"
        oRng = Ws.Range("P4", "P5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 16) = "Actual" & Chr(10) & "去年总计"
        Ws.Cells(6, 16) = "Y" & pYear
        oRng = Ws.Range("Q4", "Q5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 17) = "Rolling" & Chr(10) & "Forecast" & Chr(10) & "今年预测总计"
        Ws.Cells(6, 17) = "Y" & tYear
        oRng = Ws.Range("R4", "R5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 18) = "Budget" & Chr(10) & "今年预算总计"
        Ws.Cells(6, 18) = tYear
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
        oRng = Ws.Range("C4", "R6")
        oRng.HorizontalAlignment = xlCenter
        oRng = Ws.Range("C1", "R1")
        oRng.EntireColumn.ColumnWidth = 14
        LineZ = 7
        '190404 add by Brady END

    End Sub
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
        If tYear = 2019 Then
            oCommand2.CommandText = "select sum(t1) from ( "
            oCommand2.CommandText += "select nvl(sum(aao05-aao06),0) as t1 from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
            oCommand2.CommandText += tYear & " and aao04 between 1 and 6 "
            oCommand2.CommandText += "union all "
            oCommand2.CommandText += "select nvl(sum(tc_bud13),0) from tc_bud_file where tc_bud07 = '" & aag01 & "' and tc_bud08 = '" & gem01 & "' and tc_bud02 = "
            oCommand2.CommandText += tYear & " and tc_bud03 between 7 and " & tMonth
            oCommand2.CommandText += ") "

        Else
            oCommand2.CommandText = "select nvl(sum(tc_bud13),0) from tc_bud_file where tc_bud07 = '" & aag01 & "' and tc_bud08 = '" & gem01 & "' and tc_bud02 = "
            oCommand2.CommandText += tYear & " and tc_bud03 <= " & tMonth
        End If
        
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

        If tYear = 2019 Then
            oCommand2.CommandText = "select sum(t1) from ( "
            oCommand2.CommandText += "select nvl(sum(aao05-aao06),0) as t1 from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
            oCommand2.CommandText += tYear & " and aao04 between 1 and 6 "
            oCommand2.CommandText += "union all "
            oCommand2.CommandText += "select nvl(sum(tc_bud13),0) from tc_bud_file where tc_bud07 = '" & aag01 & "' and tc_bud08 = '" & gem01 & "' and tc_bud02 = "
            oCommand2.CommandText += tYear & " and tc_bud03 between 7 and 12 "
            oCommand2.CommandText += ") "
        Else
            oCommand2.CommandText = "select nvl(sum(tc_bud13),0) from tc_bud_file where tc_bud07 = '" & aag01 & "' and tc_bud08 = '" & gem01 & "' and tc_bud02 = "
            oCommand2.CommandText += tYear.ToString()
        End If

        oCommand2.CommandText = "select nvl(sum(tc_bud13),0) from tc_bud_file where tc_bud07 = '" & aag01 & "' and tc_bud08 = '" & gem01 & "' and tc_bud02 = "
        oCommand2.CommandText += tYear.ToString()
        Dim TYTMB As Decimal = Decimal.Round(oCommand2.ExecuteScalar() / ExchangeRate1, 3)
        Return TYTMB
    End Function

    Private Function GetLastYearSameMonthA(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
        oCommand2.CommandText += pYear & " and aao04 = " & tMonth
        Dim LYTM As Decimal = oCommand2.ExecuteScalar()
        Return LYTM
    End Function
    Private Function GetThisYearSameMonthA(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
        oCommand2.CommandText += tYear & " and aao04 = " & tMonth
        Dim TYTM As Decimal = oCommand2.ExecuteScalar()
        Return TYTM
    End Function
    Private Function GetLastMonthA(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 <> 'D9999' and aao02 = '" & gem01 & "' and aao03 = "
        oCommand2.CommandText += lYear & " and aao04 = " & lMonth
        Dim TYTM As Decimal = oCommand2.ExecuteScalar()
        Return TYTM
    End Function
    Private Function GetLastYearBeforeMonthA(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
        oCommand2.CommandText += pYear & " and aao04 <= " & tMonth & " and aao04 > 0"
        Dim LYBM As Decimal = oCommand2.ExecuteScalar()
        Return LYBM
    End Function
    Private Function GetThisYearBeforeMonthA(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
        oCommand2.CommandText += tYear & " and aao04 <= " & tMonth & " and aao04 > 0"
        Dim TYBM As Decimal = oCommand2.ExecuteScalar()
        Return TYBM
    End Function
    Private Function GetLastYearNoMonthA(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
        oCommand2.CommandText += pYear.ToString() & " and aao04 > 0"
        Dim TYNM As Decimal = oCommand2.ExecuteScalar()
        Return TYNM
    End Function
    '200422 add by Brady
    Private Function GetThisYearSameMonth_D0210(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(tc_exu06),0) from tc_exu_file where tc_exu03 = '" & aag01 & "' and tc_exu02 = 'D0210' and year(tc_exu01) = "
        oCommand2.CommandText += tYear & " and month(tc_exu01) = " & tMonth
        Dim TYTM_D0210 As Decimal = oCommand2.ExecuteScalar()
        Return TYTM_D0210
    End Function
    '200422 add by Brady END
    Private Sub SaveExcel(ByVal gem02 As String, ByVal gem01 As String)
        Dim SS As String = String.Empty
        If tMonth < 10 Then
            SS = "0" & tMonth
        Else
            SS = tMonth
        End If
        Dim SFN As String = "S:\A02_Finance_財務部\FN32-外挂报表\部门费用明细表\" & tYear & SS & "\" & gem02 & ".xlsx"
        'Dim SFN As String = "C:\TEMP\" & tYear & SS & "_" & gem02 & ".xlsx"
        Ws.SaveAs(SFN, XlFileFormat.xlOpenXMLWorkbook)
        xWorkBook.Saved = True
        xWorkBook.Close()
        xExcel.Quit()
        'If oConnection.State = ConnectionState.Open Then
        Try
            'oConnection.Close()
            Module1.KillExcelProcess(OldExcel)
            MailSend(SFN, gem01, gem02)
            'MsgBox("Finished")
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        'End If
    End Sub
    '200417 add by Brady
    Private Sub SaveExcel_2(ByVal gem02 As String, ByVal gem01 As String)
        Dim SS As String = String.Empty
        If tMonth < 10 Then
            SS = "0" & tMonth
        Else
            SS = tMonth
        End If
        Dim SFN As String = "S:\A02_Finance_財務部\FN32-外挂报表\部门费用明细表\" & tYear & SS & "\" & gem02 & "(扣减向客户收费部分).xlsx"
        'Dim SFN As String = "C:\TEMP\" & tYear & SS & "_" & gem02 & "(扣减向客户收费部分).xlsx"
        Ws.SaveAs(SFN, XlFileFormat.xlOpenXMLWorkbook)
        xWorkBook.Saved = True
        xWorkBook.Close()
        xExcel.Quit()
        'If oConnection.State = ConnectionState.Open Then
        Try
            'oConnection.Close()
            Module1.KillExcelProcess(OldExcel)
            'MailSend(SFN, gem01, gem02)
            'MsgBox("Finished")
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        'End If
    End Sub
    '200417 add by Brady END
    Private Sub AdjustExcelFormat2(ByVal hh As String)
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = hh
        Ws.Columns.ColumnWidth = 17
        Ws.Cells(1, 1) = "部门编号"
        Ws.Cells(1, 2) = "部门名称"
        Ws.Cells(1, 3) = "凭证日期"
        Ws.Cells(1, 4) = "凭证编号"
        Ws.Cells(1, 5) = "科目编码"
        Ws.Cells(1, 6) = "科目名称"
        Ws.Cells(1, 7) = "摘要"
        Ws.Cells(1, 8) = "发生额"
        'Ws.Cells(1, 8) = "借方"
        'Ws.Cells(1, 9) = "贷方"
        'Ws.Cells(1, 10) = "余额"
        oRng = Ws.Range("H1", "H1")
        oRng.EntireColumn.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        oRng.EntireColumn.ColumnWidth = 17
        LineZ = 2
    End Sub
    '20/04/02 add by Brady
    Private Sub AdjustExcelFormat3(ByVal hh As String)
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = hh
        Ws.Columns.EntireColumn.ColumnWidth = 15
        Ws.Columns.EntireColumn.WrapText = True
        'oRng = Ws.Range("A1", "H1")
        'oRng.Merge()
        oRng = Ws.Range("A1", "K2")
        oRng.EntireRow.RowHeight = 42
        oRng.EntireRow.ColumnWidth = 20
        'oRng = Ws.Range("C2", "M2")
        'oRng.EntireColumn.ColumnWidth = 17.25
        'oRng = Ws.Range("A2", "B2")
        'oRng.EntireColumn.ColumnWidth = 23.28

        Ws.Cells(1, 1) = "Month" & Chr(10) & "月份"
        oRng = Ws.Range("A1", "A2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter

        Ws.Cells(1, 2) = "Actual Part Revenue" & Chr(10) & "实际产品销售金额"
        oRng = Ws.Range("B1", "B2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter

        Ws.Cells(1, 3) = "Budget  Part Revenue" & Chr(10) & "预算产品销售金额"
        oRng = Ws.Range("C1", "C2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter

        Ws.Cells(1, 4) = "Revenue A.%(Actl/Bgt)" & Chr(10) & "收入达成% 实际收入/预算收入"
        oRng = Ws.Range("D1", "D2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter

        Ws.Cells(1, 5) = "Actl Cost Center Exp." & Chr(10) & "实际部门费用"
        oRng = Ws.Range("E1", "E2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter

        Ws.Cells(1, 6) = "Budget Cost Center Exp" & Chr(10) & "原预算费用"
        oRng = Ws.Range("F1", "F2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter

        Ws.Cells(1, 7) = "Act - But" & Chr(10) & "实际费用-原预算费用"
        oRng = Ws.Range("G1", "G2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter

        Ws.Cells(1, 8) = "Exp.A. %(Actl-Bgt)/Bgt" & Chr(10) & "(实际费用-原预算费用)/原预算费用"
        oRng = Ws.Range("H1", "H2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter

        Ws.Cells(1, 9) = "Budget Cost Center Exp.(Revenue A.%)" & Chr(10) & "重计预算费用"
        oRng = Ws.Range("I1", "I2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter

        Ws.Cells(1, 10) = "Act - But" & Chr(10) & "实际费用-重计预算费用"
        oRng = Ws.Range("J1", "J2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter

        Ws.Cells(1, 11) = "Exp.A.%(Actl-Bgt)/Bgt" & Chr(10) & "(实际费用-重计预算费用)/重计预算费用"
        oRng = Ws.Range("K1", "K2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter

        Ws.Cells(15, 1) = "Year YTD"
        oRng = Ws.Range("A15", "A15")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter

        oRng = Ws.Range("A3", "A14")
        oRng.NumberFormatLocal = "mmm-yy"
        oRng.HorizontalAlignment = xlCenter

        oRng = Ws.Range("A1", "K15")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("B3", "C15")
        oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        oRng.HorizontalAlignment = xlCenter

        oRng = Ws.Range("D3", "D15")
        oRng.NumberFormatLocal = "0.00%"
        oRng.HorizontalAlignment = xlCenter

        oRng = Ws.Range("E3", "G15")
        oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        oRng.HorizontalAlignment = xlCenter

        oRng = Ws.Range("H3", "H15")
        oRng.NumberFormatLocal = "0.00%"
        oRng.HorizontalAlignment = xlCenter

        oRng = Ws.Range("J15", "J15")
        oRng.Interior.Color = 65535

        oRng = Ws.Range("I3", "J15")
        oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        oRng.HorizontalAlignment = xlCenter

        oRng = Ws.Range("K3", "K15")
        oRng.NumberFormatLocal = "0.00%"
        oRng.HorizontalAlignment = xlCenter

        LineZ = 3
    End Sub
    '20/04/02 add by Brady END
    '20/04/11 add by Brady
    Private Sub AdjustExcelFormat4(ByVal hh As String)
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = hh
        Ws.Columns.ColumnWidth = 17
        Ws.Cells(1, 1) = "日期"
        Ws.Cells(1, 2) = "成本中心"
        Ws.Cells(1, 3) = "科目代码"
        Ws.Cells(1, 4) = "科目名称"
        Ws.Cells(1, 5) = "Debit Note币别"
        Ws.Cells(1, 6) = "原币金额"
        Ws.Cells(1, 7) = "本币金额"
        Ws.Cells(1, 8) = "Debit Note No."
        oRng = Ws.Range("H1", "H1")
        oRng.EntireColumn.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        oRng.EntireColumn.ColumnWidth = 17
        LineZ = 2
    End Sub
    Public Sub MailSend(ByVal FileName As String, ByVal gem01 As String, ByVal gem02 As String)
        Dim MS As New System.Net.Mail.MailMessage
        Dim MA As New System.Net.Mail.MailAddress("action.server@action-composites.com.cn")
        MS.From = MA
        MS.Subject = "部门费用明细表-" & gem02
        'Dim mConnectionBuilder As New SqlClient.SqlConnectionStringBuilder
        Dim mConnection As New SqlClient.SqlConnection
        Dim mSQLS1 As New SqlClient.SqlCommand
        Dim mSQLReader As SqlClient.SqlDataReader
        mConnection.ConnectionString = Module1.OpenConnectionOfERPSUPPORT()
        If mConnection.State <> ConnectionState.Open Then
            mConnection.Open()
            mSQLS1.Connection = mConnection
            mSQLS1.CommandType = CommandType.Text
            mSQLS1.CommandTimeout = 600
        End If
        mSQLS1.CommandText = "SELECT * FROM MAILLIST WHERE CC = 'TO' AND ProgramName = 'department_expense' And DepartmentCode = '" & gem01 & "'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                MS.To.Add(mSQLReader.Item("MailAddress"))
            End While
        End If
        mSQLReader.Close()

        mSQLS1.CommandText = "SELECT * FROM MAILLIST WHERE CC = 'CC' AND ProgramName = 'department_expense' And DepartmentCode = '" & gem01 & "'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                MS.CC.Add(mSQLReader.Item("MailAddress"))
            End While
        End If
        mSQLReader.Close()

        mSQLS1.CommandText = "SELECT * FROM MAILLIST WHERE CC = 'BCC' AND ProgramName = 'department_expense' And DepartmentCode = '" & gem01 & "'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                MS.Bcc.Add(mSQLReader.Item("MailAddress"))
            End While
        End If

        mSQLReader.Close()
        mConnection.Close()
        MS.Body = "Dear 各部门主管负责人:<BR/>"
        MS.Body += "Dear Cost Center Managers:<BR/>"
        MS.Body += "<BR/>     请开始检阅、分析本部门费用，如当月或本年度累计实际费用有超原预算费用需回复说明原因、金额及改善方式。财务对应人员将会与各部门主管联系，收集回复及执行情况说明。"
        MS.Body += "<BR/>     Attached is cost center expense report.  Please take time to check and review."
        MS.Body += "<BR/>     If actual expenditure amount is over budget expense amount, please provide reason of over-expenditure and feedback improvement actions to control the expenditures."
        MS.Body += "<BR/>     FN in-charge window will contact you to collect your feedback of action plans and comments.<BR/>"
        MS.Body += "<BR/>     各部门主管检阅报表时，请注意原编制预算费用已按收入达成%重计费用预算金额(见附件KPI工作表)。如KPI表中2020年YTD实际总费用超重计预算费用的需要加强管控"
        MS.Body += "<BR/>     各部门主管填写月KPI %（ 费用vs 预算费用管控），请以K栏的%为准。"
        MS.Body += "<BR/>     In the report file, there is one worksheet 'KPI'.  The KPI calculation formula is:"
        MS.Body += "<BR/>     1.      Achievement % based on actual shipment of parts over budget parts revenue."
        MS.Body += "<BR/>     2.      Budget expenses amount x achievement % to get expenditures available."
        MS.Body += "<BR/>     3.      Actual expenditures/expenditures available=KPI %"
        MS.Body += "<BR/>"
        MS.Body += "<BR/>     以上如有任何问题, 请于财务成本组姚巧联系, 谢谢!"
        MS.Body += "<BR/>     Any question, please feel free to contact Sunny, Cost accountant at FN Dept."
        MS.IsBodyHtml = True
        Dim MAM As New System.Net.Mail.Attachment(FileName)
        MS.Attachments.Add(MAM)
        ' 信件做好了
        Dim SMT As New System.Net.Mail.SmtpClient("smtp.action-composites.com.cn")
        SMT.UseDefaultCredentials = True
        'SMT.PickupDirectoryLocation = "C:\temp\ab"
        Dim UAP As New System.Net.NetworkCredential()
        UAP.UserName = "action.server@action-composites.com.cn"
        UAP.Password = "action@2017"
        SMT.Credentials = UAP

        Try
            SMT.Send(MS)
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
    End Sub
    '20/04/11 add by Brady END
End Class