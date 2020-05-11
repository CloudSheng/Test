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
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader3 As Oracle.ManagedDataAccess.Client.OracleDataReader
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
        'SaveExcel()
        'BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
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
                If PY > 0 Then
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
                    oRng = Ws.Range("A7", Ws.Cells(LineZ, 18))
                    oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
                    oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
                    oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
                    oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
                    oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                    oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
                    oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
                End If


                ' 第二頁
                oCommand2.CommandText = "SELECT nvl(sum(aao05 + aao06),0) FROM aao_file where aao03 = " & tYear & " and aao04 = " & tMonth & " and aao01 like '6604%' and aao02 = '" & oReader.Item("gem01") & "'"
                Dim PY1 As Decimal = oCommand2.ExecuteScalar()
                If PY1 > 0 Then
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
                    oRng = Ws.Range("A7", Ws.Cells(LineZ, 18))
                    oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
                    oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
                    oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
                    oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
                    oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                    oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
                    oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
                End If

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
                oRng = Ws.Range("A7", Ws.Cells(LineZ, 18))
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
                If PY2 > 0 Then
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
                    oRng = Ws.Range("A7", Ws.Cells(LineZ, 18))
                    oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
                    oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
                    oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
                    oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
                    oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                    oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
                    oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
                End If

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
                SaveExcel(oReader.Item("gem02"))
            End While
        End If
        oReader.Close()
        MsgBox("Finished")
    End Sub
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
    Private Sub SaveExcel(ByVal gem02 As String)
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
            'MsgBox("Finished")
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        'End If
    End Sub
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
        
End Class