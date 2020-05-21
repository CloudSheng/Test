Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel

Public Class Form357
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
    Dim cYear As Int16 = 0
    Dim cMonth As Int16 = 0
    Dim y_cnt As Int16 = 0
    Dim m_cnt As Int16 = 0
    Dim pYear As Int16 = 0
    Dim tDate As Date
    Dim DBC As String = String.Empty
    Dim LineZ As Integer = 0
    Dim LineOZ As Integer = 0
    Dim l_amt As Decimal = 0
    Dim ll_amt As Decimal = 0
    Dim DNP As String = String.Empty
    Dim gDatabase As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Dim SaveFileDialog1 As New SaveFileDialog
    Private Sub Form357_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'If Me.BackgroundWorker1.IsBusy() Then
        'MsgBox("处理中，请等待")
        'Return
        'End If        
        'DBC = "hkacttest"
        'oConnection.ConnectionString = Module1.OpenOracleConnection(DBC)
        'If oConnection.State <> ConnectionState.Open Then
        '    Try
        '        oConnection.Open()
        '        oCommand.Connection = oConnection
        '        oCommand.CommandType = CommandType.Text
        '        oCommand2.Connection = oConnection
        '        oCommand2.CommandType = CommandType.Text
        '        oCommand3.Connection = oConnection
        '        oCommand3.CommandType = CommandType.Text
        '    Catch ex As Exception
        '        MsgBox(ex.Message)
        '    End Try
        'End If

        'DBC = "actiontest"
        'oConnection.ConnectionString = Module1.OpenOracleConnection(DBC)
        gDatabase = Me.ComboBox2.SelectedItem.ToString()
        If String.IsNullOrEmpty(gDatabase) Then
            MsgBox("Database Error")
            Return
        End If
        Select Case gDatabase
            Case "DAC"
                oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
            Case "HAC"
                oConnection.ConnectionString = Module1.OpenOracleConnection("hkacttest")
            Case "BVI"
                oConnection.ConnectionString = Module1.OpenOracleConnection("action_bvi")
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

        tYear = Me.DateTimePicker1.Value.Year
        tMonth = Me.DateTimePicker1.Value.Month
        eYear = Me.DateTimePicker2.Value.Year
        eMonth = Me.DateTimePicker2.Value.Month

        If (tYear = eYear And eMonth < tMonth) Or (tYear > eYear) Then
            MsgBox("年月區間輸入錯誤")
            oConnection.Close()
            Return
        End If

        y_cnt = eYear - tYear

        If y_cnt >= 1 Then
            m_cnt = (y_cnt - 1) * 12 + ((12 - tMonth) + 1) + eMonth
        Else
            m_cnt = eMonth - tMonth + 1
        End If

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

        'DBC = "actiontest"
        'oConnection.ConnectionString = Module1.OpenOracleConnection(DBC)
        'If oConnection.State <> ConnectionState.Open Then
        '    Try
        '        oConnection.Open()
        '        oCommand.Connection = oConnection
        '        oCommand.CommandType = CommandType.Text
        '        oCommand2.Connection = oConnection
        '        oCommand2.CommandType = CommandType.Text
        '        oCommand3.Connection = oConnection
        '        oCommand3.CommandType = CommandType.Text
        '    Catch ex As Exception
        '        MsgBox(ex.Message)
        '    End Try
        'End If

        SaveExcel()
    End Sub

    Private Sub ExportToExcel()
        ' 第一頁 (外币余额)    
        Ws = xWorkBook.Sheets(1)
        Ws.Name = "外币余额"
        Ws.Activate()
        AdjustExcelFormat()
        'oCommand.CommandText = "select unique tah08 from tah_file where tah01 like '660129%' and aag07 = 2 order by aag01"
        oCommand.CommandText = "select azi01 from azi_file order by azi01"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()

                cYear = tYear
                cMonth = tMonth
                LineOZ = LineZ
                For i = 1 To m_cnt
                    Ws.Cells(LineZ, 1) = tDate.AddMonths(i - 1)
                    Ws.Cells(LineZ, 2) = oReader.Item("azi01")
                    l_amt = Decimal.Round(GetThisYearMonth(oReader.Item("azi01").ToString(), cYear, cMonth), 3)
                    ll_amt = Decimal.Round(GetThisYearMonth_1(oReader.Item("azi01").ToString(), cYear, cMonth), 3)
                    Ws.Cells(LineZ, 3) = l_amt + ll_amt
                    cMonth += 1
                    If cMonth > 12 Then
                        cMonth = 1
                        cYear += 1
                        End If
                    LineZ += 1
                Next
                Ws.Cells(LineZ, 1) = "小计"
                Ws.Cells(LineZ, 3) = "=SUM(C" & LineOZ & ":C" & LineZ - 1 & ")"

                cYear = tYear
                cMonth = tMonth
                LineZ = LineOZ
                For i = 1 To m_cnt
                    l_amt = Decimal.Round(GetThisYearMonth_2(oReader.Item("azi01").ToString(), cYear, cMonth), 3)
                    ll_amt = Decimal.Round(GetThisYearMonth_3(oReader.Item("azi01").ToString(), cYear, cMonth), 3)
                    Ws.Cells(LineZ, 4) = l_amt + ll_amt
                    cMonth += 1
                    If cMonth > 12 Then
                        cMonth = 1
                        cYear += 1
                    End If
                    LineZ += 1
                Next
                Ws.Cells(LineZ, 4) = "=SUM(D" & LineOZ & ":D" & LineZ - 1 & ")"

                cYear = tYear
                cMonth = tMonth
                LineZ = LineOZ
                For i = 1 To m_cnt
                    l_amt = Decimal.Round(GetThisYearMonth_4(oReader.Item("azi01").ToString(), cYear, cMonth), 3)
                    ll_amt = Decimal.Round(GetThisYearMonth_5(oReader.Item("azi01").ToString(), cYear, cMonth), 3)
                    Ws.Cells(LineZ, 5) = l_amt + ll_amt
                    cMonth += 1
                    If cMonth > 12 Then
                        cMonth = 1
                        cYear += 1
                    End If
                    LineZ += 1
                Next
                Ws.Cells(LineZ, 5) = "=SUM(E" & LineOZ & ":E" & LineZ - 1 & ")"

                cYear = tYear
                cMonth = tMonth
                LineZ = LineOZ
                For i = 1 To m_cnt
                    l_amt = Decimal.Round(GetThisYearMonth_6(oReader.Item("azi01").ToString(), cYear, cMonth), 3)
                    ll_amt = Decimal.Round(GetThisYearMonth_7(oReader.Item("azi01").ToString(), cYear, cMonth), 3)
                    Ws.Cells(LineZ, 6) = l_amt + ll_amt
                    cMonth += 1
                    If cMonth > 12 Then
                        cMonth = 1
                        cYear += 1
                    End If
                    LineZ += 1
                Next
                Ws.Cells(LineZ, 6) = "=SUM(F" & LineOZ & ":F" & LineZ - 1 & ")"

                cYear = tYear
                cMonth = tMonth
                LineZ = LineOZ
                For i = 1 To m_cnt
                    l_amt = Decimal.Round(GetThisYearMonth_8(oReader.Item("azi01").ToString(), cYear, cMonth), 3)
                    ll_amt = Decimal.Round(GetThisYearMonth_9(oReader.Item("azi01").ToString(), cYear, cMonth), 3)
                    Ws.Cells(LineZ, 7) = l_amt + ll_amt
                    cMonth += 1
                    If cMonth > 12 Then
                        cMonth = 1
                        cYear += 1
                    End If
                    LineZ += 1
                Next
                Ws.Cells(LineZ, 7) = "=SUM(G" & LineOZ & ":G" & LineZ - 1 & ")"

                cYear = tYear
                cMonth = tMonth
                LineZ = LineOZ
                For i = 1 To m_cnt
                    l_amt = Decimal.Round(GetThisYearMonth_10(oReader.Item("azi01").ToString(), cYear, cMonth), 3)
                    ll_amt = Decimal.Round(GetThisYearMonth_11(oReader.Item("azi01").ToString(), cYear, cMonth), 3)
                    Ws.Cells(LineZ, 8) = l_amt + ll_amt
                    cMonth += 1
                    If cMonth > 12 Then
                        cMonth = 1
                        cYear += 1
                    End If
                    LineZ += 1
                Next
                Ws.Cells(LineZ, 8) = "=SUM(H" & LineOZ & ":H" & LineZ - 1 & ")"

                cYear = tYear
                cMonth = tMonth
                LineZ = LineOZ
                For i = 1 To m_cnt
                    l_amt = Decimal.Round(GetThisYearMonth_12(oReader.Item("azi01").ToString(), cYear, cMonth), 3)
                    ll_amt = Decimal.Round(GetThisYearMonth_13(oReader.Item("azi01").ToString(), cYear, cMonth), 3)
                    Ws.Cells(LineZ, 9) = l_amt + ll_amt
                    cMonth += 1
                    If cMonth > 12 Then
                        cMonth = 1
                        cYear += 1
                    End If
                    LineZ += 1
                Next
                Ws.Cells(LineZ, 9) = "=SUM(I" & LineOZ & ":I" & LineZ - 1 & ")"

                cYear = tYear
                cMonth = tMonth
                LineZ = LineOZ
                For i = 1 To m_cnt
                    l_amt = Decimal.Round(GetThisYearMonth_14(oReader.Item("azi01").ToString(), cYear, cMonth), 3)
                    ll_amt = Decimal.Round(GetThisYearMonth_15(oReader.Item("azi01").ToString(), cYear, cMonth), 3)
                    Ws.Cells(LineZ, 10) = l_amt + ll_amt
                    cMonth += 1
                    If cMonth > 12 Then
                        cMonth = 1
                        cYear += 1
                    End If
                    LineZ += 1
                Next
                Ws.Cells(LineZ, 10) = "=SUM(J" & LineOZ & ":J" & LineZ - 1 & ")"

                LineZ += 2

                'Ws.Cells(6, m_cnt + 3) = "YTD " & tYear

                'Ws.Cells(LineZ, 1) = oReader.Item("aag01")
                'Ws.Cells(LineZ, 2) = oReader.Item("aag02")
                'oRng = Ws.Cells(LineZ, 2)
                'oRng.WrapText = True
                'oRng.ColumnWidth = 55

                'For i = 1 To m_cnt
                '    cMonth = tMonth + i - 1
                '    Ws.Cells(LineZ, i + 2) = Decimal.Round(GetThisYearMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                'Next

            End While
        End If
        oReader.Close()
        'Ws.Cells(LineZ, 2) = "sut total in USD"
        'Ws.Cells(LineZ, 3) = "=SUM(C7:C" & LineZ - 1 & ")"
        'oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
        'oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, m_cnt + 3)), Type:=xlFillDefault)

        ' 劃線
        'If m_cnt = 1 Then
        '    oRng = Ws.Range("B7", Ws.Cells(LineZ, 4))
        'End If
        'If m_cnt = 2 Then
        '    oRng = Ws.Range("B7", Ws.Cells(LineZ, 5))
        'End If
        'If m_cnt = 3 Then
        '    oRng = Ws.Range("B7", Ws.Cells(LineZ, 6))
        'End If
        
        'oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        'oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        'oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        'oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "

        'oRng = Ws.Range("A6", Ws.Cells(LineZ, 1))
        'oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        'oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        'oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        'If m_cnt = 1 Then
        '    oRng = Ws.Range("C6", "C6")
        'End If
        'If m_cnt = 2 Then
        '    oRng = Ws.Range("C6", "D6")
        'End If
        'If m_cnt = 3 Then
        '    oRng = Ws.Range("C6", "E6")
        'End If
        
        'oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        'oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        'oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
    End Sub

    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 15

        oRng = Ws.Range("A6", "A9999")
        oRng.NumberFormatLocal = "yyyy年mm月"
        oRng = Ws.Range("C5", "J5")
        'oRng.Font =
        oRng = Ws.Range("C6", "J9999")
        oRng.NumberFormat = "#,##0.## ;-#,##0.## "

        'Ws.Cells(1, 3) = "112201"
        'Ws.Cells(2, 3) = "to"
        'Ws.Cells(3, 3) = "112202"

        'Ws.Cells(1, 4) = "12210101"
        'Ws.Cells(2, 4) = "to"
        'Ws.Cells(3, 4) = "122102"

        'Ws.Cells(1, 5) = "220201"
        'Ws.Cells(2, 5) = "to"
        'Ws.Cells(3, 5) = "220204"

        'Ws.Cells(1, 6) = "22410101"
        'Ws.Cells(2, 6) = "to"
        'Ws.Cells(3, 6) = "224102"

        'Ws.Cells(1, 7) = "2203"
        'Ws.Cells(2, 7) = "to"
        'Ws.Cells(3, 7) = "2204"

        'Ws.Cells(3, 8) = "1123"

        'Ws.Cells(1, 9) = "100101"
        'Ws.Cells(2, 9) = "to"
        'Ws.Cells(3, 9) = "100299"

        'Ws.Cells(3, 10) = "2001+2501"

        'Ws.Cells(3, 1) = "会计科目余额"
        Ws.Cells(5, 2) = "币种"
        Ws.Cells(5, 3) = "应收账款余额"
        Ws.Cells(5, 4) = "其他应收款余额"
        Ws.Cells(5, 5) = "应付账款余额"
        Ws.Cells(5, 6) = "其他应付款余额"
        Ws.Cells(5, 7) = "预收账款余额"
        Ws.Cells(5, 8) = "预付账款余额"
        Ws.Cells(5, 9) = "现金余额"
        Ws.Cells(5, 10) = "借款余额"

        'ExchangeRate1 = 1
        'Ws.Cells(6, 1) = "Account"
        'Ws.Cells(6, 2) = "Month"
        'If m_cnt = 1 Then
        '    oRng = Ws.Range("C6", "C6")
        'End If
        'If m_cnt = 2 Then
        '    oRng = Ws.Range("C6", "D6")
        'End If
        'If m_cnt = 3 Then
        '    oRng = Ws.Range("C6", "E6")
        'End If        
        'oRng.NumberFormatLocal = "mmm-yy"

        'For i = 1 To m_cnt
        '    Ws.Cells(6, i + 2) = tDate.AddMonths(i - 1)
        'Next
        'Ws.Cells(6, m_cnt + 3) = "YTD " & tYear

        ' 劃線
        'If m_cnt = 1 Then
        '    oRng = Ws.Range("B6", "D6")
        'End If
        'If m_cnt = 2 Then
        '    oRng = Ws.Range("B6", "E6")
        'End If        
        'oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        'oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        'oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        'If m_cnt = 1 Then
        '    oRng = Ws.Range("C6", "D6")
        'End If
        'If m_cnt = 2 Then
        '    oRng = Ws.Range("C6", "E6")
        'End If
        'If m_cnt = 3 Then
        '    oRng = Ws.Range("C6", "F6")
        'End If
        'oRng.HorizontalAlignment = xlRight

        LineZ = 6
    End Sub

    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "foreign_currency_balance_sheet"
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
    Private Function GetThisYearMonth(ByVal cAzi01 As String, ccYear As Integer, ccMonth As Integer)
        oCommand2.CommandText = "select nvl(SUM(tah09-tah10),0) from tah_file,aag_file where aag07 IN ('2','3') and tah01 = aag01 and aag01 between '112201' and '112202' and tah08= '" & cAzi01 & "' and tah02 = "
        oCommand2.CommandText += ccYear & " and tah03 >= 0 and tah03 < " & ccMonth
        Dim TYTM_1 As Decimal = oCommand2.ExecuteScalar()
        Return TYTM_1
    End Function

    Private Function GetThisYearMonth_1(ByVal cAzi01 As String, ccYear As Integer, ccMonth As Integer)
        oCommand2.CommandText = "select nvl(SUM(tah09-tah10),0) from tah_file,aag_file where aag07 IN ('2','3') and tah01 = aag01 and aag01 between '112201' and '112202' and tah08= '" & cAzi01 & "' and tah02 = "
        oCommand2.CommandText += ccYear & " and tah03 = " & ccMonth
        Dim TYTM_1 As Decimal = oCommand2.ExecuteScalar()
        Return TYTM_1
    End Function

    Private Function GetThisYearMonth_2(ByVal cAzi01 As String, ccYear As Integer, ccMonth As Integer)
        oCommand2.CommandText = "select nvl(SUM(tah09-tah10),0) from tah_file,aag_file where aag07 IN ('2','3') and tah01 = aag01 and aag01 between '12210101' and '122102' and tah08= '" & cAzi01 & "' and tah02 = "
        oCommand2.CommandText += ccYear & " and tah03 >= 0 and tah03 < " & ccMonth
        Dim TYTM_2 As Decimal = oCommand2.ExecuteScalar()
        Return TYTM_2
    End Function

    Private Function GetThisYearMonth_3(ByVal cAzi01 As String, ccYear As Integer, ccMonth As Integer)
        oCommand2.CommandText = "select nvl(SUM(tah09-tah10),0) from tah_file,aag_file where aag07 IN ('2','3') and tah01 = aag01 and aag01 between '12210101' and '122102' and tah08= '" & cAzi01 & "' and tah02 = "
        oCommand2.CommandText += ccYear & " and tah03 = " & ccMonth
        Dim TYTM_3 As Decimal = oCommand2.ExecuteScalar()
        Return TYTM_3
    End Function

    Private Function GetThisYearMonth_4(ByVal cAzi01 As String, ccYear As Integer, ccMonth As Integer)
        oCommand2.CommandText = "select nvl(SUM(tah09-tah10),0) from tah_file,aag_file where aag07 IN ('2','3') and tah01 = aag01 and aag01 between '220201' and  '220204' and tah08= '" & cAzi01 & "' and tah02 = "
        oCommand2.CommandText += ccYear & " and tah03 >= 0 and tah03 < " & ccMonth
        Dim TYTM_4 As Decimal = oCommand2.ExecuteScalar()
        Return TYTM_4
    End Function

    Private Function GetThisYearMonth_5(ByVal cAzi01 As String, ccYear As Integer, ccMonth As Integer)
        oCommand2.CommandText = "select nvl(SUM(tah09-tah10),0) from tah_file,aag_file where aag07 IN ('2','3') and tah01 = aag01 and aag01 between '220201' and  '220204' and tah08= '" & cAzi01 & "' and tah02 = "
        oCommand2.CommandText += ccYear & " and tah03 = " & ccMonth
        Dim TYTM_5 As Decimal = oCommand2.ExecuteScalar()
        Return TYTM_5
    End Function

    Private Function GetThisYearMonth_6(ByVal cAzi01 As String, ccYear As Integer, ccMonth As Integer)
        oCommand2.CommandText = "select nvl(SUM(tah09-tah10),0) from tah_file,aag_file where aag07 IN ('2','3') and tah01 = aag01 and aag01 between '22410101' and '224102' and tah08= '" & cAzi01 & "' and tah02 = "
        oCommand2.CommandText += ccYear & " and tah03 >= 0 and tah03 < " & ccMonth
        Dim TYTM_6 As Decimal = oCommand2.ExecuteScalar()
        Return TYTM_6
    End Function

    Private Function GetThisYearMonth_7(ByVal cAzi01 As String, ccYear As Integer, ccMonth As Integer)
        oCommand2.CommandText = "select nvl(SUM(tah09-tah10),0) from tah_file,aag_file where aag07 IN ('2','3') and tah01 = aag01 and aag01 between '22410101' and '224102' and tah08= '" & cAzi01 & "' and tah02 = "
        oCommand2.CommandText += ccYear & " and tah03 = " & ccMonth
        Dim TYTM_7 As Decimal = oCommand2.ExecuteScalar()
        Return TYTM_7
    End Function

    Private Function GetThisYearMonth_8(ByVal cAzi01 As String, ccYear As Integer, ccMonth As Integer)
        oCommand2.CommandText = "select nvl(SUM(tah09-tah10),0) from tah_file,aag_file where aag07 IN ('2','3') and tah01 = aag01 and aag01 between '2203' and '2204' and tah08= '" & cAzi01 & "' and tah02 = "
        oCommand2.CommandText += ccYear & " and tah03 >= 0 and tah03 < " & ccMonth
        Dim TYTM_8 As Decimal = oCommand2.ExecuteScalar()
        Return TYTM_8
    End Function

    Private Function GetThisYearMonth_9(ByVal cAzi01 As String, ccYear As Integer, ccMonth As Integer)
        oCommand2.CommandText = "select nvl(SUM(tah09-tah10),0) from tah_file,aag_file where aag07 IN ('2','3') and tah01 = aag01 and aag01 between '2203' and '2204' and tah08= '" & cAzi01 & "' and tah02 = "
        oCommand2.CommandText += ccYear & " and tah03 = " & ccMonth
        Dim TYTM_9 As Decimal = oCommand2.ExecuteScalar()
        Return TYTM_9
    End Function

    Private Function GetThisYearMonth_10(ByVal cAzi01 As String, ccYear As Integer, ccMonth As Integer)
        oCommand2.CommandText = "select nvl(SUM(tah09-tah10),0) from tah_file,aag_file where aag07 IN ('2','3') and tah01 = aag01 and aag01 between '1123' and '1123' and tah08= '" & cAzi01 & "' and tah02 = "
        oCommand2.CommandText += ccYear & " and tah03 >= 0 and tah03 < " & ccMonth
        Dim TYTM_10 As Decimal = oCommand2.ExecuteScalar()
        Return TYTM_10
    End Function

    Private Function GetThisYearMonth_11(ByVal cAzi01 As String, ccYear As Integer, ccMonth As Integer)
        oCommand2.CommandText = "select nvl(SUM(tah09-tah10),0) from tah_file,aag_file where aag07 IN ('2','3') and tah01 = aag01 and aag01 between '1123' and '1123' and tah08= '" & cAzi01 & "' and tah02 = "
        oCommand2.CommandText += ccYear & " and tah03 = " & ccMonth
        Dim TYTM_11 As Decimal = oCommand2.ExecuteScalar()
        Return TYTM_11
    End Function

    Private Function GetThisYearMonth_12(ByVal cAzi01 As String, ccYear As Integer, ccMonth As Integer)
        oCommand2.CommandText = "select nvl(SUM(tah09-tah10),0) from tah_file,aag_file where aag07 IN ('2','3') and tah01 = aag01 and aag01 between '100101' and '100299' and tah08= '" & cAzi01 & "' and tah02 = "
        oCommand2.CommandText += ccYear & " and tah03 >= 0 and tah03 < " & ccMonth
        Dim TYTM_12 As Decimal = oCommand2.ExecuteScalar()
        Return TYTM_12
    End Function

    Private Function GetThisYearMonth_13(ByVal cAzi01 As String, ccYear As Integer, ccMonth As Integer)
        oCommand2.CommandText = "select nvl(SUM(tah09-tah10),0) from tah_file,aag_file where aag07 IN ('2','3') and tah01 = aag01 and aag01 between '100101' and '100299' and tah08= '" & cAzi01 & "' and tah02 = "
        oCommand2.CommandText += ccYear & " and tah03 = " & ccMonth
        Dim TYTM_13 As Decimal = oCommand2.ExecuteScalar()
        Return TYTM_13
    End Function

    Private Function GetThisYearMonth_14(ByVal cAzi01 As String, ccYear As Integer, ccMonth As Integer)
        oCommand2.CommandText = "select nvl(SUM(tah09-tah10),0) from tah_file,aag_file where aag07 IN ('2','3') and tah01 = aag01 and aag01 in ('2001','2501') and tah08= '" & cAzi01 & "' and tah02 = "
        oCommand2.CommandText += ccYear & " and tah03 >= 0 and tah03 < " & ccMonth
        Dim TYTM_14 As Decimal = oCommand2.ExecuteScalar()
        Return TYTM_14
    End Function

    Private Function GetThisYearMonth_15(ByVal cAzi01 As String, ccYear As Integer, ccMonth As Integer)
        oCommand2.CommandText = "select nvl(SUM(tah09-tah10),0) from tah_file,aag_file where aag07 IN ('2','3') and tah01 = aag01 and aag01 in ('2001','2501') and tah08= '" & cAzi01 & "' and tah02 = "
        oCommand2.CommandText += ccYear & " and tah03 = " & ccMonth
        Dim TYTM_15 As Decimal = oCommand2.ExecuteScalar()
        Return TYTM_15
    End Function

End Class