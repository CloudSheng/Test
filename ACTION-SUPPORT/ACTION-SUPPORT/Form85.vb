Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports System.Drawing
Public Class Form85
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
    Dim pYear As Int16 = 0
    Dim lYear As Int16 = 0
    Dim lMonth As Int16 = 0
    Dim tDate As Date
    Dim DBC As String = String.Empty
    Dim LineZ As Integer = 0
    Dim DNP As String = String.Empty
    Dim ExchangeRate1 As Decimal = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form85_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If IsNothing(Me.ComboBox1.SelectedItem) Then
            MsgBox("未选定营运中心")
            Return
        End If
        DBC = Me.ComboBox1.SelectedItem.ToString.ToLower()
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
        pYear = Me.DateTimePicker1.Value.AddYears(-1).Year
        tDate = Convert.ToDateTime(tYear & "/" & tMonth & "/01")
        pYear = tDate.AddYears(-1).Year
        lYear = Me.DateTimePicker1.Value.AddMonths(-1).Year
        lMonth = Me.DateTimePicker1.Value.AddMonths(-1).Month
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
        SaveFileDialog1.FileName = "ADM_Expense_Report"
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
        Ws = xWorkBook.Sheets(1)
        Ws.Name = "ADM Exp 汇总"
        Ws.Activate()
        AdjustExcelFormat()
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

        'oRng = Ws.Range("C7", Ws.Cells(LineZ, 9))
        'oRng.Interior.Color = Color.FromArgb(255, 231, 153)
        'oRng = Ws.Range("H7", Ws.Cells(LineZ, 14))
        'oRng.Interior.Color = Color.FromArgb(217, 226, 243)
        'oRng = Ws.Range("L7", Ws.Cells(LineZ, 17))
        'oRng.Interior.Color = Color.FromArgb(197, 224, 178)

        ' 第二頁
        Ws = xWorkBook.Sheets(2)
        Ws.Name = "ADM Exp 部门明细"
        Ws.Activate()
        AdjustExcelFormat1()
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
        'oRng = Ws.Range("D7", Ws.Cells(LineZ, 10))
        'oRng.Interior.Color = Color.FromArgb(255, 231, 153)
        'oRng = Ws.Range("I7", Ws.Cells(LineZ, 15))
        'oRng.Interior.Color = Color.FromArgb(217, 226, 243)
        'oRng = Ws.Range("M7", Ws.Cells(LineZ, 18))
        'oRng.Interior.Color = Color.FromArgb(197, 224, 178)
    End Sub

    Private Sub AdjustExcelFormat()
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
        Select Case DBC
            Case "actiontest"
                Ws.Cells(6, 2) = "Dongguan Action Composites LTD Co."
                Dim TYM1 As String = String.Empty
                If tMonth < 10 Then
                    TYM1 = tYear & "0" & tMonth
                Else
                    TYM1 = tYear & tMonth
                End If
                oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & TYM1 & "'"
                ExchangeRate1 = oCommand.ExecuteScalar()

            Case "hkacttest"
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
    Private Sub AdjustExcelFormat1()
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
        Select Case DBC
            Case "actiontest"
                Ws.Cells(6, 2) = "Dongguan Action Composites LTD Co."
            Case "hkacttest"
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
    Private Function GetLastMonth(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah01 = '" & aag01 & "' and aah02 = "
        oCommand2.CommandText += lYear & " and aah03 = " & lMonth
        Dim TYTM As Decimal = oCommand2.ExecuteScalar()
        Return TYTM
    End Function
    Private Function GetLastMonth(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
        oCommand2.CommandText += lYear & " and aao04 = " & lMonth
        Dim TYTM As Decimal = oCommand2.ExecuteScalar()
        Return TYTM
    End Function
End Class