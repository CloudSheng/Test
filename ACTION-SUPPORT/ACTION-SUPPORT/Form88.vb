Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form88
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim DBC As String = String.Empty
    Dim LineZ As Integer = 0
    Dim DNP As String = String.Empty
    Dim ExchangeRate1 As Decimal = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form88_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        NumericUpDown1.Value = Today.Year
        NumericUpDown2.Value = Today.Month
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
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        tYear = Me.NumericUpDown1.Value
        tMonth = Me.NumericUpDown2.Value
        oCommand.CommandText = "select count(*) from ccc_file where ccc02 = " & tYear & " and ccc03 = " & tMonth & " and ccc61 <> 0 and ccc01 not like 'S%'"
        Dim HasRows As Decimal = oCommand.ExecuteScalar()
        If HasRows = 0 Then
            MsgBox("No Data")
            Return
        End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Margin_Analysis_Report"
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
        Ws.Name = "Margin Analysis output"
        Ws.Activate()
        AdjustExcelFormat()
        oCommand.CommandText = "select ccc01,ima02,ccc23,ccc61,ccc63 from ccc_file,ima_file where ccc01 = ima01 and ccc02 = " & tYear & " and ccc03 = " & tMonth & " and ccc61 <> 0 and ccc01 not like 'S%' "
        oCommand.CommandText += "order by ccc01 "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("ccc01")
                Ws.Cells(LineZ, 3) = oReader.Item("ima02")
                Ws.Cells(LineZ, 4) = Decimal.Round(GetStandardCost(oReader.Item("ccc01")) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 5) = Decimal.Round(oReader.Item("ccc23") / ExchangeRate1, 3)
                Ws.Cells(LineZ, 6) = "=J" & LineZ & "/G" & LineZ
                Ws.Cells(LineZ, 7) = oReader.Item("ccc61") * Decimal.MinusOne
                Ws.Cells(LineZ, 8) = "=D" & LineZ & "*G" & LineZ
                Ws.Cells(LineZ, 9) = "=E" & LineZ & "*G" & LineZ
                Ws.Cells(LineZ, 10) = Decimal.Round(oReader.Item("ccc63") / ExchangeRate1, 3)
                If oReader.Item("ccc63") <> 0 Then
                    Ws.Cells(LineZ, 11) = "=(J" & LineZ & "-H" & LineZ & ")/J" & LineZ
                    Ws.Cells(LineZ, 12) = "=(J" & LineZ & "-I" & LineZ & ")/J" & LineZ
                Else
                    Ws.Cells(LineZ, 11) = 0
                    Ws.Cells(LineZ, 12) = 0
                End If
                Ws.Cells(LineZ, 13) = "=J" & LineZ & "-I" & LineZ
                LineZ += 1
            End While
        End If
        oReader.Close()
        ' 劃線
        oRng = Ws.Range("A5", Ws.Cells(LineZ - 1, 13))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        ' 加總
        Ws.Cells(LineZ, 3) = "Total Automotive"
        Ws.Cells(LineZ, 7) = "=SUM(G5:G" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 7))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 10)), Type:=xlFillDefault)

        oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 13))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("D5", Ws.Cells(LineZ, 10))
        oRng.NumberFormatLocal = "#,##0.00_ "

        oRng = Ws.Range("M5", Ws.Cells(LineZ, 13))
        oRng.NumberFormatLocal = "#,##0.00_ "

    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 21.33
        Ws.Columns.HorizontalAlignment = xlCenter
        Ws.Cells(1, 1) = "Margin Analysis for COGS"
        Ws.Cells(2, 1) = tYear & "-" & tMonth
        Select Case DBC
            Case "actiontest"
                Ws.Cells(2, 2) = "Dongguan Action Composites LTD Co."
                Dim TYM1 As String = String.Empty
                If tMonth < 10 Then
                    TYM1 = tYear & "0" & tMonth
                Else
                    TYM1 = tYear & tMonth
                End If
                oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & TYM1 & "'"
                ExchangeRate1 = oCommand.ExecuteScalar()
            Case "hkacttest"
                Ws.Cells(2, 2) = "Action Composite Technology Limited"
                ExchangeRate1 = 1
            Case "action_bvi"
                Ws.Cells(2, 2) = "Action Composites International Limited"
                ExchangeRate1 = 1
        End Select
        oRng = Ws.Range("A3", "A4")
        oRng.Merge()
        Ws.Cells(3, 1) = "Product group"
        oRng = Ws.Range("B3", "B4")
        oRng.Merge()
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(3, 2) = "Part No."
        oRng = Ws.Range("C3", "C4")
        oRng.Merge()
        Ws.Cells(3, 3) = "Part description"
        oRng = Ws.Range("D3", "D4")
        oRng.Merge()
        Ws.Cells(3, 4) = "Standard Cost"
        oRng = Ws.Range("E3", "E4")
        oRng.Merge()
        Ws.Cells(3, 5) = "Actual Cost"
        Ws.Cells(3, 6) = "Selling price"
        Ws.Cells(4, 6) = "USD"
        oRng = Ws.Range("G3", "M3")
        oRng.Merge()
        Ws.Cells(3, 7) = tYear & "/" & tMonth
        oRng = Ws.Range("C4", "D5")
        Ws.Cells(4, 7) = "Qty in Unit"
        Ws.Cells(4, 8) = "CGOS at STD"
        Ws.Cells(4, 9) = "COGS at Actual"
        Ws.Cells(4, 10) = "Sold at Selling Price"
        Ws.Cells(4, 11) = "%Margin at STD"
        Ws.Cells(4, 12) = "%Margin at Actual"
        oRng = Ws.Range("K5", "L5")
        oRng.EntireColumn.NumberFormatLocal = "0.00%"
        Ws.Cells(4, 13) = "Margin at Actual"
        ' 劃線
        oRng = Ws.Range("A3", "M4")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        LineZ = 5
    End Sub
    Private Function GetStandardCost(ByVal ccc01 As String)
        oCommand2.CommandText = "select nvl(sum(stb07+stb08+stb09),0) from stb_file where stb02 = " & tYear & " and stb03 = " & tMonth & " and stb01 = '" & ccc01 & "'"
        Dim SC As Decimal = oCommand2.ExecuteScalar()
        Return SC
    End Function
End Class