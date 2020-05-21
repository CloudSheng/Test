Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form131
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim tDate As Date
    Dim TW As Decimal = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form131_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
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
        BackgroundWorker1.RunWorkerAsync()
        'ExportToExcel()
        'SaveExcel()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        ' 20180717 加入臨時表
        oCommand.CommandText = "DROP TABLE shipfee_temp2"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception

        End Try

        oCommand.CommandText = "Create Table shipfee_temp2 (shareDate date, RMBD number(18, 3), USDD number(18,3), PN varchar2(40), Qty number(15, 3))"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        AdjustExcelFormat1()
        oCommand.CommandText = "select Updatedate, DACInvoice, Currency, oga01,oga02,ACAinvoice,ogb03,ogb04,ima02,ima021,ogb12,ogb05,oga23,oga24,ogb14,fee,oga50 "
        oCommand.CommandText += "from shipfee_temp left join oga_file on dacinvoice = oga27 left join ogb_file on oga01 = ogb01 left join ima_file on ogb04 = ima01 "
        oCommand.CommandText += "where Updatedate >= to_date('" & tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("Updatedate")
                Ws.Cells(LineZ, 2) = oReader.Item("DACInvoice")
                Ws.Cells(LineZ, 3) = oReader.Item("Currency")
                Ws.Cells(LineZ, 4) = "ACTIONTEST"
                Ws.Cells(LineZ, 5) = oReader.Item("oga01")
                Ws.Cells(LineZ, 6) = oReader.Item("oga02")
                Ws.Cells(LineZ, 7) = oReader.Item("ogb03")
                Ws.Cells(LineZ, 8) = oReader.Item("ogb04")
                Ws.Cells(LineZ, 9) = oReader.Item("ima02")
                Ws.Cells(LineZ, 10) = oReader.Item("ima021")
                Ws.Cells(LineZ, 11) = oReader.Item("ogb12")
                Ws.Cells(LineZ, 12) = oReader.Item("ogb05")
                Ws.Cells(LineZ, 13) = oReader.Item("oga23")
                Ws.Cells(LineZ, 14) = oReader.Item("oga24")
                Ws.Cells(LineZ, 15) = oReader.Item("ogb14")
                Dim ES As Decimal = Decimal.Round(oReader.Item("fee") * oReader.Item("ogb14") / oReader.Item("oga50"), 2)
                Ws.Cells(LineZ, 18) = ES

                Dim YM As String = String.Empty
                YM = Convert.ToDateTime(oReader.Item("oga02")).ToString("yyyyMM")
                Dim ExchangeRate1 As Decimal = GetExchangeRate("EUR", YM)
                Dim RMBD As Decimal = Decimal.Round(ES * ExchangeRate1, 2)
                Ws.Cells(LineZ, 16) = RMBD
                Dim ExchangeRate2 As Decimal = GetExchangeRate("USD", YM)
                Dim USDD As Decimal = Decimal.Round(ES * ExchangeRate1 / ExchangeRate2, 2)
                Ws.Cells(LineZ, 17) = USDD
                oCommand2.CommandText = "INSERT INTO shipfee_temp2 VALUES (to_date('" & oReader.Item("oga02") & "','yyyy/mm/dd'), " & RMBD & "," & USDD & ",'" & oReader.Item("ogb04") & "'," & oReader.Item("ogb12") & ")"
                Try
                    oCommand2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
                LineZ += 1
            End While
        End If
        oReader.Close()

        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustExcelFormat2()

        ' 先處理RMB 
        oCommand.CommandText = "SELECT azn05,nvl(sum(rmbd),0) FROM SHIPfee_temp2,azn_file where sharedate = azn01 and year(sharedate) = " & tDate.Year & " group by azn05 order by azn05"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                Dim nW As Decimal = oReader.Item(0)
                Ws.Cells(2, 2 + nW) = oReader.Item(1)
            End While
        End If
        oReader.Close()
        ' 其他
        oCommand.CommandText = "SELECT nvl(sum(rmbd),0) FROM SHIPfee_temp2 where year(sharedate) <> " & tDate.Year
        Dim TX As Decimal = oCommand.ExecuteScalar()
        Ws.Cells(2, 2) = TX

        ' 總合
        oCommand.CommandText = "SELECT nvl(sum(rmbd),0) FROM SHIPfee_temp2"
        Ws.Cells(2, 3 + TW) = oCommand.ExecuteScalar()

        ' 再處理USD 
        oCommand.CommandText = "SELECT azn05,nvl(sum(usdd),0) FROM SHIPfee_temp2,azn_file where sharedate = azn01 and year(sharedate) = " & tDate.Year & " group by azn05 order by azn05"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                Dim nW As Decimal = oReader.Item(0)
                Ws.Cells(6, 2 + nW) = oReader.Item(1)
            End While
        End If
        oReader.Close()
        ' 其他
        oCommand.CommandText = "SELECT nvl(sum(usdd),0) FROM SHIPfee_temp2 where year(sharedate) <> " & tDate.Year
        Dim TX1 As Decimal = oCommand.ExecuteScalar()
        Ws.Cells(6, 2) = TX1

        ' 總合
        oCommand.CommandText = "SELECT nvl(sum(usdd),0) FROM SHIPfee_temp2"
        Ws.Cells(6, 3 + TW) = oCommand.ExecuteScalar()

        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        AdjustExcelFormat3()

        oCommand.CommandText = "select pn,ima02,ima021,ima25,sum(qty) as t1,sum(rmbd) as t2,sum(usdd) as t3 from shipfee_temp2,ima_file where pn = ima01 group by pn,ima02,ima021,ima25"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("pn")
                Ws.Cells(LineZ, 2) = oReader.Item("ima02")
                Ws.Cells(LineZ, 3) = oReader.Item("ima021")
                Ws.Cells(LineZ, 4) = oReader.Item("ima25")
                Ws.Cells(LineZ, 5) = oReader.Item("t1")
                Ws.Cells(LineZ, 7) = oReader.Item("t2")
                Ws.Cells(LineZ, 8) = oReader.Item("t3")
                Ws.Cells(LineZ, 6) = "=G" & LineZ & "/E" & LineZ
                LineZ += 1
            End While
        End If
        oReader.Close()

    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10

        Ws.Name = "明細表"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireRow.WrapText = True

        Ws.Cells(1, 1) = "Date"
        Ws.Cells(1, 2) = "Invoice No. DAC"
        Ws.Cells(1, 3) = "Currency"
        Ws.Cells(1, 4) = "Operation Center"
        Ws.Cells(1, 5) = "Delivery Note No."
        Ws.Cells(1, 6) = "Shipment Date"
        Ws.Cells(1, 7) = "Position No."
        Ws.Cells(1, 8) = "Part Name"
        Ws.Cells(1, 9) = "Part Description"
        Ws.Cells(1, 10) = "Spec."
        Ws.Cells(1, 11) = "Shipping Qty"
        Ws.Cells(1, 12) = "Unit"
        Ws.Cells(1, 13) = "Currency"
        Ws.Cells(1, 14) = "Exchange Rate"
        Ws.Cells(1, 15) = "Contract Amount"
        Ws.Cells(1, 16) = "Air-freight recharge RMB"
        Ws.Cells(1, 17) = "Air-freight recharge USD"
        Ws.Cells(1, 18) = "Air-freight recharge EUR"

        oRng = Ws.Range("E1", "H1")
        oRng.EntireColumn.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        LineZ = 2
    End Sub
    Private Function GetExchangeRate(ByVal tCurrency As String, ByVal YM As String)
        oCommand2.CommandText = "SELECT nvl(azj041,1) FROM azj_file WHERE azj01 = '" & tCurrency & "' and azj02 = '" & YM & "'"
        Dim Td As Decimal = oCommand2.ExecuteScalar()
        Return Td
    End Function
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Transportation Fee"
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
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10
        Ws.Columns.HorizontalAlignment = xlCenter
        Ws.Name = "weekly"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 23.11

        Ws.Cells(1, 1) = "Air-freight recharge RMB"
        Ws.Cells(1, 2) = "Others"
        Ws.Cells(5, 1) = "Air-freight recharge USD"
        Ws.Cells(5, 2) = "Others"

        oCommand.CommandText = "select max(azn05) from azn_file where year(azn01) = " & tDate.Year
        TW = oCommand.ExecuteScalar()

        For i As Int16 = 1 To TW Step 1
            Ws.Cells(1, 2 + i) = "W" & i
            Ws.Cells(5, 2 + i) = "W" & i
            Ws.Cells(2, 2 + i) = 0
            Ws.Cells(6, 2 + i) = 0
        Next
        Ws.Cells(1, 3 + TW) = "Total"
        Ws.Cells(5, 3 + TW) = "Total"
        Ws.Cells(2, 1) = "Total"
        Ws.Cells(6, 1) = "Total"

        oRng = Ws.Range("B2", "B2")
        oRng.EntireRow.NumberFormat = "#,##0_ ;-#,##0 "

        oRng = Ws.Range("B6", "B6")
        oRng.EntireRow.NumberFormat = "#,##0_ ;-#,##0 "

        oRng = Ws.Range("A1", Ws.Cells(2, 3 + TW))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("A5", Ws.Cells(6, 3 + TW))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat3()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10

        Ws.Name = "air freight by part"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireRow.WrapText = True

        Ws.Cells(1, 1) = "Part Name"
        Ws.Cells(1, 2) = "Part Description"
        Ws.Cells(1, 3) = "Spec."
        Ws.Cells(1, 4) = "Unit"
        Ws.Cells(1, 5) = "Shipping Qty"
        Ws.Cells(1, 6) = "Air-freight recharge/pcs RMB"
        Ws.Cells(1, 7) = "Air-freight recharge RMB"
        Ws.Cells(1, 8) = "Air-freight recharge USD"
        LineZ = 2
    End Sub
End Class