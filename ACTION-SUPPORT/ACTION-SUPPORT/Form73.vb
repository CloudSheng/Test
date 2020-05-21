Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel.XlChartType
Public Class Form73
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim CC As Integer = 0
    Dim DStartN As Date
    Dim LineZ As Integer = 0
    Dim LineX As Integer = 0
    Dim DW1 As Integer = 0
    Dim DW2 As Integer = 0
    Dim DW3 As Integer = 0
    Dim DW4 As Integer = 0
    Dim C1 As Int16 = 0
    Dim C2 As Int16 = 0
    Dim DS As Data.DataSet = New DataSet()
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If Me.DateTimePicker1.Value.AddDays(1).Month = Me.DateTimePicker1.Value.Month Then
            MsgBox("请选择月底日期")
            Return
        End If
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        C1 = Me.DateTimePicker1.Value.Year
        C2 = Me.DateTimePicker1.Value.Month
        DS.Clear()
        DStartN = Me.DateTimePicker1.Value
        '先訂位
        DW1 = GetAzn02(DStartN)
        DW2 = GetAzn05(DStartN)
        Label1.Text = "处理中"
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Form73_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("action_bvi")
        Me.DateTimePicker1.Value = Today.AddDays((Today.Day) * Decimal.MinusOne)
    End Sub
    Private Function GetAzn02(ByVal eDate As Date)
        Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander2.Connection = oConnection
        oCommander2.CommandType = CommandType.Text
        oCommander2.CommandText = "SELECT azn02 FROM azn_file where azn01 = TO_DATE('" & eDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        Dim ADW1 As Integer = oCommander2.ExecuteScalar()
        Return ADW1
    End Function
    Private Function GetAzn05(ByVal eDate As Date)
        Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander2.Connection = oConnection
        oCommander2.CommandType = CommandType.Text
        oCommander2.CommandText = "SELECT azn05 FROM azn_file where azn01 = TO_DATE('" & eDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        Dim ADW2 As Integer = oCommander2.ExecuteScalar()
        Return ADW2
    End Function
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Label1.Text = "处理完毕"
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "BVI_overdue_AR_aging_report_" & DStartN.ToString("yyyyMMdd")
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
        xWorkBook.Sheets.Add()  '第四頁
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        Ws.Name = "BVIAR"
        AdjustExcelFormat()
        oCommand.CommandText = "SELECT oma68,oma32,oma01,oma67,oma09,oma11,alz08,alz09 FROM OMA_FILE,Alz_file  WHERE  alz09 > 0 and oma03 = alz01 AND alz00 = '2' AND alz02 = "
        oCommand.CommandText += C1 & " AND alz03 = " & C2 & " AND alz04 = oma01 "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("oma68")
                Ws.Cells(LineZ, 2) = oReader.Item("oma32")
                Ws.Cells(LineZ, 3) = oReader.Item("oma01")
                Ws.Cells(LineZ, 4) = oReader.Item("oma67")
                Ws.Cells(LineZ, 5) = oReader.Item("oma09")
                Ws.Cells(LineZ, 6) = oReader.Item("oma11")
                Ws.Cells(LineZ, 7) = oReader.Item("alz08")
                Ws.Cells(LineZ, 8) = oReader.Item("alz09")
                If oReader.Item("oma11") <= DStartN Then
                    Ws.Cells(LineZ, 9) = oReader.Item("alz09")
                Else
                    Setwk(oReader.Item("oma11"))
                    If DW1 = DW3 Then
                        Ws.Cells(LineZ, 9 + DW4 - DW2) = oReader.Item("alz09")
                    Else
                        Ws.Cells(LineZ, 9 + 52 + DW4 - DW2) = oReader.Item("alz09")
                    End If
                End If
                LineZ += 1
            End While
        End If
        oReader.Close()
        ' 
        ' 加總
        Ws.Cells(LineZ, 1) = "Total_forecast_AR_USD"
        Ws.Cells(LineZ, 9) = "=SUM(I2:I" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9))
        oRng.AutoFill(Destination:=Ws.Range("I" & LineZ & ":DB" & LineZ), Type:=xlFillDefault)
        '第二頁
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        Ws.Name = "Overdue AR detail"
        AdjustExcelFormat1()
        oCommand.CommandText = "SELECT oma68,oma69,oma01,oma67,oma09,oma02,oma11,alz08,alz09,oag04 FROM OMA_FILE,Alz_file,oag_file  WHERE  alz09 > 0 and oma03 = alz01 AND alz00 = '2' AND alz02 = "
        oCommand.CommandText += C1 & " AND alz03 = " & C2 & " AND alz04 = oma01  and oag01 = oma32"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("oma68")
                Ws.Cells(LineZ, 2) = oReader.Item("oma01")
                Ws.Cells(LineZ, 3) = oReader.Item("oma67")
                Ws.Cells(LineZ, 4) = oReader.Item("oma09")
                Ws.Cells(LineZ, 5) = oReader.Item("oma11")
                Ws.Cells(LineZ, 6) = oReader.Item("alz09")
                Ws.Cells(LineZ, 8) = oReader.Item("oma69")
                Ws.Cells(LineZ, 9) = oReader.Item("oag04")
                Dim SS As Decimal = DateDiff(DateInterval.Day, oReader.Item("oma11"), DStartN)
                Ws.Cells(LineZ, 10) = SS
                Select Case SS
                    Case Is < 0
                        Ws.Cells(LineZ, 11) = oReader.Item("alz09")
                    Case 0 To 29
                        Ws.Cells(LineZ, 12) = oReader.Item("alz09")
                    Case 30 To 59
                        Ws.Cells(LineZ, 13) = oReader.Item("alz09")
                    Case 60 To 89
                        Ws.Cells(LineZ, 14) = oReader.Item("alz09")
                    Case 90 To 119
                        Ws.Cells(LineZ, 15) = oReader.Item("alz09")
                    Case Is > 120
                        Ws.Cells(LineZ, 16) = oReader.Item("alz09")
                End Select
                LineZ += 1
            End While
        End If
        oReader.Close()
        ' 加總
        Ws.Cells(LineZ, 5) = "Total AR amount_USD"
        Ws.Cells(LineZ, 6) = "=SUM(F2:F" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 11) = "=SUM(K2:K" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 12) = "=SUM(L2:L" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 13) = "=SUM(M2:M" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 14) = "=SUM(N2:N" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 15) = "=SUM(O2:O" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 16) = "=SUM(P2:P" & LineZ - 1 & ")"

        ' 第三頁
        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        Ws.Name = "Overdue AR overview"
        AdjustExcelFormat2()
        Ws.Cells(2, 2) = "='Overdue AR detail'!P" & LineZ
        Ws.Cells(3, 2) = "='Overdue AR detail'!O" & LineZ
        Ws.Cells(4, 2) = "='Overdue AR detail'!N" & LineZ
        Ws.Cells(5, 2) = "='Overdue AR detail'!M" & LineZ
        Ws.Cells(6, 2) = "='Overdue AR detail'!L" & LineZ
        Ws.Cells(7, 2) = "='Overdue AR detail'!K" & LineZ
        Ws.Cells(8, 2) = "=SUM(B2:B6)"

        Dim XShape As Microsoft.Office.Interop.Excel.Shape = Ws.Shapes.AddChart(xlLineMarkers, 330, 30, 550, 250)
        XShape.Chart.ChartType = xlLineMarkers
        XShape.Chart.SetSourceData(Ws.Range("A1", "B7"))
        XShape.Chart.ChartTitle.Text = "overdue amount by days"

        oCommand.CommandText = "SELECT oma69,sum(alz09) as t2 FROM OMA_FILE,Alz_file WHERE  alz09 > 0 and oma03 = alz01 AND alz00 = '2' AND alz02 = "
        oCommand.CommandText += C1 & " AND alz03 = " & C2 & " AND oma11 < to_date('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') AND alz04 = oma01  group by oma69"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineX, 1) = oReader.Item("oma69")
                Ws.Cells(LineX, 2) = oReader.Item("t2")
                LineX += 1
            End While
        End If
        oReader.Close()
        ' 加總
        Ws.Cells(LineX, 1) = "Total overdue AR_USD"
        Ws.Cells(LineX, 2) = "=SUM(B11:B" & LineX - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineX, 1), Ws.Cells(LineX, 2))
        oRng.Interior.Color = Color.Yellow
        ' 作圖
        Dim XShape2 As Microsoft.Office.Interop.Excel.Shape = Ws.Shapes.AddChart(xlColumnClustered, 330, 300, 700, 400)
        XShape2.Chart.ChartType = xlColumnClustered
        XShape2.Chart.SetSourceData(Ws.Range("A10", Ws.Cells(LineX - 1, 2)))
        XShape2.Chart.ChartTitle.Text = "overdue amount by customer"

        ' 第四頁   20160829
        Ws = xWorkBook.Sheets(4)
        Ws.Activate()
        Ws.Name = "Customer VS Overdue days"
        AdjustExcelFormat3()
        oCommand.CommandText = "SELECT oma69,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6 from ("
        oCommand.CommandText += "SELECT oma69,(case when oma11 > to_date('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') then sum(alz09) end) as t1,"
        oCommand.CommandText += "(case when to_date('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') - oma11 >= 0 and to_date('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') - oma11 < 30 then sum(alz09) end) as t2,"
        oCommand.CommandText += "(case when to_date('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') - oma11 >= 30 and to_date('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') - oma11 < 60 then sum(alz09) end) as t3,"
        oCommand.CommandText += "(case when to_date('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') - oma11 >= 60 and to_date('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') - oma11 < 90 then sum(alz09) end) as t4,"
        oCommand.CommandText += "(case when to_date('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') - oma11 >= 90 and to_date('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') - oma11 < 120 then sum(alz09) end) as t5,"
        oCommand.CommandText += "(case when to_date('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') - oma11 >= 120 then sum(alz09) end) as t6  FROM OMA_FILE,alz_file WHERE alz09 > 0 and oma03 = alz01 AND alz00 = '2' AND alz02 = "
        oCommand.CommandText += C1 & " AND alz03 = " & C2 & " AND alz04 = oma01 group by oma69,oma11,oma02 order by oma69 ) group by oma69"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineX, 1) = oReader.Item("oma69")
                Ws.Cells(LineX, 2) = oReader.Item("t1")
                Ws.Cells(LineX, 3) = oReader.Item("t2")
                Ws.Cells(LineX, 4) = oReader.Item("t3")
                Ws.Cells(LineX, 5) = oReader.Item("t4")
                Ws.Cells(LineX, 6) = oReader.Item("t5")
                Ws.Cells(LineX, 7) = oReader.Item("t6")
                LineX += 1
            End While
        End If
        oReader.Close()
        ' 加總
        Ws.Cells(LineX, 1) = "Total AR_USD"
        Ws.Cells(LineX, 2) = "=SUM(B2:B" & LineX - 1 & ")"
        Ws.Cells(LineX, 3) = "=SUM(C2:C" & LineX - 1 & ")"
        Ws.Cells(LineX, 4) = "=SUM(D2:D" & LineX - 1 & ")"
        Ws.Cells(LineX, 5) = "=SUM(E2:E" & LineX - 1 & ")"
        Ws.Cells(LineX, 6) = "=SUM(F2:F" & LineX - 1 & ")"
        Ws.Cells(LineX, 7) = "=SUM(G2:G" & LineX - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineX, 1), Ws.Cells(LineX, 7))
        oRng.Interior.Color = Color.Yellow
    End Sub
    Private Sub Setwk(ByVal eDate As Date)
        Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander2.Connection = oConnection
        oCommander2.CommandType = CommandType.Text
        oCommander2.CommandText = "SELECT azn02 FROM azn_file where azn01 = TO_DATE('" & eDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        DW3 = 0
        DW3 = oCommander2.ExecuteScalar()
        oCommander2.CommandText = "SELECT azn05 FROM azn_file where azn01 = TO_DATE('" & eDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        DW4 = 0
        DW4 = oCommander2.ExecuteScalar()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "H1")
        oRng.EntireColumn.ColumnWidth = 20
        Ws.Cells(1, 1) = "customer_nr"
        Ws.Cells(1, 2) = "payment terms"
        Ws.Cells(1, 3) = "bill number"
        Ws.Cells(1, 4) = "Invoice_nr"
        Ws.Cells(1, 5) = "invoice_date"
        Ws.Cells(1, 6) = "due_date"
        Ws.Cells(1, 7) = "amount"
        Ws.Cells(1, 8) = "due_balance"
        oRng = Ws.Range("A1", "B1")
        oRng.EntireColumn.NumberFormat = "@"
        oCommand.CommandText = "SELECT distinct azn02,azn05 FROM AZN_FILE WHERE AZN01 >= TO_DATE('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') order by azn02,azn05"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            CC = 0
            While oReader.Read()
                Ws.Cells(1, 9 + CC) = oReader.Item("azn02") & "W" & oReader.Item("azn05")
                CC += 1
            End While
        End If
        oRng = Ws.Range("G1", "DB1")
        oRng.EntireColumn.NumberFormatLocal = "_-""US$""* #,##0.00_ ;_-""US$""* -#,##0.00 ;_-""US$""* ""-""??_ ;_-@_ "
        oReader.Close()
        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "P1")
        oRng.EntireColumn.ColumnWidth = 22.22
        Ws.Cells(1, 1) = "customer_nr"
        Ws.Cells(1, 2) = "bill number"
        Ws.Cells(1, 3) = "Invoice_nr"
        Ws.Cells(1, 4) = "invoice_date"
        Ws.Cells(1, 5) = "due_date"
        Ws.Cells(1, 6) = "due_balance"
        oRng = Ws.Range("F1", "F1")
        oRng.EntireColumn.NumberFormatLocal = "_-""US$""* #,##0.00_ ;_-""US$""* -#,##0.00 ;_-""US$""* ""-""??_ ;_-@_ "
        Ws.Cells(1, 7) = "note"
        Ws.Cells(1, 8) = "customer"
        Ws.Cells(1, 9) = "payment terms_days"
        Ws.Cells(1, 10) = "how many overdue days"
        Ws.Cells(1, 11) = "not overdue"
        Ws.Cells(1, 12) = "overdue less then 30 days"
        Ws.Cells(1, 13) = "overdue over 30 days"
        oRng = Ws.Range("M1", "M1")
        oRng.Interior.Color = Color.LightGreen
        Ws.Cells(1, 14) = "overdue over 60 days"
        oRng = Ws.Range("N1", "N1")
        oRng.Interior.Color = Color.Yellow
        Ws.Cells(1, 15) = "overdue over 90 days"
        oRng = Ws.Range("O1", "O1")
        oRng.Interior.Color = Color.Orange
        Ws.Cells(1, 16) = "overdue over 120 days"
        oRng = Ws.Range("P1", "P1")
        oRng.Interior.Color = Color.Red
        oRng = Ws.Range("K1", "P1")
        oRng.EntireColumn.NumberFormatLocal = "_-""US$""* #,##0.00_ ;_-""US$""* -#,##0.00 ;_-""US$""* ""-""??_ ;_-@_ "
        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 37.33
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 19.11
        oRng.EntireColumn.NumberFormatLocal = "_-""US$""* #,##0.00_ ;_-""US$""* -#,##0.00 ;_-""US$""* ""-""??_ ;_-@_ "
        Ws.Cells(1, 1) = "overdue days"
        Ws.Cells(1, 2) = "sum amount"
        Ws.Cells(2, 1) = "overdue over 120 days"
        Ws.Cells(3, 1) = "overdue over 90 days"
        Ws.Cells(4, 1) = "overdue over 60 days"
        Ws.Cells(5, 1) = "overdue over 30 days"
        Ws.Cells(6, 1) = "overdue less than 30 days"
        Ws.Cells(7, 1) = "not overdue"
        Ws.Cells(8, 1) = "Total overdue AR_USD"
        oRng = Ws.Range("A8", "B8")
        oRng.Interior.Color = Color.Yellow
        Ws.Cells(10, 1) = "customer"
        Ws.Cells(10, 2) = "overdue amount"
        LineX = 11
    End Sub
    Private Sub AdjustExcelFormat3()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 37.33
        oRng = Ws.Range("B1", "G1")
        oRng.EntireColumn.ColumnWidth = 19.11
        oRng.EntireColumn.NumberFormatLocal = "_-""US$""* #,##0.00_ ;_-""US$""* -#,##0.00 ;_-""US$""* ""-""??_ ;_-@_ "
        'Ws.Cells(1, 1) = "overdue days"
        'Ws.Cells(1, 2) = "sum amount"
        Ws.Cells(1, 7) = "overdue over 120 days"
        Ws.Cells(1, 6) = "overdue over 90 days"
        Ws.Cells(1, 5) = "overdue over 60 days"
        Ws.Cells(1, 4) = "overdue over 30 days"
        Ws.Cells(1, 3) = "overdue less than 30 days"
        Ws.Cells(1, 2) = "not overdue"
        'Ws.Cells(7, 1) = "Total overdue AR_USD"
        'oRng = Ws.Range("A7", "B7")
        'oRng.Interior.Color = Color.Yellow
        'Ws.Cells(9, 1) = "customer"
        'Ws.Cells(9, 2) = "overdue amount"
        LineX = 2
    End Sub
End Class