Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants

Public Class Form115
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim TYear As String = String.Empty
    Dim TMonth As String = String.Empty
    Dim CYear As String = String.Empty
    Dim CMonth As String = String.Empty
    Dim g_oga03 As String = String.Empty
    Dim LineZ As Integer = 0
    Dim LineS1 As Int16 = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form115_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        Me.DateTimePicker1.Value = Today
        Me.DateTimePicker2.Value = Today
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
        g_oga03 = String.Empty
        If Not String.IsNullOrEmpty(TextBox1.Text) Then
            g_oga03 = TextBox1.Text
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
        SaveFileDialog1.FileName = "DAC_InvoiceReport"
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
        Ws.Activate()
        AdjustExcelFormat1()
        oCommand.CommandText = "select ofa02,ofa01,ofb04,ofb06,ima021,ofa23,ofb13 from ofa_file,ofb_file,ima_file where ofa01 = ofb01 and ofaconf = 'Y' and ofb04 = ima01 and ofa02 between to_date('"
        oCommand.CommandText += Me.DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += Me.DateTimePicker2.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ofa01 not like '%PI%'"
        If Not String.IsNullOrEmpty(g_oga03) Then
            oCommand.CommandText += " AND ofb04 like = '%" & g_oga03 & "%' "
        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = oReader.Item(i)
                Next
                LineZ += 1
            End While
        End If
        oReader.Close()

        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustExcelFormat2()
        oCommand.CommandText = "select DISTINCT ofb04,ofb06,ima021,ofa23 from ofa_file,ofb_file,ima_file where ofa01 = ofb01 and ofaconf = 'Y' and ofb04 = ima01 and ofa02 between to_date('"
        oCommand.CommandText += Me.DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += Me.DateTimePicker2.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ofa01 not like '%PI%'"
        If Not String.IsNullOrEmpty(g_oga03) Then
            oCommand.CommandText += " AND ofb04 like = '%" & g_oga03 & "%' "
        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("ofb04")
                Ws.Cells(LineZ, 2) = oReader.Item("ofb06")
                Ws.Cells(LineZ, 3) = oReader.Item("ima021")
                Ws.Cells(LineZ, 4) = oReader.Item("ofa23")
                oCommand2.CommandText = "select ofb13,ofa02 from ofa_file,ofb_file where ofa01 = ofb01 and ofb04 = '"
                oCommand2.CommandText += oReader.Item("ofb04") & "' and ofaconf = 'Y' and ofa02 between to_date('"
                oCommand2.CommandText += Me.DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                oCommand2.CommandText += Me.DateTimePicker2.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ofa01 not like '%PI%'  order by ofa02 desc"
                Dim LastPrice As Decimal = 0
                oReader2 = oCommand2.ExecuteReader()
                If oReader2.HasRows() Then
                    oReader2.Read()
                    LastPrice = oReader2.Item("ofb13")
                    Ws.Cells(LineZ, 5) = LastPrice
                    Ws.Cells(LineZ, 6) = oReader2.Item("ofa02")
                End If
                oReader2.Close()
                If LastPrice <> 0 Then
                    oCommand2.CommandText = "select ofb13,ofa02 from ofa_file,ofb_file where ofa01 = ofb01 and ofb04 = '"
                    oCommand2.CommandText += oReader.Item("ofb04") & "' and ofaconf = 'Y' and ofa02 between to_date('"
                    oCommand2.CommandText += Me.DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                    oCommand2.CommandText += Me.DateTimePicker2.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ofa01 not like '%PI%' and ofb13 <> " & LastPrice & " order by ofa02 desc"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        oReader2.Read()
                        Ws.Cells(LineZ, 7) = oReader2.Item("ofb13")
                        Ws.Cells(LineZ, 8) = oReader2.Item("ofa02")
                        Ws.Cells(LineZ, 9) = "=E" & LineZ & "-G" & LineZ
                    End If
                End If
                oReader2.Close()
                LineZ += 1
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "Invoice 明细"
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 29.43
        Ws.Cells(1, 1) = "出货日期"
        Ws.Cells(1, 2) = "Invoice No"
        Ws.Cells(1, 3) = "产品编号"
        Ws.Cells(1, 4) = "品名规格"
        Ws.Cells(1, 5) = "规格"
        Ws.Cells(1, 6) = "币别"
        Ws.Cells(1, 7) = "单价"
        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "Invoice 单价分析"
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 29.43
        Ws.Cells(1, 1) = "产品编号"
        Ws.Cells(1, 2) = "品名规格"
        Ws.Cells(1, 3) = "规格"
        Ws.Cells(1, 4) = "币别"
        Ws.Cells(1, 5) = "最后一次单价"
        Ws.Cells(1, 6) = "最后一次单价的出口日期"
        Ws.Cells(1, 7) = "前一次与最后一次的不同单价"
        Ws.Cells(1, 8) = "前一次与最后一次的不同单价的出口日期"
        Ws.Cells(1, 9) = "差异"
        LineZ = 2
    End Sub
End Class