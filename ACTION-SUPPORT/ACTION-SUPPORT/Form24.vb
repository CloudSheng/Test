Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form24
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim l_ima01 As String = String.Empty
    Dim DStartN As Date
    Dim DstartE As Date
    Dim TotalRows As Integer = 0
    Dim LineX As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form24_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = String.Empty
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
                oCommander2.Connection = oConnection
                oCommander2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        l_ima01 = TextBox1.Text
        DStartN = DateTimePicker1.Value
        DstartE = DateTimePicker2.Value
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If TotalRows > 0 Then
            SaveExcel()
        End If
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "QueryPrice_Report"
        SaveFileDialog1.DefaultExt = ".xls"
        Dim SON As DialogResult = SaveFileDialog1.ShowDialog()
        If SON = DialogResult.OK Then
            Dim SFN As String = SaveFileDialog1.FileName
            Ws.SaveAs(SFN, XlFileFormat.xlExcel12)
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
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub ExportToExcel()
        If String.IsNullOrEmpty(l_ima01) Then
            l_ima01 = " 1=1 "
        Else
            If Strings.InStr(l_ima01, "*") > 0 Then
                l_ima01 = Strings.Replace(l_ima01, "*", "%")
                l_ima01 = " pmx08 LIKE '" & l_ima01 & "' "
            Else
                l_ima01 = " pmx08 = '" & l_ima01 & "' "
            End If
        End If
        oCommand.CommandText = "select count(*) from pmw_file,pmx_file where pmw01 = pmx01 and pmw06 between to_date('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ta_pmw04 = 'Y' "
        oCommand.CommandText += "AND " & l_ima01 & "order by pmw06"
        TotalRows = oCommand.ExecuteScalar()
        If TotalRows <> 0 Then
            Me.ProgressBar1.Maximum = TotalRows
            xExcel = New Microsoft.Office.Interop.Excel.Application
            xWorkBook = xExcel.Workbooks.Add()
            Ws = xWorkBook.Sheets(1)
            Ws.Activate()
            AdjustExcelFormat()
            oCommand.CommandText = "select pmw01,pmw06,pmx08,pmx081,pmx082,pmx09,ta_pmx09,ta_pmx02 from pmw_file,pmx_file where pmw01 = pmx01 and pmw06 between to_date('"
            oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ta_pmw04 = 'Y' "
            oCommand.CommandText += "AND " & l_ima01 & "order by pmw06"
            oReader = oCommand.ExecuteReader()
            If oReader.HasRows() Then
                While oReader.Read()
                    Ws.Cells(LineX, 1) = oReader.Item("pmw01")
                    Ws.Cells(LineX, 2) = oReader.Item("pmw06")
                    Ws.Cells(LineX, 3) = "'" & oReader.Item("pmx08")
                    Ws.Cells(LineX, 4) = oReader.Item("pmx081")
                    Ws.Cells(LineX, 5) = oReader.Item("pmx082")
                    Ws.Cells(LineX, 6) = oReader.Item("pmx09")
                    Ws.Cells(LineX, 7) = oReader.Item("ta_pmx09")
                    Ws.Cells(LineX, 8) = oReader.Item("ta_pmx02")
                    LineX += 1
                    Me.ProgressBar1.Value += 1
                End While
            End If
            oReader.Close()
        Else
            MsgBox("报表无资料")
            Return
        End If
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "Query_Price"
        Ws.Columns.EntireColumn.ColumnWidth = 20
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Cells(1, 1) = "Query NO."
        Ws.Cells(1, 2) = "Query Date"
        Ws.Cells(1, 3) = "Item NO."
        Ws.Cells(1, 4) = "Product Name"
        Ws.Cells(1, 5) = "Specfication"
        Ws.Cells(1, 6) = "Unit"
        Ws.Cells(1, 7) = "Purchasing Price (RMB)"
        Ws.Cells(1, 8) = "Supplier"
        LineX = 2
    End Sub
End Class