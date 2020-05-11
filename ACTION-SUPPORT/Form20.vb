Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form20
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim tTimes As Int16 = 0
    Dim DStartN As Date
    Dim DstartE As Date
    Dim LineZ As Integer = 0
    Dim LineX As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form20_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = 0
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        tTimes = TextBox1.Text
        If tTimes < 0 Then
            MsgBox("次数有误")
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
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "PriceQuery_Report"
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
        LineZ = 2
        oCommand.CommandText = "select pmx08,pmx081,pmx082,count(pmx08) as t1 from pmw_file,pmx_file where pmw01 = pmx01 and ta_pmw04 = 'Y' and pmx08 <> 'MISC'  "
        oCommand.CommandText += "group by pmx08,pmx081,pmx082 having count(pmx08) > " & tTimes
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            xExcel = New Microsoft.Office.Interop.Excel.Application
            xWorkBook = xExcel.Workbooks.Add()
            Ws = xWorkBook.Sheets(1)
            AdjustExcelFormat()
            While oReader.Read()
                Ws.Cells(LineZ, 1) = "'" & oReader.Item("pmx08")
                Ws.Cells(LineZ, 2) = oReader.Item("pmx081")
                Ws.Cells(LineZ, 3) = oReader.Item("pmx082")
                Ws.Cells(LineZ, 4) = oReader.Item("t1")
                If oReader.Item("t1") > 0 Then
                    GetUnitPrice(oReader.Item("pmx08"), oReader.Item("t1"))
                End If
                LineZ += 1
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "询价"
        Ws.Columns.EntireColumn.ColumnWidth = 20
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Cells(1, 1) = "料件编号"
        Ws.Cells(1, 2) = "品名"
        Ws.Cells(1, 3) = "规格"
        Ws.Cells(1, 4) = "询价次数"
        Ws.Cells(1, 5) = "询价日期"
        Ws.Cells(1, 6) = "供应商"
        Ws.Cells(1, 7) = "税率"
        'oRng = Ws.Range("G1", "G1")
        'oRng.EntireColumn.NumberFormatLocal = "%"
        Ws.Cells(1, 8) = "币别"
        Ws.Cells(1, 9) = "询价单位"
        Ws.Cells(1, 10) = "询价金额(第一次)"
        Ws.Cells(1, 11) = "询价日期"
        Ws.Cells(1, 12) = "供应商"
        Ws.Cells(1, 13) = "税率"
        'oRng = Ws.Range("L1", "L1")
        'oRng.EntireColumn.NumberFormatLocal = "%"
        Ws.Cells(1, 14) = "币别"
        Ws.Cells(1, 15) = "询价单位"
        Ws.Cells(1, 16) = "询价金额(第二次)"
        Ws.Cells(1, 17) = "询价日期"
        Ws.Cells(1, 18) = "供应商"
        Ws.Cells(1, 19) = "税率"
        'oRng = Ws.Range("Q1", "Q1")
        'oRng.EntireColumn.NumberFormatLocal = "%"
        Ws.Cells(1, 20) = "币别"
        Ws.Cells(1, 21) = "询价单位"
        Ws.Cells(1, 22) = "询价金额(第三次)"
        Ws.Cells(1, 23) = "询价日期"
        Ws.Cells(1, 24) = "供应商"
        Ws.Cells(1, 25) = "税率"
        'oRng = Ws.Range("V1", "V1")
        'oRng.EntireColumn.NumberFormatLocal = "%"
        Ws.Cells(1, 26) = "币别"
        Ws.Cells(1, 27) = "询价单位"
        Ws.Cells(1, 28) = "询价金额(第四次)"
    End Sub
    Private Sub GetUnitPrice(ByVal ima01 As String, ByVal sTimes As Int16)
        Dim iTimes As Int16 = 0
        oCommander2.CommandText = "select * from pmw_file,pmx_file where pmw01 = pmx01 and ta_pmw04 = 'Y'  "
        oCommander2.CommandText += "and pmx08 = '" & ima01 & "' ORDER BY Pmx04,PMx01"
        oReader2 = oCommander2.ExecuteReader()
        If sTimes > 4 Then
            sTimes = 4
        End If
        If oReader2.HasRows() Then
            While iTimes < sTimes And oReader2.Read()
                'Ws.Cells(LineZ, iTimes + 5) = oReader2.Item("pmj05") & " " & oReader2.Item("pmj07")
                Ws.Cells(LineZ, iTimes * 6 + 5) = oReader2.Item("pmx04")
                Ws.Cells(LineZ, iTimes * 6 + 6) = oReader2.Item("ta_pmx02")
                Ws.Cells(LineZ, iTimes * 6 + 7) = oReader2.Item("ta_pmx06")
                Ws.Cells(LineZ, iTimes * 6 + 8) = oReader2.Item("ta_pmx03")
                Ws.Cells(LineZ, iTimes * 6 + 9) = oReader2.Item("pmx09")
                Ws.Cells(LineZ, iTimes * 6 + 10) = oReader2.Item("pmx06")
                iTimes += 1
            End While
        End If
        oReader2.Close()
    End Sub
End Class