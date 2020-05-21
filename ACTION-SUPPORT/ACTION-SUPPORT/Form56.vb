Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form56
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form56_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "DAC_LastOrderPrice_report_" & Today.ToString("yyyyMMdd")
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
        AdjustExcelFormat()
        oCommand.CommandText = "SELECT ima01,ima02,ima021 FROM ima_file WHERE IMA06 = '103' and imaacti = 'Y'"
        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            While oReader.Read
                Ws.Cells(LineZ, 1) = oReader.Item("ima01")
                Ws.Cells(LineZ, 2) = oReader.Item("ima02")
                Ws.Cells(LineZ, 3) = oReader.Item("ima021")
                GetLastPrice(oReader.Item("ima01"))
                LineZ += 1
            End While
        End If
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "G1")
        oRng.EntireColumn.ColumnWidth = 20
        Ws.Cells(1, 1) = "料号"
        Ws.Cells(1, 2) = "品名"
        Ws.Cells(1, 3) = "规格"
        Ws.Cells(1, 4) = "订单单号"
        Ws.Cells(1, 5) = "原币"
        Ws.Cells(1, 6) = "汇率"
        Ws.Cells(1, 7) = "本币"
        oRng = Ws.Range("A1", "B1")
        oRng.EntireColumn.NumberFormat = "@"
        LineZ = 2
    End Sub
    Private Sub GetLastPrice(ByVal ima01 As String)
        oCommander2.CommandText = "select * from ( SELECT oeb01,oeb13,oea24,(OEB13 * OEA24) as t1 FROM OEB_FILE,OEA_FILE WHERE OEB01 = OEA01 AND OEB04 = '" & ima01 & "' AND oea01 LIKE 'D2302%' AND OEACONF = 'Y' AND oea99 IS NOT NULL ORDER BY OEA02 DESC) WHERE rownum = 1"
        oReader2 = oCommander2.ExecuteReader
        If oReader2.HasRows Then
            oReader2.Read()
            Ws.Cells(LineZ, 4) = oReader2.Item("oeb01")
            Ws.Cells(LineZ, 5) = oReader2.Item("oeb13")
            Ws.Cells(LineZ, 6) = oReader2.Item("oea24")
            Ws.Cells(LineZ, 7) = oReader2.Item("t1")
        End If
        oReader2.Close()
    End Sub
End Class