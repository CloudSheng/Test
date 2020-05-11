Public Class Form103
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
    Dim LineZ As Integer = 0
    Dim Time1 As DateTime
    Dim Time2 As DateTime
    Dim rvu01 As String = String.Empty
    Dim rvu02 As String = String.Empty
    Dim rvu04 As String = String.Empty
    Dim rvv36 As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")


    Private Sub Form103_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Time1 = Me.DateTimePicker1.Value
        Time2 = Me.DateTimePicker2.Value
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "验退仓退明细表"
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
        oCommand.CommandText = "select rvu00,rvu02,rvu01,rvu03,rvu04,rvu05,rvv02,rvv36,rvv31,rvv031,ima021,rvv35,rvv17,rvv38t,rvv39t,NULL,rvv26,azf03 from rvv_file "
        oCommand.CommandText += "left join rvu_file on rvv01 = rvu01 left join ima_file on rvv31 = ima01 left join azf_file on rvv26 = azf01 and azf02 = '2' where rvu00 = 3 "
        oCommand.CommandText += "AND rvu03 between to_date('" & Time1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') AND to_date('" & Time2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        If Not String.IsNullOrEmpty(rvu01) Then
            oCommand.CommandText += " AND rvu01 like '" & rvu01 & "%'"
        End If
        If Not String.IsNullOrEmpty(rvu02) Then
            oCommand.CommandText += " AND rvu02 like '" & rvu02 & "%'"
        End If
        If Not String.IsNullOrEmpty(rvu04) Then
            oCommand.CommandText += " AND rvu04 like '" & rvu04 & "%'"
        End If
        If Not String.IsNullOrEmpty(rvv36) Then
            oCommand.CommandText += " AND rvv36 like '" & rvv36 & "%'"
        End If
        oCommand.CommandText += " union all "
        oCommand.CommandText += "select rvu00,rvu02,rvu01,rvu03,rvu04,rvu05,rvv02,rvv36,rvv31,rvv031,ima021,rvv35,rvv17,rvv38t,rvv39t,qcu021,qcu04,qce03 from rvv_file "
        oCommand.CommandText += "left join rvu_file on rvv01 = rvu01 left join ima_file on rvv31 = ima01 left join qcu_file on rvu02 = qcu01 and rvv05 = qcu02 left join qce_file on qcu04 = qce01 where rvu00 = 2 "
        oCommand.CommandText += "AND rvu03 between to_date('" & Time1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') AND to_date('" & Time2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        If Not String.IsNullOrEmpty(rvu01) Then
            oCommand.CommandText += " AND rvu01 like '" & rvu01 & "%'"
        End If
        If Not String.IsNullOrEmpty(rvu02) Then
            oCommand.CommandText += " AND rvu02 like '" & rvu02 & "%'"
        End If
        If Not String.IsNullOrEmpty(rvu04) Then
            oCommand.CommandText += " AND rvu04 like '" & rvu04 & "%'"
        End If
        If Not String.IsNullOrEmpty(rvv36) Then
            oCommand.CommandText += " AND rvv36 like '" & rvv36 & "%'"
        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                If oReader.Item(0) = 2 Then
                    Ws.Cells(LineZ, 1) = "验退"
                Else
                    Ws.Cells(LineZ, 1) = "仓退"
                End If
                For i As Int16 = 1 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = oReader.Item(i)
                Next
                LineZ += 1
            End While
        End If
        oReader.Close()
        oRng = Ws.Range("A1", "R1")
        oRng.EntireColumn.AutoFit()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Cells(1, 1) = "异动类型"
        Ws.Cells(1, 2) = "收货单号"
        Ws.Cells(1, 3) = "退货单号"
        Ws.Cells(1, 4) = "退货日期"
        Ws.Cells(1, 5) = "厂商编号"
        Ws.Cells(1, 6) = "厂商简称"
        Ws.Cells(1, 7) = "项次"
        Ws.Cells(1, 8) = "采购单号"
        Ws.Cells(1, 9) = "料件编号"
        Ws.Cells(1, 10) = "品名"
        Ws.Cells(1, 11) = "规格"
        Ws.Cells(1, 12) = "单位"
        Ws.Cells(1, 13) = "退货数量"
        Ws.Cells(1, 14) = "含税单价"
        Ws.Cells(1, 15) = "含税金额(原幣)"
        Ws.Cells(1, 16) = "行序"
        Ws.Cells(1, 17) = "退货理由码"
        Ws.Cells(1, 18) = "退货理由"
        oRng = Ws.Range("N1", "O1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00"
        oRng = Ws.Range("I1", "I1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        oRng = Ws.Range("M1", "M1")
        oRng.EntireColumn.NumberFormatLocal = "0.00"
        LineZ = 2
    End Sub
End Class