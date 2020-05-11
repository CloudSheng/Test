Public Class Form98
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
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim pYear As Int16 = 0
    Dim pMonth As Int16 = 0
    Dim Start2 As Date
    Dim End2 As Date
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form98_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        tYear = Me.NumericUpDown1.Value
        tMonth = Me.NumericUpDown2.Value
        pYear = tYear
        pMonth = tMonth - 1
        If pMonth = 0 Then
            pYear = tYear - 1
            pMonth = 12
        End If
        Start2 = Convert.ToDateTime(tYear & "/" & tMonth & "/01")
        End2 = Start2.AddMonths(1).AddDays(-1)
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "BOM表物料标准成本与实际成本明细表"
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
        oCommand.CommandText = "select distinct bmb03,ima02,ima08,ima06,bmb10,nvl(tc_stc04,0) as t1,aa.ccc23 as t2,nvl(bb.ccc23,0) as t3 from bma_file "
        oCommand.CommandText += "left join bmb_file on bma01 = bmb01 left join ima_file on bmb03 = ima01 left join tc_stc_file on bmb03 = tc_stc01 "
        oCommand.CommandText += "left join ccc_file aa on bmb03 = aa.ccc01 and aa.ccc02 = " & tYear & " and aa.ccc03 = " & tMonth
        oCommand.CommandText += " left join ccc_file bb on bmb03 = bb.ccc01 and bb.ccc02 = " & pYear & " and bb.ccc03 = " & pMonth
        oCommand.CommandText += " where bmaacti = 'Y' and bma05 <= to_date('" & End2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ima08 in ('P','S') "
        If Not String.IsNullOrEmpty(Me.TextBox1.Text) Then
            oCommand.CommandText += " AND bmb03 like '" & Me.TextBox1.Text & "%'"
        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = tYear
                Ws.Cells(LineZ, 2) = tMonth
                Ws.Cells(LineZ, 3) = oReader.Item("bmb03")
                Ws.Cells(LineZ, 4) = oReader.Item("ima02")
                Ws.Cells(LineZ, 5) = oReader.Item("ima08")
                Ws.Cells(LineZ, 6) = oReader.Item("ima06")
                Ws.Cells(LineZ, 7) = oReader.Item("bmb10")
                Ws.Cells(LineZ, 8) = oReader.Item("t1")
                If IsDBNull(oReader.Item("t2")) Then
                    Ws.Cells(LineZ, 9) = oReader.Item("t3")
                Else
                    Ws.Cells(LineZ, 9) = oReader.Item("t2")
                End If
                LineZ += 1
            End While
        End If
        oRng = Ws.Range("A1", "I1")
        oRng.EntireColumn.AutoFit()
        oReader.Close()
    End Sub

    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Cells(1, 1) = "年度"
        Ws.Cells(1, 2) = "月份"
        Ws.Cells(1, 3) = "料件编号"
        Ws.Cells(1, 4) = "品名"
        Ws.Cells(1, 5) = "来源码"
        Ws.Cells(1, 6) = "分群码"
        Ws.Cells(1, 7) = "单位"
        Ws.Cells(1, 8) = "标准单位成本"
        Ws.Cells(1, 9) = "实际单位成本"
        oRng = Ws.Range("H1", "I1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.000000_ "
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        LineZ = 2
    End Sub
End Class