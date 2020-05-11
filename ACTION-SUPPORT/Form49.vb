Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form49
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim DStartN As Date
    Dim DstartE As Date
    Dim TYear As String = String.Empty
    Dim TMonth As String = String.Empty
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

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
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        TYear = Strings.Left(TextBox1.Text, 4)
        TMonth = Strings.Right(TextBox1.Text, 2)
        DStartN = Convert.ToDateTime(TYear & "/" & TMonth & "/01")
        DstartE = DStartN.AddMonths(1).AddDays(-1)
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Form49_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If Now.Month < 10 Then
            TextBox1.Text = Now.Year & "0" & Now.Month
        Else
            TextBox1.Text = Now.Year & Now.Month
        End If
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "aimt312Check_Report"
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
        oCommand.CommandText = "select inb04,ima02,ima021,inb05,sum(inb09) as t1 from ina_file,inb_file,ima_file where ina01 = inb01 and ina00 = 4 and ina02 between to_date('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and inapost = 'Y' and inb04 =  ima01 group by inb04,ima02,ima021,inb05 order by inb04"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("inb04")
                Ws.Cells(LineZ, 2) = oReader.Item("ima02")
                Ws.Cells(LineZ, 3) = oReader.Item("ima021")
                Ws.Cells(LineZ, 4) = oReader.Item("inb05")
                Ws.Cells(LineZ, 5) = oReader.Item("t1")
                Ws.Cells(LineZ, 6) = Getimg(oReader.Item("inb04"), oReader.Item("inb05"))
                LineZ += 1
            End While
        End If
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "Data"
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 30
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(1, 1) = "料号"
        Ws.Cells(1, 2) = "品名"
        Ws.Cells(1, 3) = "规格"
        Ws.Cells(1, 4) = "仓库"
        Ws.Cells(1, 5) = "杂收总量"
        Ws.Cells(1, 6) = "库存量"
        Ws.Cells(1, 7) = "不含D14开头仓库"
        LineZ = 2
    End Sub
    Private Function Getimg(ByVal img01 As String, ByVal img02 As String)
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "select nvl(sum(img10),0) from img_file where img01 = '" & img01 & "' and img02 not like 'D14%'  and img10 <> 0 and img02 = '"
        oCommander99.CommandText += img02 & "'"
        Dim MF As Decimal = oCommander99.ExecuteScalar()
        Return MF
    End Function
End Class