Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form68
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form68_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        SaveFileDialog1.FileName = "NG_ReworkOrder"
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
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        oCommand.CommandText = "select distinct img01,ima02,ima021,imd02,img10,img37,sfb82,(case when sfb82 is null then 0 else sum(sfa05-sfa06-sfa062) end) t1 from img_file "
        oCommand.CommandText += "left join ima_file on img01 = ima01 left join imd_file on img02 = imd01 left join sfa_file on img01 = sfa03 and sfa05 > (sfa06  + sfa062) "
        oCommand.CommandText += "left join sfb_file on sfa01 = sfb01 and sfb99 = 'Y' and sfb04 <> '8' and sfb87 ='Y' and substr(img02,0,5) = sfb82  where img02 in ('D356105','D356305','D356405','D356505') AND IMG10 > 0 "
        oCommand.CommandText += "group by img01,ima02,ima021,imd02,img10,img37,sfb82"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = oReader.Item(i)
                Next
                If Not IsDBNull(oReader.Item("sfb82")) Then
                    Ws.Cells(LineZ, 9) = oReader.Item("img10") - oReader.Item("t1")
                End If
                LineZ += 1
            End While
        End If
        oReader.Close()
        Ws.Columns.EntireColumn.AutoFit()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Cells(1, 1) = "料号"
        Ws.Cells(1, 2) = "品名"
        Ws.Cells(1, 3) = "规格"
        Ws.Cells(1, 4) = "仓库名称"
        Ws.Cells(1, 5) = "库存数量"
        Ws.Cells(1, 6) = "库存最后异动日期"
        Ws.Cells(1, 7) = "部门"
        Ws.Cells(1, 8) = "返工工单可用数量"
        Ws.Cells(1, 9) = "差异数"
        LineZ = 2
    End Sub
End Class