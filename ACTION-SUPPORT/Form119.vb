Imports Microsoft.Office.Interop.Excel.XlFileFormat
Public Class Form119
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
    Dim LineS1 As Int16 = 0
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
                oCommand2.Connection = oConnection
                oCommand2.CommandType = CommandType.Text
                oCommand3.Connection = oConnection
                oCommand3.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        BackgroundWorker1.RunWorkerAsync()
        'ExportToExcel()
    End Sub

    Private Sub Form119_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "StandardPriceVSRealPrice"
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
        oCommand.CommandText = "select tc_stc01,ima02,ima021,ima25,ima44,ima44_fac,tc_stc04,ima06,ima08,tc_stcacti from tc_stc_file,ima_file where tc_stc01 = ima01  "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("tc_stc01")
                Ws.Cells(LineZ, 2) = oReader.Item("ima02")
                Ws.Cells(LineZ, 3) = oReader.Item("ima021")
                Ws.Cells(LineZ, 4) = oReader.Item("ima25")
                Ws.Cells(LineZ, 5) = oReader.Item("ima44")
                Ws.Cells(LineZ, 6) = oReader.Item("ima44_fac")
                Ws.Cells(LineZ, 7) = oReader.Item("tc_stc04")
                Ws.Cells(LineZ, 16) = oReader.Item("ima06")
                Ws.Cells(LineZ, 17) = oReader.Item("ima08")
                'Ws.Cells(LineZ, 18) = oReader.Item("tc_stcacti")
                GetPrice1(oReader.Item("tc_stc01"))
                GetPrice2(oReader.Item("tc_stc01"))
                LineZ += 1
            End While
        End If
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        'Ws.Name = "Invoice 明细"
        'Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 15
        Ws.Cells(1, 1) = "料号"
        Ws.Cells(1, 2) = "品名"
        Ws.Cells(1, 3) = "规格"
        Ws.Cells(1, 4) = "库存单位"
        Ws.Cells(1, 5) = "采购单位"
        Ws.Cells(1, 6) = "单位换算率"
        Ws.Cells(1, 7) = "标准单价（RMB)"
        Ws.Cells(1, 8) = "币种（核准单价）"
        Ws.Cells(1, 9) = "核准单价（原币)"
        Ws.Cells(1, 10) = "核准单价（RMB)"
        Ws.Cells(1, 11) = "币种（实际单价）"
        Ws.Cells(1, 12) = "实际单价（原币)"
        Ws.Cells(1, 13) = "实际单价（RMB)"
        Ws.Cells(1, 14) = "汇率(核准单价）"
        Ws.Cells(1, 15) = "汇率（实际单价）"
        Ws.Cells(1, 16) = "分群码"
        Ws.Cells(1, 17) = "来源码"
        'Ws.Cells(1, 18) = "资料有效码"
        Ws.Cells(1, 18) = "入库单位"
        Ws.Cells(1, 19) = "换算率（入库单）"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        LineZ = 2
    End Sub
    Private Sub GetPrice1(ByVal pmh01 As String)
        oCommand2.CommandText = "select pmj05,pmj07 from ( select * from pmj_file where pmj03 = '" & pmh01 & "' "
        oCommand2.CommandText += "  and pmj01 in (select ta_pmx12 from pmw_file,pmx_file where pmwacti = 'Y' and pmw01 = pmx01 ) order by pmj09 desc ) where rownum = 1"
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows Then
            While oReader2.Read()
                Ws.Cells(LineZ, 8) = oReader2.Item("pmj05")
                Ws.Cells(LineZ, 9) = oReader2.Item("pmj07")
                Dim l_azj04 As Decimal = 1
                If oReader2.Item("pmj05") = "RMB" Then
                    l_azj04 = 1
                Else
                    Dim l_s As String = String.Empty
                    If Today.Month < 10 Then
                        l_s = Today.Year & "0" & Today.Month
                    Else
                        l_s = Today.Year & Today.Month
                    End If
                    l_azj04 = GetRate(oReader2.Item("pmj05"), l_s)
                End If
                Ws.Cells(LineZ, 10) = oReader2.Item("pmj07") * l_azj04
                Ws.Cells(LineZ, 14) = l_azj04
            End While
        End If
        oReader2.Close()
    End Sub
    Private Sub GetPrice2(ByVal rvv31 As String)
        oCommand2.CommandText = "select pmm22,rvv38,(rvv38 * pmm42) as t1,pmm42,rvv35,rvv35_fac  from ( select * from rvv_file,rvu_file,pmm_file where rvv01 = rvu01 and rvv36 = pmm01 and rvv31 = '" & rvv31 & "' order by rvu03 desc ) where rownum = 1"
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                Ws.Cells(LineZ, 11) = oReader2.Item("pmm22")
                Ws.Cells(LineZ, 12) = oReader2.Item("rvv38")
                Ws.Cells(LineZ, 13) = oReader2.Item("t1")
                Ws.Cells(LineZ, 15) = oReader2.Item("pmm42")
                Ws.Cells(LineZ, 18) = oReader2.Item("rvv35")
                Ws.Cells(LineZ, 19) = oReader2.Item("rvv35_fac")
            End While
        End If
        oReader2.Close()
    End Sub
    Private Function GetRate(ByVal azj01 As String, ByVal azj02 As String)
        oCommand3.CommandText = "select nvl(azj04,1) from azj_file where azj01 = '" & azj01 & "' and azj02 = '" & azj02 & "'"
        Dim l_rate As Decimal = oCommand3.ExecuteScalar
        Return l_rate
    End Function
End Class