Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form34
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim DStartN As Date
    Dim DstartE As Date
    Dim DN As Date
    Dim DE As Date
    Dim TYear As String = String.Empty
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form34_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        TextBox1.Text = Today.Year
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
        DStartN = Me.DateTimePicker1.Value
        DstartE = Me.DateTimePicker2.Value
        TYear = TextBox1.Text
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "DAC_SCRAP_Report"
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
        'Ws.Name = "Week"
        AdjustExcelFormat()

        oCommand.CommandText = "select tlf905,tlf06,gen02,gem02,tlf01,ima02,ima021,imd02,tlf11,tlf10,azf03,tlf17 from tlf_file "
        oCommand.CommandText += "left join gen_file on tlf09 = gen01 left join gem_file on tlf19 = gem01 left join ima_file on tlf01 = ima01 "
        oCommand.CommandText += "left join imd_file on tlf902 = imd01 left join azf_file on tlf14 = azf01 and azf02= '2' "
        oCommand.CommandText += "where tlf06 between to_date('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and (tlf13 = 'aimt303' or tlf13 = 'aimt313') "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("tlf905")
                Ws.Cells(LineZ, 2) = oReader.Item("tlf06")
                Ws.Cells(LineZ, 3) = oReader.Item("gen02")
                Ws.Cells(LineZ, 4) = oReader.Item("gem02")
                Ws.Cells(LineZ, 5) = oReader.Item("tlf01")
                Ws.Cells(LineZ, 6) = oReader.Item("ima02")
                Ws.Cells(LineZ, 7) = oReader.Item("ima021")
                Ws.Cells(LineZ, 8) = oReader.Item("imd02")
                Ws.Cells(LineZ, 9) = oReader.Item("tlf11")
                Ws.Cells(LineZ, 10) = oReader.Item("tlf10")
                Ws.Cells(LineZ, 11) = oReader.Item("azf03")
                Ws.Cells(LineZ, 12) = oReader.Item("tlf17")
                LineZ += 1
            End While
        End If
        oReader.Close()
        LineZ += 1

        AdjustExcelFormat1()
        DN = Convert.ToDateTime(TYear & "/01/01")
        DE = DStartN.AddMonths(1).AddDays(-1)
        For i As Integer = 1 To 12 Step 1
            Ws.Cells(LineZ + 1, 1 + i) = GetScrapCost(TYear, i)
            DN = DN.AddMonths(1)
            DE = DN.AddMonths(1).AddDays(-1)
        Next
        Ws.Cells(LineZ + 1, 14) = "=SUM(B" & LineZ + 1 & ":M" & LineZ + 1

        ' 第二頁
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustExcelFormat2()
        oCommand.CommandText = "SELECT SFU01,SFU02,GEN02,gem02,sfv11,sfv04,ima02,ima021,sfv08,sfvud07 FROM SFU_FILE JOIN SFV_FILE ON SFU01 = SFV01 AND SFVUD07 > 0 "
        oCommand.CommandText += "LEFT JOIN GEN_FILE ON SFU16 = GEN01 LEFT JOIN GEM_FILE ON sfu04 = gem01 left join ima_file on sfv04 = ima01 where SFU02 BETWEEN to_date('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and sfupost = 'Y' "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("sfu01")
                Ws.Cells(LineZ, 2) = oReader.Item("sfu02")
                Ws.Cells(LineZ, 3) = oReader.Item("gen02")
                Ws.Cells(LineZ, 4) = oReader.Item("gem02")
                Ws.Cells(LineZ, 5) = oReader.Item("sfv11")
                Ws.Cells(LineZ, 6) = oReader.Item("sfv04")
                Ws.Cells(LineZ, 7) = oReader.Item("ima02")
                Ws.Cells(LineZ, 8) = oReader.Item("ima021")
                Ws.Cells(LineZ, 9) = oReader.Item("sfv08")
                Ws.Cells(LineZ, 10) = oReader.Item("sfvud07")
                LineZ += 1
            End While
        End If
        oReader.Close()
        LineZ += 1
        AdjustExcelFormat1()
        DN = Convert.ToDateTime(TYear & "/01/01")
        DE = DStartN.AddMonths(1).AddDays(-1)
        For i As Integer = 1 To 12 Step 1
            Ws.Cells(LineZ + 1, 1 + i) = GetWIPScrapCost(TYear, i)
            DN = DN.AddMonths(1)
            DE = DN.AddMonths(1).AddDays(-1)
        Next
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "仓库杂项报废报表"
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 30
        oRng = Ws.Range("A1", "N1")
        oRng.Merge()
        Ws.Cells(1, 1) = "仓库杂项报废报表"
        Ws.Cells(2, 1) = "单据编号"
        Ws.Cells(2, 2) = "扣账日期"
        Ws.Cells(2, 3) = "申请人"
        Ws.Cells(2, 4) = "部门名称"
        oRng = Ws.Range("E2", "E2")
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(2, 5) = "料件编号"
        Ws.Cells(2, 6) = "品名"
        Ws.Cells(2, 7) = "规格"
        Ws.Cells(2, 8) = "仓库名称"
        Ws.Cells(2, 9) = "单位"
        Ws.Cells(2, 10) = "申请数量"
        Ws.Cells(2, 11) = "理由码说明"
        Ws.Cells(2, 12) = "备注"
        LineZ = 3
    End Sub
    Private Sub AdjustExcelFormat1()
        Dim prefix As String = String.Empty
        Ws.Cells(LineZ, 1) = "月份"
        For i = 1 To 12 Step 1
            If i < 10 Then
                prefix = "0" & i
            Else
                prefix = i
            End If
            Ws.Cells(LineZ, 1 + i) = TYear & "/" & prefix
        Next
        Ws.Cells(LineZ, 14) = "金额合计"
        Ws.Cells(LineZ + 1, 1) = "金额"
    End Sub
    Private Function GetScrapCost(ByVal iY As Integer, iM As Integer)
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "select nvl(sum(tlf10*tlf12*(stb07+stb08+stb09)),0) from tlf_file left join stb_file on tlf01 = stb01 and stb02 = "
        oCommander99.CommandText += iY & " and stb03 = " & iM & " where tlf06 between to_date('"
        oCommander99.CommandText += DN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommander99.CommandText += DE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and (tlf13 = 'aimt303' or tlf13 = 'aimt313')"
        Dim SC As Decimal = oCommander99.ExecuteScalar()
        Return SC
    End Function
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "工单不良品数报表"
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 30
        oRng = Ws.Range("A1", "N1")
        oRng.Merge()
        Ws.Cells(1, 1) = "工单不良品数报表"
        Ws.Cells(2, 1) = "入库单号"
        Ws.Cells(2, 2) = "入库日期"
        Ws.Cells(2, 3) = "申请人"
        Ws.Cells(2, 4) = "部门名称"
        Ws.Cells(2, 5) = "工单单号"
        oRng = Ws.Range("F2", "F2")
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(2, 6) = "料件编号"
        Ws.Cells(2, 7) = "品名"
        Ws.Cells(2, 8) = "规格"
        Ws.Cells(2, 9) = "库存单位"
        Ws.Cells(2, 10) = "报废数量"
        LineZ = 3
    End Sub
    Private Function GetWIPScrapCost(ByVal iY As Integer, iM As Integer)
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "SELECT nvl(sum(sfvud07*(stb07 + stb08 + stb09)),0) FROM SFU_FILE JOIN SFV_FILE ON SFU01 = SFV01 AND SFVUD07 > 0 left join stb_file on sfv04 = stb01 and stb02 = "
        oCommander99.CommandText += iY & " and stb03 = " & iM & " where SFU02 BETWEEN to_date('"
        oCommander99.CommandText += DN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommander99.CommandText += DE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and sfupost = 'Y'"
        Dim SC As Decimal = oCommander99.ExecuteScalar()
        Return SC
    End Function
End Class