Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form23
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim tYear As Decimal = 0
    Dim DStartN As Date
    Dim DstartE As Date
    Dim TotalRows As Integer = 0
    Dim LineX As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form23_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = Now.Year
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        Me.ProgressBar1.Value = 0
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
        tYear = Me.TextBox1.Text
        DStartN = Convert.ToDateTime(tYear & "/01/01")
        DstartE = DStartN.AddMonths(1).AddDays(-1)
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
        SaveFileDialog1.FileName = "Purchase_Annual_Report"
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
        oCommand.CommandText = "select count(*) from pmm_file,pmn_file,pmc_file,ima_file,pma_file "
        oCommand.CommandText += "where pmm01 = pmn01 and pmm18 = 'Y' and pmm09 = pmc01 and pmn04 = ima01 and pmm20 = pma01 "
        oCommand.CommandText += "and pmm02 <> 'SUB' and pmm04 between to_date('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += DStartN.AddYears(1).AddDays(-1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') order by pmm04,pmm01"
        TotalRows = oCommand.ExecuteScalar()
        oCommand.CommandText = "select count(*) from pmm_file,pmn_file,pmc_file,ima_file,pma_file "
        oCommand.CommandText += "where pmm01 = pmn01 and pmm18 = 'Y' and pmm09 = pmc01 and pmn04 = ima01 and pmm20 = pma01 "
        oCommand.CommandText += "and pmm02 = 'SUB' and pmm04 between to_date('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += DStartN.AddYears(1).AddDays(-1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') order by pmm04,pmm01"
        TotalRows += oCommand.ExecuteScalar()
        If TotalRows <> 0 Then
            Me.ProgressBar1.Maximum = TotalRows
            xExcel = New Microsoft.Office.Interop.Excel.Application
            xWorkBook = xExcel.Workbooks.Add()
            For i As Integer = 1 To 12
                If i < 4 Then
                    Ws = xWorkBook.Sheets(i)
                Else
                    Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                End If
                Ws.Activate()
                AdjustExcelFormat(i)
                oCommand.CommandText = "select pmm04,pmm01,pmn04,pmc03,ima02,pmn041,pmn07,pmn20,pmn31,pmm22,pmm43,pma02,pmn33 from pmm_file,pmn_file,pmc_file,ima_file,pma_file "
                oCommand.CommandText += "where pmm01 = pmn01 and pmm18 = 'Y' and pmm09 = pmc01 and pmn04 = ima01 and pmm20 = pma01 "
                oCommand.CommandText += "and pmm02 <> 'SUB' and pmm04 between to_date('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') order by pmm04,pmm01"
                oReader = oCommand.ExecuteReader()
                If oReader.HasRows() Then
                    While oReader.Read()
                        Ws.Cells(LineX, 1) = oReader.Item("pmm04")
                        Ws.Cells(LineX, 2) = oReader.Item("pmm01")
                        Ws.Cells(LineX, 3) = "'" & oReader.Item("pmn04")
                        Ws.Cells(LineX, 4) = oReader.Item("pmc03")
                        Ws.Cells(LineX, 5) = oReader.Item("ima02")
                        Ws.Cells(LineX, 6) = oReader.Item("pmn041")
                        Ws.Cells(LineX, 7) = oReader.Item("pmn07")
                        Ws.Cells(LineX, 8) = oReader.Item("pmn20")
                        Ws.Cells(LineX, 9) = oReader.Item("pmn31")
                        Ws.Cells(LineX, 10) = oReader.Item("pmm22")
                        Ws.Cells(LineX, 11) = oReader.Item("pmm43") & "%"
                        Ws.Cells(LineX, 12) = oReader.Item("pma02")
                        Ws.Cells(LineX, 13) = oReader.Item("pmn33")
                        LineX += 1
                        Me.ProgressBar1.Value += 1
                    End While
                End If
                oReader.Close()
                DStartN = DStartN.AddMonths(1)
                DstartE = DStartN.AddMonths(1).AddDays(-1)
            Next
            DStartN = Convert.ToDateTime(tYear & "/01/01")
            oCommand.CommandText = "select pmm04,pmm01,pmn04,pmc03,ima02,pmn041,pmn07,pmn20,pmn31,pmm22,pmm43,pma02,pmn33 from pmm_file,pmn_file,pmc_file,ima_file,pma_file "
            oCommand.CommandText += "where pmm01 = pmn01 and pmm18 = 'Y' and pmm09 = pmc01 and pmn04 = ima01 and pmm20 = pma01 "
            oCommand.CommandText += "and pmm02 = 'SUB' and pmm04 between to_date('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            oCommand.CommandText += DStartN.AddYears(1).AddDays(-1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') order by pmm04,pmm01"
            oReader = oCommand.ExecuteReader()
            If oReader.HasRows() Then
                Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                Ws.Activate()
                AdjustExcelFormat("委外")
                While oReader.Read()
                    Ws.Cells(LineX, 1) = oReader.Item("pmm04")
                    Ws.Cells(LineX, 2) = oReader.Item("pmm01")
                    Ws.Cells(LineX, 3) = "'" & oReader.Item("pmn04")
                    Ws.Cells(LineX, 4) = oReader.Item("pmc03")
                    Ws.Cells(LineX, 5) = oReader.Item("ima02")
                    Ws.Cells(LineX, 6) = oReader.Item("pmn041")
                    Ws.Cells(LineX, 7) = oReader.Item("pmn07")
                    Ws.Cells(LineX, 8) = oReader.Item("pmn20")
                    Ws.Cells(LineX, 9) = oReader.Item("pmn31")
                    Ws.Cells(LineX, 10) = oReader.Item("pmm22")
                    Ws.Cells(LineX, 11) = oReader.Item("pmm43") & "%"
                    Ws.Cells(LineX, 12) = oReader.Item("pma02")
                    Ws.Cells(LineX, 13) = oReader.Item("pmn33")
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
    Private Sub AdjustExcelFormat(ByVal Month1 As String)
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = Month1
        Ws.Columns.EntireColumn.ColumnWidth = 20
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "M1")
        oRng.Merge()
        oRng = Ws.Range("A2", "M2")
        oRng.Merge()
        Ws.Cells(1, 1) = "东莞艾可迅复合材料有限公司Dongguan Action Composites LTD  Co."
        Ws.Cells(2, 1) = "采购订单汇总表Purchasing Order List"
        Ws.Cells(3, 1) = "采购日期"
        Ws.Cells(3, 2) = "采购单号"
        Ws.Cells(3, 3) = "采购料号"
        Ws.Cells(3, 4) = "供应商名称"
        Ws.Cells(3, 5) = "品名"
        Ws.Cells(3, 6) = "规格"
        Ws.Cells(3, 7) = "单位"
        Ws.Cells(3, 8) = "数量"
        Ws.Cells(3, 9) = "未税单价"
        Ws.Cells(3, 10) = "currency"
        Ws.Cells(3, 11) = "税率"
        Ws.Cells(3, 12) = "结算方式"
        Ws.Cells(3, 13) = "交期"
        LineX = 4
    End Sub
End Class