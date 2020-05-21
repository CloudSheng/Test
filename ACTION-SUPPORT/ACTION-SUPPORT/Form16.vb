Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form16
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim tYear As Decimal = 0
    Dim DStartN As Date
    Dim DstartE As Date
    Dim LineZ As Integer = 0
    Dim LineX As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            xExcel = New Microsoft.Office.Interop.Excel.Application
            xWorkBook = xExcel.Workbooks.Open(OpenFileDialog1.FileName)
            Ws = xWorkBook.Sheets(1)
        End If
        If IsDBNull(xWorkBook) Then
            Label1.Text = "读取失败"
            Me.GroupBox2.Enabled = False
        Else
            Label1.Text = "已读入"
            Me.GroupBox2.Enabled = True
        End If
    End Sub

    Private Sub Form16_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.GroupBox2.Enabled = False
        Me.TextBox1.Text = Now.Year
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
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
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Price_Report"
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
        LineZ = 26
        For i As Integer = 0 To 11
            'Count Part A
            oCommand.CommandText = "select nvl(round(sum(t1),4),0) as t1 from ( "
            oCommand.CommandText += "select pmn04,count(pmn04),sum(pmn31*pmm42) as t1 from pmm_file,pmn_file where pmm01 = pmn01 and pmm04 between to_date('"
            oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmm18 = 'Y' and pmm25 in (1,2,6) and pmm02 in ('REG','EXP') "
            oCommand.CommandText += "and pmn04 not in ( select distinct bmb03 from bmb_file ) group by pmn04 having count(pmn04) = 1 )"
            Dim PartA As Decimal = oCommand.ExecuteScalar()
            'Count Part B
            oCommand.CommandText = "select pmn04,count(pmn04) from pmm_file,pmn_file where pmm01 = pmn01 and pmm04 between to_date('"
            oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmm18 = 'Y' and pmm25 in (1,2,6) and pmm02 in ('REG','EXP') "
            oCommand.CommandText += "and pmn04 not in ( select distinct bmb03 from bmb_file ) group by pmn04  having count(pmn04) > 1"
            oReader = oCommand.ExecuteReader()
            Dim PartB As Decimal = 0
            If oReader.HasRows() Then
                While oReader.Read()
                    oCommander2.CommandText = "select nvl((pmm42 * pmn31),0) from pmm_file,pmn_file where pmm01 = pmn01 and pmm04 between to_date('"
                    oCommander2.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                    oCommander2.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmm18 = 'Y' and pmm25 in (1,2,6) and pmm02 in ('REG','EXP') "
                    oCommander2.CommandText += "and pmn04 = '" & oReader.Item("pmn04").ToString() & "' and pmm04 = ( "
                    oCommander2.CommandText += "select max(pmm04) from pmm_file,pmn_file where pmm01 = pmn01 and pmm04 between to_date('"
                    oCommander2.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                    oCommander2.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmm18 = 'Y' and pmm25 in (1,2,6) and pmm02 in ('REG','EXP') "
                    oCommander2.CommandText += "and pmn04 = '" & oReader.Item("pmn04").ToString() & "' )"
                    PartB += oCommander2.ExecuteScalar()
                End While
            End If
            oReader.Close()
            Dim TS As Decimal = PartA + PartB
            Ws.Cells(LineZ, 3 + i) = TS
            DStartN = DStartN.AddMonths(1)
            DstartE = DStartN.AddMonths(1).AddDays(-1)
        Next
        xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(1))
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustExcelFormat()
        DStartN = Convert.ToDateTime(tYear & "/01/01")
        DstartE = DStartN.AddMonths(1).AddDays(-1)
        For i As Integer = 0 To 11
            oCommand.CommandText = "select pmn04,ima02,ima021,pmm01,pmm04,pmn07,pmn20,pmc03,pmn24,pmm22,pmn31,(pmm31*pmn20) as t1,gen02,count(pmn04) "
            oCommand.CommandText += "from pmm_file, pmn_file, ima_file, pmc_file, gen_file "
            oCommand.CommandText += "where pmm01 = pmn01 And pmn04 = ima01 And pmm12 = gen01 and pmm09 = pmc01 and pmm04 between to_date('"
            oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmm18 = 'Y' and pmm25 in (1,2,6) and pmm02 in ('REG','EXP') "
            oCommand.CommandText += "and pmn04 not in ( select distinct bmb03 from bmb_file ) group by pmn04,ima02,ima021,pmm01,pmm04,pmn07,pmn20,pmc03,pmn24,pmm22,pmn31,(pmm31*pmn20),gen02 having count(pmn04) = 1 "
            oReader = oCommand.ExecuteReader()
            If oReader.HasRows() Then
                While oReader.Read()
                    Ws.Cells(LineX, 1) = "'" & oReader.Item("pmn04")
                    Ws.Cells(LineX, 2) = oReader.Item("ima02")
                    Ws.Cells(LineX, 3) = oReader.Item("ima021")
                    Ws.Cells(LineX, 4) = oReader.Item("pmm01")
                    Ws.Cells(LineX, 5) = oReader.Item("pmm04")
                    Ws.Cells(LineX, 6) = oReader.Item("pmn07")
                    Ws.Cells(LineX, 7) = oReader.Item("pmn20")
                    Ws.Cells(LineX, 8) = oReader.Item("pmc03")
                    Ws.Cells(LineX, 9) = oReader.Item("pmn24")
                    Ws.Cells(LineX, 10) = oReader.Item("pmm22")
                    Ws.Cells(LineX, 11) = oReader.Item("pmn31")
                    Ws.Cells(LineX, 12) = oReader.Item("t1")
                    Ws.Cells(LineX, 13) = oReader.Item("gen02")
                    LineX += 1
                End While
            End If
            oReader.Close()
            oCommand.CommandText = "select pmn04,count(pmn04) from pmm_file,pmn_file where pmm01 = pmn01 and pmm04 between to_date('"
            oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmm18 = 'Y' and pmm25 in (1,2,6) and pmm02 in ('REG','EXP') "
            oCommand.CommandText += "and pmn04 not in ( select distinct bmb03 from bmb_file ) group by pmn04  having count(pmn04) > 1"
            oReader = oCommand.ExecuteReader()
            If oReader.HasRows() Then
                While oReader.Read()
                    oCommander2.CommandText = "select pmn04,ima02,ima021,pmm01,pmm04,pmn07,pmn20,pmc03,pmn24,pmm22,pmn31,(pmm31*pmn20) as t1,gen02,count(pmn04) "
                    oCommander2.CommandText += "from pmm_file, pmn_file, ima_file, pmc_file, gen_file "
                    oCommander2.CommandText += "where pmm01 = pmn01 And pmn04 = ima01 And pmm12 = gen01 and pmm09 = pmc01 and pmm04 between to_date('"
                    oCommander2.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                    oCommander2.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmm18 = 'Y' and pmm25 in (1,2,6) and pmm02 in ('REG','EXP') "
                    oCommander2.CommandText += "and pmn04 = '" & oReader.Item("pmn04").ToString() & "' and pmm04 = ( "
                    oCommander2.CommandText += "select max(pmm04) from pmm_file,pmn_file where pmm01 = pmn01 and pmm04 between to_date('"
                    oCommander2.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                    oCommander2.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmm18 = 'Y' and pmm25 in (1,2,6) and pmm02 in ('REG','EXP') "
                    oCommander2.CommandText += "and pmn04 = '" & oReader.Item("pmn04").ToString() & "' ) group by pmn04,ima02,ima021,pmm01,pmm04,pmn07,pmn20,pmc03,pmn24,pmm22,pmn31,(pmm31*pmn20),gen02"
                    oReader2 = oCommander2.ExecuteReader()
                    If oReader2.HasRows() Then
                        While oReader2.Read()
                            Ws.Cells(LineX, 1) = "'" & oReader2.Item("pmn04")
                            Ws.Cells(LineX, 2) = oReader2.Item("ima02")
                            Ws.Cells(LineX, 3) = oReader2.Item("ima021")
                            Ws.Cells(LineX, 4) = oReader2.Item("pmm01")
                            Ws.Cells(LineX, 5) = oReader2.Item("pmm04")
                            Ws.Cells(LineX, 6) = oReader2.Item("pmn07")
                            Ws.Cells(LineX, 7) = oReader2.Item("pmn20")
                            Ws.Cells(LineX, 8) = oReader2.Item("pmc03")
                            Ws.Cells(LineX, 9) = oReader2.Item("pmn24")
                            Ws.Cells(LineX, 10) = oReader2.Item("pmm22")
                            Ws.Cells(LineX, 11) = oReader2.Item("pmn31")
                            Ws.Cells(LineX, 12) = oReader2.Item("t1")
                            Ws.Cells(LineX, 13) = oReader2.Item("gen02")
                            LineX += 1
                        End While
                    End If
                    oReader2.Close()
                End While
            End If
        Next
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "明細"
        Ws.Columns.EntireColumn.ColumnWidth = 20
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Cells(1, 1) = "料件编号"
        Ws.Cells(1, 2) = "品名"
        Ws.Cells(1, 3) = "规格"
        Ws.Cells(1, 4) = "采购单号"
        Ws.Cells(1, 5) = "采购日期"
        Ws.Cells(1, 6) = "采购单位"
        Ws.Cells(1, 7) = "采购量"
        Ws.Cells(1, 8) = "供应商简称"
        Ws.Cells(1, 9) = "请购单号"
        Ws.Cells(1, 10) = "币种"
        Ws.Cells(1, 11) = "原币单价"
        Ws.Cells(1, 12) = "原币金额"
        Ws.Cells(1, 13) = "采购员姓名"
        LineX = 2
    End Sub
End Class