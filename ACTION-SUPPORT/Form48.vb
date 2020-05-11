Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form48
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander3 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader3 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineX As Integer = 0
    Dim DetailX As Integer = 0
    Dim DStartN As Date
    Dim DstartE As Date
    Dim l_sfb01 As String = String.Empty
    Dim l_sfb05 As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form48_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        TextBox3.Text = Now.Year
        TextBox4.Text = Now.Month
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If String.IsNullOrEmpty(TextBox3.Text) Then
            MsgBox("请输入年份")
            Return
        End If
        If String.IsNullOrEmpty(TextBox4.Text) Then
            MsgBox("请输入月份")
            Return
        End If
        If TextBox4.Text > 12 Or TextBox4.Text < 1 Then
            MsgBox("月份有误")
            Return
        End If
        DStartN = Convert.ToDateTime(TextBox3.Text & "/" & TextBox4.Text & "/01")
        DstartE = DStartN.AddMonths(1).AddDays(-1)
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommander2.Connection = oConnection
                oCommander2.CommandType = CommandType.Text
                oCommander3.Connection = oConnection
                oCommander3.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        If Not String.IsNullOrEmpty(TextBox1.Text) Then
            l_sfb01 = TextBox1.Text
        End If
        If Not String.IsNullOrEmpty(TextBox2.Text) Then
            l_sfb05 = TextBox2.Text
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
        SaveFileDialog1.FileName = "Cost_Analysis_Report"
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
        oCommand.CommandText = "select sfv04,ima02,ima021 from ( select distinct sfv04 from sfu_file,sfv_file where sfu01 = sfv01 and sfu02 between to_date('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and sfupost = 'Y' "
        If Not String.IsNullOrEmpty(l_sfb01) Then
            oCommand.CommandText += "AND sfv11 = '" & l_sfb01 & "' "
        End If
        If Not String.IsNullOrEmpty(l_sfb05) Then
            oCommand.CommandText += "AND sfv04 = '" & l_sfb05 & "' "
        End If
        oCommand.CommandText += "order by sfv04 ),ima_file where sfv04 = ima01  order by sfv04"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineX, 1) = "产品编号"
                Ws.Cells(LineX, 2) = "品名"
                Ws.Cells(LineX, 3) = "规格"
                Ws.Cells(LineX + 1, 1) = oReader.Item("sfv04")
                Ws.Cells(LineX + 1, 2) = oReader.Item("ima02")
                Ws.Cells(LineX + 1, 3) = oReader.Item("ima021")
                Ws.Cells(LineX + 2, 1) = "工单编号"
                Ws.Cells(LineX + 2, 2) = "生产数量"
                Ws.Cells(LineX + 2, 3) = "完工入库数量"
                Ws.Cells(LineX + 2, 4) = "报废数量"
                Ws.Cells(LineX + 2, 5) = "实际工时"
                Ws.Cells(LineX + 2, 6) = "标准工时"
                Ws.Cells(LineX + 2, 7) = "料号"
                Ws.Cells(LineX + 2, 8) = "品名"
                Ws.Cells(LineX + 2, 9) = "规格"
                Ws.Cells(LineX + 2, 10) = "发料单位"
                Ws.Cells(LineX + 2, 11) = "应发数量"
                Ws.Cells(LineX + 2, 12) = "实发数量"
                ' 開始處理各料號工單
                LineX += 3
                oCommander2.CommandText = "select sum(sfv09) as t1,sum(sfvud07) as t2,sum((sfv09 + sfvud07) * ima58) as t3,sfv11 from sfu_file,sfv_file,ima_file where sfu01 = sfv01 and sfu02 between to_date('"
                oCommander2.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                oCommander2.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and sfupost = 'Y' and sfv04 = '"
                oCommander2.CommandText += oReader.Item("sfv04") & "' and sfv04 = ima01 group by sfv11"
                oReader2 = oCommander2.ExecuteReader()
                If oReader2.HasRows() Then
                    While oReader2.Read()
                        Ws.Cells(LineX, 1) = oReader2.Item("sfv11")
                        Ws.Cells(LineX, 2) = GetSfb08(oReader2.Item("sfv11"))
                        Ws.Cells(LineX, 3) = oReader2.Item("t1")
                        Ws.Cells(LineX, 4) = oReader2.Item("t2")
                        Ws.Cells(LineX, 5) = Getccj05(oReader2.Item("sfv11"))
                        Ws.Cells(LineX, 6) = oReader2.Item("t3")
                        ' 知道工單單身筆數
                        oCommander3.CommandText = "select sfa03,ima02,ima021,sfa12,sfa161 from sfa_file,ima_file where sfa01 = '"
                        oCommander3.CommandText += oReader2.Item("sfv11") & "' and sfa05 <> 0 and sfa03 = ima01"
                        oReader3 = oCommander3.ExecuteReader()
                        DetailX = 0
                        If oReader3.HasRows() Then
                            While oReader3.Read()
                                Ws.Cells(LineX + DetailX, 7) = oReader3.Item("sfa03")
                                Ws.Cells(LineX + DetailX, 8) = oReader3.Item("ima02")
                                Ws.Cells(LineX + DetailX, 9) = oReader3.Item("ima021")
                                Ws.Cells(LineX + DetailX, 10) = oReader3.Item("sfa12")
                                Ws.Cells(LineX + DetailX, 11) = Decimal.Round(oReader3.Item("sfa161") * (oReader2.Item("t1") + oReader2.Item("t2")), 4)
                                Ws.Cells(LineX + DetailX, 12) = Gettlf(oReader3.Item("sfa03"), oReader2.Item("sfv11"))
                                DetailX += 1
                            End While
                        End If
                        oReader3.Close()
                        LineX += DetailX
                    End While
                End If
                oReader2.Close()
            End While
        End If
        oReader.Close()
        ' 匯總
        LineX += 1
        Ws.Cells(LineX, 1) = "产品编号"
        Ws.Cells(LineX, 2) = "品名"
        Ws.Cells(LineX, 3) = "规格"
        Ws.Cells(LineX, 4) = "生产数量"
        Ws.Cells(LineX, 5) = "完工入库数量"
        Ws.Cells(LineX, 6) = "报废数量"
        Ws.Cells(LineX, 7) = "实际工时"
        Ws.Cells(LineX, 8) = "标准工时"
        Ws.Cells(LineX, 9) = "应发数量"
        Ws.Cells(LineX, 10) = "实发数量"
        Ws.Cells(LineX, 11) = "实际工时-标准工时"
        Ws.Cells(LineX, 12) = "应发数量-实发数量"
        LineX += 1
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineX, 1) = oReader.Item("sfv04")
                Ws.Cells(LineX, 2) = oReader.Item("ima02")
                Ws.Cells(LineX, 3) = oReader.Item("ima021")
                S1(oReader.Item("sfv04"))
                Ws.Cells(LineX, 4) = S2(oReader.Item("sfv04"))
                Ws.Cells(LineX, 7) = S3(oReader.Item("sfv04"))
                Ws.Cells(LineX, 9) = S4(oReader.Item("sfv04"))
                Ws.Cells(LineX, 10) = S5(oReader.Item("sfv04"))
                Ws.Cells(LineX, 11) = "=G" & LineX & "-H" & LineX
                Ws.Cells(LineX, 12) = "=I" & LineX & "-J" & LineX
                LineX += 1
            End While
        End If

    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "WorkOrder"
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 20.88
        oRng = Ws.Range("G1", "G1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        LineX = 1
    End Sub
    Private Function GetSfb08(ByVal sfb01 As String)
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "SELECT sfb08 FROM sfb_file WHERE SFB01 = '" & sfb01 & "'"
        Dim sfb08 As Decimal = oCommander99.ExecuteScalar()
        Return sfb08
    End Function
    Private Function Getccj05(ByVal sfb01 As String)
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "select nvl(sum(ccj05),0) from ccj_file where ccj01 between to_date('"
        oCommander99.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommander99.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ccj04 = '" & sfb01 & "'"
        Dim ccj05 As Decimal = oCommander99.ExecuteScalar()
        Return ccj05
    End Function
    Private Function Gettlf(ByVal tlf01 As String, ByVal tlf036 As String)
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "select nvl(sum(tlf10*tlf12),0) as t5 from tlf_file where tlf06 between to_date('"
        oCommander99.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommander99.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tlf01 = '"
        oCommander99.CommandText += tlf01 & "' and tlf036 = '" & tlf036 & "' and tlf13 like 'asfi5%'"
        Dim tlf10 As Decimal = oCommander99.ExecuteScalar()
        Return tlf10
    End Function
    Private Sub S1(ByVal sfv04 As String)
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "select sum(sfv09) as t1,sum(sfvud07) as t2,sum((sfv09 + sfvud07) * ima58) as t3 from sfu_file,sfv_file,ima_file where sfu01 = sfv01 and sfu02 between to_date('"
        oCommander99.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommander99.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and sfv04 = '" & sfv04 & "' and sfupost = 'Y' and sfv04 = ima01 "
        Dim oReader99 As Oracle.ManagedDataAccess.Client.OracleDataReader
        oReader99 = oCommander99.ExecuteReader()
        If oReader99.HasRows() Then
            oReader99.Read()
            Ws.Cells(LineX, 5) = oReader99.Item("t1")
            Ws.Cells(LineX, 6) = oReader99.Item("t2")
            Ws.Cells(LineX, 8) = oReader99.Item("t3")
        End If
        oReader99.Close()
    End Sub
    Private Function S2(ByVal sfv04 As String)
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "select sum(sfb08) from sfb_file where sfb01 in ( "
        oCommander99.CommandText += "select distinct sfv11 from sfu_file,sfv_file where sfu01 = sfv01 and sfu02 between to_date('"
        oCommander99.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommander99.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and sfv04 = '"
        oCommander99.CommandText += sfv04 & "' and sfupost = 'Y' )"
        Dim sfb08 As Decimal = oCommander99.ExecuteScalar()
        Return sfb08
    End Function
    Private Function S3(ByVal sfv04 As String)
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "select sum(ccj05) from ccj_file where ccj01 between to_date('"
        oCommander99.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommander99.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ccj04 in ( "
        oCommander99.CommandText += "select distinct sfv11 from sfu_file,sfv_file where sfu01 = sfv01 and sfu02 between to_date('"
        oCommander99.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommander99.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and sfv04 = '"
        oCommander99.CommandText += sfv04 & "' and sfupost = 'Y' )"
        Dim ccj05 As Decimal = oCommander99.ExecuteScalar()
        Return ccj05
    End Function
    Private Function S4(ByVal sfv04 As String)
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "select round(sum((sfv09 + sfvud07)* sfa161),4) as t1 from sfu_file,sfv_file,sfa_file where sfu01 = sfv01 and sfu02 between to_date('"
        oCommander99.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommander99.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and sfv04 = '"
        oCommander99.CommandText += sfv04 & "' and sfupost = 'Y' and sfv11 = sfa01"
        Dim sfa161 As Decimal = oCommander99.ExecuteScalar()
        Return sfa161
    End Function
    Private Function S5(ByVal sfv04 As String)
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "select nvl(round(sum(tlf10*tlf12),4),0) as t1 from tlf_file where tlf06 between to_date('"
        oCommander99.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommander99.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tlf13 like 'asfi%' and tlf036 in ("
        oCommander99.CommandText += "select distinct sfv11 from sfu_file,sfv_file where sfu01 = sfv01 and sfu02 between to_date('"
        oCommander99.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommander99.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and sfv04 = '"
        oCommander99.CommandText += sfv04 & "' and sfupost = 'Y' )"
        Dim tlf10 As Decimal = oCommander99.ExecuteScalar()
        Return tlf10
    End Function
End Class