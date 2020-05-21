Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form89
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
    Dim SQ1 As Decimal = 0  '請購量
    Dim SQ2 As Decimal = 0  '採購量
    Dim SQ3 As Decimal = 0  '工單備料量
    Dim SQ4 As Decimal = 0  '委外在制
    Dim SQ5 As Decimal = 0  'IQC
    Dim SQ6 As Decimal = 0  '可用量(庫存量)
    Dim SQ7 As Decimal = 0  '預計可用量
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
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
        Me.ProgressBar1.Value = 0
        oCommand.CommandText = "select count(distinct bmb03) from bmb_file,bma_file,ima_file where bmb01 = bma01 and bma10 = '2' and bmb05 is null and bmb03 = ima01 and ima06 not in ('102','103')"
        Dim HasRows As Decimal = oCommand.ExecuteScalar()
        If HasRows = 0 Then
            MsgBox("NO DATA")
            Return
        End If
        Me.ProgressBar1.Maximum = HasRows
        BackgroundWorker1.RunWorkerAsync()
        'ExportToExcel()
        'SaveExcel()
    End Sub

    Private Sub Form89_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        Me.ProgressBar1.Value = 0
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Material_Report"
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
        oCommand.CommandText = "select distinct bmb03,ima02,ima021,ima25 from bmb_file,bma_file,ima_file where  bmb01 = bma01 and bma10 = '2' and bmb05 is null and bmb03 = ima01 and ima06 not in ('102','103') order by bmb03"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                ClearALLSQ()
                Ws.Cells(LineZ, 1) = oReader.Item("bmb03")
                Ws.Cells(LineZ, 2) = oReader.Item("ima02")
                Ws.Cells(LineZ, 3) = oReader.Item("ima021")
                Ws.Cells(LineZ, 4) = oReader.Item("ima25")
                SQ1 = PML_Q(oReader.Item("bmb03"))
                Ws.Cells(LineZ, 7) = SQ1
                SQ2 = PMN_Q(oReader.Item("bmb03"))
                Ws.Cells(LineZ, 8) = SQ2
                SQ3 = SFA_Q1(oReader.Item("bmb03"))
                Ws.Cells(LineZ, 6) = SQ3
                SQ4 = SFB_Q2(oReader.Item("bmb03"))
                Ws.Cells(LineZ, 9) = SQ4
                SQ5 = RVB_Q(oReader.Item("bmb03"))
                Ws.Cells(LineZ, 10) = SQ5
                SQ6 = AVL_STK_Q(oReader.Item("bmb03"))
                Ws.Cells(LineZ, 11) = SQ6
                SQ7 = ATP_Q(oReader.Item("bmb03"))
                Ws.Cells(LineZ, 5) = SQ7
                Ws.Cells(LineZ, 12) = Stock_Q(oReader.Item("bmb03"), "D146101")
                Ws.Cells(LineZ, 13) = Stock_Q(oReader.Item("bmb03"), "D146102")
                Ws.Cells(LineZ, 14) = Stock_Q(oReader.Item("bmb03"), "D146108")
                Ws.Cells(LineZ, 15) = Stock_Q(oReader.Item("bmb03"), "D180002")
                Ws.Cells(LineZ, 16) = Stock_Q(oReader.Item("bmb03"), "D353102")
                Ws.Cells(LineZ, 17) = Stock_Q(oReader.Item("bmb03"), "D353202")
                Ws.Cells(LineZ, 18) = Stock_Q(oReader.Item("bmb03"), "D353502")
                Ws.Cells(LineZ, 19) = Stock_Q(oReader.Item("bmb03"), "D353602")
                Ws.Cells(LineZ, 20) = Stock_Q(oReader.Item("bmb03"), "D356102")
                Ws.Cells(LineZ, 21) = Stock_Q(oReader.Item("bmb03"), "D356302")
                Ws.Cells(LineZ, 22) = Stock_Q(oReader.Item("bmb03"), "D356402")
                Ws.Cells(LineZ, 23) = Stock_Q(oReader.Item("bmb03"), "D356502")
                Ws.Cells(LineZ, 24) = Stock_Q(oReader.Item("bmb03"), "D356602")
                Ws.Cells(LineZ, 25) = Stock_Q(oReader.Item("bmb03"), "D363202")
                Ws.Cells(LineZ, 26) = Stock_Q(oReader.Item("bmb03"), "D363502")
                Ws.Cells(LineZ, 27) = Stock_Q(oReader.Item("bmb03"), "D373202")
                Ws.Cells(LineZ, 28) = Stock_Q(oReader.Item("bmb03"), "D373502")
                Ws.Cells(LineZ, 29) = Stock_Q(oReader.Item("bmb03"), "D230001")
                LineZ += 1
                Me.ProgressBar1.Value += 1
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.HorizontalAlignment = xlCenter
        Ws.Cells(1, 1) = "料号"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(1, 2) = "品名"
        Ws.Cells(1, 3) = "规格"
        Ws.Cells(1, 4) = "单位"
        Ws.Cells(1, 5) = "预计可用量"
        Ws.Cells(1, 6) = "工单备用量"
        Ws.Cells(1, 7) = "请购量"
        Ws.Cells(1, 8) = "采购量"
        Ws.Cells(1, 9) = "委外在制"
        Ws.Cells(1, 10) = "IQC在检"
        Ws.Cells(1, 11) = "可用量"
        Ws.Cells(1, 12) = "D146101 保税仓"
        Ws.Cells(1, 13) = "D146102 非保税仓"
        Ws.Cells(1, 14) = "D146108 客供仓"
        Ws.Cells(1, 15) = "D180002 委外仓"
        Ws.Cells(1, 16) = "D353102 裁纱线边仓"
        Ws.Cells(1, 17) = "D353202 预型线边仓"
        Ws.Cells(1, 18) = "D353502 成型线边仓"
        Ws.Cells(1, 19) = "D353602 CNC线边仓"
        Ws.Cells(1, 20) = "D356102 补土线边仓"
        Ws.Cells(1, 21) = "D356302 涂装线边仓"
        Ws.Cells(1, 22) = "D356402 胶合线边仓"
        Ws.Cells(1, 23) = "D356502 抛光线边仓"
        Ws.Cells(1, 24) = "D356602 包装线边仓"
        Ws.Cells(1, 25) = "D363202 预型线边仓"
        Ws.Cells(1, 26) = "D363502 PCM成型线边仓"
        Ws.Cells(1, 27) = "D373202 结构件预型线边仓"
        Ws.Cells(1, 28) = "D373502 结构件成型线边仓"
        Ws.Cells(1, 29) = "D230001 研发样品仓"

        LineZ = 2
    End Sub
    Private Function PML_Q(ByVal bmb03 As String)
        oCommand2.CommandText = "SELECT NVL(SUM((pml20-pml21)*pml09),0) FROM pml_file, pmk_file WHERE pml04 = '"
        oCommand2.CommandText += bmb03 & "' And pml01 = pmk01 AND pml20 > pml21 AND ( pml16 <='2' OR pml16='S' OR pml16='R' OR pml16='W') "
        oCommand2.CommandText += "AND pml011 !='SUB' AND pmk18 != 'X' AND pmk13 = 'D1461' "  'ADd by cloud 20170705
        Dim PQ As Decimal = oCommand2.ExecuteScalar
        Return PQ
    End Function
    Private Function PMN_Q(ByVal bmb03 As String)
        oCommand2.CommandText = "SELECT NVL(SUM((pmn20-pmn50+pmn55+pmn58)*pmn09),0) FROM pmn_file, pmm_file WHERE pmn04 = '"
        oCommand2.CommandText += bmb03 & "' AND pmn01 = pmm01 AND pmn20 > pmn50-pmn55-pmn58 AND ( pmn16 <='2' OR pmn16='S' OR pmn16='R' OR pmn16='W') "
        oCommand2.CommandText += "AND pmn011 !='SUB' AND pmm18 != 'X'"
        Dim PQ As Decimal = oCommand2.ExecuteScalar
        Return PQ
    End Function
    Private Function SFA_Q1(ByVal bmb03 As String)
        Dim l_sfa_q1 = 0
        oCommand2.CommandText = "SELECT sfa_file.*  FROM sfb_file,sfa_file WHERE sfa03 = '"
        oCommand2.CommandText += bmb03 & "' AND sfb01 = sfa01  AND sfb04 !='8' AND sfb87!='X' AND sfb02 != '15' "
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                oCommand3.CommandText = "SELECT nvl(SUM(rvv17),0) FROM rvv_file WHERE rvv18='" & oReader2.Item("sfa01") & "' AND rvv31='" & oReader2.Item("sfa03") & "'"
                Dim l_rvv17 = oCommand3.ExecuteScalar()
                l_sfa_q1 += (oReader2.Item("sfa05") - oReader2.Item("sfa06") - oReader2.Item("sfa065") + oReader2.Item("sfa063") - oReader2.Item("sfa062") + l_rvv17) * oReader2.Item("sfa13")
            End While
        End If
        oReader2.Close()
        Return l_sfa_q1
    End Function
    Private Function SFB_Q2(ByVal bmb03 As String)
        oCommand2.CommandText = "SELECT NVL(SUM((sfb08-sfb09-sfb10-sfb11-sfb12)*ima55_fac),0) FROM sfb_file,ima_file WHERE sfb05=ima01 AND sfb05 = '"
        oCommand2.CommandText += bmb03 & "' AND (sfb02='7' OR sfb02='8') AND sfb08 > (sfb09+sfb10+sfb11+sfb12) AND sfb87!='X' "
        Dim PQ1 As Decimal = oCommand2.ExecuteScalar
        oCommand2.CommandText = "SELECT NVL(SUM(shb114),0) FROM shb_file,sfb_file WHERE shb10 ='" & bmb03 & "' AND shb05 = sfb01 AND shb10 = sfb05 AND sfb04 < '8' AND sfb87 != 'X' AND ( sfb02 ='7' AND sfb02 ='8')"
        Dim PQ2 As Decimal = oCommand2.ExecuteScalar()
        oCommand2.CommandText = "SELECT NVL(SUM((rvb07-rvb29-rvb30)*pmn09),0) FROM rvb_file, rva_file, pmn_file WHERE rvb05 ='" & bmb03 & "' AND rvb01=rva01 AND rvb04 = pmn01 AND rvb03 = pmn02 AND rvb07 > (rvb29+rvb30) AND rvaconf='Y' AND rva10 ='SUB' AND pmn43 = 0"
        Dim PQ3 As Decimal = oCommand2.ExecuteScalar()
        Dim PQ4 As Decimal = PQ1 - PQ2 + PQ3
        Return PQ4
    End Function
    Private Function RVB_Q(ByVal bmb03 As String)
        oCommand2.CommandText = "SELECT NVL(SUM((rvb07-rvb29-rvb30)*pmn09),0) FROM rvb_file, rva_file, pmn_file WHERE rvb05 = '"
        oCommand2.CommandText += bmb03 & "' AND rvb01=rva01 AND rvb04 = pmn01 AND rvb03 = pmn02 AND rvb07 > (rvb29+rvb30) AND rvaconf='Y' AND rva10 != 'SUB'"
        Dim PQ As Decimal = oCommand2.ExecuteScalar
        Return PQ
    End Function
    Private Function AVL_STK_Q(ByVal bmb03 As String)
        oCommand2.CommandText = "SELECT NVL(SUM(img10*img21),0) FROM img_file WHERE img01 = '"
        oCommand2.CommandText += bmb03 & "' AND img23 = 'Y' and img02 not in ('D146104','D146109','D400001','D230001','D310001')"
        Dim PQ As Decimal = oCommand2.ExecuteScalar
        Return PQ
    End Function
    Private Function ATP_Q(ByVal bmb03 As String)
        ' 受訂量
        oCommand2.CommandText = "SELECT NVL(SUM((oeb12-oeb24+oeb25-oeb26)*oeb05_fac),0) FROM oeb_file, oea_file WHERE oeb04 = '"
        oCommand2.CommandText += bmb03 & "' AND oeb01 = oea01 AND oea00<>'0' AND oeb70 = 'N' AND oeb12-oeb24+oeb25-oeb26>0 AND oeaconf = 'Y'"
        Dim PQ1 As Decimal = oCommand2.ExecuteScalar
        ' 工單在制
        oCommand2.CommandText = "SELECT NVL(SUM((sfb08-sfb09-sfb10-sfb11-sfb12)*ima55_fac),0) FROM sfb_file,ima_file WHERE sfb05=ima01 AND sfb05 = '" & bmb03 & "' AND sfb04 <'8' AND ( sfb02 !='7' AND sfb02 !='8' AND sfb02 !='11' AND sfb02 != '15' ) AND sfb08 > (sfb09+sfb10+sfb11+sfb12) AND sfb87!='X'"
        Dim PQ2 As Decimal = oCommand2.ExecuteScalar()
        oCommand2.CommandText = "SELECT NVL(SUM(shb114),0) FROM shb_file,sfb_file WHERE shb10 ='" & bmb03 & "' AND shb05 = sfb01 AND shb10 = sfb05 AND sfb04 < '8' AND sfb87 != 'X' AND ( sfb02 !='7' AND sfb02 !='8' AND sfb02 !='11' AND sfb02 != '15' ) "
        Dim PQ3 As Decimal = oCommand2.ExecuteScalar()

        Dim ATP_Q_FINAL As Decimal = SQ6 - PQ1 - SQ3 + SQ1 + SQ2 + SQ5 + SQ4 + (PQ2 - PQ3)
        Return ATP_Q_FINAL
    End Function
    Private Function Stock_Q(ByVal bmb03 As String, ByVal img02 As String)
        oCommand2.CommandText = "SELECT NVL(SUM(img10*img21),0) FROM img_file WHERE img01 = '"
        oCommand2.CommandText += bmb03 & "' AND img02 = '" & img02 & "'"
        Dim PQ As Decimal = oCommand2.ExecuteScalar
        Return PQ
    End Function
    Private Sub ClearALLSQ()
        SQ1 = 0
        SQ2 = 0
        SQ3 = 0
        SQ4 = 0
        SQ5 = 0
        SQ6 = 0
    End Sub
End Class