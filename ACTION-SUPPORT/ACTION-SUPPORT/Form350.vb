Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form350
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim ptime As String = String.Empty
    Dim MaxDetailCount As Int16 = 0
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim temp_cnt As Integer = 0
    Dim temp_cnt_1 As Integer = 0
    Dim erp_price As Double = 0
    Dim erp_amt As Double = 0
    Dim HaveReport As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form350_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        ptime = Today.AddDays(-7).ToString("yyyy/MM/dd")
        ptime = ptime & " 08:00:00"
        Me.DateTimePicker1.Value = Convert.ToDateTime(ptime)
        Me.DateTimePicker2.Value = Convert.ToDateTime(ptime).AddDays(6).AddSeconds(-1)
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
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
        CreateTempTable()
        CreateTempTable1()
        TimeS1 = DateTimePicker1.Value
        TimeS2 = DateTimePicker2.Value
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets.Add()  '20160901
        Ws = xWorkBook.Sheets(1)
        Ws.Name = "报废明细"
        AdjustExcelFormat()

        mSQLS1.CommandText = "select scrap.datetime,case when (right(cf01,2) = '35' or right(cf01,3) = '35A') and station <> '0331' then 'Mold 成型' "
        mSQLS1.CommandText += "           when (right(cf01,2) = '35' or right(cf01,3) = '35A') and station = '0331' then 'PCM 工段' "
        mSQLS1.CommandText += "           when right(cf01,2) = '36' or right(cf01,3) = '36A' then 'CNC' "
        mSQLS1.CommandText += "           when right(cf01,2) = '61' or right(cf01,3) = '61A' then 'Sanding 补土' "
        mSQLS1.CommandText += "           when right(cf01,2) = '64'  then 'Gluing 胶合 1' "
        mSQLS1.CommandText += "           when right(cf01,3) = '64A' then 'Gluing 胶合 2' "
        mSQLS1.CommandText += "           when right(cf01,3) = '64B' then 'Gluing 胶合 3' "
        mSQLS1.CommandText += "           when right(cf01,2) = '63' or right(cf01,3) = '63A' then 'Painting 涂装' "
        mSQLS1.CommandText += "           when right(cf01,2) = '65' or right(cf01,3) = '65A' then 'Polishing 抛光' "
        mSQLS1.CommandText += "           when right(cf01,2) = '66' or right(cf01,3) = '66A' then 'Packing 包装' "
        mSQLS1.CommandText += "       end as right_cf01,lot.model,cf01,count(*) as l_cnt "
        mSQLS1.CommandText += "  from scrap left join scrap_sn on scrap.sn = scrap_sn.sn left join lot on scrap.lot = lot.lot "
        mSQLS1.CommandText += "       left join model_station_paravalue on model_station_paravalue.profilename = 'ERP' and model_station_paravalue.station = scrap_sn.updatedstation "
        mSQLS1.CommandText += "       and model_station_paravalue.model = lot.model "
        mSQLS1.CommandText += "       left join defect on scrap.defect = defect.defect where scrap.defect not in ('0051','0052','100','114','0042','122') and scrap.datetime between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS1.CommandText += "       group by scrap.datetime,case when (right(cf01,2) = '35' or right(cf01,3) = '35A') and station <> '0331' then 'Mold 成型' "
        mSQLS1.CommandText += "                     when (right(cf01,2) = '35' or right(cf01,3) = '35A') and station = '0331' then 'PCM 工段' "
        mSQLS1.CommandText += "                     when right(cf01,2) = '36' or right(cf01,3) = '36A' then 'CNC' "
        mSQLS1.CommandText += "                     when right(cf01,2) = '61' or right(cf01,3) = '61A' then 'Sanding 补土' "
        mSQLS1.CommandText += "                     when right(cf01,2) = '64'  then 'Gluing 胶合 1' "
        mSQLS1.CommandText += "                     when right(cf01,3) = '64A' then 'Gluing 胶合 2' "
        mSQLS1.CommandText += "                     when right(cf01,3) = '64B' then 'Gluing 胶合 3' "
        mSQLS1.CommandText += "                     when right(cf01,2) = '63' or right(cf01,3) = '63A' then 'Painting 涂装' "
        mSQLS1.CommandText += "                     when right(cf01,2) = '65' or right(cf01,3) = '65A' then 'Polishing 抛光' "
        mSQLS1.CommandText += "                     when right(cf01,2) = '66' or right(cf01,3) = '66A' then 'Packing 包装' "
        mSQLS1.CommandText += "                end,lot.model,cf01 "

        mSQLReader = mSQLS1.ExecuteReader(CommandBehavior.CloseConnection)
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                If Not mSQLReader.Item("datetime") Is DBNull.Value And
                   Not mSQLReader.Item("cf01") Is DBNull.Value And Not mSQLReader.Item("right_cf01") Is DBNull.Value And
                   Not mSQLReader.Item("model") Is DBNull.Value And Not mSQLReader.Item("l_cnt") Is DBNull.Value Then
                    If oConnection.State <> ConnectionState.Open Then
                        Try
                            oConnection.Open()
                            oCommand.Connection = oConnection
                            oCommand.CommandType = CommandType.Text
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try
                    End If
                    Ws.Cells(LineZ, 1) = mSQLReader.Item("datetime")
                    Ws.Cells(LineZ, 2) = mSQLReader.Item("right_cf01")
                    Ws.Cells(LineZ, 3) = mSQLReader.Item("model")
                    Ws.Cells(LineZ, 4) = mSQLReader.Item("cf01")
                    Ws.Cells(LineZ, 5) = mSQLReader.Item("l_cnt")
                    If Not mSQLReader.Item("datetime") Is DBNull.Value And
                       Not mSQLReader.Item("cf01") Is DBNull.Value And Not mSQLReader.Item("right_cf01") Is DBNull.Value And
                       Not mSQLReader.Item("model") Is DBNull.Value And Not mSQLReader.Item("l_cnt") Is DBNull.Value Then
                        FindERPPrice(mSQLReader.Item("cf01"), mSQLReader.Item("right_cf01"), mSQLReader.Item("model"), mSQLReader.Item("l_cnt"))
                    End If
                    Ws.Cells(LineZ, 6) = erp_price
                    Ws.Cells(LineZ, 7) = mSQLReader.Item("l_cnt") * erp_price
                    LineZ += 1
                End If
            End While
        End If
        mSQLReader.Close()

        ' 第二頁    
        Ws = xWorkBook.Sheets(2)
        Ws = xWorkBook.Sheets.Add()
        Ws.Activate()
        Ws.Name = "工段汇总"
        AdjustExcelFormat2()
        oCommand.CommandText = "select right_cf01 ,amt FROM AMT_TEMP ORDER BY amt desc"

        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("right_cf01")
                Ws.Cells(LineZ, 2) = oReader.Item("amt")
                LineZ += 1
            End While
        End If
        oReader.Close()
        Ws.Cells(LineZ, 1) = "总计"
        Ws.Cells(LineZ, 2) = "=SUM(B3:B" & LineZ - 1 & ")"

        ' 第三頁  
        Ws = xWorkBook.Sheets(3)
        Ws = xWorkBook.Sheets.Add()
        Ws.Activate()
        Ws.Name = "型号汇总"
        AdjustExcelFormat3()
        oCommand.CommandText = "select model ,amt FROM AMT_TEMP_1 ORDER BY amt desc"

        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("model")
                Ws.Cells(LineZ, 2) = oReader.Item("amt")
                LineZ += 1
            End While
        End If
        oReader.Close()
        Ws.Cells(LineZ, 1) = "总计"
        Ws.Cells(LineZ, 2) = "=SUM(B3:B" & LineZ - 1 & ")"
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        'If HaveReport > 0 Then
        SaveExcel()
        'End If
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Quantity statistics of quality scrapping weekly"
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
        If mConnection.State = ConnectionState.Open Then
            Try
                mConnection.Close()
                Module1.KillExcelProcess(OldExcel)
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 15
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.WrapText = True
        oRng = Ws.Range("A1", "G1")
        oRng.Merge()
        oRng = Ws.Range("A1", "G2")
        oRng.EntireRow.RowHeight = 42
        oRng = Ws.Range("C2", "M2")
        oRng.EntireColumn.ColumnWidth = 17.25
        oRng = Ws.Range("A2", "B2")
        oRng.EntireColumn.ColumnWidth = 23.28
        Ws.Cells(1, 1) = "品质报废周报金额统计 Quantity statistics of quality scrapping weekly"
        Ws.Cells(2, 1) = "日期时间"
        oRng = Ws.Range("A2", "A3")
        oRng.Merge()
        Ws.Cells(2, 2) = "工段"
        oRng = Ws.Range("B2", "B3")
        oRng.Merge()
        Ws.Cells(2, 3) = "型号"
        oRng = Ws.Range("C2", "C3")
        oRng.Merge()
        Ws.Cells(2, 4) = "ERP料号"
        oRng = Ws.Range("D2", "D3")
        oRng.Merge()
        Ws.Cells(2, 5) = "数量"
        oRng = Ws.Range("E2", "E3")
        oRng.Merge()
        Ws.Cells(2, 6) = "标准成本单价"
        oRng = Ws.Range("F2", "F3")
        oRng.Merge()
        Ws.Cells(2, 7) = "金额"
        oRng = Ws.Range("G2", "G3")
        oRng.Merge()
        LineZ = 4
    End Sub
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 15
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.WrapText = True
        oRng = Ws.Range("A1", "B1")
        oRng.Merge()
        oRng = Ws.Range("A1", "B2")
        oRng.EntireRow.RowHeight = 42
        oRng = Ws.Range("C2", "M2")
        oRng.EntireColumn.ColumnWidth = 17.25
        oRng = Ws.Range("A2", "B2")
        oRng.EntireColumn.ColumnWidth = 23.28
        Ws.Cells(1, 1) = "品质报废周报金额统计 Quantity statistics of quality scrapping weekly"
        oRng = Ws.Range("A2", "A3")
        oRng.Merge()
        Ws.Cells(2, 1) = "工段"
        oRng = Ws.Range("B2", "B3")
        oRng.Merge()
        Ws.Cells(2, 2) = "金额"
        LineZ = 4
    End Sub
    Private Sub AdjustExcelFormat3()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 15
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.WrapText = True
        oRng = Ws.Range("A1", "B1")
        oRng.Merge()
        oRng = Ws.Range("A1", "B2")
        oRng.EntireRow.RowHeight = 42
        oRng = Ws.Range("C2", "M2")
        oRng.EntireColumn.ColumnWidth = 17.25
        oRng = Ws.Range("A2", "B2")
        oRng.EntireColumn.ColumnWidth = 23.28
        Ws.Cells(1, 1) = "品质报废周报金额统计 Quantity statistics of quality scrapping weekly"
        oRng = Ws.Range("A2", "A3")
        oRng.Merge()
        Ws.Cells(2, 1) = "型号"
        oRng = Ws.Range("B2", "B3")
        oRng.Merge()
        Ws.Cells(2, 2) = "金额"
        LineZ = 4
    End Sub
    Private Sub FindERPPrice(ByVal stb01 As String, r2_stb01 As String, mdl As String, Qty As Integer)
        oCommand.CommandText = "select stb02,stb03,SUM(stb07+stb08+stb09+stb09a) as s_stb789 FROM STB_FILE WHERE STB01 = '"
        oCommand.CommandText += stb01 & "' group by stb02, stb03 order by stb02 desc,stb03 desc"
        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            oReader.Read()
            erp_price = oReader.Item("s_stb789")
            erp_amt = erp_price * Qty
        End If
        oReader.Close()

        oCommand.CommandText = "select count(*) as temp_cnt from amt_temp WHERE right_cf01 = '"
        oCommand.CommandText += r2_stb01 & "'"
        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            oReader.Read()
            temp_cnt = oReader.Item("temp_cnt")
        End If
        oReader.Close()

        If temp_cnt > 0 Then
            oCommand.CommandText = "update amt_temp set amt = amt + "
            oCommand.CommandText += erp_amt & "where right_cf01 = '"
            oCommand.CommandText += r2_stb01 & "'"
            oReader = oCommand.ExecuteReader
        Else
            oCommand.CommandText = "insert into amt_temp values( '"
            oCommand.CommandText += r2_stb01 & "',"
            oCommand.CommandText += erp_amt & ")"
            oReader = oCommand.ExecuteReader
        End If
        oReader.Close()

        oCommand.CommandText = "select count(*) as temp_cnt_1 from amt_temp_1 WHERE model = '"
        oCommand.CommandText += mdl & "'"
        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            oReader.Read()
            temp_cnt_1 = oReader.Item("temp_cnt_1")
        End If
        oReader.Close()

        If temp_cnt_1 > 0 Then
            oCommand.CommandText = "update amt_temp_1 set amt = amt + "
            oCommand.CommandText += erp_amt & "where model = '"
            oCommand.CommandText += mdl & "'"
            oReader = oCommand.ExecuteReader
        Else
            oCommand.CommandText = "insert into amt_temp_1 values( '"
            oCommand.CommandText += mdl & "',"
            oCommand.CommandText += erp_amt & ")"
            oReader = oCommand.ExecuteReader
        End If
        oReader.Close()
    End Sub
    Private Sub CreateTempTable()
        Me.Label2.Text = "DROP TABLE"
        oCommand.CommandText = "DROP TABLE amt_temp"

        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Me.Label2.Text = "CREATE TABLE"
        oCommand.CommandText = "CREATE TABLE amt_temp (right_cf01 varchar2(40), amt number(15,3)) "

        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

    End Sub
    Private Sub CreateTempTable1()
        Me.Label2.Text = "DROP TABLE"
        oCommand.CommandText = "DROP TABLE amt_temp_1"

        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Me.Label2.Text = "CREATE TABLE"
        oCommand.CommandText = "CREATE TABLE amt_temp_1 (model varchar2(40), amt number(15,3)) "

        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

    End Sub
End Class