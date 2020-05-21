Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form14
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Dim g_success As Boolean = True

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
        TimeS1 = DateTimePicker1.Value
        TimeS2 = DateTimePicker2.Value
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Form14_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
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
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        If g_success = True Then
            SaveFileDialog1.FileName = "Fail_Detail"
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
            If mConnection.State = ConnectionState.Open Then
                Try
                    mConnection.Close()
                    Module1.KillExcelProcess(OldExcel)
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
        End If
    End Sub
    Private Sub ExportToExcel()
        mSQLS1.CommandText = "select count(sn) from failure left join lot on failure.lot = lot.lot "
        mSQLS1.CommandText += "where failtime between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and failstation in "
        mSQLS1.CommandText += "('0670','0645','0640','0590','0475','0430','0627','0620','0490') "
        Dim HaveReport As Integer = mSQLS1.ExecuteScalar()
        If HaveReport = 0 Then
            g_success = False
            MsgBox("没有资料，请重选条件")
            Return
        End If
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat3()
        'LineZ = 3
        mSQLS1.CommandText = "select model,desc_th,desc_en ,rework,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9 from ( "
        mSQLS1.CommandText += "select model,defect.desc_th,desc_en,rework,case when failstation = '0670' then 1 else 0 end as t1,case when failstation = '0645' then 1 else 0 end as t2,"
        mSQLS1.CommandText += "case when failstation = '0640' then 1 else 0 end as t3,case when failstation = '0590' then 1 else 0 end as t4,case when failstation = '0475' then 1 else 0 end as t5,"
        mSQLS1.CommandText += "case when failstation = '0430' then 1 else 0 end as t6,case when failstation = '0627' then 1 else 0 end as t7,case when failstation = '0620' then 1 else 0 end as t8,case when failstation = '0490' then 1 else 0 end as t9 "
        mSQLS1.CommandText += "from failure left join lot on failure.lot = lot.lot left join defect on failure.defect = defect.defect where failtime between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and failstation in "
        mSQLS1.CommandText += "('0670','0645','0640','0590','0475','0430','0627','0620','0490') ) as C1 group by model,desc_th,desc_en ,rework order by model,rework"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("desc_en") & " " & mSQLReader.Item("desc_th")
                Ws.Cells(LineZ, 3) = GetERPPN(mSQLReader.Item("model").ToString(), mSQLReader.Item("rework").ToString())
                Ws.Cells(LineZ, 4) = "'" & mSQLReader.Item("rework")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("t1")
                If mSQLReader.Item("t1") > 0 Then
                    Ws.Cells(LineZ, 6) = GetERPPN(mSQLReader.Item("model").ToString(), "0670")
                End If
                Ws.Cells(LineZ, 7) = mSQLReader.Item("t2")
                If mSQLReader.Item("t2") > 0 Then
                    Ws.Cells(LineZ, 8) = GetERPPN(mSQLReader.Item("model").ToString(), "0645")
                End If
                Ws.Cells(LineZ, 9) = mSQLReader.Item("t3")
                If mSQLReader.Item("t3") > 0 Then
                    Ws.Cells(LineZ, 10) = GetERPPN(mSQLReader.Item("model").ToString(), "0640")
                End If
                Ws.Cells(LineZ, 11) = mSQLReader.Item("t4")
                If mSQLReader.Item("t4") > 0 Then
                    Ws.Cells(LineZ, 12) = GetERPPN(mSQLReader.Item("model").ToString(), "0590")
                End If
                Ws.Cells(LineZ, 13) = mSQLReader.Item("t5")
                If mSQLReader.Item("t5") > 0 Then
                    Ws.Cells(LineZ, 14) = GetERPPN(mSQLReader.Item("model").ToString(), "0475")
                End If
                Ws.Cells(LineZ, 15) = mSQLReader.Item("t6")
                If mSQLReader.Item("t6") > 0 Then
                    Ws.Cells(LineZ, 16) = GetERPPN(mSQLReader.Item("model").ToString(), "0430")
                End If
                Ws.Cells(LineZ, 17) = mSQLReader.Item("t7")
                If mSQLReader.Item("t7") > 0 Then
                    Ws.Cells(LineZ, 18) = GetERPPN(mSQLReader.Item("model").ToString(), "0627")
                End If
                Ws.Cells(LineZ, 19) = mSQLReader.Item("t8")
                If mSQLReader.Item("t8") > 0 Then
                    Ws.Cells(LineZ, 20) = GetERPPN(mSQLReader.Item("model").ToString(), "0620")
                End If
                Ws.Cells(LineZ, 21) = mSQLReader.Item("t9")
                If mSQLReader.Item("t9") > 0 Then
                    Ws.Cells(LineZ, 22) = GetERPPN(mSQLReader.Item("model").ToString(), "0490")
                End If
                Ws.Cells(LineZ, 23) = "=E" & LineZ & "+G" & LineZ & "+I" & LineZ & "+K" & LineZ & "+M" & LineZ & "+O" & LineZ & "+Q" & LineZ & "+S" & LineZ & "+U" & LineZ
                LineZ += 1
            End While
        End If
                'mSQLS1.CommandText = "SELECT SUM(T1) FROM ("
                'mSQLS1.CommandText += "select count(*) AS T1  from failure where failtime between '"
                'mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
                'mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                'mSQLS1.CommandText += "UNION ALL "
                'mSQLS1.CommandText += "select count(*)  from scrap_failure where failtime between '"
                'mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
                'mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' ) as bb"

                'Dim HaveReport As Integer = mSQLS1.ExecuteScalar()
                'If HaveReport = 0 Then
                '    g_success = False
                '    MsgBox("没有资料，请重选条件")
                '    Return
                'End If
                'xExcel = New Microsoft.Office.Interop.Excel.Application
                'xWorkBook = xExcel.Workbooks.Add()
                'Ws = xWorkBook.Sheets(1)
                'AdjustExcelFormat()
                'LineZ = 3
                'mSQLS1.CommandText = "SELECT SUM(t1) as t1,sum(t2) as t2,sum(t3) as t3,model,value from ( "
                'mSQLS1.CommandText += "SELECT (case when aa.failstation = '0590' then COUNT(distinct sn) else 0 end) as t1,"
                'mSQLS1.CommandText += "(case when aa.failstation = '0640' or aa.failstation = '0645' then COUNT(distinct sn) else 0 end) as t2,"
                'mSQLS1.CommandText += "(case when aa.failstation = '0670' then COUNT(distinct sn) else 0 end) as t3,"
                'mSQLS1.CommandText += "aa.model,model_paravalue.value  FROM ( "
                'mSQLS1.CommandText += "select failure.sn,lot.model,failstation   from failure "
                'mSQLS1.CommandText += "FULL JOIN sn  ON failure.sn = sn.sn  "
                'mSQLS1.CommandText += "FULL JOIN lot on sn.lot = lot.lot "
                'mSQLS1.CommandText += "where failtime between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
                'mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and failstation in ('0590','0640','0645','0670') "
                'mSQLS1.CommandText += "AND rework in ('0460','0583') "
                'mSQLS1.CommandText += "UNION ALL "
                'mSQLS1.CommandText += "select scrap_failure.sn,lot.model,failstation  from scrap_failure "
                'mSQLS1.CommandText += "FULL JOIN scrap_sn  ON scrap_failure.sn = scrap_sn.sn  "
                'mSQLS1.CommandText += "FULL JOIN lot on scrap_sn.lot = lot.lot "
                'mSQLS1.CommandText += "where failtime  between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
                'mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and failstation in ('0590','0640','0645','0670') "
                'mSQLS1.CommandText += "AND rework in ('0460','0583') "
                'mSQLS1.CommandText += ") as AA LEFT JOIN model_paravalue ON AA.model = model_paravalue.model and model_paravalue.parameter = 'ERP PN' "
                'mSQLS1.CommandText += "group by aa.model,model_paravalue.value,aa.failstation "
                'mSQLS1.CommandText += ") as BB group by bb.model,bb.value"
                'mSQLReader = mSQLS1.ExecuteReader()
                'If mSQLReader.HasRows() Then
                '    While mSQLReader.Read()
                '        Ws.Cells(LineZ, 1) = mSQLReader.Item("value").ToString()
                '        Ws.Cells(LineZ, 2) = mSQLReader.Item("model").ToString()
                '        Ws.Cells(LineZ, 3) = mSQLReader.Item("t1")
                '        Ws.Cells(LineZ, 4) = mSQLReader.Item("t2")
                '        Ws.Cells(LineZ, 5) = mSQLReader.Item("t3")
                '        Ws.Cells(LineZ, 6) = mSQLReader.Item("t1") + mSQLReader.Item("t2") + mSQLReader.Item("t3")
                '        LineZ += 1
                '    End While
                'End If
                'mSQLReader.Close()
                '' 以上是補土, 以下膠合
                'Ws = xWorkBook.Sheets(2)
                'Ws.Activate()
                'AdjustExcelFormat1()
                'LineZ = 3
                'mSQLS1.CommandText = "SELECT SUM(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,model,value from ( "
                'mSQLS1.CommandText += "SELECT (case when aa.failstation = '0430' then COUNT(distinct sn) else 0 end) as t1,"
                'mSQLS1.CommandText += "(case when aa.failstation = '0590' then COUNT(distinct sn) else 0 end) as t2,"
                'mSQLS1.CommandText += "(case when aa.failstation = '0640' or aa.failstation = '0645' then COUNT(distinct sn) else 0 end) as t3,"
                'mSQLS1.CommandText += "(case when aa.failstation = '0670' then COUNT(distinct sn) else 0 end) as t4,"
                'mSQLS1.CommandText += "aa.model,model_paravalue.value  FROM ( "
                'mSQLS1.CommandText += "select failure.sn,lot.model,failstation   from failure "
                'mSQLS1.CommandText += "FULL JOIN sn  ON failure.sn = sn.sn  "
                'mSQLS1.CommandText += "FULL JOIN lot on sn.lot = lot.lot "
                'mSQLS1.CommandText += "where failtime between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
                'mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and failstation in ('0430','0590','0640','0645','0670') "
                'mSQLS1.CommandText += "AND rework in ('0480','0610') "
                'mSQLS1.CommandText += "UNION ALL "
                'mSQLS1.CommandText += "select scrap_failure.sn,lot.model,failstation  from scrap_failure "
                'mSQLS1.CommandText += "FULL JOIN scrap_sn  ON scrap_failure.sn = scrap_sn.sn  "
                'mSQLS1.CommandText += "FULL JOIN lot on scrap_sn.lot = lot.lot "
                'mSQLS1.CommandText += "where failtime  between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
                'mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and failstation in ('0430','0590','0640','0645','0670') "
                'mSQLS1.CommandText += "AND rework in ('0480','0610') "
                'mSQLS1.CommandText += ") as AA LEFT JOIN model_paravalue ON AA.model = model_paravalue.model and model_paravalue.parameter = 'ERP PN' "
                'mSQLS1.CommandText += "group by aa.model,model_paravalue.value,aa.failstation "
                'mSQLS1.CommandText += ") as BB group by bb.model,bb.value"
                'mSQLReader = mSQLS1.ExecuteReader()
                'If mSQLReader.HasRows() Then
                '    While mSQLReader.Read()
                '        Ws.Cells(LineZ, 1) = mSQLReader.Item("value").ToString()
                '        Ws.Cells(LineZ, 2) = mSQLReader.Item("model").ToString()
                '        Ws.Cells(LineZ, 3) = mSQLReader.Item("t1")
                '        Ws.Cells(LineZ, 4) = mSQLReader.Item("t2")
                '        Ws.Cells(LineZ, 5) = mSQLReader.Item("t3")
                '        Ws.Cells(LineZ, 6) = mSQLReader.Item("t4")
                '        Ws.Cells(LineZ, 7) = mSQLReader.Item("t1") + mSQLReader.Item("t2") + mSQLReader.Item("t3") + mSQLReader.Item("t4")
                '        LineZ += 1
                '    End While
                'End If
                'mSQLReader.Close()
                '' 以下拋光
                'Ws = xWorkBook.Sheets(3)
                'Ws.Activate()
                'AdjustExcelFormat2()
                'LineZ = 3
                'mSQLS1.CommandText = "SELECT SUM(t1) as t1,model,value from ( "
                'mSQLS1.CommandText += "SELECT (case when aa.failstation = '0670' then COUNT(distinct sn) else 0 end) as t1,"
                'mSQLS1.CommandText += "aa.model,model_paravalue.value  FROM ( "
                'mSQLS1.CommandText += "select failure.sn,lot.model,failstation   from failure "
                'mSQLS1.CommandText += "FULL JOIN sn  ON failure.sn = sn.sn  "
                'mSQLS1.CommandText += "FULL JOIN lot on sn.lot = lot.lot "
                'mSQLS1.CommandText += "where failtime between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
                'mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and failstation in ('0670') "
                'mSQLS1.CommandText += "AND rework in ('0630','0635') "
                'mSQLS1.CommandText += "UNION ALL "
                'mSQLS1.CommandText += "select scrap_failure.sn,lot.model,failstation  from scrap_failure "
                'mSQLS1.CommandText += "FULL JOIN scrap_sn  ON scrap_failure.sn = scrap_sn.sn  "
                'mSQLS1.CommandText += "FULL JOIN lot on scrap_sn.lot = lot.lot "
                'mSQLS1.CommandText += "where failtime  between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
                'mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and failstation in ('0670') "
                'mSQLS1.CommandText += "AND rework in ('0630','0635') "
                'mSQLS1.CommandText += ") as AA LEFT JOIN model_paravalue ON AA.model = model_paravalue.model and model_paravalue.parameter = 'ERP PN' "
                'mSQLS1.CommandText += "group by aa.model,model_paravalue.value,aa.failstation "
                'mSQLS1.CommandText += ") as BB group by bb.model,bb.value"
                'mSQLReader = mSQLS1.ExecuteReader()
                'If mSQLReader.HasRows() Then
                '    While mSQLReader.Read()
                '        Ws.Cells(LineZ, 1) = mSQLReader.Item("value").ToString()
                '        Ws.Cells(LineZ, 2) = mSQLReader.Item("model").ToString()
                '        Ws.Cells(LineZ, 3) = mSQLReader.Item("t1")
                '        Ws.Cells(LineZ, 4) = mSQLReader.Item("t1")
                '        LineZ += 1
                '    End While
                'End If
                'mSQLReader.Close()
    End Sub
    'Private Sub AdjustExcelFormat()
    '    xExcel.ActiveWindow.Zoom = 75
    '    Ws.Name = "补土"
    '    Ws.Columns.EntireColumn.ColumnWidth = 25
    '    Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
    '    oRng = Ws.Range("A1", "F1")
    '    oRng.Merge()
    '    oRng.Interior.Color = Color.LightBlue
    '    Ws.Cells(1, 1) = "补土段"
    '    Ws.Cells(2, 1) = "Product No. 产品料号"
    '    Ws.Cells(2, 2) = "Product Name 产品名称"
    '    Ws.Cells(2, 3) = "接涂装返工品"
    '    Ws.Cells(2, 4) = "接抛光返工品"
    '    Ws.Cells(2, 5) = "接包装返工品"
    '    Ws.Cells(2, 6) = "合计"
    '    oRng = Ws.Range("C2:E2")
    '    oRng.Interior.Color = Color.LightYellow
    '    oRng = Ws.Range("F2")
    '    oRng.Interior.Color = Color.LightPink
    'End Sub
    'Private Sub AdjustExcelFormat1()
    '    xExcel.ActiveWindow.Zoom = 75
    '    Ws.Name = "胶合"
    '    Ws.Columns.EntireColumn.ColumnWidth = 25
    '    Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
    '    oRng = Ws.Range("A1", "G1")
    '    oRng.Merge()
    '    oRng.Interior.Color = Color.LightBlue
    '    Ws.Cells(1, 1) = "胶合段"
    '    Ws.Cells(2, 1) = "Product No. 产品料号"
    '    Ws.Cells(2, 2) = "Product Name 产品名称"
    '    Ws.Cells(2, 3) = "接补土返工品"
    '    Ws.Cells(2, 4) = "接涂装返工品"
    '    Ws.Cells(2, 5) = "接抛光返工品"
    '    Ws.Cells(2, 6) = "接包装返工品"
    '    Ws.Cells(2, 7) = "合计"
    '    oRng = Ws.Range("C2:F2")
    '    oRng.Interior.Color = Color.LightYellow
    '    oRng = Ws.Range("G2")
    '    oRng.Interior.Color = Color.LightPink
    'End Sub
    'Private Sub AdjustExcelFormat2()
    '    xExcel.ActiveWindow.Zoom = 75
    '    Ws.Name = "拋光"
    '    Ws.Columns.EntireColumn.ColumnWidth = 25
    '    Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
    '    oRng = Ws.Range("A1", "D1")
    '    oRng.Merge()
    '    oRng.Interior.Color = Color.LightBlue
    '    Ws.Cells(1, 1) = "拋光段"
    '    Ws.Cells(2, 1) = "Product No. 产品料号"
    '    Ws.Cells(2, 2) = "Product Name 产品名称"
    '    Ws.Cells(2, 3) = "接包装返工品"
    '    Ws.Cells(2, 4) = "合计"
    '    oRng = Ws.Range("C2")
    '    oRng.Interior.Color = Color.LightYellow
    '    oRng = Ws.Range("D2")
    '    oRng.Interior.Color = Color.LightPink
    'End Sub
    Private Sub AdjustExcelFormat3()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "Rework"
        Ws.Columns.EntireColumn.ColumnWidth = 15
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Cells(1, 1) = "Product Name 产品名称 "
        Ws.Cells(1, 2) = "Defect 不良缺陷"
        Ws.Cells(1, 3) = "Product No. 产品料号"
        Ws.Cells(1, 4) = "Rework Station 返工工站"
        Ws.Cells(1, 5) = "成品检（0670）"
        Ws.Cells(1, 6) = "Product No.产品料号 "
        Ws.Cells(1, 7) = "抛光检2（0645）"
        Ws.Cells(1, 8) = "Product No.产品料号 "
        Ws.Cells(1, 9) = "抛光检1（0640）"
        Ws.Cells(1, 10) = "Product No.产品料号 "
        Ws.Cells(1, 11) = "涂装检（0590）"
        Ws.Cells(1, 12) = "Product No.产品料号 "
        Ws.Cells(1, 13) = "研磨检2（0475）"
        Ws.Cells(1, 14) = "Product No.产品料号 "
        Ws.Cells(1, 15) = "研磨检1（0430）"
        Ws.Cells(1, 16) = "Product No.产品料号 "
        Ws.Cells(1, 17) = "胶合检2（0627）"
        Ws.Cells(1, 18) = "Product No.产品料号 "
        Ws.Cells(1, 19) = "胶合检2（0620）"
        Ws.Cells(1, 20) = "Product No.产品料号 "
        Ws.Cells(1, 21) = "胶合检1（0490）"
        Ws.Cells(1, 22) = "Product No.产品料号 "
        Ws.Cells(1, 23) = "合计"
        LineZ = 2
    End Sub
    Private Function GetERPPN(ByVal model As String, stationC As String)
        Dim mConnection2 As New SqlClient.SqlConnection
        Dim mSQLS2 As New SqlClient.SqlCommand
        mConnection2.ConnectionString = Module1.OpenConnectionOfMes()
        If mConnection2.State <> ConnectionState.Open Then
            Try
                mConnection2.Open()
                mSQLS2.Connection = mConnection2
                mSQLS2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        mSQLS2.CommandText = "select cf01 from model_station_paravalue where profilename = 'ERP' and model = '"
        mSQLS2.CommandText += model & "' and station = '" & stationC & "'"
        Dim RV As String = mSQLS2.ExecuteScalar()
        mSQLS2.Dispose()
        mConnection2.Close()
        mConnection2.Dispose()
        Return RV
    End Function
End Class