Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel

Public Class Form10
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim tStation1 As String
    Dim ptime As String = String.Empty
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        ptime = Today.AddDays(-1).ToString("yyyy/MM/dd")
        ptime = ptime & " 08:00:00"
        Me.DateTimePicker1.Value = Convert.ToDateTime(ptime)
        Me.DateTimePicker2.Value = Convert.ToDateTime(ptime).AddDays(1).AddSeconds(-1)
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
        BindModel_Station()
    End Sub
    Private Sub BindModel_Station()
        Me.ComboBox1.Items.Clear()
        mSQLS1.CommandText = "SELECT station FROM station "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox1.Items.Add(mSQLReader.Item(0).ToString())
            End While
        End If
        mSQLReader.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If IsNothing(ComboBox1.SelectedItem) Then
            MsgBox("请选择工站")
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
        tStation1 = Me.ComboBox1.SelectedItem.ToString()
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        LineZ = 8
        'mSQLS1.CommandText = "select model,sum(t1) as t1 from ( "
        'mSQLS1.CommandText += "select lot.model,count(sn) as t1 from tracking "
        'mSQLS1.CommandText += "left join lot on tracking.lot = lot.lot "
        'mSQLS1.CommandText += "where timeout between '"
        'mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        'mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND station = '"
        'mSQLS1.CommandText += tStation1 & "' group by lot.model "
        'mSQLS1.CommandText += "union all "
        'mSQLS1.CommandText += "select lot.model,count(sn) as t1 from tracking_dup "
        'mSQLS1.CommandText += "left join lot on tracking_dup.lot = lot.lot "
        'mSQLS1.CommandText += "where timeout between '"
        'mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        'mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND station = '"
        'mSQLS1.CommandText += tStation1 & "' group by lot.model ) as C group by model order by model"
        mSQLS1.CommandText = "select cf01,model,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6"
        mSQLS1.CommandText += ",sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11 from ( "
        mSQLS1.CommandText += "select model_station_paravalue.cf01,lot.model,0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,1 as t11 from tracking "
        mSQLS1.CommandText += "FULL JOIN LOT ON tracking.lot = lot.lot left join model_station_paravalue on model_station_paravalue.profilename = 'ERP' and lot.model = model_station_paravalue.model "
        mSQLS1.CommandText += "and model_station_paravalue.station = '" & tStation1 & "' WHERE tracking.TIMEOUT BETWEEN '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND tracking.station = '" & tStation1 & "' "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select model_station_paravalue.cf01,lot.model,0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,1 as t11 from scrap_tracking "
        mSQLS1.CommandText += "FULL JOIN LOT ON scrap_tracking.lot = lot.lot left join model_station_paravalue on model_station_paravalue.profilename = 'ERP' and lot.model = model_station_paravalue.model "
        mSQLS1.CommandText += "and model_station_paravalue.station = '" & tStation1 & "' WHERE scrap_tracking.TIMEOUT BETWEEN '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND scrap_tracking.station = '" & tStation1 & "' and scrap_tracking.station in ('0112','0120','0130','0140','0150','0160','0170') "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "SELECT model_station_paravalue.cf01, lot.model,case when failure.failstation = '0670' then 1 else 0 end as t1,"
        mSQLS1.CommandText += "case when failure.failstation = '0645' then 1 else 0 end as t2,case when failure.failstation = '0640' then 1 else 0 end as t3,"
        mSQLS1.CommandText += "case when failure.failstation = '0590' then 1 else 0 end as t4,case when failure.failstation = '0475' then 1 else 0 end as t5,"
        mSQLS1.CommandText += "case when failure.failstation = '0430' then 1 else 0 end as t6,case when failure.failstation = '0627' then 1 else 0 end as t7,"
        mSQLS1.CommandText += "case when failure.failstation = '0620' then 1 else 0 end as t8,case when failure.failstation = '0490' then 1 else 0 end as t9,"
        mSQLS1.CommandText += "case when failure.failstation is null then 1 else 0 end as t10,0 as t11  "
        mSQLS1.CommandText += "FROM TRACKING_DUP FULL JOIN LOT ON tracking_dup.lot = lot.lot left join model_station_paravalue on model_station_paravalue.profilename = 'ERP' and lot.model = model_station_paravalue.model "
        mSQLS1.CommandText += "and model_station_paravalue.station = '" & tStation1 & "' LEFT JOIN FAILURE on tracking_dup.sn = failure.sn and tracking_dup.station = failure.rework "
        mSQLS1.CommandText += "and abs(datediff(second,tracking_dup.timein,failure.reworktime)) < 3 and abs(datediff(second,tracking_dup.timeout,failure.reworktime_finish)) < 3 "
        mSQLS1.CommandText += "WHERE tracking_dup.TIMEOUT BETWEEN '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND tracking_dup.station = '" & tStation1 & "' ) as TT group by tt.cf01 ,tt.model "
        mSQLS1.CommandTimeout = 300
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Dim ERPPN As String = String.Empty
                'Dim ERPPN As String = GetERPPN(mSQLReader.Item("model").ToString(), tStation1)
                If Not IsDBNull(mSQLReader.Item("cf01")) Then
                    ERPPN = mSQLReader.Item("cf01")
                End If
                Ws.Cells(LineZ, 1) = ERPPN
                Ws.Cells(LineZ, 2) = mSQLReader.Item("model").ToString()
                Ws.Cells(LineZ, 4) = mSQLReader.Item("t1").ToString()
                Ws.Cells(LineZ, 5) = mSQLReader.Item("t2").ToString()
                Ws.Cells(LineZ, 6) = mSQLReader.Item("t3").ToString()
                Ws.Cells(LineZ, 7) = mSQLReader.Item("t4").ToString()
                Ws.Cells(LineZ, 8) = mSQLReader.Item("t5").ToString()
                Ws.Cells(LineZ, 9) = mSQLReader.Item("t6").ToString()
                Ws.Cells(LineZ, 10) = mSQLReader.Item("t7").ToString()
                Ws.Cells(LineZ, 11) = mSQLReader.Item("t8").ToString()
                Ws.Cells(LineZ, 12) = mSQLReader.Item("t9").ToString()
                Ws.Cells(LineZ, 13) = mSQLReader.Item("t10").ToString()
                Ws.Cells(LineZ, 14) = mSQLReader.Item("t11").ToString()
                Ws.Cells(LineZ, 15) = "=SUM(D" & LineZ & ":N" & LineZ & ")"
                If Not String.IsNullOrEmpty(ERPPN) Then
                    Dim ima58 As Decimal = GetStandardIE(ERPPN)
                    Ws.Cells(LineZ, 3) = ima58
                End If
                LineZ += 1
            End While
        End If
        mSQLReader.Close()

    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 19
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "I1")
        oRng.Merge()
        oRng = Ws.Range("A2", "I2")
        oRng.Merge()
        oRng = Ws.Range("A3", "I3")
        oRng.Merge()
        oRng = Ws.Range("A4", "I4")
        oRng.Merge()
        oRng = Ws.Range("A5", "I5")
        oRng.Merge()
        oRng = Ws.Range("D6", "L6")
        oRng.Merge()
        oRng = Ws.Range("A6", "A7")
        oRng.Merge()
        oRng = Ws.Range("B6", "B7")
        oRng.Merge()
        oRng = Ws.Range("C6", "C7")
        oRng.Merge()
        oRng = Ws.Range("M6", "M7")
        oRng.Merge()
        oRng = Ws.Range("N6", "N7")
        oRng.Merge()
        oRng = Ws.Range("O6", "O7")
        oRng.Merge()
        oRng = Ws.Range("P6", "P7")
        oRng.Merge()
        oRng = Ws.Range("Q6", "Q7")
        oRng.Merge()
        oRng = Ws.Range("R6", "R7")
        oRng.Merge()
        oRng = Ws.Range("S6", "S7")
        oRng.Merge()
        oRng = Ws.Range("T6", "T7")
        oRng.Merge()
        Ws.Cells(1, 1) = "东莞艾可迅复合材料有限公司"
        Ws.Cells(2, 1) = "Dongguan Action Composites LTD Co."
        Ws.Cells(3, 1) = "生产部日产量统计报表"
        Ws.Cells(4, 1) = "Production Daily Report"
        Ws.Cells(5, 1) = "日期/Date:  " & TimeS2.Year & "年/Year       " & TimeS2.Month & "月/Month      " & TimeS2.Day & "日/Day    工段别/Station：              " & tStation1
        Ws.Cells(6, 1) = "产品型号 Products No."
        Ws.Cells(6, 2) = "Product name 产品名称"
        Ws.Cells(6, 3) = "Tacttime (min)"
        Ws.Cells(6, 4) = "返工完工"
        Ws.Cells(7, 4) = "成品检（0670）"
        Ws.Cells(7, 5) = "抛光检2（0645）"
        Ws.Cells(7, 6) = "抛光检1（0640）"
        Ws.Cells(7, 7) = "涂装检（0590）"
        Ws.Cells(7, 8) = "研磨检2（0475）"
        Ws.Cells(7, 9) = "研磨检1（0430）"
        Ws.Cells(7, 10) = "胶合检3（0620）"
        Ws.Cells(7, 11) = "胶合检2（0620）"
        Ws.Cells(7, 12) = "胶合检1（0490）"
        Ws.Cells(6, 13) = "非正常完工"
        Ws.Cells(6, 14) = "正常完工"
        Ws.Cells(6, 15) = "实际完成良品量 Actual FG Quanttity (PCS)"
        Ws.Cells(6, 16) = "HC"
        Ws.Cells(6, 17) = "Man hours /day"
        Ws.Cells(6, 18) = "Collected IE time(H)"
        Ws.Cells(6, 19) = "Effiencey"
        Ws.Cells(6, 20) = "Average man h"
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Daily Finished Product"
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
    End Sub

    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Function GetERPPN(ByVal modelA As String, stationA As String)
        Dim mConnection2 As New SqlClient.SqlConnection
        Dim mSQLS2 As New SqlClient.SqlCommand
        mConnection2.ConnectionString = Module1.OpenConnectionOfMes()
        mSQLS2.Connection = mConnection2
        mSQLS2.CommandType = CommandType.Text
        If mConnection2.State <> ConnectionState.Open Then
            mConnection2.Open()
        End If
        mSQLS2.CommandText = "SELECT cf01 FROM model_station_paravalue WHERE model = '"
        mSQLS2.CommandText += modelA & "' and station = '"
        mSQLS2.CommandText += stationA & "' and profilename = 'ERP'"
        Dim erppn As String = String.Empty
        Try
            erppn = mSQLS2.ExecuteScalar()
        Catch ex As Exception
            MsgBox(ex.Message())
            erppn = ""
        Finally
            mConnection2.Close()
        End Try
        Return erppn

    End Function
    Private Function GetStandardIE(ByVal ERPPN As String)
        Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
        If oConnection.State <> ConnectionState.Open Then
            oCommand.Connection = oConnection
            oCommand.CommandType = CommandType.Text
            oConnection.Open()
        End If
        oCommand.CommandText = "SELECT nvl(ima58,0) FROM ima_file WHERE ima01 = '" & ERPPN & "'"
        Dim l_ima58 As Decimal = 0
        Try
            l_ima58 = oCommand.ExecuteScalar()
        Catch ex As Exception
            MsgBox(ex.Message())
            l_ima58 = 0
        Finally
            oConnection.Close()
        End Try
        Return l_ima58
    End Function
End Class