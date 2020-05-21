Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form133
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim mSQLReader2 As SqlClient.SqlDataReader
    Dim tStation1 As String = String.Empty
    Dim tEquipment As String = String.Empty
    Dim ptime As String = String.Empty
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form133_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        BindModel_Station()
        BindEquipment()
    End Sub
    Private Sub BindModel_Station()
        Me.ComboBox2.Items.Clear()
        mSQLS1.CommandText = "SELECT station FROM station "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox2.Items.Add(mSQLReader.Item(0).ToString())
            End While
        End If
        mSQLReader.Close()
        Me.ComboBox2.SelectedItem = "0380"
    End Sub
    Private Sub BindEquipment()
        Me.ComboBox1.Items.Clear()
        mSQLS1.CommandText = "select equipment_id from z_ms_equipment "
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
            MsgBox("请选择设备编号")
            Return
        End If
        If IsNothing(ComboBox2.SelectedItem) Then
            MsgBox("请选择工站")
            Return
        End If
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        TimeS1 = DateTimePicker1.Value
        TimeS2 = DateTimePicker2.Value
        tStation1 = Me.ComboBox2.SelectedItem.ToString()
        tEquipment = Me.ComboBox1.SelectedItem.ToString()
        'BackgroundWorker1.RunWorkerAsync()
        ExportToExcel()
        SaveExcel()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        'mSQLS1.CommandText = "select value,Dateadd(hour,-8,timeout) as c1,timeout,lot.model , ab.sn  from ( "
        'mSQLS1.CommandText += "select sn,timeout,station from tracking where tracking.timeout between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        'mSQLS1.CommandText += "union all "
        'mSQLS1.CommandText += "select sn,timeout,station from tracking_dup where tracking_dup.timeout between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        'mSQLS1.CommandText += "union all "
        'mSQLS1.CommandText += "select sn,timeout,station from scrap_tracking where scrap_tracking.timeout between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' ) as AB "
        'mSQLS1.CommandText += "left join paravalue on ab.sn = paravalue.sn and ab.station = paravalue.station left join lot on paravalue.lot = lot.lot where value = '" & tEquipment & "'"
        mSQLS1.CommandText = "select value,CONVERT(varchar(10) , Dateadd(hour,-8,timeout),101) as c1,timeout,lot.model , ab.sn  from ( "
        mSQLS1.CommandText += "select sn,timeout,station from tracking where tracking.timeout between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select sn,timeout,station from tracking_dup where tracking_dup.timeout between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS1.CommandText += ") as AB left join paravalue on ab.sn = paravalue.sn and ab.station = paravalue.station "
        mSQLS1.CommandText += "left join lot on paravalue.lot = lot.lot where value = '" & tEquipment & "' "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select value,CONVERT(varchar(10) , Dateadd(hour,-8,timeout),101) as c1,timeout,lot.model , ab.sn  from ( "
        mSQLS1.CommandText += "select sn,timeout,station from scrap_tracking where scrap_tracking.timeout between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS1.CommandText += ") as AB left join scrap_paravalue on ab.sn = scrap_paravalue.sn and ab.station = scrap_paravalue.station "
        mSQLS1.CommandText += "left join lot on scrap_paravalue.lot = lot.lot where value = '" & tEquipment & "' "

        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("value")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("c1") '.ToString("yyyy/MM/dd")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("timeout")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("sn")
                GetDefectCode(mSQLReader.Item("sn"))
                LineZ += 1
                            End While
                        End If
        mSQLReader.Close()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 19
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Cells(1, 1) = "Machine No."
        Ws.Cells(1, 2) = "Produce Date"
        Ws.Cells(1, 3) = "Machine Scanning time"
        Ws.Cells(1, 4) = "MES Description"
        Ws.Cells(1, 5) = "Sequence No."
        Ws.Cells(1, 6) = "Inspection Station"
        Ws.Cells(1, 7) = "Inspection Scanning time"
        Ws.Cells(1, 8) = "Defect code"
        Ws.Cells(1, 9) = "Defect Description"
        oRng = Ws.Range("F1", "F1")
        oRng.EntireColumn.NumberFormat = "@"
        oRng = Ws.Range("H1", "H1")
        oRng.EntireColumn.NumberFormat = "@"
        LineZ = 2
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Machine Defect Report"
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

    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub GetDefectCode(ByVal sn As String)
        mSQLS2.CommandText = "select * from ( select row_number() over(order by timeout) as t1,* from ( select station,timeout from tracking where sn = '"
        mSQLS2.CommandText += sn & "' and station = '" & tStation1 & "' "
        mSQLS2.CommandText += "union all "
        mSQLS2.CommandText += "select station,timeout from tracking_dup where sn = '" & sn & "' and station = '" & tStation1 & "' "
        mSQLS2.CommandText += "union all "
        mSQLS2.CommandText += "select station,timeout from scrap_tracking where sn = '" & sn & "' and station = '" & tStation1 & "' ) as ab ) as ac where t1 = 1"
        mSQLReader2 = mSQLS2.ExecuteReader()
        If mSQLReader2.HasRows Then
            While mSQLReader2.Read()
                Ws.Cells(LineZ, 6) = mSQLReader2.Item("station")
                Ws.Cells(LineZ, 7) = mSQLReader2.Item("timeout")
            End While
        End If
        mSQLReader2.Close()

        mSQLS2.CommandText = "select * from ( select row_number() over(order by failtime) as t1,ab.defect from ( select * from failure where sn = '"
        mSQLS2.CommandText += sn & "' AND failstation = '" & tStation1 & "' "
        mSQLS2.CommandText += "union all "
        mSQLS2.CommandText += "select * from scrap_failure where sn = '" & sn & "' AND failstation = '" & tStation1 & "' ) as ab ) as ac left join defect on ac.defect = defect.defect  where t1 = 1"
        mSQLReader2 = mSQLS2.ExecuteReader()
        If mSQLReader2.HasRows() Then
            While mSQLReader2.Read()
                Ws.Cells(LineZ, 8) = mSQLReader2.Item(1)
                Ws.Cells(LineZ, 9) = mSQLReader2.Item("desc_en")
            End While
        End If
        mSQLReader2.Close()
    End Sub
End Class