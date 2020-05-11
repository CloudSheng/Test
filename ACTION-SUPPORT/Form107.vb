Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form107
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim ptime As String = String.Empty
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim HaveReport As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form107_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
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
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        mSQLS1.CommandText = "select Convert(varchar(100),scrap.datetime,111) as c1,scrap.sn,lot.model,model.modelname,model_station_paravalue.cf01,scrap.defect,scrap_sn.updatedstation   from scrap "
        mSQLS1.CommandText += "left join scrap_sn on scrap.sn = scrap_sn.sn left join lot on scrap.lot = lot.lot left join model on lot.model = model.model left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue.profilename = 'ERP' and scrap_sn.updatedstation = model_station_paravalue.station "
        mSQLS1.CommandText += "where scrap.datetime between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and scrap.defect not in ('0042','102','025','0048','0051','0052','012','056','075','083','095','100') "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("c1")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("sn")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("modelname")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("cf01")
                Ws.Cells(LineZ, 6) = mSQLReader.Item("defect")
                Ws.Cells(LineZ, 7) = mSQLReader.Item("updatedstation")
                Dim SU As String = String.Empty
                SU = GetNameE(mSQLReader.Item("sn"))
                Ws.Cells(LineZ, 8) = SU
                Dim DU As String = String.Empty
                DU = GetNameD(mSQLReader.Item("sn"))
                Ws.Cells(LineZ, 9) = DU
                LineZ += 1
                Me.Label3.Text = LineZ
            End While
        End If
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        'If HaveReport > 0 Then
        SaveExcel()
        'End If
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Scrap_Information"
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
        Ws.Columns.EntireColumn.ColumnWidth = 25
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.WrapText = True
        Ws.Cells(1, 1) = "MES报废日期"
        Ws.Cells(1, 2) = "SN#"
        Ws.Cells(1, 3) = "产品简称"
        Ws.Cells(1, 4) = "产品名称"
        Ws.Cells(1, 5) = "ERP 料号"
        Ws.Cells(1, 6) = "报废原因代码"
        Ws.Cells(1, 7) = "报废的站点"
        Ws.Cells(1, 8) = "报废原因对应的工序"
        Ws.Cells(1, 9) = "作业员工工号"
        oRng = Ws.Range("F1", "I1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        'oRng = Ws.Range("F1", "G1")
        'oRng.EntireColumn.NumberFormatLocal = "@"
        LineZ = 2
    End Sub
    Private Function GetNameD(ByVal sn As String)
        mSQLS2.CommandText = "select top 1 users,station from ( "
        mSQLS2.CommandText += "select station,timeout,users  from scrap_tracking where sn = '" & sn & "' "
        mSQLS2.CommandText += "union all "
        mSQLS2.CommandText += "select station,timeout,users from tracking_dup where sn = '" & sn & "' ) as ab order by timeout desc"
        Dim GN As String = String.Empty
        Try
            GN = mSQLS2.ExecuteScalar()

        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        Return GN
    End Function
    Private Function GetNameE(ByVal sn As String)
        mSQLS2.CommandText = "select top 1 station from ( "
        mSQLS2.CommandText += "select station,timeout,users  from scrap_tracking where sn = '" & sn & "' "
        mSQLS2.CommandText += "union all "
        mSQLS2.CommandText += "select station,timeout,users from tracking_dup where sn = '" & sn & "' ) as ab order by timeout desc"
        Dim GN As String = String.Empty
        Try
            GN = mSQLS2.ExecuteScalar()

        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        Return GN
    End Function
End Class