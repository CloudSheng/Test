Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form22
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim MaxDetailCount As Int16 = 0
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form22_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        mSQLS1.CommandText = "select count(*) from scrap_sn left join scrap on scrap_sn.sn = scrap.sn where scrap.datetime between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND updatedstation = '0330'"
        Dim HaveReport As Integer = mSQLS1.ExecuteScalar()
        Me.ProgressBar1.Maximum = HaveReport
        Me.ProgressBar1.Value = 0
        If HaveReport = 0 Then
            MsgBox("没有资料，请重选条件")
            Return
        End If
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        LineZ = 4
        mSQLS1.CommandText = "select scrap_tracking.sn,station,timein,timeout,scrap_tracking.users,users.name,scrap.defect,defect.desc_en  from scrap_tracking "
        mSQLS1.CommandText += "left join users on scrap_tracking.users = users.id "
        mSQLS1.CommandText += "left join scrap on scrap_tracking.sn = scrap.sn "
        mSQLS1.CommandText += "left join defect on scrap.defect = defect.defect "
        mSQLS1.CommandText += "where scrap_tracking.sn in ( "
        mSQLS1.CommandText += "select scrap_sn.sn from scrap_sn left join scrap on scrap_sn.sn = scrap.sn where  updatedstation = '0330' and datetime between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "') and scrap_tracking.station in ('0330','0150','0200','0320') order by sn,station desc"
        mSQLReader = mSQLS1.ExecuteReader()
        Dim TR As String = String.Empty
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                If String.IsNullOrEmpty(TR) Then
                    TR = mSQLReader.Item("sn")
                    Ws.Cells(LineZ, 2) = TR
                End If
                If TR <> mSQLReader.Item("sn") Then
                    LineZ += 1
                    TR = mSQLReader.Item("sn")
                    Ws.Cells(LineZ, 2) = TR
                End If
                Select Case mSQLReader.Item("station").ToString()
                    Case "0330"
                        Ws.Cells(LineZ, 1) = "'0330"
                        Ws.Cells(LineZ, 3) = mSQLReader.Item("timein")
                        Ws.Cells(LineZ, 4) = mSQLReader.Item("timeout")
                        Ws.Cells(LineZ, 5) = mSQLReader.Item("users") & " " & mSQLReader.Item("name")
                        Ws.Cells(LineZ, 6) = mSQLReader.Item("defect") & " " & mSQLReader.Item("desc_en")
                    Case "0150"
                        Ws.Cells(LineZ, 7) = "'0150"
                        Ws.Cells(LineZ, 8) = mSQLReader.Item("timein")
                        Ws.Cells(LineZ, 9) = mSQLReader.Item("timeout")
                        Ws.Cells(LineZ, 10) = mSQLReader.Item("users") & " " & mSQLReader.Item("name")
                    Case "0200"
                        Ws.Cells(LineZ, 11) = "'0200"
                        Ws.Cells(LineZ, 12) = mSQLReader.Item("timein")
                        Ws.Cells(LineZ, 13) = mSQLReader.Item("timeout")
                        Ws.Cells(LineZ, 14) = mSQLReader.Item("users") & " " & mSQLReader.Item("name")
                    Case "0320"
                        Ws.Cells(LineZ, 15) = "'0320"
                        Ws.Cells(LineZ, 16) = mSQLReader.Item("timein")
                        Ws.Cells(LineZ, 17) = mSQLReader.Item("timeout")
                        Ws.Cells(LineZ, 18) = mSQLReader.Item("users") & " " & mSQLReader.Item("name")
                End Select
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 30
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "F1")
        oRng.Merge()
        oRng = Ws.Range("G1", "J1")
        oRng.Merge()
        oRng = Ws.Range("K1", "N1")
        oRng.Merge()
        oRng = Ws.Range("O1", "R1")
        oRng.Merge()
        oRng = Ws.Range("A2", "R3")
        oRng.Interior.Color = Color.Green
        Ws.Cells(1, 1) = "基准工站信息（工站代码：0330成型检验）"
        Ws.Cells(1, 7) = "追遡工站信息（工站代码：0150预型）"
        Ws.Cells(1, 11) = "追遡工站信息（工站代码：0200成型）"
        Ws.Cells(1, 15) = "追遡工站信息（工站代码：0320脱模）"
        Ws.Cells(2, 1) = "Station"
        Ws.Cells(2, 7) = "Station"
        Ws.Cells(2, 11) = "Station"
        Ws.Cells(2, 15) = "Station"
        Ws.Cells(2, 2) = "SN"
        Ws.Cells(2, 3) = "Start time"
        Ws.Cells(2, 8) = "Start time"
        Ws.Cells(2, 12) = "Start time"
        Ws.Cells(2, 16) = "Start time"
        Ws.Cells(2, 4) = "Finish time"
        Ws.Cells(2, 9) = "Finish time"
        Ws.Cells(2, 13) = "Finish time"
        Ws.Cells(2, 17) = "Finish time"
        Ws.Cells(2, 5) = "Operator1"
        Ws.Cells(2, 10) = "Operator1"
        Ws.Cells(2, 14) = "Operator1"
        Ws.Cells(2, 18) = "Operator1"
        Ws.Cells(2, 6) = "Scrap"
        Ws.Cells(3, 1) = "工站"
        Ws.Cells(3, 7) = "工站"
        Ws.Cells(3, 11) = "工站"
        Ws.Cells(3, 15) = "工站"
        Ws.Cells(3, 2) = "系列号"
        Ws.Cells(3, 3) = "开始时间"
        Ws.Cells(3, 8) = "开始时间"
        Ws.Cells(3, 12) = "开始时间"
        Ws.Cells(3, 16) = "开始时间"
        Ws.Cells(3, 4) = "结束时间"
        Ws.Cells(3, 9) = "结束时间"
        Ws.Cells(3, 13) = "结束时间"
        Ws.Cells(3, 17) = "结束时间"
        Ws.Cells(3, 5) = "作业员"
        Ws.Cells(3, 10) = "作业员"
        Ws.Cells(3, 14) = "作业员"
        Ws.Cells(3, 18) = "作业员"
        Ws.Cells(3, 6) = "报废项"
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Scrap0330_Information"
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

End Class