Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form120
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim mSQLReader2 As SqlClient.SqlDataReader
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim ptime As String = String.Empty
    Dim MaxDetailCount As Int16 = 0
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim HaveReport As Integer = 0
    Dim l_sn As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form120_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        HaveReport = 0
        l_sn = String.Empty
        If Not IsDBNull(TextBox1.Text) Then
            l_sn = TextBox1.Text
        End If
        BackgroundWorker1.RunWorkerAsync()
        'ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        mSQLS1.CommandText = "select count(sn) from sn_log where event = 'Reprint' "
        If Me.GroupBox2.Enabled = True Then
            mSQLS1.CommandText += "and logtime between '"
            mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        End If
        If Not String.IsNullOrEmpty(l_sn) Then
            mSQLS1.CommandText += " AND sn like '%" & l_sn & "%' "
        End If
        HaveReport = mSQLS1.ExecuteScalar()
        If HaveReport > 0 Then
            xExcel = New Microsoft.Office.Interop.Excel.Application
            xWorkBook = xExcel.Workbooks.Add()
            Ws = xWorkBook.Sheets(1)
            AdjustExcelFormat()
            mSQLS1.CommandText = "select *,CONVERT(varchar(100), logtime, 11) as t1,CONVERT(varchar(100), logtime, 24) as t2 from sn_log where event = 'Reprint' "
            If Me.GroupBox2.Enabled = True Then
                mSQLS1.CommandText += "and logtime between '"
                mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
            End If
            If Not String.IsNullOrEmpty(l_sn) Then
                mSQLS1.CommandText += " AND sn like '%" & l_sn & "%' "
            End If
            mSQLReader = mSQLS1.ExecuteReader()
            If mSQLReader.HasRows() Then
                While mSQLReader.Read()
                    Ws.Cells(LineZ, 1) = mSQLReader.Item("sn")
                    Ws.Cells(LineZ, 2) = mSQLReader.Item("lot")
                    Ws.Cells(LineZ, 3) = mSQLReader.Item("station")
                    Ws.Cells(LineZ, 4) = mSQLReader.Item("event")
                    Ws.Cells(LineZ, 5) = mSQLReader.Item("message")
                    Ws.Cells(LineZ, 6) = mSQLReader.Item("t1")
                    Ws.Cells(LineZ, 7) = mSQLReader.Item("t2")
                    Ws.Cells(LineZ, 8) = mSQLReader.Item("users")
                    LineZ += 1
                End While
            End If
            mSQLReader.Close()
        End If
        
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 23
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.WrapText = True
        
        Ws.Cells(1, 1) = "SN序列号"
        Ws.Cells(1, 2) = "生产制令"
        Ws.Cells(1, 3) = "重印工站"
        Ws.Cells(1, 4) = "事项"
        Ws.Cells(1, 5) = "型号"
        Ws.Cells(1, 6) = "重新列印年月日"
        Ws.Cells(1, 7) = "重新列印时间"
        Ws.Cells(1, 8) = "列印人员"
        oRng = Ws.Range("E1", "E1")
        oRng.EntireColumn.ColumnWidth = 32.26
        oRng = Ws.Range("H1", "H1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        LineZ = 2
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If HaveReport > 0 Then
            SaveExcel()
        End If
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "MES_REPRINT_LIST"
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

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Me.GroupBox2.Enabled = True
        Else
            Me.GroupBox2.Enabled = False
        End If
    End Sub
End Class