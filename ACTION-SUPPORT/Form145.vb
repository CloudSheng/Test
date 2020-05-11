Public Class Form145
    'Dim ModelList() As String = {}
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
    Dim CheckTimes As Integer = 0
    Dim ExtendCon As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        TimeS1 = DateTimePicker1.Value
        TimeS2 = DateTimePicker2.Value
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        BackgroundWorker1.RunWorkerAsync()
        'Array.Clear(ModelList, 0, ModelList.Count - 1)
        'For i As Int16 = 0 To RichTextBox1.Lines.Count - 1 Step 1
        'MsgBox(RichTextBox1.Lines(i))
        'Next
    End Sub

    Private Sub Form145_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        DateTimePicker1.Value = Convert.ToDateTime(Now.ToString("yyyy/MM/dd") & " 00:00:00")
        DateTimePicker2.Value = Convert.ToDateTime(Now.ToString("yyyy/MM/dd") & " 00:00:00")
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
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()

        For i As Int16 = 0 To RichTextBox1.Lines.Count - 1 Step 1
            If i > 2 Then
                Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
            Else
                Ws = xWorkBook.Sheets(i + 1)
            End If
            Ws.Activate()
            AdjustExcelFormat()
            Ws.Name = RichTextBox1.Lines(i)
            mSQLS1.CommandText = "select sn,count(sn) from ( "
            mSQLS1.CommandText += "select sn from tracking,lot where tracking.lot = lot.lot and lot.model in ('" & RichTextBox1.Lines(i) & "') AND tracking.station = '0590' and tracking.timeout between '"
            mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
            mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
            mSQLS1.CommandText += "union all "
            mSQLS1.CommandText += "select sn from tracking_dup,lot where tracking_dup.lot = lot.lot and lot.model in ('" & RichTextBox1.Lines(i) & "') AND tracking_dup.station = '0590' and tracking_dup.timeout between '"
            mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
            mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
            mSQLS1.CommandText += "union all "
            mSQLS1.CommandText += "select sn from scrap_tracking,lot where scrap_tracking.lot = lot.lot and lot.model in ('" & RichTextBox1.Lines(i) & "') AND scrap_tracking.station = '0590' and scrap_tracking.timeout between '"
            mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
            mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' ) as ab group by sn order by sn"

            mSQLReader = mSQLS1.ExecuteReader()
            If mSQLReader.HasRows() Then
                While mSQLReader.Read()
                    For j As Int16 = 0 To mSQLReader.FieldCount - 1 Step 1
                        Ws.Cells(LineZ, j + 1) = mSQLReader.Item(j)
                    Next
                    LineZ += 1
                End While
            End If
            mSQLReader.Close()
        Next

    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 50
        Ws.Cells(1, 1) = "产品序列号"
        Ws.Cells(1, 2) = "经过0590次数"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormat = "@"
        LineZ = 2
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "PaintPassTime_Report"
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

End Class