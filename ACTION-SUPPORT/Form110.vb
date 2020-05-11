Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form110
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
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form110_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        mSQLS1.CommandText = "select eventlog.*, users.name from eventlog left join users on UpdatedBy = users.id where event = 'Restore' and UpdateOn between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("sn")
                Ws.Cells(LineZ, 6) = mSQLReader.Item("UpdatedBy") & " " & mSQLReader.Item("name")
                Ws.Cells(LineZ, 7) = mSQLReader.Item("UpdateOn")
                Ws.Cells(LineZ, 8) = mSQLReader.Item("remark")
                ' 讀取 失敗記錄
                GetMESFailure(mSQLReader.Item("sn"))
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 23.11
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.WrapText = True
        oRng = Ws.Range("A1", "I1")
        oRng.Merge()
        oRng = Ws.Range("A1", "I2")
        oRng.EntireRow.RowHeight = 42
        
        Ws.Cells(1, 1) = "报废解鎖日报表 Daliy Scrap Unlock Report"
        oRng = Ws.Range("A2", "A3")
        oRng.Merge()
        Ws.Cells(2, 1) = "產品序列號" & Chr(10) & "SN"
        oRng = Ws.Range("B2", "B3")
        oRng.Merge()
        Ws.Cells(2, 2) = "報廢工站" & Chr(10) & "Scrap Station"
        oRng = Ws.Range("C2", "C3")
        oRng.Merge()
        Ws.Cells(2, 3) = "報廢時間" & Chr(10) & "Scrap Timing"
        oRng = Ws.Range("D2", "D3")
        oRng.Merge()
        Ws.Cells(2, 4) = "缺陷原因" & Chr(10) & "Defect Code"
        oRng = Ws.Range("E2", "E3")
        oRng.Merge()
        Ws.Cells(2, 5) = "判定報廢檢驗員 ID" & Chr(10) & "Scraped by Whom"
        oRng = Ws.Range("F2", "F3")
        oRng.Merge()
        Ws.Cells(2, 6) = "解鎖人員 ID" & Chr(10) & "Unlock by whom"
        oRng = Ws.Range("G2", "G3")
        oRng.Merge()
        Ws.Cells(2, 7) = "解鎖時間" & Chr(10) & "Unlock timing"
        oRng = Ws.Range("H2", "H3")
        oRng.Merge()
        Ws.Cells(2, 8) = "解鎖原因代碼" & Chr(10) & "Reason for unlock"
        oRng = Ws.Range("I2", "I3")
        oRng.Merge()
        Ws.Cells(2, 9) = "備註" & Chr(10) & "Note"

        oRng = Ws.Range("B2", "B2")
        oRng.EntireColumn.NumberFormatLocal = "@"
        oRng = Ws.Range("D2", "F2")
        oRng.EntireColumn.NumberFormatLocal = "@"

        LineZ = 4
    End Sub
    Private Sub GetMESFailure(ByVal sn As String)
        mSQLS2.CommandText = "select failure.*,defect.desc_en,users.name from failure left join defect on failure.defect = defect.defect left join users on failure.users = users.id where sn = '" & sn & "' and rework = 'SCRP' order by failtime desc"
        mSQLReader2 = mSQLS2.ExecuteReader()
        If mSQLReader2.HasRows() Then
            mSQLReader2.Read()  '只讀一行
            Ws.Cells(LineZ, 2) = mSQLReader2.Item("failstation")
            Ws.Cells(LineZ, 3) = mSQLReader2.Item("failtime")
            Ws.Cells(LineZ, 4) = mSQLReader2.Item("defect") & " " & mSQLReader2.Item("desc_en")
            Ws.Cells(LineZ, 5) = mSQLReader2.Item("users") & " " & mSQLReader2.Item("name")
        End If
        mSQLReader2.Close()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        'If HaveReport > 0 Then
        SaveExcel()
        'End If
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Daily_Scrap_Unlock_Information"
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
End Class