Public Class Form182
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim StartDate As Date
    Dim EndDate As Date
    Dim DateCount As Int16
    Dim ReportType As String = String.Empty
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form182_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Dim xPath As String = "C:\temp\原材料进料计划达成Template.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommand2.Connection = oConnection
                oCommand2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        StartDate = DateTimePicker1.Value
        If Me.RadioButton1.Checked = True Then
            DateCount = 6
            ReportType = "周报"
        Else
            DateCount = 30
            ReportType = "月报"
        End If
        EndDate = StartDate.AddDays(DateCount)
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = StartDate.ToString("MMdd") & ReportType
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
        Dim xPath As String = "C:\temp\原材料进料计划达成Template.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        LineZ = 4
        ' 先填入日期
        For i As Int16 = 0 To DateCount Step 1
            Ws.Cells(1, 14 + i * 3) = StartDate.AddDays(i).ToString("yyyy/MM/dd")
        Next
        ' 建立 SQL
        oCommand.CommandText = "Select tc_prp01,ima02,ima021,ima25,ima54"
        For i As Int16 = 1 To DateCount + 1 Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        For i As Int16 = 1 To DateCount + 1 Step 1
            oCommand.CommandText += ",sum(r" & i & ") as r" & i
        Next
        oCommand.CommandText += " from ( "
        oCommand.CommandText += "Select tc_prp01,ima02,ima021,ima25,ima54"
        For i As Int16 = 1 To DateCount + 1 Step 1
            oCommand.CommandText += ",sum(case when tc_prp05 = to_date('" & StartDate.AddDays(i - 1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then tc_prp04 else 0 end) as t" & i
        Next
        For i As Int16 = 1 To DateCount + 1 Step 1
            oCommand.CommandText += ",0 as r" & i
        Next
        oCommand.CommandText += " from tc_prp_file left join ima_file on tc_prp01 = ima01 Where tc_prp05 between to_date('"
        oCommand.CommandText += StartDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += EndDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') group by tc_prp01,ima02,ima021,ima25,ima54 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "Select rvb05,ima02,ima021,ima25,ima54"
        For i As Int16 = 1 To DateCount + 1 Step 1
            oCommand.CommandText += ",0"
        Next
        For i As Int16 = 1 To DateCount + 1 Step 1
            oCommand.CommandText += ",sum(case when rva06 = to_date('" & StartDate.AddDays(i - 1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then rvb87 else 0 end)"
        Next
        oCommand.CommandText += " from rvb_file, rva_file,ima_file where rvb01 = rva01 and rvaconf = 'Y' and rvb05 = ima01 and rvb36 in ('D146101','D146102','D146108') "
        oCommand.CommandText += "and rva06 between to_date('" & StartDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += EndDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') group by rvb05,ima02,ima021,ima25,ima54 ) group by tc_prp01,ima02,ima021,ima25,ima54 order by tc_prp01 "

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item(0)
                Ws.Cells(LineZ, 2) = oReader.Item(1)
                Ws.Cells(LineZ, 3) = oReader.Item(2)
                Ws.Cells(LineZ, 4) = oReader.Item(3)
                Ws.Cells(LineZ, 5) = oReader.Item(4)
                For i As Int16 = 0 To DateCount Step 1
                    If oReader.Item(i + 5) <> 0 Then
                        Ws.Cells(LineZ, 14 + i * 3) = oReader.Item(i + 5)
                    End If
                    If oReader.Item(i + 6 + DateCount) <> 0 Then
                        Ws.Cells(LineZ, 15 + i * 3) = oReader.Item(i + 6 + DateCount)
                    End If
                Next
                LineZ += 1
                Label3.Text = LineZ
                Label3.Refresh()
            End While
        End If
        oReader.Close()
    End Sub
End Class