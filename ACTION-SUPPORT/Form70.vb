Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form70
    Dim kConnection As New SqlClient.SqlConnection
    Dim kCommander As New SqlClient.SqlCommand
    Dim kReader As SqlClient.SqlDataReader
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim tDate As Date
    Dim tProject As String = String.Empty
    Dim tUser As String = String.Empty
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Me.GroupBox2.Enabled = True
        Else
            Me.GroupBox2.Enabled = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            Me.GroupBox3.Enabled = True
        Else
            Me.GroupBox3.Enabled = False
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            Me.GroupBox4.Enabled = True
        Else
            Me.GroupBox4.Enabled = False
        End If
    End Sub
    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            Me.GroupBox5.Enabled = True
        Else
            Me.GroupBox5.Enabled = False
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        If kConnection.State <> ConnectionState.Open Then
            kConnection.Open()
            kCommander.Connection = kConnection
            kCommander.CommandType = CommandType.Text
        End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Form70_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        kConnection.ConnectionString = Module1.OpenConnectionOfERPSUPPORT()
        If kConnection.State <> ConnectionState.Open Then
            kConnection.Open()
            kCommander.Connection = kConnection
            kCommander.CommandType = CommandType.Text
        End If
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Project_WorkHour_Report"
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
        If kConnection.State = ConnectionState.Open Then
            Try
                kConnection.Close()
                oConnection.Close()
                Module1.KillExcelProcess(OldExcel)
                MsgBox("Finished")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        AdjustExcelFormat1()
        kCommander.CommandText = "SELECT * FROM ERPSUPPORT.dbo.ProjectHR WHERE 1 =1"
        If Me.GroupBox2.Enabled = True Then
            kCommander.CommandText += " AND eDate between '" & Me.DateTimePicker1.Value.ToString("yyyy/MM/dd") & "' AND '" & Me.DateTimePicker4.Value.ToString("yyyy/MM/dd") & "' "
        End If
        If Me.GroupBox3.Enabled = True Then
            kCommander.CommandText += " AND EProject = '" & Me.TextBox1.Text & "' "
        End If
        If Me.GroupBox4.Enabled = True Then
            kCommander.CommandText += " AND eUser = '" & Me.TextBox2.Text & "' "
        End If
        'If Me.GroupBox5.Enabled = True Then
        'kCommander.CommandText += " AND RecordTime between '" & Me.DateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '" & Me.DateTimePicker3.Value.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        'End If
        If Me.GroupBox5.Enabled = True Then
            kCommander.CommandText += " AND EDepartNo LIKE '%" & Me.TextBox5.Text & "%' "
        End If
        If Me.GroupBox6.Enabled = True Then
            kCommander.CommandText += " AND EAP = " & Me.TextBox3.Text & "' "
        End If
        If Me.GroupBox7.Enabled = True Then
            kCommander.CommandText += " AND ModelID LIKE '%" & Me.TextBox4.Text & "%' "
        End If
        kReader = kCommander.ExecuteReader()
        If kReader.HasRows() Then
            While kReader.Read()
                Ws.Cells(LineZ, 1) = kReader.Item("EUserName")
                Ws.Cells(LineZ, 2) = kReader.Item("EProject")
                Ws.Cells(LineZ, 3) = kReader.Item("EDepartNo")
                Ws.Cells(LineZ, 4) = kReader.Item("EUser")
                Ws.Cells(LineZ, 5) = kReader.Item("Edate")
                Ws.Cells(LineZ, 6) = kReader.Item("ModelID")
                Ws.Cells(LineZ, 7) = kReader.Item("EAP")
                Ws.Cells(LineZ, 8) = GetNameofProject(kReader.Item("EProject"))
                Ws.Cells(LineZ, 9) = kReader.Item("WorkDesc")
                Ws.Cells(LineZ, 10) = kReader.Item("EHour")
                Ws.Cells(LineZ, 11) = kReader.Item("Remark")
                LineZ += 1
            End While
        End If
        kReader.Close()
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 12

        oRng = Ws.Range("H1", "H1")
        oRng.EntireColumn.ColumnWidth = 60
        oRng = Ws.Range("J1", "J1")
        oRng.EntireColumn.ColumnWidth = 20

        Ws.Cells(1, 1) = "姓名"
        Ws.Cells(1, 2) = "项目专案号"
        Ws.Cells(1, 3) = "部门代码"
        Ws.Cells(1, 4) = "工号"
        Ws.Cells(1, 5) = "日期"
        Ws.Cells(1, 6) = "产品编号"
        Ws.Cells(1, 7) = "APQP阶段"
        Ws.Cells(1, 8) = "专案名称"
        Ws.Cells(1, 9) = "工作内容"
        Ws.Cells(1, 10) = "实际工时（H）"
        Ws.Cells(1, 11) = "备注（ECR/ECN）"
        LineZ = 2
    End Sub
    Private Function GetNameofProject(ByVal pja01 As String)
        
        
        oCommand.CommandText = "SELECT pja02 FROM pja_file WHERE pja01 = '" & pja01 & "' "
        Dim NameC1 As String = oCommand.ExecuteScalar()
        Return NameC1
    End Function

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            Me.GroupBox6.Enabled = True
        Else
            Me.GroupBox6.Enabled = False
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            Me.GroupBox7.Enabled = True
        Else
            Me.GroupBox7.Enabled = False
        End If
    End Sub
End Class