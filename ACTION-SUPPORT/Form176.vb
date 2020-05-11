Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form176
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim tModel As String
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim HaveReport As Integer = 0
    Dim l_sn As String = String.Empty
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form176_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
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
        
        BindModel()

    End Sub
    Private Sub BindModel()
        Me.ComboBox2.Items.Clear()
        mSQLS1.CommandText = "select distinct lot.model,model.modelname  from lot,model " _
                          & " where lot.model = model.model and model.model_type <> 'Action'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox2.Items.Add(mSQLReader.Item(0).ToString() & "|" & mSQLReader.Item(1).ToString())
            End While
        End If
        mSQLReader.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        tModel = String.Empty
        If Not IsNothing(ComboBox2.SelectedItem) Then
            tModel = ComboBox2.SelectedItem.ToString()
            Dim stCount As Int16 = Strings.InStr(tModel, "|")
            If stCount > 0 Then
                tModel = Strings.Left(tModel, stCount - 1)
            End If
        End If

        TimeS1 = DateTimePicker1.Value
        TimeS2 = DateTimePicker2.Value
        HaveReport = 0
        l_sn = String.Empty
        If Not IsDBNull(TextBox1.Text) Then
            l_sn = TextBox1.Text
        End If

        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        mSQLS1.CommandText = "select count(sn) from EventLog where Process = 'CONCESSION'"
        If Me.GroupBox2.Enabled = True Then
            mSQLS1.CommandText += "and UpdateOn between '"
            mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        End If
        If Not String.IsNullOrEmpty(l_sn) Then
            mSQLS1.CommandText += " AND sn like '%" & l_sn & "%' "
        End If
        If Not String.IsNullOrEmpty(tModel) Then
            mSQLS1.CommandText += " AND Model = '" & tModel & "' "
        End If
        HaveReport = mSQLS1.ExecuteScalar()
        If HaveReport > 0 Then
            xExcel = New Microsoft.Office.Interop.Excel.Application
            xWorkBook = xExcel.Workbooks.Add()
            Ws = xWorkBook.Sheets(1)
            AdjustExcelFormat()
            mSQLS1.CommandText = "select model, sn,  e1.station, s1.stationname_cn , SUBSTRING(event,9, 4) as c2, s2.stationname_cn , UpdateOn , UpdatedBy , users.name "
            mSQLS1.CommandText += "from EventLog e1 left join station s1 on e1.station = s1.station left join station s2 on SUBSTRING(e1.event, 9, 4) = s2.station left join users on UpdatedBy = users.id "
            mSQLS1.CommandText += "where Process = 'CONCESSION' "
            If Me.GroupBox2.Enabled = True Then
                mSQLS1.CommandText += "and UpdateOn between '"
                mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
            End If
            If Not String.IsNullOrEmpty(l_sn) Then
                mSQLS1.CommandText += " AND sn like '%" & l_sn & "%' "
            End If
            If Not String.IsNullOrEmpty(tModel) Then
                mSQLS1.CommandText += " AND Model = '" & tModel & "' "
            End If
            mSQLReader = mSQLS1.ExecuteReader()
            If mSQLReader.HasRows() Then
                While mSQLReader.Read()
                    For i As Int16 = 0 To mSQLReader.FieldCount - 1 Step 1
                        Ws.Cells(LineZ, i + 1) = mSQLReader.Item(i)
                    Next
                    LineZ += 1
                End While
            End If
            mSQLReader.Close()
        Else
            MsgBox("No Data")
        End If
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 23
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.WrapText = True

        Ws.Cells(1, 1) = "产品型号"
        Ws.Cells(1, 2) = "SN序列号"
        Ws.Cells(1, 3) = "移出工站"
        Ws.Cells(1, 4) = "移出工站名称"
        Ws.Cells(1, 5) = "移入工站"
        Ws.Cells(1, 6) = "移入工站名称"
        Ws.Cells(1, 7) = "移站时间"
        Ws.Cells(1, 8) = "处理人员工号"
        Ws.Cells(1, 9) = "处理人员姓名"
        oRng = Ws.Range("E1", "E1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        oRng = Ws.Range("H1", "H1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        LineZ = 2
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If HaveReport > 0 Then
            SaveExcel()
        End If
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "MES_Concession_LIST"
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