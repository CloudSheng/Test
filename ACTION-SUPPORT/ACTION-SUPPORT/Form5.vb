Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form5
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim tModel_type As String
    Dim tModel As String
    Dim tLot As String
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim tStation1 As String
    Dim tPN As String
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        BindModel_Station()
        BindModel()
    End Sub
    Private Sub BindModel()
        Me.ComboBox4.Items.Clear()
        mSQLS1.CommandText = "select distinct lot.model,model.modelname  from lot,model " _
                          & " where lot.model = model.model and model.model_type <> 'Action'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox4.Items.Add(mSQLReader.Item(0).ToString() & "," & mSQLReader.Item(1).ToString())
            End While
        End If
        Me.ComboBox4.Items.Add("ALL,ALL")
        mSQLReader.Close()
    End Sub
    Private Sub BindModel_Station()
        Me.ComboBox1.Items.Clear()
        mSQLS1.CommandText = "SELECT station,stationname FROM station "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox1.Items.Add(mSQLReader.Item(0).ToString() & "," & mSQLReader.Item(1).ToString())
            End While
        End If
        Me.ComboBox1.Items.Add("ALL,ALL")
        mSQLReader.Close()
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
                'mConnection2.Open()
                'mSQLS2.Connection = mConnection2
                'mSQLS2.CommandType = CommandType.Text

            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        TimeS1 = Now()
        TimeS2 = TimeS1.AddDays(Me.NumericUpDown1.Value * Decimal.MinusOne)

        'MsgBox(TimeS1.ToString("yyyy/MM/dd HH:mm:ss"))
        If Not IsNothing(ComboBox1.SelectedItem) Then
            tStation1 = ComboBox1.SelectedItem.ToString()
            Dim stCount As Int16 = Strings.InStr(tStation1, ",")
            If stCount > 0 Then
                tStation1 = Strings.Left(tStation1, stCount - 1)
            End If
        End If
        If Not IsNothing(ComboBox4.SelectedItem) Then
            tPN = ComboBox4.SelectedItem.ToString()
            Dim stCount As Int16 = Strings.InStr(tPN, ",")
            If stCount > 0 Then
                tPN = Strings.Left(tPN, stCount - 1)
            End If
        End If
        If String.IsNullOrEmpty(tStation1) Then
            tStation1 = "ALL"
        End If
        If String.IsNullOrEmpty(tPN) Then
            tPN = "ALL"
        End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Dull_Report"
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
    Private Sub ExportToExcel()
        mSQLS1.CommandText = "select Count(*) FROM lot l JOIN model m ON l.model=m.model JOIN sn s ON l.lot=s.lot "
        mSQLS1.CommandText += "JOIN station t ON t.station=case when (s.topreworkstation is not null and s.topreworkstation<>'') then s.topreworkstation else s.updatedstation end "
        mSQLS1.CommandText += "WHERE updatedtime < '" & TimeS2.ToString("yyyy/MM/dd hh:mm:ss") & "' "
        If Not tStation1 = "ALL" Then
            mSQLS1.CommandText += " AND t.station = '" & tStation1 & "' "
        End If
        If Not tPN = "ALL" Then
            mSQLS1.CommandText += " AND l.lot = '" & tPN & "' "
        End If
        Dim HaveReport As Integer = mSQLS1.ExecuteScalar()
        If HaveReport = 0 Then
            MsgBox("没有资料，请重选条件")
            Return
        End If
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        LineZ = 3
        mSQLS1.CommandText = "select m.model,s.sn,t.stationname_cn  ,s.updatedtime FROM lot l JOIN model m ON l.model=m.model JOIN sn s ON l.lot=s.lot "
        mSQLS1.CommandText += "JOIN station t ON t.station=case when (s.topreworkstation is not null and s.topreworkstation<>'') then s.topreworkstation else s.updatedstation end "
        mSQLS1.CommandText += "WHERE updatedtime < '" & TimeS2.ToString("yyyy/MM/dd hh:mm:ss") & "' "
        If Not tStation1 = "ALL" Then
            mSQLS1.CommandText += " AND t.station = '" & tStation1 & "' "
        End If
        If Not tPN = "ALL" Then
            mSQLS1.CommandText += " AND l.lot = '" & tPN & "' "
        End If
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("sn")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("stationname_cn")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("updatedtime")
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 20
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("D1", "D1")
        oRng.EntireColumn.ColumnWidth = 47
        Ws.Cells(1, 1) = "查询时间"
        Ws.Cells(1, 2) = TimeS1
        Ws.Cells(1, 3) = "呆滞天数"
        Ws.Cells(1, 4) = Me.NumericUpDown1.Value
        Ws.Cells(2, 1) = "产品型号 P/N"
        Ws.Cells(2, 2) = "产品序列号 S/N"
        Ws.Cells(2, 3) = "工段 Section"
        Ws.Cells(2, 4) = "上一工段完成时间 Last section finished time"
    End Sub
End Class