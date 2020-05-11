Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form39
    Dim mConnection As New SqlClient.SqlConnection
    Dim mConnection2 As New SqlClient.SqlConnection
    Dim mConnection3 As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLS3 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim mSQLReader2 As SqlClient.SqlDataReader
    Dim mSQLReader3 As SqlClient.SqlDataReader
    Dim tStation1 As String
    Dim tStation2 As String
    Dim tDefect_Code As String
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

    Private Sub Form39_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        mConnection2.ConnectionString = Module1.OpenConnectionOfMes()
        mConnection3.ConnectionString = Module1.OpenConnectionOfMes()
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mConnection2.Open()
                mSQLS2.Connection = mConnection2
                mSQLS2.CommandType = CommandType.Text
                mConnection3.Open()
                mSQLS3.Connection = mConnection3
                mSQLS3.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        BindModel()
    End Sub
    Private Sub BindModel()
        Me.CheckedListBox1.Items.Clear()
        mSQLS1.CommandText = "select model from model order by model"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.CheckedListBox1.Items.Add(mSQLReader.Item("model"), False)
            End While
        End If
        mSQLReader.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If Me.CheckedListBox1.CheckedItems.Count = 0 Then
            MsgBox("请选择型号")
            Return
        End If
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mConnection2.Open()
                mSQLS2.Connection = mConnection2
                mSQLS2.CommandType = CommandType.Text
                mConnection3.Open()
                mSQLS3.Connection = mConnection3
                mSQLS3.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        ExtendCon = "'"
        For i As Integer = 0 To Me.CheckedListBox1.CheckedItems.Count - 1 Step 1
            If i = 0 Then
                ExtendCon += Me.CheckedListBox1.CheckedItems(i).ToString()
                ExtendCon += "'"
            Else
                ExtendCon += ",'"
                ExtendCon += Me.CheckedListBox1.CheckedItems(i).ToString()
                ExtendCon += "'"
            End If
        Next
        CheckTimes = Me.NumericUpDown1.Value
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        mSQLS1.CommandText = "SELECT COUNT(SN) AS T2 FROM ( "
        mSQLS1.CommandText += "select sum(t1) AS T1,sn,updatedstation  from ( "
        mSQLS1.CommandText += "select count(sn.sn) as t1,sn.sn,sn.updatedstation  from sn "
        mSQLS1.CommandText += "left join tracking on sn.sn = tracking.sn and tracking.station = '0590' "
        mSQLS1.CommandText += "right join lot on sn.lot = lot.lot and lot.model in ("
        mSQLS1.CommandText += ExtendCon & ") where updatedstation <> '9999' group by sn.sn,sn.updatedstation "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select count(sn.sn),sn.sn,sn.updatedstation from sn "
        mSQLS1.CommandText += "left join tracking_dup on sn.sn = tracking_dup.sn and tracking_dup.station = '0590' "
        mSQLS1.CommandText += "right join lot on sn.lot = lot.lot and lot.model in ("
        mSQLS1.CommandText += ExtendCon & ") where updatedstation <> '9999' group by sn.sn,sn.updatedstation "
        mSQLS1.CommandText += ") as AA group by sn,updatedstation having sum(t1) > " & CheckTimes & ") AS BB"

        Dim HaveReport As Integer = mSQLS1.ExecuteScalar()
        If HaveReport = 0 Then
            MsgBox("没有资料，请重选条件")
            Return
        End If
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        LineZ = 2
        mSQLS1.CommandText = "select sum(t1) AS T1,sn,updatedstation,lasttimeout from ( "
        mSQLS1.CommandText += "select count(sn.sn) as t1,sn.sn,sn.updatedstation,sn.lasttimeout  from sn "
        mSQLS1.CommandText += "left join tracking on sn.sn = tracking.sn and tracking.station = '0590' "
        mSQLS1.CommandText += "right join lot on sn.lot = lot.lot and lot.model in ("
        mSQLS1.CommandText += ExtendCon & ") where updatedstation <> '9999' group by sn.sn,sn.updatedstation,sn.lasttimeout "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select count(sn.sn),sn.sn,sn.updatedstation,sn.lasttimeout from sn "
        mSQLS1.CommandText += "left join tracking_dup on sn.sn = tracking_dup.sn and tracking_dup.station = '0590' "
        mSQLS1.CommandText += "right join lot on sn.lot = lot.lot and lot.model in ("
        mSQLS1.CommandText += ExtendCon & ") where updatedstation <> '9999' group by sn.sn,sn.updatedstation,sn.lasttimeout "
        mSQLS1.CommandText += ") as AA group by sn,updatedstation,lasttimeout having sum(t1) > " & CheckTimes & " order by sn"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("sn")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("t1")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("lasttimeout").ToString()
                Ws.Cells(LineZ, 4) = mSQLReader.Item("updatedstation").ToString()
                LineZ += 1
            End While
        End If
        mSQLReader.Close()

    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 50
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Cells(1, 1) = "产品序列号"
        Ws.Cells(1, 2) = "经过0590次数"
        Ws.Cells(1, 3) = "最后一次的时间"
        Ws.Cells(1, 4) = "目前所在工站"
        oRng = Ws.Range("D1", "D1")
        oRng.EntireColumn.NumberFormat = "@"
    End Sub
    
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "PaintPassTime_Report"
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