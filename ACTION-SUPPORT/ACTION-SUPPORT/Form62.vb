Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form62
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim tModel As String
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form62_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        BindModel()
    End Sub
    Private Sub BindModel()
        Me.ComboBox1.Items.Clear()
        mSQLS1.CommandText = "select distinct lot.model,model.modelname  from lot,model " _
                          & " where lot.model = model.model and model.model_type <> 'Action'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox1.Items.Add(mSQLReader.Item(0).ToString())
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
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If Not IsNothing(ComboBox1.SelectedItem) Then
            tModel = ComboBox1.SelectedItem.ToString()
        End If
        TimeS1 = DateTimePicker1.Value
        TimeS2 = DateTimePicker2.Value
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Rework_Report"
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
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        mSQLS1.CommandText = "select model,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4 from ( "
        mSQLS1.CommandText += "select lot.model,count(sn) as t1,0 as t2,0 as t3,0 as t4 from tracking left join lot on tracking.lot = lot.lot where tracking.timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and station = '0590' group by lot.model "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select lot.model,count(sn),0,0,0 from scrap_tracking left join lot on scrap_tracking.lot = lot.lot where scrap_tracking.timein between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and station = '0590' group by lot.model "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select lot.model,0,count(sn),0,0 from tracking_dup left join lot on tracking_dup.lot = lot.lot where tracking_dup.timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and station = '0590' group by lot.model "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select lot.model,0,0,count(sn),0 from tracking left join lot on tracking.lot = lot.lot where tracking.timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and station = '0640' group by lot.model "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select lot.model,0,0,count(sn),0 from scrap_tracking left join lot on scrap_tracking.lot = lot.lot where scrap_tracking.timein between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and station = '0640' group by lot.model "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select lot.model,0,0,0,count(sn) from tracking_dup left join lot on tracking_dup.lot = lot.lot where tracking_dup.timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and station = '0640' group by lot.model"
        mSQLS1.CommandText += ") AS AD "
        If Not String.IsNullOrEmpty(tModel) Then
            mSQLS1.CommandText += " WHERE model = '" & tModel & "' "
        End If
        mSQLS1.CommandText += " group by model"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = LineZ - 3
                Ws.Cells(LineZ, 2) = GetERPPN(mSQLReader.Item("model"))
                Ws.Cells(LineZ, 3) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 4) = GetModelName(mSQLReader.Item("model"))
                Ws.Cells(LineZ, 5) = "=F" & LineZ & "+G" & LineZ
                If mSQLReader.Item("t1") <> 0 Then
                    Ws.Cells(LineZ, 6) = mSQLReader.Item("t1")
                End If
                If mSQLReader.Item("t2") <> 0 Then
                    Ws.Cells(LineZ, 7) = mSQLReader.Item("t2")
                End If
                Ws.Cells(LineZ, 8) = "=I" & LineZ & "+J" & LineZ
                If mSQLReader.Item("t3") <> 0 Then
                    Ws.Cells(LineZ, 9) = mSQLReader.Item("t3")
                End If
                If mSQLReader.Item("t4") <> 0 Then
                    Ws.Cells(LineZ, 10) = mSQLReader.Item("t4")
                End If
                If mSQLReader.Item("t1") + mSQLReader.Item("t2") <> 0 Then
                    Ws.Cells(LineZ, 11) = "=F" & LineZ & "/E" & LineZ
                    Ws.Cells(LineZ, 12) = "=G" & LineZ & "/E" & LineZ
                End If
                If mSQLReader.Item("t3") + mSQLReader.Item("t4") <> 0 Then
                    Ws.Cells(LineZ, 13) = "=I" & LineZ & "/H" & LineZ
                    Ws.Cells(LineZ, 14) = "=J" & LineZ & "/H" & LineZ
                End If
                Ws.Cells(LineZ, 15) = GetMoldTime(mSQLReader.Item("model"))
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
        'Total
        Ws.Cells(1, 5) = "=SUBTOTAL(9,E4:E" & LineZ - 1 & ")"
        Ws.Cells(1, 6) = "=SUBTOTAL(9,F4:F" & LineZ - 1 & ")"
        Ws.Cells(1, 7) = "=SUBTOTAL(9,G4:G" & LineZ - 1 & ")"
        Ws.Cells(1, 8) = "=SUBTOTAL(9,H4:H" & LineZ - 1 & ")"
        Ws.Cells(1, 9) = "=SUBTOTAL(9,I4:I" & LineZ - 1 & ")"
        Ws.Cells(1, 10) = "=SUBTOTAL(9,J4:J" & LineZ - 1 & ")"
        Ws.Cells(1, 11) = "=SUBTOTAL(9,K4:K" & LineZ - 1 & ")"
        Ws.Cells(1, 12) = "=SUBTOTAL(9,L4:L" & LineZ - 1 & ")"
        Ws.Cells(1, 13) = "=SUBTOTAL(9,M4:M" & LineZ - 1 & ")"
        Ws.Cells(1, 14) = "=SUBTOTAL(9,N4:N" & LineZ - 1 & ")"
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "O1")
        oRng.EntireColumn.ColumnWidth = 16.22
        oRng = Ws.Range("D1", "D1")
        oRng.EntireColumn.ColumnWidth = 28.33
        Ws.Cells(2, 1) = "序号"
        Ws.Cells(3, 1) = "NO"
        Ws.Cells(2, 2) = "ERP 料号"
        Ws.Cells(3, 2) = "ERP PN"
        Ws.Cells(2, 3) = "产品简称"
        Ws.Cells(3, 3) = "Product name"
        Ws.Cells(2, 4) = "产品名称"
        Ws.Cells(3, 4) = "WIP_Product Description"
        oRng = Ws.Range("E2", "G2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        oRng = Ws.Range("H2", "J2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        oRng = Ws.Range("K2", "L2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        oRng = Ws.Range("M2", "N2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(2, 5) = "涂装"
        Ws.Cells(2, 8) = "抛光"
        Ws.Cells(2, 11) = "涂装"
        Ws.Cells(2, 13) = "抛光"
        Ws.Cells(3, 5) = "合计"
        Ws.Cells(3, 8) = "合计"
        Ws.Cells(3, 6) = "第一次通过数量"
        Ws.Cells(3, 9) = "第一次通过数量"
        Ws.Cells(3, 7) = "返修数量"
        Ws.Cells(3, 10) = "返修数量"
        Ws.Cells(3, 11) = "一次百分比"
        Ws.Cells(3, 13) = "一次百分比"
        Ws.Cells(3, 12) = "返修百分比"
        Ws.Cells(3, 14) = "返修百分比"
        oRng = Ws.Range("K1", "N1")
        oRng.EntireColumn.NumberFormatLocal = "0.00%"
        Ws.Cells(2, 15) = "成型首次产出时间"
        LineZ = 4
    End Sub
    Private Function GetERPPN(ByVal model As String)
        Dim mSQL99 As New SqlClient.SqlCommand
        mSQL99.Connection = mConnection
        mSQL99.CommandType = CommandType.Text
        mSQL99.CommandText = "select value  from model_paravalue  where parameter = 'ERP PN' and model = '" & model & "'"
        Dim S1 As String = mSQL99.ExecuteScalar()
        Return S1
    End Function
    Private Function GetModelName(ByVal model As String)
        Dim mSQL99 As New SqlClient.SqlCommand
        mSQL99.Connection = mConnection
        mSQL99.CommandType = CommandType.Text
        mSQL99.CommandText = "select modelname  from model  where model = '" & model & "'"
        Dim S1 As String = mSQL99.ExecuteScalar()
        Return S1
    End Function
    Private Function GetMoldTime(ByVal model As String)
        Dim mSQL99 As New SqlClient.SqlCommand
        mSQL99.Connection = mConnection
        mSQL99.CommandType = CommandType.Text
        mSQL99.CommandText = "select top 1 CONVERT(varchar(100), timein, 111) from tracking right join lot on tracking.lot = lot.lot and lot.model = '" & model & "' where tracking.station in ('0330','0331') order by timein"
        Dim S1 As String = mSQL99.ExecuteScalar()
        Return S1
    End Function
End Class