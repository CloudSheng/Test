Public Class Form61
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim TimeS1 As DateTime   'MES 開始時間
    Dim TimeS2 As DateTime   'MES 結束時間
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form61_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
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
        'TimeS1 = DateTimePicker1.Value.ToString("yyyy/MM/01 08:00:00")
        'TimeS2 = TimeS1.AddMonths(1).AddSeconds(-1)
        TimeS1 = DateTimePicker1.Value
        TimeS2 = DateTimePicker2.Value
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub ExportToExcel()
        Dim model As String = String.Empty
        Dim station1 As String = String.Empty
        Dim station2 As String = String.Empty
        Dim station3 As String = String.Empty
        Dim LineS As Decimal = 0
        Dim S0 As Decimal = 0
        'Dim Formula As String = String.Empty
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        Ws.Name = "ScrapRate"
        AdjustExcelFormat()
        mSQLS1.CommandText = "select distinct model from ( select model from tracking left join lot on tracking.lot = lot.lot where tracking.timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and tracking.station in ('0330','0331','0380','0490','0410','0590','0640','0675') "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select model from tracking_dup left join lot on tracking_dup.lot = lot.lot where tracking_dup.timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and tracking_dup.station in ('0330','0331','0380','0490','0410','0590','0640','0675') "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select model from scrap_tracking left join lot on scrap_tracking.lot = lot.lot where scrap_tracking.timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and scrap_tracking.station in ('0330','0331','0380','0490','0410','0590','0640','0675') "
        mSQLS1.CommandText += ") AA order by model"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                For i As Integer = 1 To 7 Step 1
                    Select Case i
                        Case 1
                            station1 = "'0330','0331'"
                            station2 = station1
                            station3 = "成型"
                            LineS = 0
                        Case 2
                            station1 = "'0380'"
                            station2 = station1
                            station3 = "CNC"
                            LineS = 1
                        Case 3
                            station1 = "'0490'"
                            station2 = station1
                            station3 = "胶合"
                            LineS = 2
                        Case 4
                            station1 = "'0410'"
                            station2 = station1
                            station3 = "补土"
                            LineS = 3
                        Case 5
                            station1 = "'0590'"
                            station2 = station1
                            station3 = "涂装"
                            LineS = 4
                        Case 6
                            station1 = "'0640'"
                            station2 = "'0640','0642','0645','0657'"
                            station3 = "抛光"
                            LineS = 5
                        Case 7
                            station1 = "'0675'"
                            station2 = "'0650','0652','0665','0670','0675','0685'"
                            station3 = "包装"
                            LineS = 6
                    End Select
                    Ws.Cells(LineZ, 1) = LineZ - 1
                    Ws.Cells(LineZ, 2) = GetERPPN(mSQLReader.Item("model"))
                    Ws.Cells(LineZ, 3) = mSQLReader.Item("model")
                    Ws.Cells(LineZ, 4) = GetModelName(mSQLReader.Item("model"))
                    Ws.Cells(LineZ, 5) = station3
                    Dim S9 As Decimal = GetAllProductQuantity(station1, mSQLReader.Item("model"))
                    If station3 = "成型" Then
                        S0 = S9
                    End If
                    Ws.Cells(LineZ, 6) = S9
                    Ws.Cells(LineZ, 7) = GetScrapQuantity(station2, mSQLReader.Item("model"))
                    If S0 <> 0 Then
                        Ws.Cells(LineZ, 8) = "=G" & LineZ & "/F" & LineZ - LineS
                    Else
                        Ws.Cells(LineZ, 8) = 0
                    End If
                    LineZ += 1
                    Label2.Text = LineZ
                Next
                Ws.Cells(LineZ - 7, 9) = "=H" & LineZ - 7 & "+H" & LineZ - 6 & "+H" & LineZ - 5 & "+H" & LineZ - 4 & "+H" & LineZ - 3 & "+H" & LineZ - 2 & "+H" & LineZ - 1
                Ws.Cells(LineZ - 6, 9) = "=H" & LineZ - 6 & "+H" & LineZ - 5 & "+H" & LineZ - 4 & "+H" & LineZ - 3 & "+H" & LineZ - 2 & "+H" & LineZ - 1
                Ws.Cells(LineZ - 5, 9) = "=H" & LineZ - 5 & "+H" & LineZ - 4 & "+H" & LineZ - 3 & "+H" & LineZ - 2 & "+H" & LineZ - 1
                Ws.Cells(LineZ - 4, 9) = "=H" & LineZ - 4 & "+H" & LineZ - 3 & "+H" & LineZ - 2 & "+H" & LineZ - 1
                Ws.Cells(LineZ - 3, 9) = "=H" & LineZ - 3 & "+H" & LineZ - 2 & "+H" & LineZ - 1
                Ws.Cells(LineZ - 2, 9) = "=H" & LineZ - 2 & "+H" & LineZ - 1
                Ws.Cells(LineZ - 1, 9) = "=H" & LineZ - 1
            End While
        End If
        mSQLReader.Close()

    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "MonthScrapRate"
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

    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("B1", "D1")
        oRng.EntireColumn.ColumnWidth = 40
        oRng.Interior.Color = Color.Blue
        oRng = Ws.Range("E1", "I1")
        oRng.EntireColumn.ColumnWidth = 15
        Ws.Cells(1, 1) = "序号"
        Ws.Cells(1, 2) = "ERP 料号"
        Ws.Cells(1, 3) = "产品名称"
        Ws.Cells(1, 4) = "产品名称"
        Ws.Cells(1, 5) = "工站"
        Ws.Cells(1, 6) = "生产总数量"
        Ws.Cells(1, 7) = "报废数量"
        Ws.Cells(1, 8) = "报废率"
        Ws.Cells(1, 9) = "综合报废率"
        LineZ = 2
    End Sub
    Private Function GetAllProductQuantity(ByVal station As String, ByVal model As String)
        Dim mSQL99 As New SqlClient.SqlCommand
        mSQL99.Connection = mConnection
        mSQL99.CommandType = CommandType.Text
        mSQL99.CommandText = "select count(sn) from ( select sn from tracking left join lot on tracking.lot = lot.lot where tracking.timein between '"
        mSQL99.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and tracking.station in ("
        mSQL99.CommandText += station & ") and lot.model = '" & model & "' "
        mSQL99.CommandText += "union all "
        mSQL99.CommandText += "select sn from tracking_dup left join lot on tracking_dup.lot = lot.lot where tracking_dup.timein between '"
        mSQL99.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and tracking_dup.station in ("
        mSQL99.CommandText += station & ") and lot.model = '" & model & "' "
        mSQL99.CommandText += "union all "
        mSQL99.CommandText += "select sn from scrap_tracking left join lot on scrap_tracking.lot = lot.lot where scrap_tracking.timein between '"
        mSQL99.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and scrap_tracking.station in ("
        mSQL99.CommandText += station & ") and lot.model = '" & model & "' ) AS AD "
        Dim Q1 As Decimal = mSQL99.ExecuteScalar()
        Return Q1
    End Function
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
    Private Function GetScrapQuantity(ByVal station As String, ByVal model As String)
        Dim mSQL99 As New SqlClient.SqlCommand
        mSQL99.Connection = mConnection
        mSQL99.CommandType = CommandType.Text
        mSQL99.CommandText = "select count(scrap.sn) from scrap,scrap_sn,lot where scrap.sn = scrap_sn.sn and scrap.lot = lot.lot and scrap.datetime between '"
        mSQL99.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and scrap_sn.updatedstation in ("
        mSQL99.CommandText += station & ") and lot.model = '" & model & "'"
        Dim S1 As String = mSQL99.ExecuteScalar()
        Return S1
    End Function
End Class