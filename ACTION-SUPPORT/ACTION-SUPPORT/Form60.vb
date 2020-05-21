Public Class Form60
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim TimeS1 As DateTime   'ERP 開始時間
    Dim TimeS2 As DateTime   'ERP 結束時間
    Dim TimeS3 As DateTime   'MES 開始時間
    Dim TimeS4 As DateTime   'MES 結束時間
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form60_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
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
        TimeS1 = DateTimePicker1.Value
        TimeS2 = DateTimePicker2.Value
        TimeS3 = DateTimePicker1.Value.ToString("yyyy/MM/dd 08:00:00")
        TimeS4 = TimeS3.AddDays(1).AddSeconds(-1)
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ImportData()
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub ImportData()
        Dim station As String = String.Empty
        Dim workgroup As String = String.Empty
        DropTable()
        CreateTable()
        oCommand.CommandText = "insert into aa_temp select sfv04,sfu04,sum(sfv09),0 from sfu_file,sfv_file where sfu01 = sfv01 and sfupost = 'Y' and "
        oCommand.CommandText += "sfu02 between to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += TimeS2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') group by sfv04,sfu04"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
            Return
        End Try
        For i As Int32 = 1 To 9 Step 1
            Select Case i
                Case 1
                    station = "'0112','0113'"
                    workgroup = "D3531"
                Case 2
                    station = "'0150','0151'"
                    workgroup = "D3532"
                Case 3
                    station = "'0330','0390'"
                    workgroup = "D3535"
                Case 4
                    station = "'0380','0520','0530'"
                    workgroup = "D3536"
                Case 5
                    station = "'0480','0490','0400'"
                    workgroup = "D3564"
                Case 6
                    station = "'0475'"
                    workgroup = "D3561"
                Case 7
                    station = "'0590'"
                    workgroup = "D3563"
                Case 8
                    station = "'0642'"
                    workgroup = "D3565"
                Case 9
                    station = "'0680'"
                    workgroup = "D3566"
            End Select
            mSQLS1.CommandText = "select count(cf01) as t1,cf01 from ( select cf01  from tracking left join lot on tracking.lot = lot.lot "
            mSQLS1.CommandText += "left join model_station_paravalue  on lot.model = model_station_paravalue.model  and model_station_paravalue.profilename = 'ERP' "
            mSQLS1.CommandText += "and model_station_paravalue.station = tracking.station where timeout between '" & TimeS3.ToString("yyyy/MM/dd HH:ss:mm") & "' and '"
            mSQLS1.CommandText += TimeS4.ToString("yyyy/MM/dd HH:ss:mm") & "' and tracking.station in (" & station & ") "
            mSQLS1.CommandText += "union all "
            mSQLS1.CommandText += "select cf01 from tracking_dup left join lot on tracking_dup.lot = lot.lot left join model_station_paravalue  on lot.model = model_station_paravalue.model  and model_station_paravalue.profilename = 'ERP' "
            mSQLS1.CommandText += "and model_station_paravalue.station = tracking_dup.station where timeout between '" & TimeS3.ToString("yyyy/MM/dd HH:ss:mm") & "' and '"
            mSQLS1.CommandText += TimeS4.ToString("yyyy/MM/dd HH:ss:mm") & "' and tracking_dup.station in (" & station & ") "
            mSQLS1.CommandText += "union all "
            mSQLS1.CommandText += "select cf01 from scrap_tracking left join lot on scrap_tracking.lot = lot.lot left join model_station_paravalue  on lot.model = model_station_paravalue.model  and model_station_paravalue.profilename = 'ERP' "
            mSQLS1.CommandText += "and model_station_paravalue.station = scrap_tracking.station where timeout between '" & TimeS3.ToString("yyyy/MM/dd HH:ss:mm") & "' and '"
            mSQLS1.CommandText += TimeS4.ToString("yyyy/MM/dd HH:ss:mm") & "' and scrap_tracking.station in (" & station & ") ) AS AA where cf01 is not null group by cf01"
            mSQLReader = mSQLS1.ExecuteReader()
            If mSQLReader.HasRows() Then
                While mSQLReader.Read()
                    oCommand.CommandText = "INSERT INTO bb_temp VALUES('" & mSQLReader.Item("cf01") & "','" & workgroup & "',0," & mSQLReader.Item("t1") & ")"
                    Try
                        oCommand.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.Message())
                        Exit While
                    End Try
                End While
            End If
            mSQLReader.Close()
        Next
        '處理完後

        oCommand.CommandText = "merge into aa_temp a1 USING bb_temp b1 ON (a1.ima01 = b1.ima01 AND a1.gem01 = b1.gem01 ) WHEN MATCHED THEN "
        oCommand.CommandText += "UPDATE SET a1.mesQ = b1.mesQ WHEN NOT MATCHED THEN INSERT (ima01,gem01,erpQ,mesQ) VALUES(b1.ima01,b1.gem01,b1.erpq,b1.mesQ)"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
            Return
        End Try
    End Sub
    Private Sub DropTable()
        oCommand.CommandText = "DROP TABLE aa_temp"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
        End Try
        oCommand.CommandText = "DROP TABLE bb_temp"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub CreateTable()
        oCommand.CommandText = "CREATE TABLE aa_temp (ima01 varchar2(40), gem01 varchar2(10), erpQ number(15,3), mesQ number(15,0)) "
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
            Return
        End Try
        oCommand.CommandText = "CREATE TABLE bb_temp (ima01 varchar2(40), gem01 varchar2(10), erpQ number(15,3), mesQ number(15,0)) "
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
            Return
        End Try
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        For i As Int16 = 4 To 9 Step 1
            xWorkBook.Sheets.Add()
        Next
        Dim workgroup As String = String.Empty
        Dim workgroupname As String = String.Empty
        For i As Int16 = 1 To 9 Step 1
            Select Case i
                Case 1
                    workgroup = "D3531"
                    workgroupname = "Cut"
                Case 2
                    workgroup = "D3532"
                    workgroupname = "Layup"
                Case 3
                    workgroup = "D3535"
                    workgroupname = "Mold"
                Case 4
                    workgroup = "D3536"
                    workgroupname = "CNC"
                Case 5
                    workgroup = "D3564"
                    workgroupname = "Glue"
                Case 6
                    workgroup = "D3561"
                    workgroupname = "Sanding"
                Case 7
                    workgroup = "D3563"
                    workgroupname = "Painting"
                Case 8
                    workgroup = "D3565"
                    workgroupname = "Polishing"
                Case 9
                    workgroup = "D3566"
                    workgroupname = "Package"
            End Select
            Ws = xWorkBook.Sheets(i)
            Ws.Activate()
            Ws.Name = workgroupname
            AdjustExcelFormat()
            oCommand.CommandText = "SELECT * FROM aa_temp WHERE gem01 = '" & workgroup & "' ORDER BY ima01"
            oReader = oCommand.ExecuteReader()
            If oReader.HasRows() Then
                While oReader.Read()
                    Ws.Cells(LineZ, 1) = oReader.Item("ima01")
                    Ws.Cells(LineZ, 2) = oReader.Item("erpQ")
                    Ws.Cells(LineZ, 3) = oReader.Item("mesQ")
                    LineZ += 1
                End While
            End If
            oReader.Close()
            '加總
            Ws.Cells(LineZ, 1) = "合计"
            Ws.Cells(LineZ, 2) = "=SUM(B3:B" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 3) = "=SUM(C3:C" & LineZ - 1 & ")"
        Next
        
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "C1")
        oRng.EntireColumn.ColumnWidth = 31.33
        Ws.Cells(1, 2) = "ERP日期范围：" & TimeS1.ToString("yyyy/MM/dd") & "-" & TimeS2.ToString("yyyy/MM/dd")
        Ws.Cells(1, 3) = "MES时间范围：" & TimeS3.ToString("yyyy/MM/dd HH:mm:ss") & "-" & TimeS4.ToString("yyyy/MM/dd HH:mm:ss")
        Ws.Cells(2, 1) = "料号"
        Ws.Cells(2, 2) = "ERP入库量"
        Ws.Cells(2, 3) = "MES移转量"
        LineZ = 3
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Check_ERP_MES"
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
End Class