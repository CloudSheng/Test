Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants

Public Class Form4
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
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        BindModel_Station()
        BindModel_DefectCode()
    End Sub
    Private Sub BindModel_Station()
        Me.ComboBox1.Items.Clear()
        Me.ComboBox2.Items.Clear()
        mSQLS1.CommandText = "SELECT station,stationname FROM station "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox1.Items.Add(mSQLReader.Item(0).ToString() & "," & mSQLReader.Item(1).ToString())
                Me.ComboBox2.Items.Add(mSQLReader.Item(0).ToString() & "," & mSQLReader.Item(1).ToString())
            End While
        End If
        Me.ComboBox2.Items.Add("ALL,ALL")
        mSQLReader.Close()
    End Sub
    Private Sub BindModel_DefectCode()
        Me.ComboBox3.Items.Clear()
        mSQLS1.CommandText = "SELECT defect,desc_th FROM defect "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox3.Items.Add(mSQLReader.Item(0).ToString() & "," & mSQLReader.Item(1).ToString())
            End While
        End If
        Me.ComboBox3.Items.Add("ALL,ALL")
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
        TimeS1 = DateTimePicker1.Value
        TimeS2 = DateTimePicker2.Value
        'MsgBox(TimeS1.ToString("yyyy/MM/dd HH:mm:ss"))
        If Not IsNothing(ComboBox1.SelectedItem) Then
            tStation1 = ComboBox1.SelectedItem.ToString()
            Dim stCount As Int16 = Strings.InStr(tStation1, ",")
            If stCount > 0 Then
                tStation1 = Strings.Left(tStation1, stCount - 1)
            End If
        End If
        If Not IsNothing(ComboBox2.SelectedItem) Then
            tStation2 = ComboBox2.SelectedItem.ToString()
            Dim stCount As Int16 = Strings.InStr(tStation2, ",")
            If stCount > 0 Then
                tStation2 = Strings.Left(tStation2, stCount - 1)
            End If
        End If
        If Not IsNothing(ComboBox3.SelectedItem) Then
            tDefect_Code = ComboBox3.SelectedItem.ToString()
            Dim stCount As Int16 = Strings.InStr(tDefect_Code, ",")
            If stCount > 0 Then
                tDefect_Code = Strings.Left(tDefect_Code, stCount - 1)
            End If
        End If
        If String.IsNullOrEmpty(tStation2) Then
            tStation2 = "ALL"
        End If
        If String.IsNullOrEmpty(tDefect_Code) Then
            tDefect_Code = "ALL"
        End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        mSQLS1.CommandText = "select count(*) as count1 from ("
        mSQLS1.CommandText += "SELECT tracking.station,station.stationname_cn,tracking.sn,lot.model,tracking.timein,tracking.timeout,"
        mSQLS1.CommandText += "tracking.users, users.name, station_type,failure.defect,defect.desc_th  "
        mSQLS1.CommandText += "FROM TRACKING  FULL JOIN STATION ON TRACKING.station = station.station "
        mSQLS1.CommandText += "full join users on tracking.users = users.id  full join lot  on tracking.lot = lot.lot "
        mSQLS1.CommandText += "LEFT JOIN FAILURE ON TRACKING.SN = FAILURE.SN AND TRACKING.station = FAILURE.failstation and abs(datediff(second,tracking.timeout,failure.failtime)) < 5 "
        mSQLS1.CommandText += "left join defect on failure.defect = defect.defect "
        mSQLS1.CommandText += "WHERE tracking.timein between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND tracking.station = '"
        mSQLS1.CommandText += tStation1 & "' "
        If tDefect_Code <> "ALL" Then
            mSQLS1.CommandText += " AND failure.defect  = '" & tDefect_Code & "' "
        End If
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "SELECT tracking_dup.station,station.stationname_cn,tracking_dup.sn,lot.model,tracking_dup.timein,tracking_dup.timeout,"
        mSQLS1.CommandText += "tracking_dup.users, users.name, station_type,failure.defect,defect.desc_th "
        mSQLS1.CommandText += "FROM TRACKING_dup  FULL JOIN STATION ON TRACKING_dup.station = station.station  full join users on tracking_dup.users = users.id "
        mSQLS1.CommandText += "full join lot  on tracking_dup.lot = lot.lot LEFT JOIN FAILURE ON TRACKING_dup.SN = FAILURE.SN AND TRACKING_dup.station = FAILURE.failstation "
        mSQLS1.CommandText += "and abs(datediff(second,tracking_dup.timeout,failure.failtime)) < 5 left join defect on failure.defect = defect.defect "
        mSQLS1.CommandText += "WHERE tracking_dup.timein between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND tracking_dup.station = '"
        mSQLS1.CommandText += tStation1 & "' "
        If tDefect_Code <> "ALL" Then
            mSQLS1.CommandText += " AND failure.defect  = '" & tDefect_Code & "' "
        End If
        mSQLS1.CommandText += ") as AA "
        mSQLS1.CommandTimeout = 300
        Dim HaveReport As Integer = mSQLS1.ExecuteScalar()
        Me.ProgressBar1.Maximum = HaveReport
        Me.ProgressBar1.Value = 0
        If HaveReport = 0 Then
            MsgBox("没有资料，请重选条件")
            Return
        End If
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        LineZ = 4

        'mSQLS1.CommandText = "SELECT station,stationname_cn,sn,model,timein,timeout,users,name,station_type FROM ( "
        'mSQLS1.CommandText += "SELECT tracking.station,station.stationname_cn,tracking.sn,lot.model,tracking.timein,tracking.timeout,"
        'mSQLS1.CommandText += "tracking.users, users.name, station_type "
        'mSQLS1.CommandText += "FROM tracking,station,users,lot WHERE tracking.station = station.station and tracking.users = users.id and tracking.lot = lot.lot "
        'mSQLS1.CommandText += "and tracking.timein between '"
        'mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        'mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND tracking.station = '"
        'mSQLS1.CommandText += tStation1 & "' "
        'mSQLS1.CommandText += "union all "
        'mSQLS1.CommandText += "SELECT tracking_dup.station,station.stationname_cn,tracking_dup.sn,lot.model,tracking_dup.timein,tracking_dup.timeout,"
        'mSQLS1.CommandText += "tracking_dup.users, users.name, station_type "
        'mSQLS1.CommandText += "FROM tracking_dup,station,users,lot where tracking_dup.station = station.station and tracking_dup.users = users.id and tracking_dup.lot = lot.lot "
        'mSQLS1.CommandText += "and tracking_dup.timein between '"
        'mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        'mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND tracking_dup.station = '"
        'mSQLS1.CommandText += tStation1 & "') as BB order by sn"

        mSQLS1.CommandText = "SELECT station,stationname_cn,sn,model,timein,timeout,users,name,station_type,defect,desc_th FROM ( "
        mSQLS1.CommandText += "SELECT tracking.station,station.stationname_cn,tracking.sn,lot.model,tracking.timein,tracking.timeout,"
        mSQLS1.CommandText += "tracking.users, users.name, station_type,failure.defect,defect.desc_th  "
        mSQLS1.CommandText += "FROM TRACKING  FULL JOIN STATION ON TRACKING.station = station.station "
        mSQLS1.CommandText += "full join users on tracking.users = users.id  full join lot  on tracking.lot = lot.lot "
        mSQLS1.CommandText += "LEFT JOIN FAILURE ON TRACKING.SN = FAILURE.SN AND TRACKING.station = FAILURE.failstation and abs(datediff(second,tracking.timeout,failure.failtime)) < 5 "
        mSQLS1.CommandText += "left join defect on failure.defect = defect.defect "
        mSQLS1.CommandText += "WHERE tracking.timein between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND tracking.station = '"
        mSQLS1.CommandText += tStation1 & "' "
        If tDefect_Code <> "ALL" Then
            mSQLS1.CommandText += " AND failure.defect  = '" & tDefect_Code & "' "
        End If
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "SELECT tracking_dup.station,station.stationname_cn,tracking_dup.sn,lot.model,tracking_dup.timein,tracking_dup.timeout,"
        mSQLS1.CommandText += "tracking_dup.users, users.name, station_type,failure.defect,defect.desc_th "
        mSQLS1.CommandText += "FROM TRACKING_dup  FULL JOIN STATION ON TRACKING_dup.station = station.station  full join users on tracking_dup.users = users.id "
        mSQLS1.CommandText += "full join lot  on tracking_dup.lot = lot.lot LEFT JOIN FAILURE ON TRACKING_dup.SN = FAILURE.SN AND TRACKING_dup.station = FAILURE.failstation "
        mSQLS1.CommandText += "and abs(datediff(second,tracking_dup.timeout,failure.failtime)) < 5 left join defect on failure.defect = defect.defect "
        mSQLS1.CommandText += "WHERE tracking_dup.timein between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND tracking_dup.station = '"
        mSQLS1.CommandText += tStation1 & "' "
        If tDefect_Code <> "ALL" Then
            mSQLS1.CommandText += " AND failure.defect  = '" & tDefect_Code & "' "
        End If
        mSQLS1.CommandText += " ) as BB order by sn"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("station").ToString() & mSQLReader.Item("stationname_cn").ToString()
                Ws.Cells(LineZ, 2) = mSQLReader.Item("sn").ToString()
                Ws.Cells(LineZ, 3) = mSQLReader.Item("model").ToString()
                Ws.Cells(LineZ, 4) = mSQLReader.Item("timein").ToString()
                Ws.Cells(LineZ, 5) = mSQLReader.Item("timeout").ToString()
                Ws.Cells(LineZ, 6) = mSQLReader.Item("users").ToString() & mSQLReader.Item("name").ToString()
                'If mSQLReader.Item("station_type").ToString.StartsWith("QC") Then
                'Dim NC As String = String.Empty
                'If Not IsDBNull(mSQLReader.Item("timeout")) Then
                'NC = GetDefectCode(mSQLReader.Item("sn"), mSQLReader.Item("station"), _
                '                                 mSQLReader.Item("timein"), mSQLReader.Item("timeout"))
                'Else
                '    NC = "N/A"
                'End If
                Ws.Cells(LineZ, 7) = mSQLReader.Item("defect") & " " & mSQLReader.Item("desc_th")
                'End If
                If tStation2 = "ALL" Then
                    GetDetail(mSQLReader.Item("sn"))
                Else
                    GetDetail1(mSQLReader.Item("sn"))
                End If
                LineZ += 1
                Me.ProgressBar1.Value = LineZ - 4
            End While
        End If
        mSQLReader.Close()
        AdjustExcelFormat1()

    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 30
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "G1")
        oRng.Merge()
        oRng = Ws.Range("A2", "G3")
        oRng.Interior.Color = Color.Green
        Ws.Cells(1, 1) = "基准工站信息"
        Ws.Cells(2, 1) = "Station"
        Ws.Cells(2, 2) = "SN"
        Ws.Cells(2, 3) = "Product Name"
        Ws.Cells(2, 4) = "Start time"
        Ws.Cells(2, 5) = "Finish time"
        Ws.Cells(2, 6) = "Operator1"
        Ws.Cells(2, 7) = "defect"
        Ws.Cells(3, 1) = "工站"
        Ws.Cells(3, 2) = "系列号"
        Ws.Cells(3, 3) = "产品名称"
        Ws.Cells(3, 4) = "开始时间"
        Ws.Cells(3, 5) = "结束时间"
        Ws.Cells(3, 6) = "作业员"
        Ws.Cells(3, 7) = "缺陷项"
    End Sub
    Private Function GetDefectCode(ByVal sn As String, ByVal station As String, ByVal time1 As DateTime, ByVal time2 As DateTime)
        Dim NR As String = String.Empty
        mSQLS2.CommandText = "SELECT failure.defect,defect.desc_th,defect.desc_en  FROM failure,defect  WHERE failure.defect = defect.defect and sn = '"
        mSQLS2.CommandText += sn & "' and failstation = '"
        mSQLS2.CommandText += station & "' and failure.failtime between '"
        mSQLS2.CommandText += time1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & time2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLReader2 = mSQLS2.ExecuteReader
        If mSQLReader2.HasRows() Then
            mSQLReader2.Read()
            NR = mSQLReader2.Item("defect") & " " & mSQLReader2.Item("desc_en") & " " & mSQLReader2.Item("desc_th")
        Else
            NR = String.Empty
        End If
        mSQLReader2.Close()
        Return NR
    End Function
    Private Sub GetDetail(ByVal sn As String)
        Dim CountDetail As Int16 = 1
        mSQLS3.CommandText = "SELECT station,stationname_cn,sn,timein,timeout,users,name,station_type,defect,desc_en,desc_th FROM ( "
        mSQLS3.CommandText += "SELECT tracking.station,station.stationname_cn,tracking.sn,tracking.timein,tracking.timeout,tracking.users, users.name, station_type,"
        mSQLS3.CommandText += "failure.defect,defect.desc_en, defect.desc_th "
        mSQLS3.CommandText += "FROM tracking left join failure on tracking.sn = failure.sn and tracking.station = failure.failstation "
        mSQLS3.CommandText += "and failure.failtime between tracking.timein and tracking.timeout "
        mSQLS3.CommandText += "left join defect on failure.defect = defect.defect "
        mSQLS3.CommandText += "full join station on tracking.station = station.station "
        mSQLS3.CommandText += "full join users on tracking.users = users.id WHERE  tracking.sn = '"
        mSQLS3.CommandText += sn & "' AND tracking.station < '" & tStation1 & "' "
        If tStation2 <> "ALL" Then
            mSQLS3.CommandText += "AND tracking.station between '" & tStation2 & "' AND '" & tStation1 & "' "
        End If
        'If tDefect_Code <> "ALL" Then
        'mSQLS3.CommandText += "AND failure.defect = '" & tDefect_Code & "' "
        'End If
        mSQLS3.CommandText += "union all "
        mSQLS3.CommandText += "SELECT tracking_dup.station,station.stationname_cn,tracking_dup.sn,tracking_dup.timein,tracking_dup.timeout,tracking_dup.users, users.name, station_type,"
        mSQLS3.CommandText += "failure.defect, defect.desc_en, defect.desc_th "
        mSQLS3.CommandText += "FROM tracking_dup left join failure on tracking_dup.sn = failure.sn and tracking_dup.station = failure.failstation "
        mSQLS3.CommandText += "and failure.failtime between tracking_dup.timein and tracking_dup.timeout "
        mSQLS3.CommandText += "left join defect on failure.defect = defect.defect "
        mSQLS3.CommandText += "full join station on tracking_dup.station = station.station "
        mSQLS3.CommandText += "full join users on tracking_dup.users = users.id WHERE  tracking_dup.sn = '"
        mSQLS3.CommandText += sn & "' AND tracking_dup.station < '" & tStation1 & "' "
        If tStation2 <> "ALL" Then
            mSQLS3.CommandText += "AND tracking_dup.station between '" & tStation2 & "' AND '" & tStation1 & "' "
        End If
        'If tDefect_Code <> "ALL" Then
        'mSQLS3.CommandText += "AND failure.defect = '" & tDefect_Code & "' "
        'End If
        mSQLS3.CommandText += ") as CC order by timeout DESC"
        mSQLReader3 = mSQLS3.ExecuteReader()
        If mSQLReader3.HasRows() Then
            While mSQLReader3.Read()
                If CountDetail > 410 Then
                    Exit While
                End If
                If CountDetail > MaxDetailCount Then
                    MaxDetailCount = CountDetail
                End If
                Ws.Cells(LineZ, CountDetail * 5 + 3) = mSQLReader3.Item("station").ToString() & mSQLReader3.Item("stationname_cn").ToString()
                Ws.Cells(LineZ, CountDetail * 5 + 4) = mSQLReader3.Item("sn").ToString()
                Ws.Cells(LineZ, CountDetail * 5 + 5) = mSQLReader3.Item("timein").ToString()
                Ws.Cells(LineZ, CountDetail * 5 + 6) = mSQLReader3.Item("timeout").ToString()
                Ws.Cells(LineZ, CountDetail * 5 + 7) = mSQLReader3.Item("users").ToString() & mSQLReader3.Item("name").ToString()
                'If mSQLReader3.Item("station_type").ToString.StartsWith("QC") Then
                'Dim NC As String = GetDefectCode(mSQLReader3.Item("sn"), mSQLReader3.Item("station"), _
                '                                mSQLReader3.Item("timein"), mSQLReader3.Item("timeout"))
                'Ws.Cells(LineZ, CountDetail * 6 + 7) = mSQLReader3.Item("defect") & " " & mSQLReader3.Item("desc_en") & " " & mSQLReader3.Item("desc_th")
                '  End If
                CountDetail += 1
            End While
        End If
        mSQLReader3.Close()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Output_Inspection"
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
    Private Sub AdjustExcelFormat1()
        For i As Int16 = 1 To MaxDetailCount Step 1
            oRng = Ws.Range(Ws.Cells(1, i * 5 + 3), Ws.Cells(1, i * 5 + 7))
            oRng.Merge()
            Ws.Cells(1, i * 5 + 3) = "追遡工站信息"
            Ws.Cells(2, i * 5 + 3) = "Station"
            Ws.Cells(2, i * 5 + 4) = "SN"
            Ws.Cells(2, i * 5 + 5) = "Start time"
            Ws.Cells(2, i * 5 + 6) = "Finish time"
            Ws.Cells(2, i * 5 + 7) = "Operator1"
            'Ws.Cells(2, i * 6 + 7) = "defect"
            Ws.Cells(3, i * 5 + 3) = "工站"
            Ws.Cells(3, i * 5 + 4) = "系列号"
            Ws.Cells(3, i * 5 + 5) = "开始时间"
            Ws.Cells(3, i * 5 + 6) = "结束时间"
            Ws.Cells(3, i * 5 + 7) = "作业员"
            'Ws.Cells(3, i * 6 + 7) = "缺陷项"
            oRng = Ws.Range(Ws.Cells(2, i * 5 + 3), Ws.Cells(3, i * 5 + 7))
            oRng.Interior.Color = Color.Green
        Next
    End Sub
    Private Sub GetDetail1(ByVal sn As String)
        Dim CountDetail As Int16 = 1
        mSQLS3.CommandText = "select TOP 1  * from ( "
        mSQLS3.CommandText += "SELECT tracking.station,station.stationname_cn,tracking.sn,tracking.timein,tracking.timeout,tracking.users, users.name, station_type "
        mSQLS3.CommandText += "FROM tracking full join station on tracking.station = station.station full join users on tracking.users = users.id WHERE  tracking.sn = '"
        mSQLS3.CommandText += sn & "' AND tracking.station = '" & tStation2 & "' and tracking.timeout <=  '"
        mSQLS3.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS3.CommandText += "union all "
        mSQLS3.CommandText += "SELECT tracking_dup.station,station.stationname_cn,tracking_dup.sn,tracking_dup.timein,tracking_dup.timeout,tracking_dup.users, users.name, station_type "
        mSQLS3.CommandText += "FROM tracking_dup full join station on tracking_dup.station = station.station full join users on tracking_dup.users = users.id WHERE  tracking_dup.sn = '"
        mSQLS3.CommandText += sn & "' and tracking_dup.station = '" & tStation2 & "' and tracking_dup.timeout <=  '"
        mSQLS3.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS3.CommandText += ") as CA ORDER BY TIMEOUT DESC"
        mSQLReader3 = mSQLS3.ExecuteReader()
        If mSQLReader3.HasRows() Then
            While mSQLReader3.Read()
                If CountDetail > 410 Then
                    Exit While
                        End If
                If CountDetail > MaxDetailCount Then
                    MaxDetailCount = CountDetail
                        End If
                Ws.Cells(LineZ, CountDetail * 5 + 3) = mSQLReader3.Item("station").ToString() & mSQLReader3.Item("stationname_cn").ToString()
                Ws.Cells(LineZ, CountDetail * 5 + 4) = mSQLReader3.Item("sn").ToString()
                Ws.Cells(LineZ, CountDetail * 5 + 5) = mSQLReader3.Item("timein").ToString()
                Ws.Cells(LineZ, CountDetail * 5 + 6) = mSQLReader3.Item("timeout").ToString()
                Ws.Cells(LineZ, CountDetail * 5 + 7) = mSQLReader3.Item("users").ToString() & mSQLReader3.Item("name").ToString()
                        'If mSQLReader3.Item("station_type").ToString.StartsWith("QC") Then
                        'Dim NC As String = GetDefectCode(mSQLReader3.Item("sn"), mSQLReader3.Item("station"), _
                        '                                mSQLReader3.Item("timein"), mSQLReader3.Item("timeout"))
                        'Ws.Cells(LineZ, CountDetail * 6 + 7) = mSQLReader3.Item("defect") & " " & mSQLReader3.Item("desc_en") & " " & mSQLReader3.Item("desc_th")
                        '  End If
                CountDetail += 1
                    End While
                End If
        mSQLReader3.Close()
    End Sub
End Class