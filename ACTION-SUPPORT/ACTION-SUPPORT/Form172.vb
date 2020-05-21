Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form172
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS11 As New SqlClient.SqlCommand
    Dim mSQLS12 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim mConnection2 As New SqlClient.SqlConnection
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader2 As SqlClient.SqlDataReader
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oConnection2 As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand3 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader99 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim StartWeek As String = String.Empty
    Dim EndWeek As String = String.Empty
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim oRng1 As Microsoft.Office.Interop.Excel.Range
    Dim ShipmentIndex As Int16 = 0
    Dim LineZ As Integer = 0
    Dim g_Success As Boolean = True
    Dim BiasWeek As Int16 = 0
    Dim ReportYear As Int16 = 0
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form173.Show()
        Form173.Focus()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form174.Show()
        Form174.Focus()
    End Sub

    Private Sub Form172_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfERPSUPPORT
        mConnection2.ConnectionString = Module1.OpenConnectionOfMes()
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        oConnection2.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mSQLS1.CommandTimeout = 600
                mSQLS11.Connection = mConnection
                mSQLS11.CommandType = CommandType.Text
                mSQLS11.CommandTimeout = 600
                mSQLS12.Connection = mConnection
                mSQLS12.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If

        If mConnection2.State <> ConnectionState.Open Then
            Try
                mConnection2.Open()
                mSQLS2.Connection = mConnection2
                mSQLS2.CommandType = CommandType.Text
                mSQLS2.CommandTimeout = 600
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommand2.Connection = oConnection
                oCommand2.CommandType = CommandType.Text
                oCommand3.Connection = oConnection
                oCommand3.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        If oConnection2.State <> ConnectionState.Open Then
            Try
                oConnection2.Open()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        Me.ComboBox1.SelectedItem = "本周"
        Me.ComboBox2.SelectedItem = "2019"
        Me.ComboBox3.SelectedItem = "3"
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Me.ComboBox1.SelectedItem = "本周" Then
            TextBox1.Enabled = False
            TextBox1.ReadOnly = True
            TextBox2.Enabled = False
            TextBox2.ReadOnly = True
            mSQLS1.CommandText = "select top 1 weekno  from IES7 where StartTime < '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' order by StartTime desc"
            Try
                Dim W1 As String = mSQLS1.ExecuteScalar()
                TextBox1.Text = W1
                TextBox2.Text = W1
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try

        Else
            TextBox1.Enabled = True
            TextBox1.ReadOnly = False
            TextBox2.Enabled = True
            TextBox2.ReadOnly = False
            TextBox1.Text = ""
            TextBox2.Text = ""
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' 檢查程式
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If Me.ComboBox1.SelectedItem = "区间" Then
            If String.IsNullOrEmpty(TextBox1.Text) Or String.IsNullOrEmpty(TextBox1.Text) Or Strings.Len(TextBox1.Text) <> 6 Or Strings.Len(TextBox2.Text) <> 6 Then
                MsgBox("计算范围区间有误")
                Return
            End If
        End If

        ' 檢查程式End

        StartWeek = TextBox1.Text
        EndWeek = TextBox2.Text

        BiasWeek = Me.ComboBox3.SelectedItem
        ReportYear = Me.ComboBox2.SelectedItem

        If Me.CheckedListBox1.CheckedItems.Count > 0 Then
            ShipmentIndex = CheckedListBox1.CheckedIndices(0)
        End If

        ProcessALL()
    End Sub

    Private Sub ProcessALL()
        Label3.Text = "计算中"
        Label3.Refresh()
        'ProcessStep1()
        'ProcessStep2()
        'ProcessStep3()
        'ProcessStep4()
        'ProcessStep5()
        'CheckProcessStep1()
        ProcessStep6()
        Label3.Text = "计算完毕"
        Label3.Refresh()
    End Sub

    Private Sub ProcessStep1()
        mSQLS1.CommandText = "DELETE IED21"
        Try
            mSQLS1.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        ' 裁紗
        GetD21W1(1, "IED0", "StandardTime3", "'0110','0111'")
        ' 預型
        GetD21W1(2, "IED0", "StandardTime3", "'0150','0151'", "预型Layup")
        ' PCM LayUp
        GetD21W1(3, "IED0", "StandardTime3", "'0150','0151'", "PCM Layup")
        ' 結構件預型
        GetD21W1(4, "IED0", "StandardTime3", "'0150','0151'", "结构件预型Structure Layup")
        ' 成型
        GetD21W2(5, "IED0", "StandardTime3", "StandardTime4", "'0330','0331'", "成型Molding")
        ' PCM Molding
        GetD21W2(6, "IED0", "StandardTime3", "StandardTime4", "'0330','0331'", "PCM Molding")
        ' 結構件成型
        GetD21W2(7, "IED0", "StandardTime3", "StandardTime4", "'0330','0331'", "结构件成型Structure Molding")
        ' 乳胶芯材Latex Core
        GetD21W1(8, "IED0", "StandardTime3", "'0145'", "乳胶芯材Latex Core")
        ' CNC
        GetD21W4(9, "IED0", "StandardTime3", "StandardTime4", "'0380'", "IED0", "StandardTime3", "StandardTime4", "'0530'", "'115AD0101001036','115AD0108001036'")
        ' 胶合
        GetD21W4(10, "IED0", "StandardTime3", "StandardTime4", "'0480'", "IED0", "StandardTime3", "StandardTime4", "'0610'", "'142AD0206001064'")
        ' 補土
        GetD21W3(11, "IED10", "TimeTotal", "IED0", "StandardTime4", "'0410'")
        ' 涂裝
        GetD21W2(12, "IED0", "StandardTime3", "StandardTime4", "'0590'")
        ' 拋光
        GetD21W3(13, "IED11", "StandardTime3", "IED0", "StandardTime4", "'0640'")
        ' 包裝
        GetD21W4(14, "IED0", "StandardTime3", "StandardTime4", "'0675'", "IED0", "StandardTime3", "StandardTime4", "'0650'", "'142AD0206A21066','142AD0206A11066','142AD0206A31066','142AD0206A41066','126AU0108A61066','126AU0108A71066','236BK0137010066'")
        CountD23W1()
    End Sub

    Private Sub GetD21W1(ByVal Sector As Int16, DB1 As String, F1 As String, ByVal S1 As String)
        ' 此Function 只供1工時, 且無工藝分類的用
        mSQLS1.CommandText = "SELECT * FROM IES7 WHERE WeekNo BETWEEN '" & StartWeek & "' AND '" & EndWeek & "'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Dim T1 As DateTime = mSQLReader.Item("StartTime")
                Dim T2 As DateTime = mSQLReader.Item("EndTime")
                Dim W1 As String = mSQLReader.Item("WeekNo")

                mSQLS2.CommandText = "Insert into ERPSUPPORT.dbo.IED21 "
                mSQLS2.CommandText += "select " & Sector & ",'" & W1 & "',isnull(cf01,'Norecord'),count(sn) as t1, isnull(" & F1 & ",0), Round((count(sn) * isnull(" & F1 & ",0) )/60,3) as t2, 0, 0  from tracking "
                mSQLS2.CommandText += "left join lot on tracking.lot = lot.lot "
                mSQLS2.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and tracking.station = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
                mSQLS2.CommandText += "left join ERPSUPPORT.dbo." & DB1 & " x1 on cf01 = x1.PN "
                mSQLS2.CommandText += "where tracking.timeout between '" & T1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & T2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                mSQLS2.CommandText += "and tracking.station in (" & S1 & ") group by cf01, " & F1
                Try
                    mSQLS2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub GetD21W1(ByVal Sector As Int16, DB1 As String, F1 As String, ByVal S1 As String, XX1 As String)
        ' 此Function 只供1工時, 且有工藝分類的用
        mSQLS1.CommandText = "SELECT * FROM IES7 WHERE WeekNo BETWEEN '" & StartWeek & "' AND '" & EndWeek & "'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Dim T1 As DateTime = mSQLReader.Item("StartTime")
                Dim T2 As DateTime = mSQLReader.Item("EndTime")
                Dim W1 As String = mSQLReader.Item("WeekNo")

                mSQLS2.CommandText = "Insert into ERPSUPPORT.dbo.IED21 "
                mSQLS2.CommandText += "select " & Sector & ",'" & W1 & "',isnull(cf01,'Norecord'),count(sn) as t1, isnull(" & F1 & ",0), Round((count(sn) * isnull(" & F1 & ",0) )/60,3) as t2, 0, 0  from tracking "
                mSQLS2.CommandText += "left join lot on tracking.lot = lot.lot "
                mSQLS2.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and tracking.station = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
                mSQLS2.CommandText += "left join ERPSUPPORT.dbo." & DB1 & " x1 on cf01 = x1.PN "
                mSQLS2.CommandText += "where tracking.timeout between '" & T1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & T2.ToString("yyyy/MM/dd HH:mm:ss") & "' and x1.Catalog = '" & XX1 & "' "
                mSQLS2.CommandText += "and tracking.station in (" & S1 & ") group by cf01, " & F1
                Try
                    mSQLS2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub GetD21W2(ByVal Sector As Int16, DB1 As String, F1 As String, ByVal F2 As String, ByVal S1 As String, XX1 As String)
        ' 此Function 只供2工時, 且有工藝分類的用
        mSQLS1.CommandText = "SELECT * FROM IES7 WHERE WeekNo BETWEEN '" & StartWeek & "' AND '" & EndWeek & "'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Dim T1 As DateTime = mSQLReader.Item("StartTime")
                Dim T2 As DateTime = mSQLReader.Item("EndTime")
                Dim W1 As String = mSQLReader.Item("WeekNo")

                mSQLS2.CommandText = "Insert into ERPSUPPORT.dbo.IED21 "
                mSQLS2.CommandText += "select " & Sector & ",'" & W1 & "',isnull(cf01,'Norecord'),count(sn) as t1, isnull(" & F1 & ",0), Round((count(sn) * isnull(" & F1 & ",0) )/60,3) as t2, isnull(" & F2 & ",0), Round((count(sn) * isnull(" & F2 & ",0))/60,3) as t3 from tracking "
                mSQLS2.CommandText += "left join lot on tracking.lot = lot.lot "
                mSQLS2.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and tracking.station = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
                mSQLS2.CommandText += "left join ERPSUPPORT.dbo." & DB1 & " x1 on cf01 = x1.PN "
                mSQLS2.CommandText += "where tracking.timeout between '" & T1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & T2.ToString("yyyy/MM/dd HH:mm:ss") & "' and x1.Catalog = '" & XX1 & "' "
                mSQLS2.CommandText += "and tracking.station in (" & S1 & ") group by cf01, " & F1 & "," & F2
                Try
                    mSQLS2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub GetD21W2(ByVal Sector As Int16, DB1 As String, F1 As String, ByVal F2 As String, ByVal S1 As String)
        ' 此Function 只供2工時, 且無工藝分類的用
        mSQLS1.CommandText = "SELECT * FROM IES7 WHERE WeekNo BETWEEN '" & StartWeek & "' AND '" & EndWeek & "'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Dim T1 As DateTime = mSQLReader.Item("StartTime")
                Dim T2 As DateTime = mSQLReader.Item("EndTime")
                Dim W1 As String = mSQLReader.Item("WeekNo")

                mSQLS2.CommandText = "Insert into ERPSUPPORT.dbo.IED21 "
                mSQLS2.CommandText += "select " & Sector & ",'" & W1 & "',isnull(cf01,'Norecord'),count(sn) as t1, isnull(" & F1 & ",0), Round((count(sn) * isnull(" & F1 & ",0) )/60,3) as t2, isnull(" & F2 & ",0), Round((count(sn) * isnull(" & F2 & ",0))/60,3) as t3 from tracking "
                mSQLS2.CommandText += "left join lot on tracking.lot = lot.lot "
                mSQLS2.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and tracking.station = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
                mSQLS2.CommandText += "left join ERPSUPPORT.dbo." & DB1 & " x1 on cf01 = x1.PN "
                mSQLS2.CommandText += "where tracking.timeout between '" & T1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & T2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                mSQLS2.CommandText += "and tracking.station in (" & S1 & ") group by cf01, " & F1 & "," & F2
                Try
                    mSQLS2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub GetD21W3(ByVal Sector As Int16, DB1 As String, F1 As String, ByVal DB2 As String, ByVal F2 As String, ByVal S1 As String)
        ' 此Function 只供補土, 拋光
        mSQLS1.CommandText = "SELECT * FROM IES7 WHERE WeekNo BETWEEN '" & StartWeek & "' AND '" & EndWeek & "'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Dim T1 As DateTime = mSQLReader.Item("StartTime")
                Dim T2 As DateTime = mSQLReader.Item("EndTime")
                Dim W1 As String = mSQLReader.Item("WeekNo")

                mSQLS2.CommandText = "Insert into ERPSUPPORT.dbo.IED21 "
                mSQLS2.CommandText += "select " & Sector & ",'" & W1 & "',isnull(cf01,'Norecord'),count(sn) as t1, AVG(ISNULL(x1." & F1 & ",0)), Round((count(sn) * AVG(ISNULL(x1." & F1 & ",0) ) )/60,3) as t2, AVG(x2." & F2 & "), Round((count(sn) * AVG(x2." & F2 & "))/60,3) as t3  from tracking "
                mSQLS2.CommandText += "left join lot on tracking.lot = lot.lot "
                mSQLS2.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and tracking.station = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
                mSQLS2.CommandText += "left join ERPSUPPORT.dbo." & DB1 & " x1 on lot.model = x1.ModelID "
                mSQLS2.CommandText += "left join ERPSUPPORT.dbo." & DB2 & " x2 on cf01 = x2.PN "
                mSQLS2.CommandText += "where tracking.timeout between '" & T1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & T2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                mSQLS2.CommandText += "and tracking.station in (" & S1 & ") group by cf01"
                Try
                    mSQLS2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End While
        End If
        mSQLReader.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If

        Dim xPath As String = "C:\temp\IER2_Template.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If

        mSQLS1.CommandText = "select count(*) from ied21 "
        Dim R2Counts As Decimal = mSQLS1.ExecuteScalar()
        If R2Counts = 0 Then
            MsgBox("无资料",MsgBoxStyle.Critical)
            Return
        End If

        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcelR2()
    End Sub
    Private Sub BackgroundWorker2_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        ExportToExcelR3()
    End Sub
    Private Sub ExportToExcelR2()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        Dim xPath As String = "C:\temp\IER2_Template.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)

        mSQLS1.CommandText = "select distinct Sector  from ied21  order by sector "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Dim PageNo As Decimal = mSQLReader.Item(0)
                Ws = xWorkBook.Sheets(PageNo)
                Ws.Activate()
                LineZ = 2
                mSQLS11.CommandText = "select * from ied21 where sector = " & mSQLReader.Item(0) & " order by weekno, pn "
                mSQLReader2 = mSQLS11.ExecuteReader()
                If mSQLReader2.HasRows() Then
                    While mSQLReader2.Read()
                        Ws.Cells(LineZ, 1) = LineZ - 1
                        Ws.Cells(LineZ, 2) = mSQLReader2.Item("WeekNo")
                        Ws.Cells(LineZ, 3) = mSQLReader2.Item("PN")
                        Ws.Cells(LineZ, 4) = mSQLReader2.Item("Qty1")
                        Ws.Cells(LineZ, 5) = mSQLReader2.Item("Time1")
                        Ws.Cells(LineZ, 6) = mSQLReader2.Item("TotalTime1")
                        Select Case PageNo
                            Case 5, 6, 7, 9, 10, 11, 12, 13, 14
                                Ws.Cells(LineZ, 7) = mSQLReader2.Item("Time2")
                                Ws.Cells(LineZ, 8) = mSQLReader2.Item("TotalTime2")
                        End Select
                        LineZ += 1
                        Label3.Text = PageNo & "-" & LineZ
                        Label3.Refresh()
                    End While
                End If
                mSQLReader2.Close()
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub SaveExcelR2()
        SaveFileDialog1.FileName = "Achievement"
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
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        Label3.Text = "R2已存檔"
        Label3.Refresh()
    End Sub
    Private Sub SaveExcelR3()
        SaveFileDialog1.FileName = "投入工时统计表"
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
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        Label3.Text = "R3已存檔"
        Label3.Refresh()
    End Sub
    Private Sub SaveExcelR4()
        SaveFileDialog1.FileName = "生产效率预测表"
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
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        Label3.Text = "R4已存檔"
        Label3.Refresh()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcelR2()
    End Sub
    Private Sub BackgroundWorker2_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        SaveExcelR3()
    End Sub
    Private Sub CountD23W1()
        mSQLS1.CommandText = "DELETE IED23 WHERE Section1 in (1, 2) AND Weekno between '" & StartWeek & "' AND '" & EndWeek & "'"
        Try
            mSQLS1.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        ' Section1
        mSQLS1.CommandText = "INSERT INTO IED23 select 1, sector, weekno, sum(totaltime1) from ied21 group by sector, weekno order by sector, weekno"
        Try
            mSQLS1.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

        'Section2
        mSQLS1.CommandText = "INSERT INTO IED23 select 2, sector, weekno, sum(totaltime1) from ied21 WHERE SECTOR IN (5, 9, 10, 11, 12, 13, 14) group by sector, weekno order by sector, weekno"
        Try
            mSQLS1.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

    End Sub
    Private Sub GetD21W4(ByVal Sector As Int16, DB1 As String, F1 As String, ByVal F2 As String, ByVal S1 As String, DB2 As String, F3 As String, F4 As String, ByVal S2 As String, ByVal PN2 As String)
        ' 此Function 只供2工時, 且無工藝分類的用 -- 有除外的資料 
        mSQLS1.CommandText = "SELECT * FROM IES7 WHERE WeekNo BETWEEN '" & StartWeek & "' AND '" & EndWeek & "'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Dim T1 As DateTime = mSQLReader.Item("StartTime")
                Dim T2 As DateTime = mSQLReader.Item("EndTime")
                Dim W1 As String = mSQLReader.Item("WeekNo")

                mSQLS2.CommandText = "Insert into ERPSUPPORT.dbo.IED21 "
                mSQLS2.CommandText += "select " & Sector & ",'" & W1 & "',isnull(cf01,'Norecord'),count(sn) as t1,isnull(" & F1 & ",0), Round((count(sn) * isnull(" & F1 & ",0) )/60,3) as t2, isnull(" & F2 & ",0), Round((count(sn) * isnull(" & F2 & ",0))/60,3) as t3 from tracking "
                mSQLS2.CommandText += "left join lot on tracking.lot = lot.lot "
                mSQLS2.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and tracking.station = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
                mSQLS2.CommandText += "left join ERPSUPPORT.dbo." & DB1 & " x1 on cf01 = x1.PN "
                mSQLS2.CommandText += "where tracking.timeout between '" & T1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & T2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                mSQLS2.CommandText += "and tracking.station in (" & S1 & ") group by cf01, " & F1 & "," & F2
                mSQLS2.CommandText += " union all "
                mSQLS2.CommandText += "select " & Sector & ",'" & W1 & "',isnull(cf01,'Norecord'),count(sn) as t1, isnull(" & F3 & ",0), Round((count(sn) * isnull(" & F3 & ",0) )/60,3) as t2, isnull(" & F4 & ",0), Round((count(sn) * isnull(" & F4 & ",0)) /60, 3) as t3 from tracking "
                mSQLS2.CommandText += "left join lot on tracking.lot = lot.lot "
                mSQLS2.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and tracking.station = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
                mSQLS2.CommandText += "left join ERPSUPPORT.dbo." & DB2 & " x1 on cf01 = x1.PN "
                mSQLS2.CommandText += "where tracking.timeout between '" & T1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & T2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                mSQLS2.CommandText += "and tracking.station in (" & S2 & ") and x1.pn in (" & PN2 & ") group by cf01, " & F3 & "," & F4

                Try
                    mSQLS2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End While
        End If
        mSQLReader.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Me.BackgroundWorker2.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If

        Dim xPath As String = "C:\temp\IER3_Template.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        Label3.Text = "R3报表处理中"
        BackgroundWorker2.RunWorkerAsync()
    End Sub
    Private Sub ExportToExcelR3()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        Dim xPath As String = "C:\temp\IER3_Template.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        LineZ = 2
        StartWeek = TextBox1.Text
        EndWeek = TextBox2.Text
        mSQLS1.CommandText = "select startTime from ies7 where weekno = '" & StartWeek & "'"
        Dim Date1 As Date = mSQLS1.ExecuteScalar()
        mSQLS1.CommandText = "select EndTime from ies7 where weekno = '" & EndWeek & "'"
        Dim Date2 As Date = mSQLS1.ExecuteScalar()

        Dim DateCount As Int16 = DateDiff(DateInterval.Day, Date1, Date2)
        For i As Int16 = 0 To DateCount Step 1
            Dim Date3 As Date = Date1.AddDays(i)
            Ws.Cells(LineZ, 1) = Date3.ToString("yyyy/MM/dd")
            Ws.Cells(LineZ, 2) = "=I" & LineZ
            Ws.Cells(LineZ, 4) = "=SUM(B" & LineZ & ":C" & LineZ & ")"
            Ws.Cells(LineZ, 5) = "=X" & LineZ
            Ws.Cells(LineZ, 7) = "=AM" & LineZ
            Ws.Cells(LineZ, 9) = "=SUM(J" & LineZ & ":W" & LineZ & ")"
            Ws.Cells(LineZ, 24) = "=SUM(Y" & LineZ & ":AL" & LineZ & ")"
            Ws.Cells(LineZ, 39) = "=SUM(AN" & LineZ & ":BA" & LineZ & ")"
            Ws.Cells(LineZ, 54) = "=SUM(BC" & LineZ & ":BP" & LineZ & ")"
            mSQLS1.CommandText = "select * from ied12 where date1 = '" & Date3.ToString("yyyy/MM/dd") & "'"
            mSQLReader = mSQLS1.ExecuteReader()
            If mSQLReader.HasRows() Then
                While mSQLReader.Read()
                    Ws.Cells(LineZ, 3) = mSQLReader.Item(1)
                    Ws.Cells(LineZ, 6) = mSQLReader.Item(2)
                    Ws.Cells(LineZ, 8) = mSQLReader.Item(3)
                    For j As Int16 = 1 To 14
                        Ws.Cells(LineZ, 9 + j) = mSQLReader.Item(3 + j)
                        Ws.Cells(LineZ, 24 + j) = mSQLReader.Item(17 + j)
                        Ws.Cells(LineZ, 39 + j) = mSQLReader.Item(31 + j)
                        Ws.Cells(LineZ, 54 + j) = mSQLReader.Item(45 + j)
                    Next
                    Ws.Cells(LineZ, 69) = mSQLReader.Item(60)
                    LineZ += 1
                End While
            End If
            mSQLReader.Close()
        Next

        ' 轉置第二頁
        oRng = Ws.Range("A2:BQ" & LineZ - 1)
        oRng.Select()
        oRng.Copy()
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        oRng1 = Ws.Range("B1")
        oRng1.Select()
        oRng1.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, False, True)

        ' 刪除第一頁
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        oRng.Delete()
        LineZ = 2
        For i As Int16 = 0 To DateCount Step 1
            Dim Date3 As Date = Date1.AddDays(i)
            Ws.Cells(LineZ, 1) = Date3.ToString("yyyy/MM/dd")
            Ws.Cells(LineZ, 2) = "=H" & LineZ
            Ws.Cells(LineZ, 4) = "=P" & LineZ
            Ws.Cells(LineZ, 6) = "=X" & LineZ
            Ws.Cells(LineZ, 8) = "=SUM(I" & LineZ & ":O" & LineZ & ")"
            Ws.Cells(LineZ, 16) = "=SUM(Q" & LineZ & ":W" & LineZ & ")"
            Ws.Cells(LineZ, 24) = "=SUM(Y" & LineZ & ":AE" & LineZ & ")"

            mSQLS1.CommandText = "select * from ied13 where date1 = '" & Date3.ToString("yyyy/MM/dd") & "'"
            mSQLReader = mSQLS1.ExecuteReader()
            If mSQLReader.HasRows() Then
                While mSQLReader.Read()
                    Ws.Cells(LineZ, 3) = mSQLReader.Item(1)
                    Ws.Cells(LineZ, 5) = mSQLReader.Item(2)
                    Ws.Cells(LineZ, 7) = mSQLReader.Item(3)
                    For j As Int16 = 1 To 7
                        Ws.Cells(LineZ, 8 + j) = mSQLReader.Item(3 + j)
                        Ws.Cells(LineZ, 16 + j) = mSQLReader.Item(10 + j)
                        Ws.Cells(LineZ, 24 + j) = mSQLReader.Item(17 + j)
                    Next
                    Ws.Cells(LineZ, 32) = mSQLReader.Item(25)
                    LineZ += 1
                End While
            End If
            mSQLReader.Close()
        Next

        ' 轉置第三頁
        oRng = Ws.Range("A2:AF" & LineZ - 1)
        oRng.Select()
        oRng.Copy()
        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        oRng1 = Ws.Range("B1")
        oRng1.Select()
        oRng1.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, False, True)
    End Sub

    Private Sub ProcessStep2()
        mSQLS1.CommandText = "Select weekno, StartTime, EndTime from IES7 Where Weekno between '" & StartWeek & "' AND '" & EndWeek & "' order by weekno"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Dim T1 As DateTime = mSQLReader.Item("StartTime")
                Dim T2 As DateTime = mSQLReader.Item("EndTime")
                Dim W1 As String = mSQLReader.Item("WeekNo")
                mSQLS11.CommandText = "select sector,isnull(sum(TotalTime1),0) from ied21 where weekno ='" & W1 & "' group by sector order by sector"
                mSQLReader2 = mSQLS11.ExecuteReader()
                If mSQLReader2.HasRows() Then
                    While mSQLReader2.Read()
                        mSQLS12.CommandText = "UPDATE IES7 SET s" & mSQLReader2.Item(0) & "= " & mSQLReader2.Item(1) & " WHERE weekno = '" & W1 & "'"
                        Try
                            mSQLS12.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.Message())
                        End Try
                    End While
                End If
                mSQLReader2.Close()

                ' 檢驗 5-7 
                mSQLS11.CommandText = "select isnull(sum(TotalTime2),0) from ied21 where weekno ='" & W1 & "' and sector in (5,6,7)"
                Dim Dq1 As Decimal = mSQLS11.ExecuteScalar()
                mSQLS11.CommandText = "UPDATE IES7 SET q1 = " & Dq1 & " WHERE weekno = '" & W1 & "'"
                Try
                    mSQLS11.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try

                ' 檢驗 9-14
                mSQLS11.CommandText = "select sector,isnull(sum(TotalTime2),0) from ied21 where weekno ='" & W1 & "' and sector > 8 group by sector order by sector"
                mSQLReader2 = mSQLS11.ExecuteReader()
                If mSQLReader2.HasRows() Then
                    While mSQLReader2.Read()
                        mSQLS12.CommandText = "UPDATE IES7 SET q" & Decimal.Subtract(mSQLReader2.Item(0), 7) & "= " & mSQLReader2.Item(1) & " WHERE weekno = '" & W1 & "'"
                        Try
                            mSQLS12.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.Message())
                        End Try
                    End While
                End If
                mSQLReader2.Close()

                ' 正班工時 + 加班工時
                mSQLS11.CommandText = "select sum(z1+j1) as z1,sum(z2+j2) as z2,sum(z3+j3) as z3, sum(z4+j4) as z4,sum(z5+j5) as z5,sum(z6+j6) as z6,sum(z7+j7) as z7,sum(z8+j8) as z8,sum(z9+j9) as z9, "
                mSQLS11.CommandText += "sum(z10+j10) as z10,sum(z11+j11) as z11,sum(z12+j12) as z12,sum(z13+j13) as z13,sum(z14+j14) as z14 "
                mSQLS11.CommandText += "from ied12 where date1 between '" & T1.ToString("yyyy/MM/dd") & "' and '" & T2.ToString("yyyy/MM/dd") & "' "
                mSQLReader2 = mSQLS11.ExecuteReader()
                If mSQLReader2.HasRows() Then
                    While mSQLReader2.Read()
                        For i As Int16 = 0 To mSQLReader2.FieldCount - 1 Step 1
                            mSQLS12.CommandText = "UPDATE IES7 SET st" & i + 1 & "=" & mSQLReader2.Item(i) & " WHERE weekno = '" & W1 & "'"
                            Try
                                mSQLS12.ExecuteNonQuery()
                            Catch ex As Exception
                                MsgBox(ex.Message())
                            End Try
                        Next
                    End While
                End If
                mSQLReader2.Close()

                ' 檢驗工時
                mSQLS11.CommandText = "select sum(z1+j1) as z1,sum(z2+j2) as z2,sum(z3+j3) as z3, sum(z4+j4) as z4,sum(z5+j5) as z5,sum(z6+j6) as z6,sum(z7+j7) as z7 "
                mSQLS11.CommandText += "from ied13 where date1 between '" & T1.ToString("yyyy/MM/dd") & "' and '" & T2.ToString("yyyy/MM/dd") & "' "
                mSQLReader2 = mSQLS11.ExecuteReader()
                If mSQLReader2.HasRows() Then
                    While mSQLReader2.Read()
                        For i As Int16 = 0 To mSQLReader2.FieldCount - 1 Step 1
                            mSQLS12.CommandText = "UPDATE IES7 SET qt" & i + 1 & "=" & mSQLReader2.Item(i) & " WHERE weekno = '" & W1 & "'"
                            Try
                                mSQLS12.ExecuteNonQuery()
                            Catch ex As Exception
                                MsgBox(ex.Message())
                            End Try
                        Next
                    End While
                End If
                mSQLReader2.Close()

                ' 直工效率
                For i As Int16 = 1 To 14 Step 1
                    mSQLS11.CommandText = "update ies7 set se" & i & " = ( select (case when st" & i & " =  0 then NULL else isnull(round(s" & i & "/st" & i & ",3),0) end)  from ies7 where weekno = '" & W1 & "' ) where weekno = '" & W1 & "'"
                    Try
                        mSQLS11.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                Next

                ' QC 效率
                For i As Int16 = 1 To 7 Step 1
                    mSQLS11.CommandText = "update ies7 set qe" & i & " = ( select (case when qt" & i & " =  0 then NULL else isnull(round(q" & i & "/qt" & i & ",3),0) end)  from ies7 where weekno = '" & W1 & "' ) where weekno = '" & W1 & "'"
                    Try
                        mSQLS11.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                Next

               

                ' 偏移推測值  -- 檢驗
                For i As Int16 = 1 To 7 Step 1
                    ' Step 1 看當週有無效率, 沒效率, 直接以上週推測值
                    mSQLS11.CommandText = "Select qe" & i & " from ies7 where weekno = '" & W1 & "'"
                    If IsDBNull(mSQLS11.ExecuteScalar()) Then
                        ' 取上週
                        Dim LastWeek As Int16 = Strings.Right(W1, 2)
                        Dim LastYear As Int16 = Strings.Left(W1, 4)
                        LastWeek = LastWeek - 1
                        If LastWeek = 0 Then
                            LastWeek = 52
                            LastYear = LastYear - 1
                        End If
                        Dim W2 As String = String.Empty
                        If LastWeek < 10 Then
                            W2 = LastYear & "0" & LastWeek
                        Else
                            W2 = LastYear & LastWeek
                        End If
                        'mSQLS11.CommandText = "Select sef" & i & " from ies7 WHERE weekno = '" & W2 & "'"
                        'Dim LastResult As Decimal = mSQLS11.ExecuteScalar()
                        'mSQLS11.CommandText = "UPDATE IES7 set sef" & i & " = " & LastResult & " WHERE weekno = '" & W1 & "'"
                        mSQLS11.CommandText = "update ies7 set qef" & i & " = (Select qef" & i & " from ies7 where weekno = '" & W2 & "') where WeekNo = '" & W1 & "'"
                        Try
                            mSQLS11.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.Message())
                        End Try
                    Else
                        ' Step2 , 計算有無足夠之前週數
                        Dim Week1 As Int16 = Strings.Right(W1, 2)
                        Dim Year1 As Int16 = Strings.Left(W1, 4)
                        Dim Week2 As Int16 = Strings.Right(W1, 2)
                        Dim Year2 As Int16 = Strings.Left(W1, 4)
                        Week2 = Week2 - 1
                        If Week2 = 0 Then
                            Year2 = Year2 - 1
                            Week2 = 52
                        End If
                        Week1 = Week1 - BiasWeek
                        If Week1 <= 0 Then
                            Year1 = Year1 - 1
                            Week1 = Week1 + 52
                        End If
                        Dim W2 As String = String.Empty
                        If Week1 < 10 Then
                            W2 = Year1 & "0" & Week1
                        Else
                            W2 = Year1 & Week1
                        End If
                        Dim W3 As String = String.Empty
                        If Week2 < 10 Then
                            W3 = Year2 & "0" & Week2
                        Else
                            W3 = Year2 & Week2
                        End If
                        mSQLS11.CommandText = "Select count(qe" & i & ") from ies7 where weekno between '" & W2 & "' and '" & W3 & "' and qe" & i & " is not null"
                        Dim BiasWeekX As Int16 = mSQLS11.ExecuteScalar()
                        If BiasWeekX < BiasWeek Then
                            ' 週數不夠, 取當週存檔值
                            'mSQLS11.CommandText = "select se" & i & " from ies7 where weekno = '" & W1 & "'"
                            'Dim LastResult As Decimal = mSQLS11.ExecuteScalar()
                            'mSQLS11.CommandText = "UPDATE IES7 set sef" & i & " = " & LastResult & " WHERE weekno = '" & W1 & "'"
                            mSQLS11.CommandText = "update ies7 set qef" & i & " = (Select qe" & i & " from ies7 where weekno = '" & W1 & "') where WeekNo = '" & W1 & "'"
                            Try
                                mSQLS11.ExecuteNonQuery()
                            Catch ex As Exception
                                MsgBox(ex.Message())
                            End Try
                        Else
                            ' 週數夠, 採用偏移平均值
                            mSQLS11.CommandText = "Select ROUND(avg(qe" & i & "),3) from ( SELECT TOP " & BiasWeekX - 2 & " qe" & i & " FROM IES7 AS a WHERE weekno between '"
                            mSQLS11.CommandText += W2 & "' and '" & W3 & "' and qe" & i & " is not null  AND Not Exists (Select * From (Select Top 1 qe" & i & ",WEEKNO From IES7  WHERE weekno between '"
                            mSQLS11.CommandText += W2 & "' and '" & W3 & "' and qe" & i & " is not null  order by QE" & i & ") b Where b.WeekNo =a.WeekNo  ) Order by QE" & i & " ) as AB"
                            If IsDBNull(mSQLS11.ExecuteScalar()) Then
                                ' 若取出的值是Null , 用上週推測值
                                mSQLS11.CommandText = "update ies7 set qef" & i & " = (Select qef" & i & " from ies7 where weekno = '" & W3 & "') where WeekNo = '" & W1 & "'"
                                Try
                                    mSQLS1.ExecuteNonQuery()
                                Catch ex As Exception
                                    MsgBox(ex.Message())
                                End Try
                            Else
                                ' modify by cloud 20191101
                                If mSQLS11.ExecuteScalar() = 0 Then
                                    mSQLS11.CommandText = "update ies7 set qef" & i & " = (Select qef" & i & " from ies7 where weekno = '" & W3 & "') where WeekNo = '" & W1 & "'"
                                    Try
                                        mSQLS1.ExecuteNonQuery()
                                    Catch ex As Exception
                                        MsgBox(ex.Message())
                                    End Try
                                Else
                                    ' 不是Null 和 0, 直接回寫
                                    Dim LastResult As Decimal = mSQLS11.ExecuteScalar()
                                    mSQLS11.CommandText = "UPDATE IES7 SET qef" & i & " = " & LastResult & " WHERE weekno = '" & W1 & "'"
                                    Try
                                        mSQLS11.ExecuteNonQuery()
                                    Catch ex As Exception
                                        MsgBox(ex.Message())
                                    End Try
                                End If
                            End If
                        End If
                    End If
                Next


            End While
        End If
        mSQLReader.Close()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Me.BackgroundWorker3.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If

        Dim xPath As String = "C:\temp\IER4_Template.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        Label3.Text = "R4报表处理中"
        BackgroundWorker3.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker3_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        ExportToExcelR4()
    End Sub
    Private Sub BackgroundWorker3_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        SaveExcelR4()
    End Sub
    Private Sub ExportToExcelR4()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        Dim xPath As String = "C:\temp\IER4_Template.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        ReportYear = Me.ComboBox2.SelectedItem
        Dim ReportNextYear As Int16 = ReportYear + 1
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        LineZ = 2        
        mSQLS1.CommandText = "Select * from ies7 where SUBSTRING(weekno, 1,4) = " & ReportYear & " union all Select * from ies7 where SUBSTRING(weekno, 1,4) = " & ReportNextYear & "  order by weekno"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                For j As Int16 = 6 To mSQLReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, j - 5) = mSQLReader.Item(j)
                Next
                LineZ += 1
            End While
        End If
        mSQLReader.Close()


        ' 轉置第二頁
        oRng = Ws.Range("A2:CF" & LineZ - 1)
        oRng.Select()
        oRng.Copy()
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        oRng1 = Ws.Range("D2")
        oRng1.Select()
        oRng1.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, False, True)

    End Sub

    Private Sub ProcessStep3()
        ReportYear = Me.ComboBox2.SelectedItem
        Dim ReportNextYear As Int16 = ReportYear + 1
        mSQLS1.CommandText = "Select weekno, StartTime, EndTime from IES7 Where SUBSTRING(weekno,1,4) in (" & ReportYear & "," & ReportNextYear & ") order by weekno"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Dim T1 As DateTime = mSQLReader.Item("StartTime")
                Dim T2 As DateTime = mSQLReader.Item("EndTime")
                Dim W1 As String = mSQLReader.Item("WeekNo")
                ' 偏移推測值  -- 工段
                For i As Int16 = 1 To 14 Step 1
                    ' Step 1 看當週有無效率, 沒效率, 直接以上週推測值
                    mSQLS11.CommandText = "Select se" & i & " from ies7 where weekno = '" & W1 & "'"
                    If IsDBNull(mSQLS11.ExecuteScalar()) Then
                        ' 取上週
                        Dim LastWeek As Int16 = Strings.Right(W1, 2)
                        Dim LastYear As Int16 = Strings.Left(W1, 4)
                        LastWeek = LastWeek - 1
                        If LastWeek = 0 Then
                            LastWeek = 52
                            LastYear = LastYear - 1
                        End If
                        Dim W2 As String = String.Empty
                        If LastWeek < 10 Then
                            W2 = LastYear & "0" & LastWeek
                        Else
                            W2 = LastYear & LastWeek
                        End If
                        'mSQLS11.CommandText = "Select sef" & i & " from ies7 WHERE weekno = '" & W2 & "'"
                        'Dim LastResult As Decimal = mSQLS11.ExecuteScalar()
                        'mSQLS11.CommandText = "UPDATE IES7 set sef" & i & " = " & LastResult & " WHERE weekno = '" & W1 & "'"
                        mSQLS11.CommandText = "update ies7 set sef" & i & " = (Select sef" & i & " from ies7 where weekno = '" & W2 & "') where WeekNo = '" & W1 & "'"
                        Try
                            mSQLS11.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.Message())
                        End Try
                    Else
                        ' Step2 , 計算有無足夠之前週數
                        Dim Week1 As Int16 = Strings.Right(W1, 2)
                        Dim Year1 As Int16 = Strings.Left(W1, 4)
                        Dim Week2 As Int16 = Strings.Right(W1, 2)
                        Dim Year2 As Int16 = Strings.Left(W1, 4)
                        Week2 = Week2 - 1
                        If Week2 = 0 Then
                            Year2 = Year2 - 1
                            Week2 = 52
                        End If
                        Week1 = Week1 - BiasWeek
                        If Week1 <= 0 Then
                            Year1 = Year1 - 1
                            Week1 = Week1 + 52
                        End If
                        Dim W2 As String = String.Empty
                        If Week1 < 10 Then
                            W2 = Year1 & "0" & Week1
                        Else
                            W2 = Year1 & Week1
                        End If
                        Dim W3 As String = String.Empty
                        If Week2 < 10 Then
                            W3 = Year2 & "0" & Week2
                        Else
                            W3 = Year2 & Week2
                        End If
                        mSQLS11.CommandText = "Select count(se" & i & ") from ies7 where weekno between '" & W2 & "' and '" & W3 & "' and se" & i & " is not null"
                        Dim BiasWeekX As Int16 = mSQLS11.ExecuteScalar()
                        If BiasWeekX < BiasWeek Then
                            ' 週數不夠, 取當週存檔值
                            'mSQLS11.CommandText = "select se" & i & " from ies7 where weekno = '" & W1 & "'"
                            'Dim LastResult As Decimal = mSQLS11.ExecuteScalar()
                            'mSQLS11.CommandText = "UPDATE IES7 set sef" & i & " = " & LastResult & " WHERE weekno = '" & W1 & "'"
                            mSQLS11.CommandText = "update ies7 set sef" & i & " = (Select se" & i & " from ies7 where weekno = '" & W1 & "') where WeekNo = '" & W1 & "'"
                            Try
                                mSQLS11.ExecuteNonQuery()
                            Catch ex As Exception
                                MsgBox(ex.Message())
                            End Try
                        Else
                            ' 週數夠, 採用偏移平均值
                            mSQLS11.CommandText = "Select ROUND(avg(se" & i & "),3) from ( SELECT TOP " & BiasWeekX - 2 & " se" & i & " FROM IES7 AS a WHERE weekno between '"
                            mSQLS11.CommandText += W2 & "' and '" & W3 & "' and se" & i & " is not null  AND Not Exists (Select * From (Select Top 1 se" & i & ",WEEKNO From IES7  WHERE weekno between '"
                            mSQLS11.CommandText += W2 & "' and '" & W3 & "' and se" & i & " is not null  order by se" & i & ") b Where b.WeekNo =a.WeekNo  ) Order by SE" & i & " ) as AB"
                            If IsDBNull(mSQLS11.ExecuteScalar()) Then
                                ' 若取出的值是Null , 用上週推測值
                                mSQLS11.CommandText = "update ies7 set sef" & i & " = (Select sef" & i & " from ies7 where weekno = '" & W3 & "') where WeekNo = '" & W1 & "'"
                                Try
                                    mSQLS11.ExecuteNonQuery()
                                Catch ex As Exception
                                    MsgBox(ex.Message())
                                End Try
                            Else
                                ' modify by cloud 20191101
                                Dim LastResult As Decimal = mSQLS11.ExecuteScalar()
                                If LastResult = 0 Then
                                    mSQLS11.CommandText = "update ies7 set sef" & i & " = (Select sef" & i & " from ies7 where weekno = '" & W3 & "') where WeekNo = '" & W1 & "'"
                                    Try
                                        mSQLS11.ExecuteNonQuery()
                                    Catch ex As Exception
                                        MsgBox(ex.Message())
                                    End Try
                                Else
                                    ' 不是Null 和 0, 直接回寫
                                    mSQLS11.CommandText = "UPDATE IES7 SET sef" & i & " = " & LastResult & " WHERE weekno = '" & W1 & "'"
                                    Try
                                        mSQLS11.ExecuteNonQuery()
                                    Catch ex As Exception
                                        MsgBox(ex.Message())
                                    End Try
                                End If
                            End If
                        End If
                    End If
                Next


                ' 偏移推測值  -- 檢驗
                For i As Int16 = 1 To 7 Step 1
                    ' Step 1 看當週有無效率, 沒效率, 直接以上週推測值
                    mSQLS11.CommandText = "Select qe" & i & " from ies7 where weekno = '" & W1 & "'"
                    If IsDBNull(mSQLS11.ExecuteScalar()) Then
                        ' 取上週
                        Dim LastWeek As Int16 = Strings.Right(W1, 2)
                        Dim LastYear As Int16 = Strings.Left(W1, 4)
                        LastWeek = LastWeek - 1
                        If LastWeek = 0 Then
                            LastWeek = 52
                            LastYear = LastYear - 1
                        End If
                        Dim W2 As String = String.Empty
                        If LastWeek < 10 Then
                            W2 = LastYear & "0" & LastWeek
                        Else
                            W2 = LastYear & LastWeek
                        End If
                        'mSQLS11.CommandText = "Select sef" & i & " from ies7 WHERE weekno = '" & W2 & "'"
                        'Dim LastResult As Decimal = mSQLS11.ExecuteScalar()
                        'mSQLS11.CommandText = "UPDATE IES7 set sef" & i & " = " & LastResult & " WHERE weekno = '" & W1 & "'"
                        mSQLS11.CommandText = "update ies7 set qef" & i & " = (Select qef" & i & " from ies7 where weekno = '" & W2 & "') where WeekNo = '" & W1 & "'"
                        Try
                            mSQLS11.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.Message())
                        End Try
                    Else
                        ' Step2 , 計算有無足夠之前週數
                        Dim Week1 As Int16 = Strings.Right(W1, 2)
                        Dim Year1 As Int16 = Strings.Left(W1, 4)
                        Dim Week2 As Int16 = Strings.Right(W1, 2)
                        Dim Year2 As Int16 = Strings.Left(W1, 4)
                        Week2 = Week2 - 1
                        If Week2 = 0 Then
                            Year2 = Year2 - 1
                            Week2 = 52
                        End If
                        Week1 = Week1 - BiasWeek
                        If Week1 <= 0 Then
                            Year1 = Year1 - 1
                            Week1 = Week1 + 52
                        End If
                        Dim W2 As String = String.Empty
                        If Week1 < 10 Then
                            W2 = Year1 & "0" & Week1
                        Else
                            W2 = Year1 & Week1
                        End If
                        Dim W3 As String = String.Empty
                        If Week2 < 10 Then
                            W3 = Year2 & "0" & Week2
                        Else
                            W3 = Year2 & Week2
                        End If
                        mSQLS11.CommandText = "Select count(qe" & i & ") from ies7 where weekno between '" & W2 & "' and '" & W3 & "' and qe" & i & " is not null"
                        Dim BiasWeekX As Int16 = mSQLS11.ExecuteScalar()
                        If BiasWeekX < BiasWeek Then
                            ' 週數不夠, 取當週存檔值
                            'mSQLS11.CommandText = "select se" & i & " from ies7 where weekno = '" & W1 & "'"
                            'Dim LastResult As Decimal = mSQLS11.ExecuteScalar()
                            'mSQLS11.CommandText = "UPDATE IES7 set sef" & i & " = " & LastResult & " WHERE weekno = '" & W1 & "'"
                            mSQLS11.CommandText = "update ies7 set qef" & i & " = (Select qe" & i & " from ies7 where weekno = '" & W1 & "') where WeekNo = '" & W1 & "'"
                            Try
                                mSQLS11.ExecuteNonQuery()
                            Catch ex As Exception
                                MsgBox(ex.Message())
                            End Try
                        Else
                            ' 週數夠, 採用偏移平均值
                            mSQLS11.CommandText = "Select ROUND(avg(qe" & i & "),3) from ( SELECT TOP " & BiasWeekX - 2 & " qe" & i & " FROM IES7 AS a WHERE weekno between '"
                            mSQLS11.CommandText += W2 & "' and '" & W3 & "' and qe" & i & " is not null  AND Not Exists (Select * From (Select Top 1 qe" & i & ",WEEKNO From IES7  WHERE weekno between '"
                            mSQLS11.CommandText += W2 & "' and '" & W3 & "' and qe" & i & " is not null  order by QE" & i & ") b Where b.WeekNo =a.WeekNo  ) Order by QE" & i & " ) as AB"
                            If IsDBNull(mSQLS11.ExecuteScalar()) Then
                                ' 若取出的值是Null , 用上週推測值
                                mSQLS11.CommandText = "update ies7 set qef" & i & " = (Select qef" & i & " from ies7 where weekno = '" & W3 & "') where WeekNo = '" & W1 & "'"
                                Try
                                    mSQLS11.ExecuteNonQuery()
                                Catch ex As Exception
                                    MsgBox(ex.Message())
                                End Try
                            Else
                                ' modify by cloud 20191101
                                Dim LastResult As Decimal = mSQLS11.ExecuteScalar()
                                If LastResult = 0 Then
                                    mSQLS11.CommandText = "update ies7 set qef" & i & " = (Select qef" & i & " from ies7 where weekno = '" & W3 & "') where WeekNo = '" & W1 & "'"
                                    Try
                                        mSQLS11.ExecuteNonQuery()
                                    Catch ex As Exception
                                        MsgBox(ex.Message())
                                    End Try
                                Else
                                    ' 不是Null 和 0, 直接回寫

                                    mSQLS11.CommandText = "UPDATE IES7 SET qef" & i & " = " & LastResult & " WHERE weekno = '" & W1 & "'"
                                    Try
                                        mSQLS11.ExecuteNonQuery()
                                    Catch ex As Exception
                                        MsgBox(ex.Message())
                                    End Try
                                End If
                            End If
                        End If
                    End If
                Next
            End While
        End If
        mSQLReader.Close()
    End Sub

    Private Sub CheckedListBox1_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles CheckedListBox1.ItemCheck
        SingleSelectCheckedListBoxControls(CheckedListBox1, e.Index)
    End Sub
    Public Sub SingleSelectCheckedListBoxControls(ByVal CheckedListbox1 As CheckedListBox, ByVal s1 As Decimal)
        If CheckedListbox1.CheckedItems.Count > 0 Then
            For i As Int16 = 0 To CheckedListbox1.Items.Count - 1 Step 1
                If i <> s1 Then
                    CheckedListbox1.SetItemChecked(i, CheckState.Unchecked)
                End If
            Next
        End If
    End Sub
    Private Sub ProcessStep4()
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mSQLS1.CommandTimeout = 600
                mSQLS11.Connection = mConnection
                mSQLS11.CommandType = CommandType.Text
                mSQLS11.CommandTimeout = 600
                mSQLS12.Connection = mConnection
                mSQLS12.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If

        mSQLS1.CommandText = "DELETE IEC1"
        Try
            mSQLS1.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
            Return
        End Try
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommand2.Connection = oConnection
                oCommand2.CommandType = CommandType.Text
                oCommand3.Connection = oConnection
                oCommand3.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        oCommand.CommandText = "Select pn,sum("
        For i As Int16 = 1 To 104 Step 1
            oCommand.CommandText += "+nvl(w" & i & ",0)"
        Next
        oCommand.CommandText += ") as t1 from ship_temp where "
        If ShipmentIndex = 0 Then
            oCommand.CommandText += "etype = 1 "
        End If
        If ShipmentIndex = 1 Then
            oCommand.CommandText += "etype = 2 "
        End If
        If ShipmentIndex = 2 Then
            oCommand.CommandText += "etype = 4 "
        End If
        If ShipmentIndex = 3 Then
            oCommand.CommandText += "etype = 5 "
        End If
        If ShipmentIndex = 4 Then
            oCommand.CommandText += "etype in (2, 5) "
        End If
        oCommand.CommandText += " and pn not like 'S%' and length(pn) >= 15 and ("
        For i As Int16 = 1 To 104 Step 1
            oCommand.CommandText += "+nvl(w" & i & ",0)"
        Next
        oCommand.CommandText += ") <> 0 group by pn order by pn"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                mSQLS1.CommandText = "INSERT INTO IEC1 (PN,S1) VALUES ('" & oReader.Item(0) & "'," & oReader.Item(1) & ")"
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End While
        Else
            mSQLS1.CommandText = "UPDATE IEE1 SET Result = 'NG' WHERE ErrorCode = '19001'"
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                g_Success = False
            End Try
        End If
        oReader.Close()
        If g_Success = False Then
            Return
        End If
        ' IEC1 有數據 , 查 ERP-BOM 狀態回寫
        mSQLS1.CommandText = "Select * from IEC1 "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                oCommand.CommandText = "Select bma10 from bma_file where bma01  ='" & mSQLReader.Item(0) & "'"
                Try
                    Dim BOMS As Int16 = oCommand.ExecuteScalar()
                    mSQLS12.CommandText = "UPDATE IEC1 SET BomStatus = " & BOMS & " WHERE PN = '" & mSQLReader.Item(0) & "'"
                    mSQLS12.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try

            End While
        End If
        mSQLReader.Close()
        ' 查臨時IES6表狀態
        mSQLS1.CommandText = "Select IEC1.PN from IEC1, IES6 WHERE IEC1.PN = IES6.PN"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                mSQLS11.CommandText = "UPDATE IEC1 SET TempIETime = 1 WHERE PN = '" & mSQLReader.Item(0) & "'"
                Try
                    mSQLS11.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    Return
                End Try
            End While
        End If
        mSQLReader.Close()
        ' 把多餘的從 S6 去除
        mSQLS1.CommandText = "Select PN from IEC1 WHERE BomStatus = 2 and TempIETime = 1"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                mSQLS11.CommandText = "DELETE IES6 WHERE PN = '" & mSQLReader.Item(0) & "'"
                Try
                    mSQLS11.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    Return
                End Try
            End While
        End If
        mSQLReader.Close()

        ' 把缺的塞進 S6
        mSQLS1.CommandText = "Select PN from IEC1 WHERE BomStatus <> 2 and TempIETime IS NULL"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                mSQLS11.CommandText = "INSERT INTO IES6 (PN) VALUES ('" & mSQLReader.Item(0) & "')"
                Try
                    mSQLS11.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    Return
                End Try
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub ProcessStep5()
        mSQLS1.CommandText = "DELETE IED1"
        Try
            mSQLS1.ExecuteNonQuery()
        Catch ex As Exception

        End Try

        mSQLS1.CommandText = "Select PN from iec1 order by pn"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            Label3.Text = "计算IED1"
            Label3.Refresh()
            While mSQLReader.Read()
                For i As Int16 = 1 To 104 Step 1
                    oCommand.CommandText = "Select nvl(sum(w" & i & "),0) from ship_temp where pn = '" & mSQLReader.Item(0) & "' "
                    If ShipmentIndex = 0 Then
                        oCommand.CommandText += "and etype = 1 "
                    End If
                    If ShipmentIndex = 1 Then
                        oCommand.CommandText += "and etype = 2 "
                    End If
                    If ShipmentIndex = 2 Then
                        oCommand.CommandText += "and etype = 4 "
                    End If
                    If ShipmentIndex = 3 Then
                        oCommand.CommandText += "and etype = 5 "
                    End If
                    If ShipmentIndex = 4 Then
                        oCommand.CommandText += "and etype in (2, 5) "
                    End If
                    oCommand.CommandText += " and w" & i & " is not null"
                    Dim WR1 As Decimal = oCommand.ExecuteScalar()
                    Dim WeekNo1 As String = String.Empty
                    Select Case i
                        Case Is < 10
                            WeekNo1 = ReportYear & "0" & i
                        Case 10 To 52
                            WeekNo1 = ReportYear & i
                        Case 53 To 61
                            WeekNo1 = ReportYear + 1 & "0" & i - 52
                        Case Is > 61
                            WeekNo1 = ReportYear + 1 & i - 52
                    End Select
                    mSQLS11.CommandText = "INSERT INTO IED1 (PN,WeekNo,Package1) VALUES ('" & mSQLReader.Item(0) & "','" & WeekNo1 & "'," & WR1 & ") "
                    Try
                        mSQLS11.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.Message())
                        Return
                    End Try
                Next
                ' 要加 裁切的週數
                mSQLS11.CommandText = "Select PreparePeriod  from ies1 where id = 1"
                Dim ExtendWeek As Int16 = mSQLS11.ExecuteScalar()
                Dim ExtendWeek1 As Int16 = 104 + ExtendWeek

                For i As Int16 = 105 To ExtendWeek1 Step 1
                    Dim WeekNo1 As String = ReportYear + 2 & "0" & i - 104
                    mSQLS11.CommandText = "INSERT INTO IED1 (PN,WeekNo,Package1) VALUES ('" & mSQLReader.Item(0) & "','" & WeekNo1 & "',0)"
                    Try
                        mSQLS11.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.Message())
                        Return
                    End Try
                Next
            End While
        End If
        mSQLReader.Close()

        '拋光處理
        mSQLS1.CommandText = "Select PreparePeriod from ies1 where id = 13"
        Dim PP1 As Int16 = mSQLS1.ExecuteScalar()
        If PP1 = 0 Then
            mSQLS1.CommandText = "UPDATE IED1 SET Polishing1 = Package1 "
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        Else
            mSQLS1.CommandText = "Update B SET Polishing1 = A.Package1 from IED1 A LEFT JOIN IED1 B ON A.pn = B.pn and A.WeekNo = (case when right(B.WeekNo,2) + " & PP1 & " > 52 then B.WeekNo + 51 else B.WeekNo +" & PP1 & " end )"
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        End If

        '涂裝處理
        mSQLS1.CommandText = "Select PreparePeriod from ies1 where id = 12"
        PP1 = mSQLS1.ExecuteScalar()
        If PP1 = 0 Then
            mSQLS1.CommandText = "UPDATE IED1 SET Painting1 = Package1 "
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        Else
            mSQLS1.CommandText = "Update B SET Painting1 = A.Package1 from IED1 A LEFT JOIN IED1 B ON A.pn = B.pn and A.WeekNo = (case when right(B.WeekNo,2) + " & PP1 & " > 52 then B.WeekNo + 51 else B.WeekNo +" & PP1 & " end )"
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        End If

        'Sanding處理
        mSQLS1.CommandText = "Select PreparePeriod from ies1 where id = 11"
        PP1 = mSQLS1.ExecuteScalar()
        If PP1 = 0 Then
            mSQLS1.CommandText = "UPDATE IED1 SET Sanding1 = Package1 "
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        Else
            mSQLS1.CommandText = "Update B SET Sanding1 = A.Package1 from IED1 A LEFT JOIN IED1 B ON A.pn = B.pn and A.WeekNo = (case when right(B.WeekNo,2) + " & PP1 & " > 52 then B.WeekNo + 51 else B.WeekNo +" & PP1 & " end )"
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        End If

        'Glueing處理
        mSQLS1.CommandText = "Select PreparePeriod from ies1 where id = 10"
        PP1 = mSQLS1.ExecuteScalar()
        If PP1 = 0 Then
            mSQLS1.CommandText = "UPDATE IED1 SET Glueing1 = Package1 "
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        Else
            mSQLS1.CommandText = "Update B SET Glueing1 = A.Package1 from IED1 A LEFT JOIN IED1 B ON A.pn = B.pn and A.WeekNo = (case when right(B.WeekNo,2) + " & PP1 & " > 52 then B.WeekNo + 51 else B.WeekNo +" & PP1 & " end )"
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        End If

        'CNC處理
        mSQLS1.CommandText = "Select PreparePeriod from ies1 where id = 9"
        PP1 = mSQLS1.ExecuteScalar()
        If PP1 = 0 Then
            mSQLS1.CommandText = "UPDATE IED1 SET CNC1 = Package1 "
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        Else
            mSQLS1.CommandText = "Update B SET CNC1 = A.Package1 from IED1 A LEFT JOIN IED1 B ON A.pn = B.pn and A.WeekNo = (case when right(B.WeekNo,2) + " & PP1 & " > 52 then B.WeekNo + 51 else B.WeekNo +" & PP1 & " end )"
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        End If

        'MOLDING處理
        mSQLS1.CommandText = "Select PreparePeriod from ies1 where id = 3"
        PP1 = mSQLS1.ExecuteScalar()
        If PP1 = 0 Then
            mSQLS1.CommandText = "UPDATE IED1 SET MOLDING1 = Package1 "
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        Else
            mSQLS1.CommandText = "Update B SET MOLDING1 = A.Package1 from IED1 A LEFT JOIN IED1 B ON A.pn = B.pn and A.WeekNo = (case when right(B.WeekNo,2) + " & PP1 & " > 52 then B.WeekNo + 51 else B.WeekNo +" & PP1 & " end )"
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        End If

        'LAYUP處理
        mSQLS1.CommandText = "Select PreparePeriod from ies1 where id = 2"
        PP1 = mSQLS1.ExecuteScalar()
        If PP1 = 0 Then
            mSQLS1.CommandText = "UPDATE IED1 SET LAYUP1 = Package1 "
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        Else
            mSQLS1.CommandText = "Update B SET LAYUP1 = A.Package1 from IED1 A LEFT JOIN IED1 B ON A.pn = B.pn and A.WeekNo = (case when right(B.WeekNo,2) + " & PP1 & " > 52 then B.WeekNo + 51 else B.WeekNo +" & PP1 & " end )"
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        End If

        'CUTTING處理
        mSQLS1.CommandText = "Select PreparePeriod from ies1 where id = 1"
        PP1 = mSQLS1.ExecuteScalar()
        If PP1 = 0 Then
            mSQLS1.CommandText = "UPDATE IED1 SET CUTTING1 = Package1 "
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        Else
            mSQLS1.CommandText = "Update B SET CUTTING1 = A.Package1 from IED1 A LEFT JOIN IED1 B ON A.pn = B.pn and A.WeekNo = (case when right(B.WeekNo,2) + " & PP1 & " > 52 then B.WeekNo + 51 else B.WeekNo +" & PP1 & " end )"
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        End If
    End Sub
    Private Sub CheckProcessStep1()
        mSQLS1.CommandText = "Select SUM(ISNULL(package1,0) + ISNULL(Polishing1,0) + ISNULL(Painting1,0) + ISNULL(Sanding1,0) + ISNULL(Glueing1,0) + ISNULL(CNC1,0) + ISNULL(Molding1,0) + ISNULL(Layup1,0) + ISNULL(Cutting1,0)) from ies7 left join ied1 on ies7.WeekNo = ied1.WeekNo WHERE IES7.WeekOff = 'Y'"
        Dim CS As Int16 = 0
        Try
            CS = mSQLS1.ExecuteScalar()
        Catch ex As Exception

        End Try
        If CS > 0 Then
            mSQLS1.CommandText = "update IEE1 set Result = 'NG' WHERE ErrorCode = '21001'"
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If Me.BackgroundWorker4.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If

        Dim xPath As String = "C:\temp\IER1_Template.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        Label3.Text = "R1报表处理中"
        BackgroundWorker4.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker4_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        ExportToExcelR1()
    End Sub
    Private Sub BackgroundWorker4_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
        SaveExcelR1()
    End Sub
    Private Sub SaveExcelR1()
        SaveFileDialog1.FileName = "订单需求预测表"
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
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        Label3.Text = "R1已存檔"
        Label3.Refresh()
    End Sub
    Private Sub ExportToExcelR1()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        Dim xPath As String = "C:\temp\IER1_Template.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        LineZ = 4
        ReportYear = Me.ComboBox2.SelectedItem
        mSQLS1.CommandText = "Select pn"
        For i As Int16 = 1 To 104 Step 1
            mSQLS1.CommandText += ",sum(t" & i & ") as t" & i
        Next
        mSQLS1.CommandText += " from ( Select Pn"
        For i As Int16 = 1 To 104 Step 1
            Dim WeekX As String = String.Empty
            Select Case i
                Case Is < 10
                    WeekX = ReportYear & "0" & i
                Case 11 To 52
                    WeekX = ReportYear & i
                Case 53 To 61
                    WeekX = ReportYear + 1 & "0" & i - 52
                Case Is > 61
                    WeekX = ReportYear + 1 & i - 52
            End Select
            mSQLS1.CommandText += ",(case when weekno = '" & WeekX & "' then package1 else 0 end ) as t" & i
        Next
        mSQLS1.CommandText += " from ied1 ) as AB group by pn"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item(0)
                Ws.Cells(LineZ, 3) = "=SUM(D" & LineZ & ":DE" & LineZ & ")"
                For i As Int16 = 1 To mSQLReader.FieldCount - 1 Step 1
                    If mSQLReader.Item(i) <> 0 Then
                        Ws.Cells(LineZ, 3 + i) = mSQLReader.Item(i)
                    End If
                Next
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub ProcessStep6()
        mSQLS1.CommandText = "DELETE IED2 "
        Try
            mSQLS1.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        
        'CAllBom("101AA0105011066", "101AA0105011066", 1)
       
        mSQLS1.CommandText = "Select distinct pn from iec1 where BomStatus = 2 and TempIETime is null"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            Dim SC As Int16 = 1
            While mSQLReader.Read()
                CAllBom(mSQLReader.Item(0), mSQLReader.Item(0), 1)
                If oConnection2.State <> ConnectionState.Closed Then
                    oConnection2.Close()
                End If
                Label3.Text = "展开BOM表" & SC
                Label3.Refresh()
                SC += 1
            End While
        End If
    End Sub

    Public Sub CAllBom(ByVal g_bmb01 As String, ByVal bmb01 As String, ByVal Qty1 As Decimal)
        If oConnection2.State <> ConnectionState.Open Then
            oConnection2.Open()
        End If
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection2
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "select bmb03,SUM(bmb06*" & Qty1 & ") from bmb_file,ima_file where bmb03 = ima01 and bmb29 = ima910 and bmb01 = '" & bmb01 & "' and bmb05 is null and bmb19 = '2' and ima08 = 'M' GROUP BY bmb03"
        oReader99 = oCommander99.ExecuteReader()
        If oReader99.HasRows() Then
            While oReader99.Read()
                mSQLS11.CommandText = "INSERT INTO IED2 VALUES ('" & g_bmb01 & "','" & oReader99.Item(0) & "'," & oReader99.Item(1) & ")"
                Try
                    mSQLS11.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
                Call CAllBom(g_bmb01, oReader99.Item(0), oReader99.Item(1))
            End While
        End If
        'oReader99.Close()
    End Sub
End Class