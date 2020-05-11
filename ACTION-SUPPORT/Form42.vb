﻿Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form42
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim tStation1 As String
    Dim ptime As String = String.Empty
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    'Dim CheckLastStation As Boolean = False
    Dim LastStation As String = String.Empty
    Dim ERPPN As String = String.Empty
    Dim tModel As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form42_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        ptime = Today.AddDays(-1).ToString("yyyy/MM/dd")
        ptime = ptime & " 08:00:00"
        Me.DateTimePicker1.Value = Convert.ToDateTime(ptime)
        Me.DateTimePicker2.Value = Convert.ToDateTime(ptime).AddDays(1)
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mSQLS1.CommandTimeout = 600
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
                mSQLS2.CommandTimeout = 600
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        BindModel_Station()
        BindModel()
    End Sub
    Private Sub BindModel_Station()
        Me.ComboBox1.Items.Clear()
        mSQLS1.CommandText = "SELECT station FROM station "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox1.Items.Add(mSQLReader.Item(0).ToString())
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub BindModel()
        Me.ComboBox2.Items.Clear()
        mSQLS1.CommandText = "SELECT model FROM model "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox2.Items.Add(mSQLReader.Item(0).ToString())
            End While
        End If
        mSQLReader.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If IsNothing(ComboBox1.SelectedItem) Then
            MsgBox("请选择工站")
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
        TimeS1 = DateTimePicker1.Value
        TimeS2 = DateTimePicker2.Value
        tStation1 = Me.ComboBox1.SelectedItem.ToString()
        'CheckLastStation = Me.CheckBox1.Checked
        tModel = String.Empty
        If Not IsNothing(ComboBox2.SelectedItem) Then
            tModel = Me.ComboBox2.SelectedItem.ToString()
        End If
        'ExportToExcel()
        'SaveExcel()
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        ' add by cloud 20180201
        Dim DBS As String = "trans" & Now.ToString("yyyyMMddHHmmss")
        mSQLS1.CommandText = "CREATE TABLE " & DBS & " (model nvarchar(20), ERPPN nvarchar(40), ERPDESC nvarchar(255),station nvarchar(255),GQ numeric(18,0),SQ numeric(18,0))"
        Try
            mSQLS1.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Name = "Detail"
        AdjustExcelFormat()
        LineZ = 6
        Ws.Cells(1, 2) = TimeS1.ToString("yyyy/MM/dd HH:mm:ss")
        Ws.Cells(2, 2) = TimeS2.ToString("yyyy/MM/dd HH:mm:ss")
        'mSQLS1.CommandText = "SELECT model,cf01,modelname,sn,station,stationname_cn,timein,timeout,users,name,fresh FROM ( "
        'mSQLS1.CommandText += "select lot.model,model_station_paravalue.cf01,model.modelname,tracking.sn,tracking.station,station.stationname_cn,"
        'mSQLS1.CommandText += "tracking.timein, tracking.timeout, tracking.users, users.name, tracking.fresh from tracking "
        'mSQLS1.CommandText += "left join lot on tracking.lot = lot.lot left join model on lot.model = model.model "
        'mSQLS1.CommandText += "left join station on tracking.station =station.station left join users on tracking.users = users.id "
        'mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue .profilename = 'ERP' and tracking.station = model_station_paravalue.station "
        'mSQLS1.CommandText += "WHERE tracking.TIMEOUT BETWEEN '"
        'mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        'mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND tracking.station  = '" & tStation1 & "' "
        'mSQLS1.CommandText += "union all "
        'mSQLS1.CommandText += "select lot.model,model_station_paravalue.cf01,model.modelname,scrap_tracking.sn,scrap_tracking.station,station.stationname_cn,"
        'mSQLS1.CommandText += "scrap_tracking.timein,scrap_tracking.timeout,scrap_tracking.users,users.name,scrap_tracking.fresh from scrap_tracking  "
        'mSQLS1.CommandText += "left join lot on scrap_tracking.lot = lot.lot left join model on lot.model = model.model "
        'mSQLS1.CommandText += "left join station on scrap_tracking.station =station.station left join users on scrap_tracking.users = users.id "
        'mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue.profilename = 'ERP' and scrap_tracking.station = model_station_paravalue.station "
        'mSQLS1.CommandText += "WHERE scrap_tracking.TIMEOUT BETWEEN '"
        'mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        'mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND scrap_tracking.station  = '" & tStation1 & "' "
        'mSQLS1.CommandText += ") as AA order by model"
        mSQLS1.CommandText = "SELECT model,cf01,modelname,sn,station,stationname_cn,timein,timeout,users,name,route,seq,laststation FROM ( "
        mSQLS1.CommandText += "select lot.model,model_station_paravalue.cf01,model.modelname,tracking.sn,tracking.station,station.stationname_cn,"
        mSQLS1.CommandText += "tracking.timein, tracking.timeout, tracking.users, users.name, lot.route, s1.seq,s2.station as laststation from tracking "
        mSQLS1.CommandText += "left join lot on tracking.lot = lot.lot left join model on lot.model = model.model "
        mSQLS1.CommandText += "left join station on tracking.station =station.station left join users on tracking.users = users.id "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue .profilename = 'ERP' and tracking.station = model_station_paravalue.station "
        mSQLS1.CommandText += "left join routing s1 on lot.route = s1.route and tracking.station = s1.station left join routing s2 on lot.route = s2.route and (s1.seq - 1 ) = s2.seq "
        mSQLS1.CommandText += "WHERE tracking.TIMEOUT BETWEEN '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND tracking.station  = '" & tStation1 & "' "
        If Not String.IsNullOrEmpty(tModel) Then
            mSQLS1.CommandText += " AND lot.model = '" & tModel & "' "
        End If
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select lot.model,model_station_paravalue.cf01,model.modelname,scrap_tracking.sn,scrap_tracking.station,station.stationname_cn,"
        mSQLS1.CommandText += "scrap_tracking.timein,scrap_tracking.timeout,scrap_tracking.users,users.name,lot.route, s1.seq,s2.station from scrap_tracking  "
        mSQLS1.CommandText += "left join lot on scrap_tracking.lot = lot.lot left join model on lot.model = model.model "
        mSQLS1.CommandText += "left join station on scrap_tracking.station =station.station left join users on scrap_tracking.users = users.id "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue.profilename = 'ERP' and scrap_tracking.station = model_station_paravalue.station "
        mSQLS1.CommandText += "left join routing s1 on lot.route = s1.route and scrap_tracking.station = s1.station left join routing s2 on lot.route = s2.route and (s1.seq - 1 ) = s2.seq "
        mSQLS1.CommandText += "WHERE scrap_tracking.TIMEOUT BETWEEN  '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND scrap_tracking.station  = '" & tStation1 & "' "
        If Not String.IsNullOrEmpty(tModel) Then
            mSQLS1.CommandText += " AND lot.model = '" & tModel & "' "
        End If
        mSQLS1.CommandText += ") as AA order by model"

        mSQLS1.CommandTimeout = 300
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                LastStation = String.Empty
                LastStation = GetLastStation(mSQLReader.Item("sn"), mSQLReader.Item("timeout"))
                If LastStation = "0080" Then
                    Continue While
                End If
                Ws.Cells(LineZ, 9) = LastStation
                ERPPN = String.Empty
                ERPPN = GetLastStationERPPN(mSQLReader.Item("sn"), LastStation)
                Ws.Cells(LineZ, 1) = mSQLReader.Item("model")
                'Ws.Cells(LineZ, 2) = mSQLReader.Item("cf01")
                Ws.Cells(LineZ, 2) = ERPPN
                Ws.Cells(LineZ, 3) = mSQLReader.Item("modelname")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("sn")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("station") & " " & mSQLReader.Item("stationname_cn")
                Ws.Cells(LineZ, 6) = mSQLReader.Item("timein")
                Ws.Cells(LineZ, 7) = mSQLReader.Item("timeout")
                Ws.Cells(LineZ, 8) = mSQLReader.Item("users") & " " & mSQLReader.Item("name")
                'Ws.Cells(LineZ, 9) = mSQLReader.Item("fresh")
                'If CheckLastStation = True Then
                'Ws.Cells(LineZ, 9) = mSQLReader.Item("laststation")
                mSQLS2.CommandText = "INSERT INTO " & DBS & " VALUES ('" & mSQLReader.Item("model") & "','" & ERPPN & "','" & mSQLReader.Item("modelname") & "','"
                mSQLS2.CommandText += mSQLReader.Item("station") & " " & mSQLReader.Item("stationname_cn") & "',1,0)"
                Try
                    mSQLS2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
                'End If
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
        ' 20170518 程式
        oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 9))
        oRng.Merge()
        oRng.Interior.Color = Color.LightGray
        Ws.Cells(LineZ, 1) = "返工品"
        LineZ += 1
        mSQLS1.CommandText = "select lot.model,model_station_paravalue.cf01,model.modelname,tracking_dup.sn,tracking_dup.station,station.stationname_cn,"
        mSQLS1.CommandText += "tracking_dup.timein,tracking_dup.timeout,tracking_dup.users,users.name,lot.route, s1.seq,s2.station as laststation from tracking_dup "
        mSQLS1.CommandText += "left join lot on tracking_dup.lot = lot.lot left join model on lot.model = model.model "
        mSQLS1.CommandText += "left join station on tracking_dup.station =station.station left join users on tracking_dup.users = users.id "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue .profilename = 'ERP' and tracking_dup.station = model_station_paravalue.station "
        mSQLS1.CommandText += "left join routing s1 on lot.route = s1.route and tracking_dup.station = s1.station left join routing s2 on lot.route = s2.route and (s1.seq - 1 ) = s2.seq "
        mSQLS1.CommandText += "WHERE tracking_dup.TIMEOUT BETWEEN '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND tracking_dup.station  = '" & tStation1 & "' "
        If Not String.IsNullOrEmpty(tModel) Then
            mSQLS1.CommandText += " AND lot.model = '" & tModel & "' "
        End If
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                LastStation = String.Empty
                LastStation = GetLastStation(mSQLReader.Item("sn"), mSQLReader.Item("timeout"))
                If LastStation = "0080" Then
                    Continue While
                End If
                Ws.Cells(LineZ, 9) = LastStation
                ERPPN = String.Empty
                ERPPN = GetLastStationERPPN(mSQLReader.Item("sn"), LastStation)
                Ws.Cells(LineZ, 1) = mSQLReader.Item("model")
                'Ws.Cells(LineZ, 2) = mSQLReader.Item("cf01")
                Ws.Cells(LineZ, 2) = ERPPN
                Ws.Cells(LineZ, 3) = mSQLReader.Item("modelname")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("sn")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("station") & " " & mSQLReader.Item("stationname_cn")
                Ws.Cells(LineZ, 6) = mSQLReader.Item("timein")
                Ws.Cells(LineZ, 7) = mSQLReader.Item("timeout")
                Ws.Cells(LineZ, 8) = mSQLReader.Item("users") & " " & mSQLReader.Item("name")
                'Ws.Cells(LineZ, 9) = mSQLReader.Item("fresh")
                'If CheckLastStation = True Then
                'GetLastStation(mSQLReader.Item("sn"), mSQLReader.Item("timeout"))
                'Ws.Cells(LineZ, 9) = mSQLReader.Item("laststation")
                'End If
                mSQLS2.CommandText = "INSERT INTO " & DBS & " VALUES ('" & mSQLReader.Item("model") & "','" & ERPPN & "','" & mSQLReader.Item("modelname") & "','"
                mSQLS2.CommandText += mSQLReader.Item("station") & " " & mSQLReader.Item("stationname_cn") & "',0,1)"
                Try
                    mSQLS2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
        Ws.Cells(3, 2) = Now()
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        Ws.Name = "Summary"
        AdjustExcelFormat1()
        LineZ = 5
        'mSQLS1.CommandText = "SELECT model,cf01,modelname,station,stationname_cn,sum(t1) as t1,sum(t2) as t2 FROM ( "
        'mSQLS1.CommandText += "select lot.model,model_station_paravalue.cf01,model.modelname,tracking.station,station.stationname_cn,count(sn) as t1,0 as t2 "
        'mSQLS1.CommandText += "from tracking "
        'mSQLS1.CommandText += "left join lot on tracking.lot = lot.lot left join model on lot.model = model.model "
        'mSQLS1.CommandText += "left join station on tracking.station =station.station "
        'mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue .profilename = 'ERP' and tracking.station = model_station_paravalue.station "
        'mSQLS1.CommandText += "WHERE tracking.TIMEOUT BETWEEN '"
        'mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        'mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND tracking.station  = '" & tStation1 & "' "
        'If Not String.IsNullOrEmpty(tModel) Then
        '    mSQLS1.CommandText += " AND lot.model = '" & tModel & "' "
        'End If
        'mSQLS1.CommandText += "group by lot.model,model_station_paravalue.cf01,model.modelname,tracking.station,station.stationname_cn "
        'mSQLS1.CommandText += "union all "
        'mSQLS1.CommandText += "select lot.model,model_station_paravalue.cf01,model.modelname,tracking_dup.station,station.stationname_cn,0,count(sn) "
        'mSQLS1.CommandText += "from tracking_dup "
        'mSQLS1.CommandText += "left join lot on tracking_dup.lot = lot.lot left join model on lot.model = model.model "
        'mSQLS1.CommandText += "left join station on tracking_dup.station =station.station "
        'mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue .profilename = 'ERP' and tracking_dup.station = model_station_paravalue.station "
        'mSQLS1.CommandText += "WHERE tracking_dup.TIMEOUT BETWEEN '"
        'mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        'mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND tracking_dup.station  = '" & tStation1 & "' "
        'If Not String.IsNullOrEmpty(tModel) Then
        '    mSQLS1.CommandText += " AND lot.model = '" & tModel & "' "
        'End If
        'mSQLS1.CommandText += "group by lot.model,model_station_paravalue.cf01,model.modelname,tracking_dup.station,station.stationname_cn "
        'mSQLS1.CommandText += "union all "
        'mSQLS1.CommandText += "select lot.model,model_station_paravalue.cf01,model.modelname,scrap_tracking.station,station.stationname_cn,count(scrap_tracking.sn),0 "
        'mSQLS1.CommandText += "from scrap_tracking  "
        'mSQLS1.CommandText += "left join lot on scrap_tracking.lot = lot.lot left join model on lot.model = model.model "
        'mSQLS1.CommandText += "left join station on scrap_tracking.station =station.station "
        'mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue.profilename = 'ERP' and scrap_tracking.station = model_station_paravalue.station "
        'mSQLS1.CommandText += "WHERE scrap_tracking.TIMEOUT BETWEEN '"
        'mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        'mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND scrap_tracking.station  = '" & tStation1 & "' "
        'If Not String.IsNullOrEmpty(tModel) Then
        '    mSQLS1.CommandText += " AND lot.model = '" & tModel & "' "
        'End If
        'mSQLS1.CommandText += "group by lot.model,model_station_paravalue.cf01,model.modelname,scrap_tracking.station,station.stationname_cn "
        'mSQLS1.CommandText += ") as AA group by model,cf01,modelname,station,stationname_cn order by model"
        mSQLS1.CommandText = "select model,erppn,ERPDESC,station,sum(GQ) as GQ, sum(SQ) as SQ from " & DBS
        mSQLS1.CommandText += " group by model,erppn,ERPDESC,station order by model"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("model")
                'Ws.Cells(LineZ, 2) = mSQLReader.Item("cf01")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("ERPPN")
                'Ws.Cells(LineZ, 3) = mSQLReader.Item("modelname")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("ERPDESC")
                'Ws.Cells(LineZ, 4) = mSQLReader.Item("station") & " " & mSQLReader.Item("stationname_cn")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("station")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("GQ")
                Ws.Cells(LineZ, 6) = mSQLReader.Item("SQ")
                Ws.Cells(LineZ, 7) = mSQLReader.Item("GQ") + mSQLReader.Item("SQ")
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
        LineZ += 1
        Ws.Cells(LineZ, 1) = "交接人"
        Ws.Cells(LineZ, 3) = "接收人"
        oRng = Ws.Range(Ws.Cells(LineZ, 5), Ws.Cells(LineZ, 5))
        oRng.NumberFormatLocal = "@"
        Ws.Cells(LineZ, 5) = Now.ToString("yyyy/MM/dd HH:mm:ss")
        ' add by cloud 20180201
        mSQLS1.CommandText = "DROP TABLE " & DBS
        Try
            mSQLS1.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        Ws.Name = "Fail Detail"
        AdjustExcelFormat2()
        LineZ = 5
        mSQLS1.CommandText = "select lot.model,sn,failstation,a.cf01 ,rework,b.cf01  from failure left join lot on failure.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue a on lot.model = a.model and a.profilename = 'ERP' and failstation = a.station "
        mSQLS1.CommandText += "left join model_station_paravalue b on lot.model = b.model and b.profilename = 'ERP' and rework = b.station where failtime between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and rework <> 'SCRP' and a.cf01 <> b.cf01 "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select lot.model,failstation,a.cf01 ,rework,b.cf01,sn  from scrap_failure left join lot on scrap_failure.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue a on lot.model = a.model and a.profilename = 'ERP' and failstation = a.station "
        mSQLS1.CommandText += "left join model_station_paravalue b on lot.model = b.model and b.profilename = 'ERP' and rework = b.station where failtime between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and rework <> 'SCRP' and a.cf01 <> b.cf01 "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                For i As Int16 = 0 To 5 Step 1
                    Ws.Cells(LineZ, i + 1) = mSQLReader.Item(i)
                Next
                LineZ += 1
            End While
        End If
        mSQLReader.Close()

        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws.Activate()
        Ws.Name = "Fail Summary"
        AdjustExcelFormat3()
        LineZ = 5
        mSQLS1.CommandText = "select model,failstation,cf01,rework,cf01a,count(sn) from ( "
        mSQLS1.CommandText += "select lot.model,failstation,a.cf01 ,rework,b.cf01 as cf01a,sn  from failure left join lot on failure.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue a on lot.model = a.model and a.profilename = 'ERP' and failstation = a.station "
        mSQLS1.CommandText += "left join model_station_paravalue b on lot.model = b.model and b.profilename = 'ERP' and rework = b.station where failtime between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and rework <> 'SCRP' and a.cf01 <> b.cf01 "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select lot.model,failstation,a.cf01 ,rework,b.cf01,sn  from scrap_failure left join lot on scrap_failure.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue a on lot.model = a.model and a.profilename = 'ERP' and failstation = a.station "
        mSQLS1.CommandText += "left join model_station_paravalue b on lot.model = b.model and b.profilename = 'ERP' and rework = b.station where failtime between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and rework <> 'SCRP' and a.cf01 <> b.cf01 ) as ab group by model,failstation,cf01,rework,cf01a"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                For i As Int16 = 0 To 5 Step 1
                    Ws.Cells(LineZ, 1 + i) = mSQLReader.Item(i)
                Next
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
        LineZ += 1
        Ws.Cells(LineZ, 1) = "交接人"
        Ws.Cells(LineZ, 3) = "接收人"
        oRng = Ws.Range(Ws.Cells(LineZ, 5), Ws.Cells(LineZ, 5))
        oRng.NumberFormatLocal = "@"
        Ws.Cells(LineZ, 5) = Now.ToString("yyyy/MM/dd HH:mm:ss")


        mSQLS1.CommandText = "DROP TABLE " & DBS
        Try
            mSQLS1.ExecuteNonQuery()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 25
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("C1", "I1")
        oRng.Merge()
        oRng = Ws.Range("C2", "I2")
        oRng.Merge()
        oRng = Ws.Range("C3", "I3")
        oRng.Merge()
        Ws.Cells(1, 1) = "取数开始时间"
        Ws.Cells(2, 1) = "取数结束时间"
        Ws.Cells(3, 1) = "报表打印时间"
        Ws.Cells(1, 3) = "东莞艾可迅复合材料有限公司"
        Ws.Cells(2, 3) = "Dongguan Action Composites LTD Co."
        Ws.Cells(3, 3) = "客制交接报表"
        oRng = Ws.Range("A4", "I4")
        oRng.Merge()
        oRng.Interior.Color = Color.LightGray
        Ws.Cells(4, 1) = "正常品"
        Ws.Cells(5, 1) = "品号"
        Ws.Cells(5, 2) = "ERP料号"
        Ws.Cells(5, 3) = "产品描述"
        Ws.Cells(5, 4) = "序列号"
        Ws.Cells(5, 5) = "工作站"
        Ws.Cells(5, 6) = "开始时间"
        Ws.Cells(5, 7) = "完成时间"
        Ws.Cells(5, 8) = "作业员"
        Ws.Cells(5, 9) = "上一工站"
        oRng = Ws.Range("I4", "I4")
        oRng.EntireColumn.NumberFormatLocal = "@"
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 25
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "E1")
        oRng.Merge()
        oRng = Ws.Range("A2", "E2")
        oRng.Merge()
        oRng = Ws.Range("A3", "E3")
        oRng.Merge()
        Ws.Cells(1, 1) = "东莞艾可迅复合材料有限公司"
        Ws.Cells(2, 1) = "Dongguan Action Composites LTD Co."
        Ws.Cells(3, 1) = "客制交接报表"
        Ws.Cells(4, 1) = "品号"
        Ws.Cells(4, 2) = "ERP料号"
        Ws.Cells(4, 3) = "产品描述"
        'Ws.Cells(4, 4) = "序列号"
        Ws.Cells(4, 4) = "工作站"
        Ws.Cells(4, 5) = "良品数量"
        Ws.Cells(4, 6) = "返工品数量"
        Ws.Cells(4, 7) = "总合计数量"
    End Sub
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 25
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "E1")
        oRng.Merge()
        oRng = Ws.Range("A2", "E2")
        oRng.Merge()
        oRng = Ws.Range("A3", "E3")
        oRng.Merge()
        oRng = Ws.Range("C4", "C4")
        oRng.EntireColumn.NumberFormatLocal = "@"
        oRng = Ws.Range("E4", "E4")
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(1, 1) = "东莞艾可迅复合材料有限公司"
        Ws.Cells(2, 1) = "Dongguan Action Composites LTD Co."
        Ws.Cells(3, 1) = "失败返工明细"
        Ws.Cells(4, 1) = "品号"
        Ws.Cells(4, 2) = "序列号"
        Ws.Cells(4, 3) = "失败工站"
        Ws.Cells(4, 4) = "失败ERP料号"
        Ws.Cells(4, 5) = "返工工站"
        Ws.Cells(4, 6) = "返工ERP料号"
    End Sub
    Private Sub AdjustExcelFormat3()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 25
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "E1")
        oRng.Merge()
        oRng = Ws.Range("A2", "E2")
        oRng.Merge()
        oRng = Ws.Range("A3", "E3")
        oRng.Merge()
        oRng = Ws.Range("B4", "B4")
        oRng.EntireColumn.NumberFormatLocal = "@"
        oRng = Ws.Range("D4", "D4")
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(1, 1) = "东莞艾可迅复合材料有限公司"
        Ws.Cells(2, 1) = "Dongguan Action Composites LTD Co."
        Ws.Cells(3, 1) = "失败返工汇总"
        Ws.Cells(4, 1) = "品号"
        Ws.Cells(4, 2) = "失败工站"
        Ws.Cells(4, 3) = "失败ERP料号"
        Ws.Cells(4, 4) = "返工工站"
        Ws.Cells(4, 5) = "返工ERP料号"
        Ws.Cells(4, 6) = "数量"
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Custom Transfer Report"
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
    Private Function GetLastStation(ByVal sn As String, ByVal lasttimeout As Date)
        mSQLS2.CommandText = "select top 1 station from ( select sn,station,timeout from tracking where sn = '" & sn & "' and timeout < '"
        mSQLS2.CommandText += lasttimeout.ToString("yyyy/MM/dd HH:mm:ss") & "' union all select sn,station,timeout from tracking_dup where sn = '"
        mSQLS2.CommandText += sn & "' and timeout < '" & lasttimeout.ToString("yyyy/MM/dd HH:mm:ss") & "' union all select sn,station,timeout from scrap_tracking where sn = '"
        mSQLS2.CommandText += sn & "' and timeout < '" & lasttimeout.ToString("yyyy/MM/dd HH:mm:ss") & "' ) as AE order by timeout desc"
        Dim LastS As String = mSQLS2.ExecuteScalar()
        'Ws.Cells(LineZ, 9) = LastS
        Return LastS
    End Function
    Private Function GetLastStationERPPN(ByVal sn As String, ByVal station As String)
        mSQLS2.CommandText = "select cf01 from sn left join lot on sn.lot = lot.lot left join model_station_paravalue on lot.model = model_station_paravalue.model "
        mSQLS2.CommandText += "and model_station_paravalue.profilename = 'ERP' and model_station_paravalue.station = '" & station & "' where sn = '"
        mSQLS2.CommandText += sn & "' "
        mSQLS2.CommandText += "union all "
        mSQLS2.CommandText += "select cf01 from scrap_sn left join lot on scrap_sn.lot = lot.lot left join model_station_paravalue on lot.model = model_station_paravalue.model "
        mSQLS2.CommandText += "and model_station_paravalue.profilename = 'ERP' and model_station_paravalue.station = '" & station & "' where sn = '"
        mSQLS2.CommandText += sn & "' "
        Dim ERP_PN As String = String.Empty
        If IsDBNull(mSQLS2.ExecuteScalar()) Then
            ERP_PN = String.Empty
        Else
            ERP_PN = mSQLS2.ExecuteScalar()
        End If
        Return ERP_PN
    End Function
End Class