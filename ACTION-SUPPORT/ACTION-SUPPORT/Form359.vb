Public Class Form359
    Dim oConnectionBuilder As New Oracle.ManagedDataAccess.Client.OracleConnectionStringBuilder
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oSQLS1 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oSQLReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim mConnectionBuilder As New SqlClient.SqlConnectionStringBuilder
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLS3 As New SqlClient.SqlCommand    '200312 add by Brady
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim D1 As Date   '週期第一天
    Dim D2 As Date   '週期最後一天
    Dim D3 As Date   ' 前一週週日
    Dim D4 As Date
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Dim LastStation As String = String.Empty   '200312 add by Brady
    Dim ERPPN As String = String.Empty         '200312 add by Brady

    Private Property xlOpenXMLWorkbook As Object

    Private Sub Form359_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'CheckForIllegalCrossThreadCalls = False
        D1 = Convert.ToDateTime(Today().Year & "/" & Today.Month() & "/" & Today.Day & " 08:00:00")
        'D1 = "2018/05/10 08:00:00"   

        D2 = D1.AddDays(-2)
        D3 = D1.AddDays(-1)

        Me.DateTimePicker1.Value = Convert.ToDateTime(D2)
        Me.DateTimePicker2.Value = Convert.ToDateTime(D3)

        mConnectionBuilder.DataSource = "192.168.10.254"
        mConnectionBuilder.InitialCatalog = "IQMES3"
        mConnectionBuilder.IntegratedSecurity = False
        mConnectionBuilder.UserID = "sa"
        mConnectionBuilder.Password = "p@$$w0rd"
        mConnectionBuilder.MultipleActiveResultSets = True
        mConnection.ConnectionString = mConnectionBuilder.ConnectionString

        If mConnection.State <> ConnectionState.Open Then
            mConnection.Open()
            mSQLS1.Connection = mConnection
            mSQLS1.CommandType = CommandType.Text
            mSQLS1.CommandTimeout = 1800
            mSQLS2.Connection = mConnection
            mSQLS2.CommandType = CommandType.Text
        End If
        oConnectionBuilder.DataSource = "topprod"
        oConnectionBuilder.PersistSecurityInfo = False
        oConnectionBuilder.UserID = "actiontest"
        oConnectionBuilder.Password = "actiontest"
        oConnection.ConnectionString = oConnectionBuilder.ConnectionString
        If oConnection.State <> ConnectionState.Open Then
            oConnection.Open()
            oSQLS1.Connection = oConnection
            oSQLS1.CommandType = CommandType.Text
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        TimeS1 = DateTimePicker1.Value
        TimeS2 = DateTimePicker2.Value

        D2 = TimeS1.AddDays(0)
        D3 = TimeS2.AddDays(0)
        D4 = TimeS2.AddDays(-1)

        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        Ws.Name = "良品"
        AdjustExcelFormat()

        '200312 add by Brady
        mSQLS1.CommandText = "select model,cf01,station,stationname_cn,sum(t1) as t1 from ( "
        mSQLS1.CommandText += "select lot.model,cf01,tracking.station,station.stationname_cn,count(sn) as t1 from tracking left join lot on tracking.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and tracking.station = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
        mSQLS1.CommandText += "left join station on tracking.station = station.station "
        mSQLS1.CommandText += "where tracking.timeout between '" & D2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & D3.ToString("yyyy/MM/dd HH:mm:ss") & "' and (cf01 like '%31' or cf01 like '%32' or cf01 like '%35' or cf01 like '%32A' or cf01 like '%35A') "

        '181212 add by Brady 
        'mSQLS1.CommandText += "and tracking.station in ('0112','0113','0180','0193','0165','0395','0478','0405') group by lot.model,cf01,tracking.station,stationname_cn "
        mSQLS1.CommandText += "and tracking.station in ('0112','0113','0180','0193','0165','0395','0478','0405','0146') group by lot.model,cf01,tracking.station,stationname_cn "
        '181212 add by Brady END

        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select lot.model,cf01,tracking_dup.station,station.stationname_cn,count(sn) from tracking_dup left join lot on tracking_dup.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and tracking_dup.station = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
        mSQLS1.CommandText += "left join station on tracking_dup.station = station.station "
        mSQLS1.CommandText += "where tracking_dup.timeout between '" & D2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & D3.ToString("yyyy/MM/dd HH:mm:ss") & "' and (cf01 like '%31' or cf01 like '%32' or cf01 like '%35' or cf01 like '%32A' or cf01 like '%35A') "

        '181212 add by Brady
        'mSQLS1.CommandText += "and tracking_dup.station in ('0112','0113','0180','0193','0165','0395','0478','0405') group by lot.model,cf01,tracking_dup.station,stationname_cn "
        mSQLS1.CommandText += "and tracking_dup.station in ('0112','0113','0180','0193','0165','0395','0478','0405','0146') group by lot.model,cf01,tracking_dup.station,stationname_cn "
        '181212 add by Brady END

        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select lot.model,cf01,scrap_tracking.station,station.stationname_cn,count(sn) from scrap_tracking left join lot on scrap_tracking.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and scrap_tracking.station = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
        mSQLS1.CommandText += "left join station on scrap_tracking.station = station.station "
        mSQLS1.CommandText += "where scrap_tracking.timeout between '" & D2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & D3.ToString("yyyy/MM/dd HH:mm:ss") & "' and (cf01 like '%31' or cf01 like '%32' or cf01 like '%35' or cf01 like '%32A' or cf01 like '%35A') "

        '181212 add by Brady
        'mSQLS1.CommandText += "and scrap_tracking.station in ('0112','0113','0180','0193','0165','0395','0478','0405') group by lot.model,cf01,scrap_tracking.station,stationname_cn ) as AB group by model,cf01,station,stationname_cn order by cf01"
        mSQLS1.CommandText += "and scrap_tracking.station in ('0112','0113','0180','0193','0165','0395','0478','0405','0146') group by lot.model,cf01,scrap_tracking.station,stationname_cn ) as AB group by model,cf01,station,stationname_cn order by cf01"
        '181212 add by Brady END

        'mSQLS1.CommandText = "select model,cf01,station,stationname_cn,max(sn) as sn,max(timeout) as timeout,sum(t1) as t1 from ( "
        'mSQLS1.CommandText += "select lot.model,cf01,tracking.station,station.stationname_cn,max(sn) as sn,max(timeout) as timeout,count(sn) as t1 from tracking left join lot on tracking.lot = lot.lot "
        'mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and tracking.station = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
        'mSQLS1.CommandText += "left join station on tracking.station = station.station "
        'mSQLS1.CommandText += "where tracking.timeout between '" & D2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & D3.ToString("yyyy/MM/dd HH:mm:ss") & "' and (cf01 like '%31' or cf01 like '%32' or cf01 like '%35' or cf01 like '%32A' or cf01 like '%35A') "
        'mSQLS1.CommandText += "and tracking.station in ('0112','0113','0180','0193','0165','0395','0478','0405','0146') group by lot.model,cf01,tracking.station,stationname_cn "
        'mSQLS1.CommandText += "union all "
        'mSQLS1.CommandText += "select lot.model,cf01,tracking_dup.station,station.stationname_cn,max(sn),max(timeout),count(sn) from tracking_dup left join lot on tracking_dup.lot = lot.lot "
        'mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and tracking_dup.station = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
        'mSQLS1.CommandText += "left join station on tracking_dup.station = station.station "
        'mSQLS1.CommandText += "where tracking_dup.timeout between '" & D2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & D3.ToString("yyyy/MM/dd HH:mm:ss") & "' and (cf01 like '%31' or cf01 like '%32' or cf01 like '%35' or cf01 like '%32A' or cf01 like '%35A') "
        'mSQLS1.CommandText += "and tracking_dup.station in ('0112','0113','0180','0193','0165','0395','0478','0405','0146') group by lot.model,cf01,tracking_dup.station,stationname_cn "
        'mSQLS1.CommandText += "union all "
        'mSQLS1.CommandText += "select lot.model,cf01,scrap_tracking.station,station.stationname_cn,max(sn),max(timeout),count(sn) from scrap_tracking left join lot on scrap_tracking.lot = lot.lot "
        'mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and scrap_tracking.station = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
        'mSQLS1.CommandText += "left join station on scrap_tracking.station = station.station "
        'mSQLS1.CommandText += "where scrap_tracking.timeout between '" & D2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & D3.ToString("yyyy/MM/dd HH:mm:ss") & "' and (cf01 like '%31' or cf01 like '%32' or cf01 like '%35' or cf01 like '%32A' or cf01 like '%35A') "
        'mSQLS1.CommandText += "and scrap_tracking.station in ('0112','0113','0180','0193','0165','0395','0478','0405','0146') group by lot.model,cf01,scrap_tracking.station,stationname_cn ) as AB group by model,cf01,station,stationname_cn order by cf01"
        '200312 add by Brady END

        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()

                '200312 add by Brady
                Ws.Cells(LineZ, 1) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("cf01")
                'LastStation = String.Empty
                'LastStation = GetLastStation(mSQLReader.Item("sn"), mSQLReader.Item("timeout"))
                'If LastStation = "0080" Then
                '    Continue While
                'End If
                'ERPPN = String.Empty
                'ERPPN = GetLastStationERPPN(mSQLReader.Item("sn"), LastStation)
                'Ws.Cells(LineZ, 1) = mSQLReader.Item("model")
                'Ws.Cells(LineZ, 2) = ERPPN
                '200312 add by Brady END

                Ws.Cells(LineZ, 3) = mSQLReader.Item("station") & " " & mSQLReader.Item("stationname_cn")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("t1")
                Ws.Cells(LineZ, 5) = GetERPQ(mSQLReader.Item("cf01"))
                Ws.Cells(LineZ, 6) = "=D" & LineZ & "-E" & LineZ
                LineZ += 1
            End While
        End If
        mSQLReader.Close()

        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        Ws.Name = "报废品"
        AdjustExcelFormat2()

        mSQLS1.CommandText = "select lot.model,cf01,scrap_sn.updatedstation,station.stationname_cn,count(scrap.sn) as t1   from scrap left join scrap_sn on scrap.sn = scrap_sn.sn "
        mSQLS1.CommandText += "left join lot on scrap.lot = lot.lot left join model_station_paravalue on lot.model = model_station_paravalue.model and scrap_sn.updatedstation = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
        mSQLS1.CommandText += "left join station on scrap_sn.updatedstation = station.station where scrap.datetime between '"
        mSQLS1.CommandText += D2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & D3.ToString("yyyy/MM/dd HH:mm:ss") & "' and scrap_sn.updatedstation in ('0330','0331') group by lot.model,cf01,scrap_sn.updatedstation,station.stationname_cn order by cf01"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("cf01")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("updatedstation") & " " & mSQLReader.Item("stationname_cn")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("t1")
                If Not IsDBNull(mSQLReader.Item("cf01")) Then
                    Ws.Cells(LineZ, 5) = GetERPQ2(mSQLReader.Item("cf01"))
                End If
                Ws.Cells(LineZ, 6) = "=D" & LineZ & "-E" & LineZ
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
        ' CNC

        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        Ws.Name = "CNC良品"
        AdjustExcelFormat()

        mSQLS1.CommandText = "select model,cf01,station,stationname_cn,sum(t1) as t1 from ( "
        mSQLS1.CommandText += "select lot.model,cf01,tracking.station,station.stationname_cn,count(sn) as t1 from tracking left join lot on tracking.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and tracking.station = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
        mSQLS1.CommandText += "left join station on tracking.station = station.station "
        mSQLS1.CommandText += "where tracking.timeout between '" & D2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & D3.ToString("yyyy/MM/dd HH:mm:ss") & "' and (cf01 like '%36') "
        mSQLS1.CommandText += "and tracking.station in ('0405','0478','0605') group by lot.model,cf01,tracking.station,stationname_cn "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select lot.model,cf01,tracking_dup.station,station.stationname_cn,count(sn) from tracking_dup left join lot on tracking_dup.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and tracking_dup.station = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
        mSQLS1.CommandText += "left join station on tracking_dup.station = station.station "
        mSQLS1.CommandText += "where tracking_dup.timeout between '" & D2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & D3.ToString("yyyy/MM/dd HH:mm:ss") & "' and (cf01 like '%36') "
        mSQLS1.CommandText += "and tracking_dup.station in ('0405','0478','0605') group by lot.model,cf01,tracking_dup.station,stationname_cn "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select lot.model,cf01,scrap_tracking.station,station.stationname_cn,count(sn) from scrap_tracking left join lot on scrap_tracking.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and scrap_tracking.station = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
        mSQLS1.CommandText += "left join station on scrap_tracking.station = station.station "
        mSQLS1.CommandText += "where scrap_tracking.timeout between '" & D2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & D3.ToString("yyyy/MM/dd HH:mm:ss") & "' and (cf01 like '%36') "
        mSQLS1.CommandText += "and scrap_tracking.station in ('0405','0478','0605') group by lot.model,cf01,scrap_tracking.station,stationname_cn ) as AB group by model,cf01,station,stationname_cn order by cf01"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("cf01")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("station") & " " & mSQLReader.Item("stationname_cn")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("t1")
                Ws.Cells(LineZ, 5) = GetERPQ(mSQLReader.Item("cf01"))
                Ws.Cells(LineZ, 6) = "=D" & LineZ & "-E" & LineZ
                LineZ += 1
            End While
        End If
        mSQLReader.Close()

        '第四頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(4)
        Ws.Activate()
        Ws.Name = "CNC报废品"
        AdjustExcelFormat2()

        mSQLS1.CommandText = "select lot.model,cf01,scrap_sn.updatedstation,station.stationname_cn,count(scrap.sn) as t1   from scrap left join scrap_sn on scrap.sn = scrap_sn.sn "
        mSQLS1.CommandText += "left join lot on scrap.lot = lot.lot left join model_station_paravalue on lot.model = model_station_paravalue.model and scrap_sn.updatedstation = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
        mSQLS1.CommandText += "left join station on scrap_sn.updatedstation = station.station where scrap.datetime between '"
        mSQLS1.CommandText += D2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & D3.ToString("yyyy/MM/dd HH:mm:ss") & "' and scrap_sn.updatedstation in ('0380','0530') group by lot.model,cf01,scrap_sn.updatedstation,station.stationname_cn order by cf01"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("cf01")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("updatedstation") & " " & mSQLReader.Item("stationname_cn")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("t1")
                If Not IsDBNull(mSQLReader.Item("cf01")) Then
                    Ws.Cells(LineZ, 5) = GetERPQ2(mSQLReader.Item("cf01"))
                End If

                Ws.Cells(LineZ, 6) = "=D" & LineZ & "-E" & LineZ
                LineZ += 1
            End While
        End If
        mSQLReader.Close()

        '第五頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(5)
        Ws.Activate()
        Ws.Name = "补土良品及返工"
        AdjustExcelFormat()

        mSQLS1.CommandText = "select model,cf01,station,stationname_cn,sum(t1) as t1 from ( "
        mSQLS1.CommandText += "select lot.model,cf01,tracking.station,station.stationname_cn,count(sn) as t1 from tracking left join lot on tracking.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and tracking.station = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
        mSQLS1.CommandText += "left join station on tracking.station = station.station "
        mSQLS1.CommandText += "where tracking.timeout between '" & D2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & D3.ToString("yyyy/MM/dd HH:mm:ss") & "' and ((cf01 like '%61') or (cf01 like '%61A')) "
        mSQLS1.CommandText += "and tracking.station in ('0455','0478','0605','0673') group by lot.model,cf01,tracking.station,stationname_cn "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select lot.model,cf01,tracking_dup.station,station.stationname_cn,count(sn) from tracking_dup left join lot on tracking_dup.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and tracking_dup.station = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
        mSQLS1.CommandText += "left join station on tracking_dup.station = station.station "
        mSQLS1.CommandText += "where tracking_dup.timeout between '" & D2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & D3.ToString("yyyy/MM/dd HH:mm:ss") & "' and ((cf01 like '%61') or (cf01 like '%61A')) "
        mSQLS1.CommandText += "and tracking_dup.station in ('0455','0478','0605','0673') group by lot.model,cf01,tracking_dup.station,stationname_cn "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select lot.model,cf01,scrap_tracking.station,station.stationname_cn,count(sn) from scrap_tracking left join lot on scrap_tracking.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and scrap_tracking.station = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
        mSQLS1.CommandText += "left join station on scrap_tracking.station = station.station "
        mSQLS1.CommandText += "where scrap_tracking.timeout between '" & D2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & D3.ToString("yyyy/MM/dd HH:mm:ss") & "' and ((cf01 like '%61') or (cf01 like '%61A')) "
        mSQLS1.CommandText += "and scrap_tracking.station in ('0455','0478','0605','0673') group by lot.model,cf01,scrap_tracking.station,stationname_cn "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select model,cf01,failstation, stationname_cn,count(sn) from ( "
        mSQLS1.CommandText += "select lot.model,a.cf01 ,failstation,stationname_cn,sn  from failure left join lot on failure.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue a on lot.model = a.model and a.profilename = 'ERP' and failstation = a.station "
        mSQLS1.CommandText += "left join model_station_paravalue b on lot.model = b.model and b.profilename = 'ERP' and rework = b.station "
        mSQLS1.CommandText += "left join station on failstation = station.station where failtime between '" & D2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & D3.ToString("yyyy/MM/dd HH:mm:ss") & "' and failstation in ('0475') and rework <> 'SCRP' and a.cf01 <> b.cf01 "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select lot.model,a.cf01 ,failstation,stationname_cn,sn  from scrap_failure left join lot on scrap_failure.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue a on lot.model = a.model and a.profilename = 'ERP' and failstation = a.station "
        mSQLS1.CommandText += "left join model_station_paravalue b on lot.model = b.model and b.profilename = 'ERP' and rework = b.station "
        mSQLS1.CommandText += "left join station on failstation = station.station where failtime between '" & D2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & D3.ToString("yyyy/MM/dd HH:mm:ss") & "' and failstation in ('0475') and rework <> 'SCRP' and a.cf01 <> b.cf01 ) as ab group by model,cf01,failstation ,stationname_cn "
        mSQLS1.CommandText += ") as AB group by model,cf01,station,stationname_cn order by cf01"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("cf01")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("station") & " " & mSQLReader.Item("stationname_cn")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("t1")
                Ws.Cells(LineZ, 5) = GetERPQ(mSQLReader.Item("cf01"))
                Ws.Cells(LineZ, 6) = "=D" & LineZ & "-E" & LineZ
                LineZ += 1
            End While
        End If
        mSQLReader.Close()

        '第六頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(6)
        Ws.Activate()
        Ws.Name = "补土报废品"
        AdjustExcelFormat2()

        mSQLS1.CommandText = "select lot.model,cf01,scrap_sn.updatedstation,station.stationname_cn,count(scrap.sn) as t1   from scrap left join scrap_sn on scrap.sn = scrap_sn.sn "
        mSQLS1.CommandText += "left join lot on scrap.lot = lot.lot left join model_station_paravalue on lot.model = model_station_paravalue.model and scrap_sn.updatedstation = model_station_paravalue.station and model_station_paravalue.profilename = 'ERP' "
        mSQLS1.CommandText += "left join station on scrap_sn.updatedstation = station.station where scrap.datetime between '"
        mSQLS1.CommandText += D2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & D3.ToString("yyyy/MM/dd HH:mm:ss") & "' and scrap_sn.updatedstation in ('0410','0415','0417','0420','0430','0440','0445','0450','0460','0465','0470','0475') "
        mSQLS1.CommandText += "and  (cf01 like '%61' or cf01 like '%61A'or cf01 like '%61B')  group by lot.model,cf01,scrap_sn.updatedstation,station.stationname_cn order by cf01"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("cf01")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("updatedstation") & " " & mSQLReader.Item("stationname_cn")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("t1")
                If Not IsDBNull(mSQLReader.Item("cf01")) Then
                    Ws.Cells(LineZ, 5) = GetERPQ2(mSQLReader.Item("cf01"))
                End If
                Ws.Cells(LineZ, 6) = "=D" & LineZ & "-E" & LineZ
                LineZ += 1
            End While
        End If
        mSQLReader.Close()


        mConnection.Close()
        oConnection.Close()

        ''Dim FileName1 As String = "E:\SHARE DRIVER\A08_MP&L_資材部\09.Share Folder 共用资料\仓库进料日报表\" & Today.AddDays(-1).Date.ToString("yyyy-MM-dd") & "-Receive.xlsx"
        'Dim FileName1 As String = "c:\temp\" & D2.ToString("yyyy-MM-dd") & "-MESERP入库数量比较.xlsx"
        'Try
        '    Ws.SaveAs(FileName1, xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing)
        '    xWorkBook.Saved = True
        '    xWorkBook.Close()
        '    xExcel.Quit()
        'Catch ex As Exception
        '    MsgBox(ex.Message())
        'End Try

        'Try
        '    Module1.KillExcelProcess(OldExcel)
        'Catch ex As Exception

        'End Try

        SaveFileDialog1.FileName = D2.ToString("yyyy-MM-dd") & "至" & D3.ToString("yyyy-MM-dd") & " MES&ERP入库数量比较"
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
        If oConnection.State = ConnectionState.Open Then
            Try
                oConnection.Close()
                Module1.KillExcelProcess(OldExcel)
                MsgBox("Finished")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If

        ' try to send out
        'MailSend(FileName1)
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "F1")
        oRng.EntireColumn.ColumnWidth = 21
        oRng.Merge()
        oRng = Ws.Range("A2", "F2")
        oRng.Merge()
        oRng = Ws.Range("C3", "C3")
        oRng.EntireColumn.NumberFormatLocal = "@"

        Ws.Cells(1, 1) = "MES接收工站与ERP入库数量比对（良品）"
        Ws.Cells(2, 1) = "数据导出日期/时间：" & D2.ToString("yyyy/MM/dd HH:mm:ss") & " ~ " & D3.ToString("yyyy/MM/dd HH:mm:ss")
        Ws.Cells(3, 1) = "品号"
        Ws.Cells(3, 2) = "ERP料号"
        Ws.Cells(3, 3) = "工作站"
        Ws.Cells(3, 4) = "MES接收总合计数量"
        Ws.Cells(3, 5) = "ERP入库数量"
        Ws.Cells(3, 6) = "差异"
        LineZ = 4
    End Sub
    Private Function GetERPQ(ByVal ima01 As String)
        '181225 add by Brady
        'oSQLS1.CommandText = "select nvl(sum(sfv09),0) from sfv_file,sfu_file where sfv01 =sfu01 and sfv04 = '" & ima01 & "' and sfu02 = to_date('"
        'oSQLS1.CommandText += D2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and sfupost = 'Y'"
        oSQLS1.CommandText = "select nvl(sum(sfv09),0) from sfv_file,sfu_file where sfv01 =sfu01 and sfv04 = '" & ima01 & "' and sfu02 between to_date('"
        oSQLS1.CommandText += D2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & D4.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and sfupost = 'Y'"
        '181225 add by Brady END
        Dim GEQ As Decimal = oSQLS1.ExecuteScalar()
        Return GEQ
    End Function
    Public Sub KillExcelProcess(ByVal oldExcel() As Process)
        Dim NewExcelProcess() As Process = Process.GetProcessesByName("Excel")
        For i As Int16 = 0 To NewExcelProcess.Length - 1 Step 1
            Dim FoundExcel As Boolean = False
            Dim NewProcessInteger As Integer = NewExcelProcess(i).Id
            For j As Int16 = 0 To oldExcel.Length - 1 Step 1
                Dim OldProcessIntger As Integer = oldExcel(j).Id
                If NewProcessInteger = OldProcessIntger Then
                    FoundExcel = True
                    Exit For
                End If
            Next
            If FoundExcel = False Then
                Process.GetProcessById(NewExcelProcess(i).Id).Kill()
                Exit For
            End If
        Next
    End Sub
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "F1")
        oRng.EntireColumn.ColumnWidth = 21
        oRng.Merge()
        oRng = Ws.Range("A2", "F2")
        oRng.Merge()
        oRng = Ws.Range("C3", "C3")
        oRng.EntireColumn.NumberFormatLocal = "@"

        Ws.Cells(1, 1) = "MES接收工站与ERP入库数量比对（报废品）"
        Ws.Cells(2, 1) = "数据导出日期/时间：" & D2.ToString("yyyy/MM/dd HH:mm:ss") & " ~ " & D3.ToString("yyyy/MM/dd HH:mm:ss")
        Ws.Cells(3, 1) = "品号"
        Ws.Cells(3, 2) = "ERP料号"
        Ws.Cells(3, 3) = "工作站"
        Ws.Cells(3, 4) = "MES接收总合计数量"
        Ws.Cells(3, 5) = "ERP入库数量"
        Ws.Cells(3, 6) = "差异"
        LineZ = 4
    End Sub
    Private Function GetERPQ2(ByVal ima01 As String)
        '181225 add by Brady
        'oSQLS1.CommandText = "select nvl(sum(sfvud07),0) from sfv_file,sfu_file where sfv01 =sfu01 and sfv04 = '" & ima01 & "' and sfu02 = to_date('"
        'oSQLS1.CommandText += D2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and sfupost = 'Y'"
        oSQLS1.CommandText = "select nvl(sum(sfvud07),0) from sfv_file,sfu_file where sfv01 =sfu01 and sfv04 = '" & ima01 & "' and sfu02 between to_date('"
        oSQLS1.CommandText += D2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & D4.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and sfupost = 'Y'"
        '181225 add by Brady END
        Dim GEQ As Decimal = oSQLS1.ExecuteScalar()
        Return GEQ
    End Function
    '200312 add by Brady
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
    '200312 add by Brady END
    'Public Sub MailSend(ByVal FileName)
    '    Dim MS As New System.Net.Mail.MailMessage
    '    Dim MA As New System.Net.Mail.MailAddress("action.server@action-composites.com.cn")
    '    MS.From = MA
    '    MS.Subject = "MES & ERP Compare Report " & D2.ToString("yyyy-MM-dd")
    '    MS.To.Add("wen.lee@action-composites.com.cn")
    '    MS.To.Add("sanding@action-composites.com.cn")
    '    MS.To.Add("DAC_Costing@action-composites.com.cn")
    '    MS.To.Add("fanpeixia@action-composites.com.cn")

    '    MS.IsBodyHtml = True
    '    Dim MAM As New System.Net.Mail.Attachment(FileName)
    '    MS.Attachments.Add(MAM)
    '    ' 信件做好了
    '    Dim SMT As New System.Net.Mail.SmtpClient("smtp.action-composites.com.cn")
    '    SMT.UseDefaultCredentials = True
    '    'SMT.PickupDirectoryLocation = "C:\temp\ab"
    '    Dim UAP As New System.Net.NetworkCredential()
    '    UAP.UserName = "action.server@action-composites.com.cn"
    '    UAP.Password = "action@2017"
    '    SMT.Credentials = UAP

    '    Try
    '        SMT.Send(MS)
    '    Catch ex As Exception
    '        MsgBox(ex.Message())
    '    End Try
    'End Sub


End Class