Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form43
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim mSQLReader2 As SqlClient.SqlDataReader
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim g_success As Boolean = True
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form43_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommander2.Connection = oConnection
                oCommander2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
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
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommander2.Connection = oConnection
                oCommander2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        GetReady()
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        ' First Page
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        AdjustExcelFormat()
        oCommand.CommandText = "select erppn,sum(c1) as c1,sum(c2) as c2,sum(c3) as c3,sum(c4) as c4,sum(c5) as c5,sum(c6) as c6,sum(c7) as c7,sum(c8) as c8 "
        oCommand.CommandText += ",sum(c9) as c9,sum(c10) as c10,sum(c11) as c11,sum(c12) as c12,sum(c13) as c13,sum(c14) as c14,sum(c15) as c15,sum(c16) as c16,sum(c17) as c17"
        oCommand.CommandText += ",sum(c18) as c18,sum(c19) as c19,sum(c20) as c20,sum(c21) as c21,sum(c22) as c22,sum(c23) as c23,sum(c24) as c24,sum(c25) as c25,sum(c26) as c26 from ( "
        oCommand.CommandText += "select erppn,(case when ntype = '2' and warehouse = 'D353201' then sum(t1) else 0 end) as c1,"
        oCommand.CommandText += "             (case when ntype = '0' and warehouse = 'D353201' then sum(t1) else 0 end) as c2,"
        oCommand.CommandText += "             (case when ntype = '2' and warehouse = 'D353501' then sum(t1) else 0 end) as c3,"
        oCommand.CommandText += "             (case when ntype = '0' and warehouse = 'D353501' then sum(t1) else 0 end) as c4,"
        oCommand.CommandText += "             (case when ntype = '2' and warehouse = 'D353601' then sum(t1) else 0 end) as c5,"
        oCommand.CommandText += "             (case when ntype = '0' and warehouse = 'D353601' then sum(t1) else 0 end) as c6,"
        oCommand.CommandText += "             (case when ntype = '2' and warehouse = 'D356401' then sum(t1) else 0 end) as c7,"
        oCommand.CommandText += "             (case when ntype = '0' and warehouse = 'D356401' then sum(t1) else 0 end) as c8,"
        oCommand.CommandText += "             (case when ntype = '2' and warehouse = 'D356405' then sum(t1) else 0 end) as c9,"
        oCommand.CommandText += "             (case when ntype = '1' and warehouse = 'D356405' then sum(t1) else 0 end) as c10,"
        oCommand.CommandText += "             (case when ntype = '2' and warehouse = 'D356101' then sum(t1) else 0 end) as c11,"
        oCommand.CommandText += "             (case when ntype = '0' and warehouse = 'D356101' then sum(t1) else 0 end) as c12,"
        oCommand.CommandText += "             (case when ntype = '2' and warehouse = 'D356105' then sum(t1) else 0 end) as c13,"
        oCommand.CommandText += "             (case when ntype = '1' and warehouse = 'D356105' then sum(t1) else 0 end) as c14,"
        oCommand.CommandText += "             (case when ntype = '2' and warehouse = 'D356301' then sum(t1) else 0 end) as c15,"
        oCommand.CommandText += "             (case when ntype = '0' and warehouse = 'D356301' then sum(t1) else 0 end) as c16,"
        oCommand.CommandText += "             (case when ntype = '2' and warehouse = 'D356305' then sum(t1) else 0 end) as c17,"
        oCommand.CommandText += "             (case when ntype = '1' and warehouse = 'D356305' then sum(t1) else 0 end) as c18,"
        oCommand.CommandText += "             (case when ntype = '2' and warehouse = 'D356501' then sum(t1) else 0 end) as c19,"
        oCommand.CommandText += "             (case when ntype = '0' and warehouse = 'D356501' then sum(t1) else 0 end) as c20,"
        oCommand.CommandText += "             (case when ntype = '2' and warehouse = 'D356505' then sum(t1) else 0 end) as c21,"
        oCommand.CommandText += "             (case when ntype = '1' and warehouse = 'D356505' then sum(t1) else 0 end) as c22,"
        oCommand.CommandText += "             (case when ntype = '2' and warehouse = 'D356601' then sum(t1) else 0 end) as c23,"
        oCommand.CommandText += "             (case when ntype = '0' and warehouse = 'D356601' then sum(t1) else 0 end) as c24,"
        oCommand.CommandText += "             (case when ntype = '2' and warehouse = 'D146103' then sum(t1) else 0 end) as c25,"
        oCommand.CommandText += "             (case when ntype = '0' and warehouse = 'D146103' then sum(t1) else 0 end) as c26 from mes_temp2 group by erppn,ntype,warehouse ) group by erppn"
        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("erppn")
                If oReader.Item("c1") <> 0 Then
                    Ws.Cells(LineZ, 2) = oReader.Item("c1")
                    Ws.Cells(LineZ, 4) = "=B" & LineZ & "-C" & LineZ
                End If
                If oReader.Item("c2") <> 0 Then
                    Ws.Cells(LineZ, 3) = oReader.Item("c2")
                    Ws.Cells(LineZ, 4) = "=B" & LineZ & "-C" & LineZ
                End If
                If oReader.Item("c3") <> 0 Then
                    Ws.Cells(LineZ, 5) = oReader.Item("c3")
                    Ws.Cells(LineZ, 7) = "=E" & LineZ & "-F" & LineZ
                End If
                If oReader.Item("c4") <> 0 Then
                    Ws.Cells(LineZ, 6) = oReader.Item("c4")
                    Ws.Cells(LineZ, 7) = "=E" & LineZ & "-F" & LineZ
                End If
                If oReader.Item("c5") <> 0 Then
                    Ws.Cells(LineZ, 8) = oReader.Item("c5")
                    Ws.Cells(LineZ, 10) = "=H" & LineZ & "-I" & LineZ
                End If
                If oReader.Item("c6") <> 0 Then
                    Ws.Cells(LineZ, 9) = oReader.Item("c6")
                    Ws.Cells(LineZ, 10) = "=H" & LineZ & "-I" & LineZ
                End If
                If oReader.Item("c7") <> 0 Then
                    Ws.Cells(LineZ, 11) = oReader.Item("c7")
                    Ws.Cells(LineZ, 13) = "=K" & LineZ & "-L" & LineZ
                End If
                If oReader.Item("c8") <> 0 Then
                    Ws.Cells(LineZ, 12) = oReader.Item("c8")
                    Ws.Cells(LineZ, 13) = "=K" & LineZ & "-L" & LineZ
                End If
                If oReader.Item("c9") <> 0 Then
                    Ws.Cells(LineZ, 14) = oReader.Item("c9")
                    Ws.Cells(LineZ, 16) = "=N" & LineZ & "-O" & LineZ
                End If
                If oReader.Item("c10") <> 0 Then
                    Ws.Cells(LineZ, 15) = oReader.Item("c10")
                    Ws.Cells(LineZ, 16) = "=N" & LineZ & "-O" & LineZ
                End If
                If oReader.Item("c11") <> 0 Then
                    Ws.Cells(LineZ, 17) = oReader.Item("c11")
                    Ws.Cells(LineZ, 19) = "=Q" & LineZ & "-R" & LineZ
                End If
                If oReader.Item("c12") <> 0 Then
                    Ws.Cells(LineZ, 18) = oReader.Item("c12")
                    Ws.Cells(LineZ, 19) = "=Q" & LineZ & "-R" & LineZ
                End If
                If oReader.Item("c13") <> 0 Then
                    Ws.Cells(LineZ, 20) = oReader.Item("c13")
                    Ws.Cells(LineZ, 22) = "=T" & LineZ & "-U" & LineZ
                End If
                If oReader.Item("c14") <> 0 Then
                    Ws.Cells(LineZ, 21) = oReader.Item("c14")
                    Ws.Cells(LineZ, 22) = "=T" & LineZ & "-U" & LineZ
                End If
                If oReader.Item("c15") <> 0 Then
                    Ws.Cells(LineZ, 23) = oReader.Item("c15")
                    Ws.Cells(LineZ, 25) = "=W" & LineZ & "-X" & LineZ
                End If
                If oReader.Item("c16") <> 0 Then
                    Ws.Cells(LineZ, 24) = oReader.Item("c16")
                    Ws.Cells(LineZ, 25) = "=W" & LineZ & "-X" & LineZ
                End If
                If oReader.Item("c17") <> 0 Then
                    Ws.Cells(LineZ, 26) = oReader.Item("c17")
                    Ws.Cells(LineZ, 28) = "=Z" & LineZ & "-AA" & LineZ
                End If
                If oReader.Item("c18") <> 0 Then
                    Ws.Cells(LineZ, 27) = oReader.Item("c18")
                    Ws.Cells(LineZ, 28) = "=Z" & LineZ & "-AA" & LineZ
                End If
                If oReader.Item("c19") <> 0 Then
                    Ws.Cells(LineZ, 29) = oReader.Item("c19")
                    Ws.Cells(LineZ, 31) = "=AC" & LineZ & "-AD" & LineZ
                End If
                If oReader.Item("c20") <> 0 Then
                    Ws.Cells(LineZ, 30) = oReader.Item("c20")
                    Ws.Cells(LineZ, 31) = "=AC" & LineZ & "-AD" & LineZ
                End If
                If oReader.Item("c21") <> 0 Then
                    Ws.Cells(LineZ, 32) = oReader.Item("c21")
                    Ws.Cells(LineZ, 34) = "=AF" & LineZ & "-AG" & LineZ
                End If
                If oReader.Item("c22") <> 0 Then
                    Ws.Cells(LineZ, 33) = oReader.Item("c22")
                    Ws.Cells(LineZ, 34) = "=AF" & LineZ & "-AG" & LineZ
                End If
                If oReader.Item("c23") <> 0 Then
                    Ws.Cells(LineZ, 35) = oReader.Item("c23")
                    Ws.Cells(LineZ, 37) = "=AI" & LineZ & "-AJ" & LineZ
                End If
                If oReader.Item("c24") <> 0 Then
                    Ws.Cells(LineZ, 36) = oReader.Item("c24")
                    Ws.Cells(LineZ, 37) = "=AI" & LineZ & "-AJ" & LineZ
                End If
                If oReader.Item("c25") <> 0 Then
                    Ws.Cells(LineZ, 38) = oReader.Item("c25")
                    Ws.Cells(LineZ, 40) = "=AL" & LineZ & "-AM" & LineZ
                End If
                If oReader.Item("c26") <> 0 Then
                    Ws.Cells(LineZ, 39) = oReader.Item("c26")
                    Ws.Cells(LineZ, 40) = "=AL" & LineZ & "-AM" & LineZ
                End If
                ' 20170323 增加加總
                Ws.Cells(LineZ, 41) = "=B" & LineZ & "+E" & LineZ & "+H" & LineZ & "+K" & LineZ & "+N" & LineZ & "+Q" & LineZ & "+T" & LineZ & "+W" & LineZ & "+Z" & LineZ & "+AC" & LineZ & "+AF" & LineZ & "+AI" & LineZ & "+AL" & LineZ
                Ws.Cells(LineZ, 42) = "=C" & LineZ & "+F" & LineZ & "+I" & LineZ & "+L" & LineZ & "+O" & LineZ & "+R" & LineZ & "+U" & LineZ & "+X" & LineZ & "+AA" & LineZ & "+AD" & LineZ & "+AG" & LineZ & "+AJ" & LineZ & "+AM" & LineZ
                Ws.Cells(LineZ, 43) = "=AO" & LineZ & "-AP" & LineZ
                LineZ += 1
            End While
        End If

        '第二頁 20170323 
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustExcelFormat1()
        mSQLS1.CommandText = "select sn1,sn2,cf01,profilename,laststation,station.stationname  from ( "
        mSQLS1.CommandText += "select sn1,sn2,cf01,profilename,laststation,(case when c1 is null then null when c1 in ('0080','0100','0110','0111') then '裁纱( Prepreg)' "
        mSQLS1.CommandText += "when c1 in ('0130','0140','0150','0151','0160','0170','0175','0180','0193') then '预型（Layup）' "
        mSQLS1.CommandText += "when c1 in ('0165','0190','0195','0200','0210','0215','0220','0223','0225','0230','0231','0240','0250','0260','0280','0300','0315','0316','0320','0321','0325','0326','0330','0331','0333','0390','0395') then '成型（Molding）' "
        mSQLS1.CommandText += "when c1 in ('0335','0340','0350','0360','0370','0380','0390','0495','0500','0510','0520','0530') then 'CNC' "
        mSQLS1.CommandText += "when c1 in ('0400','0435','0478','0480','0490','0492','0493','0605','0610','0620','0623','0627') then '胶合（Glueing)' "
        mSQLS1.CommandText += "when c1 in ('0629','0630','0633','0635','0640','0645') then '抛光（Polishing）' "
        mSQLS1.CommandText += "when c1 in ('0642','0650','0652','0657','0658','0659','0660','0665','0670','0673','0675','0680','0690') then '包装（Packing）' "
        mSQLS1.CommandText += "when c1 in ('0405','0410','0415','0417','0440','0475','0418','0420','0430','0445','0450','0455') then '研磨（Sanding）' "
        mSQLS1.CommandText += "when c1 in ('0460','0465','0540','0545','0567','0570','0575','0583','0584','0470','0550','0560','0563','0580','0585','0587','0590','0592','0595') then '涂装（Painting）' "
        mSQLS1.CommandText += "when c1 in ('0720','0730','0799') then '成品（FG）' end) as c2 from ( "
        mSQLS1.CommandText += "select InventoryCount.sn sn1,sn_temp.sn sn2,cf01,InventoryCount.profileName,"
        mSQLS1.CommandText += "(case when sn_temp.topreworkstation is null or sn_temp.topreworkstation = '' then sn_temp.updatedstation else sn_temp.topreworkstation end) as c1,laststation  "
        mSQLS1.CommandText += "from InventoryCount left join sn_temp on InventoryCount.sn = sn_temp.sn left join lot on sn_temp.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue.profilename  = 'ERP' "
        mSQLS1.CommandText += "and (case when sn_temp.topreworkstation is null or sn_temp.topreworkstation = '' then sn_temp.updatedstation else sn_temp.topreworkstation end ) = model_station_paravalue.station "
        mSQLS1.CommandText += ") as AA ) as AB,station where profileName <> c2 and ab.laststation  = station.station"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("sn1")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("cf01")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("profilename")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("laststation")
                Ws.Cells(LineZ, 6) = mSQLReader.Item("stationname")
                GetQCStatus(mSQLReader.Item("sn1"))
                LineZ += 1
            End While
        End If
                'AdjustExcelFormat("裁纱 Cutting")
                'GetWipData("'0090','0100','0110','0112'")
                '' 2nd Page
                'Ws = xWorkBook.Sheets(2)
                'Ws.Activate()
                'AdjustExcelFormat("PCM裁纱 PCM Cutting")
                'GetWipData("'0111','0113'")
                '' 3rd Page
                'Ws = xWorkBook.Sheets(3)
                'Ws.Activate()
                'AdjustExcelFormat("预型 Layup")
                'GetWipData("'0120','0130','0140','0150','0160','0170'")
                '' 4th Page
                'Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                'Ws.Activate()
                'AdjustExcelFormat("PCM预型 PCM Layup")
                'GetWipData("'0151'")
                '' 5th Page
                'Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                'Ws.Activate()
                'AdjustExcelFormat("成型 Molding")
                'GetWipData("'0190','0195','0200','0210','0215','0220','0223','0225','0230','0315','0320','0330','0390'")
                ''6th Page
                'Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                'Ws.Activate()
                'AdjustExcelFormat("PCM成型 PCM Molding")
                'GetWipData("'0231','0240','0250','0260','0280','0300','0316','0321','0326','0331'")
                '' 7th Page
                'Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                'Ws.Activate()
                'AdjustExcelFormat("CNC")
                'GetWipData("'0335','0340','0350','0360','0380','0395','0495','0500','0510','0520','0530'")
                '' 8th Page
                'Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                'Ws.Activate()
                'AdjustExcelFormat("补土 Sanding")
                'GetWipData("'0325','0410','0415','0417','0430','0440','0445','0460','0465','0540','0545','0570','0575','0583','0584'")
                '' 9th Page
                'Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                'Ws.Activate()
                'AdjustExcelFormat("胶合 Gluing")
                'GetWipData("'0400','0480','0490','0492','0600','0610','0620','0623','0627'")
                '' 10th Page
                'Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                'Ws.Activate()
                'AdjustExcelFormat("涂装 Painting")
                'GetWipData("'0418','0420','0450','0470','0475','0550','0560','0563','0580','0585','0587','0590'")
                '' 11th Page
                'Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                'Ws.Activate()
                'AdjustExcelFormat("抛光 Polishing")
                'GetWipData("'0630','0635','0640','0642','0645','0657'")
                '' 12th Page
                'Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                'Ws.Activate()
                'AdjustExcelFormat("包装 Packing")
                'GetWipData("'0650','0652','0658','0659','0660','0665','0670','0675','0680','0690'")
                '' 13rd Page
                'Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                'Ws.Activate()
                'AdjustExcelFormat("成品 FinishGoods")
                'GetWipData("'0720','0730'")
    End Sub
    ' Private Sub AdjustExcelFormat(ByVal eTitle As String)
    Private Sub AdjustExcelFormat()
        Ws.Name = "MES vs ERP 盘点比对"
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "AQ1")
        oRng.EntireColumn.ColumnWidth = 25
        oRng.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Cells(1, 1) = "料号（ERP PN）"
        Ws.Cells(1, 2) = "ERP预型半成品仓"
        Ws.Cells(1, 3) = "MES预型半成品仓"
        Ws.Cells(1, 4) = "预型半成品仓差异数量"
        Ws.Cells(1, 5) = "ERP成型半成品仓"
        Ws.Cells(1, 6) = "MES成型半成品仓"
        Ws.Cells(1, 7) = "成型半成品仓差异数量"
        Ws.Cells(1, 8) = "ERPCNC半成品仓"
        Ws.Cells(1, 9) = "MES CNC半成品仓"
        Ws.Cells(1, 10) = "CNC半成品仓差异数量"
        Ws.Cells(1, 11) = "ERP胶合半成品仓"
        Ws.Cells(1, 12) = "MES胶合半成品仓"
        Ws.Cells(1, 13) = "胶合半成品仓差异数量"
        Ws.Cells(1, 14) = "ERP胶合不良品仓"
        Ws.Cells(1, 15) = "MES胶合不良品仓"
        Ws.Cells(1, 16) = "胶合不良品仓差异数量"
        Ws.Cells(1, 17) = "ERP补土半成品仓"
        Ws.Cells(1, 18) = "MES补土半成品仓"
        Ws.Cells(1, 19) = "补土半成品仓差异数量"
        Ws.Cells(1, 20) = "ERP补土不良品仓"
        Ws.Cells(1, 21) = "MES补土不良品仓"
        Ws.Cells(1, 22) = "补土不良品仓差异数量"
        Ws.Cells(1, 23) = "ERP涂装半成品仓"
        Ws.Cells(1, 24) = "MES涂装半成品仓"
        Ws.Cells(1, 25) = "涂装半成品仓差异数量"
        Ws.Cells(1, 26) = "ERP涂装不良品仓"
        Ws.Cells(1, 27) = "MES涂装不良品仓"
        Ws.Cells(1, 28) = "涂装不良品仓差异数量"
        Ws.Cells(1, 29) = "ERP抛光半成品仓"
        Ws.Cells(1, 30) = "MES抛光半成品仓"
        Ws.Cells(1, 31) = "抛光半成品仓差异数量"
        Ws.Cells(1, 32) = "ERP抛光不良品仓"
        Ws.Cells(1, 33) = "MES抛光不良品仓"
        Ws.Cells(1, 34) = "抛光不良品仓差异数量"
        Ws.Cells(1, 35) = "ERP包装半成品仓"
        Ws.Cells(1, 36) = "MES包装半成品仓"
        Ws.Cells(1, 37) = "包装半成品仓差异数量"
        Ws.Cells(1, 38) = "ERP成品仓"
        Ws.Cells(1, 39) = "MES成品仓"
        Ws.Cells(1, 40) = "成品仓差异数量"
        Ws.Cells(1, 41) = "ERP  total"
        Ws.Cells(1, 42) = "MES  total"
        Ws.Cells(1, 43) = "差异数量"
        LineZ = 2
    End Sub
    Private Sub GetWipData(ByVal eStationList As String)
        mSQLS1.CommandText = "select count(sn) as t1,m.model,c.value,m.modelname from lot l JOIN model m ON l.model=m.model JOIN sn s ON l.lot=s.lot JOIN station t ON t.station=case when (s.topreworkstation is not null and s.topreworkstation<>'') then s.topreworkstation else s.updatedstation end "
        mSQLS1.CommandText += "LEFT JOIN model_paravalue c on m.model = c.model and c.parameter = 'ERP PN' WHERE t.station  in ("
        mSQLS1.CommandText += eStationList & ") group by m.model,c.value,m.modelname order by m.model"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("model") & " " & mSQLReader.Item("modelname")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("value")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("t1")
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Inventory_Count"
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
        'g_success = False
        If g_success = True Then
            SaveExcel()
        End If
        MsgBox("DONE")
    End Sub
    Private Sub GetReady()
        oCommand.CommandText = "DELETE FROM MES_TEMP2"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
            g_success = False
            Return
        End Try
        mSQLS1.CommandText = "select cf01,warehouse,count(sn) as t1 from ( select sn,cf01,updatedstation,"
        mSQLS1.CommandText += "(case when updatedstation in ('0730', '0799') then 'D146103' "
        mSQLS1.CommandText += "      when updatedstation in ('0130', '0140', '0150', '0151', '0160', '0170','0175', '0180', '0193') then 'D353201'"
        mSQLS1.CommandText += "      when updatedstation in ('0165', '0190', '0195', '0200', '0210', '0215', '0220', '0223', '0225', '0230', '0231', '0240', '0250', '0260', '0280', '0300', '0315', '0316', '0320', '0321', '0325', '0326', '0330', '0331','0333', '0390', '0395' ) then 'D353501'"
        mSQLS1.CommandText += "      when updatedstation in ('0335', '0340', '0350', '0360', '0370', '0380', '0495', '0500', '0510', '0520', '0530') then 'D353601'"
        mSQLS1.CommandText += "      when updatedstation in ('0405', '0410', '0415', '0417', '0418', '0420', '0430', '0440', '0445', '0450', '0455', '0475') then 'D356101'"
        mSQLS1.CommandText += "      when updatedstation in ('0460', '0465', '0470', '0540', '0545', '0550', '0560', '0563', '0567', '0570', '0575', '0580', '0583', '0584', '0585', '0587', '0590', '0592', '0595') then 'D356301'"
        mSQLS1.CommandText += "      when updatedstation in ('0400', '0435', '0478', '0480', '0490', '0492', '0493', '0605', '0610', '0620', '0623', '0627') then 'D356401'"
        mSQLS1.CommandText += "      when updatedstation in ('0629', '0630', '0633', '0635', '0640', '0645') then 'D356501'"
        mSQLS1.CommandText += "      when updatedstation in ('0642', '0650', '0652', '0657', '0658', '0659', '0660', '0665', '0670', '0673', '0675', '0680', '0690') then 'D356601' end ) as warehouse "
        mSQLS1.CommandText += "from sn_temp left join lot on sn_temp.lot = lot.lot left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue.profilename = 'ERP' and sn_temp.updatedstation = model_station_paravalue.station "
        mSQLS1.CommandText += "where sn_temp.topreworkstation is null or sn_temp.topreworkstation = '' ) as AA where warehouse is NOT null group by cf01,warehouse "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                If g_success = False Then
                    Exit While
                End If
                oCommander2.CommandText = "INSERT INTO mes_temp2 (erppn,ntype,warehouse,t1) VALUES ('" & mSQLReader.Item("cf01") & "','0','" & mSQLReader.Item("warehouse") & "'," & mSQLReader.Item("t1") & ")"
                Try
                    oCommander2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    g_success = False
                    Return
                End Try
            End While
        End If
        mSQLReader.Close()
        mSQLS1.CommandText = "select cf01,warehouse,count(sn) as t1 from ( select sn,cf01,updatedstation,"
        mSQLS1.CommandText += "(case when topreworkstation in ('0410', '0440') then 'D356105' "
        mSQLS1.CommandText += "      when topreworkstation in ('0460', '0540','0570') then 'D356305'"
        mSQLS1.CommandText += "      when topreworkstation in ('0480', '0610') then 'D356405'"
        mSQLS1.CommandText += "      when topreworkstation in ('0630') then 'D356505' end ) as warehouse from sn_temp left join lot on sn_temp.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue.profilename = 'ERP' and sn_temp.topreworkstation = model_station_paravalue.station "
        mSQLS1.CommandText += "where sn_temp.topreworkstation is not null and sn_temp.topreworkstation <> '' ) as AA where warehouse is NOT null group by cf01,warehouse "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                If g_success = False Then
                    Exit While
                End If
                oCommander2.CommandText = "INSERT INTO mes_temp2 (erppn,ntype,warehouse,t1) VALUES ('" & mSQLReader.Item("cf01") & "','1','" & mSQLReader.Item("warehouse") & "'," & mSQLReader.Item("t1") & ")"
                Try
                    oCommander2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    g_success = False
                    Return
                End Try
            End While
        End If
        mSQLReader.Close()
        oCommand.CommandText = "INSERT INTO mes_temp2 select pia02,'2',pia03,pia30 from pia_file where pia01 like 'D1501-1703%' and pia19 = 'Y' and pia03 in ("
        oCommand.CommandText += "'D146103','D353201','D353501','D353601','D356101','D356301','D356401','D356501','D356601','D356105','D356305','D356405','D356505')"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
            g_success = False
            Return
        End Try
    End Sub
    Private Sub AdjustExcelFormat1()
        Ws.Name = "MES盘点站与记录站不符"
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "I1")
        oRng.EntireColumn.ColumnWidth = 25
        oRng.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Cells(1, 1) = "MES SN"
        Ws.Cells(1, 2) = "料号（ERP PN）"
        Ws.Cells(1, 3) = "盘点站"
        Ws.Cells(1, 4) = "站别名称"
        Ws.Cells(1, 5) = "最后一次扫描工站"
        Ws.Cells(1, 6) = "最后一次扫描工站名称"
        Ws.Cells(1, 7) = "最后一次QC工站"
        Ws.Cells(1, 8) = "最后一次QC工站名称"
        Ws.Cells(1, 9) = "最后一次QC工站扫描员"
        LineZ = 2
    End Sub
    Private Sub GetQCStatus(ByVal sn As String)
        mSQLS2.CommandText = "select top 1 *  from ( "
        mSQLS2.CommandText += "select timeout,tracking.station,station.stationname ,users.name from tracking "
        mSQLS2.CommandText += "left join station on tracking.station = station.station left join users on tracking.users = users.id where tracking.sn = '"
        mSQLS2.CommandText += sn & "' and tracking.station in ('0330','0331','0380','0530','0430','0475','0490','0620','0590','0592','0595','0640','0645','0659','0670') "
        mSQLS2.CommandText += "union all "
        mSQLS2.CommandText += "select timeout,tracking_dup.station,station.stationname ,users.name from tracking_dup left join station on tracking_dup.station = station.station left join users on tracking_dup.users = users.id "
        mSQLS2.CommandText += "where tracking_dup.sn = '" & sn & "' and tracking_dup.station in ('0330','0331','0380','0530','0430','0475','0490','0620','0590','0592','0595','0640','0645','0659','0670') ) AS aa order by timeout desc"
        Try
            mSQLReader2 = mSQLS2.ExecuteReader()
            If mSQLReader2.HasRows() Then
                mSQLReader2.Read()
                Ws.Cells(LineZ, 7) = mSQLReader2.Item("station")
                Ws.Cells(LineZ, 8) = mSQLReader2.Item("stationname")
                Ws.Cells(LineZ, 9) = mSQLReader2.Item("name")
            End If
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        mSQLReader2.Close()
    End Sub
End Class