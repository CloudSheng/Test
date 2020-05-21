Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form129
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim tModel_type As String
    Dim tModel As String
    Dim tLot As String
    Dim tStation As String
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    
    Private Sub Form129_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        BindStation()
        BindModel_Type()
        Dim Model_Type As String = String.Empty
        BindModel(Model_Type)
        BindLot(Model_Type)
    End Sub
    Private Sub BindModel_Type()
        Me.ComboBox1.Items.Clear()
        mSQLS1.CommandText = "SELECT * FROM model_type WHERE model_type <> 'Action'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox1.Items.Add(mSQLReader.Item(0).ToString())
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub BindStation()
        Me.ComboBox1.Items.Clear()
        mSQLS1.CommandText = "SELECT station FROM station order by station"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox4.Items.Add(mSQLReader.Item(0).ToString())
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub BindModel(ByVal Models1 As String)
        Me.ComboBox2.Items.Clear()
        mSQLS1.CommandText = "select distinct lot.model,model.modelname  from lot,model " _
                          & " where lot.model = model.model and model.model_type <> 'Action'"
        If Not String.IsNullOrEmpty(Models1) Then
            mSQLS1.CommandText += " AND model.model_type = '" & Models1 & "'"
        End If
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox2.Items.Add(mSQLReader.Item(0).ToString() & "|" & mSQLReader.Item(1).ToString())
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim model_type As String = ComboBox1.Items(ComboBox1.SelectedIndex).ToString()
        BindModel(model_type)
        BindLot(model_type)
    End Sub

    Private Sub BindLot(ByVal Models1 As String)
        Me.ComboBox3.Items.Clear()
        mSQLS1.CommandText = "select distinct lot.lot from lot,model " _
                          & " where lot.model = model.model and model.model_type <> 'Action'"
        If Not String.IsNullOrEmpty(Models1) Then
            mSQLS1.CommandText += " AND model.model_type = '" & Models1 & "'"
        End If
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox3.Items.Add(mSQLReader.Item(0).ToString())
            End While
        End If
        mSQLReader.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        ProgressBar1.Value = 0
        tModel_type = String.Empty
        tModel = String.Empty
        tLot = String.Empty
        tStation = String.Empty
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
            tModel_type = ComboBox1.SelectedItem.ToString()
        End If
        If Not IsNothing(ComboBox2.SelectedItem) Then
            tModel = ComboBox2.SelectedItem.ToString()
            Dim stCount As Int16 = Strings.InStr(tModel, "|")
            If stCount > 0 Then
                tModel = Strings.Left(tModel, stCount - 1)
            End If
        End If
        If Not IsNothing(ComboBox3.SelectedItem) Then
            tLot = ComboBox3.SelectedItem.ToString()
        End If
        If Not IsNothing(ComboBox4.SelectedItem) Then
            tStation = ComboBox4.SelectedItem.ToString()
        End If
        BackgroundWorker1.RunWorkerAsync()
        'ExportToExcel()
        'SaveExcel()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        ' 前面加 計數器
        mSQLS1.CommandText = "select count(distinct model) FROM ( Select m.model from model m "
        mSQLS1.CommandText += "left join lot l on l.model= m.model left join sn s on s.lot = l.lot  "
        mSQLS1.CommandText += "left join station t on t.station=case when (s.topreworkstation is not null and s.topreworkstation<>'') then s.topreworkstation else s.updatedstation end and t.station <> '9999' "
        mSQLS1.CommandText += "left join model_paravalue mp on m.model = mp.model and mp.parameter = 'ERP PN' "
        mSQLS1.CommandText += "where l.remark not in ('TEST','test','Test') "
        If Not String.IsNullOrEmpty(tModel_type) Then
            mSQLS1.CommandText += "and m.model_type like '" & tModel_type & "' "
        End If
        If Not String.IsNullOrEmpty(tModel) Then
            mSQLS1.CommandText += "and l.model like '" & tModel & "' "
        End If
        If Not String.IsNullOrEmpty(tLot) Then
            mSQLS1.CommandText += "and l.lot like '" & tLot & "' "
        End If
        If Not String.IsNullOrEmpty(tStation) Then
            mSQLS1.CommandText += "and t.station = '" & tStation & "' "
        End If
        mSQLS1.CommandText += " ) AS AB"

        Dim RowsT As Decimal = mSQLS1.ExecuteScalar()
        If RowsT = 0 Then
            MsgBox("没有资料，请重选条件")
            Return
        End If
        Me.ProgressBar1.Maximum = RowsT
        ' 開啟 Excel 調好格式
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()

        ' 開始算WIP
        ' 首先知道要多有多少行
        mSQLS1.CommandText = "select max(wiporder) FROM ERPSUPPORT.dbo.StationDefine where WipOrder > 0"
        Dim MaxColumn As Int16 = mSQLS1.ExecuteScalar()
        ' 然後開始寫SQL
        mSQLS1.CommandText = "select value,model,modelname"
        For i As Int16 = 1 To MaxColumn + 3 Step 1
            mSQLS1.CommandText += ",sum(t" & i & ") as t" & i
        Next
        mSQLS1.CommandText += " FROM ( select mp.value,m.model,m.modelname"
        For i As Int16 = 1 To MaxColumn Step 1
            mSQLS2.CommandText = "select station FROM ERPSUPPORT.dbo.StationDefine where WipOrder =" & i
            mSQLReader = mSQLS2.ExecuteReader()
            If mSQLReader.HasRows() Then
                mSQLS1.CommandText += ",(case when station in ("
                Dim FirstPlace As Int16 = 0
                While mSQLReader.Read()
                    If FirstPlace = 0 Then
                        mSQLS1.CommandText += "'" & mSQLReader.Item(0) & "'"
                    Else
                        mSQLS1.CommandText += ",'" & mSQLReader.Item(0) & "'"
                    End If
                    FirstPlace += 1
                End While
                mSQLS1.CommandText += ") and s.block = 'N' then 1 else 0 end) as t" & i
            End If
            mSQLReader.Close()
        Next
        mSQLS1.CommandText += ",(case when s.block = 'Y' and t.station <> '9999' then 1 else 0 end ) as t" & MaxColumn + 1
        mSQLS1.CommandText += ",(case when station in ('BLCK') and s.block = 'N' then 1 else 0 end ) as t" & MaxColumn + 2

        mSQLS2.CommandText = "SELECT station FROM ERPSUPPORT.dbo.StationDefine where WipOrder > 0"
        mSQLReader = mSQLS2.ExecuteReader()
        If mSQLReader.HasRows() Then
            Dim FirstPlace As Integer = 0
            mSQLS1.CommandText += ",(case when station not in ("
            While mSQLReader.Read()
                If FirstPlace = 0 Then
                    mSQLS1.CommandText += "'" & mSQLReader.Item(0) & "'"
                Else
                    mSQLS1.CommandText += ",'" & mSQLReader.Item(0) & "'"
                End If
                FirstPlace += 1
            End While
            mSQLS1.CommandText += ",'BLCK') and s.block = 'N' then 1 else 0 end) as t" & MaxColumn + 3
        End If
        mSQLReader.Close()
        'mSQLS1.CommandText += ",case when s.block = 'Y' then 1 else 0 end ) as t" & MaxColumn + 1
        mSQLS1.CommandText += " from model m left join lot l on l.model= m.model left join sn s on s.lot = l.lot "
        mSQLS1.CommandText += "left join station t on t.station=case when (s.topreworkstation is not null and s.topreworkstation<>'') then s.topreworkstation else s.updatedstation end and t.station <> '9999' "
        mSQLS1.CommandText += "left join model_paravalue mp on m.model = mp.model and mp.parameter = 'ERP PN' "
        mSQLS1.CommandText += "where l.remark not in ('TEST','test','Test')  "
        If Not String.IsNullOrEmpty(tModel_type) Then
            mSQLS1.CommandText += "and m.model_type like '" & tModel_type & "' "
        End If
        If Not String.IsNullOrEmpty(tModel) Then
            mSQLS1.CommandText += "and l.model like '" & tModel & "' "
        End If
        If Not String.IsNullOrEmpty(tLot) Then
            mSQLS1.CommandText += "and l.lot like '" & tLot & "' "
        End If
        If Not String.IsNullOrEmpty(tStation) Then
            mSQLS1.CommandText += "and t.station = '" & tStation & "' "
        End If
        mSQLS1.CommandText += ") AS AB GROUP BY value,model,modelname order by model "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                For i As Int16 = 0 To mSQLReader.FieldCount - 1 Step 1
                    If i < 3 Then
                        Ws.Cells(LineZ, i + 1) = mSQLReader.Item(i)
                    Else
                        If mSQLReader.Item(i) <> 0 Then
                            Ws.Cells(LineZ, i + 1) = mSQLReader.Item(i)
                        End If
                    End If
                Next
                Ws.Cells(LineZ, 70) = "=SUM(D" & LineZ & ":BQ" & LineZ & ")"
                Ws.Cells(LineZ, 71) = "=SUM(O" & LineZ & ":BQ" & LineZ & ")"
                Ws.Cells(LineZ, 72) = "=T" & LineZ & "+V" & LineZ & "+SUM(W" & LineZ & ":BN" & LineZ & ")"
                Ws.Cells(LineZ, 73) = "=Z" & LineZ & "+AA" & LineZ & "+AD" & LineZ & "+SUM(AE" & LineZ & ":BN" & LineZ & ")"
                Ws.Cells(LineZ, 74) = "=SUM(AX" & LineZ & ":BN" & LineZ & ")"
                Ws.Cells(LineZ, 75) = "=BA" & LineZ & "+SUM(BC" & LineZ & ":BN" & LineZ & ")"
                Ws.Cells(LineZ, 76) = "=SUM(H" & LineZ & ":BQ" & LineZ & ")-BZ" & LineZ
                If Not IsDBNull(mSQLReader.Item("model")) Then
                    GetMESOrder(mSQLReader.Item("model"))
                End If
                'If Not IsDBNull(mSQLReader.Item("value")) Then
                'GetERPOrder(mSQLReader.Item("value"))
                'End If
                LineZ += 1
                ProgressBar1.Value += 1
            End While
        End If
        SumColumn()
        mSQLReader.Close()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.HorizontalAlignment = xlCenter
        Ws.Columns.Font.Name = "Arial Unicode MS"
        Ws.Columns.Font.Size = 9
        Ws.Columns.ShrinkToFit = True

        oRng = Ws.Range("A1", "C1")
        oRng.EntireColumn.ColumnWidth = 40
        oRng = Ws.Range("D1", "BY1")
        oRng.EntireColumn.ColumnWidth = 15
        oRng = Ws.Range("A2", "C3")
        oRng.Interior.Color = Color.FromArgb(242, 220, 219)
        oRng = Ws.Range("AE2", "AK3")
        oRng.Interior.Color = Color.FromArgb(197, 217, 241)
        oRng = Ws.Range("AL2", "AW3")
        oRng.Interior.Color = Color.FromArgb(83, 141, 213)
        oRng = Ws.Range("BK2", "BN3")
        oRng.Interior.Color = Color.FromArgb(0, 176, 80)
        oRng = Ws.Range("BO2", "BO3")
        oRng.Interior.Color = Color.Red
        oRng = Ws.Range("BP2", "BP3")
        oRng.Interior.Color = Color.Yellow
        oRng = Ws.Range("BQ2", "BQ3")
        oRng.Interior.Color = Color.FromArgb(247, 150, 70)
        oRng = Ws.Range("BR2", "BR3")
        oRng.Interior.Color = Color.FromArgb(146, 208, 80)
        oRng = Ws.Range("BS2", "BS3")
        oRng.Interior.Color = Color.FromArgb(0, 176, 240)
        oRng = Ws.Range("BT2", "BX3")
        oRng.Interior.Color = Color.FromArgb(250, 191, 143)
        oRng = Ws.Range("BY2", "BZ3")
        oRng.Interior.Color = Color.FromArgb(146, 208, 80)

        Ws.Cells(2, 1) = "ERP PN"
        Ws.Cells(3, 1) = "ERP PN"
        Ws.Cells(2, 2) = "Product name"
        Ws.Cells(3, 2) = "Product name"
        Ws.Cells(2, 3) = "WIP_Product Description"
        Ws.Cells(3, 3) = "WIP_Product Description"
        Ws.Cells(2, 4) = "Making Label"
        Ws.Cells(3, 4) = "标签数量"
        Ws.Cells(2, 5) = "Prepreg Cut"
        Ws.Cells(3, 5) = "裁紗作业"
        Ws.Cells(2, 6) = "PREPARE FILE"
        Ws.Cells(3, 6) = "备料夹"
        Ws.Cells(2, 7) = "Layup"
        Ws.Cells(3, 7) = "预型作业"
        Ws.Cells(2, 8) = "Inspection-Layup"
        Ws.Cells(3, 8) = "预型检验"
        Ws.Cells(2, 9) = "Await Molding"
        Ws.Cells(3, 9) = "待成型"
        Ws.Cells(2, 10) = "Molding"
        Ws.Cells(3, 10) = "成型作业"
        Ws.Cells(2, 11) = "Await Layup 2"
        Ws.Cells(3, 11) = "待预型2"
        Ws.Cells(2, 12) = "Layup 2"
        Ws.Cells(3, 12) = "预型2作业"
        Ws.Cells(2, 13) = "Await Molding 2"
        Ws.Cells(3, 13) = "待成型2"
        Ws.Cells(2, 14) = "Molding 2"
        Ws.Cells(3, 14) = "成型2作业"
        Ws.Cells(2, 15) = "Molding Aftertreatment"
        Ws.Cells(3, 15) = "成型后加工"
        Ws.Cells(2, 16) = "Inspection-Molding"
        Ws.Cells(3, 16) = "成型检验"
        Ws.Cells(2, 17) = "POST CURING"
        Ws.Cells(3, 17) = "後加溫(固化)"
        Ws.Cells(2, 18) = "Await CNC"
        Ws.Cells(3, 18) = "待CNC"
        Ws.Cells(2, 19) = "CNC作业"
        Ws.Cells(3, 19) = "CNC作业"
        Ws.Cells(2, 20) = "Inspection-CNC"
        Ws.Cells(3, 20) = "CNC检验"
        Ws.Cells(2, 21) = "CNC 2"
        Ws.Cells(3, 21) = "CNC2作业"
        Ws.Cells(2, 22) = "Inspection-CNC 2"
        Ws.Cells(3, 22) = "CNC2检验"
        Ws.Cells(2, 23) = "SAND BLASTING"
        Ws.Cells(3, 23) = "喷砂作业"
        Ws.Cells(2, 24) = "Await Gluing"
        Ws.Cells(3, 24) = "待胶合"
        Ws.Cells(2, 25) = "Gluing"
        Ws.Cells(3, 25) = "胶合作业"
        Ws.Cells(2, 26) = "Inspection-Gluing"
        Ws.Cells(3, 26) = "胶合检验"
        Ws.Cells(2, 27) = "Gluing CURING"
        Ws.Cells(3, 27) = "胶合硬化"
        Ws.Cells(2, 28) = "Await Gluing 2&3"
        Ws.Cells(3, 28) = "待胶合2&3"
        Ws.Cells(2, 29) = "Gluing 2&3"
        Ws.Cells(3, 29) = "胶合2&3作业"
        Ws.Cells(2, 30) = "Inspection-Gluing 2&3"
        Ws.Cells(3, 30) = "胶合2&3检验"
        Ws.Cells(2, 31) = "Sanding 1"
        Ws.Cells(3, 31) = "补1作业"
        Ws.Cells(2, 32) = "Replenish"
        Ws.Cells(3, 32) = "点补作业"
        Ws.Cells(2, 33) = "Sanding 2"
        Ws.Cells(3, 33) = "补2作业"
        Ws.Cells(2, 34) = "Painting 1"
        Ws.Cells(3, 34) = "涂1作业"
        Ws.Cells(2, 35) = "Inspection-Painting 1"
        Ws.Cells(3, 35) = "涂1检验"
        Ws.Cells(2, 36) = "Painting 2"
        Ws.Cells(3, 36) = "涂2作业"
        Ws.Cells(2, 37) = "Inspection-Painting 2"
        Ws.Cells(3, 37) = "涂2检验"
        Ws.Cells(2, 38) = "Sanding 3"
        Ws.Cells(3, 38) = "补3作业"
        Ws.Cells(3, 39) = "Inspection-Sanding 3"
        Ws.Cells(2, 39) = "补3检验"
        Ws.Cells(2, 40) = "Decal&Inspect"
        Ws.Cells(3, 40) = "贴水标&检验"
        Ws.Cells(2, 41) = "Sanding 4"
        Ws.Cells(3, 41) = "补4作业"
        Ws.Cells(2, 42) = "Inspection-Sanding 4"
        Ws.Cells(3, 42) = "补4检验"
        Ws.Cells(2, 43) = "Sanding 5"
        Ws.Cells(3, 43) = "补5作业"
        Ws.Cells(2, 44) = "Sanding 6"
        Ws.Cells(3, 44) = "补6作业"
        Ws.Cells(3, 45) = "Painting 3"
        Ws.Cells(2, 45) = "涂3作业"
        Ws.Cells(2, 46) = "Painting 4"
        Ws.Cells(3, 46) = "涂4作业"
        Ws.Cells(2, 47) = "Inspection-top coat"
        Ws.Cells(3, 47) = "首次面漆检"
        Ws.Cells(2, 48) = "Painting 5"
        Ws.Cells(3, 48) = "涂5作业"
        Ws.Cells(2, 49) = "Painting 6"
        Ws.Cells(3, 49) = "涂6作业"
        Ws.Cells(2, 50) = "Painting final inspection"
        Ws.Cells(3, 50) = "涂装终检"
        Ws.Cells(2, 51) = "Polishing Let stand"
        Ws.Cells(3, 51) = "待抛光"
        Ws.Cells(2, 52) = "Polishing"
        Ws.Cells(3, 52) = "抛光作业"
        Ws.Cells(2, 53) = "Inspection-Polishing"
        Ws.Cells(3, 53) = "抛光检验1"
        Ws.Cells(2, 54) = "Polishing 2"
        Ws.Cells(3, 54) = "抛光2作业"
        Ws.Cells(3, 55) = "Inspection-Polishing 2"
        Ws.Cells(2, 55) = "抛光检验2"
        Ws.Cells(2, 56) = "Spring rate test"
        Ws.Cells(3, 56) = "测试作业"
        Ws.Cells(2, 57) = "Test at Lab"
        Ws.Cells(3, 57) = "测试检验"
        Ws.Cells(2, 58) = "Xray"
        Ws.Cells(3, 58) = "X光"
        Ws.Cells(2, 59) = "Xray Inspection"
        Ws.Cells(3, 59) = "X光检验"
        Ws.Cells(2, 60) = "Cleaning&Assembling&Let stand"
        Ws.Cells(3, 60) = "清洁/组装/静置"
        Ws.Cells(2, 61) = "Pack-Repair"
        Ws.Cells(3, 61) = "包装返修"
        Ws.Cells(2, 62) = "FQC"
        Ws.Cells(3, 62) = "成品检验"
        Ws.Cells(2, 63) = "Packing"
        Ws.Cells(3, 63) = "包装作业"
        Ws.Cells(2, 64) = "Complete packing"
        Ws.Cells(3, 64) = "包装完成"
        Ws.Cells(3, 65) = "B FG"
        Ws.Cells(2, 65) = "0799 B类"
        Ws.Cells(2, 66) = "FG"
        Ws.Cells(3, 66) = "成品仓存"
        Ws.Cells(2, 67) = "Block / Hold"
        Ws.Cells(3, 67) = "不良品隔离"
        Ws.Cells(2, 68) = "Block / Hold"
        Ws.Cells(3, 68) = "隔离区呆滞"
        Ws.Cells(2, 69) = "Scattered"
        Ws.Cells(3, 69) = "分类数量"
        Ws.Cells(2, 70) = "Total quantity"
        Ws.Cells(3, 70) = "含标签总量"
        Ws.Cells(2, 71) = "Number of products"
        Ws.Cells(3, 71) = "成型完成数"
        Ws.Cells(2, 72) = "Completed CNC"
        Ws.Cells(3, 72) = "CNC之后总量"
        Ws.Cells(2, 73) = "Completed Gluing"
        Ws.Cells(3, 73) = "胶合后总量"
        Ws.Cells(2, 74) = "Completed Painting"
        Ws.Cells(3, 74) = "面漆后总量"
        Ws.Cells(3, 75) = "Completed Polishing"
        Ws.Cells(2, 75) = "抛光后总量"
        Ws.Cells(2, 76) = "Layup owe number"
        Ws.Cells(3, 76) = "待预型量"
        Ws.Cells(2, 77) = "MES Order"
        Ws.Cells(3, 77) = "MES可用制令"
        Ws.Cells(2, 78) = "ERP Order"
        Ws.Cells(3, 78) = "ERP 未结订单"

        LineZ = 4
    End Sub
    Private Sub GetMESOrder(ByVal model As String)
        mSQLS2.CommandText = "select isnull(sum(qty),0) from lot where status = 'N' and lot.model = '" & model & "' "
        Dim TQ As Decimal = mSQLS2.ExecuteScalar()
        mSQLS2.CommandText = "select ISNULL(count(sn),0) from sn,lot where sn.lot =lot.lot and status = 'N' and lot.model = '" & model & "' "
        Dim RQ As Decimal = mSQLS2.ExecuteScalar()
        Dim SQ As Decimal = TQ - RQ
        Ws.Cells(LineZ, 77) = SQ
    End Sub
    Private Sub GetERPOrder(ByVal oeb04 As String)
        oCommand.CommandText = "select nvl(sum(oeb12 - oeb24 + oeb25),0) from oeb_file where oeb04 = '" & oeb04 & "' and oeb70 = 'N'"
        Dim SQ As Decimal = oCommand.ExecuteScalar()
        Ws.Cells(LineZ, 77) = SQ
    End Sub
    Private Sub SumColumn()
        Ws.Cells(1, 4) = "=SUBTOTAL(9,D4:D" & LineZ & ")"
        oRng = Ws.Range("D1", "D1")
        oRng.AutoFill(Destination:=Ws.Range("D1", "BZ1"), Type:=xlFillDefault)
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "WIP_STATUS"
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
End Class