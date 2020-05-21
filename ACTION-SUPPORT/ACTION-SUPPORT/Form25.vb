Public Class Form25
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader98 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader97 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim ptime As String = String.Empty
    Dim r_percentage As Decimal = 0
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim DS As Data.DataSet = New DataSet()
    Dim LineZ As Integer = 0
    Dim PaperDate As Date
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form25_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        PaperDate = Now.AddDays(-1)
        ptime = Today.AddDays(-1).ToString("yyyy/MM/dd")
        ptime = ptime & " 08:00:00"
        Me.DateTimePicker1.Value = Convert.ToDateTime(ptime)
        Me.DateTimePicker2.Value = Convert.ToDateTime(ptime).AddDays(1).AddSeconds(-1)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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
        mSQLS1.CommandText = "select lot.model,failstation,a.cf01,rework,b.cf01,count(sn) as t1 from failure left join lot on failure.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue as a on a.profilename = 'ERP' and a.model = lot.model and a.station = failstation "
        mSQLS1.CommandText += "left join model_station_paravalue as b on b.profilename = 'ERP' and b.model = lot.model and b.station = rework "
        mSQLS1.CommandText += "where failtime between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") '& "' and failstation in ('0670','0645','0640','0590','0475','0430','0627','0620','0490') "
        mSQLS1.CommandText += "' AND rework <> 'SCRP' and a.cf01 <> b.cf01 AND failstation not in ('0659', '0670') "
        mSQLS1.CommandText += "group by lot.model,failstation,a.cf01,rework,b.cf01"
        Dim mSQLAD As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(mSQLS1.CommandText, mConnection)
        Try
            DS.Clear()
            mSQLAD.Fill(DS, "table1")
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        Me.DataGridView1.DataSource = DS.Tables("table1")
        ModifyHeader()
        CompareAllData()
        Try
            mConnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
    End Sub
    Private Sub ModifyHeader()
        Me.DataGridView1.Columns.Item(0).HeaderText = "型号"
        Me.DataGridView1.Columns.Item(1).HeaderText = "失败工站"
        Me.DataGridView1.Columns.Item(2).HeaderText = "失败ERP料号"
        Me.DataGridView1.Columns.Item(3).HeaderText = "返工工站"
        Me.DataGridView1.Columns.Item(4).HeaderText = "返工ERP料号"
        Me.DataGridView1.Columns.Item(5).HeaderText = "数量"
    End Sub
    Private Sub CompareAllData()
        Dim CF01A As Integer = 0
        Dim CF01B As Integer = 0
        For i As Integer = 0 To Me.DataGridView1.Rows.Count - 1 Step 1
            If Not String.IsNullOrEmpty(Me.DataGridView1.Rows(i).Cells(2).Value.ToString()) Then
                CF01A += 1
            End If
            If Not String.IsNullOrEmpty(Me.DataGridView1.Rows(i).Cells(4).Value.ToString()) Then
                CF01B += 1
            End If
        Next
        If CF01A = Me.DataGridView1.Rows.Count And CF01B = Me.DataGridView1.Rows.Count Then
            Me.Button2.Enabled = True
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Me.DataGridView1.Rows.Count = 0 Then
            MsgBox("无资料可处理")
            Return
        End If
        If MsgBox("确认执行", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            Return
        End If
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommander2.Connection = oConnection
                oCommander2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

        For i As Integer = 0 To Me.DataGridView1.Rows.Count - 1 Step 1
            If Not IsDBNull(Me.DataGridView1.Rows(i).Cells(2).Value) And Not IsDBNull(Me.DataGridView1.Rows(i).Cells(4).Value) Then
                'Select Case Me.DataGridView1.Rows(i).Cells(3).Value
                'Case "0480"
                '    AutoGenerateReworkOrder(Me.DataGridView1.Rows(i).Cells(2).Value.ToString(), Me.DataGridView1.Rows(i).Cells(4).Value.ToString(),
                '                    Me.DataGridView1.Rows(i).Cells(5).Value, 4, 1, "0480")
                '    If Me.DataGridView1.Rows(i).Cells(2).Value <> Me.DataGridView1.Rows(i).Cells(4).Value Then
                '        ExtendWorkOrder(Me.DataGridView1.Rows(i).Cells(2).Value.ToString(), Me.DataGridView1.Rows(i).Cells(4).Value.ToString(),
                '                            Me.DataGridView1.Rows(i).Cells(5).Value, "0480")
                '    End If
                'Case "0460", "0540", "0583", "0570"
                ' add by cloud 20170701
                Select Case Strings.Right(Me.DataGridView1.Rows(i).Cells(4).Value.ToString(), 1)
                    Case "A", "B"
                        Select Case Strings.Right(Me.DataGridView1.Rows(i).Cells(4).Value.ToString(), 3)
                            Case "63A"
                                r_percentage = 0.4
                            Case "65A"
                                r_percentage = 0.3
                            Case Else
                                r_percentage = 1
                        End Select
                    Case Else
                        Select Case Strings.Right(Me.DataGridView1.Rows(i).Cells(4).Value.ToString(), 2)
                            Case "63"
                                r_percentage = 0.4
                            Case "65"
                                r_percentage = 0.3
                            Case Else
                                r_percentage = 1
                        End Select
                End Select
                AutoGenerateReworkOrder(Me.DataGridView1.Rows(i).Cells(2).Value.ToString(), Me.DataGridView1.Rows(i).Cells(4).Value.ToString(),
                                Me.DataGridView1.Rows(i).Cells(5).Value, 1, r_percentage, Me.DataGridView1.Rows(i).Cells(3).Value)
                'If Me.DataGridView1.Rows(i).Cells(2).Value <> Me.DataGridView1.Rows(i).Cells(4).Value Then
                ExtendWorkOrder(Me.DataGridView1.Rows(i).Cells(2).Value.ToString(), Me.DataGridView1.Rows(i).Cells(4).Value.ToString(),
                                    Me.DataGridView1.Rows(i).Cells(5).Value, Me.DataGridView1.Rows(i).Cells(3).Value)
                'End If
                '    Case "0630"
                'AutoGenerateReworkOrder(Me.DataGridView1.Rows(i).Cells(2).Value.ToString(), Me.DataGridView1.Rows(i).Cells(4).Value.ToString(),
                '                Me.DataGridView1.Rows(i).Cells(5).Value, 1, 1, "0630")
                'End Select
            End If
        Next
        ' 當張返工工單
        'AutoGenerateReworkOrder("101AQ0305011065", "101AQ0305011064", 1, 1, 1, "0640")

        ' 後續返工工單
        'ExtendWorkOrder("101AQ0305011065", "101AQ0305011064", 1, "0640")
        MsgBox("DONE")
    End Sub
    Public Sub AutoGenerateReworkOrder(ByVal erp1 As String, ByVal erp2 As String, quantity As Decimal, ByVal typeA As Int16, ByVal percentage As Decimal, ByVal station1 As String)
        Dim sfb01 As String = String.Empty
        Dim sfb82 As String = "D35"
        sfb01 = Me.Getsfb01(typeA)
        '20160113
        If String.IsNullOrEmpty(Me.TextBox1.Text) Then
            Me.TextBox1.Text = sfb01
            Me.TextBox2.Text = sfb01
        End If
        Me.TextBox2.Text = sfb01
        If Strings.Right(erp2, 1) = "A" Then
            sfb82 = sfb82 & Strings.Right(erp2, 3)
            sfb82 = sfb82.Remove(sfb82.Count() - 1)
        Else
            sfb82 = sfb82 & Strings.Right(erp2, 2)
        End If
        'sfb82 = sfb82 & Strings.Right(erp2, 2)
        If typeA = 1 Then
            oCommand.CommandText = "INSERT INTO sfb_file VALUES ('" & sfb01 & "',5,NULL,2,'" & erp2 & "',NULL,NULL,to_date('" & PaperDate.ToString("yyyy/MM/dd") & "','yyyy/MM/dd')"
            oCommand.CommandText += "," & quantity & ",0,0,0,0,0,0,0,NULL,to_date('" & PaperDate.ToString("yyyy/MM/dd") & "','yyyy/MM/dd'),'00:00',to_date('" & PaperDate.ToString("yyyy/MM/dd") & "','yyyy/MM/dd')"
            oCommand.CommandText += ",'00:00',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'Y','N',NULL,to_date('" & PaperDate.ToString("yyyy/MM/dd") & "','yyyy/MM/dd')"
            oCommand.CommandText += ",NULL,NULL,NULL,NULL,'Y',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,NULL,'N',NULL,to_date('" & PaperDate.ToString("yyyy/MM/dd") & "','yyyy/MM/dd')"
            oCommand.CommandText += ",'" & sfb82 & "',NULL,NULL,'Y',NULL,NULL,NULL,'N','N',NULL,NULL,NULL,NULL,'Y',1,NULL,'Y','DA99018','D1461',NULL,NULL,NULL,'N',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" & station1 & "'"
            oCommand.CommandText += ",0,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,'DA99018','N','ACTIONTEST','ACTIONTEST','DA99018','D1461','N',NULL)"
        Else
            oCommand.CommandText = "INSERT INTO sfb_file VALUES ('" & sfb01 & "',1,NULL,2,'" & erp2 & "',NULL,NULL,to_date('" & PaperDate.ToString("yyyy/MM/dd") & "','yyyy/MM/dd')"
            oCommand.CommandText += "," & quantity & ",0,0,0,0,0,0,0,NULL,to_date('" & PaperDate.ToString("yyyy/MM/dd") & "','yyyy/MM/dd'),'00:00',to_date('" & PaperDate.ToString("yyyy/MM/dd") & "','yyyy/MM/dd')"
            oCommand.CommandText += ",'00:00',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'Y','N',NULL,to_date('" & PaperDate.ToString("yyyy/MM/dd") & "','yyyy/MM/dd')"
            oCommand.CommandText += ",NULL,NULL,NULL,NULL,'Y',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,NULL,'N',NULL,to_date('" & PaperDate.ToString("yyyy/MM/dd") & "','yyyy/MM/dd')"
            oCommand.CommandText += ",'" & sfb82 & "',NULL,NULL,'Y',NULL,NULL,NULL,'N','N',NULL,NULL,NULL,NULL,'N',1,NULL,'Y','DA99018','D1461',NULL,NULL,NULL,'N',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" & station1 & "'"
            oCommand.CommandText += ",0,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,'DA99018','N','ACTIONTEST','ACTIONTEST','DA99018','D1461','N',NULL)"
        End If
        Try
            Dim ED As Int16 = oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        AutoGenerateReworkOrderDetail(erp1, erp2, quantity, sfb01, typeA, percentage)
    End Sub
    Public Function Getsfb01(ByVal typeA As String)
        Dim AB As String = String.Empty
        If typeA = 1 Then
            AB = "D5111-" & PaperDate.ToString("yy") & PaperDate.ToString("MM")
        Else
            AB = "D5114-" & PaperDate.ToString("yy") & PaperDate.ToString("MM")
        End If
        oCommander2.CommandText = "select nvl(MAX(SUBSTR(SFB01,11,4)),0) from sfb_file where sfb01 LIKE '" & AB & "%'"
        Dim MaxInt As Integer = oCommander2.ExecuteScalar()
        MaxInt += 1
        Select Case Strings.Len(MaxInt.ToString())
            Case 1
                AB = AB & "000" & MaxInt
            Case 2
                AB = AB & "00" & MaxInt
            Case 3
                AB = AB & "0" & MaxInt
            Case 4
                AB = AB & MaxInt
        End Select
        Return AB
    End Function
    Public Sub AutoGenerateReworkOrderDetail(ByVal erp1 As String, ByVal erp2 As String, ByVal quantity As Decimal, ByVal s1 As String, ByVal typeA As Int16, ByVal percentage As Decimal)
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        Dim oCommander98 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander98.Connection = oConnection
        oCommander98.CommandType = CommandType.Text
        Select Case typeA
            Case 1
                ' 第一張返工工單單身
                oCommander99.CommandText = "INSERT INTO sfa_file VALUES ('" & s1 & "',5,'" & erp1 & "'," & quantity & "," & quantity & ",0,0,0,0,0,0,0,NULL,' ',0,NULL,'N','PCS',1,'PCS',1,1,1,1,0,'"
                oCommander99.CommandText += erp1 & "',1,'" & erp2 & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,0,'Y',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'ACTIONTEST','ACTIONTEST',' ',0)"
                Try
                    oCommander99.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
                oCommander98.CommandText = "select round(sum(bmb06/bmb07),3) as t1,round(sum(bmb06/bmb07 * (1+ bmb08 /100)),3) as t2,bmb01,bmb03,ima70,bmb10,ima86,ima64,ima86_fac,bmb16,ima64 from bmb_file full join ima_file on bmb03 = ima01 where bmb01 = '" & erp2 & "' and bmb05 is NULL and bmb19 = 1 group by bmb01,bmb03,ima70,bmb10,ima86,ima64,ima86_fac,bmb16,ima64 order by bmb03"
                oReader98 = oCommander98.ExecuteReader()
                If oReader98.HasRows() Then
                    While oReader98.Read()
                        If erp2 = "139AH0505011063" And erp1 = "139AH0505011065" And oReader98.Item("bmb03") = "304000020007" Then
                            Continue While
                        End If
                        If erp2 = "139AH0506011063" And erp1 = "139AH0506011065" And oReader98.Item("bmb03") = "304000020008" Then
                            Continue While
                        End If
                        If erp2 = "139AH0507011063" And erp1 = "139AH0507011065" And oReader98.Item("bmb03") = "304000020009" Then
                            Continue While
                        End If
                        If erp2 = "138AH0513011063" And erp1 = "138AH0513011065" And oReader98.Item("bmb03") = "304000020010" Then
                            Continue While
                        End If
                        If erp2 = "138AH0514011063" And erp1 = "138AH0514011065" And oReader98.Item("bmb03") = "304000020010" Then
                            Continue While
                        End If
                        If erp2 = "121AS0108013063" And erp1 = "121AS0108013065" And oReader98.Item("bmb03") = "303000020021" Then
                            Continue While
                        End If
                        If erp2 = "121AS0108013063" And erp1 = "121AS0108013063A" And oReader98.Item("bmb03") = "303000020021" Then
                            Continue While
                        End If
                        If erp2 = "121AS0109013063" And erp1 = "121AS0109013065" And oReader98.Item("bmb03") = "303000020021" Then
                            Continue While
                        End If
                        If erp2 = "121AS0109013063" And erp1 = "121AS0109013063A" And oReader98.Item("bmb03") = "303000020021" Then
                            Continue While
                        End If
                        If erp2 = "101AF0202011063" And erp1 = "101AF0202011065" And oReader98.Item("bmb03") = "303000020001" Then
                            Continue While
                        End If
                        If erp2 = "118AF0101001063" And erp1 = "118AF0101001065" And oReader98.Item("bmb03") = "303000020001" Then
                            Continue While
                        End If
                        Dim Usage As Decimal = oReader98.Item("t2") * quantity
                        Dim UnitUsage As Decimal = oReader98.Item("t1")
                        Dim UnitUsageR As Decimal = oReader98.Item("t2")
                        Dim sfa11 As String = String.Empty
                        If oReader98.Item("ima70") = "Y" Then
                            sfa11 = "E"
                        Else
                            sfa11 = "N"
                        End If
                        ' 20160113
                        If sfa11 = "E" Then
                            If oReader98.Item("bmb03") = "508000020009" Then
                                Usage = Usage * 0.8
                                Usage = Decimal.Round(Usage, 3)
                                UnitUsage = Usage / quantity
                                UnitUsageR = UnitUsage
                            Else
                                Usage = Usage * percentage
                                Usage = Decimal.Round(Usage, 3)
                                UnitUsage = Usage / quantity
                                UnitUsageR = UnitUsage
                            End If
                        End If
                        Dim sfa13 As Decimal = 1
                        If oReader98.Item("bmb10").ToString() <> oReader98.Item("ima86").ToString() Then
                            sfa13 = Gsfa13(oReader98.Item("bmb10").ToString(), oReader98.Item("ima86").ToString(), oReader98.Item("bmb03").ToString())
                        End If
                        If sfa11 = "N" And oReader98.Item("ima64") = 1 Then
                            Usage = Decimal.Ceiling(Usage)
                            UnitUsageR = Usage / quantity
                        End If
                        oCommander99.CommandText = "INSERT INTO sfa_file VALUES ('" & s1 & "',5,'" & oReader98.Item("bmb03") & "'," & Usage & "," & Usage & ",0,0,0,0,0,0,0,NULL,' ',0,NULL,'" & sfa11 & "','" & oReader98.Item("bmb10") & "'," & sfa13 & ",'" & oReader98.Item("ima86") & "'," & oReader98.Item("ima86_fac") & "," & UnitUsage & "," & UnitUsageR & ",0," & oReader98.Item("bmb16") & ",'"
                        oCommander99.CommandText += oReader98.Item("bmb03") & "',1,'" & erp2 & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,0,'Y',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'ACTIONTEST','ACTIONTEST',' ',0)"
                        Try
                            oCommander99.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.Message())
                        End Try
                    End While
                End If
                oReader98.Close()
            Case 0  '後續返工工單 - 非 66
                '後續返工工單
                oCommander98.CommandText = "select round(sum(bmb06/bmb07),3) as t1,round(sum(bmb06/bmb07 * (1+ bmb08 /100)),3) as t2,bmb01,bmb03,ima70,bmb10,ima86,ima64,ima86_fac,bmb16,ima64 from bmb_file full join ima_file on bmb03 = ima01 where bmb01 = '" & erp2 & "' and bmb05 is NULL group by bmb01,bmb03,ima70,bmb10,ima86,ima64,ima86_fac,bmb16,ima64 order by bmb03"
                oReader98 = oCommander98.ExecuteReader()
                If oReader98.HasRows() Then
                    While oReader98.Read()
                        If erp2 = "139AH0505011063" And erp1 = "139AH0505011065" And oReader98.Item("bmb03") = "304000020007" Then
                            Continue While
                        End If
                        If erp2 = "139AH0506011063" And erp1 = "139AH0506011065" And oReader98.Item("bmb03") = "304000020008" Then
                            Continue While
                        End If
                        If erp2 = "139AH0507011063" And erp1 = "139AH0507011065" And oReader98.Item("bmb03") = "304000020009" Then
                            Continue While
                        End If
                        If erp2 = "138AH0513011063" And erp1 = "138AH0513011065" And oReader98.Item("bmb03") = "304000020010" Then
                            Continue While
                        End If
                        If erp2 = "138AH0514011063" And erp1 = "138AH0514011065" And oReader98.Item("bmb03") = "304000020010" Then
                            Continue While
                        End If
                        If erp2 = "121AS0108013063" And erp1 = "121AS0108013065" And oReader98.Item("bmb03") = "303000020021" Then
                            Continue While
                        End If
                        If erp2 = "121AS0108013063" And erp1 = "121AS0108013063A" And oReader98.Item("bmb03") = "303000020021" Then
                            Continue While
                        End If
                        If erp2 = "121AS0109013063" And erp1 = "121AS0109013065" And oReader98.Item("bmb03") = "303000020021" Then
                            Continue While
                        End If
                        If erp2 = "121AS0109013063" And erp1 = "121AS0109013063A" And oReader98.Item("bmb03") = "303000020021" Then
                            Continue While
                        End If
                        Dim Usage As Decimal = oReader98.Item("t2") * quantity
                        Dim UnitUsage As Decimal = oReader98.Item("t1")
                        Dim UnitUsageR As Decimal = oReader98.Item("t2")
                        Dim sfa11 As String = String.Empty
                        If oReader98.Item("ima70") = "Y" Then
                            sfa11 = "E"
                        Else
                            sfa11 = "N"
                        End If
                        If sfa11 = "E" Then
                            If oReader98.Item("bmb03") = "508000020009" Then
                                Usage = Usage * 0.8
                                Usage = Decimal.Round(Usage, 3)
                                UnitUsage = Usage / quantity
                                UnitUsageR = UnitUsage
                            Else
                                Usage = Usage * percentage
                                Usage = Decimal.Round(Usage, 3)
                                UnitUsage = Usage / quantity
                                UnitUsageR = UnitUsage
                            End If
                        End If
                        Dim sfa13 As Decimal = 1
                        If oReader98.Item("bmb10").ToString() <> oReader98.Item("ima86").ToString() Then
                            sfa13 = Gsfa13(oReader98.Item("bmb10").ToString(), oReader98.Item("ima86").ToString(), oReader98.Item("bmb03").ToString())
                        End If
                        If sfa11 = "N" And oReader98.Item("ima64") = 1 Then
                            Usage = Decimal.Ceiling(Usage)
                            UnitUsageR = Usage / quantity
                        End If
                        oCommander99.CommandText = "INSERT INTO sfa_file VALUES ('" & s1 & "',1,'" & oReader98.Item("bmb03") & "'," & Usage & "," & Usage & ",0,0,0,0,0,0,0,NULL,' ',0,NULL,'" & sfa11 & "','" & oReader98.Item("bmb10") & "'," & sfa13 & ",'" & oReader98.Item("ima86") & "'," & oReader98.Item("ima86_fac") & "," & UnitUsage & "," & UnitUsageR & ",0," & oReader98.Item("bmb16") & ",'"
                        oCommander99.CommandText += oReader98.Item("bmb03") & "',1,'" & erp2 & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,0,'Y',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'ACTIONTEST','ACTIONTEST',' ',0)"
                        Try
                            oCommander99.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.Message())
                        End Try
                    End While
                End If
                oReader98.Close()
                'Case 2  '包裝工單特別
                '    oCommander98.CommandText = "select * from bmb_file full join ima_file on bmb03 = ima01 where bmb01 = '" & erp2 & "' and bmb05 is NULL and (ima08 = 'M' or ima02 like '%胶带%') order by bmb03"
                '    oReader98 = oCommander98.ExecuteReader()
                '    If oReader98.HasRows() Then
                '        While oReader98.Read()
                '            Dim Usage As Decimal = oReader98.Item("bmb06") / oReader98.Item("bmb07") * (1 + oReader98.Item("bmb08") / 100) * quantity
                '            Dim UnitUsage As Decimal = oReader98.Item("bmb06") / oReader98.Item("bmb07") * (1 + oReader98.Item("bmb08") / 100)
                '            Dim UnitUsageR As Decimal = UnitUsage
                '            Dim sfa11 As String = String.Empty
                '            If oReader98.Item("ima70") = "Y" Then
                '                sfa11 = "E"
                '            Else
                '                sfa11 = "N"
                '            End If
                '            Dim sfa13 As Decimal = 1
                '            If oReader98.Item("bmb10").ToString() <> oReader98.Item("ima86").ToString() Then
                '                sfa13 = Gsfa13(oReader98.Item("bmb10").ToString(), oReader98.Item("ima86").ToString(), oReader98.Item("bmb03").ToString())
                '            End If
                '            If sfa11 = "N" And oReader98.Item("ima64") = 1 Then
                '                Usage = Decimal.Ceiling(Usage)
                '                UnitUsageR = Usage / quantity
                '            End If
                '            oCommander99.CommandText = "INSERT INTO sfa_file VALUES ('" & s1 & "',5,'" & oReader98.Item("bmb03") & "'," & Usage & "," & Usage & ",0,0,0,0,0,0,0,NULL,' ',0,NULL,'" & sfa11 & "','" & oReader98.Item("bmb10") & "'," & sfa13 & ",'" & oReader98.Item("ima86") & "'," & oReader98.Item("ima86_fac") & "," & UnitUsage & "," & UnitUsageR & ",0," & oReader98.Item("bmb16") & ",'"
                '            oCommander99.CommandText += oReader98.Item("bmb03") & "',1,'" & erp2 & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,0,'Y',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'ACTIONTEST','ACTIONTEST',' ',0)"
                '            Try
                '                oCommander99.ExecuteNonQuery()
                '            Catch ex As Exception
                '                MsgBox(ex.Message())
                '            End Try
                '        End While
                '    End If
                '    oReader98.Close()
                'Case 3   ' 0480 後續專用
                '    oCommander98.CommandText = "select * from bmb_file full join ima_file on bmb03 = ima01 where bmb01 = '" & erp2 & "' and bmb19 = 2 and bmb05 is NULL order by bmb03"
                '    oReader98 = oCommander98.ExecuteReader()
                '    If oReader98.HasRows() Then
                '        While oReader98.Read()
                '            Dim Usage As Decimal = quantity
                '            Dim UnitUsage As Decimal = 1
                '            Dim UnitUsageR As Decimal = 1
                '            Dim sfa11 As String = String.Empty
                '            'If oReader98.Item("ima70") = "Y" Then
                '            'sfa11 = "E"
                '            'Else
                '            sfa11 = "N"
                '            'End If
                '            Dim sfa13 As Decimal = 1
                '            'If oReader98.Item("bmb10").ToString() <> oReader98.Item("ima86").ToString() Then
                '            'sfa13 = Gsfa13(oReader98.Item("bmb10").ToString(), oReader98.Item("ima86").ToString(), oReader98.Item("bmb03").ToString())
                '            'End If
                '            'If sfa11 = "N" And oReader98.Item("ima64") = 1 Then
                '            'Usage = Decimal.Ceiling(Usage)
                '            'UnitUsageR = Usage / quantity
                '            'End If
                '            oCommander99.CommandText = "INSERT INTO sfa_file VALUES ('" & s1 & "',5,'" & oReader98.Item("bmb03") & "'," & Usage & "," & Usage & ",0,0,0,0,0,0,0,NULL,' ',0,NULL,'" & sfa11 & "','" & oReader98.Item("bmb10") & "'," & sfa13 & ",'" & oReader98.Item("ima86") & "'," & oReader98.Item("ima86_fac") & "," & UnitUsage & "," & UnitUsageR & ",0," & oReader98.Item("bmb16") & ",'"
                '            oCommander99.CommandText += oReader98.Item("bmb03") & "',1,'" & erp2 & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,0,'Y',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'ACTIONTEST','ACTIONTEST',' ',0)"
                '            Try
                '                oCommander99.ExecuteNonQuery()
                '            Catch ex As Exception
                '                MsgBox(ex.Message())
                '            End Try
                '        End While
                '    End If
                'Case 4  '0480 當站專用
                '    ' 第一張返工工單單身
                '    oCommander99.CommandText = "INSERT INTO sfa_file VALUES ('" & s1 & "',5,'" & erp1 & "'," & quantity & "," & quantity & ",0,0,0,0,0,0,0,NULL,' ',0,NULL,'N','PCS',1,'PCS',1,1,1,1,0,'"
                '    oCommander99.CommandText += erp1 & "',1,'" & erp2 & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,0,'Y',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'ACTIONTEST','ACTIONTEST',' ',0)"
                '    Try
                '        oCommander99.ExecuteNonQuery()
                '    Catch ex As Exception
                '        MsgBox(ex.Message())
                '    End Try
                '    oCommander98.CommandText = "select sum(bmb06/bmb07) as t1,sum(bmb06/bmb07 * (1+ bmb08 /100)) as t2,bmb01,bmb03,ima70,bmb10,ima86,ima64,ima86_fac,bmb16,ima64 from bmb_file full join ima_file on bmb03 = ima01 where bmb01 = '" & erp2 & "' and bmb05 is NULL and bmb03 not in ('513000020001') and bmb19 = 1 group by bmb01,bmb03,ima70,bmb10,ima86,ima64,ima86_fac,bmb16,ima64 order by bmb03"
                '    oReader98 = oCommander98.ExecuteReader()
                '    If oReader98.HasRows() Then
                '        While oReader98.Read()
                '            Dim Usage As Decimal = oReader98.Item("t2") * quantity
                '            Dim UnitUsage As Decimal = oReader98.Item("t1")
                '            Dim UnitUsageR As Decimal = oReader98.Item("t2")
                '            Dim sfa11 As String = String.Empty
                '            If oReader98.Item("ima70") = "Y" Then
                '                sfa11 = "E"
                '            Else
                '                sfa11 = "N"
                '            End If
                '            ' 20160113
                '            If sfa11 = "E" Then
                '                Usage = Usage * percentage
                '                Usage = Decimal.Round(Usage, 3)
                '                UnitUsage = Usage / quantity
                '                UnitUsageR = UnitUsage
                '            End If
                '            Dim sfa13 As Decimal = 1
                '            If oReader98.Item("bmb10").ToString() <> oReader98.Item("ima86").ToString() Then
                '                sfa13 = Gsfa13(oReader98.Item("bmb10").ToString(), oReader98.Item("ima86").ToString(), oReader98.Item("bmb03").ToString())
                '            End If
                '            If sfa11 = "N" And oReader98.Item("ima64") = 1 Then
                '                Usage = Decimal.Ceiling(Usage)
                '                UnitUsageR = Usage / quantity
                '            End If
                '            oCommander99.CommandText = "INSERT INTO sfa_file VALUES ('" & s1 & "',5,'" & oReader98.Item("bmb03") & "'," & Usage & "," & Usage & ",0,0,0,0,0,0,0,NULL,' ',0,NULL,'" & sfa11 & "','" & oReader98.Item("bmb10") & "'," & sfa13 & ",'" & oReader98.Item("ima86") & "'," & oReader98.Item("ima86_fac") & "," & UnitUsage & "," & UnitUsageR & ",0," & oReader98.Item("bmb16") & ",'"
                '            oCommander99.CommandText += oReader98.Item("bmb03") & "',1,'" & erp2 & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,0,'Y',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'ACTIONTEST','ACTIONTEST',' ',0)"
                '            Try
                '                oCommander99.ExecuteNonQuery()
                '            Catch ex As Exception
                '                MsgBox(ex.Message())
                '            End Try
                '        End While
                '    End If
                '    oReader98.Close()
        End Select
        'If typeA = 1 Then

        '        Else

        'End If
    End Sub
    Private Sub ExtendWorkOrder(ByVal erp1 As String, ByVal erp2 As String, ByVal quantity As Decimal, ByVal station As String)
        If Strings.Right(erp2, 2) = "64" Or Strings.Right(erp2, 3) = "64A" Or Strings.Right(erp2, 3) = "64B" Then
            Return
        End If
        Dim l_percentage As Decimal = 1
        Dim oCommander97 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander97.Connection = oConnection
        oCommander97.CommandType = CommandType.Text
        oCommander97.CommandText = "select bmb01,bmb03 from bmb_file where bmb01 = '" & erp1 & "' and bmb05 is NULL and bmb19 = 2 order by bmb03"
        oReader97 = oCommander97.ExecuteReader()
        If oReader97.HasRows() Then
            While oReader97.Read()
                If Strings.Right(oReader97.Item("bmb03").ToString(), 2) = "64" Or
                    Strings.Right(oReader97.Item("bmb03").ToString(), 3) = "64A" Or
                    Strings.Right(oReader97.Item("bmb03").ToString(), 3) = "64B" Then
                    Continue While
                End If
                If oReader97.Item("bmb03") <> erp2 Then
                    'If oReader97.Item("bmb01").ToString.EndsWith("66") Then
                    'AutoGenerateReworkOrder(oReader97.Item("bmb03"), oReader97.Item("bmb01"), quantity, 2)
                    'Else
                    'If station = "0480" And Strings.Right(oReader97.Item("bmb01").ToString(), 2) <> "66" Then
                    'AutoGenerateReworkOrder(oReader97.Item("bmb03"), oReader97.Item("bmb01"), quantity, 3, 1, station)
                    'End If
                    'If (station = "0460" Or station = "0540" Or station = "0583" Or station = "0570") And Strings.Right(oReader97.Item("bmb01").ToString(), 2) <> "66" Then
                    AutoGenerateReworkOrder(oReader97.Item("bmb03"), oReader97.Item("bmb01"), quantity, 0, 1, station)
                    'End If
                    'End If
                    Call ExtendWorkOrder(oReader97.Item("bmb03"), erp2, quantity, station)
                Else
                    'If station = "0480" And Strings.Right(oReader97.Item("bmb01").ToString(), 2) <> "66" Then
                    '    AutoGenerateReworkOrder(oReader97.Item("bmb03"), oReader97.Item("bmb01"), quantity, 3, 1, station)
                    'End If
                    'If (station = "0460" Or station = "0540" Or station = "0583" Or station = "0570") And Strings.Right(oReader97.Item("bmb01").ToString(), 2) <> "66" Then
                    Select Case Strings.Right(oReader97.Item("bmb01").ToString(), 1)
                        Case "A", "B"
                            Select Case Strings.Right(oReader97.Item("bmb01").ToString(), 3)
                                Case "63A"
                                    l_percentage = 0.4
                                Case "65A"
                                    l_percentage = 0.3
                                Case Else
                                    l_percentage = 1
                            End Select
                        Case Else
                            Select Case Strings.Right(oReader97.Item("bmb01").ToString(), 2)
                                Case "63"
                                    l_percentage = 0.4
                                Case "65"
                                    l_percentage = 0.3
                                Case Else
                                    l_percentage = 1
                            End Select
                    End Select
                    AutoGenerateReworkOrder(oReader97.Item("bmb03"), oReader97.Item("bmb01"), quantity, 0, l_percentage, station)
                    'End If
                    'AutoGenerateReworkOrder(oReader97.Item("bmb03"), oReader97.Item("bmb01"), quantity, 0)
                End If

            End While
        End If
        'oReader97.Close()

    End Sub
    Private Function Gsfa13(ByVal v1 As String, ByVal v2 As String, ByVal erppn As String)
        oCommander2.CommandText = "select nvl((smd04/smd06),0) from smd_file where smd01 = '" & erppn & "' and smd03 = '" & v1 & "' and smd02 = '" & v2 & "'"
        Dim sfa13 As Decimal = 1
        Try
            sfa13 = oCommander2.ExecuteScalar()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        If IsDBNull(sfa13) Then
            sfa13 = 1
        End If
        Return sfa13
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Me.DataGridView1.Rows.Count = 0 Then
            MsgBox("无资料，不可汇出")
            Return
        End If
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        'Label1.Text = "已完成"
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Fail_Summary_Report"
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
        If oConnection.State = ConnectionState.Open Then
            Try
                oConnection.Close()
                Module1.KillExcelProcess(OldExcel)
                MsgBox("Finished")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        AdjustExcelFormat()
        For i As Integer = 0 To Me.DataGridView1.Rows.Count - 1 Step 1
            For j As Integer = 0 To Me.DataGridView1.Columns.Count - 1 Step 1
                Ws.Cells(LineZ, j + 1) = Me.DataGridView1.Rows(i).Cells(j).Value
            Next
            LineZ += 1
        Next
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "Fail Summary"
        'Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 20
        Ws.Cells(1, 1) = "型号"
        Ws.Cells(1, 2) = "失败工站"
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.NumberFormat = "@"
        Ws.Cells(1, 3) = "失败ERP料号"
        Ws.Cells(1, 4) = "返工工站"
        oRng = Ws.Range("D1", "D1")
        oRng.EntireColumn.NumberFormat = "@"
        Ws.Cells(1, 5) = "返工ERP料号"
        Ws.Cells(1, 6) = "数量"
        'Ws.Cells(1, 7) = "预计回厂时间"
        LineZ = 2
    End Sub
End Class