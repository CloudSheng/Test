Imports Microsoft.VisualBasic.Strings
Public Class Form69
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim DS As Data.DataSet = New DataSet()
    Dim modifyDate As Date


    Private Sub TextBox1_Leave(sender As Object, e As EventArgs) Handles TextBox1.Leave
        If Not String.IsNullOrEmpty(TextBox1.Text) Then
            oCommander2.CommandText = "SELECT COUNT(*) FROM gen_file where gen01 = '" & TextBox1.Text & "' AND genacti = 'Y'"
            Dim l_t1 As Int16 = oCommander2.ExecuteScalar()
            If l_t1 <= 0 Then
                MsgBox("请购账号有误")
                TextBox1.Focus()
                Return
            End If
            ' 通過驗證, 帶出預設部門
            oCommander2.CommandText = "SELECT GEN03 FROM gen_file where gen01 = '" & TextBox1.Text & "' AND genacti = 'Y'"
            Dim l_t2 As String = oCommander2.ExecuteScalar()
            TextBox2.Text = l_t2
        Else
            TextBox2.Text = ""
        End If
    End Sub

    Private Sub Form69_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
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
    End Sub

    Private Sub DateTimePicker1_Leave(sender As Object, e As EventArgs) Handles DateTimePicker1.Leave
        oCommander2.CommandText = "SELECT SMA53 FROM SMA_FILE"
        Dim l_d1 As Date = oCommander2.ExecuteScalar()
        If l_d1 >= DateTimePicker1.Value Then
            MsgBox("异动日期不可小于关账日期")
            DateTimePicker1.Focus()
            Return
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString = "SELECT * FROM [sheet1$]"
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)
            'Dim DS As Data.DataSet = New DataSet()
            Try
                ExcelAdapater.Fill(DS, "table1")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            Me.DataGridView1.DataSource = DS.Tables("table1")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        modifyDate = DateTimePicker1.Value.ToString("yyyy/MM/dd")
        Label6.Text = "处理中"
        '先檢查各資料
        If String.IsNullOrEmpty(TextBox1.Text) Or String.IsNullOrEmpty(TextBox2.Text) Then
            MsgBox("请输入请购账号")
            TextBox1.Focus()
            Label6.Text = "失败"
            Return
        End If
        If IsNothing(ComboBox1.SelectedItem) Then
            MsgBox("请选择请购单性质")
            Label6.Text = "失败"
            Return
        End If
        Dim Pmktype As String = Strings.Left(ComboBox1.SelectedItem.ToString(), 3)
        'If DS.Tables("table1").Rows.Count = 0 Then
        If IsNothing(DS.Tables("table1")) Then
            MsgBox("无单身资料，请检查")
            Label6.Text = "失败"
            Return
        End If
        ' 單身資料
        For i As Integer = 0 To DS.Tables("table1").Rows.Count - 1 Step 1
            If String.IsNullOrEmpty(DS.Tables("table1").Rows(i).Item("料件编号").ToString) Or _
                String.IsNullOrEmpty(DS.Tables("table1").Rows(i).Item("请购单位").ToString) Or _
                String.IsNullOrEmpty(DS.Tables("table1").Rows(i).Item("请购数量").ToString) Or _
                String.IsNullOrEmpty(DS.Tables("table1").Rows(i).Item("交货日期").ToString) Or _
                String.IsNullOrEmpty(DS.Tables("table1").Rows(i).Item("项目编号").ToString) Then
                MsgBox("单身资料不完整，请检查")
                Label6.Text = "失败"
                Return
            End If
            If DS.Tables("table1").Rows(i).Item("请购数量") <= 0 Then
                MsgBox(DS.Tables("table1").Rows(i).Item("料件编号").ToString.Trim() & "数量有误，请检查")
                Label6.Text = "失败"
                Return
            End If
            oCommander2.CommandText = "SELECT COUNT(*) FROM ima_file WHERE ima01 = '" & DS.Tables("table1").Rows(i).Item("料件编号").ToString.Trim() & "' AND imaacti = 'Y' and ima08 = 'P'"
            Dim l_v1 As Int16 = oCommander2.ExecuteScalar()
            If l_v1 <= 0 Then
                MsgBox("单身料号" & DS.Tables("table1").Rows(i).Item("料件编号").ToString.Trim() & "有误，请检查")
                Label6.Text = "失败"
                Return
            End If
            If DS.Tables("table1").Rows(i).Item("交货日期") < modifyDate Then
                MsgBox(DS.Tables("table1").Rows(i).Item("料件编号").ToString.Trim() & "交货日期不可小于请购日期")
                Label6.Text = "失败"
                Return
            End If
            oCommander2.CommandText = "SELECT (ima48+ima49+ima491+ima50) from ima_file where ima01 = '" & DS.Tables("table1").Rows(i).Item("料件编号").ToString.Trim() & "'"
            Dim l_v2 As Decimal = oCommander2.ExecuteScalar()
            If modifyDate.AddDays(l_v2) > DS.Tables("table1").Rows(i).Item("交货日期") Then
                MsgBox(DS.Tables("table1").Rows(i).Item("料件编号").ToString.Trim() & "前置日期为" & l_v2 & "天,交货日期不可小于推算日")
                Label6.Text = "失败"
                Return
            End If
            oCommander2.CommandText = "SELECT COUNT(*) FROM pja_file WHERE pja01 = '" & DS.Tables("table1").Rows(i).Item("项目编号").ToString.Trim() & "' AND pjaacti = 'Y'"
            Dim l_v3 As Int16 = oCommander2.ExecuteScalar()
            If l_v3 <= 0 Then
                MsgBox(DS.Tables("table1").Rows(i).Item("料件编号").ToString.Trim() & "专案号有误")
                Label6.Text = "失败"
                Return
            End If
        Next
        Label6.Text = "检查完毕，汇入中"
        ' 檢查完畢, 開始建單頭, 先給定號
        Dim HC1 As String = "D3101-" & modifyDate.ToString("yyMM")
        Dim MonthA As Int16 = modifyDate.ToString("MM")
        oCommander2.CommandText = "select nvl(max(pmk01),'N') from pmk_file where pmk01 like '" & HC1 & "%'"
        Dim HC2 As String = oCommander2.ExecuteScalar()
        If HC2 = "N" Then ' 表示沒有任何單, 從1號開始
            HC1 += "0001"
        Else
            Dim HC3 As Int16 = Strings.Right(HC2, 4)
            HC3 += 1
            Dim HC4 As Int16 = Strings.Len(Convert.ToString(HC3))
            Select Case HC4
                Case 1
                    HC1 += "000" & HC3
                Case 2
                    HC1 += "00" & HC3
                Case 3
                    HC1 += "0" & HC3
                Case 4
                    HC1 += "HC3"
            End Select
        End If
        Dim DT As Oracle.ManagedDataAccess.Client.OracleTransaction = oConnection.BeginTransaction()
        oCommand.CommandText = "INSERT INTO pmk_file VALUES ('" & HC1 & "','" & Pmktype & "',0,to_date('" & modifyDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),NULL,"
        oCommand.CommandText += "NULL,NULL,NULL,NULL,NULL,NULL,'" & TextBox1.Text & "','" & TextBox2.Text & "',NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,0,NULL,to_date('" & modifyDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),"
        oCommand.CommandText += "NULL,NULL,'Y'," & modifyDate.ToString("yy") & "," & MonthA & ",0,0,NULL,NULL,0,'N','Y',0,NULL,'N',' ',0,NULL,0,0,'Y','" & TextBox1.Text & "','" & TextBox2.Text & "',"
        oCommand.CommandText += "NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,'ACTIONTEST','" & Now.ToString("HH/mm/ss") & "',NULL,NULL,"
        oCommand.CommandText += "to_date('" & modifyDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),NULL,'ACTIONTEST','ACTIONTEST','" & TextBox1.Text & "','" & TextBox2.Text & "',NULL)"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            DT.Rollback()
            Label6.Text = "失败"
            MsgBox(ex.Message())
            Return
        End Try
        ' 處理單身
        For i As Integer = 0 To DS.Tables("table1").Rows.Count - 1 Step 1
            ' 查會計科目
            oCommander2.CommandText = "SELECT imz39 from ima_file,imz_file where ima06 = imz01 and ima01 = '" & DS.Tables("table1").Rows(i).Item("料件编号").ToString().Trim() & "'"
            Dim imz39 = oCommander2.ExecuteScalar()

            oCommand.CommandText = "INSERT INTO pml_file VALUES ('" & HC1 & "','" & Pmktype & "'," & i + 1 & ",NULL,'" & DS.Tables("table1").Rows(i).Item("料件编号").ToString().Trim() & "','"
            oCommand.CommandText += DS.Tables("table1").Rows(i).Item("品名").ToString().Trim() & "',NULL,'','" & DS.Tables("table1").Rows(i).Item("请购单位").ToString().Trim() & "','"
            oCommand.CommandText += DS.Tables("table1").Rows(i).Item("请购单位").ToString().Trim() & "',1,NULL,'N','" & DS.Tables("table1").Rows(i).Item("项目编号").ToString().Trim() & "',' ',' ',NULL,"
            oCommand.CommandText += "20,'Y','Y',0,NULL," & DS.Tables("table1").Rows(i).Item("请购数量") & ",0,'Y',0,NULL,0," & "to_date('" & DS.Tables("table1").Rows(i).Item("交货日期") & "','yyyy/mm/dd'),"
            oCommand.CommandText += "to_date('" & DS.Tables("table1").Rows(i).Item("交货日期") & "','yyyy/mm/dd')," & "to_date('" & DS.Tables("table1").Rows(i).Item("交货日期") & "','yyyy/mm/dd'),"
            oCommand.CommandText += "'N','" & imz39 & "',NULL,0,NULL,NULL,0,NULL,NULL,NULL,'" & TextBox2.Text & "',NULL,NULL,NULL,NULL,NULL,NULL,'" & DS.Tables("table1").Rows(i).Item("请购单位").ToString().Trim() & "',"
            oCommand.CommandText += DS.Tables("table1").Rows(i).Item("请购数量") & ",NULL,NULL,NULL,'N',NULL,'N',NULL,NULL,NULL,NULL,' ','" & DS.Tables("table1").Rows(i).Item("厂商料号（备注）") & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,"
            oCommand.CommandText += "1,1,NULL,NULL,NULL,2,NULL,1,'ACTIONTEST','ACTIONTEST','N',NULL,NULL,NULL,NULL,NULL,NULL,NULL)"
            Try
                oCommand.ExecuteNonQuery()
            Catch ex As Exception
                DT.Rollback()
                Label6.Text = "失败"
                MsgBox(ex.Message())
                Return
            End Try
        Next
        DT.Commit()
        Label6.Text = "完成"
        TextBox3.Text = HC1
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Not String.IsNullOrEmpty(TextBox3.Text) Then
            Clipboard.SetText(TextBox3.Text)
        End If
    End Sub
End Class