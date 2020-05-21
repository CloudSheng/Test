
Imports Microsoft.VisualBasic.Strings
Public Class Form305
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
                MsgBox("单据录入人员账号有误")
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

    Private Sub Form305_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        'oCommander2.CommandText = "SELECT SMA53 FROM SMA_FILE"
        ' Dim l_d1 As Date = oCommander2.ExecuteScalar()
        ' If l_d1 >= DateTimePicker1.Value Then
        If DateTimePicker1.Value < Now.Date Then
            MsgBox("变更日期不可小于作业当日日期")
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
            'IE不想每次将工作簿改为sheet1，默认读取工作簿名称。
            Dim SheetName As String = Excelconn.GetSchema("Tables").Rows(0)("TABLE_NAME").ToString.Trim()
            Dim ExcelString = "SELECT * FROM [" & SheetName & "]"
            'Dim ExcelString = "SELECT * FROM [sheet1$]"
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
            MsgBox("请输入变更人员ERP账号")
            TextBox1.Focus()
            Label6.Text = "失败"
            Return
        End If
        
        If IsNothing(DS.Tables("table1")) Then
            MsgBox("无单身资料，请检查")
            Label6.Text = "失败"
            Return
        End If
        ' 單身資料
        For i As Integer = 0 To DS.Tables("table1").Rows.Count - 1 Step 1
            If String.IsNullOrEmpty(DS.Tables("table1").Rows(i).Item("料件编号").ToString) Or _
                String.IsNullOrEmpty(DS.Tables("table1").Rows(i).Item("变更后工时").ToString) Or _
                String.IsNullOrEmpty(DS.Tables("table1").Rows(i).Item("变更原因").ToString) Then
                MsgBox("单身资料不完整，请检查")
                Label6.Text = "失败"
                Return
            End If
            '检查变更后工时
            If DS.Tables("table1").Rows(i).Item("变更后工时") <= 0 Then
                MsgBox(DS.Tables("table1").Rows(i).Item("料件编号").ToString.Trim() & "工时有误，请检查")
                Label6.Text = "失败"
                Return
            End If
            '检查料件编号是否是‘M：自制料件’。
            oCommander2.CommandText = "SELECT COUNT(*) FROM ima_file WHERE ima01 = '" & DS.Tables("table1").Rows(i).Item("料件编号").ToString.Trim() & "' AND imaacti = 'Y' and ima08 = 'M'"
            Dim l_v1 As Int16 = oCommander2.ExecuteScalar()
            If l_v1 <= 0 Then
                MsgBox("单身料号" & DS.Tables("table1").Rows(i).Item("料件编号").ToString.Trim() & "有误，请检查")
                Label6.Text = "失败"
                Return
            End If
            
            '检查备注栏长度
            If System.Text.ASCIIEncoding.Default.GetByteCount(DS.Tables("table1").Rows(i).Item("变更原因").ToString) > 52 Then
                MsgBox(DS.Tables("table1").Rows(i).Item("料件编号").ToString.Trim() & "备注过长，请检查")
                Return
            End If

        Next
        Label6.Text = "检查完毕，汇入中"
        ' 檢查完畢, 開始建單頭, 先給定號
        Dim HC1 As String = String.Empty

        HC1 = "CI-" & modifyDate.ToString("yyyyMMdd")

        Dim MonthA As Int16 = modifyDate.ToString("MM")
        oCommander2.CommandText = "select nvl(max(TC_CID_01),'N') from tc_cid_file where TC_CID_01 like '" & HC1 & "%'"
        Dim HC2 As String = oCommander2.ExecuteScalar()
        If HC2 = "N" Then ' 表示沒有任何單, 從1號開始
            HC1 += "00001"
        Else
            Dim HC3 As Int16 = Strings.Right(HC2, 5)
            HC3 += 1
            Dim HC4 As Int16 = Strings.Len(Convert.ToString(HC3))
            Select Case HC4
                Case 1
                    HC1 += "0000" & HC3
                Case 2
                    HC1 += "000" & HC3
                Case 3
                    HC1 += "00" & HC3
                Case 4
                    HC1 += "0" & HC3
                Case 5
                    HC1 += "HC3"
            End Select
        End If

        Dim DT As Oracle.ManagedDataAccess.Client.OracleTransaction = oConnection.BeginTransaction()
        oCommand.CommandText = "INSERT INTO tc_cid_file (tc_cid_01,tc_cid_02,tc_cid_03,tc_cid_04,tc_cid_05,tc_cid_confirm,tc_cid_acti,tc_cidplant) VALUES ('" & HC1 & "','" & TextBox1.Text & "',to_date('" & modifyDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),NULL,NULL,'N','Y','ACTIONTEST')"

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
            '查最后人工工时
            oCommander2.CommandText = "SELECT ima58 from ima_file  where ima01 = '" & DS.Tables("table1").Rows(i).Item("料件编号").ToString().Trim() & "' AND imaacti = 'Y' and ima08='M'"
            Dim ima58 As Decimal = oCommander2.ExecuteScalar()
            '查最后机器工时
            oCommander2.CommandText = "SELECT ima912 from ima_file  where ima01 = '" & DS.Tables("table1").Rows(i).Item("料件编号").ToString().Trim() & "' AND imaacti = 'Y' and ima08='M'"
            Dim ima912 As Decimal = oCommander2.ExecuteScalar()
            'IE确定只更新人工工时，新机器工时默认为0.
            Dim ASS As String = DS.Tables("table1").Rows(i).Item("变更原因").ToString().Trim()
            ASS = ASS.Replace("'", "''")
            'oCommand.CommandText = "INSERT INTO tc_cie_file VALUES ('" & HC1 & "'," & i + 1 & ",'" & DS.Tables("table1").Rows(i).Item("料件编号").ToString().Trim() & "','" & ima58 & "','" & ima912 & "','" & DS.Tables("table1").Rows(i).Item("变更后工时").ToString().Trim() & "',0,'" & DS.Tables("table1").Rows(i).Item("变更原因").ToString().Trim() & "',NULL)"
            oCommand.CommandText = "INSERT INTO tc_cie_file VALUES ('" & HC1 & "'," & i + 1 & ",'" & DS.Tables("table1").Rows(i).Item("料件编号").ToString().Trim() & "','" & ima58 & "','" & ima912 & "','" & DS.Tables("table1").Rows(i).Item("变更后工时").ToString().Trim() & "',0,'" & ASS & "',NULL)"
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
        MsgBox("资料导入成功！")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Not String.IsNullOrEmpty(TextBox3.Text) Then
            Clipboard.SetText(TextBox3.Text)
        End If
    End Sub
End Class