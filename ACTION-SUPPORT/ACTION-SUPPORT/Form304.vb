Imports Microsoft.VisualBasic.Strings
Public Class Form304
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

    Private Sub Form304_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
            Dim ExcelString = "SELECT * FROM [sheet2$]"
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
            MsgBox("请输入杂发人员ERP账号")
            TextBox1.Focus()
            Label6.Text = "失败"
            Return
        End If
        If IsNothing(ComboBox1.SelectedItem) Then
            MsgBox("请选择仓库杂项单据性质")
            Label6.Text = "失败"
            Return
        End If
        Dim Inatype As String = Strings.Left(ComboBox1.SelectedItem.ToString(), 1)
        'If DS.Tables("table1").Rows.Count = 0 Then
        If IsNothing(DS.Tables("table1")) Then
            MsgBox("无单身资料，请检查")
            Label6.Text = "失败"
            Return
        End If
        ' 單身資料
        For i As Integer = 0 To DS.Tables("table1").Rows.Count - 1 Step 1
            If String.IsNullOrEmpty(DS.Tables("table1").Rows(i).Item("品號").ToString) Or _
                String.IsNullOrEmpty(DS.Tables("table1").Rows(i).Item("申请数量").ToString) Or _
                String.IsNullOrEmpty(DS.Tables("table1").Rows(i).Item("理由").ToString) Or _
                String.IsNullOrEmpty(DS.Tables("table1").Rows(i).Item("单位").ToString) Or _
                String.IsNullOrEmpty(DS.Tables("table1").Rows(i).Item("仓库").ToString) Then
                MsgBox("单身资料不完整，请检查")
                Label6.Text = "失败"
                Return
            End If
            '检查申请数量
            If DS.Tables("table1").Rows(i).Item("申请数量") <= 0 Then
                MsgBox(DS.Tables("table1").Rows(i).Item("品號").ToString.Trim() & "数量有误，请检查")
                Label6.Text = "失败"
                Return
            End If
            '与仓库确定，要求只导入半成品资料，所以只限定检查是否是‘M：自制料件’的料件编号
            oCommander2.CommandText = "SELECT COUNT(*) FROM ima_file WHERE ima01 = '" & DS.Tables("table1").Rows(i).Item("品號").ToString.Trim() & "' AND imaacti = 'Y' and ima08 = 'M'"
            Dim l_v1 As Int16 = oCommander2.ExecuteScalar()
            If l_v1 <= 0 Then
                MsgBox("单身料号" & DS.Tables("table1").Rows(i).Item("品號").ToString.Trim() & "有误，请检查")
                Label6.Text = "失败"
                Return
            End If
            '检查仓库编号
            oCommander2.CommandText = "SELECT count(*) from imd_file where imd01 = '" & DS.Tables("table1").Rows(i).Item("仓库").ToString.Trim() & "'"
            Dim l_v2 As Int16 = oCommander2.ExecuteScalar()
            If l_v2 <= 0 Then
                MsgBox(DS.Tables("table1").Rows(i).Item("仓库").ToString.Trim() & "仓库编号有误，请检查")
                Label6.Text = "失败"
                Return
            End If
            '检查库存单位
            oCommander2.CommandText = "SELECT COUNT(*) FROM ima_file WHERE ima01 = '" & DS.Tables("table1").Rows(i).Item("品號").ToString.Trim() & "' AND imaacti = 'Y' and ima08 = 'M'"
            oCommander2.CommandText += " and ima25 = '" & DS.Tables("table1").Rows(i).Item("单位").ToString.Trim() & "'"
            Dim l_v3 As Int16 = oCommander2.ExecuteScalar()
            If l_v3 <= 0 Then
                MsgBox(DS.Tables("table1").Rows(i).Item("品號").ToString.Trim() & "单位有误，请检查")
                Label6.Text = "失败"
                Return
            End If
            '检查理由码
            oCommander2.CommandText = "SELECT COUNT(*) FROM azf_file WHERE azf01 = '" & DS.Tables("table1").Rows(i).Item("理由").ToString.Trim() & "' AND  azfacti = 'Y' and azf02 = '2'"
            Dim l_v4 As Int16 = oCommander2.ExecuteScalar()
            If l_v4 <= 0 Then
                MsgBox(DS.Tables("table1").Rows(i).Item("品號").ToString.Trim() & "理由码有误，请检查")
                Label6.Text = "失败"
                Return
            End If
            '检查备注栏长度
            If System.Text.ASCIIEncoding.Default.GetByteCount(DS.Tables("table1").Rows(i).Item("备注").ToString) > 40 Then
                MsgBox(DS.Tables("table1").Rows(i).Item("品號").ToString.Trim() & "备注过长，请检查")
                Return
            End If

        Next
        Label6.Text = "检查完毕，汇入中"
        ' 檢查完畢, 開始建單頭, 先給定號
        Dim HC1 As String = String.Empty
        Select Case Inatype
            Case 1
                HC1 = "D1101-" & modifyDate.ToString("yyMM")
            Case 3
                HC1 = "D1201-" & modifyDate.ToString("yyMM")
            Case 5
                HC1 = "D1301-" & modifyDate.ToString("yyMM")
        End Select

        Dim MonthA As Int16 = modifyDate.ToString("MM")
        oCommander2.CommandText = "select nvl(max(ina01),'N') from ina_file where  ina01 like '" & HC1 & "%'"
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
        oCommand.CommandText = "INSERT INTO ina_file VALUES ('" & Inatype & "','" & HC1 & "',to_date('" & modifyDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),to_date('" & modifyDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),'" & TextBox2.Text & "',NULL,NULL,NULL,"
        oCommand.CommandText += "0,NULL,NULL,NULL,'N','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox1.Text & "',to_date('" & modifyDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),'N','" & TextBox1.Text & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,"
        oCommand.CommandText += "NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',0,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'N',"
        oCommand.CommandText += "'ACTIONTEST','ACTIONTEST','" & TextBox1.Text & "','" & TextBox2.Text & "',NULL)"
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
            ' 查仓库名称，个别仓库名称长度大于inbud02栏位长度，无法插入数据。
            oCommander2.CommandText = "SELECT substr(imd02,1,17)  from imd_file  where imd01 = '" & DS.Tables("table1").Rows(i).Item("仓库").ToString().Trim() & "'"
            Dim imd02 As String = oCommander2.ExecuteScalar()
            '查自制料件的库存单位换算率
            oCommander2.CommandText = "SELECT ima55_fac from ima_file  where ima01 = '" & DS.Tables("table1").Rows(i).Item("品號").ToString().Trim() & "' AND imaacti = 'Y' and ima08='M'"
            Dim inbima55_fac As Int16 = oCommander2.ExecuteScalar()
            ' 查料件来源码
            oCommander2.CommandText = "SELECT ima08 from ima_file  where ima01 = '" & DS.Tables("table1").Rows(i).Item("品號").ToString().Trim() & "'"
            Dim ima08 As String = oCommander2.ExecuteScalar()
            Dim inbima08 As String = String.Empty
            Select Case ima08
                Case "M"
                    inbima08 = "M:自製料件"
                Case "P"
                    inbima08 = "P:採購料件"
                Case "S"
                    inbima08 = "S:委外加工料件"
            End Select

            oCommand.CommandText = "INSERT INTO inb_file VALUES ('" & HC1 & "'," & i + 1 & ",'" & DS.Tables("table1").Rows(i).Item("品號").ToString().Trim() & "','" & DS.Tables("table1").Rows(i).Item("仓库").ToString().Trim() & "',' ',' ','"
            oCommand.CommandText += DS.Tables("table1").Rows(i).Item("单位").ToString().Trim() & "','" & inbima55_fac & "','" & DS.Tables("table1").Rows(i).Item("申请数量") & "','N',NULL,NULL,0,NULL,'" & DS.Tables("table1").Rows(i).Item("理由").ToString().Trim() & "',"
            oCommand.CommandText += "NULL,NULL,NULL,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" & imd02 & "','" & DS.Tables("table1").Rows(i).Item("备注").ToString().Trim() & "','" & inbima08 & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" & DS.Tables("table1").Rows(i).Item("申请数量").ToString().Trim() & "',NULL,"
            oCommand.CommandText += "NULL,NULL,NULL,NULL,NULL,'ACTIONTEST','ACTIONTEST',0,0,0,0,0,0,0)"

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