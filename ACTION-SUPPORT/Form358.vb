Public Class Form358
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand3 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader3 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim tYear As Decimal = 0
    Dim tWeek As Decimal = 0
    Dim temp_cnt As Integer = 0
    Dim OpenFileDialog1 As Object

    Private Sub Form358_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        tYear = Today.Year
        Me.TextBox1.Text = tYear
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommand3.Connection = oConnection
                oCommand3.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        oCommand.CommandText = "SELECT tc_azn05 FROM tc_azn_file WHERE tc_azn01 = to_date('" & Today.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        tWeek = oCommand.ExecuteScalar()
        Me.TextBox2.Text = tWeek
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog2.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog2.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog2.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            'Dim ExcelString = "SELECT ERPPN,Year,WeekNum,Inbound FROM [sheet1$] Where Year > " & tYear & " or (year = " & tYear & " and WeekNum >=" & tWeek & ")"
            Dim ExcelString = "SELECT ERPPN,Year,WeekNum,Inbound FROM [sheet1$] Where Year >= " & tYear
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)
            Dim DS As Data.DataSet = New DataSet()
            Try
                ExcelAdapater.Fill(DS, "table1")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            Me.DataGridView1.DataSource = DS.Tables("table1")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        tYear = TextBox1.Text
        '刪除所有當週週及之後的資料
        Me.Label3.Text = "DELETE DATA"
        '181226 add by Brady
        'For I As Integer = 0 To DataGridView1.Rows.Count - 1 Step 1   
        '    'oCommand.CommandText = "DELETE tc_prm_file WHERE tc_prm01 = '" & DataGridView1.Rows(I).Cells("ERPPN").Value & "'"
        '    Try
        '        oCommand.ExecuteNonQuery()
        '    Catch ex As Exception
        '        'MsgBox(ex.Message())
        '        'Return
        '    End Try
        'Next
        oCommand.CommandText = "DELETE FROM tc_prm_file"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            'MsgBox(ex.Message())
            'Return
        End Try
        '181226 add by Brady END

        ' 匯入Datagridview 資料
        For i As Integer = 0 To DataGridView1.Rows.Count - 1 Step 1
            Me.Label3.Text = "INSERT DATA" & i
            Me.Label3.Refresh()
            oCommand.CommandText = "INSERT INTO tc_prm_file (tc_prm01,tc_prm02,tc_prm03,tc_prm04,tc_prmlegal) VALUES ('"
            oCommand.CommandText += DataGridView1.Rows(i).Cells("ERPPN").Value & "'," & DataGridView1.Rows(i).Cells("Year").Value & "," & DataGridView1.Rows(i).Cells("WeekNum").Value & "," & DataGridView1.Rows(i).Cells("Inbound").Value & ",'ACTIONTEST')"
            Try
                oCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        Next
        Me.Label3.Text = "FINISHED"
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        tYear = TextBox1.Text
        For I As Integer = 0 To DataGridView1.Rows.Count - 1 Step 1
            oCommand3.CommandText = "select count(*) as temp_cnt from tc_prm_file "
            oCommand3.CommandText += " WHERE tc_prm01 = '" & DataGridView1.Rows(I).Cells("ERPPN").Value & "' AND "
            oCommand3.CommandText += "       tc_prm02 = " & DataGridView1.Rows(I).Cells("Year").Value & " AND "
            oCommand3.CommandText += "       tc_prm03 = " & DataGridView1.Rows(I).Cells("WeekNum").Value
            oReader3 = oCommand3.ExecuteReader
            If oReader3.HasRows() Then
                oReader3.Read()
                temp_cnt = oReader3.Item("temp_cnt")
            End If
            oReader3.Close()

            If temp_cnt > 0 Then
                oCommand.CommandText = "update tc_prm_file SET tc_prm04 = '" & DataGridView1.Rows(I).Cells("Inbound").Value & "'"
                oCommand.CommandText += " WHERE tc_prm01 = '" & DataGridView1.Rows(I).Cells("ERPPN").Value & "' AND "
                oCommand.CommandText += "       tc_prm02 = " & DataGridView1.Rows(I).Cells("Year").Value & " AND "
                oCommand.CommandText += "       tc_prm03 = " & DataGridView1.Rows(I).Cells("WeekNum").Value
                Try
                    oCommand.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            Else
                oCommand.CommandText = "INSERT INTO tc_prm_file (tc_prm01,tc_prm02,tc_prm03,tc_prm04,tc_prmlegal) VALUES ('"
                oCommand.CommandText += DataGridView1.Rows(I).Cells("ERPPN").Value & "'," & DataGridView1.Rows(I).Cells("Year").Value & "," & DataGridView1.Rows(I).Cells("WeekNum").Value & "," & DataGridView1.Rows(I).Cells("Inbound").Value & ",'ACTIONTEST')"
                Try
                    oCommand.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    Return
                End Try
            End If
        Next
        Me.Label3.Text = "FINISHED"
    End Sub
End Class