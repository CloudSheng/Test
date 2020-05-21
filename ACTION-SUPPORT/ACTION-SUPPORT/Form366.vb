Public Class Form366
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim tYear As Decimal = 0
    Dim tWeek As Decimal = 0
    Dim c_year As String = String.Empty
    Dim n_year As Integer = 0
    Dim c_date As String = String.Empty
    'Dim c_mon As String = String.Empty
    'Dim n_mon As Integer = 0
    Dim c_tc_azn02 As Decimal = 0
    Dim c_tc_azn05 As Decimal = 0
    Dim n_tc_azn02 As Integer = 0
    Dim n_tc_azn05 As Integer = 0

    Private Sub Form116_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        tYear = Today.Year
        Me.TextBox1.Text = tYear
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        oCommand.CommandText = "SELECT tc_azn05 FROM tc_azn_file WHERE tc_azn01 = to_date('" & Today.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        tWeek = oCommand.ExecuteScalar()
        Me.TextBox2.Text = tWeek
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            'Dim ExcelString = "SELECT ERPPN,Year,WeekNum,Inbound FROM [sheet1$] Where Year > " & tYear & " or (year = " & tYear & " and WeekNum >=" & tWeek & ")"
            Dim ExcelString = "SELECT ERPPN,Year,WeekNum,Inbound FROM [sheet1$] Where Year > " & tYear & " or (year = " & tYear & " and WeekNum >=" & Today & ")"
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
        For I As Integer = 0 To DataGridView1.Rows.Count - 1 Step 1
            oCommand.CommandText = "DELETE tc_prp_file WHERE tc_prp01 = '" & DataGridView1.Rows(I).Cells("ERPPN").Value & "'"
            Try
                oCommand.ExecuteNonQuery()
            Catch ex As Exception
                'MsgBox(ex.Message())
                'Return
            End Try
        Next

        ' 匯入Datagridview 資料
        For i As Integer = 0 To DataGridView1.Rows.Count - 1 Step 1
            'c_year = DataGridView1.Rows(i).Cells("Year").Value
            'n_year = Val(c_year)
            ''c_year = Mid(l_year, 1, Len(l_year) - 1)
            'c_mon = DataGridView1.Rows(i).Cells("WeekNum").Value
            'n_mon = Val(c_mon)
            ''c_mon = Mid(l_mon, 1, Len(l_mon) - 1)
            'oCommand.CommandText = " select max(tc_azn01) as tc_azn01 from tc_azn_file "
            'oCommand.CommandText += " where tc_azn02 = " & n_year & " and tc_azn05 = " & n_mon
            'oReader = oCommand.ExecuteReader()
            'If oReader.HasRows Then
            '    oReader.Read()
            '    l_tc_azn01 = oReader.Item("tc_azn01")
            'End If
            'oReader.Close()
            'Me.Label3.Text = "INSERT DATA" & i
            'Me.Label3.Refresh()
            'oCommand.CommandText = "INSERT INTO tc_prp_file (tc_prp01,tc_prp02,tc_prp03,tc_prp04,tc_prp05,tc_prplegal) VALUES ('"
            'oCommand.CommandText += DataGridView1.Rows(i).Cells("ERPPN").Value & "'," & DataGridView1.Rows(i).Cells("Year").Value & "," & DataGridView1.Rows(i).Cells("WeekNum").Value & "," & DataGridView1.Rows(i).Cells("Inbound").Value & ",to_date('" & l_tc_azn01 & "','yyyy/MM/dd'),'ACTIONTEST')"

            'c_year = DataGridView1.Rows(i).Cells("Year").Value
            'n_year = Val(c_year)
            c_date = DataGridView1.Rows(i).Cells("WeekNum").Value
            oCommand.CommandText = " select tc_azn02,tc_azn05 from tc_azn_file "
            oCommand.CommandText += " where tc_azn01 = to_date('" & c_date & "','yyyy/mm/dd')"
            oReader = oCommand.ExecuteReader()
            If oReader.HasRows Then
                oReader.Read()
                If Not oReader.Item("tc_azn02") Is DBNull.Value And Not oReader.Item("tc_azn05") Is DBNull.Value Then
                    c_tc_azn02 = oReader.Item("tc_azn02")
                    n_tc_azn02 = Val(c_tc_azn02)
                    c_tc_azn05 = oReader.Item("tc_azn05")
                    n_tc_azn05 = Val(c_tc_azn05)
                End If
            End If
            oReader.Close()

            Me.Label3.Text = "INSERT DATA" & i
            Me.Label3.Refresh()
            oCommand.CommandText = "INSERT INTO tc_prp_file (tc_prp01,tc_prp02,tc_prp03,tc_prp04,tc_prp05,tc_prplegal) VALUES ('"
            oCommand.CommandText += DataGridView1.Rows(i).Cells("ERPPN").Value & "'," & n_tc_azn02 & "," & n_tc_azn05 & "," & DataGridView1.Rows(i).Cells("Inbound").Value & ",to_date('" & DataGridView1.Rows(i).Cells("WeekNum").Value & "','yyyy/MM/dd'),'ACTIONTEST')"

            Try
                oCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        Next
        Me.Label3.Text = "FINISHED"
    End Sub
End Class