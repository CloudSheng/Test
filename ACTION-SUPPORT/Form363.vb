Public Class Form363
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand3 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader3 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim tYear As Decimal = 0
    Dim tWeek As Decimal = 0
    Dim temp_cnt As Integer = 0
    Dim OpenFileDialog1 As Object

    Private Sub Form363_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        tYear = Today.Year
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
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog2.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog2.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog2.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            
            Dim ExcelString = "SELECT ERPPN,EOP_YW FROM [sheet1$] "

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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        For I As Integer = 0 To DataGridView1.Rows.Count - 1 Step 1
            oCommand3.CommandText = "select count(*) as temp_cnt from tc_ime_file "
            oCommand3.CommandText += " WHERE tc_ime01 = '" & DataGridView1.Rows(I).Cells("ERPPN").Value & "'"
            oReader3 = oCommand3.ExecuteReader
            If oReader3.HasRows() Then
                oReader3.Read()
                temp_cnt = oReader3.Item("temp_cnt")
            End If
            oReader3.Close()

            If temp_cnt > 0 Then
                oCommand.CommandText = "update tc_ime_file SET tc_ime08 = '" & DataGridView1.Rows(I).Cells("EOP_YW").Value & "'"
                oCommand.CommandText += "       ,tc_ime02 = to_date('" & Today & "','yyyy/mm/dd')"
                oCommand.CommandText += " WHERE tc_ime01 = '" & DataGridView1.Rows(I).Cells("ERPPN").Value & "'"
                Try
                    oCommand.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            Else
                oCommand.CommandText = "INSERT INTO tc_ime_file (tc_ime01,tc_ime08) VALUES ('"
                oCommand.CommandText += DataGridView1.Rows(I).Cells("ERPPN").Value & "','" & DataGridView1.Rows(I).Cells("EOP_YW").Value & "')"
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