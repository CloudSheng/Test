Public Class Form367
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
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
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
            Dim ExcelString = "SELECT sys_no,close_date,vat_no,open_date FROM [sheet1$] "
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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'tYear = TextBox1.Text

        ' 匯入Datagridview 資料
        For i As Integer = 0 To DataGridView1.Rows.Count - 1 Step 1
            Me.Label3.Text = "UPDATE DATA" & i
            Me.Label3.Refresh()

            'tc_xma05,tc_xma03,tc_xmaud13
            oCommand.CommandText = "update tc_xma_file SET tc_xma05 = " & "to_date('" & DataGridView1.Rows(i).Cells("close_date").Value & "','yyyy/MM/dd'), "
            oCommand.CommandText += " tc_xma03 = '" & DataGridView1.Rows(i).Cells("vat_no").Value & "', "
            oCommand.CommandText += " tc_xmaud13 = " & "to_date('" & DataGridView1.Rows(i).Cells("open_date").Value & "','yyyy/MM/dd') "
            oCommand.CommandText += " WHERE tc_xma01 = '" & DataGridView1.Rows(i).Cells("sys_no").Value & "'"
            Try
                oCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        Next
        'oReader.Close()
        Me.Label3.Text = "FINISHED"
    End Sub
End Class