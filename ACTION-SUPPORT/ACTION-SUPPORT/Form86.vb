Public Class Form86
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim DS As Data.DataSet = New DataSet()
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString = "SELECT Year,WK,bmb01,D1,D2,D3,D4,D5,D6,D7 FROM [sheet1$]"
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)
            'Dim DS As Data.DataSet = New DataSet()
            Try
                DS.Clear()
                ExcelAdapater.Fill(DS, "table1")
                Label2.Text = ExcelPath
                Label3.Text = "读入成功"
            Catch ex As Exception
                MsgBox(ex.Message())
                Label3.Text = "读入失败"
            End Try
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If IsNothing(DS.Tables("table1")) Then
            MsgBox("无单身资料，请检查")
            Label3.Text = "写入失败"
            Return
        End If
        Label3.Text = "开始写入"
        For i As Integer = 0 To DS.Tables("table1").Rows.Count - 1 Step 1
            For j As Integer = 3 To 9 Step 1
                If String.IsNullOrEmpty(DS.Tables("table1").Rows(i).Item(j).ToString()) Then
                    DS.Tables("table1").Rows(i).Item(j) = 0
                End If
            Next
        Next
        Dim DT As Oracle.ManagedDataAccess.Client.OracleTransaction = oConnection.BeginTransaction()
        For i As Integer = 0 To DS.Tables("table1").Rows.Count - 1 Step 1
            oCommand.CommandText = "INSERT INTO zaction VALUES (" & DS.Tables("table1").Rows(i).Item("Year").ToString() & ","
            oCommand.CommandText += DS.Tables("table1").Rows(i).Item("WK").ToString() & ",'"
            oCommand.CommandText += DS.Tables("table1").Rows(i).Item("bmb01").ToString().Trim() & "',"
            oCommand.CommandText += DS.Tables("table1").Rows(i).Item("D1").ToString() & ","
            oCommand.CommandText += DS.Tables("table1").Rows(i).Item("D2").ToString() & ","
            oCommand.CommandText += DS.Tables("table1").Rows(i).Item("D3").ToString() & ","
            oCommand.CommandText += DS.Tables("table1").Rows(i).Item("D4").ToString() & ","
            oCommand.CommandText += DS.Tables("table1").Rows(i).Item("D5").ToString() & ","
            oCommand.CommandText += DS.Tables("table1").Rows(i).Item("D6").ToString() & ","
            oCommand.CommandText += DS.Tables("table1").Rows(i).Item("D7").ToString() & ")"
            Try
                oCommand.ExecuteNonQuery()
            Catch ex As Exception
                DT.Rollback()
                Label3.Text = "失败"
                MsgBox(ex.Message())
                Return
            End Try
        Next
        DT.Commit()
        Label3.Text = "完成"
        MsgBox("资料导入成功！")
    End Sub

    Private Sub Form86_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
End Class