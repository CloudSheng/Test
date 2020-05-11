Public Class Form351
    'Dim Sda As New Oracle.ManagedDataAccess.Client.OracleDataAdapter
    'Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    'Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    'Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim Ds As New DataSet()
    Dim Sda As New SqlClient.SqlDataAdapter
    Dim conn As New SqlClient.SqlConnection()
    Private Function OpenConnectionOfTESTMes()
        Dim mConnectionBuilder As New SqlClient.SqlConnectionStringBuilder
        mConnectionBuilder.DataSource = "192.168.10.254"
        mConnectionBuilder.InitialCatalog = "IQMES-TEST"
        mConnectionBuilder.IntegratedSecurity = False
        mConnectionBuilder.MultipleActiveResultSets = True
        mConnectionBuilder.UserID = "sa"
        mConnectionBuilder.Password = "p@$$w0rd"
        Return mConnectionBuilder.ConnectionString
    End Function
    Private Sub Form351_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        'If oConnection.State <> ConnectionState.Open Then
        'Try
        'oConnection.Open()
        'oCommand.Connection = oConnection
        'oCommand.CommandType = CommandType.Text
        'Catch ex As Exception
        'MsgBox(ex.Message)
        'End Try
        'End If
        'CreateTempTable()
        'Sda = New Oracle.ManagedDataAccess.Client.OracleDataAdapter("select * from ord_temp", conn)        

        conn.ConnectionString = OpenConnectionOfTESTMes()
        If conn.State <> ConnectionState.Open Then
            Try
                conn.Open()
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        Sda = New SqlClient.SqlDataAdapter("select * from mpl_ord_temp", conn)
        Sda.Fill(Ds)
        Me.DataGridView1.DataSource = Ds.Tables(0)
        Me.DataGridView1.Show()

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Ds.HasChanges() Then
            Dim cb As New SqlClient.SqlCommandBuilder(Sda)
            Sda.Update(Ds.Tables(0))
            Ds.Tables(0).AcceptChanges()
            Me.DataGridView1.Update()
        End If
    End Sub
    'Private Sub CreateTempTable()
    'Me.Label2.Text = "DROP TABLE"
    '    oCommand.CommandText = "DROP TABLE ord_temp"

    '    Try
    '        oCommand.ExecuteNonQuery()
    '    Catch ex As Exception
    '    End Try

    'Me.Label2.Text = "CREATE TABLE"
    '    oCommand.CommandText = "CREATE TABLE ord_temp (ord_date date, ord_no varchar2(20), ord_item varchar2(40)) "

    '    Try
    '        oCommand.ExecuteNonQuery()
    '    Catch ex As Exception
    '        MsgBox(ex.Message())
    '    End Try

    'End Sub
End Class