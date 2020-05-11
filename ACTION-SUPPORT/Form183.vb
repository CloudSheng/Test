Public Class Form183
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim tModel As String
    Dim tValue As String
    
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.SelectedIndex = 0 Then
            MsgBox("请选择型号")
            Return
        End If
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        tModel = ComboBox1.SelectedItem.ToString()
        tValue = TextBox1.Text
        mSQLS1.CommandText = "Select Count(*) from paravalue left join lot on paravalue.lot = lot.lot where parameter = 'Mold_ID' and station in ('0150','0151') and lot.model = '"
        mSQLS1.CommandText += tModel & "'"
        If Not String.IsNullOrEmpty(tValue) Then
            mSQLS1.CommandText += " and paravalue.value like '" & tValue & "%'"
        End If
        Dim CA As Integer = mSQLS1.ExecuteScalar()
        Label5.Text = CA


        mSQLS1.CommandText = "Select Count(*) from scrap_paravalue left join lot on scrap_paravalue.lot = lot.lot where parameter = 'Mold_ID' and station in ('0150','0151') and lot.model = '"
        mSQLS1.CommandText += tModel & "'"
        If Not String.IsNullOrEmpty(tValue) Then
            mSQLS1.CommandText += " and scrap_paravalue.value like '" & tValue & "%'"
        End If
        Dim CB As Integer = mSQLS1.ExecuteScalar()
        Label6.Text = CB

        Label8.Text = CA + CB
    End Sub

    Private Sub Form183_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        BindModel()
    End Sub
    Private Sub BindModel()
        Me.ComboBox1.Items.Clear()
        mSQLS1.CommandText = "select distinct model  from model where model.model_type <> 'Action'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox1.Items.Add(mSQLReader.Item(0).ToString())
            End While
        End If
        mSQLReader.Close()
    End Sub

End Class