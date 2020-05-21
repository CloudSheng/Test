Public Class Form8
    Dim Ds As New DataSet()
    Dim Sda As New SqlClient.SqlDataAdapter
    Dim conn As New SqlClient.SqlConnection()
    Private Sub Form8_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        
        conn.ConnectionString = Module1.OpenConnectionOfMes()
        If conn.State <> ConnectionState.Open Then
            Try
                conn.Open()
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        Sda = New SqlClient.SqlDataAdapter("select * from inventoryCountPrepare", conn)
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
End Class