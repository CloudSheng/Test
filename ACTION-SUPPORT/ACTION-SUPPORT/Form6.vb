Public Class Form6
    Dim mConnection As New SqlClient.SqlConnection
    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        Me.Label2.Text = mConnection.State.ToString()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        Me.Label2.Text = mConnection.State.ToString()
    End Sub
End Class