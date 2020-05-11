Public Class Form17

    Private Sub Form17_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label2.Text = My.Application.Info.Version.ToString()
        Label4.Text = My.Application.Info.Description
        Try
            Label6.Text = My.Application.Deployment.CurrentVersion.ToString()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

    End Sub
End Class