Imports System.Security.Principal
Public Class Form12

    Private Sub Form12_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Me.Label2.Text = oConnection.State.ToString()
        Label4.Text = My.Computer.Name
        Dim Uss As String = Environment.GetEnvironmentVariable("USERNAME")
        Label6.Text = Uss
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
            oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
            If oConnection.State <> ConnectionState.Open Then
                Try
                    oConnection.Open()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
            Me.Label2.Text = oConnection.State.ToString()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
End Class