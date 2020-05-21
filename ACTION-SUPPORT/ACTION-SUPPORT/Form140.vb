Public Class Form140
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim l_ima69 As Decimal = 0
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Me.TextBox1.Enabled = False
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Me.TextBox1.Enabled = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Enabled = True And IsDBNull(TextBox1.Text) Then
            MsgBox("请输入料号")
            Return
        End If

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
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Form140_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Recount()
    End Sub
    Private Sub Recount()
        If TextBox1.Enabled = True Then
            oCommand.CommandText = "select distinct ima01,nvl(ima58,0) as ima58 from ima_file,bma_file where ima01 = bma01 and imaacti = 'Y' and bmaacti = 'Y' and ima08 = 'M' and ima910 = bma06 and bma01 LIKE '" & TextBox1.Text & "'"
        Else
            oCommand.CommandText = "select distinct ima01,nvl(ima58,0) as ima58 from ima_file,bma_file where ima01 = bma01 and ima06 = 103 and imaacti = 'Y' and bmaacti = 'Y' and ima910 = bma06 "
        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                l_ima69 = 0
                Label2.Text = oReader.Item("ima01")
                Label2.Refresh()
                CountBOM(oReader.Item("ima01"))
                l_ima69 += oReader.Item("ima58")
                oCommander2.CommandText = "Update ima_file SET ima69 = " & l_ima69 & " WHERE ima01 = '" & oReader.Item("ima01") & "'"
                Try
                    oCommander2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End While
        End If
        oReader.Close()

    End Sub
    Private Sub CountBOM(ByVal erp1 As String)
        Dim oCommander97 As New Oracle.ManagedDataAccess.Client.OracleCommand
        Dim oreader97 As Oracle.ManagedDataAccess.Client.OracleDataReader
        oCommander97.Connection = oConnection
        oCommander97.CommandType = CommandType.Text

        oCommander97.CommandText = "select * from bmb_file full join ima_file on bmb03 = ima01 where bmb01 = '"
        oCommander97.CommandText += erp1 & "' and bmb05 is NULL and bmb29 = ima910 and bmb19 = 2 order by bmb03"
        oreader97 = oCommander97.ExecuteReader()
        If oreader97.HasRows() Then
            While oreader97.Read()
                If IsDBNull(oreader97.Item("ima58")) Then
                    Continue While
                End If
                l_ima69 += oreader97.Item("ima58")
                Call CountBOM(oreader97.Item("bmb03"))
            End While

        End If
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        MsgBox("Finished")
    End Sub
End Class