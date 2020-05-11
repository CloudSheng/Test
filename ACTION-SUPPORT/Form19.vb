Public Class Form19
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader

    Dim ServerString As String = String.Empty
    Dim NewServerString As String = String.Empty
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ServerString = "hkacttest"
        NewServerString = "action_hk"
        oConnection.ConnectionString = Module1.OpenOracleConnection(ServerString)
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
        Label3.Text = Now.ToShortTimeString()
        Label3.Refresh()
        Dim oAdapter As New Oracle.ManagedDataAccess.Client.OracleDataAdapter("", oConnection)
        oAdapter.SelectCommand.CommandText = "select zta01 from ZTA_FILE WHERE zta07  = 'T' and zta02 = '" & ServerString & "'"
        oAdapter.SelectCommand.Connection = oConnection
        oAdapter.Fill(DataSet1, "AllTables")
        Dim Allrows As Integer = DataSet1.Tables(0).Rows.Count()
        Label6.Text = Allrows
        Label6.Refresh()
        If Allrows > 0 Then
            For i = 1 To Allrows Step 1
                Label8.Text = i
                Label8.Refresh()
                Dim Tran As Oracle.ManagedDataAccess.Client.OracleTransaction = oConnection.BeginTransaction
                Dim Tables1 As String = DataSet1.Tables(0).Rows(i - 1).Item(0).ToString()
                oCommand.CommandText = "DELETE " & NewServerString & "." & Tables1
                Try
                    oCommand.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    Tran.Rollback()
                    Exit For
                End Try
                oCommand.CommandText = "INSERT INTO " & NewServerString & "." & Tables1 & " (SELECT * FROM "
                oCommand.CommandText += ServerString & "." & Tables1 & ")"
                Try
                    oCommand.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    Tran.Rollback()
                    Exit For
                End Try
                Dim CK As Int16 = Tables1.LastIndexOf("_file")
                Dim CutField As String = Tables1.Substring(0, CK)
                oCommand.CommandText = "UPDATE " & NewServerString & "." & Tables1 & " SET " & CutField
                oCommand.CommandText += "LEGAL = '" & NewServerString.ToUpper() & "'," & CutField
                oCommand.CommandText += "PLANT = '" & NewServerString.ToUpper() & "' "
                Try
                    oCommand.ExecuteNonQuery()
                Catch ex As Exception

                End Try
                Tran.Commit()
            Next
            Label4.Text = Now.ToShortTimeString()
            Label4.Refresh()
        Else
            MsgBox("NO TABLES")
            Label4.Text = Now.ToShortTimeString()
            Label4.Refresh()
            Return
        End If

        
    End Sub

    Private Sub Form19_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ServerString = "actiontest"
        NewServerString = "action_dc"
        oConnection.ConnectionString = Module1.OpenOracleConnection(ServerString)
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
        Label3.Text = Now.ToShortTimeString()
        Label3.Refresh()
        Dim oAdapter As New Oracle.ManagedDataAccess.Client.OracleDataAdapter("", oConnection)
        oAdapter.SelectCommand.CommandText = "select zta01 from ZTA_FILE WHERE zta07  = 'T' and zta02 = '" & ServerString & "'"
        oAdapter.SelectCommand.Connection = oConnection
        oAdapter.Fill(DataSet1, "AllTables")
        Dim Allrows As Integer = DataSet1.Tables(0).Rows.Count()
        Label6.Text = Allrows
        Label6.Refresh()
        If Allrows > 0 Then
            For i = 1 To Allrows Step 1
                Label8.Text = i
                Label8.Refresh()
                Dim Tran As Oracle.ManagedDataAccess.Client.OracleTransaction = oConnection.BeginTransaction
                Dim Tables1 As String = DataSet1.Tables(0).Rows(i - 1).Item(0).ToString()
                oCommand.CommandText = "DELETE " & NewServerString & "." & Tables1
                Try
                    oCommand.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    Tran.Rollback()
                    Exit For
                End Try
                oCommand.CommandText = "INSERT INTO " & NewServerString & "." & Tables1 & " (SELECT * FROM "
                oCommand.CommandText += ServerString & "." & Tables1 & ")"
                Try
                    oCommand.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    Tran.Rollback()
                    Exit For
                End Try
                Dim CK As Int16 = Tables1.LastIndexOf("_file")
                Dim CutField As String = Tables1.Substring(0, CK)
                oCommand.CommandText = "UPDATE " & NewServerString & "." & Tables1 & " SET " & CutField
                oCommand.CommandText += "LEGAL = '" & NewServerString.ToUpper() & "'," & CutField
                oCommand.CommandText += "PLANT = '" & NewServerString.ToUpper() & "' "
                Try
                    oCommand.ExecuteNonQuery()
                Catch ex As Exception

                End Try
                Tran.Commit()
            Next
            Label4.Text = Now.ToShortTimeString()
            Label4.Refresh()
        Else
            MsgBox("NO TABLES")
            Label4.Text = Now.ToShortTimeString()
            Label4.Refresh()
            Return
        End If
    End Sub
End Class