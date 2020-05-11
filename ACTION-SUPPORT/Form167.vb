Public Class Form167
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Private Sub Form167_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString As New OleDb.OleDbCommand
            ExcelString.CommandText = "SELECT * FROM [Sample$]"
            ExcelString.Connection = Excelconn
            Dim ExcelDataReader As OleDb.OleDbDataReader = ExcelString.ExecuteReader
            If ExcelDataReader.HasRows() Then
                Dim Tran1 As Oracle.ManagedDataAccess.Client.OracleTransaction = oConnection.BeginTransaction()
                While ExcelDataReader.Read()
                    oCommand.CommandText = "INSERT INTO abb_file VALUES ("
                    For i As Int16 = 0 To ExcelDataReader.FieldCount - 1 Step 1
                        'MsgBox(ExcelDataReader.Item(i).GetType().ToString())
                        If ExcelDataReader.Item(i).GetType().ToString() = "System.String" Then
                            oCommand.CommandText += "'" & ExcelDataReader.Item(i) & "',"
                        End If
                        If ExcelDataReader.Item(i).GetType().ToString() = "System.Double" Then
                            oCommand.CommandText += ExcelDataReader.Item(i) & ","
                        End If
                        If ExcelDataReader.Item(i).GetType().ToString() = "System.DBNull" Then
                            oCommand.CommandText += "NULL,"
                        End If
                        If ExcelDataReader.Item(i).GetType().ToString() = "System.DateTime" Then
                            oCommand.CommandText += "to_date('" & ExcelDataReader.Item(i) & "','yyyy/MM/dd'),"
                        End If
                    Next
                    oCommand.CommandText = oCommand.CommandText.Remove(oCommand.CommandText.Length - 1)
                    oCommand.CommandText += ") "
                    Try
                        oCommand.ExecuteNonQuery()
                    Catch ex As Exception
                        Tran1.Rollback()
                        MsgBox(ex.Message())
                        Exit While
                    End Try
                End While
                Tran1.Commit()
            End If
            ExcelDataReader.Close()
            MsgBox("Done")
        End If
    End Sub
End Class