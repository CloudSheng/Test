Public Class Form163
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString = "SELECT * FROM [Sheet1$] WHERE factory = 'DAC'"
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)
            Dim DS As Data.DataSet = New DataSet()
            Try
                ExcelAdapater.Fill(DS, "table1")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            'oCommand.CommandText = "truncate table ACA_Calloff_list"
            'Try
            'oCommand.ExecuteNonQuery()
            'Catch ex As Exception
            'MsgBox(ex.Message())
            'End Try

            For i As Int16 = 0 To DS.Tables("table1").Rows.Count - 1 Step 1
                If String.IsNullOrEmpty(DS.Tables("table1").Rows(i).Item(4).ToString()) Then
                    Continue For
                End If
                For j As Int16 = 10 To 168 Step 1
                    If String.IsNullOrEmpty(DS.Tables("table1").Rows(i).Item(j).ToString()) Then
                        Continue For
                    End If
                    If DS.Tables("table1").Rows(i).Item(j) <= 0 Then
                        Continue For
                    End If
                    Dim Year1 As Decimal = 0
                    Dim Week1 As Decimal = 0
                    Select Case j
                        Case 10 To 62
                            Year1 = 2019
                            Week1 = j - 9
                        Case 63 To 115
                            Year1 = 2020
                            Week1 = j - 62
                        Case 116 To 168
                            Year1 = 2021
                            Week1 = j - 115
                    End Select


                    oCommand.CommandText = "INSERT INTO ACA_Calloff_list VALUES ('" & DS.Tables("table1").Rows(i).Item(1).ToString() & "','" & DS.Tables("table1").Rows(i).Item(4) & "',"
                    oCommand.CommandText += Year1 & "," & Week1 & "," & DS.Tables("table1").Rows(i).Item(j) & ") "
                    Try
                        oCommand.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                Next
            Next
            MsgBox("Done")
        End If
    End Sub

    Private Sub Form163_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
End Class