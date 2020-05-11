Public Class Form118
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand

    Private Sub Form118_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
            Dim ExcelString = "SELECT 單號,單號項次,交貨項次,數量,交貨日期 FROM [Sheet1$] "
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)
            Dim DS As Data.DataSet = New DataSet()
            Try
                ExcelAdapater.Fill(DS, "table1")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            Me.DataGridView1.DataSource = DS.Tables("table1")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        For i As Integer = 0 To DataGridView1.RowCount - 1 Step 1
            oCommand.CommandText = "DELETE tc_cif_file WHERE tc_cif_01 = '" & DataGridView1.Rows(i).Cells(0).Value
            oCommand.CommandText += "' AND tc_cif_02 = " & DataGridView1.Rows(i).Cells(1).Value
            Try
                oCommand.ExecuteNonQuery()
            Catch ex As Exception

            End Try
        Next

        For i As Integer = 0 To DataGridView1.Rows.Count - 1 Step 1
            'oCommand.CommandText = "DELETE tc_cif_file WHERE tc_cif_01 = '" & DataGridView1.Rows(i).Cells(0).Value
            'oCommand.CommandText += "' AND tc_cif_02 = " & DataGridView1.Rows(i).Cells(1).Value
            'oCommand.CommandText += " AND tc_cif_03 = " & DataGridView1.Rows(i).Cells(2).Value
            'Try
            'oCommand.ExecuteNonQuery()
            'Catch ex As Exception

            'End Try
            oCommand.CommandText = "INSERT INTO tc_cif_file (tc_cif_01,tc_cif_02,tc_cif_03,tc_cif_04,tc_cif_05,tc_cifplant) VALUES ('"
            oCommand.CommandText += DataGridView1.Rows(i).Cells(0).Value & "'," & DataGridView1.Rows(i).Cells(1).Value & "," & DataGridView1.Rows(i).Cells(2).Value
            oCommand.CommandText += "," & DataGridView1.Rows(i).Cells(3).Value & ",to_date('" & DataGridView1.Rows(i).Cells(4).Value & "','yy/MM/dd'),'ACTIONTEST')"
            Try
                oCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        Next
        MsgBox("FINISHED")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        For i As Integer = 0 To DataGridView1.RowCount - 1 Step 1
            oCommand.CommandText = "DELETE tc_cig_file WHERE tc_cig_01 = '" & DataGridView1.Rows(i).Cells(0).Value
            oCommand.CommandText += "' AND tc_cig_02 = " & DataGridView1.Rows(i).Cells(1).Value
            Try
                oCommand.ExecuteNonQuery()
            Catch ex As Exception

            End Try
        Next

        For i As Integer = 0 To DataGridView1.Rows.Count - 1 Step 1
            'oCommand.CommandText = "DELETE tc_cig_file WHERE tc_cig_01 = '" & DataGridView1.Rows(i).Cells(0).Value
            'oCommand.CommandText += "' AND tc_cig_02 = " & DataGridView1.Rows(i).Cells(1).Value
            'oCommand.CommandText += " AND tc_cig_03 = " & DataGridView1.Rows(i).Cells(2).Value
            'Try
            'oCommand.ExecuteNonQuery()
            'Catch ex As Exception

            'End Try
            oCommand.CommandText = "INSERT INTO tc_cig_file (tc_cig_01,tc_cig_02,tc_cig_03,tc_cig_04,tc_cig_05,tc_cigplant) VALUES ('"
            oCommand.CommandText += DataGridView1.Rows(i).Cells(0).Value & "'," & DataGridView1.Rows(i).Cells(1).Value & "," & DataGridView1.Rows(i).Cells(2).Value
            oCommand.CommandText += "," & DataGridView1.Rows(i).Cells(3).Value & ",to_date('" & DataGridView1.Rows(i).Cells(4).Value & "','yy/MM/dd'),'ACTIONTEST')"
            Try
                oCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        Next
        MsgBox("FINISHED")
    End Sub
End Class