Public Class Form3
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim l_station As String = String.Empty
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString = "SELECT SN FROM [sheet1$]"
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
        If Me.DataGridView1.Rows.Count = 0 Then
            MsgBox("无资料可处理")
            Return
        End If
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If Me.RadioButton1.Checked Then
            l_station = "0670"
        Else
            l_station = "0669"
        End If
        For i As Integer = 0 To DataGridView1.Rows.Count - 1 Step 1
            mSQLS1.CommandText = "SELECT COUNT(*) FROM SN where sn.updatedstation = '0665' and sn = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "'"
            Dim CheckF As Decimal = mSQLS1.ExecuteScalar()
            If CheckF = 0 Then
                MsgBox(DataGridView1.Rows(i).Cells(0).Value & " not at station 0665")
                Continue For
            End If
            mSQLS1.CommandText = "UPDATE sn SET currentstation = '" & l_station & "', updatedstation = '" & l_station & "', "
            mSQLS1.CommandText += "remark = remark + '" & Today.ToString("MM/dd") & " Concession' where sn ='"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "'"
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        Next
        MsgBox("处理完毕")
    End Sub
End Class