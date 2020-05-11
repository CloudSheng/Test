Public Class Form139
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader

    Private Sub Form139_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommander2.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommander2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
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
            Dim ExcelString = "SELECT 料号,数量 FROM [Sheet3$]"
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
        ' 
        For i As Integer = 0 To DataGridView1.Rows.Count - 1 Step 1
            oCommand.CommandText = "select count(*) from pia_file where pia01 like 'D1501-1912%' and pia19 = 'N' and pia03 in ('D356601') and pia02 = '" & DataGridView1.Rows(i).Cells("料号").Value & "'"
            Dim HData As Decimal = oCommand.ExecuteScalar()
            If HData > 0 Then  ' 有資料就回寫
                oCommand.CommandText = "UPDATE PIA_FILE SET PIA30 = " & DataGridView1.Rows(i).Cells("数量").Value & ",pia901 = 'X' WHERE pia01 like 'D1501-1912%' and pia19 = 'N' and pia03 in ('D356601') and pia02 = '" & DataGridView1.Rows(i).Cells("料号").Value & "'"
                Try
                    oCommand.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            Else   ' 無資料要新增
                oCommand.CommandText = "select Min(pia01) from pia_file where pia01 like 'D1501-1912%' and pia16 = 'Y' and pia02 is null"
                Dim PIa01 As String = oCommand.ExecuteScalar()
                oCommand.CommandText = "select ima25 from ima_file where ima01 = '" & DataGridView1.Rows(i).Cells("料号").Value & "'"
                Dim pia09 As String = oCommand.ExecuteScalar()
                oCommand.CommandText = "UPDATE pia_file SET pia02 = '" & DataGridView1.Rows(i).Cells("料号").Value & "',pia03 ='D356601',pia04 = ' ',pia05 = ' ',pia08 =0,pia09 = '" & pia09 & "',pia10 = 1,pia30 = " & DataGridView1.Rows(i).Cells("数量").Value & ",pia901 = 'X' where pia01 = '" & PIa01 & "'"
                Try
                    oCommand.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
        Next

        oCommand.CommandText = "UPDATE pia_file set pia30 = 0 WHERE pia01 like 'D1501-1912%' and pia19 = 'N' and pia03 in ('D356601') and pia901 is null "
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

        MsgBox("Done")
    End Sub
End Class