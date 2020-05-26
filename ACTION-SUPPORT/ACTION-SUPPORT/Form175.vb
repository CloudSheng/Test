Public Class Form175
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim Ds As New DataSet()
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim Sda As New SqlClient.SqlDataAdapter
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString = "SELECT * FROM [Summary$]"
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)
            Dim DS As Data.DataSet = New DataSet()
            Try
                ExcelAdapater.Fill(DS, "table1")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            Me.DataGridView1.DataSource = DS.Tables(0)
            Me.DataGridView1.Show()

            ' Delete S6
            Dim Tran1 As SqlClient.SqlTransaction = mConnection.BeginTransaction()
            mSQLS1.Transaction = Tran1
            mSQLS1.CommandText = "DELETE IES6"
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message)
                Tran1.Rollback()
                Return
            End Try
            Dim SAVED As Boolean = True
            For i As Int16 = 0 To DS.Tables("table1").Rows.Count - 1 Step 1
                mSQLS1.CommandText = "INSERT INTO IES6 VALUES ('" & DS.Tables("table1").Rows(i).Item(0) & "','" & DS.Tables("table1").Rows(i).Item(1) & "'," & DS.Tables("table1").Rows(i).Item(2) & "," & DS.Tables("table1").Rows(i).Item(3) & ","
                mSQLS1.CommandText += DS.Tables("table1").Rows(i).Item(5) & "," & DS.Tables("table1").Rows(i).Item(6) & "," & DS.Tables("table1").Rows(i).Item(7) & "," & DS.Tables("table1").Rows(i).Item(8) & "," & DS.Tables("table1").Rows(i).Item(10) & ","
                mSQLS1.CommandText += DS.Tables("table1").Rows(i).Item(11) & "," & DS.Tables("table1").Rows(i).Item(12) & "," & DS.Tables("table1").Rows(i).Item(13) & "," & DS.Tables("table1").Rows(i).Item(14) & "," & DS.Tables("table1").Rows(i).Item(15) & ","
                mSQLS1.CommandText += DS.Tables("table1").Rows(i).Item(16) & "," & DS.Tables("table1").Rows(i).Item(17) & "," & DS.Tables("table1").Rows(i).Item(19) & "," & DS.Tables("table1").Rows(i).Item(20) & "," & DS.Tables("table1").Rows(i).Item(21) & ","
                mSQLS1.CommandText += DS.Tables("table1").Rows(i).Item(22) & "," & DS.Tables("table1").Rows(i).Item(24) & "," & DS.Tables("table1").Rows(i).Item(25) & "," & DS.Tables("table1").Rows(i).Item(27) & "," & DS.Tables("table1").Rows(i).Item(29) & ","
                mSQLS1.CommandText += DS.Tables("table1").Rows(i).Item(30) & "," & DS.Tables("table1").Rows(i).Item(31) & "," & DS.Tables("table1").Rows(i).Item(32) & ",'" & DS.Tables("table1").Rows(i).Item(33) & "','" & DS.Tables("table1").Rows(i).Item(34) & "','"
                mSQLS1.CommandText += DS.Tables("table1").Rows(i).Item(35) & "'," & DS.Tables("table1").Rows(i).Item(36) & ",'" & DS.Tables("table1").Rows(i).Item(37) & "'," & DS.Tables("table1").Rows(i).Item(38) & "," & DS.Tables("table1").Rows(i).Item(39) & ","
                mSQLS1.CommandText += DS.Tables("table1").Rows(i).Item(40) & "," & DS.Tables("table1").Rows(i).Item(41) & "," & DS.Tables("table1").Rows(i).Item(42) & "," & DS.Tables("table1").Rows(i).Item(43) & "," & DS.Tables("table1").Rows(i).Item(44) & ","
                mSQLS1.CommandText += DS.Tables("table1").Rows(i).Item(45) & "," & DS.Tables("table1").Rows(i).Item(46) & "," & DS.Tables("table1").Rows(i).Item(47) & ")"
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    Tran1.Rollback()
                    SAVED = False
                    Exit For
                End Try
            Next
            If SAVED = True Then
                Tran1.Commit()
            End If

        End If
    End Sub

    Private Sub Form175_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfERPSUPPORT()

        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
       
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Dim XPath = "C:\temp\IES6_Template.xlsx"
        If Not My.Computer.FileSystem.FileExists(XPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        BackgroundWorker1.RunWorkerAsync()

    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        Dim Name1 As String = String.Empty
        SaveFileDialog1.FileName = "IES3"
        SaveFileDialog1.DefaultExt = ".xlsx"
        Dim SON As DialogResult = SaveFileDialog1.ShowDialog()
        If SON = DialogResult.OK Then
            Dim SFN As String = SaveFileDialog1.FileName
            Ws.SaveAs(SFN, XlFileFormat.xlOpenXMLWorkbook)
        Else
            MsgBox("没有储存文件", MsgBoxStyle.Critical)
        End If
        xWorkBook.Saved = True
        xWorkBook.Close()
        xExcel.Quit()
       
    End Sub
    Private Sub ExportToExcel()
        Dim XPath = "C:\temp\IES6_Template.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        LineZ = 2
        mSQLS1.CommandText = "Select * from IES6"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("PN")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("PName")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("Time1")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("Time2")
                Ws.Cells(LineZ, 6) = mSQLReader.Item("Time3")
                Ws.Cells(LineZ, 7) = mSQLReader.Item("Time4")
                Ws.Cells(LineZ, 8) = mSQLReader.Item("Time5")
                Ws.Cells(LineZ, 9) = mSQLReader.Item("Time6")
                Ws.Cells(LineZ, 11) = mSQLReader.Item("Time7")
                Ws.Cells(LineZ, 12) = mSQLReader.Item("Time8")
                Ws.Cells(LineZ, 13) = mSQLReader.Item("Time9")
                Ws.Cells(LineZ, 14) = mSQLReader.Item("Time10")
                Ws.Cells(LineZ, 15) = mSQLReader.Item("Time11")
                Ws.Cells(LineZ, 16) = mSQLReader.Item("Time12")
                Ws.Cells(LineZ, 17) = mSQLReader.Item("Time13")
                Ws.Cells(LineZ, 18) = mSQLReader.Item("Time14")
                Ws.Cells(LineZ, 20) = mSQLReader.Item("Time15")
                Ws.Cells(LineZ, 21) = mSQLReader.Item("Time16")
                Ws.Cells(LineZ, 22) = mSQLReader.Item("Time17")
                Ws.Cells(LineZ, 23) = mSQLReader.Item("Time18")
                Ws.Cells(LineZ, 25) = mSQLReader.Item("Time19")
                Ws.Cells(LineZ, 26) = mSQLReader.Item("Time20")
                Ws.Cells(LineZ, 28) = mSQLReader.Item("Time21")
                Ws.Cells(LineZ, 30) = mSQLReader.Item("Time22")
                Ws.Cells(LineZ, 31) = mSQLReader.Item("Time23")
                Ws.Cells(LineZ, 32) = mSQLReader.Item("Time24")
                Ws.Cells(LineZ, 32) = mSQLReader.Item("Time25")
                Ws.Cells(LineZ, 33) = mSQLReader.Item("Remark1")
                Ws.Cells(LineZ, 34) = mSQLReader.Item("MName1")
                Ws.Cells(LineZ, 35) = mSQLReader.Item("MName2")
                Ws.Cells(LineZ, 36) = mSQLReader.Item("Time26")
                Ws.Cells(LineZ, 37) = mSQLReader.Item("Remark2")
                Ws.Cells(LineZ, 38) = mSQLReader.Item("Weight1")
                Ws.Cells(LineZ, 39) = mSQLReader.Item("Size1")
                Ws.Cells(LineZ, 40) = mSQLReader.Item("Cav1")
                Ws.Cells(LineZ, 41) = mSQLReader.Item("Time27")
                Ws.Cells(LineZ, 42) = mSQLReader.Item("Time28")
                Ws.Cells(LineZ, 43) = mSQLReader.Item("Time29")
                Ws.Cells(LineZ, 44) = mSQLReader.Item("Time30")
                Ws.Cells(LineZ, 45) = mSQLReader.Item("Time31")
                Ws.Cells(LineZ, 46) = mSQLReader.Item("Time32")
                Ws.Cells(LineZ, 47) = mSQLReader.Item("Time33")
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
    End Sub
    
End Class