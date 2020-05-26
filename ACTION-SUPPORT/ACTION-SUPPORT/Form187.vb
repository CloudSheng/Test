Public Class Form187
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
            Dim ExcelString = "SELECT * FROM [S4_Template$]"
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
            Dim SAVED As Boolean = True
            For i As Int16 = 0 To DS.Tables("table1").Rows.Count - 1 Step 1
                mSQLS1.CommandText = "INSERT INTO IES4 VALUES ('" & DS.Tables("table1").Rows(i).Item(0) & "','" & DS.Tables("table1").Rows(i).Item(1) & "','" & DS.Tables("table1").Rows(i).Item(2) & "','" & DS.Tables("table1").Rows(i).Item(3) & "','"
                mSQLS1.CommandText += DS.Tables("table1").Rows(i).Item(4) & "','" & DS.Tables("table1").Rows(i).Item(5) & "','" & DS.Tables("table1").Rows(i).Item(6) & "','" & DS.Tables("table1").Rows(i).Item(7) & "')"

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

    Private Sub Form187_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Dim XPath = "C:\temp\IES4_Template.xlsx"
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
        SaveFileDialog1.FileName = "IES4"
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
        Dim XPath = "C:\temp\IES4_Template.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        LineZ = 2
        mSQLS1.CommandText = "Select * from IES4"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                For i As Int16 = 0 To mSQLReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = mSQLReader.Item(i)
                Next
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
    End Sub
End Class