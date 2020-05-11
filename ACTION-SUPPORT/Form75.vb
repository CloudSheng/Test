Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form75
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim ExcelPath As String = String.Empty
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim ArrayX1 As String() = {"", "", ""}
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form75_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        
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
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            ExcelPath = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString = "SELECT SN号 FROM [报表1$]"
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)
            Dim DS As Data.DataSet = New DataSet()
            Try
                ExcelAdapater.Fill(DS, "table1")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            Me.DataGridView1.DataSource = DS.Tables("table1")
            Me.DataGridView1.AutoResizeColumn(0)
            CheckStatus()
        End If
        If DataGridView1.Rows.Count > 0 Then
            If mConnection.State <> ConnectionState.Open Then
                Try
                    mConnection.Open()
                    mSQLS1.Connection = mConnection
                    mSQLS1.CommandType = CommandType.Text
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
            'BackgroundWorker1.RunWorkerAsync()
            ExportToExcel()
        End If
    End Sub
    Private Sub CheckStatus()
        If Me.DataGridView1.Rows.Count > 0 Then
            Me.Label3.Text = "已读取来源档"
        Else
            Me.Label3.Text = "未读取来源档"
        End If
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Label3.Text = "导出Excel中"
        ExportToExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "MES样品进度报表"
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
        If mConnection.State = ConnectionState.Open Then
            Try
                mConnection.Close()
                Module1.KillExcelProcess(OldExcel)
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Me.Label3.Text = "导出完毕"
        SaveExcel()
    End Sub
    Private Sub ExportToExcel()
        Label3.Text = "导出Excel中"
        xExcel = New Microsoft.Office.Interop.Excel.Application
        If Not My.Computer.FileSystem.FileExists(ExcelPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(ExcelPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        LineZ = 2
        Try


            For i As Integer = 0 To DataGridView1.Rows.Count - 1 Step 1
                If Not IsDBNull(DataGridView1.Rows(i).Cells("SN号").Value) Then
                    Ws.Cells(LineZ, 1) = Now.ToString("yyyy/MM/dd HH:mm:ss")
                    ArrayX1 = GetModel(DataGridView1.Rows(i).Cells("SN号").Value)
                    Ws.Cells(LineZ, 4) = ArrayX1(0).ToString()
                    Ws.Cells(LineZ, 5) = ArrayX1(1).ToString()
                    Dim Date1 As Object = Get0330Finish(DataGridView1.Rows(i).Cells("SN号").Value)
                    Ws.Cells(LineZ, 6) = Date1
                    Select Case ArrayX1(2).ToString()
                        Case "0080", "0110", "0111"
                            Ws.Cells(LineZ, 7) = 1
                        Case "0112", "0113"
                            Ws.Cells(LineZ, 8) = 1
                        Case "0150", "0151", "0165", "0170"
                            Ws.Cells(LineZ, 9) = 1
                        Case "0180", "0190", "0193", "0195", "0200", "0210", "0215", "0220", "0223", "0225", "0230", "0231", "0240", "0250", "0260", "0280", "0300", "0315", "0316", "0320", "0321", "0325", "0326"
                            Ws.Cells(LineZ, 10) = 1
                        Case "0330", "0331"
                            Ws.Cells(LineZ, 11) = 1
                        Case "0390", "0395"
                            Ws.Cells(LineZ, 12) = 1
                        Case "0335", "0340", "0350", "0360", "0370", "0493", "0495", "0500", "0510", "0520"
                            Ws.Cells(LineZ, 13) = 1
                        Case "0380", "0530"
                            Ws.Cells(LineZ, 14) = 1
                        Case "0400"
                            Ws.Cells(LineZ, 15) = 1
                        Case "0478", "0480"
                            Ws.Cells(LineZ, 16) = 1
                        Case "0605", "0610", "0623"
                            Ws.Cells(LineZ, 17) = 1
                        Case "0490", "0492"
                            Ws.Cells(LineZ, 18) = 1
                        Case "0620", "0627"
                            Ws.Cells(LineZ, 19) = 1
                        Case "0405", "0410"
                            Ws.Cells(LineZ, 20) = 1
                        Case "0415", "0417"
                            Ws.Cells(LineZ, 21) = 1
                        Case "0435", "0440"
                            Ws.Cells(LineZ, 22) = 1
                        Case "0460"
                            Ws.Cells(LineZ, 23) = 1
                        Case "0540"
                            Ws.Cells(LineZ, 24) = 1
                        Case "0570"
                            Ws.Cells(LineZ, 25) = 1
                        Case "0583"
                            Ws.Cells(LineZ, 26) = 1
                        Case "0418", "0420", "0455"
                            Ws.Cells(LineZ, 27) = 1
                        Case "0430", "0445", "0450", "0567"
                            Ws.Cells(LineZ, 28) = 1
                        Case "0465", "0470", "0475"
                            Ws.Cells(LineZ, 29) = 1
                        Case "0545", "0550"
                            Ws.Cells(LineZ, 30) = 1
                        Case "0560", "0563"
                            Ws.Cells(LineZ, 31) = 1
                        Case "0575", "0580"
                            Ws.Cells(LineZ, 32) = 1
                        Case "0584", "0585"
                            Ws.Cells(LineZ, 33) = 1
                        Case "0587", "0590", "0591", "0592", "0595", "0600"
                            Ws.Cells(LineZ, 34) = 1
                        Case "0629", "0630", "0633", "0635"
                            Ws.Cells(LineZ, 35) = 1
                        Case "0640", "0642", "0645", "0657"
                            Ws.Cells(LineZ, 36) = 1
                        Case "0650", "0652", "0658", "0659", "0665", "0670", "0673"
                            Ws.Cells(LineZ, 37) = 1
                        Case "0675", "0680", "0690"
                            Ws.Cells(LineZ, 38) = 1
                        Case "0720", "0730"
                            Ws.Cells(LineZ, 39) = 1
                        Case "9999"
                            Ws.Cells(LineZ, 41) = 1
                    End Select
                    Ws.Cells(LineZ, 42) = GetLastStation(DataGridView1.Rows(i).Cells("SN号").Value)
                End If
                LineZ += 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        Me.Label3.Text = "导出完毕"
        SaveExcel()
    End Sub
    Private Function GetModel(ByVal sn As String)
        mSQLS1.CommandText = "select model.model,model.modelname,(case when topreworkstation is null or topreworkstation = '' then updatedstation else topreworkstation end) as c1 from sn left join lot on sn.lot = lot.lot left join model on lot.model = model.model where sn  = '" & sn & "'"
        mSQLReader = mSQLS1.ExecuteReader()
        Dim ArrayX2 As String() = {"", "", ""}
        If mSQLReader.HasRows() Then
            mSQLReader.Read()
            For k As Integer = 0 To mSQLReader.FieldCount - 1 Step 1
                ArrayX2.SetValue(mSQLReader.Item(k), k)
            Next
        End If
        mSQLReader.Close()
        Return ArrayX2

    End Function
    Private Function Get0330Finish(ByVal sn As String)
        mSQLS1.CommandText = "select timeout from tracking where sn = '" & sn & "' and station in ('0330','0331')"
        Dim DateS As Object = mSQLS1.ExecuteScalar()
        Return DateS
    End Function
    Private Function GetLastStation(ByVal sn As String)
        Dim LS As String = String.Empty
        mSQLS1.CommandText = "select updatedstation from scrap_sn where sn = '" & sn & "'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            mSQLReader.Read()
            LS = "已报废于" & mSQLReader.Item(0).ToString()
        End If
        mSQLReader.Close()
        Return LS
    End Function
End Class