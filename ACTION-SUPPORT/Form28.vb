Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form28
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim tStation1 As String
    Dim ptime As String = String.Empty
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form28_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        ptime = Today.AddDays(-1).ToString("yyyy/MM/dd")
        ptime = ptime & " 08:00:00"
        Me.DateTimePicker1.Value = Convert.ToDateTime(ptime)
        Me.DateTimePicker2.Value = Convert.ToDateTime(ptime).AddDays(1).AddSeconds(-1)
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
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
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString = "SELECT ERPPN,Productname,QTY FROM [sheet1$]"
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)
            Dim DS As Data.DataSet = New DataSet()
            Try
                ExcelAdapater.Fill(DS, "table1")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            Me.DataGridView1.DataSource = DS.Tables("table1")
            CheckStatus()
        End If
    End Sub
    Private Sub CheckStatus()
        If Me.DataGridView1.Rows.Count > 0 Then
            Me.Button2.Enabled = True
            Me.Label3.Text = "已读取来源档"
        Else
            Me.Button2.Enabled = False
            Me.Label3.Text = "未读取来源档"
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If Me.DataGridView1.Rows.Count <= 0 Then
            MsgBox("资料有误")
            Return
        End If
        Dim xPath As String = "C:\temp\Production plan to achieve statistical report.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
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
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If

        TimeS1 = DateTimePicker1.Value
        TimeS2 = DateTimePicker2.Value
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Production plan to achieve statistical report"
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
        SaveExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        Dim xPath As String = "C:\temp\Production plan to achieve statistical report.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Cells(2, 7) = TimeS1.ToString("yyyy/MM/dd")
        Ws.Cells(2, 10) = TimeS2.ToString("yyyy/MM/dd")
        LineZ = 6
        For i As Integer = 0 To DataGridView1.Rows.Count - 1 Step 1
            If Not IsDBNull(DataGridView1.Rows(i).Cells("Productname").Value) Then
                Ws.Cells(LineZ, 1) = i + 1
                Dim ERPPN As String = DataGridView1.Rows(i).Cells("ERPPN").Value
                Dim l_model As String = DataGridView1.Rows(i).Cells("Productname").Value
                Ws.Cells(LineZ, 2) = ERPPN
                'Ws.Cells(LineZ, 3) = DataGridView1.Rows(i).Cells("Product").Value
                Ws.Cells(LineZ, 4) = DataGridView1.Rows(i).Cells("QTY").Value
                If Not String.IsNullOrEmpty(l_model) Then
                    'Dim ima58 As Decimal = GetStandardIE(ERPPN)
                    'Ws.Cells(LineZ, 6) = ima58
                    'Ws.Cells(LineZ, 4) = DataGridView1.Rows(i).Cells("QTY").Value * ima58
                    'Dim l_model As String = GetModelName(ERPPN)
                    Ws.Cells(LineZ, 3) = l_model
                    Dim abNormal As Integer = GetTracking_DUP_Data(ERPPN, l_model)
                    Ws.Cells(LineZ, 10) = abNormal
                    Dim NormalOutPut As Integer = GetTrackingData(ERPPN, l_model)
                    Ws.Cells(LineZ, 9) = NormalOutPut
                    Ws.Cells(LineZ, 6) = GetTrackingData_ALL(ERPPN, l_model)
                    'Ws.Cells(LineZ, 6) = (abNormal + NormalOutPut)

                    'Ws.Cells(LineZ, 10) = Decimal.Round(ima58 * (abNormal + NormalOutPut) / 60, 1)
                    'If (abNormal + NormalOutPut) > DataGridView1.Rows(i).Cells("QTY").Value Then
                    '    Ws.Cells(LineZ, 11) = DataGridView1.Rows(i).Cells("QTY").Value
                    '    Ws.Cells(LineZ, 12) = 0
                    'Else
                    '    Ws.Cells(LineZ, 11) = (abNormal + NormalOutPut)
                    '    Ws.Cells(LineZ, 12) = (abNormal + NormalOutPut) - DataGridView1.Rows(i).Cells("QTY").Value
                    'End If
                    Dim FailCount As Decimal = GetFailureData(ERPPN, l_model)
                    Ws.Cells(LineZ, 7) = FailCount
                    Dim ScrapCount As Decimal = GetScrapData(ERPPN, l_model)
                    Ws.Cells(LineZ, 8) = ScrapCount
                End If
            End If
            LineZ += 1
        Next
        'SUMALL()

    End Sub
    Private Function GetStandardIE(ByVal ERPPN As String)
        oCommand.CommandText = "SELECT nvl(ima58,0) FROM ima_file WHERE ima01 = '" & ERPPN & "'"
        Dim l_ima58 As Decimal = 0
        Try
            l_ima58 = oCommand.ExecuteScalar()
        Catch ex As Exception
            MsgBox(ex.Message())
            l_ima58 = 0
        End Try
        Return l_ima58
    End Function
    Private Function GetModelName(ByVal erppn As String)
        mSQLS1.CommandText = "SELECT distinct model FROM model_station_paravalue WHERE cf01 = '" & erppn & "'"
        Dim l_model As String = String.Empty
        Try
            l_model = mSQLS1.ExecuteScalar()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        Return l_model
    End Function
    Private Function GetTrackingData(ByVal ERPPN As String, ByVal l_model As String)
        Select Case Strings.Right(ERPPN, 2)
            Case 66
                tStation1 = "'0675'"
            Case 65
                tStation1 = "'0640'"
            Case 63
                tStation1 = "'0590'"
            Case 61
                tStation1 = "'0410'"
            Case 64
                tStation1 = "'0480'"
            Case 36
                tStation1 = "'0380'"
            Case 35
                tStation1 = "'0330','0331'"
            Case 32
                tStation1 = "'0150','0151'"
            Case 31
                tStation1 = "'0110','0111'"
        End Select
        mSQLS1.CommandText = "select sum(t1) as t1 from ( "
        mSQLS1.CommandText += "SELECT count(sn) as t1 from tracking right join lot on tracking.lot =lot.lot "
        mSQLS1.CommandText += "where lot.model = '" & l_model & "' AND timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and result = 'P' and station in (" & tStation1 & ") "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "SELECT count(sn) from scrap_tracking right join lot on scrap_tracking.lot =lot.lot "
        mSQLS1.CommandText += "where lot.model = '" & l_model & "' AND timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and result = 'P' and station in (" & tStation1 & ") ) AS ab "
        Dim NormalOutput As Integer = 0
        Try
            NormalOutput = mSQLS1.ExecuteScalar()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        Return NormalOutput
    End Function
    Private Function GetTracking_DUP_Data(ByVal ERPPN As String, ByVal l_model As String)
        Select Case Strings.Right(ERPPN, 2)
            Case 66
                tStation1 = "'0675'"
            Case 65
                tStation1 = "'0640'"
            Case 63
                tStation1 = "'0590'"
            Case 61
                tStation1 = "'0410'"
            Case 64
                tStation1 = "'0480'"
            Case 36
                tStation1 = "'0380'"
            Case 35
                tStation1 = "'0330','0331'"
            Case 32
                tStation1 = "'0150','0151'"
            Case 31
                tStation1 = "'0110','0111'"
        End Select
        mSQLS1.CommandText = "SELECT count(sn) from tracking_dup right join lot on tracking_dup.lot =lot.lot and lot.model = '"
        mSQLS1.CommandText += l_model & "' where timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and result = 'P' and station in (" & tStation1 & ")"
        Dim NormalOutput As Integer = 0
        Try
            NormalOutput = mSQLS1.ExecuteScalar()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        Return NormalOutput
    End Function
    Private Sub SUMALL()
        Ws.Cells(5, 3) = "=SUM(C8:C" & Me.DataGridView1.Rows.Count + 7 & ")"
        Ws.Cells(5, 4) = "=SUM(D8:D" & Me.DataGridView1.Rows.Count + 7 & ")"
        Ws.Cells(5, 7) = "=SUM(G8:G" & Me.DataGridView1.Rows.Count + 7 & ")"
        Ws.Cells(5, 8) = "=SUM(H8:H" & Me.DataGridView1.Rows.Count + 7 & ")"
        Ws.Cells(5, 9) = "=SUM(I8:I" & Me.DataGridView1.Rows.Count + 7 & ")"
        Ws.Cells(5, 10) = "=SUM(J8:J" & Me.DataGridView1.Rows.Count + 7 & ")"
        Ws.Cells(5, 11) = "=SUM(K8:K" & Me.DataGridView1.Rows.Count + 7 & ")"
        Ws.Cells(5, 12) = "=SUM(L8:L" & Me.DataGridView1.Rows.Count + 7 & ")"

    End Sub
    Private Function GetFailureData(ByVal ERPPN As String, ByVal l_model As String)
        Select Case Strings.Right(ERPPN, 2)
            Case 66
                tStation1 = "'0670'"
            Case 65
                tStation1 = "'0640'"
            Case 63
                tStation1 = "'0590','0591','0592'"
            Case 61
                tStation1 = "'0430'"
            Case 64
                tStation1 = "'0490','0491','0620'"
            Case 36
                tStation1 = "'0380','0530'"
            Case 35
                tStation1 = "'0330','0331'"
            Case 32
                tStation1 = "'0172'"
                'Case 31
                '   tStation1 = "'0112','0113'"
        End Select
        mSQLS1.CommandText = "select sum(t1) as t1 from ( "
        mSQLS1.CommandText += "select count(*) as t1 from failure left join lot on failure.lot = lot.lot "
        mSQLS1.CommandText += "WHERE lot.model = '" & l_model & "' AND failtime between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and failstation in (" & tStation1 & ") "
        mSQLS1.CommandText += "and rework <> 'SCRP' AND rework <> 'BLCK' "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select count(*) from scrap_failure left join lot on scrap_failure.lot = lot.lot "
        mSQLS1.CommandText += "WHERE lot.model = '" & l_model & "' AND failtime between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and failstation in (" & tStation1 & ") "
        mSQLS1.CommandText += "and rework <> 'SCRP' AND rework <> 'BLCK' ) AS AB"
        
        Dim FailureCount As Integer = 0
        Try
            FailureCount = mSQLS1.ExecuteScalar()
        Catch ex As Exception
            MsgBox(ex.Message())
                End Try
        Return FailureCount
    End Function
    Private Function GetScrapData(ByVal ERPPN As String, ByVal l_model As String)
        Select Case Strings.Right(ERPPN, 2)
            Case 66
                tStation1 = "'0670'"
            Case 65
                tStation1 = "'0640'"
            Case 63
                tStation1 = "'0590','0591','0592'"
            Case 61
                tStation1 = "'0430'"
            Case 64
                tStation1 = "'0490','0491','0620'"
            Case 36
                tStation1 = "'0380','0530'"
            Case 35
                tStation1 = "'0330','0331'"
            Case 32
                tStation1 = "'0172'"
                'Case 31
                '   tStation1 = "'0112','0113'"
        End Select
        mSQLS1.CommandText = "select count(scrap.sn)  from scrap left join scrap_sn on scrap.sn = scrap_sn.sn left join lot on scrap.lot = lot.lot where lot.model = '"
        mSQLS1.CommandText += l_model & "' AND scrap.datetime between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and updatedstation in (" & tStation1 & ") "

        Dim ScrapCount As Integer = 0
        Try
            ScrapCount = mSQLS1.ExecuteScalar()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        Return ScrapCount
    End Function
    Private Function GetTrackingData_ALL(ByVal ERPPN As String, ByVal l_model As String)
        Select Case Strings.Right(ERPPN, 2)
            Case 66
                tStation1 = "'0675'"
            Case 65
                tStation1 = "'0640'"
            Case 63
                tStation1 = "'0590'"
            Case 61
                tStation1 = "'0410'"
            Case 64
                tStation1 = "'0490','0491','0620'"
            Case 36
                tStation1 = "'0380','0530'"
            Case 35
                tStation1 = "'0330','0331'"
            Case 32
                tStation1 = "'0150','0151'"
            Case 31
                tStation1 = "'0110','0111'"
        End Select
        mSQLS1.CommandText = "select sum(t1) as t1 from ( "
        mSQLS1.CommandText += "SELECT count(sn) as t1 from tracking right join lot on tracking.lot =lot.lot "
        mSQLS1.CommandText += "where lot.model = '" & l_model & "' AND timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and result = 'P' and station in (" & tStation1 & ") "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "SELECT count(sn) from tracking_dup right join lot on tracking_dup.lot =lot.lot and lot.model = '"
        mSQLS1.CommandText += l_model & "' where timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and result = 'P' and station in (" & tStation1 & ") "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "SELECT count(sn) from scrap_tracking right join lot on scrap_tracking.lot =lot.lot "
        mSQLS1.CommandText += "where lot.model = '" & l_model & "' AND timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and result = 'P' and station in (" & tStation1 & ") ) AS ab "
        Dim NormalOutput As Integer = 0
        Try
            NormalOutput = mSQLS1.ExecuteScalar()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        Return NormalOutput
    End Function
End Class