Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form35
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim DStartN As Date
    Dim DstartE As Date
    Dim TYear As String = String.Empty
    Dim TMonth As String = String.Empty
    Dim Tmonth1 As String = String.Empty
    Dim CYear As String = String.Empty
    Dim CMonth As String = String.Empty
    Dim LineZ As Integer = 0
    Dim mAdapter1 As New SqlClient.SqlDataAdapter
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form35_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If Now.Month < 10 Then
            TextBox1.Text = Now.Year & "0" & Now.Month
        Else
            TextBox1.Text = Now.Year & Now.Month
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        TYear = Strings.Left(TextBox1.Text, 4)
        TMonth = Strings.Right(TextBox1.Text, 2)
        DStartN = Convert.ToDateTime(TYear & "/" & TMonth & "/01")
        DstartE = DStartN.AddMonths(1).AddDays(-1)
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Sales Amount-WIP COST by customer"
        SaveFileDialog1.DefaultExt = ".xls"
        Dim SON As DialogResult = SaveFileDialog1.ShowDialog()
        If SON = DialogResult.OK Then
            Dim SFN As String = SaveFileDialog1.FileName
            Ws.SaveAs(SFN, XlFileFormat.xlExcel12)
        Else
            MsgBox("没有储存文件", MsgBoxStyle.Critical)
        End If
        xWorkBook.Saved = True
        xWorkBook.Close()
        xExcel.Quit()
        If oConnection.State = ConnectionState.Open Then
            Try
                oConnection.Close()
                Module1.KillExcelProcess(OldExcel)
                MsgBox("Finished")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub ExportToExcel()
        oCommand.CommandText = "CREATE TABLE MES_TEMP (sERPPN nvarchar2(500),sCustomer nvarchar2(50),sQty DEC(10,0))"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
            Return
        End Try
        ' 處理wip
        ' 讀入MES 資料
        Dim mConnectionBuilder As New SqlClient.SqlConnectionStringBuilder
        Dim mConnection As New SqlClient.SqlConnection
        Dim mSQLS1 As New SqlClient.SqlCommand
        'Dim DS As New DataSet()

        mConnectionBuilder.DataSource = "192.168.10.254"
        mConnectionBuilder.InitialCatalog = "ERPSUPPORT"
        mConnectionBuilder.IntegratedSecurity = False
        mConnectionBuilder.UserID = "sa"
        mConnectionBuilder.Password = "p@$$w0rd"
        mConnection.ConnectionString = mConnectionBuilder.ConnectionString

        If mConnection.State <> ConnectionState.Open Then
            mConnection.Open()
            mSQLS1.Connection = mConnection
            mSQLS1.CommandType = CommandType.Text
        End If
        mSQLS1.CommandText = "select * from WIPSaveData where sYear = " & TYear
        mSQLS1.CommandText += " and sMonth = " & TMonth & " and sStation <> '0730' and sERPPN <> ''"
        Dim mSQLReader As SqlClient.SqlDataReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                oCommand.CommandText = "insert into mes_temp (sERPPN,sCustomer,sQty) VALUES ('"
                oCommand.CommandText += mSQLReader.Item("sERPPN").ToString & "','" & mSQLReader.Item("sCustomer").ToString & "'," & mSQLReader.Item("sQty") & ")"
                Try
                    oCommand.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    Return
                End Try
            End While
        End If
        mSQLReader.Close()

        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        AdjustExcelFormat1()
        oCommand.CommandText = "select sc,customername,sum(t1) as t1,sum(t2) as t2, sum(t3) as t3 from ( "
        oCommand.CommandText += "select sc,customername,round(sum(ogb14 * oga24 /  azj041),2) as t1,0 as t2,0 as t3 from oga_file "
        oCommand.CommandText += "join ogb_file on oga01 = ogb01 left join aaa_action on ogb04 = ima01 left join azj_file on azj01 = 'USD' And azj02 = '"
        oCommand.CommandText += TYear & TMonth & "' where oga02 between to_date('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ogapost = 'Y' group by sc,customername "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select sc,customername,0,round(sum(imk09 * ccc23 /azj041),2),0 from imk_file left join aaa_action on imk01 = ima01 left join ccc_file on imk01 = ccc01 and ccc02 = "
        oCommand.CommandText += TYear & " and ccc03 = " & TMonth & " left join azj_file on azj01 = 'USD' And azj02 = '"
        oCommand.CommandText += TYear & TMonth & "' where imk09 > 0 AND imk05 = " & TYear & " and imk06 = " & TMonth & " and imk02 = 'D146103' group by sc,customername "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select sc,to_char(scustomer) as customername,0,0,round(sum(sqty * ccc23 / azj041),2) from mes_temp left join (select distinct customername,sc from aaa_action) bb on bb.customername = sCustomer "
        oCommand.CommandText += "left join ccc_file on sERPPN = ccc01 and ccc02 = " & TYear & " and ccc03 = " & TMonth & " left join azj_file on azj01 = 'USD' and azj02 = '"
        oCommand.CommandText += TYear & TMonth & "' group by sc,scustomer "
        oCommand.CommandText += ") group by sc,customername order by sc"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("sc")
                Ws.Cells(LineZ, 2) = oReader.Item("customername")
                'Ws.Cells(LineZ, 3) = "FG"
                Ws.Cells(LineZ, 3) = oReader.Item("t1")
                Ws.Cells(LineZ, 4) = oReader.Item("t3")
                Ws.Cells(LineZ, 5) = oReader.Item("t2")
                LineZ += 1
            End While
        End If
        oReader.Close()
        oCommand.CommandText = "DROP TABLE mes_temp"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
            Return
        End Try

        'mAdapter1 = New SqlClient.SqlDataAdapter(mSQLS1.CommandText, mConnection)
        'mAdapter1.Fill(DS, "wipdata")
        'If DS.Tables("wipdata").Rows.Count > 0 Then
        '    For i As Integer = 0 To DS.Tables("wipdata").Rows.Count - 1 Step 1
        '        Dim RC As Decimal = GetRC(DS.Tables("wipdata").Rows(i).Item("sERPPN"))
        '        DS.Tables("wipdata").Rows(i).Item("sCost") = RC
        '    Next
        '    If DS.HasChanges() Then
        '        Dim cb As New SqlClient.SqlCommandBuilder(mAdapter1)
        '        mAdapter1.Update(DS.Tables("wipdata"))
        '        DS.Tables("wipdata").AcceptChanges()
        '    End If
        '    ' 讀入 azj_file
        '    If Strings.Len(TMonth) = 1 Then
        '        Tmonth1 = "0" & TMonth
        '    Else
        '        Tmonth1 = TMonth
        '    End If
        '    oCommand.CommandText = "select nvl(sum(azj04),1) from azj_file where azj01 = 'USD' and azj02 = " & Tmonth1
        '    Dim ER As Decimal = oCommand.ExecuteScalar()
        '    If ER <> 0 Then
        '        mSQLS1.CommandText = "select sum(scost * sqty) as t1,sCustomer  from WIPSaveData  where syear = " & TYear & " and smonth = "
        '        mSQLS1.CommandText += TMonth & " group by sCustomer order by sCustomer "
        '        'Dim mSQLReader As SqlClient.SqlDataReader = mSQLS1.ExecuteReader()
        '        If mSQLReader.HasRows() Then
        '            While mSQLReader.Read()
        '                Ws.Cells(LineZ, 2) = mSQLReader.Item("sCustomer")
        '                Ws.Cells(LineZ, 3) = "WIP"
        '                Ws.Cells(LineZ, 5) = mSQLReader.Item("t1") / ER
        '                LineZ += 1
        '            End While
        '        End If
        '        mSQLReader.Close()
        '    End If
        'End If
        mConnection.Close()
        oConnection.Close()
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = TMonth
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 30
        Ws.Cells(1, 1) = "inventory by customer"
        Ws.Cells(1, 2) = "customer"
        'Ws.Cells(1, 3) = "type"
        Ws.Cells(1, 3) = "sales by customer"
        Ws.Cells(1, 4) = "WIP inventory amount by customer"
        Ws.Cells(1, 5) = "FG inventory amount by customer"
        oRng = Ws.Range("C1", "E1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00_ "
        LineZ = 2
    End Sub
    Private Function GetRC(ByVal ccc01 As String)
        oCommand.CommandText = "SELECT NVL(SUM(CCC23),0) FROM CCC_FILE WHERE CCC02 = " & TYear & " AND CCC03 = "
        oCommand.CommandText += TMonth & " AND CCC01 = '" & ccc01 & "'"
        Dim RC As Decimal = oCommand.ExecuteScalar()
        Return RC
    End Function
End Class