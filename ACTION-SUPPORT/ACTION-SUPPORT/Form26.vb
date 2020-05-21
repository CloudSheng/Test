Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form26
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader97 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim CountField As Integer = 0
    Dim LineX As Integer = 0
    'Dim oCommander97 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form26_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString = "SELECT * FROM [011~25SG$],[ACA_customer_list$] where [011~25SG$].customer_nr = [ACA_customer_list$].NO " 'and [011~25SG$].due_date <= " & Today.ToString("yyyy/MM/dd")
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)
            Dim DS As Data.DataSet = New DataSet()
            Try
                ExcelAdapater.Fill(DS, "table1")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            Me.DataGridView1.DataSource = DS.Tables("table1")
            MsgBox(DS.Tables("table1").Columns("due_date").DataType.ToString())
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommand2.Connection = oConnection
                oCommand2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        'AdjustExcelFormat()
        Ws.Activate()
        AdjustExcelFormat()

        For i As Integer = 0 To Me.DataGridView1.Rows.Count - 1 Step 1
            If oConnection.State = ConnectionState.Open Then
                oConnection.Close()
                oConnection.Open()
            End If
            Ws.Cells(LineX, 1) = "訂單編號"
            Ws.Cells(LineX, 2) = Me.DataGridView1.Rows(i).Cells(0).Value
            Ws.Cells(LineX, 3) = "項次"
            Ws.Cells(LineX, 4) = Me.DataGridView1.Rows(i).Cells(1).Value
            Ws.Cells(LineX, 5) = "料號"
            Ws.Cells(LineX, 6) = Me.DataGridView1.Rows(i).Cells(2).Value
            Ws.Cells(LineX, 7) = "數量"
            Ws.Cells(LineX, 8) = Me.DataGridView1.Rows(i).Cells(3).Value
            Ws.Cells(LineX + 1, 2) = Me.DataGridView1.Rows(i).Cells(2).Value
            Dim GetStock As Decimal = GetStock1(Me.DataGridView1.Rows(i).Cells(2).Value)
            Ws.Cells(LineX + 2, 1) = "倉庫量"
            Ws.Cells(LineX + 2, 2) = GetStock
            CountField = 0
            Ws.Cells(LineX + 3, 1) = "預計工單量"
            Ws.Cells(LineX + 3, 2) = Me.DataGridView1.Rows(i).Cells(3).Value - GetStock
            ExtendOrder(Me.DataGridView1.Rows(i).Cells(2).Value, (Me.DataGridView1.Rows(i).Cells(3).Value - GetStock))
            Ws.Cells(LineX + 4, 1) = "不良品倉"
            Ws.Cells(LineX + 5, 1) = "實際開工單量"
            LineX += 6
        Next
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        Ws.Name = "Paint NG"
        LineX = 1
        oCommand2.CommandText = "select IMG01,sum(img10) as t1 from img_file where img02 = 'D356305' and img10 <> 0 group by img01 order by img01"
        oReader = oCommand2.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineX, 1) = oReader.Item("img01")
                Ws.Cells(LineX, 2) = oReader.Item("t1")
                Ws.Cells(LineX, 3) = Strings.Left(oReader.Item("img01").ToString(), 13) & "63"
                LineX += 1
            End While
        End If
        oReader.Close()
        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        Ws.Name = "Glue NG"
        LineX = 1
        oCommand2.CommandText = "select IMG01,sum(img10) as t1 from img_file where img02 = 'D356405' and img10 <> 0 group by img01 order by img01"
        oReader = oCommand2.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineX, 1) = oReader.Item("img01")
                Ws.Cells(LineX, 2) = oReader.Item("t1")
                Ws.Cells(LineX, 3) = Strings.Left(oReader.Item("img01").ToString(), 13) & "64"
                LineX += 1
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "数据"
        Ws.Columns.EntireColumn.ColumnWidth = 20
        Ws.Columns.EntireColumn.NumberFormatLocal = "@"
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        LineX = 1
    End Sub
    Private Function GetStock1(ByVal img01 As String)
        oCommand2.CommandText = "SELECT NVL(SUM(IMG10),0) FROM IMG_FILE WHERE (IMG02 LIKE '%01' or img02 = 'D146103') AND IMG01 = '" & img01 & "'"
        Dim stockvalue As Decimal = oCommand2.ExecuteScalar()
        Return stockvalue
    End Function
    Private Sub ExtendOrder(ByVal bmb01 As String, ByVal restamount As Decimal)
        Dim oCommander97 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander97.Connection = oConnection
        oCommander97.CommandType = CommandType.Text
        oCommander97.CommandText = "select bmb01,bmb03 from bmb_file where bmb01 = '" & bmb01 & "' and bmb05 is NULL and bmb19 = 2 order by bmb03"
        oReader97 = oCommander97.ExecuteReader()
        If oReader97.HasRows() Then
            While oReader97.Read()
                Dim NNN As Decimal = 0
                Ws.Cells(LineX + 1, 3 + CountField) = oReader97.Item("bmb03")
                Dim Stock2 As Decimal = GetStock1(oReader97.Item("bmb03"))
                Ws.Cells(LineX + 2, 3 + CountField) = Stock2
                Dim restamount2 As Decimal = restamount - Stock2
                Ws.Cells(LineX + 3, 3 + CountField) = restamount2
                ' 不良品
                Dim AA As String = Strings.Right(oReader97.Item("bmb03"), 2)
                Select Case AA
                    Case 63
                        NNN = NG(Strings.Left(oReader97.Item("bmb03").ToString(), 13), "D356305")
                    Case 64
                        NNN = NG(Strings.Left(oReader97.Item("bmb03").ToString(), 13), "D356405")
                End Select
                Ws.Cells(LineX + 4, 3 + CountField) = NNN
                restamount2 -= NNN
                Ws.Cells(LineX + 5, 3 + CountField) = restamount2
                CountField += 1
                Call ExtendOrder(oReader97.Item("bmb03"), restamount2)
            End While
            'oReader97.Close()
        End If
        'oReader97.Close()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Work_Order_Report"
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
    Private Function NG(ByVal img01 As String, ByVal img02 As String)
        oCommand.CommandText = "SELECT nvl(sum(img10),0) as t1 from img_file where img01 LIKE '" & img01 & "%' and img02 = '" & img02 & "'"
        Dim NGM As Decimal = oCommand.ExecuteScalar()
        Return NGM
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommand2.Connection = oConnection
                oCommand2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        Dim Adapter As New Oracle.ManagedDataAccess.Client.OracleDataAdapter()
        Dim DS As Data.DataSet = New DataSet()
        oCommand.CommandText = "select imk01,sum(t1) AS T1,sum(t2) as t2,sum(t3) as t3,sum(t1+t2-t3) as t3a "
        oCommand.CommandText += "from ( select imk01,sum(imk09) as t1,0 as t2,0 as t3 from imk_file where imk05 = 2016 and imk06 = 6 AND IMK02 = 'D146103' and imk09 <> 0 group by imk01 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select tlf01,sum(tlf10 * tlf12 * tlf907),0,0 from tlf_file where tlf06 between to_date('2016/07/01','yyyy/mm/dd') "
        oCommand.CommandText += "and to_date('2016/07/30','yyyy/mm/dd') and tlf907 <> 0 and tlf902 = 'D146103' group by tlf01 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select tlf01,0,sum(case when tlf907 = 1 then tlf10 * tlf12 ELSE 0 end ),sum(case when tlf907 = -1 then tlf10 * tlf12 ELSE 0 end) from tlf_file "
        oCommand.CommandText += "where tlf06 between to_date('2016/07/31','yyyy/mm/dd') and to_date('2016/08/06','yyyy/mm/dd') and tlf907 <> 0  and tlf902 = 'D146103' group by tlf01 ) group by imk01 order by imk01"
        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            While oReader.Read()
                oCommand2.CommandText = "select oeb04,t1 from ( Select Round((oeb13 * oea24 / 6.4), 4) as t1,oeb04 from oeb_file,oea_file where oea01 = oeb01 and oeaconf  = 'Y' and oeb04 = '"
                oCommand2.CommandText += oReader.Item("imk01") & "'  and oea02 < to_date('2016/08/08','yyyy/mm/dd')  order by oea02 desc ) where rownum <= 1"
                Adapter.SelectCommand = oCommand2
                Try
                    Adapter.Fill(DS, "table1")
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End While
        End If
        oReader.Close()
        Me.DataGridView1.DataSource = DS.Tables("table1")
    End Sub
End Class