Public Class Form181
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand3 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim DeleteM As Boolean = False
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim LineS1 As Int16 = 0
    Dim OpendCursor As Int16 = 0
    Dim ReportYES As Boolean = False
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form181_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
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
                oCommand2.Connection = oConnection
                oCommand2.CommandType = CommandType.Text
                oCommand3.Connection = oConnection
                oCommand3.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        CreateTempDB()
        ProcessData()
        ExportToExcel()
    End Sub
    Private Sub CreateTempDB()
        oCommand.CommandText = "DROP TABLE Sales_Temp1"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception

        End Try
        oCommand.CommandText = "Create Table Sales_Temp1 (PN varchar2(40), ima25 varchar2(4))"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
    End Sub
    Private Sub ProcessData()
        ' 先存入初步條件符合的
        oCommand.CommandText = "Insert into Sales_Temp1 select distinct ima01,ima25 from ima_file left join img_file on img01 = ima01 where ima06 = '103' and ima01 like '%66' and imaacti = 'Y' and ima01 not in ( "
        oCommand.CommandText += "Select distinct ogb04 from ogb_file,oga_file where ogb01 = oga01 and oga02 >= to_date('" & Today.AddYears(-1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ogapost = 'Y' ) group by ima01,ima25 having nvl(sum(img10),0) = 0 "

        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

        ' 再排除掉半成品有庫存的
        Label3.Text = "Processing"
        Label3.Refresh()
        oCommand.CommandText = "Select * from Sales_temp1"
        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            While oReader.Read()
                DeleteM = False
                CallBom(oReader.Item(0), oReader.Item(0))
            End While
        End If
        oReader.Close()
        'Label3.Text = "Done"
        'Label3.Refresh()

    End Sub
    Private Sub CallBom(ByVal pn1 As String, ByVal MasterPN As String)
        If DeleteM = True Then
            Exit Sub
        End If
        oCommand2.CommandText = "Select bmb03,nvl(sum(img10),0) from bmb_file left join img_file on bmb03 = img01  where bmb01 = '" & pn1 & "' and bmb05 is null and bmb19 = 2 group by bmb03"
        Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
        oReader2 = oCommand2.ExecuteReader()
        OpendCursor += 1
        If oReader2.HasRows() Then
            While oReader2.Read()
                Dim SS1 As Decimal = oReader2.Item(1)
                If SS1 > 0 Then
                    oCommand3.CommandText = "DELETE Sales_Temp1 WHERE pn = '" & MasterPN & "'"
                    Try
                        oCommand3.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    DeleteM = True
                    Exit While
                Else
                    CallBom(oReader2.Item(0), MasterPN)
                End If
            End While
        End If
        oReader2.Close()
        OpendCursor -= 1
    End Sub
    Private Sub ExportToExcel()
        oCommand.CommandText = "Select * from Sales_Temp1"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            ReportYES = True
            xExcel = New Microsoft.Office.Interop.Excel.Application
            xWorkBook = xExcel.Workbooks.Add()
            Ws = xWorkBook.Sheets(1)
            Ws.Activate()
            Ws.Cells(1, 1) = "ERP PN"
            Ws.Cells(1, 2) = "单位"
            Ws.Cells(1, 3) = "最后一次订单日期"
            Ws.Cells(1, 4) = "最后一次出货时间"
            LineZ = 2
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item(0)
                Ws.Cells(LineZ, 2) = oReader.Item(1)
                oCommand2.CommandText = "Select max(oea02) from oeb_file,oea_file where oeb01 = oea01 and oeb04 = '" & oReader.Item(0) & "'"
                Ws.Cells(LineZ, 3) = oCommand2.ExecuteScalar()
                oCommand2.CommandText = "Select max(oga02) from ogb_file,oga_file where ogb01 = oga01 and ogapost = 'Y' and ogb04 = '" & oReader.Item(0) & "'"
                Ws.Cells(LineZ, 4) = oCommand2.ExecuteScalar()
                LineZ += 1
                Label3.Text = LineZ
                Label3.Refresh()
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If ReportYES = True Then
            SaveExcel()
        End If

    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "待结案料号明细表"
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
End Class