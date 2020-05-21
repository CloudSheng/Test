Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel

Public Class Form11
    Dim mConnection As New SqlClient.SqlConnection
    Dim mConnection2 As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim msqlS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                'mConnection2.Open()
                'mSQLS2.Connection = mConnection2
                'mSQLS2.CommandType = CommandType.Text

            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
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
        SaveFileDialog1.FileName = "MFG_ORDER_Report"
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
        If mConnection.State = ConnectionState.Open Then
            Try
                mConnection.Close()
                Module1.KillExcelProcess(OldExcel)
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        LineZ = 2
        mSQLS1.CommandText = "select lot.lot, lot.users + ' - ' +name as users,m.model,substring(m.model+ ' - ' + modelname,0,30) as modelName,m.model+ ' - ' + modelname as modelNameFull, lot.rev, convert(varchar, lot.datetime, 103) + ' ' + CONVERT(varchar, lot.datetime, 108) as datetime, route, qty, status, lot.remark, prefix, case status when 'B' then 'Block' when 'Y' then 'Close' else 'Open' end as stext, sum(case when sn.sn is null then 0 else 1 end) as wip "
        mSQLS1.CommandText += "from lot join users on lot.users = users.id left join sn on lot.lot = sn.lot  join model as m on m.model = lot.model "
        mSQLS1.CommandText += "group by lot.lot, lot.users, name, m.model, lot.rev, lot.datetime, route, qty, status, lot.remark, prefix, status,m.model+ ' - ' + modelname "
        mSQLS1.CommandText += "order by datetime"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = GetERPPN(mSQLReader.Item("model"))
                Ws.Cells(LineZ, 2) = mSQLReader.Item("lot")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("prefix")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("modelName")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("route")
                Ws.Cells(LineZ, 6) = mSQLReader.Item("qty")
                Ws.Cells(LineZ, 7) = mSQLReader.Item("wip")
                Ws.Cells(LineZ, 8) = GETScrapQty(mSQLReader.Item("lot"))
                Ws.Cells(LineZ, 9) = "=F" & LineZ & "-G" & LineZ
                Ws.Cells(LineZ, 10) = "=F" & LineZ & "-G" & LineZ & "-H" & LineZ
                Ws.Cells(LineZ, 11) = mSQLReader.Item("stext")
                Ws.Cells(LineZ, 12) = mSQLReader.Item("remark")
                Ws.Cells(LineZ, 13) = mSQLReader.Item("dateTime")
                Ws.Cells(LineZ, 14) = mSQLReader.Item("users")
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 30
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Cells(1, 1) = "ERP PN"
        Ws.Cells(1, 2) = "MFG Order"
        Ws.Cells(1, 3) = "Prefix"
        Ws.Cells(1, 4) = "Part NO."
        Ws.Cells(1, 5) = "Routing"
        Ws.Cells(1, 6) = "Qty"
        Ws.Cells(1, 7) = "WIP"
        Ws.Cells(1, 8) = "Scrap Qty"
        Ws.Cells(1, 9) = "Orders to owe Qty"
        Ws.Cells(1, 10) = "Actual owe number"
        Ws.Cells(1, 11) = "Status"
        Ws.Cells(1, 12) = "Remark"
        Ws.Cells(1, 13) = "DateTime"
        Ws.Cells(1, 14) = "By"
    End Sub

    Private Sub Form11_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
    Private Function GetERPPN(ByVal model As String)
        mConnection2.ConnectionString = Module1.OpenConnectionOfMes()
        If mConnection2.State <> ConnectionState.Open Then
            Try
                mConnection2.Open()
                msqlS2.Connection = mConnection2
                msqlS2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        msqlS2.CommandText = "SELECT VALUE FROM model_paravalue WHERE parameter = 'ERP PN' AND model = '" & model & "'"
        Dim ERPPN As String = msqlS2.ExecuteScalar()
        mConnection2.Close()
        Return ERPPN
    End Function
    Private Function GETScrapQty(ByVal lot As String)
        mConnection2.ConnectionString = Module1.OpenConnectionOfMes()
        If mConnection2.State <> ConnectionState.Open Then
            Try
                mConnection2.Open()
                msqlS2.Connection = mConnection2
                msqlS2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        msqlS2.CommandText = "select count(*) from scrap where lot = '" & lot & "'"
        Dim ScrapQty As Decimal = msqlS2.ExecuteScalar()
        mConnection2.Close()
        Return ScrapQty
    End Function
End Class