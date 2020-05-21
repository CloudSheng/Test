Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel
Public Class Form64
    Dim mConnection As New SqlClient.SqlConnection
    Dim mConnection2 As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
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
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
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
        BackgroundWorker1.RunWorkerAsync()
        'ExportToExcel()
        'SaveExcel()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "MES_ERP_quantity_Report"
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

    Private Sub Form64_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        mSQLS1.CommandText = "select sum(qty) as qty,sum(t1) as wip,sum(t2) as scrap,cf01,aa.model, model_paravalue.value from ("
        mSQLS1.CommandText += "select sum(qty) as qty,sum(t1) as t1,sum(t2) as t2,model from ("
        mSQLS1.CommandText += "select lot.lot,lot.qty,lot.model,count(sn.sn)as t1,0 as t2 from lot left join sn on lot.lot =sn.lot where status = 'N' group by lot.lot, lot.model,lot.qty "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select lot.lot,0,lot.model,0 as t1,count(scrap.sn) as t2 from lot left join scrap on lot.lot =scrap.lot where status = 'N' group by lot.lot, lot.model "
        mSQLS1.CommandText += ") AS ab group by model ) aS aa left join model on aa.model = model.model "
        mSQLS1.CommandText += "right join routing on model.default_route = route and station in ('0110','0111') "
        mSQLS1.CommandText += "left join model_station_paravalue on aa.model = model_station_paravalue.model and model_station_paravalue.station = routing.station "
        mSQLS1.CommandText += "left join model_paravalue on aa.model = model_paravalue.model and model_paravalue.parameter = 'ERP PN' "
        mSQLS1.CommandText += " where cf01 is not null group by aa.model,cf01,value order by aa.model"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("value")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("cf01")
                Ws.Cells(LineZ, 6) = mSQLReader.Item("qty") - mSQLReader.Item("wip")
                Ws.Cells(LineZ, 7) = mSQLReader.Item("scrap")
                SetERPDate(mSQLReader.Item("cf01"))
                Ws.Cells(LineZ, 8) = "=E" & LineZ & "-F" & LineZ

                Dim TotalWip As Decimal = 0
                TotalWip = GetTotalWIP(mSQLReader.Item("model"))
                Ws.Cells(LineZ, 9) = TotalWip

                Dim EWip As Decimal = 0
                EWip = GetEWIP(mSQLReader.Item("model"))
                Ws.Cells(LineZ, 10) = EWip

                Ws.Cells(LineZ, 12) = "=F" & LineZ & "+I" & LineZ & "-K" & LineZ
                Ws.Cells(LineZ, 14) = "=E" & LineZ & "+J" & LineZ & "-K" & LineZ

                Label1.Text = LineZ
                LineZ += 1
            End While
        End If
        mSQLReader.Close()

        mSQLS1.CommandText = "SELECT MODEL.MODEL,model_paravalue.value  FROM MODEL LEFT JOIN model_paravalue ON model.model = model_paravalue.model and model_paravalue.parameter = 'ERP PN' WHERE MODEL.MODEL NOT IN ( SELECT DISTINCT MODEL FROM MODEL ,ROUTING WHERE model.default_route = route AND station in ('0110','0111') )"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("value")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("model")
                LineZ += 1
            End While
        End If
        
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 30
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Cells(1, 1) = "ERP PN"
        Ws.Cells(1, 2) = "Product name"
        Ws.Cells(1, 3) = "主件料号"
        Ws.Cells(1, 4) = "品名"
        'Ws.Cells(1, 3) = "规格"
        Ws.Cells(1, 5) = "工单数量"
        Ws.Cells(1, 6) = "MES生产制令数量"
        Ws.Cells(1, 7) = "MES生产制令报废数量"
        Ws.Cells(1, 8) = "差异数量"
        Ws.Cells(1, 9) = "总WIP数量"
        Ws.Cells(1, 10) = "MES已裁纱WIP"
        Ws.Cells(1, 11) = "订单余量"
        Ws.Cells(1, 12) = "MES制令差异"
        Ws.Cells(1, 13) = "MES处理结果"
        Ws.Cells(1, 14) = "ERP工单差异"
        Ws.Cells(1, 15) = "MES处理结果"
        LineZ = 2
    End Sub
    Private Sub SetERPDate(ByVal sfb05 As String)
        oCommand.CommandText = "select SUM(sfb08-sfb09-sfb12) as t1,ima02,ima021 FROM SFB_FILE left join ima_file on sfb05 = ima01 WHERE SFB05 = '"
        oCommand.CommandText += sfb05 & "' AND SFB04 <> '8' and sfb87 <> 'X' group by ima02,ima021"
        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            oReader.Read()
            Ws.Cells(LineZ, 4) = oReader.Item("ima02")
            'Ws.Cells(LineZ, 3) = oReader.Item("ima021")
            Ws.Cells(LineZ, 5) = oReader.Item("t1")
        End If
        oReader.Close()
    End Sub
    Private Function GetTotalWIP(ByVal model As String)
        mSQLS2.CommandText = "select count(sn) from sn,lot where sn.lot = lot.lot and lot.model = '" & model & "' and sn.updatedstation <> '9999'"
        Dim WipS As Decimal = mSQLS2.ExecuteScalar()
        Return WipS
    End Function
    Private Function GetEWIP(ByVal model As String)
        mSQLS2.CommandText = "select count(sn) from sn,lot where sn.lot = lot.lot and lot.model = '" & model & "' and sn.updatedstation not in ('9999','0080','0090','0100','0110','0111','0112','0113')"
        Dim WipS As Decimal = mSQLS2.ExecuteScalar()
        Return WipS
    End Function
End Class