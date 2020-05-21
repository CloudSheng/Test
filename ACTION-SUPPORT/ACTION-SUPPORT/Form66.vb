Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form66
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim tModel As String
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form66_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
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
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "MES_Basic_Data_Report"
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
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        AdjustExcelFormat()
        mSQLS1.CommandText = "select stationname_cn,station,stationdesc from station  order by station "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("stationname_cn")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("station")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("stationdesc")
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustExcelFormat1()
        mSQLS1.CommandText = "select equipment_id,equipment_name  from z_ms_equipment "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("equipment_id")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("equipment_name")
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        AdjustExcelFormat2()
        mSQLS1.CommandText = "select model.model,model.modelname,model.model_type,model_paravalue.value,SUBSTRING(model_paravalue.value,4,2) as c1 from model left join model_paravalue on model_paravalue.model = model.model and parameter = 'ERP PN'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("value")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("c1")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("modelname")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("model_type")
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
        ' 第4頁  20171229
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws.Activate()
        AdjustExcelFormat3()
        mSQLS1.CommandText = "SELECT model.model,model.modelname,model_station_paravalue.cf01,routing.route,routing.seq,routing.station,station.stationname,model_paravalue.value  FROM MODEL "
        mSQLS1.CommandText += "left join model_paravalue on model.model = model_paravalue.model and model_paravalue.parameter = 'Accessory' "
        mSQLS1.CommandText += "LEFT JOIN ROUTING ON MODEL.default_route = ROUTING.ROUTE left join model_station_paravalue on model_station_paravalue.model = model.model and model_station_paravalue.profilename = 'ERP' "
        mSQLS1.CommandText += "and model_station_paravalue.station = routing.station left join station on station.station = model_station_paravalue.station and station.station = routing.station order by model.model,routing.seq"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("modelname")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("cf01")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("route")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("seq")
                Ws.Cells(LineZ, 6) = mSQLReader.Item("station")
                Ws.Cells(LineZ, 7) = mSQLReader.Item("stationname")
                Ws.Cells(LineZ, 8) = mSQLReader.Item("value")
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "工站代码清单"
        oRng = Ws.Range("A1", "C1")
        oRng.EntireColumn.ColumnWidth = 16.5
        Ws.Cells(1, 1) = "Name工站"
        Ws.Cells(1, 2) = "工站ID"
        Ws.Cells(1, 3) = "stationdesc"
        oRng = Ws.Range("B1", "C1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "设备清单"
        oRng = Ws.Range("A1", "B1")
        oRng.EntireColumn.ColumnWidth = 20.2
        Ws.Cells(1, 1) = "Equipment ID"
        Ws.Cells(1, 2) = "Equipment Name"
        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "产品型号清单"
        oRng = Ws.Range("A1", "E1")
        oRng.EntireColumn.ColumnWidth = 20.2
        Ws.Cells(1, 1) = "ERP NO"
        Ws.Cells(1, 2) = "客户代码"
        Ws.Cells(1, 3) = "Part No."
        Ws.Cells(1, 4) = "Description"
        Ws.Cells(1, 5) = "Customer"
        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat3()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "Routing资料"
        oRng = Ws.Range("A1", "H1")
        oRng.EntireColumn.ColumnWidth = 18.2
        Ws.Cells(1, 1) = "model"
        Ws.Cells(1, 2) = "modelname"
        Ws.Cells(1, 3) = "ERP"
        Ws.Cells(1, 4) = "Routing"
        Ws.Cells(1, 5) = "seq"
        Ws.Cells(1, 6) = "station"
        Ws.Cells(1, 7) = "stationname"
        Ws.Cells(1, 8) = "Accessory"
        oRng = Ws.Range("F1", "F1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        LineZ = 2
    End Sub
End Class