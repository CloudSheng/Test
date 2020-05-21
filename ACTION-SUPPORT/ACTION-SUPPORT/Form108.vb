Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.XlThemeColor
Public Class Form108
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim mSQLReader2 As SqlClient.SqlDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form108_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        'ExportToExcel()
        'SaveExcel()
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        Dim T1 As Date = Now()
        mSQLS1.CommandText = "select stationn,sum(c1) as c1,sum(c2) as c2,sum(c3) as c3,sum(c4) as c4,sum(c5) as c5,sum(c6) as c6,sum(c7) as c7,sum(c8) as c8,sum(c9) as c9,sum(c10) as c10,"
        mSQLS1.CommandText += "sum(c11) as c11,sum(c12) as c12,sum(c13) as c13,sum(c14) as c14 from ( "
        mSQLS1.CommandText += "select stationn,(case when t1 < 4 then 1 else 0 end) as c1, (case when t1 >= 4 and t1 < 8 then 1 else 0 end) as c2, (case when t1 >= 8 and t1 < 15 then 1 else 0 end) as c3, "
        mSQLS1.CommandText += "(case when t1 >= 15 and t1 < 22 then 1 else 0 end) as c4, (case when t1 >= 22 and t1 < 29 then 1 else 0 end) as c5, (case when t1 >= 29 and t1 < 36 then 1 else 0 end) as c6, "
        mSQLS1.CommandText += "(case when t1 >= 36 and t1 < 43 then 1 else 0 end) as c7, (case when t1 >= 43 and t1 < 50 then 1 else 0 end) as c8, (case when t1 >= 50 and t1 < 57 then 1 else 0 end) as c9, "
        mSQLS1.CommandText += "(case when t1 >= 57 and t1 < 64 then 1 else 0 end) as c10, (case when t1 >= 64 and t1 < 71 then 1 else 0 end) as c11, (case when t1 >= 71 and t1 < 78 then 1 else 0 end) as c12, "
        mSQLS1.CommandText += "(case when t1 >= 78 and t1 < 84 then 1 else 0 end) as c13, (case when t1 >= 84  then 1 else 0 end) as c14 from ( "
        mSQLS1.CommandText += "select sn,lasttimeout,(case when topreworkstation is null or topreworkstation = '' then updatedstation else topreworkstation end) as stationn, isnull(datediff(day,lasttimeout,getdate()),0) as t1 "
        mSQLS1.CommandText += "from sn left join lot on sn.lot = lot.lot where sn.updatedstation = '0730' and lot.remark not like '%PPAP%' ) as AA ) AS AB group by stationn"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = Now.ToString("yyyy/MM/dd")
                For i As Integer = 0 To mSQLReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, 2 + i) = mSQLReader.Item(i)
                Next
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
        ' 加總
        Ws.Cells(3, 17) = "=SUM(C3:P3)"
        oRng = Ws.Range("C2", "C2")

        ' 百分比
        Ws.Cells(1, 3) = "=IF(C3=0,0,C3/$Q$3)"
        oRng = Ws.Range("C1", "C1")
        oRng.AutoFill(Destination:=Ws.Range("C1", "Q1"), Type:=xlFillDefault)
        oRng = Ws.Range("A1", "Q1")
        oRng.EntireColumn.Select()
        oRng.EntireColumn.AutoFit()

        'mSQLS1.CommandText = "select distinct SUBSTRING(cf01,4,2)  as t1 from sn left join lot on sn.lot = lot.lot "
        'mSQLS1.CommandText += "left join model_station_paravalue on model_station_paravalue.profilename = 'ERP' and lot.model = model_station_paravalue.model and updatedstation = model_station_paravalue.station "
        'mSQLS1.CommandText += "where sn.updatedstation = '0730' order by t1"

        'mSQLReader = mSQLS1.ExecuteReader()
        'Dim PageNo As Decimal = 2
        'If mSQLReader.HasRows() Then
        'While mSQLReader.Read()
        'If PageNo > 3 Then
        'Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        'Else
        'Ws = xWorkBook.Sheets(PageNo)
        'End If
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        Ws.Name = "Detailed"
        'AdjustExcelFormat1(mSQLReader.Item("t1"))
        AdjustExcelFormat1()
        'Detail(mSQLReader.Item("t1"))
        Detail(1)
        'PageNo += 1
        'End While
        'End If
        mSQLReader.Close()

        ' 20170516 新增程式碼
        'Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        AdjustExcelFormat2()
        mSQLS1.CommandText = "select cf01,model,modelname,stationn,sn,t1 from ( "
        'mSQLS1.CommandText += "select sn,cf01,lot.model,modelname,lasttimeout,(case when topreworkstation is null or topreworkstation = '' then updatedstation else topreworkstation end) as stationn, datediff(day,lasttimeout,getdate()) as t1 "
        mSQLS1.CommandText += "select sn,cf01,lot.model,modelname,lasttimeout,(case when topreworkstation is null or topreworkstation = '' then updatedstation else topreworkstation end) as stationn, isnull(datediff(day,lasttimeout,getdate()),0) as t1 "
        mSQLS1.CommandText += "from sn left join lot on sn.lot = lot.lot left join model on lot.model = model.model left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue.profilename = 'ERP' "
        mSQLS1.CommandText += "and (case when topreworkstation is null or topreworkstation = '' then updatedstation else topreworkstation end) = model_station_paravalue.station "
        mSQLS1.CommandText += "where sn.updatedstation = '0730' and lot.remark not like '%PPAP%' ) as AA where t1 >= 43 order by sn"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("cf01")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("modelname")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("stationn")
                Ws.Cells(LineZ, 5) = "成品"
                Ws.Cells(LineZ, 6) = mSQLReader.Item("sn")
                Ws.Cells(LineZ, 7) = mSQLReader.Item("t1")
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
        oRng = Ws.Range("A1", "G1")
        oRng.EntireColumn.Select()
        oRng.EntireColumn.AutoFit()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "Summary"
        Ws.Cells(1, 2) = "Percentage"
        Ws.Cells(2, 1) = "取数时点Acquisition Time"
        Ws.Cells(2, 2) = "工站Station"
        Ws.Cells(2, 3) = "<4 Day"
        Ws.Cells(2, 4) = "<8 Day"
        Ws.Cells(2, 5) = "<15 Day"
        Ws.Cells(2, 6) = "<22 Day"
        Ws.Cells(2, 7) = "<29 Day"
        Ws.Cells(2, 8) = "<36 Day"
        Ws.Cells(2, 9) = "<43 Day"
        Ws.Cells(2, 10) = "<50 Day"
        Ws.Cells(2, 11) = "<57 Day"
        Ws.Cells(2, 12) = "<64 Day"
        Ws.Cells(2, 13) = "<71 Day"
        Ws.Cells(2, 14) = "<78 Day"
        Ws.Cells(2, 15) = "<85 Day"
        Ws.Cells(2, 16) = ">=85 Day"
        Ws.Cells(2, 17) = "Total Quantity"
        LineZ = 3
        oRng = Ws.Range("B1", "Q1")
        oRng.Interior.Color = Color.Yellow
        oRng.NumberFormatLocal = "0.00%"
        oRng = Ws.Range("A2", "Q2")
        oRng.Interior.Color = Color.MistyRose
        oRng = Ws.Range("B3", "B3")
        oRng.NumberFormatLocal = "@"

    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        'Ws.Name = WorkSector
        Ws.Cells(1, 1) = "Customer"
        Ws.Cells(1, 2) = "ERP PN"
        Ws.Cells(1, 3) = "Product Name"
        Ws.Cells(1, 4) = "WIP_Product Description"
        Ws.Cells(1, 5) = "<4 Day"
        Ws.Cells(1, 6) = "<8 Day"
        Ws.Cells(1, 7) = "<15 Day"
        Ws.Cells(1, 8) = "<22 Day"
        Ws.Cells(1, 9) = "<29 Day"
        Ws.Cells(1, 10) = "<36 Day"
        Ws.Cells(1, 11) = "<43 Day"
        Ws.Cells(1, 12) = "<50 Day"
        Ws.Cells(1, 13) = "<57 Day"
        Ws.Cells(1, 14) = "<64 Day"
        Ws.Cells(1, 15) = "<71 Day"
        Ws.Cells(1, 16) = "<78 Day"
        Ws.Cells(1, 17) = "<85 Day"
        Ws.Cells(1, 18) = ">=85 Day"
        Ws.Cells(1, 19) = "Total Quantity"
        LineZ = 2
        oRng = Ws.Range("A1", "D1")
        oRng.Interior.Color = Color.LightBlue
        oRng = Ws.Range("E1", "S1")
        oRng.Interior.Color = Color.MistyRose
    End Sub
    Private Sub Detail(ByVal WSN As String)
        mSQLS2.CommandText = "select substring(cf01,4,2),cf01,model,modelname,sum(c1) as c1,sum(c2) as c2,sum(c3) as c3,sum(c4) as c4,sum(c5) as c5,sum(c6) as c6,sum(c7) as c7,sum(c8) as c8,sum(c9) as c9,sum(c10) as c10,"
        mSQLS2.CommandText += "sum(c11) as c11,sum(c12) as c12,sum(c13) as c13,sum(c14) as c14 from ( "
        mSQLS2.CommandText += "select cf01,model,modelname,(case when t1 < 4 then 1 else 0 end) as c1, (case when t1 >= 4 and t1 < 8 then 1 else 0 end) as c2, (case when t1 >= 8 and t1 < 15 then 1 else 0 end) as c3, "
        mSQLS2.CommandText += "(case when t1 >= 15 and t1 < 22 then 1 else 0 end) as c4, (case when t1 >= 22 and t1 < 29 then 1 else 0 end) as c5, (case when t1 >= 29 and t1 < 36 then 1 else 0 end) as c6, "
        mSQLS2.CommandText += "(case when t1 >= 36 and t1 < 43 then 1 else 0 end) as c7, (case when t1 >= 43 and t1 < 50 then 1 else 0 end) as c8, (case when t1 >= 50 and t1 < 57 then 1 else 0 end) as c9, "
        mSQLS2.CommandText += "(case when t1 >= 57 and t1 < 64 then 1 else 0 end) as c10, (case when t1 >= 64 and t1 < 71 then 1 else 0 end) as c11, (case when t1 >= 71 and t1 < 78 then 1 else 0 end) as c12, "
        mSQLS2.CommandText += "(case when t1 >= 78 and t1 < 84 then 1 else 0 end) as c13, (case when t1 >= 84  then 1 else 0 end) as c14 from ( "
        mSQLS2.CommandText += "select cf01,lot.model,model.modelname,lasttimeout,(case when topreworkstation is null or topreworkstation = '' then updatedstation else topreworkstation end) as stationn, isnull(datediff(day,lasttimeout,getdate()),0) as t1 "
        mSQLS2.CommandText += "from sn left join lot on sn.lot = lot.lot left join model on lot.model = model.model left join model_station_paravalue on model_station_paravalue.profilename = 'ERP' and lot.model = model_station_paravalue.model and updatedstation = model_station_paravalue.station "
        'mSQLS2.CommandText += "and model.model = model_station_paravalue.model where sn.updatedstation = '0730' and substring(cf01,4,2) = '" & WSN & "' ) as AA ) AS AB group by cf01,model,modelname"
        mSQLS2.CommandText += "and model.model = model_station_paravalue.model where sn.updatedstation = '0730' ) as AA ) AS AB group by cf01,model,modelname"
        mSQLReader2 = mSQLS2.ExecuteReader()
        If mSQLReader2.HasRows() Then
            While mSQLReader2.Read()
                For i As Int16 = 0 To 17 Step 1
                    Ws.Cells(LineZ, i + 1) = mSQLReader2.Item(i)
                Next
                Ws.Cells(LineZ, 19) = "=SUM(E" & LineZ & ":R" & LineZ & ")"
                LineZ += 1
            End While
            oRng = Ws.Range("A2", Ws.Cells(LineZ - 1, 4))
            oRng.Interior.Color = Color.LightBlue

            oRng = Ws.Range("A1", "S1")
            oRng.EntireColumn.Select()
            oRng.EntireColumn.AutoFit()
        End If
        mSQLReader2.Close()

    End Sub
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = ">=43 SN"
        Ws.Cells(1, 1) = "ERP PN"
        Ws.Cells(1, 2) = "Product Name"
        Ws.Cells(1, 3) = "WIP_Product Description"
        Ws.Cells(1, 4) = "Station"
        Ws.Cells(1, 5) = "Sector"
        Ws.Cells(1, 6) = "SN"
        Ws.Cells(1, 7) = "Inventory Days"
        LineZ = 2
        oRng = Ws.Range("A1", "C1")
        oRng.Interior.Color = Color.LightBlue
        oRng = Ws.Range("D1", "E1")
        oRng.Interior.Color = Color.MistyRose
        oRng = Ws.Range("D1", "D1")
        oRng.EntireColumn.NumberFormatLocal = "@"
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "成品库龄报表"
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
End Class