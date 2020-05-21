Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.XlThemeColor
Public Class Form78
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form78_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        mSQLS1.CommandText = "select sgroup1,sgroup1en,sum(c1) as c1,sum(c2) as c2,sum(c3) as c3,sum(c4) as c4,sum(c5) as c5,sum(c6) as c6,sum(c7) as c7,sum(c8) as c8,sum(c9) as c9,sum(c10) as c10,"
        mSQLS1.CommandText += "sum(c11) as c11,sum(c12) as c12,sum(c13) as c13,sum(c14) as c14 from ( "
        mSQLS1.CommandText += "select stationdefine.sgroup1,sgroup1en,sGroup1PrintOrder,(case when t1 < 4 then 1 else 0 end) as c1, (case when t1 >= 4 and t1 < 8 then 1 else 0 end) as c2, (case when t1 >= 8 and t1 < 15 then 1 else 0 end) as c3, "
        mSQLS1.CommandText += "(case when t1 >= 15 and t1 < 22 then 1 else 0 end) as c4, (case when t1 >= 22 and t1 < 29 then 1 else 0 end) as c5, (case when t1 >= 29 and t1 < 36 then 1 else 0 end) as c6, "
        mSQLS1.CommandText += "(case when t1 >= 36 and t1 < 43 then 1 else 0 end) as c7, (case when t1 >= 43 and t1 < 50 then 1 else 0 end) as c8, (case when t1 >= 50 and t1 < 57 then 1 else 0 end) as c9, "
        mSQLS1.CommandText += "(case when t1 >= 57 and t1 < 64 then 1 else 0 end) as c10, (case when t1 >= 64 and t1 < 71 then 1 else 0 end) as c11, (case when t1 >= 71 and t1 < 78 then 1 else 0 end) as c12, "
        mSQLS1.CommandText += "(case when t1 >= 78 and t1 < 85 then 1 else 0 end) as c13, (case when t1 >= 85  then 1 else 0 end) as c14 from ( "
        'mSQLS1.CommandText += "select sn,lasttimeout,(case when topreworkstation is null or topreworkstation = '' then updatedstation else topreworkstation end) as stationn, datediff(day,lasttimeout,getdate()) as t1 "
        mSQLS1.CommandText += "select sn,lasttimeout,(case when topreworkstation is null or topreworkstation = '' then updatedstation else topreworkstation end) as stationn, datediff(day,updatedtime,getdate()) as t1 "
        mSQLS1.CommandText += "from sn,lot where sn.lot = lot.lot and lot.model not in ('AB0214','AB0215','AB0216','AB0217') AND lot.remark not like '%PPAP%' AND sn.updatedstation <> '9999' ) as AA left join ERPSUPPORT.dbo.StationDefine on stationn = station left join ERPSUPPORT.dbo.StationDefineExt on StationDefine.SGroup1 =StationDefineExt.sGroup1 "
        mSQLS1.CommandText += "where stationdefine.sgroup1 is not null and stationdefine.SGroup1 <> '成品' ) AS AB group by SGroup1,sgroup1en,sGroup1PrintOrder  order by sGroup1PrintOrder "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = T1.Date
                Ws.Cells(LineZ, 2) = mSQLReader.Item("sgroup1en") & "_" & mSQLReader.Item("sgroup1")
                For i As Int16 = 0 To 13 Step 1
                    If mSQLReader.Item(2 + i) <> 0 Then
                        Ws.Cells(LineZ, 3 + i) = mSQLReader.Item(2 + i)
                    End If
                Next
                Ws.Cells(LineZ, 17) = "=SUM(C" & LineZ & ":P" & LineZ & ")"
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
        ' 加總
        Ws.Cells(2, 3) = "=SUBTOTAL(9,C4:C" & LineZ - 1 & ")"
        oRng = Ws.Range("C2", "C2")
        oRng.AutoFill(Destination:=Ws.Range("C2", "Q2"), Type:=xlFillDefault)
        ' 百分比
        Ws.Cells(1, 3) = "=IF(C2=0,0,C2/$Q$2)"
        oRng = Ws.Range("C1", "C1")
        oRng.AutoFill(Destination:=Ws.Range("C1", "Q1"), Type:=xlFillDefault)
        oRng = Ws.Range("A1", "Q1")
        oRng.EntireColumn.Select()
        oRng.EntireColumn.AutoFit()

        LineZ = 18

        mSQLS1.CommandText = "select sgroup1,sgroup1en,sum(c1) as c1,sum(c2) as c2,sum(c3) as c3,sum(c4) as c4,sum(c5) as c5,sum(c6) as c6,sum(c7) as c7,sum(c8) as c8,sum(c9) as c9,sum(c10) as c10,"
        mSQLS1.CommandText += "sum(c11) as c11,sum(c12) as c12,sum(c13) as c13,sum(c14) as c14 from ( "
        mSQLS1.CommandText += "select stationdefine.sgroup1,sgroup1en,sGroup1PrintOrder,(case when t1 < 4 then 1 else 0 end) as c1, (case when t1 >= 4 and t1 < 8 then 1 else 0 end) as c2, (case when t1 >= 8 and t1 < 15 then 1 else 0 end) as c3, "
        mSQLS1.CommandText += "(case when t1 >= 15 and t1 < 22 then 1 else 0 end) as c4, (case when t1 >= 22 and t1 < 29 then 1 else 0 end) as c5, (case when t1 >= 29 and t1 < 36 then 1 else 0 end) as c6, "
        mSQLS1.CommandText += "(case when t1 >= 36 and t1 < 43 then 1 else 0 end) as c7, (case when t1 >= 43 and t1 < 50 then 1 else 0 end) as c8, (case when t1 >= 50 and t1 < 57 then 1 else 0 end) as c9, "
        mSQLS1.CommandText += "(case when t1 >= 57 and t1 < 64 then 1 else 0 end) as c10, (case when t1 >= 64 and t1 < 71 then 1 else 0 end) as c11, (case when t1 >= 71 and t1 < 78 then 1 else 0 end) as c12, "
        mSQLS1.CommandText += "(case when t1 >= 78 and t1 < 85 then 1 else 0 end) as c13, (case when t1 >= 85  then 1 else 0 end) as c14 from ( "
        'mSQLS1.CommandText += "select sn,lasttimeout,(case when topreworkstation is null or topreworkstation = '' then updatedstation else topreworkstation end) as stationn, datediff(day,lasttimeout,getdate()) as t1 "
        mSQLS1.CommandText += "select sn,lasttimeout,(case when topreworkstation is null or topreworkstation = '' then updatedstation else topreworkstation end) as stationn, datediff(day,updatedtime,getdate()) as t1 "
        mSQLS1.CommandText += "from sn,lot where sn.lot = lot.lot and lot.model in ('AB0214','AB0215','AB0216','AB0217') AND lot.remark not like '%PPAP%' AND sn.updatedstation <> '9999' ) as AA left join ERPSUPPORT.dbo.StationDefine on stationn = station left join ERPSUPPORT.dbo.StationDefineExt on StationDefine.SGroup1 =StationDefineExt.sGroup1 "
        mSQLS1.CommandText += "where stationdefine.sgroup1 is not null and stationdefine.SGroup1 <> '成品' ) AS AB group by SGroup1,sgroup1en,sGroup1PrintOrder  order by sGroup1PrintOrder "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = T1.Date
                Ws.Cells(LineZ, 2) = mSQLReader.Item("sgroup1en") & "_" & mSQLReader.Item("sgroup1")
                For i As Int16 = 0 To 13 Step 1
                    If mSQLReader.Item(2 + i) <> 0 Then
                        Ws.Cells(LineZ, 3 + i) = mSQLReader.Item(2 + i)
                    End If
                Next
                Ws.Cells(LineZ, 17) = "=SUM(C" & LineZ & ":P" & LineZ & ")"
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
        ' 加總
        Ws.Cells(16, 3) = "=SUBTOTAL(9,C18:C" & LineZ - 1 & ")"
        oRng = Ws.Range("C16", "C16")
        oRng.AutoFill(Destination:=Ws.Range("C16", "Q16"), Type:=xlFillDefault)
        ' 百分比
        Ws.Cells(15, 3) = "=IF(C16=0,0,C16/$Q$16)"
        oRng = Ws.Range("C15", "C15")
        oRng.AutoFill(Destination:=Ws.Range("C15", "Q15"), Type:=xlFillDefault)
        oRng = Ws.Range("A1", "Q1")
        oRng.EntireColumn.Select()
        oRng.EntireColumn.AutoFit()


        ' 20170510 新增程式碼
        For i As Int16 = 1 To 9 Step 1
            If i > 2 Then
                Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
            Else
                Ws = xWorkBook.Sheets(i + 1)
            End If
            Ws.Activate()
            Dim WSN As String = String.Empty
            Select Case i
                Case 1
                    WSN = "Cutting"
                Case 2
                    WSN = "Lay-up"
                Case 3
                    WSN = "Molding"
                Case 4
                    WSN = "CNC"
                Case 5
                    WSN = "Gluing"
                Case 6
                    WSN = "Sanding"
                Case 7
                    WSN = "Painting"
                Case 8
                    WSN = "Polishing"
                Case 9
                    WSN = "Packing"
            End Select
            AdjustExcelFormat1(WSN)
            Detail(WSN)
        Next
        ' 20170516 新增程式碼
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws.Activate()
        AdjustExcelFormat2()
        mSQLS1.CommandText = "select cf01,model,modelname,value,stationn,SGroup1,SGroup1En,sn from ( "
        'mSQLS1.CommandText += "select sn,cf01,lot.model,modelname,lasttimeout,(case when topreworkstation is null or topreworkstation = '' then updatedstation else topreworkstation end) as stationn, datediff(day,lasttimeout,getdate()) as t1 "
        mSQLS1.CommandText += "select sn,cf01,lot.model,modelname,value,lasttimeout,(case when topreworkstation is null or topreworkstation = '' then updatedstation else topreworkstation end) as stationn, datediff(day,updatedtime,getdate()) as t1 "
        mSQLS1.CommandText += "from sn left join lot on sn.lot = lot.lot left join model on lot.model = model.model left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue.profilename = 'ERP' "
        mSQLS1.CommandText += "and (case when topreworkstation is null or topreworkstation = '' then updatedstation else topreworkstation end) = model_station_paravalue.station left join model_paravalue on model.model = model_paravalue.model and model_paravalue.parameter = 'ERP PN' "
        mSQLS1.CommandText += "where sn.updatedstation <> '9999' AND lot.remark not like '%PPAP%' ) as AA left join ERPSUPPORT.dbo.StationDefine on stationn = station where t1 >= 43 and SGroup1 is not null and SGroup1 <> '成品' order by sn"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("cf01")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("modelname")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("stationn")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("SGroup1") & " " & mSQLReader.Item("SGroup1En").ToString().Trim()
                Ws.Cells(LineZ, 6) = mSQLReader.Item("sn")
                Ws.Cells(LineZ, 7) = mSQLReader.Item("Value")
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
        oRng = Ws.Range("A1", "G1")
        'oRng.Select()
        oRng.EntireColumn.AutoFit()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "Summary"
        Ws.Cells(1, 2) = "Percentage"
        Ws.Cells(2, 2) = "Total Quantity"
        Ws.Cells(2, 1) = "非结构件"
        Ws.Cells(3, 1) = "取数时点Acquisition Time"
        Ws.Cells(3, 2) = "工站Station"
        Ws.Cells(3, 3) = "<4 Day"
        Ws.Cells(3, 4) = "<8 Day"
        Ws.Cells(3, 5) = "<15 Day"
        Ws.Cells(3, 6) = "<22 Day"
        Ws.Cells(3, 7) = "<29 Day"
        Ws.Cells(3, 8) = "<36 Day"
        Ws.Cells(3, 9) = "<43 Day"
        Ws.Cells(3, 10) = "<50 Day"
        Ws.Cells(3, 11) = "<57 Day"
        Ws.Cells(3, 12) = "<64 Day"
        Ws.Cells(3, 13) = "<71 Day"
        Ws.Cells(3, 14) = "<78 Day"
        Ws.Cells(3, 15) = "<85 Day"
        Ws.Cells(3, 16) = ">=85 Day"
        Ws.Cells(3, 17) = "Total Quantity"
        LineZ = 4
        oRng = Ws.Range("B1", "Q1")
        oRng.Interior.Color = Color.Yellow
        oRng.NumberFormatLocal = "0.00%"
        oRng = Ws.Range("B2", "Q2")
        oRng.Interior.Color = Color.LightBlue
        oRng = Ws.Range("A3", "Q3")
        oRng.Interior.Color = Color.MistyRose

        ' add by cloud 20180622

        Ws.Cells(15, 2) = "Percentage"
        Ws.Cells(16, 2) = "Total Quantity"
        Ws.Cells(16, 1) = "结构件"
        Ws.Cells(17, 1) = "取数时点Acquisition Time"
        Ws.Cells(17, 2) = "工站Station"
        Ws.Cells(17, 3) = "<4 Day"
        Ws.Cells(17, 4) = "<8 Day"
        Ws.Cells(17, 5) = "<15 Day"
        Ws.Cells(17, 6) = "<22 Day"
        Ws.Cells(17, 7) = "<29 Day"
        Ws.Cells(17, 8) = "<36 Day"
        Ws.Cells(17, 9) = "<43 Day"
        Ws.Cells(17, 10) = "<50 Day"
        Ws.Cells(17, 11) = "<57 Day"
        Ws.Cells(17, 12) = "<64 Day"
        Ws.Cells(17, 13) = "<71 Day"
        Ws.Cells(17, 14) = "<78 Day"
        Ws.Cells(17, 15) = "<85 Day"
        Ws.Cells(17, 16) = ">=85 Day"
        Ws.Cells(17, 17) = "Total Quantity"
        'LineZ = 4
        oRng = Ws.Range("B15", "Q15")
        oRng.Interior.Color = Color.Yellow
        oRng.NumberFormatLocal = "0.00%"
        oRng = Ws.Range("B16", "Q16")
        oRng.Interior.Color = Color.LightBlue
        oRng = Ws.Range("A17", "Q17")
        oRng.Interior.Color = Color.MistyRose



    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "WIP库龄报表"
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
    Private Sub AdjustExcelFormat1(ByVal WorkSector As String)
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = WorkSector
        Ws.Cells(1, 3) = "Percentage"
        Ws.Cells(2, 3) = "Total Quantity"
        Ws.Cells(3, 1) = "ERP PN"
        Ws.Cells(3, 2) = "Product Name"
        Ws.Cells(3, 3) = "WIP_Product Description"
        Ws.Cells(3, 4) = "<4 Day"
        Ws.Cells(3, 5) = "<8 Day"
        Ws.Cells(3, 6) = "<15 Day"
        Ws.Cells(3, 7) = "<22 Day"
        Ws.Cells(3, 8) = "<29 Day"
        Ws.Cells(3, 9) = "<36 Day"
        Ws.Cells(3, 10) = "<43 Day"
        Ws.Cells(3, 11) = "<50 Day"
        Ws.Cells(3, 12) = "<57 Day"
        Ws.Cells(3, 13) = "<64 Day"
        Ws.Cells(3, 14) = "<71 Day"
        Ws.Cells(3, 15) = "<78 Day"
        Ws.Cells(3, 16) = "<85 Day"
        Ws.Cells(3, 17) = ">=85 Day"
        Ws.Cells(3, 18) = "Total Quantity"
        LineZ = 4
        oRng = Ws.Range("C1", "R1")
        oRng.Interior.Color = Color.Yellow
        oRng.NumberFormatLocal = "0.00%"
        oRng = Ws.Range("A2", "R2")
        oRng.Interior.Color = Color.LightBlue
        oRng = Ws.Range("D3", "R3")
        oRng.Interior.Color = Color.MistyRose
        oRng = Ws.Range("A3", "C3")
        oRng.Interior.Color = Color.LightBlue
    End Sub
    Private Sub Detail(ByVal WSN As String)
        mSQLS1.CommandText = "select cf01,model,modelname,sum(c1) as c1,sum(c2) as c2,sum(c3) as c3,sum(c4) as c4,sum(c5) as c5,sum(c6) as c6,sum(c7) as c7,sum(c8) as c8,sum(c9) as c9,sum(c10) as c10,"
        mSQLS1.CommandText += "sum(c11) as c11,sum(c12) as c12,sum(c13) as c13,sum(c14) as c14 from ( "
        mSQLS1.CommandText += "select cf01,model,modelname,(case when t1 < 4 then 1 else 0 end) as c1, (case when t1 >= 4 and t1 < 8 then 1 else 0 end) as c2, (case when t1 >= 8 and t1 < 15 then 1 else 0 end) as c3, "
        mSQLS1.CommandText += "(case when t1 >= 15 and t1 < 22 then 1 else 0 end) as c4, (case when t1 >= 22 and t1 < 29 then 1 else 0 end) as c5, (case when t1 >= 29 and t1 < 36 then 1 else 0 end) as c6, "
        mSQLS1.CommandText += "(case when t1 >= 36 and t1 < 43 then 1 else 0 end) as c7, (case when t1 >= 43 and t1 < 50 then 1 else 0 end) as c8, (case when t1 >= 50 and t1 < 57 then 1 else 0 end) as c9, "
        mSQLS1.CommandText += "(case when t1 >= 57 and t1 < 64 then 1 else 0 end) as c10, (case when t1 >= 64 and t1 < 71 then 1 else 0 end) as c11, (case when t1 >= 71 and t1 < 78 then 1 else 0 end) as c12, "
        mSQLS1.CommandText += "(case when t1 >= 78 and t1 < 85 then 1 else 0 end) as c13, (case when t1 >= 84  then 1 else 0 end) as c14 from ( "
        'mSQLS1.CommandText += "select sn,cf01,lot.model,modelname,lasttimeout,(case when topreworkstation is null or topreworkstation = '' then updatedstation else topreworkstation end) as stationn, datediff(day,lasttimeout,getdate()) as t1 "
        mSQLS1.CommandText += "select sn,cf01,lot.model,modelname,lasttimeout,(case when topreworkstation is null or topreworkstation = '' then updatedstation else topreworkstation end) as stationn, datediff(day,updatedtime,getdate()) as t1 "
        mSQLS1.CommandText += "from sn left join lot on sn.lot = lot.lot left join model on lot.model = model.model "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue.profilename = 'ERP' and (case when topreworkstation is null or topreworkstation = '' then updatedstation else topreworkstation end) = model_station_paravalue.station "
        mSQLS1.CommandText += "where sn.updatedstation <> '9999' AND lot.remark not like '%PPAP%' ) as AA left join ERPSUPPORT.dbo.StationDefine on stationn = station  "
        mSQLS1.CommandText += "where stationdefine.sgroup1 is not null and stationdefine.SGroup1En = '" & WSN & "' ) AS AB group by cf01,model,modelname  order by cf01"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                For i As Int16 = 0 To 16 Step 1
                    Ws.Cells(LineZ, i + 1) = mSQLReader.Item(i)
                Next
                Ws.Cells(LineZ, 18) = "=SUM(D" & LineZ & ":Q" & LineZ & ")"
                LineZ += 1
            End While
            oRng = Ws.Range("A4", Ws.Cells(LineZ - 1, 3))
            oRng.Interior.Color = Color.LightBlue
            ' 加總
            Ws.Cells(2, 4) = "=SUBTOTAL(9,D4:D" & LineZ - 1 & ")"
            oRng = Ws.Range("D2", "D2")
            oRng.AutoFill(Destination:=Ws.Range("D2", "R2"), Type:=xlFillDefault)
            ' 百分比
            Ws.Cells(1, 4) = "=IF(D2=0,0,D2/$R$2)"
            oRng = Ws.Range("D1", "D1")
            oRng.AutoFill(Destination:=Ws.Range("D1", "R1"), Type:=xlFillDefault)
            oRng = Ws.Range("A1", "R1")
            oRng.EntireColumn.Select()
            oRng.EntireColumn.AutoFit()
        End If
        mSQLReader.Close()

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
        Ws.Cells(1, 7) = "对应成品ERPPN"
        LineZ = 2
        oRng = Ws.Range("A1", "C1")
        oRng.Interior.Color = Color.LightBlue
        oRng = Ws.Range("D1", "E1")
        oRng.Interior.Color = Color.MistyRose
        oRng = Ws.Range("D1", "D1")
        oRng.EntireColumn.NumberFormatLocal = "@"
    End Sub
End Class