Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form95
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim pYear As Int16 = 0
    Dim pMonth As Int16 = 0
    Dim Start2 As Date
    Dim End2 As Date
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineS As Integer = 0
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form95_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        Label3.Text = "未开始"
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
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        tYear = Me.NumericUpDown1.Value
        tMonth = Me.NumericUpDown2.Value
        pYear = tYear
        pMonth = tMonth - 1
        If pMonth = 0 Then
            pMonth = 12
            pYear -= 1
        End If

        Start2 = Convert.ToDateTime(tYear & "/" & tMonth & "/01 08:00:00")
        End2 = Start2.AddMonths(1)
        ExportToExcel()
        SaveExcel()
        'BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "MES_MONTH_SCRAP"
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
    Private Sub ExportToExcel()
        Label3.Text = "处理明细中"
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Name = "报废明细表"
        Ws.Activate()
        AdjustExcelFormat()
        mSQLS1.CommandText = "select scrap_sn.updatedstation,scrap.defect ,defect.desc_th ,defect.desc_en,lot.model,cf01"
        For i As Int16 = 1 To 31 Step 1
            mSQLS1.CommandText += ",(case when day(scrap.datetime - '08:00:00') = " & i & " then count(scrap_sn.sn) else 0 end  ) as t" & i
        Next
        mSQLS1.CommandText += " from scrap left join scrap_sn on scrap.sn = scrap_sn.sn left join lot on scrap.lot = lot.lot left join defect on scrap.defect = defect.defect "
        mSQLS1.CommandText += "left join model_station_paravalue on model_station_paravalue.profilename = 'ERP' and scrap_sn.updatedstation = model_station_paravalue.station "
        mSQLS1.CommandText += "and lot.model = model_station_paravalue.model where scrap.datetime between '"
        mSQLS1.CommandText += Start2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS1.CommandText += End2.ToString("yyyy/MM/dd HH:mm:ss") & "' group by scrap_sn.updatedstation,scrap.defect ,defect.desc_th ,defect.desc_en,lot.model,cf01,scrap.datetime "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Dim WorkSector As String = String.Empty
                Ws.Cells(LineZ, 1) = mSQLReader.Item("updatedstation")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("defect") & " " & mSQLReader.Item("desc_en") & " " & mSQLReader.Item("desc_th")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("cf01")
                If Strings.Right(mSQLReader.Item("cf01").ToString(), 1) = "A" Or Strings.Right(mSQLReader.Item("cf01").ToString(), 1) = "B" Then
                    WorkSector = Strings.Right(mSQLReader.Item("cf01").ToString(), 3)
                    Select Case WorkSector
                        Case "32A", "32B"
                            Ws.Cells(LineZ, 2) = "预型"
                        Case "35A", "35B"
                            Ws.Cells(LineZ, 2) = "成型"
                        Case "36A", "36B"
                            Ws.Cells(LineZ, 2) = "CNC"
                        Case "61A", "61B"
                            Ws.Cells(LineZ, 2) = "补土"
                        Case "63A", "63B"
                            Ws.Cells(LineZ, 2) = "涂装"
                        Case "64A"
                            Ws.Cells(LineZ, 2) = "二次胶合"
                        Case "64B"
                            Ws.Cells(LineZ, 2) = "三次胶合"
                        Case "65A", "65B"
                            Ws.Cells(LineZ, 2) = "抛光"
                        Case "66A", "66B"
                            Ws.Cells(LineZ, 2) = "包装"
                    End Select
                Else
                    WorkSector = Strings.Right(mSQLReader.Item("cf01").ToString(), 2)
                    Select Case WorkSector
                        Case "32"
                            Ws.Cells(LineZ, 2) = "预型"
                        Case "35"
                            Ws.Cells(LineZ, 2) = "成型"
                        Case "36"
                            Ws.Cells(LineZ, 2) = "CNC"
                        Case "61"
                            Ws.Cells(LineZ, 2) = "补土"
                        Case "63"
                            Ws.Cells(LineZ, 2) = "涂装"
                        Case "64"
                            Ws.Cells(LineZ, 2) = "胶合"
                        Case "65"
                            Ws.Cells(LineZ, 2) = "抛光"
                        Case "66"
                            Ws.Cells(LineZ, 2) = "包装"
                    End Select
                End If
                Dim GRC As Decimal = 0
                If Not IsDBNull(mSQLReader.Item("cf01")) Then
                    GRC = GetRealCost(mSQLReader.Item("cf01"))
                End If
                'Dim GRC As Decimal = GetRealCost(mSQLReader.Item("cf01"))
                Dim TRW As Decimal = 0
                Dim TRM As Decimal = 0
                Ws.Cells(LineZ, 6) = GRC
                For i As Int16 = 1 To 31 Step 1
                    Ws.Cells(LineZ, 5 + i * 2) = mSQLReader.Item(5 + i)
                    TRW += mSQLReader.Item(5 + i)
                    Ws.Cells(LineZ, 6 + i * 2) = mSQLReader.Item(5 + i) * GRC
                    TRM += mSQLReader.Item(5 + i) * GRC
                Next
                Ws.Cells(LineZ, 69) = TRW
                Ws.Cells(LineZ, 70) = TRM
                LineZ += 1
            End While
        End If
        mSQLReader.Close()

        ' 加總欄
        Ws.Cells(LineZ, 7) = "=SUM(G3:G" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 7))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 70)), Type:=xlFillDefault)
        ' 劃線
        oRng = Ws.Range("A1", Ws.Cells(LineZ, 70))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        Label3.Text = "处理成本中"
        mSQLS1.CommandText = "CREATE TABLE ERPSUPPORT.dbo.CostTemp (ccc01 nvarchar(40), ccc23 numeric(20,6))"
        Try
            mSQLS1.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
            Return
        End Try
        oCommand.CommandText = "select ccc01,ccc23 from ccc_file,ima_file where ccc02 = " & pYear & " and ccc03 = " & pMonth & " and ccc01 = ima01 and ima08 = 'M' and ccc23 <> 0"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                mSQLS1.CommandText = "INSERT INTO ERPSUPPORT.dbo.CostTemp Values ('" & oReader.Item("ccc01") & "'," & oReader.Item("ccc23") & ")"
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End While
        End If
        oReader.Close()
        Label3.Text = "处理汇总表中"
        Ws = xWorkBook.Sheets(2)
        Ws.Name = "MES报废报表"
        Ws.Activate()
        AdjustExcelFormat1()

        ' 日數量
        mSQLS1.CommandText = "select c1 "
        For i As Int16 = 1 To 31 Step 1
            mSQLS1.CommandText += ",sum(t" & i & ") as t" & i
        Next
        mSQLS1.CommandText += " from ( select 0 as c1"
        For i As Int16 = 1 To 31 Step 1
            mSQLS1.CommandText += ",(case when day(scrap.datetime - '08:00:00') = " & i & " then count(scrap.sn) else 0 end  ) as t" & i
        Next
        mSQLS1.CommandText += " from scrap  where scrap.datetime between '"
        mSQLS1.CommandText += Start2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS1.CommandText += End2.ToString("yyyy/MM/dd HH:mm:ss") & "'group by datetime ) AS AB group by c1 "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                For i As Int16 = 1 To mSQLReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = mSQLReader.Item(i)
                Next
                Ws.Cells(LineZ, 33) = "=SUM(B3:AF3)"
            End While
        End If
        mSQLReader.Close()
        LineZ += 1

        ' 日金額
        mSQLS1.CommandText = "select c1"
        For i As Int16 = 1 To 31 Step 1
            mSQLS1.CommandText += ",sum(t" & i & ") as t" & i
        Next
        mSQLS1.CommandText += " from ( select 0 as c1"
        For i As Int16 = 1 To 31 Step 1
            mSQLS1.CommandText += ",(case when day(scrap.datetime - '08:00:00') = " & i & " then count(scrap.sn) * ccc23 else 0 end  ) as t" & i
        Next
        mSQLS1.CommandText += " from scrap left join scrap_sn on scrap.sn = scrap_sn.sn left join lot on scrap.lot = lot.lot  "
        mSQLS1.CommandText += "left join model_station_paravalue on model_station_paravalue.profilename = 'ERP' and scrap_sn.updatedstation = model_station_paravalue.station "
        mSQLS1.CommandText += "and lot.model = model_station_paravalue.model left join [ERPSUPPORT].dbo.CostTemp on cf01 = ccc01 where scrap.datetime between '"
        mSQLS1.CommandText += Start2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS1.CommandText += End2.ToString("yyyy/MM/dd HH:mm:ss") & "' group by scrap.datetime,ccc23 ) aS AB group by c1"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                For i As Int16 = 1 To mSQLReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = mSQLReader.Item(i)
                Next
                Ws.Cells(LineZ, 33) = "=SUM(B4:AF4)"
            End While
        End If
        mSQLReader.Close()
        ' 劃線
        oRng = Ws.Range("A2", Ws.Cells(LineZ, 33))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        ' 日部門數量
        LineZ += 4
        mSQLS1.CommandText = "select c1"
        For i As Int16 = 1 To 31 Step 1
            mSQLS1.CommandText += ",sum(t" & i & ") as t" & i
        Next
        mSQLS1.CommandText += " from (select (case when len(cf01) = 15 then (case when SUBSTRING(cf01,14,2) = '31' then '裁纱'  when SUBSTRING(cf01,14,2) = '32' then '预型' "
        mSQLS1.CommandText += " when SUBSTRING(cf01,14, 2) = '35' then '成型' when SUBSTRING(cf01,14,2) = '36' then 'CNC'  when SUBSTRING(cf01,14,2) = '61' then '补土' "
        mSQLS1.CommandText += " when SUBSTRING(cf01,14,2) = '63' then '涂装'  when SUBSTRING(cf01,14,2) = '64' then '胶合'  when SUBSTRING(cf01,14,2) = '65' then '抛光' "
        mSQLS1.CommandText += " when SUBSTRING(cf01,14,2) = '66' then '包装'  else '其他'  end)  else  (case when SUBSTRING(cf01,14,3) = '31A' then '裁纱' "
        mSQLS1.CommandText += " when SUBSTRING(cf01,14,3) = '32A' then '预型'  when SUBSTRING(cf01,14, 3) = '35A' then '成型'  when SUBSTRING(cf01,14,3) = '36A' then 'CNC' "
        mSQLS1.CommandText += " when SUBSTRING(cf01,14,3) = '61A' then '补土'  when SUBSTRING(cf01,14,3) = '63A' then '涂装'  when SUBSTRING(cf01,14,3) = '64A' then '胶合' "
        mSQLS1.CommandText += " when SUBSTRING(cf01,14,3) = '64B' then '胶合'  when SUBSTRING(cf01,14,3) = '65A' then '抛光'  when SUBSTRING(cf01,14,3) = '66A' then '包装' "
        mSQLS1.CommandText += " else '其他'  end) end) as c1 "
        For i As Int16 = 1 To 31 Step 1
            mSQLS1.CommandText += ",(case when day(scrap.datetime - '08:00:00') = " & i & " then count(scrap.sn)  else 0 end  ) as t" & i
        Next
        mSQLS1.CommandText += " from scrap left join scrap_sn on scrap.sn = scrap_sn.sn left join lot on scrap.lot = lot.lot  "
        mSQLS1.CommandText += "left join model_station_paravalue on model_station_paravalue.profilename = 'ERP' and scrap_sn.updatedstation = model_station_paravalue.station "
        mSQLS1.CommandText += "and lot.model = model_station_paravalue.model where scrap.datetime between '"
        mSQLS1.CommandText += Start2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS1.CommandText += End2.ToString("yyyy/MM/dd HH:mm:ss") & "' group by scrap.datetime,cf01 ) aS AB group by c1"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                For i As Int16 = 0 To mSQLReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = mSQLReader.Item(i)
                Next
                Ws.Cells(LineZ, 33) = "=SUM(B" & LineZ & ":AF" & LineZ & ")"
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
        ' 劃線
        oRng = Ws.Range("A7", Ws.Cells(LineZ - 1, 33))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        ' 定錨 
        LineS = LineZ
        AdjustExcelFormat2()

        ' 日部門金額
        mSQLS1.CommandText = "select c1"
        For i As Int16 = 1 To 31 Step 1
            mSQLS1.CommandText += ",sum(t" & i & ") as t" & i
        Next
        mSQLS1.CommandText += " from (select (case when len(cf01) = 15 then (case when SUBSTRING(cf01,14,2) = '31' then '裁纱'  when SUBSTRING(cf01,14,2) = '32' then '预型' "
        mSQLS1.CommandText += " when SUBSTRING(cf01,14, 2) = '35' then '成型' when SUBSTRING(cf01,14,2) = '36' then 'CNC'  when SUBSTRING(cf01,14,2) = '61' then '补土' "
        mSQLS1.CommandText += " when SUBSTRING(cf01,14,2) = '63' then '涂装'  when SUBSTRING(cf01,14,2) = '64' then '胶合'  when SUBSTRING(cf01,14,2) = '65' then '抛光' "
        mSQLS1.CommandText += " when SUBSTRING(cf01,14,2) = '66' then '包装'  else '其他'  end)  else  (case when SUBSTRING(cf01,14,3) = '31A' then '裁纱' "
        mSQLS1.CommandText += " when SUBSTRING(cf01,14,3) = '32A' then '预型'  when SUBSTRING(cf01,14, 3) = '35A' then '成型'  when SUBSTRING(cf01,14,3) = '36A' then 'CNC' "
        mSQLS1.CommandText += " when SUBSTRING(cf01,14,3) = '61A' then '补土'  when SUBSTRING(cf01,14,3) = '63A' then '涂装'  when SUBSTRING(cf01,14,3) = '64A' then '胶合' "
        mSQLS1.CommandText += " when SUBSTRING(cf01,14,3) = '64B' then '胶合'  when SUBSTRING(cf01,14,3) = '65A' then '抛光'  when SUBSTRING(cf01,14,3) = '66A' then '包装' "
        mSQLS1.CommandText += " else '其他'  end) end) as c1 "
        For i As Int16 = 1 To 31 Step 1
            mSQLS1.CommandText += ",(case when day(scrap.datetime - '08:00:00') = " & i & " then count(scrap.sn) * ccc23  else 0 end  ) as t" & i
        Next
        mSQLS1.CommandText += " from scrap left join scrap_sn on scrap.sn = scrap_sn.sn left join lot on scrap.lot = lot.lot  "
        mSQLS1.CommandText += "left join model_station_paravalue on model_station_paravalue.profilename = 'ERP' and scrap_sn.updatedstation = model_station_paravalue.station "
        mSQLS1.CommandText += "and lot.model = model_station_paravalue.model left join [ERPSUPPORT].dbo.CostTemp on cf01 = ccc01 where scrap.datetime between '"
        mSQLS1.CommandText += Start2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS1.CommandText += End2.ToString("yyyy/MM/dd HH:mm:ss") & "' group by scrap.datetime,cf01,ccc23 ) aS AB group by c1"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                For i As Int16 = 0 To mSQLReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = mSQLReader.Item(i)
                Next
                Ws.Cells(LineZ, 33) = "=SUM(B" & LineZ & ":AF" & LineZ & ")"
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
        ' 劃線
        oRng = Ws.Range(Ws.Cells(LineS + 2, 1), Ws.Cells(LineZ - 1, 33))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        ' 定錨 
        LineS = LineZ
        AdjustExcelFormat3()

        mSQLS1.CommandText = "select defect,desc_th,desc_en"
        For i As Int16 = 1 To 31 Step 1
            mSQLS1.CommandText += ",sum(t" & i & ") as t" & i
        Next
        mSQLS1.CommandText += " from ( select scrap.defect ,defect.desc_th ,defect.desc_en "
        For i As Int16 = 1 To 31 Step 1
            mSQLS1.CommandText += ",(case when day(scrap.datetime - '08:00:00') = " & i & " then count(scrap.sn) else 0 end  ) as t" & i
        Next
        mSQLS1.CommandText += " from scrap left join defect on scrap.defect = defect.defect where scrap.datetime between '"
        mSQLS1.CommandText += Start2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS1.CommandText += End2.ToString("yyyy/MM/dd HH:mm:ss") & "' group by scrap.defect ,defect.desc_th ,defect.desc_en,scrap.datetime "
        mSQLS1.CommandText += ") AS AB group by defect,desc_th,desc_en"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("defect") & " " & mSQLReader.Item("desc_th") & " " & mSQLReader.Item("desc_en")
                For i As Int16 = 1 To 31 Step 1
                    Ws.Cells(LineZ, i + 1) = mSQLReader.Item(2 + i)
                Next
                Ws.Cells(LineZ, 33) = "=SUM(B" & LineZ & ":AF" & LineZ & ")"
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
        ' 劃線
        oRng = Ws.Range(Ws.Cells(LineS + 2, 1), Ws.Cells(LineZ - 1, 33))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        Label3.Text = "丟棄臨時表"
        mSQLS1.CommandText = "DROP TABLE [ERPSUPPORT].dbo.CostTemp "
        Try
            mSQLS1.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        Label3.Text = "完成"
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 14.78
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.ColumnWidth = 40.89
        oRng = Ws.Range("A1", "A2")
        oRng.Merge()
        oRng.EntireColumn.NumberFormatLocal = "@"
        'oRng = Ws.Range("A1", "A1")
        oRng.AutoFill(Destination:=Ws.Range("A1", "F2"), Type:=xlFillDefault)
        Ws.Cells(1, 1) = "station"
        Ws.Cells(1, 2) = "对应部门"
        Ws.Cells(1, 3) = "缺陷原因"
        Ws.Cells(1, 4) = "Part Name"
        Ws.Cells(1, 5) = "ERP料号"
        Ws.Cells(1, 6) = "上月期末加权平均单价"
        oRng = Ws.Range("F1", "F2")
        oRng.WrapText = True
        oRng = Ws.Range("G1", "H1")
        oRng.Merge()
        Ws.Cells(1, 7) = tMonth & "/1"
        Ws.Cells(2, 7) = "数量"
        Ws.Cells(2, 8) = "金额"
        oRng = Ws.Range("G1", "H2")
        oRng.AutoFill(Destination:=Ws.Range("G1", "BR2"), Type:=xlFillDefault)
        Ws.Cells(1, 69) = "MTD"

        LineZ = 3
    End Sub
    Private Function GetRealCost(ByVal ccc01 As String)
        oCommand.CommandText = "select nvl(ccc23,0) from ccc_file where ccc01 = '" & ccc01 & "' and ccc02 = " & pYear & " and ccc03 = " & pMonth
        Dim GRC As Decimal = oCommand.ExecuteScalar()
        Return GRC
    End Function
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 8.22
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 37.11
        oRng = Ws.Range("A1", "AG1")
        oRng.Merge()
        Ws.Cells(1, 1) = "部门报废数量及金额汇总表"
        Ws.Cells(2, 1) = "汇总/Tatal"
        Ws.Cells(3, 1) = "数量qty"
        Ws.Cells(4, 1) = "金额RMB"
        Ws.Cells(2, 2) = tMonth & "/1"
        oRng = Ws.Range("B2", "B2")
        oRng.AutoFill(Destination:=Ws.Range("B2", "AF2"), Type:=xlFillDefault)
        Ws.Cells(2, 33) = "MTD"

        oRng = Ws.Range("A6", "AG6")
        oRng.Merge()
        Ws.Cells(6, 1) = "部门报废数量明细表"
        Ws.Cells(7, 1) = "部门Section"
        Ws.Cells(7, 2) = tMonth & "/1"
        oRng = Ws.Range("B7", "B7")
        oRng.AutoFill(Destination:=Ws.Range("B7", "AF7"), Type:=xlFillDefault)
        Ws.Cells(7, 33) = "MTD"

        LineZ = 3
    End Sub
    Private Sub AdjustExcelFormat2()

        oRng = Ws.Range(Ws.Cells(LineS + 1, 1), Ws.Cells(LineS + 1, 33))
        oRng.Merge()
        Ws.Cells(LineS + 2, 1) = "部门报废金额明细表"
        Ws.Cells(LineS + 2, 1) = "部门Section"
        Ws.Cells(LineS + 2, 2) = tMonth & "/1"
        oRng = Ws.Range(Ws.Cells(LineS + 2, 2), Ws.Cells(LineS + 2, 2))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineS + 2, 2), Ws.Cells(LineS + 2, 32)), Type:=xlFillDefault)
        Ws.Cells(LineS + 2, 33) = "MTD"

        LineZ += 3
    End Sub
    Private Sub AdjustExcelFormat3()

        oRng = Ws.Range(Ws.Cells(LineS + 1, 1), Ws.Cells(LineS + 1, 33))
        oRng.Merge()
        Ws.Cells(LineS + 2, 1) = "缺陷原因明细表"
        Ws.Cells(LineS + 2, 1) = "缺陷原因Defect Code"
        Ws.Cells(LineS + 2, 2) = tMonth & "/1"
        oRng = Ws.Range(Ws.Cells(LineS + 2, 2), Ws.Cells(LineS + 2, 2))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineS + 2, 2), Ws.Cells(LineS + 2, 32)), Type:=xlFillDefault)
        Ws.Cells(LineS + 2, 33) = "MTD"

        LineZ += 3
    End Sub
End Class