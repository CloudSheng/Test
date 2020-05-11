Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form76
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim Date1 As Date
    Dim Date2 As Date
    Dim Date3 As Date
    Dim Day4 As Long = 0
    Dim OCC01 As String = String.Empty
    Dim Occ02 As String = String.Empty
    Dim ogb04 As String = String.Empty
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form76_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("hkacttest")
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
        oCommand.CommandText = "SELECT OCC01 FROM OCC_FILE WHERE OCCACTI = 'Y'"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Me.ComboBox1.Items.Add(oReader.Item(0).ToString())
            End While
        End If
        oReader.Close()
        oCommand.CommandText = "SELECT OCC02 FROM OCC_FILE WHERE OCCACTI = 'Y'"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Me.ComboBox2.Items.Add(oReader.Item(0).ToString())
            End While
        End If
        oReader.Close()

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
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        ' 客户
        If Not IsNothing(ComboBox1.SelectedItem) Then
            OCC01 = ComboBox1.SelectedItem.ToString()
        Else
            OCC01 = String.Empty
        End If
        If Not IsNothing(ComboBox2.SelectedItem) Then
            Occ02 = ComboBox2.SelectedItem.ToString()
        Else
            Occ02 = String.Empty
        End If
        If Not String.IsNullOrEmpty(TextBox1.Text) Then
            ogb04 = TextBox1.Text
        Else
            ogb04 = String.Empty
        End If
        Date1 = DateTimePicker1.Value '指定日期
        Date2 = Date1.AddDays(Date1.Day * Decimal.MinusOne + 1)   '該月月初
        Date3 = Date1.AddDays(Date1.DayOfYear * Decimal.MinusOne + 1)   '該年年初
        Day4 = DateDiff(DateInterval.Day, Date2, Date1) + 1
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "HAC_SALES_DAILY_REPORT"
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
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets.Add()   '第四頁
        Ws = xWorkBook.Sheets.Add()   '第五頁
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        Ws.Name = "by customer销售日报表"
        AdjustExcelFormat()
        oCommand.CommandText = "select oga03,oga032,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,"
        oCommand.CommandText += "sum(t10) as t10,sum(t11) as t11,sum(t12) as t12,sum(t13) as t13,sum(t14) as t14,sum(t15) as t15,sum(t16) as t16,sum(t17) as t17,sum(t18) as t18,sum(t19) as t19,sum(t20) as t20,sum(t21) as t21,"
        oCommand.CommandText += "sum(t22) as t22,sum(t23) as t23,sum(t24) as t24,sum(t25) as t25,sum(t26) as t26,sum(t27) as t27,sum(t28) as t28,sum(t29) as t29,sum(t30) as t30,sum(t31) as t31 from ( "
        oCommand.CommandText += "SELECT oga03,oga032,case when day(oga02) = 1 then ogb14t * oga24 else 0 end as t1,case when day(oga02) = 2 then ogb14t * oga24 else 0 end as t2,case when day(oga02) = 3 then ogb14t * oga24 else 0 end as t3,"
        oCommand.CommandText += "case when day(oga02) = 4 then ogb14t * oga24 else 0 end as t4,case when day(oga02) = 5 then ogb14t * oga24 else 0 end as t5,case when day(oga02) = 6 then ogb14t * oga24 else 0 end as t6,"
        oCommand.CommandText += "case when day(oga02) = 7 then ogb14t * oga24 else 0 end as t7,case when day(oga02) = 8 then ogb14t * oga24 else 0 end as t8,case when day(oga02) = 9 then ogb14t * oga24 else 0 end as t9,"
        oCommand.CommandText += "case when day(oga02) = 10 then ogb14t * oga24 else 0 end as t10,case when day(oga02) = 11 then ogb14t * oga24 else 0 end as t11,case when day(oga02) = 12 then ogb14t * oga24 else 0 end as t12,"
        oCommand.CommandText += "case when day(oga02) = 13 then ogb14t * oga24 else 0 end as t13,case when day(oga02) = 14 then ogb14t * oga24 else 0 end as t14,case when day(oga02) = 15 then ogb14t * oga24 else 0 end as t15,"
        oCommand.CommandText += "case when day(oga02) = 16 then ogb14t * oga24 else 0 end as t16,case when day(oga02) = 17 then ogb14t * oga24 else 0 end as t17,case when day(oga02) = 18 then ogb14t * oga24 else 0 end as t18,"
        oCommand.CommandText += "case when day(oga02) = 19 then ogb14t * oga24 else 0 end as t19,case when day(oga02) = 20 then ogb14t * oga24 else 0 end as t20,case when day(oga02) = 21 then ogb14t * oga24 else 0 end as t21,"
        oCommand.CommandText += "case when day(oga02) = 22 then ogb14t * oga24 else 0 end as t22,case when day(oga02) = 23 then ogb14t * oga24 else 0 end as t23,case when day(oga02) = 24 then ogb14t * oga24 else 0 end as t24,"
        oCommand.CommandText += "case when day(oga02) = 25 then ogb14t * oga24 else 0 end as t25,case when day(oga02) = 26 then ogb14t * oga24 else 0 end as t26,case when day(oga02) = 27 then ogb14t * oga24 else 0 end as t27,"
        oCommand.CommandText += "case when day(oga02) = 28 then ogb14t * oga24 else 0 end as t28,case when day(oga02) = 29 then ogb14t * oga24 else 0 end as t29,case when day(oga02) = 30 then ogb14t * oga24 else 0 end as t30,"
        oCommand.CommandText += "case when day(oga02) = 31 then ogb14t * oga24 else 0 end as t31  FROM oga_file,ogb_file where oga01 = ogb01 and ogapost = 'Y' AND ogb04 <> 'AC0000000000' "
        If Not String.IsNullOrEmpty(OCC01) Then
            oCommand.CommandText += "AND oga03 = '" & OCC01 & "' "
        End If
        If Not String.IsNullOrEmpty(Occ02) Then
            oCommand.CommandText += "AND oga032 = '" & Occ02 & "' "
        End If
        If Not String.IsNullOrEmpty(ogb04) Then
            oCommand.CommandText += "AND ogb04 LIKE '%" & ogb04 & "%' "
        End If
        oCommand.CommandText += "and oga02 between to_date('"
        oCommand.CommandText += Date2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & Date1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        ' MODIFY 20170331
        oCommand.CommandText += "union all select oga03,oga032,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 from oga_file,ogb_file where oga01 = ogb01 and ogapost = 'Y' and ogb04 <> 'AC0000000000' and oga02 between to_date('"
        oCommand.CommandText += Date3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += Date2.AddDays(-1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        If Not String.IsNullOrEmpty(OCC01) Then
            oCommand.CommandText += "AND oga03 = '" & OCC01 & "' "
        End If
        If Not String.IsNullOrEmpty(Occ02) Then
            oCommand.CommandText += "AND oga032 = '" & Occ02 & "' "
        End If
        If Not String.IsNullOrEmpty(ogb04) Then
            oCommand.CommandText += "AND ogb04 LIKE '%" & ogb04 & "%' "
        End If
        oCommand.CommandText += ") group by oga03,oga032 order by oga03"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("oga03")
                Ws.Cells(LineZ, 2) = oReader.Item("oga032")
                Ws.Cells(LineZ, 3) = GetYearSales(oReader.Item("oga03"))
                Ws.Cells(LineZ, 4) = GetMonthSales(oReader.Item("oga03"))
                For j As Integer = 0 To Day4 - 1 Step 1
                    Ws.Cells(LineZ, 5 + j) = oReader.Item(2 + j)
                Next
                LineZ += 1
            End While
        End If
        oReader.Close()
        '加總
        Ws.Cells(LineZ, 2) = "Total合計"
        Ws.Cells(LineZ, 3) = "=SUM(C2:C" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 4 + Day4)), Type:=xlFillDefault)
        
        '第二頁 20170330
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        Ws.Name = "by product产品别销售日报表"
        AdjustExcelFormat1()
        oCommand.CommandText = "select oga032,ogb04,ima02,ima021,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,"
        oCommand.CommandText += "sum(t10) as t10,sum(t11) as t11,sum(t12) as t12,sum(t13) as t13,sum(t14) as t14,sum(t15) as t15,sum(t16) as t16,sum(t17) as t17,sum(t18) as t18,sum(t19) as t19,sum(t20) as t20,sum(t21) as t21,"
        oCommand.CommandText += "sum(t22) as t22,sum(t23) as t23,sum(t24) as t24,sum(t25) as t25,sum(t26) as t26,sum(t27) as t27,sum(t28) as t28,sum(t29) as t29,sum(t30) as t30,sum(t31) as t31 from ( "
        oCommand.CommandText += "SELECT oga032,ogb04,ima02,ima021,case when day(oga02) = 1 then ogb14t * oga24 else 0 end as t1,case when day(oga02) = 2 then ogb14t * oga24 else 0 end as t2,"
        oCommand.CommandText += "case when day(oga02) = 3 then ogb14t * oga24 else 0 end as t3,case when day(oga02) = 4 then ogb14t * oga24 else 0 end as t4,case when day(oga02) = 5 then ogb14t * oga24 else 0 end as t5,"
        oCommand.CommandText += "case when day(oga02) = 6 then ogb14t * oga24 else 0 end as t6,case when day(oga02) = 7 then ogb14t * oga24 else 0 end as t7,case when day(oga02) = 8 then ogb14t * oga24 else 0 end as t8,"
        oCommand.CommandText += "case when day(oga02) = 9 then ogb14t * oga24 else 0 end as t9,case when day(oga02) = 10 then ogb14t * oga24 else 0 end as t10,case when day(oga02) = 11 then ogb14t * oga24 else 0 end as t11,"
        oCommand.CommandText += "case when day(oga02) = 12 then ogb14t * oga24 else 0 end as t12,case when day(oga02) = 13 then ogb14t * oga24 else 0 end as t13,case when day(oga02) = 14 then ogb14t * oga24 else 0 end as t14,"
        oCommand.CommandText += "case when day(oga02) = 15 then ogb14t * oga24 else 0 end as t15,case when day(oga02) = 16 then ogb14t * oga24 else 0 end as t16,case when day(oga02) = 17 then ogb14t * oga24 else 0 end as t17,"
        oCommand.CommandText += "case when day(oga02) = 18 then ogb14t * oga24 else 0 end as t18,case when day(oga02) = 19 then ogb14t * oga24 else 0 end as t19,case when day(oga02) = 20 then ogb14t * oga24 else 0 end as t20,"
        oCommand.CommandText += "case when day(oga02) = 21 then ogb14t * oga24 else 0 end as t21,case when day(oga02) = 22 then ogb14t * oga24 else 0 end as t22,case when day(oga02) = 23 then ogb14t * oga24 else 0 end as t23,"
        oCommand.CommandText += "case when day(oga02) = 24 then ogb14t * oga24 else 0 end as t24,case when day(oga02) = 25 then ogb14t * oga24 else 0 end as t25,case when day(oga02) = 26 then ogb14t * oga24 else 0 end as t26,"
        oCommand.CommandText += "case when day(oga02) = 27 then ogb14t * oga24 else 0 end as t27,case when day(oga02) = 28 then ogb14t * oga24 else 0 end as t28,case when day(oga02) = 29 then ogb14t * oga24 else 0 end as t29,"
        oCommand.CommandText += "case when day(oga02) = 30 then ogb14t * oga24 else 0 end as t30,case when day(oga02) = 31 then ogb14t * oga24 else 0 end as t31 FROM oga_file,ogb_file,ima_file where oga01 = ogb01 and ogapost = 'Y' AND ogb04 <> 'AC0000000000' "
        If Not String.IsNullOrEmpty(OCC01) Then
            oCommand.CommandText += "AND oga03 = '" & OCC01 & "' "
        End If
        If Not String.IsNullOrEmpty(Occ02) Then
            oCommand.CommandText += "AND oga032 = '" & Occ02 & "' "
        End If
        If Not String.IsNullOrEmpty(ogb04) Then
            oCommand.CommandText += "AND ogb04 LIKE '%" & ogb04 & "%' "
        End If
        oCommand.CommandText += "and ogb04 = ima01 and oga02 between to_date('" & Date2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & Date1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        ' MODIFY 20170331
        oCommand.CommandText += "union all select oga032,ogb04,ima02,ima021,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 from oga_file,ogb_file,ima_file where oga01 = ogb01 and ogapost = 'Y' and ogb04 <> 'AC0000000000' and ogb04 = ima01 and oga02 between to_date('"
        oCommand.CommandText += Date3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += Date2.AddDays(-1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        If Not String.IsNullOrEmpty(OCC01) Then
            oCommand.CommandText += "AND oga03 = '" & OCC01 & "' "
        End If
        If Not String.IsNullOrEmpty(Occ02) Then
            oCommand.CommandText += "AND oga032 = '" & Occ02 & "' "
        End If
        If Not String.IsNullOrEmpty(ogb04) Then
            oCommand.CommandText += "AND ogb04 LIKE '%" & ogb04 & "%' "
        End If
        oCommand.CommandText += ") group by oga032,ogb04,ima02,ima021 order by oga032"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("oga032")
                Ws.Cells(LineZ, 2) = oReader.Item("ogb04")
                Ws.Cells(LineZ, 3) = oReader.Item("ima02")
                Ws.Cells(LineZ, 4) = oReader.Item("ima021")
                Ws.Cells(LineZ, 5) = GetYearSales(oReader.Item("oga032"), oReader.Item("ogb04"))
                Ws.Cells(LineZ, 6) = GetMonthSales(oReader.Item("oga032"), oReader.Item("ogb04"))
                For j As Integer = 0 To Day4 - 1 Step 1
                    Ws.Cells(LineZ, 7 + j) = oReader.Item(4 + j)
                Next
                LineZ += 1
            End While
        End If
        oReader.Close()
        '加總
        Ws.Cells(LineZ, 4) = "Total合計"
        Ws.Cells(LineZ, 5) = "=SUM(E2:E" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 5), Ws.Cells(LineZ, 5))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 5), Ws.Cells(LineZ, 6 + Day4)), Type:=xlFillDefault)

        '第三頁 20170331
        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        Ws.Name = "HAC Tooling"
        AdjustExcelFormat2()
        oCommand.CommandText = "select ofa032,ofbud02,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,"
        oCommand.CommandText += "sum(t10) as t10,sum(t11) as t11,sum(t12) as t12,sum(t13) as t13,sum(t14) as t14,sum(t15) as t15,sum(t16) as t16,sum(t17) as t17,sum(t18) as t18,sum(t19) as t19,sum(t20) as t20,sum(t21) as t21,"
        oCommand.CommandText += "sum(t22) as t22,sum(t23) as t23,sum(t24) as t24,sum(t25) as t25,sum(t26) as t26,sum(t27) as t27,sum(t28) as t28,sum(t29) as t29,sum(t30) as t30,sum(t31) as t31 from ( "
        oCommand.CommandText += "SELECT ofa032,ofbud02,case when day(ofa02) = 1 then ofb14t * ofa24 else 0 end as t1,case when day(ofa02) = 2 then ofb14t * ofa24 else 0 end as t2,"
        oCommand.CommandText += "case when day(ofa02) = 3 then ofb14t * ofa24 else 0 end as t3,case when day(ofa02) = 4 then ofb14t * ofa24 else 0 end as t4,case when day(ofa02) = 5 then ofb14t * ofa24 else 0 end as t5,"
        oCommand.CommandText += "case when day(ofa02) = 6 then ofb14t * ofa24 else 0 end as t6,case when day(ofa02) = 7 then ofb14t * ofa24 else 0 end as t7,case when day(ofa02) = 8 then ofb14t * ofa24 else 0 end as t8,"
        oCommand.CommandText += "case when day(ofa02) = 9 then ofb14t * ofa24 else 0 end as t9,case when day(ofa02) = 10 then ofb14t * ofa24 else 0 end as t10,case when day(ofa02) = 11 then ofb14t * ofa24 else 0 end as t11,"
        oCommand.CommandText += "case when day(ofa02) = 12 then ofb14t * ofa24 else 0 end as t12,case when day(ofa02) = 13 then ofb14t * ofa24 else 0 end as t13,case when day(ofa02) = 14 then ofb14t * ofa24 else 0 end as t14,"
        oCommand.CommandText += "case when day(ofa02) = 15 then ofb14t * ofa24 else 0 end as t15,case when day(ofa02) = 16 then ofb14t * ofa24 else 0 end as t16,case when day(ofa02) = 17 then ofb14t * ofa24 else 0 end as t17,"
        oCommand.CommandText += "case when day(ofa02) = 18 then ofb14t * ofa24 else 0 end as t18,case when day(ofa02) = 19 then ofb14t * ofa24 else 0 end as t19,case when day(ofa02) = 20 then ofb14t * ofa24 else 0 end as t20,"
        oCommand.CommandText += "case when day(ofa02) = 21 then ofb14t * ofa24 else 0 end as t21,case when day(ofa02) = 22 then ofb14t * ofa24 else 0 end as t22,case when day(ofa02) = 23 then ofb14t * ofa24 else 0 end as t23,"
        oCommand.CommandText += "case when day(ofa02) = 24 then ofb14t * ofa24 else 0 end as t24,case when day(ofa02) = 25 then ofb14t * ofa24 else 0 end as t25,case when day(ofa02) = 26 then ofb14t * ofa24 else 0 end as t26,"
        oCommand.CommandText += "case when day(ofa02) = 27 then ofb14t * ofa24 else 0 end as t27,case when day(ofa02) = 28 then ofb14t * ofa24 else 0 end as t28,case when day(ofa02) = 29 then ofb14t * ofa24 else 0 end as t29,"
        oCommand.CommandText += "case when day(ofa02) = 30 then ofb14t * ofa24 else 0 end as t30,case when day(ofa02) = 31 then ofb14t * ofa24 else 0 end as t31 FROM ofa_file,ofb_file where ofa01 = ofb01 and ofaconf = 'Y' AND ofb04 = 'AC0000000000' "
        If Not String.IsNullOrEmpty(OCC01) Then
            oCommand.CommandText += "AND ofa03 = '" & OCC01 & "' "
        End If
        If Not String.IsNullOrEmpty(Occ02) Then
            oCommand.CommandText += "AND ofa032 = '" & Occ02 & "' "
        End If
        If Not String.IsNullOrEmpty(ogb04) Then
            oCommand.CommandText += "AND ofb04 LIKE '%" & ogb04 & "%' "
        End If
        oCommand.CommandText += "and ofa02 between to_date('" & Date2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & Date1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        ' MODIFY 20170331
        oCommand.CommandText += "union all select ofa032,ofbud02,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 from ofa_file,ofb_file where ofa01 = ofb01 and ofaconf = 'Y' and ofb04 = 'AC0000000000' and ofa02 between to_date('"
        oCommand.CommandText += Date3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += Date2.AddDays(-1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        If Not String.IsNullOrEmpty(OCC01) Then
            oCommand.CommandText += "AND ofa03 = '" & OCC01 & "' "
        End If
        If Not String.IsNullOrEmpty(Occ02) Then
            oCommand.CommandText += "AND ofa032 = '" & Occ02 & "' "
        End If
        If Not String.IsNullOrEmpty(ogb04) Then
            oCommand.CommandText += "AND ofb04 LIKE '%" & ogb04 & "%' "
        End If
        oCommand.CommandText += ") group by ofa032,ofbud02 order by ofa032"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("ofa032")
                'Ws.Cells(LineZ, 2) = oReader.Item("ofb04")
                'Ws.Cells(LineZ, 3) = oReader.Item("ima02")
                Ws.Cells(LineZ, 2) = oReader.Item("ofbud02")
                Ws.Cells(LineZ, 3) = GetYearSales1(oReader.Item("ofa032"), oReader.Item("ofbud02"))
                Ws.Cells(LineZ, 4) = GetMonthSales1(oReader.Item("ofa032"), oReader.Item("ofbud02"))
                For j As Integer = 0 To Day4 - 1 Step 1
                    Ws.Cells(LineZ, 5 + j) = oReader.Item(2 + j)
                Next
                LineZ += 1
            End While
        End If
        oReader.Close()
        '加總
        Ws.Cells(LineZ, 2) = "Total合計"
        Ws.Cells(LineZ, 3) = "=SUM(C2:C" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 4 + Day4)), Type:=xlFillDefault)

        '第四頁  20170406

        Ws = xWorkBook.Sheets(4)
        Ws.Activate()
        Ws.Name = "by customer-Qty"
        AdjustExcelFormat()
        oCommand.CommandText = "select oga03,oga032,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,"
        oCommand.CommandText += "sum(t10) as t10,sum(t11) as t11,sum(t12) as t12,sum(t13) as t13,sum(t14) as t14,sum(t15) as t15,sum(t16) as t16,sum(t17) as t17,sum(t18) as t18,sum(t19) as t19,sum(t20) as t20,sum(t21) as t21,"
        oCommand.CommandText += "sum(t22) as t22,sum(t23) as t23,sum(t24) as t24,sum(t25) as t25,sum(t26) as t26,sum(t27) as t27,sum(t28) as t28,sum(t29) as t29,sum(t30) as t30,sum(t31) as t31 from ( "
        oCommand.CommandText += "SELECT oga03,oga032,case when day(oga02) = 1 then ogb12 else 0 end as t1,case when day(oga02) = 2 then ogb12 else 0 end as t2,case when day(oga02) = 3 then ogb12 else 0 end as t3,"
        oCommand.CommandText += "case when day(oga02) = 4 then ogb12 else 0 end as t4,case when day(oga02) = 5 then ogb12 else 0 end as t5,case when day(oga02) = 6 then ogb12 else 0 end as t6,"
        oCommand.CommandText += "case when day(oga02) = 7 then ogb12 else 0 end as t7,case when day(oga02) = 8 then ogb12 else 0 end as t8,case when day(oga02) = 9 then ogb12 else 0 end as t9,"
        oCommand.CommandText += "case when day(oga02) = 10 then ogb12 else 0 end as t10,case when day(oga02) = 11 then ogb12 else 0 end as t11,case when day(oga02) = 12 then ogb12 else 0 end as t12,"
        oCommand.CommandText += "case when day(oga02) = 13 then ogb12 else 0 end as t13,case when day(oga02) = 14 then ogb12 else 0 end as t14,case when day(oga02) = 15 then ogb12 else 0 end as t15,"
        oCommand.CommandText += "case when day(oga02) = 16 then ogb12 else 0 end as t16,case when day(oga02) = 17 then ogb12 else 0 end as t17,case when day(oga02) = 18 then ogb12 else 0 end as t18,"
        oCommand.CommandText += "case when day(oga02) = 19 then ogb12 else 0 end as t19,case when day(oga02) = 20 then ogb12 else 0 end as t20,case when day(oga02) = 21 then ogb12 else 0 end as t21,"
        oCommand.CommandText += "case when day(oga02) = 22 then ogb12 else 0 end as t22,case when day(oga02) = 23 then ogb12 else 0 end as t23,case when day(oga02) = 24 then ogb12 else 0 end as t24,"
        oCommand.CommandText += "case when day(oga02) = 25 then ogb12 else 0 end as t25,case when day(oga02) = 26 then ogb12 else 0 end as t26,case when day(oga02) = 27 then ogb12 else 0 end as t27,"
        oCommand.CommandText += "case when day(oga02) = 28 then ogb12 else 0 end as t28,case when day(oga02) = 29 then ogb12 else 0 end as t29,case when day(oga02) = 30 then ogb12 else 0 end as t30,"
        oCommand.CommandText += "case when day(oga02) = 31 then ogb12 else 0 end as t31  FROM oga_file,ogb_file where oga01 = ogb01 and ogapost = 'Y' AND ogb04 <> 'AC0000000000' "
        If Not String.IsNullOrEmpty(OCC01) Then
            oCommand.CommandText += "AND oga03 = '" & OCC01 & "' "
        End If
        If Not String.IsNullOrEmpty(Occ02) Then
            oCommand.CommandText += "AND oga032 = '" & Occ02 & "' "
        End If
        If Not String.IsNullOrEmpty(ogb04) Then
            oCommand.CommandText += "AND ogb04 LIKE '%" & ogb04 & "%' "
        End If
        oCommand.CommandText += "and oga02 between to_date('"
        oCommand.CommandText += Date2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & Date1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        ' MODIFY 20170331
        oCommand.CommandText += "union all select oga03,oga032,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 from oga_file,ogb_file where oga01 = ogb01 and ogapost = 'Y' and ogb04 <> 'AC0000000000' and oga02 between to_date('"
        oCommand.CommandText += Date3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += Date2.AddDays(-1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        If Not String.IsNullOrEmpty(OCC01) Then
            oCommand.CommandText += "AND oga03 = '" & OCC01 & "' "
        End If
        If Not String.IsNullOrEmpty(Occ02) Then
            oCommand.CommandText += "AND oga032 = '" & Occ02 & "' "
        End If
        If Not String.IsNullOrEmpty(ogb04) Then
            oCommand.CommandText += "AND ogb04 LIKE '%" & ogb04 & "%' "
        End If
        oCommand.CommandText += ") group by oga03,oga032 order by oga03"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("oga03")
                Ws.Cells(LineZ, 2) = oReader.Item("oga032")
                Ws.Cells(LineZ, 3) = GetYearSalesQuantity(oReader.Item("oga03"))
                Ws.Cells(LineZ, 4) = GetMonthSalesQuantity(oReader.Item("oga03"))
                For j As Integer = 0 To Day4 - 1 Step 1
                    Ws.Cells(LineZ, 5 + j) = oReader.Item(2 + j)
                Next
                LineZ += 1
            End While
        End If
        oReader.Close()
        '加總
        Ws.Cells(LineZ, 2) = "Total合計"
        Ws.Cells(LineZ, 3) = "=SUM(C2:C" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 4 + Day4)), Type:=xlFillDefault)

        '第五頁
        Ws = xWorkBook.Sheets(5)
        Ws.Activate()
        Ws.Name = "by product-Qty"
        AdjustExcelFormat1()
        oCommand.CommandText = "select oga032,ogb04,ima02,ima021,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,"
        oCommand.CommandText += "sum(t10) as t10,sum(t11) as t11,sum(t12) as t12,sum(t13) as t13,sum(t14) as t14,sum(t15) as t15,sum(t16) as t16,sum(t17) as t17,sum(t18) as t18,sum(t19) as t19,sum(t20) as t20,sum(t21) as t21,"
        oCommand.CommandText += "sum(t22) as t22,sum(t23) as t23,sum(t24) as t24,sum(t25) as t25,sum(t26) as t26,sum(t27) as t27,sum(t28) as t28,sum(t29) as t29,sum(t30) as t30,sum(t31) as t31 from ( "
        oCommand.CommandText += "SELECT oga032,ogb04,ima02,ima021,case when day(oga02) = 1 then ogb12 else 0 end as t1,case when day(oga02) = 2 then ogb12 else 0 end as t2,"
        oCommand.CommandText += "case when day(oga02) = 3 then ogb12 else 0 end as t3,case when day(oga02) = 4 then ogb12 else 0 end as t4,case when day(oga02) = 5 then ogb12 else 0 end as t5,"
        oCommand.CommandText += "case when day(oga02) = 6 then ogb12 else 0 end as t6,case when day(oga02) = 7 then ogb12 else 0 end as t7,case when day(oga02) = 8 then ogb12 else 0 end as t8,"
        oCommand.CommandText += "case when day(oga02) = 9 then ogb12 else 0 end as t9,case when day(oga02) = 10 then ogb12 else 0 end as t10,case when day(oga02) = 11 then ogb12 else 0 end as t11,"
        oCommand.CommandText += "case when day(oga02) = 12 then ogb12 else 0 end as t12,case when day(oga02) = 13 then ogb12 else 0 end as t13,case when day(oga02) = 14 then ogb12 else 0 end as t14,"
        oCommand.CommandText += "case when day(oga02) = 15 then ogb12 else 0 end as t15,case when day(oga02) = 16 then ogb12 else 0 end as t16,case when day(oga02) = 17 then ogb12 else 0 end as t17,"
        oCommand.CommandText += "case when day(oga02) = 18 then ogb12 else 0 end as t18,case when day(oga02) = 19 then ogb12 else 0 end as t19,case when day(oga02) = 20 then ogb12 else 0 end as t20,"
        oCommand.CommandText += "case when day(oga02) = 21 then ogb12 else 0 end as t21,case when day(oga02) = 22 then ogb12 else 0 end as t22,case when day(oga02) = 23 then ogb12 else 0 end as t23,"
        oCommand.CommandText += "case when day(oga02) = 24 then ogb12 else 0 end as t24,case when day(oga02) = 25 then ogb12 else 0 end as t25,case when day(oga02) = 26 then ogb12 else 0 end as t26,"
        oCommand.CommandText += "case when day(oga02) = 27 then ogb12 else 0 end as t27,case when day(oga02) = 28 then ogb12 else 0 end as t28,case when day(oga02) = 29 then ogb12 else 0 end as t29,"
        oCommand.CommandText += "case when day(oga02) = 30 then ogb12 else 0 end as t30,case when day(oga02) = 31 then ogb12 else 0 end as t31 FROM oga_file,ogb_file,ima_file where oga01 = ogb01 and ogapost = 'Y' AND ogb04 <> 'AC0000000000' "
        If Not String.IsNullOrEmpty(OCC01) Then
            oCommand.CommandText += "AND oga03 = '" & OCC01 & "' "
        End If
        If Not String.IsNullOrEmpty(Occ02) Then
            oCommand.CommandText += "AND oga032 = '" & Occ02 & "' "
        End If
        If Not String.IsNullOrEmpty(ogb04) Then
            oCommand.CommandText += "AND ogb04 LIKE '%" & ogb04 & "%' "
        End If
        oCommand.CommandText += "and ogb04 = ima01 and oga02 between to_date('" & Date2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & Date1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        ' MODIFY 20170331
        oCommand.CommandText += "union all select oga032,ogb04,ima02,ima021,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 from oga_file,ogb_file,ima_file where oga01 = ogb01 and ogapost = 'Y' and ogb04 <> 'AC0000000000' and ogb04 = ima01 and oga02 between to_date('"
        oCommand.CommandText += Date3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += Date2.AddDays(-1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        If Not String.IsNullOrEmpty(OCC01) Then
            oCommand.CommandText += "AND oga03 = '" & OCC01 & "' "
        End If
        If Not String.IsNullOrEmpty(Occ02) Then
            oCommand.CommandText += "AND oga032 = '" & Occ02 & "' "
        End If
        If Not String.IsNullOrEmpty(ogb04) Then
            oCommand.CommandText += "AND ogb04 LIKE '%" & ogb04 & "%' "
        End If
        oCommand.CommandText += ") group by oga032,ogb04,ima02,ima021 order by oga032"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("oga032")
                Ws.Cells(LineZ, 2) = oReader.Item("ogb04")
                Ws.Cells(LineZ, 3) = oReader.Item("ima02")
                Ws.Cells(LineZ, 4) = oReader.Item("ima021")
                Ws.Cells(LineZ, 5) = GetYearSalesQuantity(oReader.Item("oga032"), oReader.Item("ogb04"))
                Ws.Cells(LineZ, 6) = GetMonthSalesQuantity(oReader.Item("oga032"), oReader.Item("ogb04"))
                For j As Integer = 0 To Day4 - 1 Step 1
                    Ws.Cells(LineZ, 7 + j) = oReader.Item(4 + j)
                Next
                LineZ += 1
            End While
        End If
        oReader.Close()
        '加總
        Ws.Cells(LineZ, 4) = "Total合計"
        Ws.Cells(LineZ, 5) = "=SUM(E2:E" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 5), Ws.Cells(LineZ, 5))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 5), Ws.Cells(LineZ, 6 + Day4)), Type:=xlFillDefault)
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.ColumnWidth = 17.44
        Ws.Cells(1, 1) = "客户编号Customer Code"
        Ws.Cells(1, 2) = "客户简称C_SName"
        Ws.Cells(1, 3) = "YTD-Yearly"
        Ws.Cells(1, 4) = "YTD-Monthly"
        For i As Integer = 1 To Day4 Step 1
            Ws.Cells(1, 4 + i) = Date2.AddDays(i - 1).ToString("yyyy/MM/dd")
        Next
        oRng = Ws.Range(Ws.Cells(1, 3), Ws.Cells(1, 4 + Day4))
        oRng.EntireColumn.NumberFormatLocal = "#,##.0000_ "
        oRng.NumberFormatLocal = "yyyy/mm/dd"
        LineZ = 2
    End Sub
    Private Function GetYearSales(ByVal oga03 As String)
        oCommand2.CommandText = "select nvl(sum(ogb14t*oga24),0) from oga_file,ogb_file where oga01 =  ogb01 and ogapost = 'Y' AND ogb04 <> 'AC0000000000' and oga02 between to_date('"
        oCommand2.CommandText += Date3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand2.CommandText += Date1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and oga03 = '"
        oCommand2.CommandText += oga03 & "' "
        If Not String.IsNullOrEmpty(ogb04) Then
            oCommand2.CommandText += "AND ogb04 LIKE '%" & ogb04 & "%' "
        End If
        Dim YearSales As Decimal = oCommand2.ExecuteScalar()
        Return YearSales
    End Function
    Private Function GetMonthSales(ByVal oga03 As String)
        oCommand2.CommandText = "select nvl(sum(ogb14t*oga24),0) from oga_file,ogb_file where oga01 =  ogb01 and ogapost = 'Y' AND ogb04 <> 'AC0000000000' and oga02 between to_date('"
        oCommand2.CommandText += Date2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand2.CommandText += Date1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and oga03 = '"
        oCommand2.CommandText += oga03 & "' "
        If Not String.IsNullOrEmpty(ogb04) Then
            oCommand2.CommandText += "AND ogb04 LIKE '%" & ogb04 & "%' "
        End If
        Dim MonthSales As Decimal = oCommand2.ExecuteScalar()
        Return MonthSales
    End Function
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.ColumnWidth = 17.44
        Ws.Cells(1, 1) = "客户简称C_SName"
        Ws.Cells(1, 2) = "产品料号Part code"
        Ws.Cells(1, 3) = "产品名称Part_N"
        Ws.Cells(1, 4) = "产品规格Part_Desc."
        Ws.Cells(1, 5) = "YTD-Yearly"
        Ws.Cells(1, 6) = "YTD-Monthly"
        For i As Integer = 1 To Day4 Step 1
            Ws.Cells(1, 6 + i) = Date2.AddDays(i - 1).ToString("yyyy/MM/dd")
        Next
        oRng = Ws.Range(Ws.Cells(1, 5), Ws.Cells(1, 6 + Day4))
        oRng.EntireColumn.NumberFormatLocal = "#,##.0000_ "
        oRng.NumberFormatLocal = "yyyy/mm/dd"
        LineZ = 2
    End Sub
    Private Function GetYearSales(ByVal oga032 As String, ByVal ogb04 As String)
        oCommand2.CommandText = "select nvl(sum(ogb14t*oga24),0) from oga_file,ogb_file where oga01 =  ogb01 and ogapost = 'Y' and oga02 between to_date('"
        oCommand2.CommandText += Date3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand2.CommandText += Date1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and oga032 = '"
        oCommand2.CommandText += oga032 & "' AND ogb04 = '" & ogb04 & "' "
        Dim YearSales As Decimal = oCommand2.ExecuteScalar()
        Return YearSales
    End Function
    Private Function GetMonthSales(ByVal oga032 As String, ByVal ogb04 As String)
        oCommand2.CommandText = "select nvl(sum(ogb14t*oga24),0) from oga_file,ogb_file where oga01 =  ogb01 and ogapost = 'Y' and oga02 between to_date('"
        oCommand2.CommandText += Date2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand2.CommandText += Date1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and oga032 = '"
        oCommand2.CommandText += oga032 & "' AND ogb04 = '" & ogb04 & "' "
        Dim MonthSales As Decimal = oCommand2.ExecuteScalar()
        Return MonthSales
    End Function
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.ColumnWidth = 17.44
        Ws.Cells(1, 1) = "客户简称C_SName"
        Ws.Cells(1, 2) = "产品规格Part_Desc."
        Ws.Cells(1, 3) = "YTD-Yearly"
        Ws.Cells(1, 4) = "YTD-Monthly"
        For i As Integer = 1 To Day4 Step 1
            Ws.Cells(1, 4 + i) = Date2.AddDays(i - 1).ToString("yyyy/MM/dd")
        Next
        oRng = Ws.Range(Ws.Cells(1, 3), Ws.Cells(1, 4 + Day4))
        oRng.EntireColumn.NumberFormatLocal = "#,##.0000_ "
        oRng.NumberFormatLocal = "yyyy/mm/dd"
        LineZ = 2
    End Sub
    Private Function GetYearSales1(ByVal ofa032 As String, ByVal ofbud02 As String)
        oCommand2.CommandText = "select nvl(sum(ofb14t*ofa24),0) from ofa_file,ofb_file where ofa01 =  ofb01 and ofaconf = 'Y' and ofa02 between to_date('"
        oCommand2.CommandText += Date3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand2.CommandText += Date1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ofa032 = '"
        oCommand2.CommandText += ofa032 & "' AND ofbud02 = '" & ofbud02 & "' "
        Dim YearSales As Decimal = oCommand2.ExecuteScalar()
        Return YearSales
    End Function
    Private Function GetMonthSales1(ByVal ofa032 As String, ByVal ofbud02 As String)
        oCommand2.CommandText = "select nvl(sum(ofb14t*ofa24),0) from ofa_file,ofb_file where ofa01 =  ofb01 and ofaconf = 'Y' and ofa02 between to_date('"
        oCommand2.CommandText += Date2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand2.CommandText += Date1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ofa032 = '"
        oCommand2.CommandText += ofa032 & "' AND ofbud02 = '" & ofbud02 & "' "
        Dim MonthSales As Decimal = oCommand2.ExecuteScalar()
        Return MonthSales
    End Function
    Private Function GetYearSalesQuantity(ByVal oga03 As String)
        oCommand2.CommandText = "select nvl(sum(ogb12),0) from oga_file,ogb_file where oga01 =  ogb01 and ogapost = 'Y' AND ogb04 <> 'AC0000000000' and oga02 between to_date('"
        oCommand2.CommandText += Date3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand2.CommandText += Date1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and oga03 = '"
        oCommand2.CommandText += oga03 & "' "
        If Not String.IsNullOrEmpty(ogb04) Then
            oCommand2.CommandText += "AND ogb04 LIKE '%" & ogb04 & "%' "
        End If
        Dim YearSales As Decimal = oCommand2.ExecuteScalar()
        Return YearSales
    End Function
    Private Function GetMonthSalesQuantity(ByVal oga03 As String)
        oCommand2.CommandText = "select nvl(sum(ogb12),0) from oga_file,ogb_file where oga01 =  ogb01 and ogapost = 'Y' AND ogb04 <> 'AC0000000000' and oga02 between to_date('"
        oCommand2.CommandText += Date2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand2.CommandText += Date1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and oga03 = '"
        oCommand2.CommandText += oga03 & "' "
        If Not String.IsNullOrEmpty(ogb04) Then
            oCommand2.CommandText += "AND ogb04 LIKE '%" & ogb04 & "%' "
        End If
        Dim MonthSales As Decimal = oCommand2.ExecuteScalar()
        Return MonthSales
    End Function
    Private Function GetYearSalesQuantity(ByVal oga032 As String, ByVal ogb04 As String)
        oCommand2.CommandText = "select nvl(sum(ogb12),0) from oga_file,ogb_file where oga01 =  ogb01 and ogapost = 'Y' and oga02 between to_date('"
        oCommand2.CommandText += Date3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand2.CommandText += Date1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and oga032 = '"
        oCommand2.CommandText += oga032 & "' AND ogb04 = '" & ogb04 & "' "
        Dim YearSales As Decimal = oCommand2.ExecuteScalar()
        Return YearSales
    End Function
    Private Function GetMonthSalesQuantity(ByVal oga032 As String, ByVal ogb04 As String)
        oCommand2.CommandText = "select nvl(sum(ogb12),0) from oga_file,ogb_file where oga01 =  ogb01 and ogapost = 'Y' and oga02 between to_date('"
        oCommand2.CommandText += Date2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand2.CommandText += Date1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and oga032 = '"
        oCommand2.CommandText += oga032 & "' AND ogb04 = '" & ogb04 & "' "
        Dim MonthSales As Decimal = oCommand2.ExecuteScalar()
        Return MonthSales
    End Function
End Class