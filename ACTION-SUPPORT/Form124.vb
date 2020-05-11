Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.XlChartType
Imports Microsoft.Office.Core.MsoChartElementType
Imports Microsoft.Office.Core.MsoTriState


Public Class Form124
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand3 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim LineS1 As Int16 = 0
    Dim tYear As Int16 = 0
    Dim pYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim pMonth As Int16 = 0
    Dim lYear As Int16 = 0
    Dim tCurrency As String = String.Empty
    Dim ExchangeRate As Decimal = 0
    Dim ExchangeRate1 As Decimal = 0
    Dim gDataBase As String = String.Empty
    Dim lMonth As Int16 = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form124_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        

        If Today.Month < 10 Then
            TextBox1.Text = Today.Year & "0" & Today.Month
        Else
            TextBox1.Text = Today.Year & Today.Month
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        gDataBase = Me.ComboBox1.SelectedItem.ToString()
        Select Case gDataBase
            Case "DAC"
                oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
            Case "HAC"
                oConnection.ConnectionString = Module1.OpenOracleConnection("hkacttest")
        End Select
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
        tYear = Strings.Left(Me.TextBox1.Text, 4)
        pYear = tYear - 1
        tMonth = Strings.Right(Me.TextBox1.Text, 2)
        pMonth = tMonth - 1
        If pMonth = 0 Then
            pMonth = 12
            lYear = tYear - 1
        Else
            lYear = tYear
        End If
        BackgroundWorker1.RunWorkerAsync()
        'ExportToExcel()
        'SaveExcel()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        AdjustExcelFormat1()
        If gDataBase = "DAC" Then
            oCommand.CommandText = "select tqa02,sum(t1) as t1 from ( "
            oCommand.CommandText += "select tqa02,round(sum(ogb14 * oga24 /azj041),2) as t1 from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oga02) || (case when length(month(oga02))= 1 then '0' end) || month(oga02) "
            oCommand.CommandText += "where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oga02) = " & tYear & " and month(oga02) = " & tMonth & " group by tqa02 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,round(sum(ohb14 * oha24 * -1 /azj041),2) as t1 from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oha02) || (case when length(month(oha02))= 1 then '0' end) || month(oha02) "
            oCommand.CommandText += "where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oha02) = " & tYear & " and month(oha02) = " & tMonth & " group by tqa02 ) group by tqa02 order by t1 desc"
        Else
            oCommand.CommandText = "select tqa02,sum(t1) as t1 from ( "
            oCommand.CommandText += "select tqa02,round(sum(ogb14 * oga24),2) as t1 from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oga02) = " & tYear & " and month(oga02) = " & tMonth & " group by tqa02 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,round(sum(ohb14 * oha24 * -1),2) as t1 from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oha02) = " & tYear & " and month(oha02) = " & tMonth & " group by tqa02 ) group by tqa02 order by t1 desc"
        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            Dim TT As Int16 = 0
            While oReader.Read()
                TT += 1
                If TT > 10 Then
                    Continue While
                End If
                Ws.Cells(LineZ, 2) = oReader.Item("tqa02")
                Ws.Cells(LineZ, 3) = oReader.Item("t1")
                Ws.Cells(LineZ, 4) = GetBudget1(oReader.Item("tqa02"))
                LineZ += 1
            End While
        End If
        oReader.Close()
        LineZ = 4
        If gDataBase = "DAC" Then
            oCommand.CommandText = "select tqa02,sum(t1) as t1 from ( "
            oCommand.CommandText += "select tqa02,round(sum(ogb14 * oga24 /azj041),2) as t1 from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oga02) || (case when length(month(oga02))= 1 then '0' end) || month(oga02) "
            oCommand.CommandText += "where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oga02) = " & tYear & " and month(oga02) <= " & tMonth & " group by tqa02 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,round(sum(ohb14 * oha24 * -1 /azj041),2) as t1 from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oha02) || (case when length(month(oha02))= 1 then '0' end) || month(oha02) "
            oCommand.CommandText += "where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oha02) = " & tYear & " and month(oha02) <= " & tMonth & " group by tqa02 ) group by tqa02 order by t1 desc"
        Else
            oCommand.CommandText = "select tqa02,sum(t1) as t1 from ( "
            oCommand.CommandText += "select tqa02,round(sum(ogb14 * oga24),2) as t1 from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oga02) = " & tYear & " and month(oga02) <= " & tMonth & " group by tqa02 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,round(sum(ohb14 * oha24 * -1),2) as t1 from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oha02) = " & tYear & " and month(oha02) <= " & tMonth & " group by tqa02 ) group by tqa02 order by t1 desc"
        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            Dim TT As Int16 = 0
            While oReader.Read()
                TT += 1
                If TT > 10 Then
                    Continue While
                End If
                Ws.Cells(LineZ, 6) = oReader.Item("tqa02")
                Ws.Cells(LineZ, 7) = oReader.Item("t1")
                If Not IsDBNull(oReader.Item("tqa02")) Then
                    Ws.Cells(LineZ, 8) = GetBudget2(oReader.Item("tqa02"))
                End If
                LineZ += 1
            End While
        End If
        oReader.Close()

        ' 第二頁
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustExcelFormat2()
        Ws.Name = "2.Sales Report 02（Month）"
        If gDataBase = "DAC" Then
            oCommand.CommandText = "select tqa02,ogb04,ima02,sum(t1) as t1,sum(t2) as t2 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,nvl(sum(ogb12),0) as t1,nvl(round(sum(ogb14 * oga24 / azj041),2),0) as t2  from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oga02) || (case when length(month(oga02))= 1 then '0' end) || month(oga02) "
            oCommand.CommandText += "where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oga02) = " & tYear & " and month(oga02) = " & tMonth & " group by tqa02,ogb04,ima02 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,ohb04,ima02,nvl(sum(ohb12 * -1),0) as t1,nvl(round(sum(ohb14 * oha24 * -1 / azj041),2),0) as t2  from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oha02) || (case when length(month(oha02))= 1 then '0' end) || month(oha02) "
            oCommand.CommandText += "where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oha02) = " & tYear & " and month(oha02) = " & tMonth & " group by tqa02,ohb04,ima02 ) group by tqa02,ogb04,ima02 order by t2 desc "
        Else
            oCommand.CommandText = "select tqa02,ogb04,ima02,sum(t1) as t1,sum(t2) as t2 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,nvl(sum(ogb12),0) as t1,nvl(round(sum(ogb14 * oga24),2),0) as t2  from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oga02) = " & tYear & " and month(oga02) = " & tMonth & " group by tqa02,ogb04,ima02 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,ohb04,ima02,nvl(sum(ohb12 * -1),0) as t1,nvl(round(sum(ohb14 * oha24 * -1),2),0) as t2  from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oha02) = " & tYear & " and month(oha02) = " & tMonth & " group by tqa02,ohb04,ima02 ) group by tqa02,ogb04,ima02 order by t2 desc "
        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            Dim TT As Int16 = 0
            While oReader.Read()
                TT += 1
                If TT > 20 Then
                    Continue While
                End If
                Ws.Cells(LineZ, 2) = oReader.Item("tqa02")
                Ws.Cells(LineZ, 3) = oReader.Item("ogb04")
                Ws.Cells(LineZ, 4) = oReader.Item("ima02")
                Ws.Cells(LineZ, 5) = oReader.Item("t1")
                Ws.Cells(LineZ, 6) = GetBudget3(oReader.Item("ogb04"))
                Ws.Cells(LineZ, 7) = oReader.Item("t2")
                LineZ += 1
            End While
        End If
        oReader.Close()

        ' 第三頁
        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        AdjustExcelFormat2()
        Ws.Name = "3.Sales Report 02 (Year)"
        Ws.Cells(2, 2) = "Output YTD " & tYear
        If gDataBase = "DAC" Then
            oCommand.CommandText = "select tqa02,ogb04,ima02,sum(t1) as t1,sum(t2) as t2 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,nvl(sum(ogb12),0) as t1,nvl(round(sum(ogb14 * oga24 / azj041),2),0) as t2  from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oga02) || (case when length(month(oga02))= 1 then '0' end) || month(oga02) "
            oCommand.CommandText += "where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oga02) = " & tYear & " and month(oga02) <= " & tMonth & " group by tqa02,ogb04,ima02 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,ohb04,ima02,nvl(sum(ohb12 * -1),0) as t1,nvl(round(sum(ohb14 * oha24 * -1 / azj041),2),0) as t2  from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oha02) || (case when length(month(oha02))= 1 then '0' end) || month(oha02) "
            oCommand.CommandText += "where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oha02) = " & tYear & " and month(oha02) <= " & tMonth & " group by tqa02,ohb04,ima02 ) group by tqa02,ogb04,ima02 order by t2 desc "
        Else
            oCommand.CommandText = "select tqa02,ogb04,ima02,sum(t1) as t1,sum(t2) as t2 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,nvl(sum(ogb12),0) as t1,nvl(round(sum(ogb14 * oga24),2),0) as t2  from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oga02) = " & tYear & " and month(oga02) <= " & tMonth & " group by tqa02,ogb04,ima02 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,ohb04,ima02,nvl(sum(ohb12 * -1),0) as t1,nvl(round(sum(ohb14 * oha24 * -1),2),0) as t2  from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oha02) = " & tYear & " and month(oha02) <= " & tMonth & " group by tqa02,ohb04,ima02 ) group by tqa02,ogb04,ima02 order by t2 desc "
        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            Dim TT As Int16 = 0
            While oReader.Read()
                TT += 1
                If TT > 20 Then
                    Continue While
                End If
                Ws.Cells(LineZ, 2) = oReader.Item("tqa02")
                Ws.Cells(LineZ, 3) = oReader.Item("ogb04")
                Ws.Cells(LineZ, 4) = oReader.Item("ima02")
                Ws.Cells(LineZ, 5) = oReader.Item("t1")
                Ws.Cells(LineZ, 6) = GetBudget4(oReader.Item("ogb04"))
                Ws.Cells(LineZ, 7) = oReader.Item("t2")
                LineZ += 1
            End While
        End If
        oReader.Close()
        ' 第四頁

        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(4)
        Ws.Activate()
        AdjustExcelFormat3()
        If gDataBase = "DAC" Then
            oCommand.CommandText = "select tqa02,t1,t2,case when t1 <> 0 then round(((t1-t2)/t1),4) else 0 end as t3 from ( select tqa02,sum(t1) as t1,sum(t2) as t2 from ( "
            oCommand.CommandText += "select tqa02,round(sum(ogb14 * oga24 /azj041),2) as t1,round(sum(ogb12 * ccc23 /azj041),2) as t2 from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oga02) || (case when length(month(oga02))= 1 then '0' end) || month(oga02) "
            oCommand.CommandText += "left join ccc_file on ogb04 = ccc01 and year(oga02) = ccc02 and month(oga02) = ccc03 "
            oCommand.CommandText += "where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oga02) = " & tYear & " and month(oga02) = " & tMonth & " group by tqa02 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,round(sum(ohb14 * oha24 * -1 /azj041),2) as t1,round(sum(ohb12 * ccc23 * -1 /azj041),2) as t2 from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oha02) || (case when length(month(oha02))= 1 then '0' end) || month(oha02) "
            oCommand.CommandText += "left join ccc_file on ohb04 = ccc01 and year(oha02) = ccc02 and month(oha02) = ccc03 "
            oCommand.CommandText += "where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oha02) = " & tYear & " and month(oha02) = " & tMonth & " group by tqa02 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,0,0 from tc_bud_file left join ima_file on tc_bud04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where tc_bud01 = 1 and tc_bud02 = " & tYear & " and tc_bud03 = " & tMonth & " and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' ) group by tqa02 order by t2 desc ) order by t3 desc"
        Else
            oCommand.CommandText = "select tqa02,t1,t2,case when t1 <> 0 then round(((t1-t2)/t1),4) else 0 end as t3 from ( select tqa02,sum(t1) as t1,sum(t2) as t2 from ( "
            oCommand.CommandText += "select tqa02,round(sum(ogb14 * oga24),2) as t1,round(sum(ogb12 * ccc23),2) as t2 from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "left join ccc_file on ogb04 = ccc01 and year(oga02) = ccc02 and month(oga02) = ccc03 "
            oCommand.CommandText += "where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oga02) = " & tYear & " and month(oga02) = " & tMonth & " group by tqa02 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,round(sum(ohb14 * oha24 * -1),2) as t1,round(sum(ohb12 * ccc23 * -1),2) as t2 from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "left join ccc_file on ohb04 = ccc01 and year(oha02) = ccc02 and month(oha02) = ccc03 "
            oCommand.CommandText += "where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oha02) = " & tYear & " and month(oha02) = " & tMonth & " group by tqa02 ) "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,0,0 from tc_bud_file left join ima_file on tc_bud04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where tc_bud01 = 1 and tc_bud02 = " & tYear & " and tc_bud03 = " & tMonth & " and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' ) group by tqa02 order by t2 desc ) order by t3 desc "

        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("tqa02")
                Ws.Cells(LineZ, 3) = oReader.Item("t1")
                Ws.Cells(LineZ, 4) = GetBudget1(oReader.Item("tqa02"))
                Ws.Cells(LineZ, 5) = oReader.Item("t2")
                Ws.Cells(LineZ, 6) = oReader.Item("t3")
                LineZ += 1
            End While
            '加總
            Ws.Cells(LineZ, 2) = "Total"
            Ws.Cells(LineZ, 3) = "=SUM(C4:C" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 4) = "=SUM(D4:D" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 5) = "=SUM(E4:E" & LineZ - 1 & ")"
            '格式
            oRng = Ws.Range("C4", Ws.Cells(LineZ, 5))
            oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
            oRng = Ws.Range("F4", Ws.Cells(LineZ - 1, 6))
            oRng.NumberFormat = "0.00%"
            ' 添加 負數為紅色 20180531
            oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
            oRng.FormatConditions(1).FONT.COLOR = Color.Red
            ' 劃線
            oRng = Ws.Range("B4", Ws.Cells(LineZ, 6))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        oReader.Close()

        LineZ = 4
        If gDataBase = "DAC" Then
            oCommand.CommandText = "select tqa02,t1,t2,case when t1 <> 0 then round(((t1-t2)/t1),4) else 0 end as t3 from ( select tqa02,sum(t1) as t1,sum(t2) as t2 from ( "
            oCommand.CommandText += "select tqa02,round(sum(ogb14 * oga24 /azj041),2) as t1,round(sum(ogb12 * ccc23 /azj041),2) as t2 from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oga02) || (case when length(month(oga02))= 1 then '0' end) || month(oga02) "
            oCommand.CommandText += "left join ccc_file on ogb04 = ccc01 and year(oga02) = ccc02 and month(oga02) = ccc03 "
            oCommand.CommandText += "where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oga02) = " & tYear & " and month(oga02) <= " & tMonth & " group by tqa02 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,round(sum(ohb14 * oha24 * -1 /azj041),2) as t1,round(sum(ohb12 * ccc23 * -1 /azj041),2) as t2 from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oha02) || (case when length(month(oha02))= 1 then '0' end) || month(oha02) "
            oCommand.CommandText += "left join ccc_file on ohb04 = ccc01 and year(oha02) = ccc02 and month(oha02) = ccc03 "
            oCommand.CommandText += "where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oha02) = " & tYear & " and month(oha02) <= " & tMonth & " group by tqa02 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,0,0 from tc_bud_file left join ima_file on tc_bud04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where tc_bud01 = 1 and tc_bud02 = " & tYear & " and tc_bud03 <= " & tMonth & " and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' ) group by tqa02 order by t2 desc ) order by t3 desc"

        Else
            oCommand.CommandText = "select tqa02,t1,t2,case when t1 <> 0 then round(((t1-t2)/t1),4) else 0 end as t3 from ( select tqa02,sum(t1) as t1,sum(t2) as t2 from ( "
            oCommand.CommandText += "select tqa02,round(sum(ogb14 * oga24),2) as t1,round(sum(ogb12 * ccc23),2) as t2 from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "left join ccc_file on ogb04 = ccc01 and year(oga02) = ccc02 and month(oga02) = ccc03 "
            oCommand.CommandText += "where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oga02) = " & tYear & " and month(oga02) <= " & tMonth & " group by tqa02 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,round(sum(ohb14 * oha24 * -1),2) as t1,round(sum(ohb12 * ccc23 * -1),2) as t2 from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "left join ccc_file on ohb04 = ccc01 and year(oha02) = ccc02 and month(oha02) = ccc03 "
            oCommand.CommandText += "where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 <> 'AC9999' and year(oha02) = " & tYear & " and month(oha02) <= " & tMonth & " group by tqa02 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,0,0 from tc_bud_file left join ima_file on tc_bud04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where tc_bud01 = 1 and tc_bud02 = " & tYear & " and tc_bud03 <= " & tMonth & " and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' ) group by tqa02 order by t2 desc ) order by t3 desc"
        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 8) = oReader.Item("tqa02")
                Ws.Cells(LineZ, 9) = oReader.Item("t1")
                If Not IsDBNull(oReader.Item("tqa02")) Then
                    Ws.Cells(LineZ, 10) = GetBudget2(oReader.Item("tqa02"))
                End If
                Ws.Cells(LineZ, 11) = oReader.Item("t2")
                Ws.Cells(LineZ, 12) = oReader.Item("t3")
                LineZ += 1
            End While
            '加總
            Ws.Cells(LineZ, 8) = "Total"
            Ws.Cells(LineZ, 9) = "=SUM(I4:I" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 10) = "=SUM(J4:J" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 11) = "=SUM(K4:K" & LineZ - 1 & ")"
            '格式
            oRng = Ws.Range("I4", Ws.Cells(LineZ, 11))
            oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
            oRng = Ws.Range("L4", Ws.Cells(LineZ - 1, 12))
            oRng.NumberFormat = "0.00%"
            oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
            oRng.FormatConditions(1).FONT.COLOR = Color.Red
            ' 劃線
            oRng = Ws.Range("H4", Ws.Cells(LineZ, 12))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        oReader.Close()

        ' 第五頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(5)
        Ws.Activate()
        AdjustExcelFormat4()
        Ws.Name = "6.STD GM by Mon"
        Ws.Cells(2, 6) = "Output of " & tYear & "/" & tMonth
        If gDataBase = "DAC" Then
            oCommand.CommandText = "select tqa02,ogb04,ima02,ima31,t1,t2,t3,t4,case when t1 <> 0 then round(((t4-t2)/t4),4) else 0 end as t5 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,ima31,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,ima31,sum(ogb12) as t1,round(sum(ogb12 * (stb07+stb08+stb09+stb09a) /azj041),2) as t2,0 as t3,round(sum(ogb14 * oga24 /azj041),2) as t4 from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oga02) || (case when length(month(oga02))= 1 then '0' end) || month(oga02) "
            oCommand.CommandText += "left join stb_file on ogb04 = stb01 and year(oga02) = stb02 and month(oga02) = stb03 where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oga02) = " & tYear & " and month(oga02) = " & tMonth & " group by tqa02,ogb04,ima02,ima31 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,ohb04,ima02,ima31,sum(ohb12 * -1) as t1,round(sum(ohb12 * (stb07+stb08+stb09+stb09a) * -1 /azj041),2) as t2,0 as t3,round(sum(ohb14 * oha24 * -1 /azj041),2) as t4 from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oha02) || (case when length(month(oha02))= 1 then '0' end) || month(oha02) "
            oCommand.CommandText += "left join stb_file on ohb04 = stb01 and year(oha02) = stb02 and month(oha02) = stb03 where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oha02) = " & tYear & " and month(oha02) = " & tMonth & " group by tqa02,ohb04,ima02,ima31 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,tc_bud04,ima02,ima31,0,0,case when tc_bud14 = 'EUR' THEN tc_bud13 *1.2 else tc_bud13 end,0 from tc_bud_file left join ima_file on tc_bud04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where tc_bud01 = 1 and tc_bud02 = " & tYear & " and tc_bud03 = " & tMonth & " and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' ) group by tqa02,ogb04,ima02,ima31 ) order by t5 desc"
        Else
            oCommand.CommandText = "select tqa02,ogb04,ima02,ima31,t1,t2,t3,t4,case when t1 <> 0 then round(((t4-t2)/t4),4) else 0 end as t5 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,ima31,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,ima31,sum(ogb12) as t1,round(sum(ogb12 * (stb07+stb08+stb09+stb09a)),2) as t2,0 as t3,round(sum(ogb14 * oga24),2) as t4 from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "left join stb_file on ogb04 = stb01 and year(oga02) = stb02 and month(oga02) = stb03 where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oga02) = " & tYear & " and month(oga02) = " & tMonth & " group by tqa02,ogb04,ima02,ima31 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,ohb04,ima02,ima31,sum(ohb12 * -1) as t1,round(sum(ohb12 * (stb07+stb08+stb09+stb09a) * -1 ),2) as t2,0 as t3,round(sum(ohb14 * oha24 * -1 ),2) as t4 from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "left join stb_file on ohb04 = stb01 and year(oha02) = stb02 and month(oha02) = stb03 where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oha02) = " & tYear & " and month(oha02) = " & tMonth & " group by tqa02,ohb04,ima02,ima31 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,tc_bud04,ima02,ima31,0,0,case when tc_bud14 = 'EUR' THEN tc_bud13 *1.2 else tc_bud13 end,0 from tc_bud_file left join ima_file on tc_bud04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where tc_bud01 = 1 and tc_bud02 = " & tYear & " and tc_bud03 = " & tMonth & " and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' ) group by tqa02,ogb04,ima02,ima31 ) order by t5 desc"
        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("tqa02")
                Ws.Cells(LineZ, 3) = oReader.Item("ogb04")
                Ws.Cells(LineZ, 4) = oReader.Item("ima02")
                Ws.Cells(LineZ, 5) = oReader.Item("ima31")
                Ws.Cells(LineZ, 6) = oReader.Item("t1")
                Ws.Cells(LineZ, 7) = oReader.Item("t2")
                Ws.Cells(LineZ, 8) = oReader.Item("t3")
                Ws.Cells(LineZ, 9) = oReader.Item("t4")
                Ws.Cells(LineZ, 10) = oReader.Item("t5")
                LineZ += 1
            End While
            '加總
            Ws.Cells(LineZ, 2) = "Total"
            Ws.Cells(LineZ, 6) = "=SUM(F4:F" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 7) = "=SUM(G4:G" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 8) = "=SUM(H4:H" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 9) = "=SUM(I4:I" & LineZ - 1 & ")"
            '格式
            oRng = Ws.Range("F4", Ws.Cells(LineZ, 9))
            oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
            oRng = Ws.Range("J4", Ws.Cells(LineZ - 1, 10))
            oRng.NumberFormat = "0.00%"
            ' 添加 負數為紅色 20180531
            oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
            oRng.FormatConditions(1).FONT.COLOR = Color.Red
            ' 劃線
            oRng = Ws.Range("B4", Ws.Cells(LineZ, 10))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        oReader.Close()

        '新增第六頁  20180606

        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(6)
        Ws.Activate()
        AdjustExcelFormat4()
        Ws.Name = "7.Actl GM by Mon"
        Ws.Cells(2, 6) = "Output of " & tYear & "/" & tMonth
        Ws.Cells(3, 7) = "Output at Actl cost"
        If gDataBase = "DAC" Then
            oCommand.CommandText = "select tqa02,ogb04,ima02,ima31,t1,t2,t3,t4,case when t1 <> 0 then round(((t4-t2)/t4),4) else 0 end as t5 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,ima31,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,ima31,sum(ogb12) as t1,round(sum(ogb12 * ccc23 /azj041),2) as t2,0 as t3,round(sum(ogb14 * oga24 /azj041),2) as t4 from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oga02) || (case when length(month(oga02))= 1 then '0' end) || month(oga02) "
            oCommand.CommandText += "left join ccc_file on ogb04 = ccc01 and year(oga02) = ccc02 and month(oga02) = ccc03 where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oga02) = " & tYear & " and month(oga02) = " & tMonth & " group by tqa02,ogb04,ima02,ima31 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,ohb04,ima02,ima31,sum(ohb12 * -1) as t1,round(sum(ohb12 * ccc23 * -1 /azj041),2) as t2,0 as t3,round(sum(ohb14 * oha24 * -1 /azj041),2) as t4 from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oha02) || (case when length(month(oha02))= 1 then '0' end) || month(oha02) "
            oCommand.CommandText += "left join ccc_file on ohb04 = ccc01 and year(oha02) = ccc02 and month(oha02) = ccc03 where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oha02) = " & tYear & " and month(oha02) = " & tMonth & " group by tqa02,ohb04,ima02,ima31 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,tc_bud04,ima02,ima31,0,0,case when tc_bud14 = 'EUR' THEN tc_bud13 *1.2 else tc_bud13 end,0 from tc_bud_file left join ima_file on tc_bud04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where tc_bud01 = 1 and tc_bud02 = " & tYear & " and tc_bud03 = " & tMonth & " and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' ) group by tqa02,ogb04,ima02,ima31 ) order by t5 desc"
        Else
            oCommand.CommandText = "select tqa02,ogb04,ima02,ima31,t1,t2,t3,t4,case when t1 <> 0 then round(((t4-t2)/t4),4) else 0 end as t5 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,ima31,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,ima31,sum(ogb12) as t1,round(sum(ogb12 * ccc23),2) as t2,0 as t3,round(sum(ogb14 * oga24),2) as t4 from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "left join ccc_file on ogb04 = ccc01 and year(oga02) = ccc02 and month(oga02) = ccc03 where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oga02) = " & tYear & " and month(oga02) = " & tMonth & " group by tqa02,ogb04,ima02,ima31 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,ohb04,ima02,ima31,sum(ohb12 * -1) as t1,round(sum(ohb12 * ccc23 * -1 ),2) as t2,0 as t3,round(sum(ohb14 * oha24 * -1 ),2) as t4 from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "left join ccc_file on ohb04 = ccc01 and year(oha02) = ccc02 and month(oha02) = ccc03 where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oha02) = " & tYear & " and month(oha02) = " & tMonth & " group by tqa02,ohb04,ima02,ima31 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,tc_bud04,ima02,ima31,0,0,case when tc_bud14 = 'EUR' THEN tc_bud13 *1.2 else tc_bud13 end,0 from tc_bud_file left join ima_file on tc_bud04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where tc_bud01 = 1 and tc_bud02 = " & tYear & " and tc_bud03 = " & tMonth & " and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' ) group by tqa02,ogb04,ima02,ima31 ) order by t5 desc"
        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("tqa02")
                Ws.Cells(LineZ, 3) = oReader.Item("ogb04")
                Ws.Cells(LineZ, 4) = oReader.Item("ima02")
                Ws.Cells(LineZ, 5) = oReader.Item("ima31")
                Ws.Cells(LineZ, 6) = oReader.Item("t1")
                Ws.Cells(LineZ, 7) = oReader.Item("t2")
                Ws.Cells(LineZ, 8) = oReader.Item("t3")
                Ws.Cells(LineZ, 9) = oReader.Item("t4")
                Ws.Cells(LineZ, 10) = oReader.Item("t5")
                LineZ += 1
            End While
            '加總
            Ws.Cells(LineZ, 2) = "Total"
            Ws.Cells(LineZ, 6) = "=SUM(F4:F" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 7) = "=SUM(G4:G" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 8) = "=SUM(H4:H" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 9) = "=SUM(I4:I" & LineZ - 1 & ")"
            '格式
            oRng = Ws.Range("F4", Ws.Cells(LineZ, 9))
            oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
            oRng = Ws.Range("J4", Ws.Cells(LineZ - 1, 10))
            oRng.NumberFormat = "0.00%"
            ' 添加 負數為紅色 20180531
            oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
            oRng.FormatConditions(1).FONT.COLOR = Color.Red
            ' 劃線
            oRng = Ws.Range("B4", Ws.Cells(LineZ, 10))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        oReader.Close()



        ' 第六頁 改第七頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(7)
        Ws.Activate()
        AdjustExcelFormat4()
        Ws.Name = "8.STD GM by Year"
        Ws.Cells(2, 6) = "Output YTD " & tYear
        If gDataBase = "DAC" Then
            oCommand.CommandText = "select tqa02,ogb04,ima02,ima31,t1,t2,t3,t4,case when t1 <> 0 then round(((t4-t2)/t4),4) else 0 end as t5 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,ima31,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,ima31,sum(ogb12) as t1,round(sum(ogb12 * (stb07+stb08+stb09+stb09a) /azj041),2) as t2,0 as t3,round(sum(ogb14 * oga24 /azj041),2) as t4 from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oga02) || (case when length(month(oga02))= 1 then '0' end) || month(oga02) "
            oCommand.CommandText += "left join stb_file on ogb04 = stb01 and year(oga02) = stb02 and month(oga02) = stb03 where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oga02) = " & tYear & " and month(oga02) <= " & tMonth & " group by tqa02,ogb04,ima02,ima31 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,ohb04,ima02,ima31,sum(ohb12 * -1) as t1,round(sum(ohb12 * (stb07+stb08+stb09+stb09a) * -1 /azj041),2) as t2,0 as t3,round(sum(ohb14 * oha24 * -1 /azj041),2) as t4 from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oha02) || (case when length(month(oha02))= 1 then '0' end) || month(oha02) "
            oCommand.CommandText += "left join stb_file on ohb04 = stb01 and year(oha02) = stb02 and month(oha02) = stb03 where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oha02) = " & tYear & " and month(oha02) <= " & tMonth & " group by tqa02,ohb04,ima02,ima31 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,tc_bud04,ima02,ima31,0,0,case when tc_bud14 = 'EUR' THEN tc_bud13 *1.2 else tc_bud13 end,0 from tc_bud_file left join ima_file on tc_bud04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where tc_bud01 = 1 and tc_bud02 = " & tYear & " and tc_bud03 <= " & tMonth & " and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' ) group by tqa02,ogb04,ima02,ima31 ) order by t5 desc"
        Else
            oCommand.CommandText = "select tqa02,ogb04,ima02,ima31,t1,t2,t3,t4,case when t1 <> 0 then round(((t4-t2)/t4),4) else 0 end as t5 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,ima31,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,ima31,sum(ogb12) as t1,round(sum(ogb12 * (stb07+stb08+stb09+stb09a)),2) as t2,0 as t3,round(sum(ogb14 * oga24),2) as t4 from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "left join stb_file on ogb04 = stb01 and year(oga02) = stb02 and month(oga02) = stb03 where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oga02) = " & tYear & " and month(oga02) <= " & tMonth & " group by tqa02,ogb04,ima02,ima31 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,ohb04,ima02,ima31,sum(ohb12 * -1) as t1,round(sum(ohb12 * (stb07+stb08+stb09+stb09a) * -1 ),2) as t2,0 as t3,round(sum(ohb14 * oha24 * -1 ),2) as t4 from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "left join stb_file on ohb04 = stb01 and year(oha02) = stb02 and month(oha02) = stb03 where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oha02) = " & tYear & " and month(oha02) <= " & tMonth & " group by tqa02,ohb04,ima02,ima31 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,tc_bud04,ima02,ima31,0,0,case when tc_bud14 = 'EUR' THEN tc_bud13 *1.2 else tc_bud13 end,0 from tc_bud_file left join ima_file on tc_bud04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where tc_bud01 = 1 and tc_bud02 = " & tYear & " and tc_bud03 <= " & tMonth & " and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' ) group by tqa02,ogb04,ima02,ima31 ) order by t5 desc"
        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("tqa02")
                Ws.Cells(LineZ, 3) = oReader.Item("ogb04")
                Ws.Cells(LineZ, 4) = oReader.Item("ima02")
                Ws.Cells(LineZ, 5) = oReader.Item("ima31")
                Ws.Cells(LineZ, 6) = oReader.Item("t1")
                Ws.Cells(LineZ, 7) = oReader.Item("t2")
                Ws.Cells(LineZ, 8) = oReader.Item("t3")
                Ws.Cells(LineZ, 9) = oReader.Item("t4")
                Ws.Cells(LineZ, 10) = oReader.Item("t5")
                LineZ += 1
            End While
            '加總
            Ws.Cells(LineZ, 2) = "Total"
            Ws.Cells(LineZ, 6) = "=SUM(F4:F" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 7) = "=SUM(G4:G" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 8) = "=SUM(H4:H" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 9) = "=SUM(I4:I" & LineZ - 1 & ")"
            '格式
            oRng = Ws.Range("F4", Ws.Cells(LineZ, 9))
            oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
            oRng = Ws.Range("J4", Ws.Cells(LineZ - 1, 10))
            oRng.NumberFormat = "0.00%"
            ' 添加 負數為紅色 20180531
            oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
            oRng.FormatConditions(1).FONT.COLOR = Color.Red
            ' 劃線
            oRng = Ws.Range("B4", Ws.Cells(LineZ, 10))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        oReader.Close()

        '第八頁 20180606
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(8)
        Ws.Activate()
        AdjustExcelFormat4()
        Ws.Name = "9.Actl GM by Year"
        Ws.Cells(2, 6) = "Output YTD " & tYear
        Ws.Cells(3, 7) = "Output at Actl cost"
        If gDataBase = "DAC" Then
            oCommand.CommandText = "select tqa02,ogb04,ima02,ima31,t1,t2,t3,t4,case when t1 <> 0 then round(((t4-t2)/t4),4) else 0 end as t5 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,ima31,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,ima31,sum(ogb12) as t1,round(sum(ogb12 * ccc23 /azj041),2) as t2,0 as t3,round(sum(ogb14 * oga24 /azj041),2) as t4 from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oga02) || (case when length(month(oga02))= 1 then '0' end) || month(oga02) "
            oCommand.CommandText += "left join ccc_file on ogb04 = ccc01 and year(oga02) = ccc02 and month(oga02) = ccc03 where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oga02) = " & tYear & " and month(oga02) <= " & tMonth & " group by tqa02,ogb04,ima02,ima31 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,ohb04,ima02,ima31,sum(ohb12 * -1) as t1,round(sum(ohb12 * ccc23 * -1 /azj041),2) as t2,0 as t3,round(sum(ohb14 * oha24 * -1 /azj041),2) as t4 from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' left join azj_file on azj01 = 'USD' and azj02 = year(oha02) || (case when length(month(oha02))= 1 then '0' end) || month(oha02) "
            oCommand.CommandText += "left join ccc_file on ohb04 = ccc01 and year(oha02) = ccc02 and month(oha02) = ccc03 where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oha02) = " & tYear & " and month(oha02) <= " & tMonth & " group by tqa02,ohb04,ima02,ima31 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,tc_bud04,ima02,ima31,0,0,case when tc_bud14 = 'EUR' THEN tc_bud13 *1.2 else tc_bud13 end,0 from tc_bud_file left join ima_file on tc_bud04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where tc_bud01 = 1 and tc_bud02 = " & tYear & " and tc_bud03 <= " & tMonth & " and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' ) group by tqa02,ogb04,ima02,ima31 ) order by t5 desc"
        Else
            oCommand.CommandText = "select tqa02,ogb04,ima02,ima31,t1,t2,t3,t4,case when t1 <> 0 then round(((t4-t2)/t4),4) else 0 end as t5 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,ima31,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4 from ( "
            oCommand.CommandText += "select tqa02,ogb04,ima02,ima31,sum(ogb12) as t1,round(sum(ogb12 * ccc23),2) as t2,0 as t3,round(sum(ogb14 * oga24),2) as t4 from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "left join ccc_file on ogb04 = ccc01 and year(oga02) = ccc02 and month(oga02) = ccc03 where ogapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oga02) = " & tYear & " and month(oga02) <= " & tMonth & " group by tqa02,ogb04,ima02,ima31 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,ohb04,ima02,ima31,sum(ohb12 * -1) as t1,round(sum(ohb12 * ccc23 * -1 ),2) as t2,0 as t3,round(sum(ohb14 * oha24 * -1 ),2) as t4 from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 "
            oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "left join ccc_file on ohb04 = ccc01 and year(oha02) = ccc02 and month(oha02) = ccc03 where ohapost = 'Y'  and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' and year(oha02) = " & tYear & " and month(oha02) <= " & tMonth & " group by tqa02,ohb04,ima02,ima31 "
            oCommand.CommandText += "union all "
            oCommand.CommandText += "select tqa02,tc_bud04,ima02,ima31,0,0,case when tc_bud14 = 'EUR' THEN tc_bud13 *1.2 else tc_bud13 end,0 from tc_bud_file left join ima_file on tc_bud04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = '2' "
            oCommand.CommandText += "where tc_bud01 = 1 and tc_bud02 = " & tYear & " and tc_bud03 <= " & tMonth & " and ima06 = '103' and ima01 not like 'S%' and ima01 not like 'A%' ) group by tqa02,ogb04,ima02,ima31 ) order by t5 desc"
        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("tqa02")
                Ws.Cells(LineZ, 3) = oReader.Item("ogb04")
                Ws.Cells(LineZ, 4) = oReader.Item("ima02")
                Ws.Cells(LineZ, 5) = oReader.Item("ima31")
                Ws.Cells(LineZ, 6) = oReader.Item("t1")
                Ws.Cells(LineZ, 7) = oReader.Item("t2")
                Ws.Cells(LineZ, 8) = oReader.Item("t3")
                Ws.Cells(LineZ, 9) = oReader.Item("t4")
                Ws.Cells(LineZ, 10) = oReader.Item("t5")
                LineZ += 1
            End While
            '加總
            Ws.Cells(LineZ, 2) = "Total"
            Ws.Cells(LineZ, 6) = "=SUM(F4:F" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 7) = "=SUM(G4:G" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 8) = "=SUM(H4:H" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 9) = "=SUM(I4:I" & LineZ - 1 & ")"
            '格式
            oRng = Ws.Range("F4", Ws.Cells(LineZ, 9))
            oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
            oRng = Ws.Range("J4", Ws.Cells(LineZ - 1, 10))
            oRng.NumberFormat = "0.00%"
            ' 添加 負數為紅色 20180531
            oRng.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
            oRng.FormatConditions(1).FONT.COLOR = Color.Red
            ' 劃線
            oRng = Ws.Range("B4", Ws.Cells(LineZ, 10))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        oReader.Close()

        ' 第七頁  改第九頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(9)
        Ws.Activate()
        AdjustExcelFormat5()
        'Ws.Cells(4, 3) = GetD146103(0, tMonth)
        'Ws.Cells(4, 7) = GetD146103(1, tMonth)
        Ws.Cells(4, 3) = NewGetD146103(0, tMonth)
        Ws.Cells(4, 7) = NewGetD146103(1, tMonth)
        oCommand.CommandText = "select nvl(sum(tc_bud11),0) as t1,nvl(sum(tc_bud11 * stb07),0) as t2,nvl(sum(tc_bud11 * stb08),0) as t3,"
        oCommand.CommandText += "nvl(sum(tc_bud11 * (stb09 + stb09a)),0) as t4 from tc_bud_file,stb_file where tc_bud01 = 1 and tc_bud02 = " & tYear & " and tc_bud03 = " & tMonth & " and tc_bud04 = stb01 and tc_bud02 = stb02 and tc_bud03 = stb03"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            oReader.Read()
            Ws.Cells(4, 4) = oReader.Item("t1")
            Ws.Cells(8, 4) = oReader.Item("t2")
            Ws.Cells(9, 4) = oReader.Item("t3")
            Ws.Cells(10, 4) = oReader.Item("t4")

        End If
        oReader.Close()
        oCommand.CommandText = "select nvl(sum(tc_bud11),0) as t1,nvl(sum(tc_bud11 * stb07),0) as t2,nvl(sum(tc_bud11 * stb08),0) as t3,"
        oCommand.CommandText += "nvl(sum(tc_bud11 * (stb09 + stb09a)),0) as t4 from tc_bud_file,stb_file where tc_bud01 = 1 and tc_bud02 = " & tYear & " and tc_bud03 <= " & tMonth & " and tc_bud04 = stb01 and tc_bud02 = stb02 and tc_bud03 = stb03"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            oReader.Read()
            Ws.Cells(4, 8) = oReader.Item("t1")
            Ws.Cells(8, 8) = oReader.Item("t2")
            Ws.Cells(9, 8) = oReader.Item("t3")
            Ws.Cells(10, 8) = oReader.Item("t4")

        End If
        oReader.Close()
        oCommand.CommandText = "select nvl(sum(ccc22a + ccc28a +ccc26a + ((ccc31 + ccc41 + ccc51 + ccc71) * ccc23a)),0) as t1,"
        oCommand.CommandText += "nvl(sum(ccc22b + ccc28b + ccc26b + ((ccc31 + ccc41 + ccc51 + ccc71) * ccc23b)),0) as t2,nvl(sum(ccc22c + ccc28c + ccc22d + ccc28d + ccc26c + ccc26d + ((ccc31 + ccc41 + ccc51 + ccc71) * (ccc23c+ccc23d))),0) as t3,"
        oCommand.CommandText += "nvl(sum(ccc93),0) as t4 from ccc_file,ima_file where ccc01 = ima01 and ima06 ='103' and ima01 not like 'S%' and ima01 <> 'AC9999' and ccc02 = " & tYear & " and ccc03 = " & tMonth
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            oReader.Read()
            Ws.Cells(8, 3) = oReader.Item("t1")
            Ws.Cells(9, 3) = oReader.Item("t2")
            Ws.Cells(10, 3) = oReader.Item("t3")
            Ws.Cells(11, 3) = oReader.Item("t4")
        End If
        oReader.Close()
        oCommand.CommandText = "select nvl(sum(ccc22a + ccc28a +ccc26a + ((ccc31 + ccc41 + ccc51 + ccc71) * ccc23a)),0) as t1,"
        oCommand.CommandText += "nvl(sum(ccc22b + ccc28b + ccc26b + ((ccc31 + ccc41 + ccc51 + ccc71) * ccc23b)),0) as t2,nvl(sum(ccc22c + ccc28c + ccc22d + ccc28d + ccc26c + ccc26d + ((ccc31 + ccc41 + ccc51 + ccc71) * (ccc23c+ccc23d))),0) as t3,"
        oCommand.CommandText += "nvl(sum(ccc93),0) as t4 from ccc_file,ima_file where ccc01 = ima01 and ima06 ='103' and ima01 not like 'S%' and ima01 <> 'AC9999' and ccc02 = " & tYear & " and ccc03 <= " & tMonth
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            oReader.Read()
            Ws.Cells(8, 7) = oReader.Item("t1")
            Ws.Cells(9, 7) = oReader.Item("t2")
            Ws.Cells(10, 7) = oReader.Item("t3")
            Ws.Cells(11, 7) = oReader.Item("t4")
        End If
        oReader.Close()
        Ws.Cells(13, 3) = GetVoucher(tMonth, 0)
        Ws.Cells(13, 7) = GetVoucher(1, 0)
        Ws.Cells(14, 3) = GetVoucher(tMonth, 1)
        Ws.Cells(14, 7) = GetVoucher(tMonth, 1)


        ' 第八頁 Scrap Report 改第10頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(10)
        Ws.Activate()
        AdjustExcelFormat6()
        For i As Int16 = 1 To tMonth Step 1
            'Ws.Cells(3, 2 + i) = GetD146103(0, i)
            Ws.Cells(3, 2 + i) = NewGetD146103(0, i)
            Ws.Cells(4, 2 + i) = NewGetD146103USD(0, i)
            Ws.Cells(5, 2 + i) = GetTlfCost(i)
            Ws.Cells(6, 2 + i) = GetD146109(i)
            Ws.Cells(7, 2 + i) = GetD146109Cost(i)
            Ws.Cells(8, 2 + i) = GetD146109USD(i)
        Next

        Dim YB As Excel.Chart = Ws.Shapes.AddChart(xlLine, 97, 180, 1265, 250).Chart
        oRng = Ws.Range("B2:N2,B9:N9")
        YB.SetSourceData(oRng, Microsoft.Office.Interop.Excel.XlRowCol.xlRows)
        'YB.SeriesCollection(1).Format.Line.ForeColor.RGB = 
        YB.SeriesCollection(1).Format.Line.Visible = msoTrue
        YB.SeriesCollection(1).Format.Line.Weight = 3
        YB.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        YB.SeriesCollection.NewSeries()
        YB.SeriesCollection(2).Name = "='12.Scrap report'!$B$10"
        YB.SeriesCollection(2).Values = "='12.Scrap report'!$C$10:$N$10"
        YB.SeriesCollection(2).XValues = "='12.Scrap report'!$C$2:$N$2"
        YB.SeriesCollection(2).Format.Line.Visible = msoCTrue
        YB.SeriesCollection(2).Format.Line.Weight = 3.5
        YB.SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(146, 208, 80)
        YB.ApplyLayout(5)
        YB.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Delete()
        YB.ChartTitle.Text = "Scrap %"
        'YB.PlotArea.Width = 1073.07
        'YB.SetElement(msoElementChartTitleAboveChart)
        'XA.ActiveChart.SetElement(msoElementLegendNone)
        'YB.ChartTitle.Text = "Scrap（Sales） %"
        'YB.SeriesCollection(1).XValues = "='29.Scrap report'!$C$2:$N$2"
        'YB.SetElement(msoElementLegendNone)
        'YB.SetElement(msoElementDataTableWithLegendKeys)

        

        ' 第十頁 改第十二頁  再改十一頁  20180621
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(11)
        Ws.Activate()
        AdjustExcelFormat8()
        oCommand.CommandText = "select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "sum(t" & i & ") as t" & i & ","
        Next
        oCommand.CommandText += "1 from ( select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "(case when aah03 = " & i & " then nvl(sum(aah04-aah05),0) else 0 end ) as t" & i & ","
        Next
        oCommand.CommandText += "1 from aah_file,aag_file where aah00 = aag00 and aah01 = aag01 and aah02 = " & tYear & " and aah03 <= " & tMonth & " AND aag08 in ('6401') and aag07 in ('2','3') group by aah03 ) "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 1 To tMonth Step 1
                    Ws.Cells(2, i + 1) = oReader.Item(i - 1)

                Next
            End While
        End If
        oReader.Close()

        oCommand.CommandText = "select nvl(sum(aah05-aah04),0) as t1 from aah_file,aag_file where aah00 = aag00 and aah01 = aag01 and aah02 = " & tYear & " and aah03 = 0 AND aag08 in ('2202') and aag07 in ('2','3')"
        Dim AAB As Decimal = oCommand.ExecuteScalar()
        Ws.Cells(3, 2) = AAB

        oCommand.CommandText = "select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "sum(t" & i & ") as t" & i & ","
        Next
        oCommand.CommandText += "sum(s1) as s1 from ( select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "(case when aah03 = " & i & " then nvl(sum(aah05-aah04),0) else 0 end ) as t" & i & ","
        Next
        oCommand.CommandText += "(case when aah03 = 0 then nvl(sum(aah05-aah04),0) else 0 end ) as s1 from aah_file,aag_file where aah00 = aag00 and aah01 = aag01 and aah02 = " & tYear & " and aah03 <= " & tMonth & " AND aag08 in ('2202') and aag07 in ('2','3') group by aah03 ) "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 1 To tMonth Step 1
                    ' 設定期初
                    Dim Scount As Decimal = 0
                    Scount += oReader.Item("s1")
                    For j As Int16 = 1 To i Step 1
                        Scount += oReader.Item(j - 1)
                    Next
                    Ws.Cells(4, i + 1) = Scount
                Next
            End While
        End If
        oReader.Close()
        ' 第十一頁 改第十三頁  再改十二頁 20180621
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(12)
        Ws.Activate()
        AdjustExcelFormat9()
        oCommand.CommandText = "select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "sum(t" & i & ") as t" & i & ","
        Next
        oCommand.CommandText += "1 from ( select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "(case when aah03 = " & i & " then nvl(sum(aah04-aah05),0) else 0 end ) as t" & i & ","
        Next
        oCommand.CommandText += "1 from aah_file,aag_file where aah00 = aag00 and aah01 = aag01 and aah02 = " & tYear & " and aah03 <= " & tMonth & " AND aag08 in ('6401') and aag07 in ('2','3') group by aah03 ) "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 1 To tMonth Step 1
                    Ws.Cells(2, i + 1) = oReader.Item(i - 1)

                Next
            End While
        End If
        oReader.Close()

        oCommand.CommandText = "select nvl(sum(aah04-aah05),0) as t1 from aah_file,aag_file where aah00 = aag00 and aah01 = aag01 and aah02 = " & tYear & " and aah03 = 0 AND aag01 in ('140301','1404','1405','1406','1407','1408','1409','1410','1412','1413','500101','500102','500103','500104','1471') and aag07 in ('2','3')"
        Dim AAC As Decimal = oCommand.ExecuteScalar()
        Ws.Cells(3, 2) = AAC

        oCommand.CommandText = "select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "sum(t" & i & ") as t" & i & ","
        Next
        oCommand.CommandText += "sum(s1) as s1 from ( select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "(case when aah03 = " & i & " then nvl(sum(aah04-aah05),0) else 0 end ) as t" & i & ","
        Next
        oCommand.CommandText += "(case when aah03 = 0 then nvl(sum(aah04-aah05),0) else 0 end ) as s1 from aah_file,aag_file where aah00 = aag00 and aah01 = aag01 and aah02 = " & tYear & " and aah03 <= " & tMonth & " AND aag01 in ('140301','1404','1405','1406','1407','1408','1409','1410','1412','1413','500101','500102','500103','500104','1471') and aag07 in ('2','3') group by aah03 ) "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 1 To tMonth Step 1
                    ' 設定期初
                    Dim Scount As Decimal = 0
                    Scount += oReader.Item("s1")
                    For j As Int16 = 1 To i Step 1
                        Scount += oReader.Item(j - 1)
                    Next
                    Ws.Cells(4, i + 1) = Scount
                Next
            End While
        End If
        oReader.Close()

        ' 第九頁 AR-Days 改第十一頁 再改第十三頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(13)
        Ws.Activate()
        AdjustExcelFormat7()
        oCommand.CommandText = "select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "sum(t" & i & ") as t" & i & ","
        Next
        oCommand.CommandText += "1 from ( select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "(case when aah03 = " & i & " then nvl(sum(aah05-aah04),0) else 0 end ) as t" & i & ","
        Next
        oCommand.CommandText += "1 from aah_file,aag_file where aah00 = aag00 and aah01 = aag01 and aah02 = " & tYear & " and aah03 <= " & tMonth & " AND aag08 in ('6001','6051','6061') and aag07 in ('2','3') group by aah03 ) "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 1 To tMonth Step 1
                    Ws.Cells(2, i + 1) = oReader.Item(i - 1)

                Next
            End While
        End If
        oReader.Close()

        oCommand.CommandText = "select nvl(sum(aah04-aah05),0) as t1 from aah_file,aag_file where aah00 = aag00 and aah01 = aag01 and aah02 = " & tYear & " and aah03 = 0 AND aag08 in ('1122') and aag07 in ('2','3')"
        Dim AAS As Decimal = oCommand.ExecuteScalar()
        Ws.Cells(3, 2) = AAS

        oCommand.CommandText = "select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "sum(t" & i & ") as t" & i & ","
        Next
        oCommand.CommandText += "sum(s1) as s1 from ( select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "(case when aah03 = " & i & " then nvl(sum(aah04-aah05),0) else 0 end ) as t" & i & ","
        Next
        oCommand.CommandText += "(case when aah03 = 0 then nvl(sum(aah04-aah05),0) else 0 end ) as s1 from aah_file,aag_file where aah00 = aag00 and aah01 = aag01 and aah02 = " & tYear & " and aah03 <= " & tMonth & " AND aag08 in ('1122') and aag07 in ('2','3') group by aah03 ) "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 1 To tMonth Step 1
                    ' 設定期初
                    Dim Scount As Decimal = 0
                    Scount += oReader.Item("s1")
                    For j As Int16 = 1 To i Step 1
                        Scount += oReader.Item(j - 1)
                    Next
                    Ws.Cells(4, i + 1) = Scount
                Next
            End While
        End If
        oReader.Close()



        ' 第十四頁 新增 20180606
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(14)
        Ws.Activate()
        Ws.Name = "16.Overhead EX"
        AdjustExcelFormat10()
        oCommand.CommandText = "select aag01,aag02 from aag_file where aag01 like '5101%' and aag07 = 2 order by aag01"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("aag01")
                Ws.Cells(LineZ, 2) = oReader.Item("aag02")
                Ws.Cells(LineZ, 3) = Decimal.Round(GetLastYearSameMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 4) = Decimal.Round(GetLastMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 5) = Decimal.Round(GetThisYearSameMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 6) = GetThisYearSameMonthBudget(oReader.Item("aag01").ToString())
                Ws.Cells(LineZ, 7) = "=E" & LineZ & "-F" & LineZ
                Ws.Cells(LineZ, 8) = "=E" & LineZ & "-C" & LineZ
                Ws.Cells(LineZ, 9) = "=E" & LineZ & "-D" & LineZ
                Ws.Cells(LineZ, 10) = Decimal.Round(GetLastYearBeforeMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 11) = Decimal.Round(GetThisYearBeforeMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 12) = GetThisYearBeforeMonthBudget(oReader.Item("aag01").ToString())
                Ws.Cells(LineZ, 13) = "=K" & LineZ & "-L" & LineZ
                Ws.Cells(LineZ, 14) = "=K" & LineZ & "-J" & LineZ
                Ws.Cells(LineZ, 15) = Decimal.Round(GetLastYearNoMonth(oReader.Item("aag01").ToString()) / ExchangeRate1, 3)
                Ws.Cells(LineZ, 16) = "=Q" & LineZ & "-L" & LineZ & "+K" & LineZ
                Ws.Cells(LineZ, 17) = GetThisYearBudget(oReader.Item("aag01").ToString())
                LineZ += 1
            End While
        End If
        oReader.Close()
        Ws.Cells(LineZ, 2) = "Total Overhead"
        Ws.Cells(LineZ, 3) = "=SUM(C7:C" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 17)), Type:=xlFillDefault)
        ' 劃線
        oRng = Ws.Range("A7", Ws.Cells(LineZ, 17))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng.NumberFormat = "#,##0_ ;[Red]-#,##0 "
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "1.Sales Report 01 "
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 8.22
        oRng = Ws.Range("B2", "D2")
        oRng.EntireColumn.ColumnWidth = 24.11
        oRng.Merge()
        'oRng.NumberFormatLocal = "mmm-yy"
        oRng.HorizontalAlignment = xlCenter
        oRng = Ws.Range("E2", "E2")
        oRng.EntireColumn.ColumnWidth = 2.56
        oRng = Ws.Range("F2", "H2")
        oRng.EntireColumn.ColumnWidth = 24.11
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        oRng = Ws.Range("B1", "H1")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(1, 2) = "Sales Margin Analysis Customer"
        Ws.Cells(2, 2) = tYear & "/" & tMonth & "/01"
        Ws.Cells(2, 6) = "YTD " & tYear
        Ws.Cells(3, 2) = "Top 10 Customers"
        Ws.Cells(3, 3) = "Sales USD"
        Ws.Cells(3, 4) = "Buget USD"
        Ws.Cells(3, 6) = "Top 10 Customers"
        Ws.Cells(3, 7) = "Sales USD"
        Ws.Cells(3, 8) = "Buget USD"
        Ws.Cells(14, 2) = "Total"
        Ws.Cells(14, 6) = "Total"
        Ws.Cells(14, 3) = "=SUM(C4:C13)"
        Ws.Cells(14, 4) = "=SUM(D4:D13)"
        Ws.Cells(14, 7) = "=SUM(G4:G13)"
        Ws.Cells(14, 8) = "=SUM(H4:H13)"
        oRng = Ws.Range("B2", "D14")
        oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng = Ws.Range("F2", "H14")
        oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("B2", "B2")
        oRng.NumberFormatLocal = "mmm-yy"
        LineZ = 4
    End Sub
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 8.22
        oRng = Ws.Range("B2", "G2")
        oRng.EntireColumn.ColumnWidth = 40
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(2, 2) = "Output of " & tYear & "/" & tMonth

        oRng = Ws.Range("B3", "G3")
        oRng.EntireRow.HorizontalAlignment = xlCenter
        Ws.Cells(3, 2) = "Customer"
        Ws.Cells(3, 3) = "Part Name"
        Ws.Cells(3, 4) = "Part Description"
        Ws.Cells(3, 5) = "Actl Qty in unit"
        Ws.Cells(3, 6) = "Bgt. Revenue"
        Ws.Cells(3, 7) = "Actl Revenue"
        Ws.Cells(24, 2) = "Total"
        Ws.Cells(24, 5) = "=SUM(E4:E23)"
        Ws.Cells(24, 6) = "=SUM(F4:F23)"
        Ws.Cells(24, 7) = "=SUM(G4:G23)"
        oRng = Ws.Range("E4", "G24")
        oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        oRng = Ws.Range("B2", "G24")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        LineZ = 4
    End Sub
    Private Sub AdjustExcelFormat3()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "5.Margin Analysis Customer "
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 8.22
        oRng = Ws.Range("B2", "F2")
        oRng.EntireColumn.ColumnWidth = 24.11
        oRng.Merge()
        'oRng.NumberFormatLocal = "mmm-yy"
        oRng.HorizontalAlignment = xlCenter
        oRng = Ws.Range("G2", "G2")
        oRng.EntireColumn.ColumnWidth = 2.56
        oRng = Ws.Range("H2", "L2")
        oRng.EntireColumn.ColumnWidth = 24.11
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        oRng = Ws.Range("B1", "L1")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(1, 2) = "Sales Margin Analysis Customer"
        Ws.Cells(2, 2) = tYear & "/" & tMonth & "/01"
        Ws.Cells(2, 8) = "YTD " & tYear
        Ws.Cells(3, 2) = "Customer"
        Ws.Cells(3, 3) = "Sales USD"
        Ws.Cells(3, 4) = "Buget USD"
        Ws.Cells(3, 5) = "Actual Cost USD"
        Ws.Cells(3, 6) = "Actual Margin%"
        Ws.Cells(3, 8) = "Customer"
        Ws.Cells(3, 9) = "Sales USD"
        Ws.Cells(3, 10) = "Buget USD"
        Ws.Cells(3, 11) = "Actual Cost USD"
        Ws.Cells(3, 12) = "Actual Margin%"

        oRng = Ws.Range("B2", "F3")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng = Ws.Range("H2", "L3")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("B2", "B2")
        oRng.NumberFormatLocal = "mmm-yy"
        LineZ = 4
    End Sub
    Private Sub AdjustExcelFormat4()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 8.22
        oRng = Ws.Range("B2", "J2")
        oRng.EntireColumn.ColumnWidth = 30
        oRng.HorizontalAlignment = xlCenter
        oRng = Ws.Range("B2", "B3")
        oRng.Merge()
        Ws.Cells(2, 2) = "Customer"
        oRng = Ws.Range("C2", "C3")
        oRng.Merge()
        Ws.Cells(2, 3) = "Part Name"
        oRng = Ws.Range("D2", "D3")
        oRng.Merge()
        Ws.Cells(2, 4) = "Part Description"
        oRng = Ws.Range("E2", "E3")
        oRng.Merge()
        oRng.EntireColumn.ColumnWidth = 7.33
        Ws.Cells(2, 5) = "Unit"
        oRng = Ws.Range("F2", "J2")
        oRng.Merge()
        Ws.Cells(2, 2) = "Customer"

        oRng = Ws.Range("B3", "J3")
        oRng.EntireRow.HorizontalAlignment = xlCenter
        oRng.EntireRow.RowHeight = 30

        Ws.Cells(3, 6) = "Actl Qty in unit"
        Ws.Cells(3, 7) = "Output at STD cost"
        Ws.Cells(3, 8) = "Bgt. Revenue"
        Ws.Cells(3, 9) = "Actl Revenue"
        Ws.Cells(3, 10) = "Actl Margin%"
        
        oRng = Ws.Range("B2", "J3")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        LineZ = 4
    End Sub
    Private Sub AdjustExcelFormat5()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "10.Production output and CO"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 2.33
        Ws.Cells(1, 2) = "DAC Production Output and COGM-" & tYear & "/" & tMonth
        oRng = Ws.Range("B1", "I1")
        oRng.EntireColumn.ColumnWidth = 24.11
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        oRng = Ws.Range("F1", "F1")
        oRng.EntireColumn.ColumnWidth = 1.67

        oRng = Ws.Range("B2", "B3")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(2, 2) = "OUTPUT(Qty in pcs）"
        oRng = Ws.Range("C2", "E2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(2, 3) = tYear & "/" & tMonth & "/01"

        oRng = Ws.Range("C2", "C3")
        oRng.EntireRow.HorizontalAlignment = xlCenter

        oRng = Ws.Range("G2", "I2")
        oRng.Merge()
        Ws.Cells(2, 7) = "YTD " & tYear

        Ws.Cells(3, 3) = "Actual"
        Ws.Cells(3, 4) = "Budget"
        Ws.Cells(3, 5) = "Variance"
        Ws.Cells(3, 7) = "Actual"
        Ws.Cells(3, 8) = "Budget"
        Ws.Cells(3, 9) = "Variance"

        Ws.Cells(4, 2) = "FG (Qty in pcs)"
        Ws.Cells(4, 5) = "=C4-D4"
        Ws.Cells(4, 9) = "=G4-H4"
        Ws.Cells(5, 2) = "Semi-FG"
        Ws.Cells(5, 5) = "=C5-D5"
        Ws.Cells(5, 9) = "=G5-H5"
        Ws.Cells(6, 2) = "Total"
        Ws.Cells(6, 3) = "=C4+C5"
        Ws.Cells(6, 4) = "=D4+D5"
        Ws.Cells(6, 5) = "=C6-D6"
        Ws.Cells(6, 7) = "=G4+G5"
        Ws.Cells(6, 8) = "=H4+H5"
        Ws.Cells(6, 9) = "=G5-H5"
        Ws.Cells(7, 2) = "COGM(USDˊ000)"
        Ws.Cells(8, 2) = "Material"
        Ws.Cells(8, 5) = "=C8-D8"
        Ws.Cells(8, 9) = "=G8-H8"
        Ws.Cells(9, 2) = "Labor"
        Ws.Cells(9, 5) = "=C9-D9"
        Ws.Cells(9, 9) = "=G9-H9"
        Ws.Cells(10, 2) = "Overhead"
        Ws.Cells(10, 5) = "=C10-D10"
        Ws.Cells(10, 9) = "=G10-H10"
        Ws.Cells(11, 2) = "Others"
        Ws.Cells(11, 5) = "=C11-D11"
        Ws.Cells(11, 9) = "=G11-H11"
        Ws.Cells(12, 2) = "MFG Costs incurred"
        Ws.Cells(12, 3) = "=SUM(C8:C11)"
        Ws.Cells(12, 4) = "=SUM(D8:D10)"
        Ws.Cells(12, 5) = "=SUM(E8:E11)"
        Ws.Cells(12, 7) = "=SUM(G8:G10)"
        Ws.Cells(12, 8) = "=SUM(H8:H10)"
        Ws.Cells(12, 9) = "=SUM(I8:I11)"
        Ws.Cells(13, 2) = "Add Beginning WIP"
        Ws.Cells(13, 5) = "=C13-D13"
        Ws.Cells(13, 9) = "=G13-H13"
        Ws.Cells(14, 2) = "Deduct Ending WIP"
        Ws.Cells(14, 5) = "=C14-D14"
        Ws.Cells(14, 9) = "=G14-H14"
        Ws.Cells(15, 2) = "Total MFG Costs"
        Ws.Cells(15, 3) = "=C12+C13-C14"
        Ws.Cells(15, 4) = "=D12"
        Ws.Cells(15, 5) = "=E12+E13-E14"
        Ws.Cells(15, 7) = "=G12+G13-G14"
        Ws.Cells(15, 8) = "=H12"
        Ws.Cells(15, 9) = "=I12+I13-I14"

        oRng = Ws.Range("B2", "E15")
        oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng = Ws.Range("G2", "I15")
        oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("C2", "C2")
        oRng.NumberFormatLocal = "mmm-yy"
        LineZ = 4
    End Sub
    Private Sub AdjustExcelFormat6()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "12.Scrap report"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 18.89
        '
        oRng = Ws.Range("B1", "O1")
        oRng.EntireColumn.ColumnWidth = 15.89
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(1, 2) = tYear & " ERP Scrap Cost"

        oRng = Ws.Range("B2", "B2")
        oRng.EntireColumn.HorizontalAlignment = xlCenter
        oRng.EntireRow.HorizontalAlignment = xlCenter
        Ws.Cells(2, 2) = "Month"
        For i As Int16 = 1 To 12 Step 1
            'If i < 10 Then
            'Ws.Cells(2, 2 + i) = tYear & "/0" & i
            'Else
            Ws.Cells(2, 2 + i) = tYear & "/" & i & "/01"
            'End If
        Next
        Ws.Cells(2, 15) = "YTD " & tYear
        Ws.Cells(3, 2) = "FG Qty(Pcs)"
        Ws.Cells(3, 15) = "=SUM(C3:N3)"
        Ws.Cells(4, 2) = "value output"
        Ws.Cells(4, 15) = "=SUM(C4:N4)"
        Ws.Cells(5, 2) = "Sales（DAC）"
        Ws.Cells(5, 15) = "=SUM(C5:N5)"
        Ws.Cells(6, 2) = "Scrap qty(Pcs)"
        Ws.Cells(6, 15) = "=SUM(C6:N6)"
        Ws.Cells(7, 2) = "Scrap value"
        Ws.Cells(7, 15) = "=SUM(C7:N7)"
        Ws.Cells(8, 2) = "Scrap USD"
        Ws.Cells(8, 15) = "=SUM(C8:N8)"
        Ws.Cells(9, 2) = "Scrap（Output） %"
        Ws.Cells(9, 3) = "=IF(C4="""","""",C8/C4)"
        Ws.Cells(10, 2) = "Scrap（Sales） %"
        Ws.Cells(10, 3) = "=IF(C5="""","""",C8/C5)"
        oRng = Ws.Range("C9", "C9")
        oRng.AutoFill(Destination:=Ws.Range("C9", "O9"), Type:=xlFillDefault)
        oRng = Ws.Range("C10", "C10")
        oRng.AutoFill(Destination:=Ws.Range("C10", "O10"), Type:=xlFillDefault)
        oRng.Interior.Color = Color.FromArgb(217, 217, 217)
        oRng = Ws.Range("O3", "O9")
        oRng.Interior.Color = Color.FromArgb(217, 217, 217)
        oRng = Ws.Range("B2", "O10")
        oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng = Ws.Range("C8", "O10")
        oRng.NumberFormatLocal = "0.00%"
        'oRng = Ws.Range("G2", "I15")
        'oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        'oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        'oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        'oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        'oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("C2", "N2")
        oRng.NumberFormatLocal = "mmm-yy"
        LineZ = 3
    End Sub
    Private Sub AdjustExcelFormat7()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "15.AR Days"
        oRng = Ws.Range("A1", "O1")
        oRng.EntireColumn.ColumnWidth = 18.56
        oRng.Interior.Color = Color.LightSkyBlue
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 27.78
        For i As Int16 = 1 To 12 Step 1
            'If i < 10 Then
            'Ws.Cells(1, i + 1) = tYear & "/0" & i
            'Else
            Ws.Cells(1, i + 1) = tYear & "/" & i & "/01"
            'End If
        Next
        Ws.Cells(1, 14) = "YTD"
        Ws.Cells(1, 15) = "Target"
        Ws.Cells(2, 1) = "Revenue (after return & discount)"
        Ws.Cells(2, 14) = "=IF(SUM(B2:M2)=0,0,AVERAGE(B2:M2))"
        Ws.Cells(3, 1) = "beginning AR balance"
        Ws.Cells(4, 1) = "Ending AR balance"
        Ws.Cells(5, 1) = "Average AR"
        Ws.Cells(6, 1) = "AR turnover Rate"
        Ws.Cells(7, 1) = "AR days"
        Ws.Cells(8, 1) = "CCC days"
        Ws.Cells(3, 3) = "=B4"
        oRng = Ws.Range("C3", "C3")
        oRng.AutoFill(Destination:=Ws.Range("C3", "M3"), Type:=xlFillDefault)

        Ws.Cells(5, 2) = "=(B3+B4)/2"
        oRng = Ws.Range("B5", "B5")
        oRng.AutoFill(Destination:=Ws.Range("B5", "M5"), Type:=xlFillDefault)

        Ws.Cells(5, 14) = "=AVERAGE(B3:M3)"

        Ws.Cells(6, 2) = "=IF(B5=0,0,B2/B5)"
        oRng = Ws.Range("B6", "B6")
        oRng.AutoFill(Destination:=Ws.Range("B6", "N6"), Type:=xlFillDefault)

        Ws.Cells(7, 2) = "=IF(B6=0,0,30/B6)"
        oRng = Ws.Range("B7", "B7")
        oRng.AutoFill(Destination:=Ws.Range("B7", "N7"), Type:=xlFillDefault)

        Ws.Cells(8, 2) = "=B7+'14.Inventory Days'!B7-'13.AP Days'!B7"
        oRng = Ws.Range("B8", "B8")
        oRng.AutoFill(Destination:=Ws.Range("B8", "N8"), Type:=xlFillDefault)

        oRng = Ws.Range("O2", "O5")
        oRng.Interior.Color = Color.FromArgb(89, 89, 89)

        oRng = Ws.Range("A1", "O8")
        oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("B1", "M1")
        oRng.NumberFormatLocal = "mmm-yy"

        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat8()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "13.AP Days"
        oRng = Ws.Range("A1", "O1")
        oRng.EntireColumn.ColumnWidth = 18.56
        oRng.Interior.Color = Color.LightSkyBlue
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 27.78
        For i As Int16 = 1 To 12 Step 1
            'If i < 10 Then
            'Ws.Cells(1, i + 1) = tYear & "/0" & i
            'Else
            Ws.Cells(1, i + 1) = tYear & "/" & i & "/01"
            'End If
        Next
        Ws.Cells(1, 14) = "YTD"
        Ws.Cells(1, 15) = "Target"
        Ws.Cells(2, 1) = "COGS"
        Ws.Cells(2, 14) = "=IF(SUM(B2:M2)=0,0,AVERAGE(B2:M2))"
        Ws.Cells(3, 1) = "beginning AP balance"
        Ws.Cells(4, 1) = "Ending AP balance"
        Ws.Cells(5, 1) = "Average AP"
        Ws.Cells(6, 1) = "AP turnover Rate"
        Ws.Cells(7, 1) = "AP days"
        Ws.Cells(3, 3) = "=B4"
        oRng = Ws.Range("C3", "C3")
        oRng.AutoFill(Destination:=Ws.Range("C3", "M3"), Type:=xlFillDefault)

        Ws.Cells(5, 2) = "=(B3+B4)/2"
        oRng = Ws.Range("B5", "B5")
        oRng.AutoFill(Destination:=Ws.Range("B5", "M5"), Type:=xlFillDefault)

        Ws.Cells(5, 14) = "=AVERAGE(B3:M3)"

        Ws.Cells(6, 2) = "=IF(B5=0,0,B2/B5)"
        oRng = Ws.Range("B6", "B6")
        oRng.AutoFill(Destination:=Ws.Range("B6", "N6"), Type:=xlFillDefault)

        Ws.Cells(7, 2) = "=IF(B6=0,0,30/B6)"
        oRng = Ws.Range("B7", "B7")
        oRng.AutoFill(Destination:=Ws.Range("B7", "N7"), Type:=xlFillDefault)

        oRng = Ws.Range("O2", "O5")
        oRng.Interior.Color = Color.FromArgb(89, 89, 89)

        oRng = Ws.Range("A1", "O7")
        oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("B1", "M1")
        oRng.NumberFormatLocal = "mmm-yy"

        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat9()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "14.Inventory Days"
        oRng = Ws.Range("A1", "O1")
        oRng.EntireColumn.ColumnWidth = 18.56
        oRng.Interior.Color = Color.LightSkyBlue
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 27.78
        For i As Int16 = 1 To 12 Step 1
            'If i < 10 Then
            'Ws.Cells(1, i + 1) = tYear & "/0" & i
            'Else
            Ws.Cells(1, i + 1) = tYear & "/" & i & "/01"
            'End If
        Next
        Ws.Cells(1, 14) = "YTD"
        Ws.Cells(1, 15) = "Target"
        Ws.Cells(2, 1) = "COGS"
        Ws.Cells(2, 14) = "=IF(SUM(B2:M2)=0,0,AVERAGE(B2:M2))"
        Ws.Cells(3, 1) = "beginning Inventory balance"
        Ws.Cells(4, 1) = "Ending Inventory balance"
        Ws.Cells(5, 1) = "Average Inventory"
        Ws.Cells(6, 1) = "Inventory turnover Rate"
        Ws.Cells(7, 1) = "Inventory days"
        Ws.Cells(3, 3) = "=B4"
        oRng = Ws.Range("C3", "C3")
        oRng.AutoFill(Destination:=Ws.Range("C3", "M3"), Type:=xlFillDefault)

        Ws.Cells(5, 2) = "=(B3+B4)/2"
        oRng = Ws.Range("B5", "B5")
        oRng.AutoFill(Destination:=Ws.Range("B5", "M5"), Type:=xlFillDefault)

        Ws.Cells(5, 14) = "=AVERAGE(B3:M3)"

        Ws.Cells(6, 2) = "=IF(B5=0,0,B2/B5)"
        oRng = Ws.Range("B6", "B6")
        oRng.AutoFill(Destination:=Ws.Range("B6", "N6"), Type:=xlFillDefault)

        Ws.Cells(7, 2) = "=IF(B6=0,0,30/B6)"
        oRng = Ws.Range("B7", "B7")
        oRng.AutoFill(Destination:=Ws.Range("B7", "N7"), Type:=xlFillDefault)

        oRng = Ws.Range("O2", "O5")
        oRng.Interior.Color = Color.FromArgb(89, 89, 89)

        oRng = Ws.Range("A1", "O7")
        oRng.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("B1", "M1")
        oRng.NumberFormatLocal = "mmm-yy"

        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat10()
        Dim tDate As Date = Convert.ToDateTime(tYear & "/" & tMonth & "/01")
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 10.44
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 60
        oRng = Ws.Range("A3", "Q3")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        lMonth = tDate.AddMonths(-1).Month
        'oRng.Interior.Color = Color.FromArgb(169, 209, 141)
        Ws.Cells(3, 1) = "Overhead  By account"
        oRng = Ws.Range("A4", "B4")
        oRng.Merge()
        oRng.HorizontalAlignment = xlLeft
        Ws.Cells(4, 1) = "USD"
        oRng = Ws.Range("A5", "B5")
        oRng.Merge()
        oRng.NumberFormatLocal = "mmm-yy"
        oRng.HorizontalAlignment = xlLeft
        Ws.Cells(5, 1) = tDate
        Ws.Cells(6, 2) = "Dongguan Action Composites LTD Co."
        Dim TYM1 As String = String.Empty
        If tMonth < 10 Then
            TYM1 = tYear & "0" & tMonth
        Else
            TYM1 = tYear & tMonth
        End If
        oCommand.CommandText = "select azj041 from azj_file where azj01 = 'USD' and azj02 = '" & TYM1 & "'"
        ExchangeRate1 = oCommand.ExecuteScalar()
        oRng = Ws.Range("C4", "E5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 3) = "Actual"
        Ws.Cells(6, 3) = tDate.AddYears(-1)
        Ws.Cells(6, 4) = tDate.AddMonths(-1)
        Ws.Cells(6, 5) = tDate
        Ws.Cells(6, 6) = tDate
        oRng = Ws.Range("C6", "F6")
        oRng.NumberFormatLocal = "mmm-yy"
        oRng = Ws.Range("F4", "F5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 6) = "Budget"
        oRng = Ws.Range("G4", "I4")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 7) = "Variance" ' & Chr(10) & "Act& Bud"
        Ws.Cells(5, 7) = "Act & But"
        Ws.Cells(5, 8) = "year-on-year"
        Ws.Cells(5, 9) = "Month-on-month"
        Ws.Cells(6, 7) = "USD"
        Ws.Cells(6, 8) = "USD"
        Ws.Cells(6, 9) = "USD"
        'oRng = Ws.Range("C4", "G6")
        'oRng.Interior.Color = Color.FromArgb(255, 218, 101)
        oRng = Ws.Range("J4", "K5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 10) = "Actual"
        Ws.Cells(6, 10) = "YTD " & pYear
        Ws.Cells(6, 11) = "YTD " & tYear
        oRng = Ws.Range("L4", "L5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 12) = "Budget"
        Ws.Cells(6, 12) = "YTD " & tYear
        oRng = Ws.Range("M4", "N4")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 13) = "Variance" '& Chr(10) & "Act& Bud"
        Ws.Cells(5, 13) = "Act & But"
        Ws.Cells(5, 14) = "year-on-year"
        Ws.Cells(6, 13) = "USD"
        Ws.Cells(6, 14) = "USD"
        'oRng = Ws.Range("J4", "N6")
        'oRng.Interior.Color = Color.FromArgb(156, 195, 230)
        oRng = Ws.Range("O4", "O5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 15) = "Actual"
        Ws.Cells(6, 15) = "Y" & pYear
        oRng = Ws.Range("P4", "P5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 16) = "Rollling" & Chr(10) & "Forecast"
        Ws.Cells(6, 16) = "Y" & tYear
        oRng = Ws.Range("Q4", "Q5")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 17) = "Budget"
        Ws.Cells(6, 17) = "Y" & tYear
        'oRng = Ws.Range("O4", "Q6")
        'oRng.Interior.Color = Color.FromArgb(146, 208, 80)
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        ' 劃線
        oRng = Ws.Range("B3", "Q6")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("C6", "Q6")
        oRng.HorizontalAlignment = xlRight
        LineZ = 7
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "costing reports_" & gDataBase
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
    Private Function GetBudget1(ByVal tqa02 As String)
        oCommand2.CommandText = "select nvl(round(sum(tc_bud13 * (case when tc_bud14 = 'EUR' THEN 1.2 ELSE 1 END)),2),0) AS t1 from tc_bud_file where tc_bud01 =1 and tc_bud02 = " & tYear & " and tc_bud03 = " & tMonth & " and tc_bud05 = '" & tqa02 & "'"
        Dim Result1 As Decimal = oCommand2.ExecuteScalar()
        Return Result1
    End Function
    Private Function GetBudget2(ByVal tqa02 As String)
        oCommand2.CommandText = "select nvl(round(sum(tc_bud13 * (case when tc_bud14 = 'EUR' THEN 1.2 ELSE 1 END)),2),0) AS t1 from tc_bud_file where tc_bud01 =1 and tc_bud02 = " & tYear & " and tc_bud03 <= " & tMonth & " and tc_bud05 = '" & tqa02 & "'"
        Dim Result2 As Decimal = oCommand2.ExecuteScalar()
        Return Result2
    End Function
    Private Function GetBudget3(ByVal ogb04 As String)
        oCommand2.CommandText = "select nvl(round(sum(tc_bud13 * (case when tc_bud14 = 'EUR' THEN 1.2 ELSE 1 END)),2),0) AS t1 from tc_bud_file where tc_bud01 =1 and tc_bud02 = " & tYear & " and tc_bud03 = " & tMonth & " and tc_bud04 = '" & ogb04 & "'"
        Dim Result3 As Decimal = oCommand2.ExecuteScalar()
        Return Result3
    End Function
    Private Function GetBudget4(ByVal ogb04 As String)
        oCommand2.CommandText = "select nvl(round(sum(tc_bud13 * (case when tc_bud14 = 'EUR' THEN 1.2 ELSE 1 END)),2),0) AS t1 from tc_bud_file where tc_bud01 =1 and tc_bud02 = " & tYear & " and tc_bud03 <= " & tMonth & " and tc_bud04 = '" & ogb04 & "'"
        Dim Result3 As Decimal = oCommand2.ExecuteScalar()
        Return Result3
    End Function
    Private Function NewGetD146103(ByVal eType As Int16, ByVal sMonth As Int16)
        oCommand2.CommandText = "select nvl(sum(tlf10 * tlf12 * tlf907),0)   from tlf_file,ima_file where tlf01 = ima01 and ima06 = '103' and ima08 = 'M' and ima01 <> 'AC9999' "
        oCommand2.CommandText += "and tlf13 = 'aimt324' and tlf902 = 'D146103' and year(tlf06) = " & tYear & " and month(tlf06) "
        If eType = 1 Then
            oCommand2.CommandText += " <= "
        Else
            oCommand2.CommandText += " = "
        End If
        oCommand2.CommandText += sMonth.ToString()
        Dim SSA As Decimal = oCommand2.ExecuteScalar()
        Return SSA
    End Function
    Private Function NewGetD146103USD(ByVal eType As Int16, ByVal sMonth As Int16)
        oCommand2.CommandText = "select nvl(round(sum(tlf10 * tlf12 * tlf907 * (stb07+stb08+stb09+stb09a) / azj041),2),0)  from tlf_file,ima_file,stb_file,azj_file "
        oCommand2.CommandText += "where tlf01 = ima01 and tlf01 = stb01 and ima06 = '103' and ima08 = 'M' and ima01 <> 'AC9999' and stb02 = year(tlf06) and stb03 = month(tlf06) "
        oCommand2.CommandText += "and azj01 = 'USD' and azj02 = stb02 || (case when length(stb03) = 1 then '0' end) || stb03 and tlf13 = 'aimt324' and tlf902 = 'D146103' and year(tlf06) = " & tYear & " and month(tlf06) "
        If eType = 1 Then
            oCommand2.CommandText += " <= "
        Else
            oCommand2.CommandText += " = "
        End If
        oCommand2.CommandText += sMonth.ToString()
        Dim SSA As Decimal = oCommand2.ExecuteScalar()
        Return SSA
    End Function
    Private Function GetD146103(ByVal eType As Int16, ByVal sMonth As Int16)
        oCommand2.CommandText = "select nvl(sum(tlf10 * tlf12),0)  from tlf_file,ima_file where tlf01 = ima01 and ima06 = '103' and ima08 = 'M' and ima01 <> 'AC9999' "
        oCommand2.CommandText += "and tlf13 = 'aimt324' and tlf902 = 'D146103' and tlf907 = 1 and year(tlf06) = " & tYear & " and month(tlf06) "
        If eType = 1 Then
            oCommand2.CommandText += " <= "
        Else
            oCommand2.CommandText += " = "
        End If
        oCommand2.CommandText += sMonth.ToString()
        Dim SSA As Decimal = oCommand2.ExecuteScalar()
        Return SSA
    End Function
    Private Function GetVoucher(ByVal sMonth As Int16, ByVal eType As Int16)
        oCommand2.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file where aah02 = " & tYear & " and aah03 "
        If eType = 1 Then
            oCommand2.CommandText += " <= "
        Else
            oCommand2.CommandText += " < "
        End If
        oCommand2.CommandText += sMonth.ToString() & " and aah01 = '1405' "
        Dim AAS As Decimal = oCommand2.ExecuteScalar()
        Return AAS
    End Function
    Private Function GetTlfCost(ByVal sMonth As Int16)
        'oCommand2.CommandText = "select nvl(round(sum(tlf10 * tlf12 * (stb07+stb08+stb09+stb09a) / azj041),2),0) "
        'oCommand2.CommandText += "from tlf_file, ima_file, stb_file, azj_file where tlf01 = ima01 and tlf01 = stb01 and ima06 = '103' and ima08 = 'M' and ima01 not like 'A%' and ima01 not like 'S%' and azj01  = 'USD' and azj02 = stb02 || case when length(stb03) = 1 then '0' end || stb03 "
        'oCommand2.CommandText += "and tlf13 = 'aimt324' and tlf902 = 'D146103' and tlf907 = 1 and year(tlf06) = " & tYear & " and month(tlf06) = " & sMonth & " and year(tlf06) =stb02 and month(tlf06) = stb03"
        oCommand2.CommandText = "select nvl(round(sum(ccc63 / azj041),2),0) from ccc_file, ima_file, azj_file where ccc01 = ima01 and ima06 = '103' and ima08 = 'M' and ima01 not like 'A%' and ima01 not like 'S%' "
        oCommand2.CommandText += "and azj01  = 'USD' and azj02 = ccc02 || case when length(ccc03) = 1 then '0' end || ccc03 and ccc02 = " & tYear & " and ccc03 = " & sMonth
        Dim ASS As Decimal = oCommand2.ExecuteScalar()
        Return ASS
    End Function
    Private Function GetD146109(ByVal sMonth As Int16)
        'oCommand2.CommandText = "select nvl(sum(tlf10 * tlf12),0) as t1 from tlf_file, ima_file where tlf01 = ima01 and ima06 in ('102','103') and ima08 = 'M' and ima01 <> 'AC9999' "
        'oCommand2.CommandText += "and tlf13 = 'aimt302' and tlf902 = 'D146109' and tlf907 = 1 and year(tlf06) = " & tYear & " and month(tlf06) = " & sMonth
        oCommand2.CommandText = "select nvl(sum(inb09 * inb08_fac),0) as t1 from ina_file,inb_file, ima_file where ina01 = inb01 and inb04 = ima01 "
        oCommand2.CommandText += "and ima06 in ('102','103') and ima01 <> 'AC9999' and ina00 in ('3', '4') and inb05 = 'D146109' and ina04 not in ('D2300','D3100','D1600','D1461') and inapost = 'Y' and year(ina02) = " & tYear & " and month(ina02) = " & sMonth
        Dim ASS As Decimal = oCommand2.ExecuteScalar()
        Return ASS
    End Function
    Private Function GetD146109Cost(ByVal sMonth As Int16)
        'oCommand2.CommandText = "select nvl(round(sum(tlf10 * tlf12 * ccc23),2),0) from tlf_file, ima_file, ccc_file "
        'oCommand2.CommandText += "where tlf01 = ima01 and tlf01 = ccc01 and ima06 in ('103','102') and ima08 = 'M' and ima01 <> 'AC9999' "
        'oCommand2.CommandText += "and tlf13 = 'aimt302' and tlf902 = 'D146109' and tlf907 = 1 and year(tlf06) = " & tYear & " and month(tlf06) = " & sMonth & " and year(tlf06) =ccc02 and month(tlf06) = ccc03"
        'oCommand2.CommandText = "select nvl(round(sum(inb09 * inb08_fac * ccc23),2),0) as t1 from ina_file,inb_file, ima_file,ccc_file where ina01 = inb01 and inb04 = ima01 and inb04 = ccc01 "
        'oCommand2.CommandText += "and ima06 in ('102','103') and ima08 = 'M' and ima01 <> 'AC9999' and year(ina02) = ccc02 and month(ina02) = ccc03 "
        'oCommand2.CommandText += "and ina00 in ('3', '4') and inb05 = 'D146109' and inapost = 'Y' and year(ina02) = " & tYear & " and month(ina02) = " & sMonth
        oCommand2.CommandText = "select nvl(round(sum(inb09 * inb08_fac * (stb07 + stb08 + stb09 + stb09a)),2),0) as t1 from ina_file,inb_file, ima_file,stb_file where ina01 = inb01 and inb04 = ima01 and inb04 = stb01 "
        oCommand2.CommandText += "and ima06 in ('102','103') and ima01 <> 'AC9999' and year(ina02) = stb02 and month(ina02) = stb03 "
        oCommand2.CommandText += "and ina00 in ('3', '4') and inb05 = 'D146109' and ina04 not in ('D2300','D3100','D1600','D1461') and inapost = 'Y' and year(ina02) = " & tYear & " and month(ina02) = " & sMonth
        Dim ASS As Decimal = oCommand2.ExecuteScalar()
        Return ASS
    End Function
    Private Function GetD146109USD(ByVal sMonth As Int16)
        'oCommand2.CommandText = "select nvl(round(sum(tlf10 * tlf12 * ccc23 /azj041),2),0) from tlf_file, ima_file, ccc_file, azj_file "
        'oCommand2.CommandText += "where tlf01 = ima01 and tlf01 = ccc01 and ima06 in ('103','102') and ima08 = 'M' and ima01 <> 'AC9999' and azj01 ='USD' AND azj02 = ccc02 || case when length(ccc03) = 1 then '0' end || ccc03 "
        'oCommand2.CommandText += "and tlf13 = 'aimt302' and tlf902 = 'D146109' and tlf907 = 1 and year(tlf06) = " & tYear & " and month(tlf06) = " & sMonth & " and year(tlf06) =ccc02 and month(tlf06) = ccc03"
        'oCommand2.CommandText = "select nvl(round(sum(inb09 * inb08_fac * ccc23 /azj041),2),0) as t1 from ina_file,inb_file, ima_file,ccc_file,azj_file "
        'oCommand2.CommandText += "where ina01 = inb01 and inb04 = ima01 and inb04 = ccc01 and azj01 = 'USD' and azj02 = year(ina02) || case when length(month(ina02)) = 1 then '0' end || month(ina02) "
        'oCommand2.CommandText += "and ima06 in ('102','103') and ima08 = 'M' and ima01 <> 'AC9999' and year(ina02) = ccc02 and month(ina02) = ccc03 "
        'oCommand2.CommandText += "and ina00 in ('3', '4') and inb05 = 'D146109' and inapost = 'Y' and year(ina02) = " & tYear & " and month(ina02) = " & sMonth
        oCommand2.CommandText = "select nvl(round(sum(inb09 * inb08_fac * (stb07 + stb08 + stb09 + stb09a) /azj041),2),0) as t1 from ina_file,inb_file, ima_file,stb_file,azj_file "
        oCommand2.CommandText += "where ina01 = inb01 and inb04 = ima01 and inb04 = stb01 and azj01 = 'USD' and azj02 = year(ina02) || case when length(month(ina02)) = 1 then '0' end || month(ina02) "
        oCommand2.CommandText += "and ima06 in ('102','103') and ima08 = 'M' and ima01 <> 'AC9999' and year(ina02) = stb02 and month(ina02) = stb03 "
        oCommand2.CommandText += "and ina00 in ('3', '4') and inb05 = 'D146109' and inapost = 'Y' and year(ina02) = " & tYear & " and month(ina02) = " & sMonth
        Dim ASS As Decimal = oCommand2.ExecuteScalar()
        Return ASS
    End Function
    Private Function GetLastYearSameMonth(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 <> 'D9999' and aao03 = "
        oCommand2.CommandText += pYear & " and aao04 = " & tMonth
        Dim LYTM As Decimal = oCommand2.ExecuteScalar()
        Return LYTM
    End Function
    Private Function GetThisYearSameMonth(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 <> 'D9999' and aao03 = "
        oCommand2.CommandText += tYear & " and aao04 = " & tMonth
        Dim TYTM As Decimal = oCommand2.ExecuteScalar()
        Return TYTM
    End Function
    Private Function GetThisYearSameMonthBudget(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(tc_bud13),0) from tc_bud_file where tc_bud07 = '" & aag01 & "' and tc_bud02 = "
        oCommand2.CommandText += tYear & " and tc_bud03 = " & tMonth
        Dim TYTMB As Decimal = Decimal.Round(oCommand2.ExecuteScalar() / ExchangeRate1, 3)
        Return TYTMB
    End Function
    Private Function GetLastYearBeforeMonth(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 <> 'D9999' and aao03 = "
        oCommand2.CommandText += pYear & " and aao04 <= " & tMonth & " and aao04 > 0"
        Dim LYBM As Decimal = oCommand2.ExecuteScalar()
        Return LYBM
    End Function
    Private Function GetThisYearBeforeMonth(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 <> 'D9999' and aao03 = "
        oCommand2.CommandText += tYear & " and aao04 <= " & tMonth & " and aao04 > 0"
        Dim TYBM As Decimal = oCommand2.ExecuteScalar()
        Return TYBM
    End Function
    Private Function GetThisYearBeforeMonthBudget(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(tc_bud13),0) from tc_bud_file where tc_bud07 = '" & aag01 & "' and tc_bud02 = "
        oCommand2.CommandText += tYear & " and tc_bud03 <= " & tMonth
        Dim TYBMB As Decimal = Decimal.Round(oCommand2.ExecuteScalar() / ExchangeRate1, 3)
        Return TYBMB
    End Function
    Private Function GetLastYearNoMonth(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 <> 'D9999' and aao03 = "
        oCommand2.CommandText += pYear.ToString() & " and aao04 > 0"
        Dim TYNM As Decimal = oCommand2.ExecuteScalar()
        Return TYNM
    End Function
    Private Function GetThisYearBudget(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(tc_bud13),0) from tc_bud_file where tc_bud07 = '" & aag01 & "' and tc_bud02 = "
        oCommand2.CommandText += tYear.ToString()
        Dim TYTMB As Decimal = Decimal.Round(oCommand2.ExecuteScalar() / ExchangeRate1, 3)
        Return TYTMB
    End Function
    Private Function GetLastMonth(ByVal aag01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 <> 'D9999' and aao03 = "
        oCommand2.CommandText += lYear & " and aao04 = " & lMonth
        Dim TYTM As Decimal = oCommand2.ExecuteScalar()
        Return TYTM
    End Function
End Class