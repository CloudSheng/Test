Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlConditionValueTypes
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form138
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
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim pYear As Int16 = 0
    Dim LineZ As Integer = 0
    Dim DNP As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
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
        tYear = Me.TextBox2.Text
        tMonth = Me.TextBox3.Text
        pYear = tYear - 1
        DNP = Me.TextBox1.Text
        
        'ExportToExcel()
        'SaveExcel()
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Form138_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        TextBox2.Text = Now.Year
        TextBox3.Text = Now.Month
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "产品分析明细表"
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
        xWorkBook.Sheets.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Name = "按年汇总非样品"
        Ws.Activate()
        AdjustExcelFormat()
        oCommand.CommandText = "select oga08,ogb04,ima02,ima021,ima06,ima08,ogb05,'',sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,"
        oCommand.CommandText += "sum(t5) as t5,sum(t6) as t6 from ( "
        oCommand.CommandText += "select (case when oga08 = '2' then '外銷' else '內銷' end) as oga08,ogb04, ima02,ima021,ima06,ima08,ogb05,'',"
        oCommand.CommandText += "sum(ogb12) as t1,round(sum(oga24 * ogb14),2) as t2,"
        oCommand.CommandText += "round(sum(ogb12 * ogb05_fac * ccc23),2) as t3,0 as t4,0 as t5,0 as t6 from oga_file left join ogb_file on oga01 = ogb01 "
        oCommand.CommandText += "left join ima_file on ogb04 = ima01 left join ccc_file on ogb04 = ccc01 and year(oga02) = ccc02 and month(oga02) = ccc03 "
        oCommand.CommandText += "left join oea_file on ogb31 = oea01 where year(oga02) =" & pYear & " and month(oga02) <= " & tMonth & " and ogapost = 'Y' and (ta_oea01 <> 'Y'  or ta_oea01 is null) and ogb09 not in (select jce02 from jce_file) "
        If Not String.IsNullOrEmpty(DNP) Then
            oCommand.CommandText += " AND ogb04 like '" & DNP & "%' "
        End If
        oCommand.CommandText += "group by oga08,ogb04, ima02,ima021,ima06,ima08,ogb05 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select (case when oha08 = '2' then '外銷' else '內銷' end) as oha08,ohb04, ima02,ima021,ima06,ima08,ohb05,'',"
        oCommand.CommandText += "sum(ohb12 * ohb05_fac * -1) as t1,round(sum(oha24 * ohb14 * -1),2) as t2,"
        oCommand.CommandText += "round(sum(ohb12 * ccc23 * -1),2) as t3,0 as t4,0 as t5,0 as t6 from oha_file left join ohb_file on oha01 = ohb01 "
        oCommand.CommandText += "left join ima_file on ohb04 = ima01 left join ccc_file on ohb04 = ccc01 and year(oha02) = ccc02 and month(oha02) = ccc03 "
        oCommand.CommandText += "left join oea_file on ohb31 = oea01 where year(oha02) =" & pYear & " and month(oha02) <= " & tMonth & " and ohapost = 'Y' and (ta_oea01 <> 'Y'  or ta_oea01 is null) and ohb09 not in (select jce02 from jce_file) "
        If Not String.IsNullOrEmpty(DNP) Then
            oCommand.CommandText += " AND ohb04 like '" & DNP & "%' "
        End If
        oCommand.CommandText += "group by oha08,ohb04, ima02,ima021,ima06,ima08,ohb05 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select (case when oga08 = '2' then '外銷' else '內銷' end) as oga08,ogb04, ima02,ima021,ima06,ima08,ogb05,'',0,0,0,"
        oCommand.CommandText += "sum(ogb12) as t4, round(sum(oga24 * ogb14),2) as t5,"
        oCommand.CommandText += "round(sum(ogb12 * ogb05_fac * ccc23),2) as t6 from oga_file left join ogb_file on oga01 = ogb01 left join ima_file on ogb04 = ima01 left join ccc_file on ogb04 = ccc01 and year(oga02) = ccc02 and month(oga02) = ccc03 "
        oCommand.CommandText += "left join oea_file on ogb31 = oea01 where year(oga02) =" & tYear & " and month(oga02) <= " & tMonth & " and ogapost = 'Y' and (ta_oea01 <> 'Y'  or ta_oea01 is null) and ogb09 not in (select jce02 from jce_file) "
        If Not String.IsNullOrEmpty(DNP) Then
            oCommand.CommandText += " AND ogb04 like '" & DNP & "%' "
        End If
        oCommand.CommandText += "group by oga08,ogb04, ima02,ima021,ima06,ima08,ogb05 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select (case when oha08 = '2' then '外銷' else '內銷' end) as oha08,ohb04, ima02,ima021,ima06,ima08,ohb05,'',0,0,0,"
        oCommand.CommandText += "sum(ohb12 * ohb05_fac * -1) as t4,round(sum(oha24 * ohb14 * -1),2) as t5,"
        oCommand.CommandText += "round(sum(ohb12 * ccc23 * -1),2) as t6 from oha_file left join ohb_file on oha01 = ohb01 "
        oCommand.CommandText += "left join ima_file on ohb04 = ima01 left join ccc_file on ohb04 = ccc01 and year(oha02) = ccc02 and month(oha02) = ccc03 "
        oCommand.CommandText += "left join oea_file on ohb31 = oea01 where year(oha02) =" & tYear & " and month(oha02) <= " & tMonth & " and ohapost = 'Y' and (ta_oea01 <> 'Y'  or ta_oea01 is null) and ohb09 not in (select jce02 from jce_file) "
        If Not String.IsNullOrEmpty(DNP) Then
            oCommand.CommandText += " AND ohb04 like '" & DNP & "%' "
        End If
        oCommand.CommandText += "group by oha08,ohb04, ima02,ima021,ima06,ima08,ohb05 ) group by oga08,ogb04,ima02,ima021,ima06,ima08,ogb05"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("oga08")
                Ws.Cells(LineZ, 3) = oReader.Item("ogb04")
                Ws.Cells(LineZ, 4) = oReader.Item("ima02")
                Ws.Cells(LineZ, 5) = oReader.Item("ima021")
                Ws.Cells(LineZ, 6) = oReader.Item("ima06")
                Ws.Cells(LineZ, 7) = oReader.Item("ima08")
                Ws.Cells(LineZ, 8) = oReader.Item("ogb05")
                Ws.Cells(LineZ, 10) = oReader.Item("t1")
                Ws.Cells(LineZ, 11) = "=IFERROR(L" & LineZ & "/J" & LineZ & ",0)"
                Ws.Cells(LineZ, 12) = oReader.Item("t2")
                Ws.Cells(LineZ, 13) = "=IFERROR(N" & LineZ & "/J" & LineZ & ",0)"
                Ws.Cells(LineZ, 14) = oReader.Item("t3")
                Ws.Cells(LineZ, 15) = "=L" & LineZ & "-N" & LineZ
                Ws.Cells(LineZ, 16) = "=IFERROR(O" & LineZ & "/L" & LineZ & ",0)"
                Ws.Cells(LineZ, 17) = oReader.Item("t4")
                Ws.Cells(LineZ, 18) = "=IFERROR(S" & LineZ & "/Q" & LineZ & ",0)"
                Ws.Cells(LineZ, 19) = oReader.Item("t5")
                Ws.Cells(LineZ, 20) = "=IFERROR(U" & LineZ & "/Q" & LineZ & ",0)"
                Ws.Cells(LineZ, 21) = oReader.Item("t6")
                Ws.Cells(LineZ, 22) = "=S" & LineZ & "-U" & LineZ
                Ws.Cells(LineZ, 23) = "=IFERROR(V" & LineZ & "/S" & LineZ & ",0)"
                Ws.Cells(LineZ, 24) = "=Q" & LineZ & "-J" & LineZ
                Ws.Cells(LineZ, 25) = "=S" & LineZ & "-L" & LineZ
                Ws.Cells(LineZ, 26) = "=U" & LineZ & "-N" & LineZ
                Ws.Cells(LineZ, 27) = "=R" & LineZ & "-K" & LineZ
                Ws.Cells(LineZ, 28) = "=T" & LineZ & "-M" & LineZ
                Ws.Cells(LineZ, 29) = "=IFERROR(AA" & LineZ & "/R" & LineZ & ",0)"
                Ws.Cells(LineZ, 30) = "=IFERROR(AB" & LineZ & "/T" & LineZ & ",0)"
                Ws.Cells(LineZ, 31) = "=W" & LineZ & "-P" & LineZ
                LineZ += 1
            End While
            ' 加總
            Ws.Cells(LineZ, 2) = "Total:"
            Ws.Cells(LineZ, 10) = "=SUM(J6:J" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 12) = "=SUM(L6:L" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 14) = "=SUM(N6:N" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 15) = "=SUM(O6:O" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 16) = "=IFERROR(O" & LineZ & "/L" & LineZ & ",0)"
            Ws.Cells(LineZ, 17) = "=SUM(Q6:Q" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 19) = "=SUM(S6:S" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 21) = "=SUM(U6:U" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 22) = "=SUM(V6:V" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 23) = "=IFERROR(V" & LineZ & "/S" & LineZ & ",0)"
            Ws.Cells(LineZ, 24) = "=SUM(X6:X" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 25) = "=SUM(Y6:Y" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 26) = "=SUM(Z6:Z" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 27) = "=R" & LineZ & "-K" & LineZ
            Ws.Cells(LineZ, 28) = "=T" & LineZ & "-M" & LineZ
            Ws.Cells(LineZ, 31) = "=W" & LineZ & "-P" & LineZ

            ' 劃線
            oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 31))
            oRng.Borders(xlEdgeLeft).LineStyle = xlNone
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlDouble
            oRng.Borders(xlEdgeRight).LineStyle = xlNone
            oRng.Borders(xlInsideHorizontal).LineStyle = xlNone
            oRng.Borders(xlInsideVertical).LineStyle = xlNone
        End If
        oReader.Close()
        oRng = Ws.Range("C1", "AE1")
        oRng.EntireColumn.AutoFit()

        ' 第二頁
        Ws = xWorkBook.Sheets(2)
        Ws.Name = "按月汇总非样品"
        Ws.Activate()
        AdjustExcelFormat1()
        oCommand.CommandText = "select oga08,oga02,ogb04,ima02,ima021,ima06,ima08,ogb05,'',sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,"
        oCommand.CommandText += "sum(t5) as t5,sum(t6) as t6 from ( "
        oCommand.CommandText += "select (case when oga08 = '2' then '外銷' else '內銷' end) as oga08,month(oga02) as oga02,ogb04, ima02,ima021,ima06,ima08,ogb05,'',"
        oCommand.CommandText += "sum(ogb12) as t1,round(sum(oga24 * ogb14),2) as t2,"
        oCommand.CommandText += "round(sum(ogb12 * ogb05_fac * ccc23),2) as t3,0 as t4,0 as t5,0 as t6 from oga_file left join ogb_file on oga01 = ogb01 "
        oCommand.CommandText += "left join ima_file on ogb04 = ima01 left join ccc_file on ogb04 = ccc01 and year(oga02) = ccc02 and month(oga02) = ccc03 "
        oCommand.CommandText += "left join oea_file on ogb31 = oea01 where year(oga02) =" & pYear & " and month(oga02) <= " & tMonth & " and ogapost = 'Y' and (ta_oea01 <> 'Y'  or ta_oea01 is null) and ogb09 not in (select jce02 from jce_file) "
        If Not String.IsNullOrEmpty(DNP) Then
            oCommand.CommandText += " AND ogb04 like '" & DNP & "%' "
        End If
        oCommand.CommandText += "group by oga08,oga02,ogb04, ima02,ima021,ima06,ima08,ogb05 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select (case when oha08 = '2' then '外銷' else '內銷' end) as oha08,month(oha02),ohb04, ima02,ima021,ima06,ima08,ohb05,'',"
        oCommand.CommandText += "sum(ohb12 * -1) as t1,round(sum(oha24 * ohb14 * -1),2) as t2,"
        oCommand.CommandText += "round(sum(ohb12 * ohb05_fac * ccc23 * -1),2) as t3,0 as t4,0 as t5,0 as t6 from oha_file left join ohb_file on oha01 = ohb01 "
        oCommand.CommandText += "left join ima_file on ohb04 = ima01 left join ccc_file on ohb04 = ccc01 and year(oha02) = ccc02 and month(oha02) = ccc03 "
        oCommand.CommandText += "left join oea_file on ohb31 = oea01 where year(oha02) =" & pYear & " and month(oha02) <= " & tMonth & " and ohapost = 'Y' and (ta_oea01 <> 'Y'  or ta_oea01 is null) and ohb09 not in (select jce02 from jce_file) "
        If Not String.IsNullOrEmpty(DNP) Then
            oCommand.CommandText += " AND ohb04 like '" & DNP & "%' "
        End If
        oCommand.CommandText += "group by oha08,oha02,ohb04, ima02,ima021,ima06,ima08,ohb05 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select (case when oga08 = '2' then '外銷' else '內銷' end) as oga08,month(oga02),ogb04, ima02,ima021,ima06,ima08,ogb05,'',0,0,0,"
        oCommand.CommandText += "sum(ogb12) as t4, round(sum(oga24 * ogb14),2) as t5,"
        oCommand.CommandText += "round(sum(ogb12 * ogb05_fac * ccc23),2) as t6 from oga_file left join ogb_file on oga01 = ogb01 left join ima_file on ogb04 = ima01 left join ccc_file on ogb04 = ccc01 and year(oga02) = ccc02 and month(oga02) = ccc03 "
        oCommand.CommandText += "left join oea_file on ogb31 = oea01 where year(oga02) =" & tYear & " and month(oga02) <= " & tMonth & " and ogapost = 'Y' and (ta_oea01 <> 'Y'  or ta_oea01 is null) and ogb09 not in (select jce02 from jce_file) "
        If Not String.IsNullOrEmpty(DNP) Then
            oCommand.CommandText += " AND ogb04 like '" & DNP & "%' "
        End If
        oCommand.CommandText += "group by oga08,oga02,ogb04, ima02,ima021,ima06,ima08,ogb05 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select (case when oha08 = '2' then '外銷' else '內銷' end) as oha08,month(oha02),ohb04, ima02,ima021,ima06,ima08,ohb05,'',0,0,0,"
        oCommand.CommandText += "sum(ohb12 * -1) as t4,round(sum(oha24 * ohb14 * -1),2) as t5,"
        oCommand.CommandText += "round(sum(ohb12 * ohb05_fac * ccc23 * -1),2) as t6 from oha_file left join ohb_file on oha01 = ohb01 "
        oCommand.CommandText += "left join ima_file on ohb04 = ima01 left join ccc_file on ohb04 = ccc01 and year(oha02) = ccc02 and month(oha02) = ccc03 "
        oCommand.CommandText += "left join oea_file on ohb31 = oea01 where year(oha02) =" & tYear & " and month(oha02) <= " & tMonth & " and ohapost = 'Y' and (ta_oea01 <> 'Y'  or ta_oea01 is null) and ohb09 not in (select jce02 from jce_file) "
        If Not String.IsNullOrEmpty(DNP) Then
            oCommand.CommandText += " AND ohb04 like '" & DNP & "%' "
        End If
        oCommand.CommandText += "group by oha08,oha02,ohb04, ima02,ima021,ima06,ima08,ohb05 ) group by oga08,oga02,ogb04,ima02,ima021,ima06,ima08,ogb05 order by ogb04,oga02"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("oga08")
                Ws.Cells(LineZ, 3) = oReader.Item("oga02")
                Ws.Cells(LineZ, 4) = oReader.Item("ogb04")
                Ws.Cells(LineZ, 5) = oReader.Item("ima02")
                Ws.Cells(LineZ, 6) = oReader.Item("ima021")
                Ws.Cells(LineZ, 7) = oReader.Item("ima06")
                Ws.Cells(LineZ, 8) = oReader.Item("ima08")
                Ws.Cells(LineZ, 9) = oReader.Item("ogb05")
                Ws.Cells(LineZ, 11) = oReader.Item("t1")
                Ws.Cells(LineZ, 12) = "=IFERROR(M" & LineZ & "/K" & LineZ & ",0)"
                Ws.Cells(LineZ, 13) = oReader.Item("t2")
                Ws.Cells(LineZ, 14) = "=IFERROR(O" & LineZ & "/K" & LineZ & ",0)"
                Ws.Cells(LineZ, 15) = oReader.Item("t3")
                Ws.Cells(LineZ, 16) = "=M" & LineZ & "-O" & LineZ
                Ws.Cells(LineZ, 17) = "=IFERROR(P" & LineZ & "/M" & LineZ & ",0)"
                Ws.Cells(LineZ, 18) = oReader.Item("t4")
                Ws.Cells(LineZ, 19) = "=IFERROR(T" & LineZ & "/R" & LineZ & ",0)"
                Ws.Cells(LineZ, 20) = oReader.Item("t5")
                Ws.Cells(LineZ, 21) = "=IFERROR(V" & LineZ & "/R" & LineZ & ",0)"
                Ws.Cells(LineZ, 22) = oReader.Item("t6")
                Ws.Cells(LineZ, 23) = "=T" & LineZ & "-V" & LineZ
                Ws.Cells(LineZ, 24) = "=IFERROR(W" & LineZ & "/T" & LineZ & ",0)"
                Ws.Cells(LineZ, 25) = "=R" & LineZ & "-K" & LineZ
                Ws.Cells(LineZ, 26) = "=T" & LineZ & "-M" & LineZ
                Ws.Cells(LineZ, 27) = "=V" & LineZ & "-O" & LineZ
                Ws.Cells(LineZ, 28) = "=S" & LineZ & "-L" & LineZ
                Ws.Cells(LineZ, 29) = "=U" & LineZ & "-N" & LineZ
                Ws.Cells(LineZ, 30) = "=IFERROR(AB" & LineZ & "/S" & LineZ & ",0)"
                Ws.Cells(LineZ, 31) = "=IFERROR(AC" & LineZ & "/U" & LineZ & ",0)"
                Ws.Cells(LineZ, 32) = "=X" & LineZ & "-Q" & LineZ
                LineZ += 1
            End While
            ' 加總
            Ws.Cells(LineZ, 2) = "Total:"
            Ws.Cells(LineZ, 11) = "=SUM(K6:K" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 13) = "=SUM(M6:M" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 15) = "=SUM(O6:O" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 16) = "=SUM(P6:P" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 17) = "=IFERROR(P" & LineZ & "/M" & LineZ & ",0)"
            Ws.Cells(LineZ, 18) = "=SUM(R6:R" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 20) = "=SUM(T6:T" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 22) = "=SUM(V6:V" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 23) = "=SUM(W6:W" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 24) = "=IFERROR(W" & LineZ & "/T" & LineZ & ",0)"
            Ws.Cells(LineZ, 25) = "=SUM(Y6:Y" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 26) = "=SUM(Z6:Z" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 27) = "=SUM(AA6:AA" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 28) = "=S" & LineZ & "-L" & LineZ
            Ws.Cells(LineZ, 29) = "=U" & LineZ & "-N" & LineZ
            Ws.Cells(LineZ, 32) = "=X" & LineZ & "-Q" & LineZ

            ' 劃線
            oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 32))
            oRng.Borders(xlEdgeLeft).LineStyle = xlNone
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlDouble
            oRng.Borders(xlEdgeRight).LineStyle = xlNone
            oRng.Borders(xlInsideHorizontal).LineStyle = xlNone
            oRng.Borders(xlInsideVertical).LineStyle = xlNone
        End If
        oReader.Close()
        oRng = Ws.Range("C1", "AF1")
        oRng.EntireColumn.AutoFit()
        '第三頁
        Ws = xWorkBook.Sheets(3)
        Ws.Name = "按年汇总样品"
        Ws.Activate()
        AdjustExcelFormat()
        oCommand.CommandText = "select oga08,ogb04,ima02,ima021,ima06,ima08,ogb05,'',sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,"
        oCommand.CommandText += "sum(t5) as t5,sum(t6) as t6 from ( "
        oCommand.CommandText += "select (case when oga08 = '2' then '外銷' else '內銷' end) as oga08,ogb04, ima02,ima021,ima06,ima08,ogb05,'',"
        oCommand.CommandText += "sum(ogb12) as t1,round(sum(oga24 * ogb14),2) as t2,"
        oCommand.CommandText += "round(sum(ogb12 * ogb05_fac * ccc23),2) as t3,0 as t4,0 as t5,0 as t6 from oga_file left join ogb_file on oga01 = ogb01 "
        oCommand.CommandText += "left join ima_file on ogb04 = ima01 left join ccc_file on ogb04 = ccc01 and year(oga02) = ccc02 and month(oga02) = ccc03 "
        oCommand.CommandText += "left join oea_file on ogb31 = oea01 where year(oga02) =" & pYear & " and month(oga02) <= " & tMonth & " and ogapost = 'Y' and ta_oea01 = 'Y' and ogb09 not in (select jce02 from jce_file) "
        If Not String.IsNullOrEmpty(DNP) Then
            oCommand.CommandText += " AND ogb04 like '" & DNP & "%' "
        End If
        oCommand.CommandText += "group by oga08,ogb04, ima02,ima021,ima06,ima08,ogb05 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select (case when oha08 = '2' then '外銷' else '內銷' end) as oha08,ohb04, ima02,ima021,ima06,ima08,ohb05,'',"
        oCommand.CommandText += "sum(ohb12 * -1) as t1,round(sum(oha24 * ohb14 * -1),2) as t2,"
        oCommand.CommandText += "round(sum(ohb12 * ohb05_fac * ccc23 * -1),2) as t3,0 as t4,0 as t5,0 as t6 from oha_file left join ohb_file on oha01 = ohb01 "
        oCommand.CommandText += "left join ima_file on ohb04 = ima01 left join ccc_file on ohb04 = ccc01 and year(oha02) = ccc02 and month(oha02) = ccc03 "
        oCommand.CommandText += "left join oea_file on ohb31 = oea01 where year(oha02) =" & pYear & " and month(oha02) <= " & tMonth & " and ohapost = 'Y' and ta_oea01 = 'Y' and ohb09 not in (select jce02 from jce_file) "
        If Not String.IsNullOrEmpty(DNP) Then
            oCommand.CommandText += " AND ohb04 like '" & DNP & "%' "
        End If
        oCommand.CommandText += "group by oha08,ohb04, ima02,ima021,ima06,ima08,ohb05 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select (case when oga08 = '2' then '外銷' else '內銷' end) as oga08,ogb04, ima02,ima021,ima06,ima08,ogb05,'',0,0,0,"
        oCommand.CommandText += "sum(ogb12) as t4, round(sum(oga24 * ogb14),2) as t5,"
        oCommand.CommandText += "round(sum(ogb12 * ogb05_fac * ccc23),2) as t6 from oga_file left join ogb_file on oga01 = ogb01 left join ima_file on ogb04 = ima01 left join ccc_file on ogb04 = ccc01 and year(oga02) = ccc02 and month(oga02) = ccc03 "
        oCommand.CommandText += "left join oea_file on ogb31 = oea01 where year(oga02) =" & tYear & " and month(oga02) <= " & tMonth & " and ogapost = 'Y' and ta_oea01 = 'Y' and ogb09 not in (select jce02 from jce_file) "
        If Not String.IsNullOrEmpty(DNP) Then
            oCommand.CommandText += " AND ogb04 like '" & DNP & "%' "
        End If
        oCommand.CommandText += "group by oga08,ogb04, ima02,ima021,ima06,ima08,ogb05 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select (case when oha08 = '2' then '外銷' else '內銷' end) as oha08,ohb04, ima02,ima021,ima06,ima08,ohb05,'',0,0,0,"
        oCommand.CommandText += "sum(ohb12 * -1) as t4,round(sum(oha24 * ohb14 * -1),2) as t5,"
        oCommand.CommandText += "round(sum(ohb12 * ohb05_fac * ccc23 * -1),2) as t6 from oha_file left join ohb_file on oha01 = ohb01 "
        oCommand.CommandText += "left join ima_file on ohb04 = ima01 left join ccc_file on ohb04 = ccc01 and year(oha02) = ccc02 and month(oha02) = ccc03 "
        oCommand.CommandText += "left join oea_file on ohb31 = oea01 where year(oha02) =" & tYear & " and month(oha02) <= " & tMonth & " and ohapost = 'Y' and ta_oea01 = 'Y' and ohb09 not in (select jce02 from jce_file) "
        If Not String.IsNullOrEmpty(DNP) Then
            oCommand.CommandText += " AND ohb04 like '" & DNP & "%' "
        End If
        oCommand.CommandText += "group by oha08,ohb04, ima02,ima021,ima06,ima08,ohb05 ) group by oga08,ogb04,ima02,ima021,ima06,ima08,ogb05"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("oga08")
                Ws.Cells(LineZ, 3) = oReader.Item("ogb04")
                Ws.Cells(LineZ, 4) = oReader.Item("ima02")
                Ws.Cells(LineZ, 5) = oReader.Item("ima021")
                Ws.Cells(LineZ, 6) = oReader.Item("ima06")
                Ws.Cells(LineZ, 7) = oReader.Item("ima08")
                Ws.Cells(LineZ, 8) = oReader.Item("ogb05")
                Ws.Cells(LineZ, 10) = oReader.Item("t1")
                Ws.Cells(LineZ, 11) = "=IFERROR(L" & LineZ & "/J" & LineZ & ",0)"
                Ws.Cells(LineZ, 12) = oReader.Item("t2")
                Ws.Cells(LineZ, 13) = "=IFERROR(N" & LineZ & "/J" & LineZ & ",0)"
                Ws.Cells(LineZ, 14) = oReader.Item("t3")
                Ws.Cells(LineZ, 15) = "=L" & LineZ & "-N" & LineZ
                Ws.Cells(LineZ, 16) = "=IFERROR(O" & LineZ & "/L" & LineZ & ",0)"
                Ws.Cells(LineZ, 17) = oReader.Item("t4")
                Ws.Cells(LineZ, 18) = "=IFERROR(S" & LineZ & "/Q" & LineZ & ",0)"
                Ws.Cells(LineZ, 19) = oReader.Item("t5")
                Ws.Cells(LineZ, 20) = "=IFERROR(U" & LineZ & "/Q" & LineZ & ",0)"
                Ws.Cells(LineZ, 21) = oReader.Item("t6")
                Ws.Cells(LineZ, 22) = "=S" & LineZ & "-U" & LineZ
                Ws.Cells(LineZ, 23) = "=IFERROR(V" & LineZ & "/S" & LineZ & ",0)"
                Ws.Cells(LineZ, 24) = "=Q" & LineZ & "-J" & LineZ
                Ws.Cells(LineZ, 25) = "=S" & LineZ & "-L" & LineZ
                Ws.Cells(LineZ, 26) = "=U" & LineZ & "-N" & LineZ
                Ws.Cells(LineZ, 27) = "=R" & LineZ & "-K" & LineZ
                Ws.Cells(LineZ, 28) = "=T" & LineZ & "-M" & LineZ
                Ws.Cells(LineZ, 29) = "=IFERROR(AA" & LineZ & "/R" & LineZ & ",0)"
                Ws.Cells(LineZ, 30) = "=IFERROR(AB" & LineZ & "/T" & LineZ & ",0)"
                Ws.Cells(LineZ, 31) = "=W" & LineZ & "-P" & LineZ
                LineZ += 1
            End While
            ' 加總
            Ws.Cells(LineZ, 2) = "Total:"
            Ws.Cells(LineZ, 10) = "=SUM(J6:J" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 12) = "=SUM(L6:L" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 14) = "=SUM(N6:N" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 15) = "=SUM(O6:O" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 16) = "=IFERROR(O" & LineZ & "/L" & LineZ & ",0)"
            Ws.Cells(LineZ, 17) = "=SUM(Q6:Q" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 19) = "=SUM(S6:S" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 21) = "=SUM(U6:U" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 22) = "=SUM(V6:V" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 23) = "=IFERROR(V" & LineZ & "/S" & LineZ & ",0)"
            Ws.Cells(LineZ, 24) = "=SUM(X6:X" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 25) = "=SUM(Y6:Y" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 26) = "=SUM(Z6:Z" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 27) = "=R" & LineZ & "-K" & LineZ
            Ws.Cells(LineZ, 28) = "=T" & LineZ & "-M" & LineZ
            Ws.Cells(LineZ, 31) = "=W" & LineZ & "-P" & LineZ

            ' 劃線
            oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 31))
            oRng.Borders(xlEdgeLeft).LineStyle = xlNone
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlDouble
            oRng.Borders(xlEdgeRight).LineStyle = xlNone
            oRng.Borders(xlInsideHorizontal).LineStyle = xlNone
            oRng.Borders(xlInsideVertical).LineStyle = xlNone
        End If
        oReader.Close()
        oRng = Ws.Range("C1", "AE1")
        oRng.EntireColumn.AutoFit()
        ' 第四頁
        Ws = xWorkBook.Sheets(4)
        Ws.Name = "按月汇总样品"
        Ws.Activate()
        AdjustExcelFormat1()
        oCommand.CommandText = "select oga08,oga02,ogb04,ima02,ima021,ima06,ima08,ogb05,'',sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,"
        oCommand.CommandText += "sum(t5) as t5,sum(t6) as t6 from ( "
        oCommand.CommandText += "select (case when oga08 = '2' then '外銷' else '內銷' end) as oga08,month(oga02) as oga02,ogb04, ima02,ima021,ima06,ima08,ogb05,'',"
        oCommand.CommandText += "sum(ogb12) as t1,round(sum(oga24 * ogb14),2) as t2,"
        oCommand.CommandText += "round(sum(ogb12 * ogb05_fac * ccc23),2) as t3,0 as t4,0 as t5,0 as t6 from oga_file left join ogb_file on oga01 = ogb01 "
        oCommand.CommandText += "left join ima_file on ogb04 = ima01 left join ccc_file on ogb04 = ccc01 and year(oga02) = ccc02 and month(oga02) = ccc03 "
        oCommand.CommandText += "left join oea_file on ogb31 = oea01 where year(oga02) =" & pYear & " and month(oga02) <= " & tMonth & " and ogapost = 'Y' and ta_oea01 = 'Y' and ogb09 not in (select jce02 from jce_file) "
        If Not String.IsNullOrEmpty(DNP) Then
            oCommand.CommandText += " AND ogb04 like '" & DNP & "%' "
        End If
        oCommand.CommandText += "group by oga08,oga02,ogb04, ima02,ima021,ima06,ima08,ogb05 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select (case when oha08 = '2' then '外銷' else '內銷' end) as oha08,month(oha02),ohb04, ima02,ima021,ima06,ima08,ohb05,'',"
        oCommand.CommandText += "sum(ohb12 * ohb05_fac * -1) as t1,round(sum(oha24 * ohb14 * -1),2) as t2,"
        oCommand.CommandText += "round(sum(ohb12 * ccc23 * -1),2) as t3,0 as t4,0 as t5,0 as t6 from oha_file left join ohb_file on oha01 = ohb01 "
        oCommand.CommandText += "left join ima_file on ohb04 = ima01 left join ccc_file on ohb04 = ccc01 and year(oha02) = ccc02 and month(oha02) = ccc03 "
        oCommand.CommandText += "left join oea_file on ohb31 = oea01 where year(oha02) =" & pYear & " and month(oha02) <= " & tMonth & " and ohapost = 'Y' and ta_oea01 = 'Y' and ohb09 not in (select jce02 from jce_file) "
        If Not String.IsNullOrEmpty(DNP) Then
            oCommand.CommandText += " AND ohb04 like '" & DNP & "%' "
        End If
        oCommand.CommandText += "group by oha08,oha02,ohb04, ima02,ima021,ima06,ima08,ohb05 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select (case when oga08 = '2' then '外銷' else '內銷' end) as oga08,month(oga02),ogb04, ima02,ima021,ima06,ima08,ogb05,'',0,0,0,"
        oCommand.CommandText += "sum(ogb12) as t4, round(sum(oga24 * ogb14),2) as t5,"
        oCommand.CommandText += "round(sum(ogb12 * ogb05_fac * ccc23),2) as t6 from oga_file left join ogb_file on oga01 = ogb01 left join ima_file on ogb04 = ima01 left join ccc_file on ogb04 = ccc01 and year(oga02) = ccc02 and month(oga02) = ccc03 "
        oCommand.CommandText += "left join oea_file on ogb31 = oea01 where year(oga02) =" & tYear & " and month(oga02) <= " & tMonth & " and ogapost = 'Y' and ta_oea01 = 'Y' and ogb09 not in (select jce02 from jce_file) "
        If Not String.IsNullOrEmpty(DNP) Then
            oCommand.CommandText += " AND ogb04 like '" & DNP & "%' "
        End If
        oCommand.CommandText += "group by oga08,oga02,ogb04, ima02,ima021,ima06,ima08,ogb05 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select (case when oha08 = '2' then '外銷' else '內銷' end) as oha08,month(oha02),ohb04, ima02,ima021,ima06,ima08,ohb05,'',0,0,0,"
        oCommand.CommandText += "sum(ohb12 * -1) as t4,round(sum(oha24 * ohb14 * -1),2) as t5,"
        oCommand.CommandText += "round(sum(ohb12 * ohb05_fac * ccc23 * -1),2) as t6 from oha_file left join ohb_file on oha01 = ohb01 "
        oCommand.CommandText += "left join ima_file on ohb04 = ima01 left join ccc_file on ohb04 = ccc01 and year(oha02) = ccc02 and month(oha02) = ccc03 "
        oCommand.CommandText += "left join oea_file on ohb31 = oea01 where year(oha02) =" & tYear & " and month(oha02) <= " & tMonth & " and ohapost = 'Y' and ta_oea01 = 'Y' and ohb09 not in (select jce02 from jce_file) "
        If Not String.IsNullOrEmpty(DNP) Then
            oCommand.CommandText += " AND ohb04 like '" & DNP & "%' "
        End If
        oCommand.CommandText += "group by oha08,oha02,ohb04, ima02,ima021,ima06,ima08,ohb05 ) group by oga08,oga02,ogb04,ima02,ima021,ima06,ima08,ogb05 order by ogb04,oga02"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("oga08")
                Ws.Cells(LineZ, 3) = oReader.Item("oga02")
                Ws.Cells(LineZ, 4) = oReader.Item("ogb04")
                Ws.Cells(LineZ, 5) = oReader.Item("ima02")
                Ws.Cells(LineZ, 6) = oReader.Item("ima021")
                Ws.Cells(LineZ, 7) = oReader.Item("ima06")
                Ws.Cells(LineZ, 8) = oReader.Item("ima08")
                Ws.Cells(LineZ, 9) = oReader.Item("ogb05")
                Ws.Cells(LineZ, 11) = oReader.Item("t1")
                Ws.Cells(LineZ, 12) = "=IFERROR(M" & LineZ & "/K" & LineZ & ",0)"
                Ws.Cells(LineZ, 13) = oReader.Item("t2")
                Ws.Cells(LineZ, 14) = "=IFERROR(O" & LineZ & "/K" & LineZ & ",0)"
                Ws.Cells(LineZ, 15) = oReader.Item("t3")
                Ws.Cells(LineZ, 16) = "=M" & LineZ & "-O" & LineZ
                Ws.Cells(LineZ, 17) = "=IFERROR(P" & LineZ & "/M" & LineZ & ",0)"
                Ws.Cells(LineZ, 18) = oReader.Item("t4")
                Ws.Cells(LineZ, 19) = "=IFERROR(T" & LineZ & "/R" & LineZ & ",0)"
                Ws.Cells(LineZ, 20) = oReader.Item("t5")
                Ws.Cells(LineZ, 21) = "=IFERROR(V" & LineZ & "/R" & LineZ & ",0)"
                Ws.Cells(LineZ, 22) = oReader.Item("t6")
                Ws.Cells(LineZ, 23) = "=T" & LineZ & "-V" & LineZ
                Ws.Cells(LineZ, 24) = "=IFERROR(W" & LineZ & "/T" & LineZ & ",0)"
                Ws.Cells(LineZ, 25) = "=R" & LineZ & "-K" & LineZ
                Ws.Cells(LineZ, 26) = "=T" & LineZ & "-M" & LineZ
                Ws.Cells(LineZ, 27) = "=V" & LineZ & "-O" & LineZ
                Ws.Cells(LineZ, 28) = "=S" & LineZ & "-L" & LineZ
                Ws.Cells(LineZ, 29) = "=U" & LineZ & "-N" & LineZ
                Ws.Cells(LineZ, 30) = "=IFERROR(AB" & LineZ & "/S" & LineZ & ",0)"
                Ws.Cells(LineZ, 31) = "=IFERROR(AC" & LineZ & "/U" & LineZ & ",0)"
                Ws.Cells(LineZ, 32) = "=X" & LineZ & "-Q" & LineZ
                LineZ += 1
            End While
            ' 加總
            Ws.Cells(LineZ, 2) = "Total:"
            Ws.Cells(LineZ, 11) = "=SUM(K6:K" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 13) = "=SUM(M6:M" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 15) = "=SUM(O6:O" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 16) = "=SUM(P6:P" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 17) = "=IFERROR(P" & LineZ & "/M" & LineZ & ",0)"
            Ws.Cells(LineZ, 18) = "=SUM(R6:R" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 20) = "=SUM(T6:T" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 22) = "=SUM(V6:V" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 23) = "=SUM(W6:W" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 24) = "=IFERROR(W" & LineZ & "/T" & LineZ & ",0)"
            Ws.Cells(LineZ, 25) = "=SUM(Y6:Y" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 26) = "=SUM(Z6:Z" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 27) = "=SUM(AA6:AA" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 28) = "=S" & LineZ & "-L" & LineZ
            Ws.Cells(LineZ, 29) = "=U" & LineZ & "-N" & LineZ
            Ws.Cells(LineZ, 32) = "=X" & LineZ & "-Q" & LineZ

            ' 劃線
            oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 32))
            oRng.Borders(xlEdgeLeft).LineStyle = xlNone
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlDouble
            oRng.Borders(xlEdgeRight).LineStyle = xlNone
            oRng.Borders(xlInsideHorizontal).LineStyle = xlNone
            oRng.Borders(xlInsideVertical).LineStyle = xlNone
        End If
        oReader.Close()
        oRng = Ws.Range("C1", "AF1")
        oRng.EntireColumn.AutoFit()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        'Ws.Columns.EntireColumn.ColumnWidth = 20.56
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 4
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 8
        oRng = Ws.Range("B1", "B2")
        oRng.HorizontalAlignment = xlLeft
        Ws.Cells(1, 2) = "公司名称：东莞艾可迅符合材料有限公司"
        Ws.Cells(2, 2) = "币别：人民币"
        Ws.Cells(5, 2) = "出货别"
        Ws.Cells(5, 3) = "产品编号"
        Ws.Cells(5, 4) = "品名"
        Ws.Cells(5, 5) = "规格"
        Ws.Cells(5, 6) = "分群码"
        Ws.Cells(5, 7) = "来源码"
        Ws.Cells(5, 8) = "销售单位"
        Ws.Cells(5, 9) = "注释"
        Ws.Cells(5, 10) = "销售数量"
        Ws.Cells(5, 11) = "销售单价"
        Ws.Cells(5, 12) = "主营业务收入"
        Ws.Cells(5, 13) = "单位成本"
        Ws.Cells(5, 14) = "主营业务成本"
        Ws.Cells(5, 15) = "毛利"
        Ws.Cells(5, 16) = "毛利率"
        Ws.Cells(5, 17) = "销售数量"
        Ws.Cells(5, 18) = "销售单价"
        Ws.Cells(5, 19) = "主营业务收入"
        Ws.Cells(5, 20) = "单位成本"
        Ws.Cells(5, 21) = "主营业务成本"
        Ws.Cells(5, 22) = "毛利"
        Ws.Cells(5, 23) = "毛利率"
        Ws.Cells(5, 24) = "销售数量"
        Ws.Cells(5, 25) = "主营业务收入"
        Ws.Cells(5, 26) = "主营业务成本"
        Ws.Cells(5, 27) = "销售单价"
        Ws.Cells(5, 28) = "单位成本"
        Ws.Cells(5, 29) = "销售单价变化率"
        Ws.Cells(5, 30) = "单位成本变化率"
        Ws.Cells(5, 31) = "毛利率变化"

        oRng = Ws.Range("J4", "P4")
        oRng.Merge()
        Ws.Cells(4, 10) = "YTD" & pYear

        oRng = Ws.Range("Q1", "W1")
        oRng.Merge()
        Ws.Cells(4, 17) = "YTD" & tYear

        oRng = Ws.Range("X1", "AE1")
        oRng.Merge()
        Ws.Cells(4, 24) = "变动分析"

        oRng = Ws.Range("I1", "N1")
        oRng.EntireColumn.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        oRng = Ws.Range("P1", "U1")
        oRng.EntireColumn.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        oRng = Ws.Range("W1", "AA1")
        oRng.EntireColumn.NumberFormat = "#,##0_ ;[Red]-#,##0 "

        oRng = Ws.Range("P1", "P1")
        oRng.EntireColumn.NumberFormat = "0.00%"
        oRng.EntireColumn.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
        oRng.FormatConditions(1).FONT.COLOR = Color.Red

        oRng = Ws.Range("W1", "W1")
        oRng.EntireColumn.NumberFormat = "0.00%"
        oRng.EntireColumn.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
        oRng.FormatConditions(1).FONT.COLOR = Color.Red

        oRng = Ws.Range("AC1", "AE1")
        oRng.EntireColumn.NumberFormat = "0.00%"
        oRng.EntireColumn.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
        oRng.FormatConditions(1).FONT.COLOR = Color.Red

        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.NumberFormat = "@"

        LineZ = 6
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 20.56
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 4
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 8
        oRng = Ws.Range("B1", "B2")
        oRng.HorizontalAlignment = xlLeft
        Ws.Cells(1, 2) = "公司名称：东莞艾可迅符合材料有限公司"
        Ws.Cells(2, 2) = "币别：人民币"
        Ws.Cells(5, 2) = "出货别"
        Ws.Cells(5, 3) = "月份"
        Ws.Cells(5, 4) = "产品编号"
        Ws.Cells(5, 5) = "品名"
        Ws.Cells(5, 6) = "规格"
        Ws.Cells(5, 7) = "分群码"
        Ws.Cells(5, 8) = "来源码"
        Ws.Cells(5, 9) = "销售单位"
        Ws.Cells(5, 10) = "注释"
        Ws.Cells(5, 11) = "销售数量"
        Ws.Cells(5, 12) = "销售单价"
        Ws.Cells(5, 13) = "主营业务收入"
        Ws.Cells(5, 14) = "单位成本"
        Ws.Cells(5, 15) = "主营业务成本"
        Ws.Cells(5, 16) = "毛利"
        Ws.Cells(5, 17) = "毛利率"
        Ws.Cells(5, 18) = "销售数量"
        Ws.Cells(5, 19) = "销售单价"
        Ws.Cells(5, 20) = "主营业务收入"
        Ws.Cells(5, 21) = "单位成本"
        Ws.Cells(5, 22) = "主营业务成本"
        Ws.Cells(5, 23) = "毛利"
        Ws.Cells(5, 24) = "毛利率"
        Ws.Cells(5, 25) = "销售数量"
        Ws.Cells(5, 26) = "主营业务收入"
        Ws.Cells(5, 27) = "主营业务成本"
        Ws.Cells(5, 28) = "销售单价"
        Ws.Cells(5, 29) = "单位成本"
        Ws.Cells(5, 30) = "销售单价变化率"
        Ws.Cells(5, 31) = "单位成本变化率"
        Ws.Cells(5, 32) = "毛利率变化"

        oRng = Ws.Range("K4", "Q4")
        oRng.Merge()
        Ws.Cells(4, 11) = "YTD" & pYear

        oRng = Ws.Range("R1", "X1")
        oRng.Merge()
        Ws.Cells(4, 18) = "YTD" & tYear

        oRng = Ws.Range("Y1", "AF1")
        oRng.Merge()
        Ws.Cells(4, 25) = "变动分析"

        oRng = Ws.Range("K1", "P1")
        oRng.EntireColumn.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        oRng = Ws.Range("R1", "W1")
        oRng.EntireColumn.NumberFormat = "#,##0_ ;[Red]-#,##0 "
        oRng = Ws.Range("Y1", "AC1")
        oRng.EntireColumn.NumberFormat = "#,##0_ ;[Red]-#,##0 "

        oRng = Ws.Range("Q1", "Q1")
        oRng.EntireColumn.NumberFormat = "0.00%"
        oRng.EntireColumn.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
        oRng.FormatConditions(1).FONT.COLOR = Color.Red

        oRng = Ws.Range("X1", "X1")
        oRng.EntireColumn.NumberFormat = "0.00%"
        oRng.EntireColumn.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
        oRng.FormatConditions(1).FONT.COLOR = Color.Red

        oRng = Ws.Range("AD1", "AF1")
        oRng.EntireColumn.NumberFormat = "0.00%"
        oRng.EntireColumn.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, "=0", Type.Missing, Type.Missing)
        oRng.FormatConditions(1).FONT.COLOR = Color.Red

        oRng = Ws.Range("D1", "D1")
        oRng.EntireColumn.NumberFormat = "@"

        LineZ = 6
    End Sub
End Class