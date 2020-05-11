Public Class Form165
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
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim tYear1 As Int16 = 0
    Dim tMonth1 As Int16 = 0
    Dim tYear2 As Int16 = 0
    Dim tMonth2 As Int16 = 0
    Dim sTime1 As Date
    Dim sTime2 As Date
    Dim sTime3 As Date
    Dim eTime1 As Date
    Dim eTime2 As Date
    Dim eTime3 As Date
    Dim USDE As Decimal = 1
    Dim EURE As Decimal = 1
    Dim EURTOUSD As Decimal = 1

    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form165_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        Me.NumericUpDown1.Value = Now.Year
        Me.NumericUpDown2.Value = Now.Month
    End Sub


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
        tYear = Me.NumericUpDown1.Value
        tMonth = Me.NumericUpDown2.Value
        ' 報表月
        sTime1 = Convert.ToDateTime(tYear & "/" & tMonth & "/01")
        eTime1 = sTime1.AddMonths(1).AddDays(-1)
        ' 報表月 -1
        sTime2 = sTime1.AddMonths(-1)
        eTime2 = sTime1.AddDays(-1)
        tYear1 = sTime2.Year
        tMonth1 = sTime2.Month
        ' 報表月 -2
        sTime3 = sTime2.AddMonths(-1)
        eTime3 = sTime2.AddDays(-1)
        tYear2 = sTime3.Year
        tMonth2 = sTime3.Month
        ' 匯率
        oCommand.CommandText = "select nvl(ER, 1) from exchangeratebyyear where year1 = " & tYear & " and currency = 'USD'"
        USDE = oCommand.ExecuteScalar()
        oCommand.CommandText = "select nvl(ER, 1) from exchangeratebyyear where year1 = " & tYear & " and currency = 'EUR'"
        EURE = oCommand.ExecuteScalar()

        EURTOUSD = Decimal.Round((EURE / USDE), 6)

        BackgroundWorker1.RunWorkerAsync()

    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "DAC 销售 成本 毛利"
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
        Dim xPath As String = "C:\temp\DAC 销售 成本 毛利资料模板.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        AdjustExcelFormat()
        LineZ = 7

        oCommand.CommandText = "select ogb04,ima02,ima021,ima25,t1,t2,t3,t4,t5,t6,round((c1.ccc62 * -1 / " & USDE & "),0) cc1,round((c2.ccc62 * -1 / " & USDE & "),0) cc2,round((c3.ccc62 * -1 / " & USDE & "),0) cc3,"
        oCommand.CommandText += "round((c1.ccc62a * -1 / " & USDE & "),0) cc4,round((c2.ccc62a * -1 / " & USDE & "),0) cc5,round((c3.ccc62a * -1 / " & USDE & "),0) cc6,"
        oCommand.CommandText += "round((c1.ccc62b * -1 / " & USDE & "),0) cc7,round((c2.ccc62b * -1 / " & USDE & "),0) cc8,round((c3.ccc62b * -1 / " & USDE & "),0) cc9,"
        oCommand.CommandText += "round(((c1.ccc62c + c1.ccc62d) * -1 / " & USDE & "),0) cc10,round(((c2.ccc62c + c2.ccc62d) * -1 / " & USDE & "),0) cc11,round(((c3.ccc62c + c3.ccc62d) * -1 / " & USDE & "),0) cc12 "
        oCommand.CommandText += "from ( select ogb04,ima02,ima021,ima25,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3, round(sum(t4),0) as t4, round(sum(t5),0) as t5,round(sum(t6),0) as t6 from ( "
        oCommand.CommandText += "select ogb04,ima02,ima021,ima25,(case when oga02 between to_date('" & sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ogb12 else 0 end) as t1,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ogb12 else 0 end) as t2,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ogb12 else 0 end) as t3,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oga23 = 'RMB' THEN ogb14 / " & USDE & " when oga23 = 'EUR' then ogb14 * " & EURTOUSD & " when oga23 = 'USD' then ogb14 end)  else 0 end) as t4,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oga23 = 'RMB' THEN ogb14 / " & USDE & " when oga23 = 'EUR' then ogb14 * " & EURTOUSD & " when oga23 = 'USD' then ogb14 end)  else 0 end) as t5,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oga23 = 'RMB' THEN ogb14 / " & USDE & " when oga23 = 'EUR' then ogb14 * " & EURTOUSD & " when oga23 = 'USD' then ogb14 end)  else 0 end) as t6 "
        oCommand.CommandText += "from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 where ogapost = 'Y' and oga02 between to_date('"
        oCommand.CommandText += sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ogb09 not in (select jce02 from jce_file) "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select ohb04,ima02,ima021,ima25,(case when oha02 between to_date('" & sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ohb12 * -1 else 0 end) as t1,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ohb12 * -1 else 0 end) as t2,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ohb12 * -1 else 0 end) as t3,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oha23 = 'RMB' THEN ohb14 / " & USDE & " * -1 when oha23 = 'EUR' then ohb14 * " & EURTOUSD & " * -1 when oha23 = 'USD' then ohb14 * -1 end)  else 0 end) as t4,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oha23 = 'RMB' THEN ohb14 / " & USDE & " * -1 when oha23 = 'EUR' then ohb14 * " & EURTOUSD & " * -1 when oha23 = 'USD' then ohb14 * -1 end)  else 0 end) as t5,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oha23 = 'RMB' THEN ohb14 / " & USDE & " * -1 when oha23 = 'EUR' then ohb14 * " & EURTOUSD & " * -1 when oha23 = 'USD' then ohb14 * -1 end)  else 0 end) as t6 "
        oCommand.CommandText += "from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 where ohapost = 'Y' and oha02 between to_date('"
        oCommand.CommandText += sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ohb09 not in (select jce02 from jce_file) ) group by ogb04,ima02,ima021,ima25  ) ag "
        oCommand.CommandText += " left join ccc_file c1 on ag.ogb04 = c1.ccc01 and c1.ccc02 = " & tYear2 & " and c1.ccc03 = " & tMonth2
        oCommand.CommandText += " left join ccc_file c2 on ag.ogb04 = c2.ccc01 and c2.ccc02 = " & tYear1 & " and c2.ccc03 = " & tMonth1
        oCommand.CommandText += " left join ccc_file c3 on ag.ogb04 = c3.ccc01 and c3.ccc02 = " & tYear & " and c3.ccc03 = " & tMonth
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("ogb04")
                Ws.Cells(LineZ, 3) = oReader.Item("ima02")
                Ws.Cells(LineZ, 4) = oReader.Item("ima021")
                Ws.Cells(LineZ, 5) = oReader.Item("ima25")
                Ws.Cells(LineZ, 6) = oReader.Item("t1")
                Ws.Cells(LineZ, 7) = oReader.Item("t2")
                Ws.Cells(LineZ, 8) = oReader.Item("t3")
                Ws.Cells(LineZ, 9) = oReader.Item("t4")
                Ws.Cells(LineZ, 10) = oReader.Item("t5")
                Ws.Cells(LineZ, 11) = oReader.Item("t6")
                Ws.Cells(LineZ, 12) = oReader.Item("cc1")
                Ws.Cells(LineZ, 13) = oReader.Item("cc2")
                Ws.Cells(LineZ, 14) = oReader.Item("cc3")
                Ws.Cells(LineZ, 15) = "=I" & LineZ & "-L" & LineZ
                Ws.Cells(LineZ, 16) = "=J" & LineZ & "-M" & LineZ
                Ws.Cells(LineZ, 17) = "=K" & LineZ & "-N" & LineZ
                Ws.Cells(LineZ, 18) = "=IFERROR(O" & LineZ & "/I" & LineZ & ",)"
                Ws.Cells(LineZ, 19) = "=IFERROR(P" & LineZ & "/J" & LineZ & ",)"
                Ws.Cells(LineZ, 20) = "=IFERROR(Q" & LineZ & "/K" & LineZ & ",)"
                Ws.Cells(LineZ, 21) = "=IFERROR(I" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 22) = "=IFERROR(J" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 23) = "=IFERROR(K" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 24) = "=U" & LineZ & "-W" & LineZ
                Ws.Cells(LineZ, 25) = "=IFERROR(L" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 26) = "=IFERROR(M" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 27) = "=IFERROR(N" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 28) = "=U" & LineZ & "-Y" & LineZ
                Ws.Cells(LineZ, 29) = "=V" & LineZ & "-Z" & LineZ
                Ws.Cells(LineZ, 30) = "=W" & LineZ & "-AA" & LineZ
                Ws.Cells(LineZ, 31) = "=Y" & LineZ & "-AA" & LineZ
                Ws.Cells(LineZ, 32) = oReader.Item("cc4")
                Ws.Cells(LineZ, 33) = oReader.Item("cc5")
                Ws.Cells(LineZ, 34) = oReader.Item("cc6")
                Ws.Cells(LineZ, 35) = "=IFERROR(AF" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 36) = "=IFERROR(AG" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 37) = "=IFERROR(AH" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 38) = "=AI" & LineZ & "-AK" & LineZ
                Ws.Cells(LineZ, 39) = oReader.Item("cc7")
                Ws.Cells(LineZ, 40) = oReader.Item("cc8")
                Ws.Cells(LineZ, 41) = oReader.Item("cc9")
                Ws.Cells(LineZ, 42) = "=IFERROR(AM" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 43) = "=IFERROR(AN" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 44) = "=IFERROR(AO" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 45) = "=AP" & LineZ & "-AR" & LineZ
                Ws.Cells(LineZ, 46) = oReader.Item("cc10")
                Ws.Cells(LineZ, 47) = oReader.Item("cc11")
                Ws.Cells(LineZ, 48) = oReader.Item("cc12")
                Ws.Cells(LineZ, 49) = "=IFERROR(AT" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 50) = "=IFERROR(AU" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 51) = "=IFERROR(AV" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 52) = "=AW" & LineZ & "-AY" & LineZ
                Ws.Cells(LineZ, 53) = "=IFERROR(AF" & LineZ & "/I" & LineZ & ",)"
                Ws.Cells(LineZ, 54) = "=IFERROR(AG" & LineZ & "/J" & LineZ & ",)"
                Ws.Cells(LineZ, 55) = "=IFERROR(AH" & LineZ & "/K" & LineZ & ",)"
                Ws.Cells(LineZ, 56) = "=IFERROR(AM" & LineZ & "/I" & LineZ & ",)"
                Ws.Cells(LineZ, 57) = "=IFERROR(AN" & LineZ & "/J" & LineZ & ",)"
                Ws.Cells(LineZ, 58) = "=IFERROR(AO" & LineZ & "/K" & LineZ & ",)"
                Ws.Cells(LineZ, 59) = "=IFERROR(AT" & LineZ & "/I" & LineZ & ",)"
                Ws.Cells(LineZ, 60) = "=IFERROR(AU" & LineZ & "/J" & LineZ & ",)"
                Ws.Cells(LineZ, 61) = "=IFERROR(AV" & LineZ & "/K" & LineZ & ",)"
                Ws.Cells(LineZ, 62) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",$T" & LineZ & "-$S" & LineZ & ")"
                Ws.Cells(LineZ, 63) = "=IFERROR(IF($T" & LineZ & ">$S" & LineZ & ","""",-(BC" & LineZ & "-BB" & LineZ & ")/$BJ" & LineZ & "),)"
                Ws.Cells(LineZ, 64) = "=IFERROR(IF($T" & LineZ & ">$S" & LineZ & ","""",-(BF" & LineZ & "-BE" & LineZ & ")/$BJ" & LineZ & "),)"
                Ws.Cells(LineZ, 65) = "=IFERROR(IF($T" & LineZ & ">$S" & LineZ & ","""",-(BI" & LineZ & "-BH" & LineZ & ")/$BJ" & LineZ & "),)"
                LineZ += 1

            End While
        End If
        oReader.Close()
        oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 1))
        oRng.EntireRow.RowHeight = 25.5
        oRng.Font.Bold = True
        Ws.Cells(LineZ, 5) = "合计" & Chr(10) & "Total"
        Ws.Cells(LineZ, 6) = "=SUM(F7:F" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 7) = "=SUM(G7:G" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 8) = "=SUM(H7:H" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 9) = "=SUM(I7:I" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 10) = "=SUM(J7:J" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 11) = "=SUM(K7:K" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 12) = "=SUM(L7:L" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 13) = "=SUM(M7:M" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 14) = "=SUM(N7:N" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 15) = "=SUM(O7:O" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 16) = "=SUM(P7:P" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 17) = "=SUM(Q7:Q" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 18) = "=IFERROR(O" & LineZ & "/I" & LineZ & ",)"
        Ws.Cells(LineZ, 19) = "=IFERROR(P" & LineZ & "/J" & LineZ & ",)"
        Ws.Cells(LineZ, 20) = "=IFERROR(Q" & LineZ & "/K" & LineZ & ",)"

        Ws.Cells(LineZ, 32) = "=SUM(AF7:AF" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 33) = "=SUM(AG7:AG" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 34) = "=SUM(AH7:AH" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 39) = "=SUM(AM7:AM" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 40) = "=SUM(AN7:AN" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 41) = "=SUM(AO7:AO" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 46) = "=SUM(AT7:AT" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 47) = "=SUM(AU7:AU" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 48) = "=SUM(AV7:AV" & LineZ - 1 & ")"

        Ws.Cells(LineZ, 53) = "=IFERROR(AF" & LineZ & "/I" & LineZ & ",)"
        Ws.Cells(LineZ, 54) = "=IFERROR(AG" & LineZ & "/J" & LineZ & ",)"
        Ws.Cells(LineZ, 55) = "=IFERROR(AH" & LineZ & "/K" & LineZ & ",)"
        Ws.Cells(LineZ, 56) = "=IFERROR(AM" & LineZ & "/I" & LineZ & ",)"
        Ws.Cells(LineZ, 57) = "=IFERROR(AN" & LineZ & "/J" & LineZ & ",)"
        Ws.Cells(LineZ, 58) = "=IFERROR(AO" & LineZ & "/K" & LineZ & ",)"
        Ws.Cells(LineZ, 59) = "=IFERROR(AT" & LineZ & "/I" & LineZ & ",)"
        Ws.Cells(LineZ, 60) = "=IFERROR(AU" & LineZ & "/J" & LineZ & ",)"
        Ws.Cells(LineZ, 61) = "=IFERROR(AV" & LineZ & "/K" & LineZ & ",)"
        Ws.Cells(LineZ, 62) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",$T" & LineZ & "-$S" & LineZ & ")"
        Ws.Cells(LineZ, 63) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",-(BC" & LineZ & "-BB" & LineZ & ")/$BJ" & LineZ & ")"
        Ws.Cells(LineZ, 64) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",-(BF" & LineZ & "-BE" & LineZ & ")/$BJ" & LineZ & ")"
        Ws.Cells(LineZ, 65) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",-(BI" & LineZ & "-BH" & LineZ & ")/$BJ" & LineZ & ")"


        ' 第二頁

        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustExcelFormat()
        LineZ = 7

        oCommand.CommandText = "select ogb04,ima02,ima021,ima25,t1,t2,t3,t4,t5,t6,round((c1.ccc23 * t1 / " & USDE & "),0) cc1,round((c2.ccc23 * t2 / " & USDE & "),0) cc2,round((c3.ccc23 * t3 / " & USDE & "),0) cc3,"
        oCommand.CommandText += "round((c1.ccc23a * t1 / " & USDE & "),0) cc4,round((c2.ccc23a * t2 / " & USDE & "),0) cc5,round((c3.ccc23a * t3 / " & USDE & "),0) cc6,"
        oCommand.CommandText += "round((c1.ccc23b * t1 / " & USDE & "),0) cc7,round((c2.ccc23b * t2 / " & USDE & "),0) cc8,round((c3.ccc23b * t3 / " & USDE & "),0) cc9,"
        oCommand.CommandText += "round(((c1.ccc23c + c1.ccc23d) * t1 / " & USDE & "),0) cc10,round(((c2.ccc23c + c2.ccc23d) * t2 / " & USDE & "),0) cc11,round(((c3.ccc23c + c3.ccc23d) * t3 / " & USDE & "),0) cc12 "
        oCommand.CommandText += "from ( select ogb04,ima02,ima021,ima25,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3, round(sum(t4),0) as t4, round(sum(t5),0) as t5,round(sum(t6),0) as t6 from ( "
        oCommand.CommandText += "select ogb04,ima02,ima021,ima25,(case when oga02 between to_date('" & sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ogb12 else 0 end) as t1,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ogb12 else 0 end) as t2,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ogb12 else 0 end) as t3,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oga23 = 'RMB' THEN ogb14 / " & USDE & " when oga23 = 'EUR' then ogb14 * " & EURTOUSD & " when oga23 = 'USD' then ogb14 end)  else 0 end) as t4,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oga23 = 'RMB' THEN ogb14 / " & USDE & " when oga23 = 'EUR' then ogb14 * " & EURTOUSD & " when oga23 = 'USD' then ogb14 end)  else 0 end) as t5,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oga23 = 'RMB' THEN ogb14 / " & USDE & " when oga23 = 'EUR' then ogb14 * " & EURTOUSD & " when oga23 = 'USD' then ogb14 end)  else 0 end) as t6 "
        oCommand.CommandText += "from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 left join oea_file on ogb31 = oea01 where ogapost = 'Y' and oga02 between to_date('"
        oCommand.CommandText += sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ogb09 not in (select jce02 from jce_file) and ta_oea01 = 'Y'"
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select ohb04,ima02,ima021,ima25,(case when oha02 between to_date('" & sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ohb12 * -1 else 0 end) as t1,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ohb12 * -1 else 0 end) as t2,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ohb12 * -1 else 0 end) as t3,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oha23 = 'RMB' THEN ohb14 / " & USDE & " * -1 when oha23 = 'EUR' then ohb14 * " & EURTOUSD & " * -1 when oha23 = 'USD' then ohb14 * -1 end)  else 0 end) as t4,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oha23 = 'RMB' THEN ohb14 / " & USDE & " * -1 when oha23 = 'EUR' then ohb14 * " & EURTOUSD & " * -1 when oha23 = 'USD' then ohb14 * -1 end)  else 0 end) as t5,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oha23 = 'RMB' THEN ohb14 / " & USDE & " * -1 when oha23 = 'EUR' then ohb14 * " & EURTOUSD & " * -1 when oha23 = 'USD' then ohb14 * -1 end)  else 0 end) as t6 "
        oCommand.CommandText += "from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 left join oea_file on ohb33 = oea01 where ohapost = 'Y' and oha02 between to_date('"
        oCommand.CommandText += sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ohb09 not in (select jce02 from jce_file) and ta_oea01 = 'Y' ) group by ogb04,ima02,ima021,ima25  ) ag "
        oCommand.CommandText += " left join ccc_file c1 on ag.ogb04 = c1.ccc01 and c1.ccc02 = " & tYear2 & " and c1.ccc03 = " & tMonth2
        oCommand.CommandText += " left join ccc_file c2 on ag.ogb04 = c2.ccc01 and c2.ccc02 = " & tYear1 & " and c2.ccc03 = " & tMonth1
        oCommand.CommandText += " left join ccc_file c3 on ag.ogb04 = c3.ccc01 and c3.ccc02 = " & tYear & " and c3.ccc03 = " & tMonth
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("ogb04")
                Ws.Cells(LineZ, 3) = oReader.Item("ima02")
                Ws.Cells(LineZ, 4) = oReader.Item("ima021")
                Ws.Cells(LineZ, 5) = oReader.Item("ima25")
                Ws.Cells(LineZ, 6) = oReader.Item("t1")
                Ws.Cells(LineZ, 7) = oReader.Item("t2")
                Ws.Cells(LineZ, 8) = oReader.Item("t3")
                Ws.Cells(LineZ, 9) = oReader.Item("t4")
                Ws.Cells(LineZ, 10) = oReader.Item("t5")
                Ws.Cells(LineZ, 11) = oReader.Item("t6")
                Ws.Cells(LineZ, 12) = oReader.Item("cc1")
                Ws.Cells(LineZ, 13) = oReader.Item("cc2")
                Ws.Cells(LineZ, 14) = oReader.Item("cc3")
                Ws.Cells(LineZ, 15) = "=I" & LineZ & "-L" & LineZ
                Ws.Cells(LineZ, 16) = "=J" & LineZ & "-M" & LineZ
                Ws.Cells(LineZ, 17) = "=K" & LineZ & "-N" & LineZ
                Ws.Cells(LineZ, 18) = "=IFERROR(O" & LineZ & "/I" & LineZ & ",)"
                Ws.Cells(LineZ, 19) = "=IFERROR(P" & LineZ & "/J" & LineZ & ",)"
                Ws.Cells(LineZ, 20) = "=IFERROR(Q" & LineZ & "/K" & LineZ & ",)"
                Ws.Cells(LineZ, 21) = "=IFERROR(I" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 22) = "=IFERROR(J" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 23) = "=IFERROR(K" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 24) = "=U" & LineZ & "-W" & LineZ
                Ws.Cells(LineZ, 25) = "=IFERROR(L" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 26) = "=IFERROR(M" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 27) = "=IFERROR(N" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 28) = "=U" & LineZ & "-Y" & LineZ
                Ws.Cells(LineZ, 29) = "=V" & LineZ & "-Z" & LineZ
                Ws.Cells(LineZ, 30) = "=W" & LineZ & "-AA" & LineZ
                Ws.Cells(LineZ, 31) = "=Y" & LineZ & "-AA" & LineZ
                Ws.Cells(LineZ, 32) = oReader.Item("cc4")
                Ws.Cells(LineZ, 33) = oReader.Item("cc5")
                Ws.Cells(LineZ, 34) = oReader.Item("cc6")
                Ws.Cells(LineZ, 35) = "=IFERROR(AF" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 36) = "=IFERROR(AG" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 37) = "=IFERROR(AH" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 38) = "=AI" & LineZ & "-AK" & LineZ
                Ws.Cells(LineZ, 39) = oReader.Item("cc7")
                Ws.Cells(LineZ, 40) = oReader.Item("cc8")
                Ws.Cells(LineZ, 41) = oReader.Item("cc9")
                Ws.Cells(LineZ, 42) = "=IFERROR(AM" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 43) = "=IFERROR(AN" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 44) = "=IFERROR(AO" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 45) = "=AP" & LineZ & "-AR" & LineZ
                Ws.Cells(LineZ, 46) = oReader.Item("cc10")
                Ws.Cells(LineZ, 47) = oReader.Item("cc11")
                Ws.Cells(LineZ, 48) = oReader.Item("cc12")
                Ws.Cells(LineZ, 49) = "=IFERROR(AT" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 50) = "=IFERROR(AU" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 51) = "=IFERROR(AV" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 52) = "=AW" & LineZ & "-AY" & LineZ
                Ws.Cells(LineZ, 53) = "=IFERROR(AF" & LineZ & "/I" & LineZ & ",)"
                Ws.Cells(LineZ, 54) = "=IFERROR(AG" & LineZ & "/J" & LineZ & ",)"
                Ws.Cells(LineZ, 55) = "=IFERROR(AH" & LineZ & "/K" & LineZ & ",)"
                Ws.Cells(LineZ, 56) = "=IFERROR(AM" & LineZ & "/I" & LineZ & ",)"
                Ws.Cells(LineZ, 57) = "=IFERROR(AN" & LineZ & "/J" & LineZ & ",)"
                Ws.Cells(LineZ, 58) = "=IFERROR(AO" & LineZ & "/K" & LineZ & ",)"
                Ws.Cells(LineZ, 59) = "=IFERROR(AT" & LineZ & "/I" & LineZ & ",)"
                Ws.Cells(LineZ, 60) = "=IFERROR(AU" & LineZ & "/J" & LineZ & ",)"
                Ws.Cells(LineZ, 61) = "=IFERROR(AV" & LineZ & "/K" & LineZ & ",)"
                Ws.Cells(LineZ, 62) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",$T" & LineZ & "-$S" & LineZ & ")"
                Ws.Cells(LineZ, 63) = "=IFERROR(IF($T" & LineZ & ">$S" & LineZ & ","""",-(BC" & LineZ & "-BB" & LineZ & ")/$BJ" & LineZ & "),)"
                Ws.Cells(LineZ, 64) = "=IFERROR(IF($T" & LineZ & ">$S" & LineZ & ","""",-(BF" & LineZ & "-BE" & LineZ & ")/$BJ" & LineZ & "),)"
                Ws.Cells(LineZ, 65) = "=IFERROR(IF($T" & LineZ & ">$S" & LineZ & ","""",-(BI" & LineZ & "-BH" & LineZ & ")/$BJ" & LineZ & "),)"
                LineZ += 1

            End While
        End If
        oReader.Close()
        oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 1))
        oRng.EntireRow.RowHeight = 25.5
        oRng.Font.Bold = True
        Ws.Cells(LineZ, 5) = "合计" & Chr(10) & "Total"
        Ws.Cells(LineZ, 6) = "=SUM(F7:F" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 7) = "=SUM(G7:G" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 8) = "=SUM(H7:H" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 9) = "=SUM(I7:I" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 10) = "=SUM(J7:J" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 11) = "=SUM(K7:K" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 12) = "=SUM(L7:L" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 13) = "=SUM(M7:M" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 14) = "=SUM(N7:N" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 15) = "=SUM(O7:O" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 16) = "=SUM(P7:P" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 17) = "=SUM(Q7:Q" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 18) = "=IFERROR(O" & LineZ & "/I" & LineZ & ",)"
        Ws.Cells(LineZ, 19) = "=IFERROR(P" & LineZ & "/J" & LineZ & ",)"
        Ws.Cells(LineZ, 20) = "=IFERROR(Q" & LineZ & "/K" & LineZ & ",)"

        Ws.Cells(LineZ, 32) = "=SUM(AF7:AF" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 33) = "=SUM(AG7:AG" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 34) = "=SUM(AH7:AH" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 39) = "=SUM(AM7:AM" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 40) = "=SUM(AN7:AN" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 41) = "=SUM(AO7:AO" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 46) = "=SUM(AT7:AT" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 47) = "=SUM(AU7:AU" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 48) = "=SUM(AV7:AV" & LineZ - 1 & ")"

        Ws.Cells(LineZ, 53) = "=IFERROR(AF" & LineZ & "/I" & LineZ & ",)"
        Ws.Cells(LineZ, 54) = "=IFERROR(AG" & LineZ & "/J" & LineZ & ",)"
        Ws.Cells(LineZ, 55) = "=IFERROR(AH" & LineZ & "/K" & LineZ & ",)"
        Ws.Cells(LineZ, 56) = "=IFERROR(AM" & LineZ & "/I" & LineZ & ",)"
        Ws.Cells(LineZ, 57) = "=IFERROR(AN" & LineZ & "/J" & LineZ & ",)"
        Ws.Cells(LineZ, 58) = "=IFERROR(AO" & LineZ & "/K" & LineZ & ",)"
        Ws.Cells(LineZ, 59) = "=IFERROR(AT" & LineZ & "/I" & LineZ & ",)"
        Ws.Cells(LineZ, 60) = "=IFERROR(AU" & LineZ & "/J" & LineZ & ",)"
        Ws.Cells(LineZ, 61) = "=IFERROR(AV" & LineZ & "/K" & LineZ & ",)"
        Ws.Cells(LineZ, 62) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",$T" & LineZ & "-$S" & LineZ & ")"
        Ws.Cells(LineZ, 63) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",-(BC" & LineZ & "-BB" & LineZ & ")/$BJ" & LineZ & ")"
        Ws.Cells(LineZ, 64) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",-(BF" & LineZ & "-BE" & LineZ & ")/$BJ" & LineZ & ")"
        Ws.Cells(LineZ, 65) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",-(BI" & LineZ & "-BH" & LineZ & ")/$BJ" & LineZ & ")"


        ' 第三頁

        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        AdjustExcelFormat()
        LineZ = 7

        oCommand.CommandText = "select ogb04,ima02,ima021,ima25,t1,t2,t3,t4,t5,t6,round((c1.ccc23 * t1 / " & USDE & "),0) cc1,round((c2.ccc23 * t2 / " & USDE & "),0) cc2,round((c3.ccc23 * t3 / " & USDE & "),0) cc3,"
        oCommand.CommandText += "round((c1.ccc23a * t1 / " & USDE & "),0) cc4,round((c2.ccc23a * t2 / " & USDE & "),0) cc5,round((c3.ccc23a * t3 / " & USDE & "),0) cc6,"
        oCommand.CommandText += "round((c1.ccc23b * t1 / " & USDE & "),0) cc7,round((c2.ccc23b * t2 / " & USDE & "),0) cc8,round((c3.ccc23b * t3 / " & USDE & "),0) cc9,"
        oCommand.CommandText += "round(((c1.ccc23c + c1.ccc23d) * t1 / " & USDE & "),0) cc10,round(((c2.ccc23c + c2.ccc23d) * t2 / " & USDE & "),0) cc11,round(((c3.ccc23c + c3.ccc23d) * t3 / " & USDE & "),0) cc12 "
        oCommand.CommandText += "from ( select ogb04,ima02,ima021,ima25,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3, round(sum(t4),0) as t4, round(sum(t5),0) as t5,round(sum(t6),0) as t6 from ( "
        oCommand.CommandText += "select ogb04,ima02,ima021,ima25,(case when oga02 between to_date('" & sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ogb12 else 0 end) as t1,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ogb12 else 0 end) as t2,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ogb12 else 0 end) as t3,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oga23 = 'RMB' THEN ogb14 / " & USDE & " when oga23 = 'EUR' then ogb14 * " & EURTOUSD & " when oga23 = 'USD' then ogb14 end)  else 0 end) as t4,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oga23 = 'RMB' THEN ogb14 / " & USDE & " when oga23 = 'EUR' then ogb14 * " & EURTOUSD & " when oga23 = 'USD' then ogb14 end)  else 0 end) as t5,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oga23 = 'RMB' THEN ogb14 / " & USDE & " when oga23 = 'EUR' then ogb14 * " & EURTOUSD & " when oga23 = 'USD' then ogb14 end)  else 0 end) as t6 "
        oCommand.CommandText += "from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 left join oea_file on ogb31 = oea01 where ogapost = 'Y' and oga02 between to_date('"
        oCommand.CommandText += sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ogb09 not in (select jce02 from jce_file) and ta_oea01 = 'N'"
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select ohb04,ima02,ima021,ima25,(case when oha02 between to_date('" & sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ohb12 * -1 else 0 end) as t1,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ohb12 * -1 else 0 end) as t2,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ohb12 * -1 else 0 end) as t3,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oha23 = 'RMB' THEN ohb14 / " & USDE & " * -1 when oha23 = 'EUR' then ohb14 * " & EURTOUSD & " * -1 when oha23 = 'USD' then ohb14 * -1 end)  else 0 end) as t4,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oha23 = 'RMB' THEN ohb14 / " & USDE & " * -1 when oha23 = 'EUR' then ohb14 * " & EURTOUSD & " * -1 when oha23 = 'USD' then ohb14 * -1 end)  else 0 end) as t5,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oha23 = 'RMB' THEN ohb14 / " & USDE & " * -1 when oha23 = 'EUR' then ohb14 * " & EURTOUSD & " * -1 when oha23 = 'USD' then ohb14 * -1 end)  else 0 end) as t6 "
        oCommand.CommandText += "from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 left join oea_file on ohb33 = oea01 where ohapost = 'Y' and oha02 between to_date('"
        oCommand.CommandText += sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ohb09 not in (select jce02 from jce_file) and ta_oea01 = 'N' ) group by ogb04,ima02,ima021,ima25  ) ag "
        oCommand.CommandText += " left join ccc_file c1 on ag.ogb04 = c1.ccc01 and c1.ccc02 = " & tYear2 & " and c1.ccc03 = " & tMonth2
        oCommand.CommandText += " left join ccc_file c2 on ag.ogb04 = c2.ccc01 and c2.ccc02 = " & tYear1 & " and c2.ccc03 = " & tMonth1
        oCommand.CommandText += " left join ccc_file c3 on ag.ogb04 = c3.ccc01 and c3.ccc02 = " & tYear & " and c3.ccc03 = " & tMonth
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("ogb04")
                Ws.Cells(LineZ, 3) = oReader.Item("ima02")
                Ws.Cells(LineZ, 4) = oReader.Item("ima021")
                Ws.Cells(LineZ, 5) = oReader.Item("ima25")
                Ws.Cells(LineZ, 6) = oReader.Item("t1")
                Ws.Cells(LineZ, 7) = oReader.Item("t2")
                Ws.Cells(LineZ, 8) = oReader.Item("t3")
                Ws.Cells(LineZ, 9) = oReader.Item("t4")
                Ws.Cells(LineZ, 10) = oReader.Item("t5")
                Ws.Cells(LineZ, 11) = oReader.Item("t6")
                Ws.Cells(LineZ, 12) = oReader.Item("cc1")
                Ws.Cells(LineZ, 13) = oReader.Item("cc2")
                Ws.Cells(LineZ, 14) = oReader.Item("cc3")
                Ws.Cells(LineZ, 15) = "=I" & LineZ & "-L" & LineZ
                Ws.Cells(LineZ, 16) = "=J" & LineZ & "-M" & LineZ
                Ws.Cells(LineZ, 17) = "=K" & LineZ & "-N" & LineZ
                Ws.Cells(LineZ, 18) = "=IFERROR(O" & LineZ & "/I" & LineZ & ",)"
                Ws.Cells(LineZ, 19) = "=IFERROR(P" & LineZ & "/J" & LineZ & ",)"
                Ws.Cells(LineZ, 20) = "=IFERROR(Q" & LineZ & "/K" & LineZ & ",)"
                Ws.Cells(LineZ, 21) = "=IFERROR(I" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 22) = "=IFERROR(J" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 23) = "=IFERROR(K" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 24) = "=U" & LineZ & "-W" & LineZ
                Ws.Cells(LineZ, 25) = "=IFERROR(L" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 26) = "=IFERROR(M" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 27) = "=IFERROR(N" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 28) = "=U" & LineZ & "-Y" & LineZ
                Ws.Cells(LineZ, 29) = "=V" & LineZ & "-Z" & LineZ
                Ws.Cells(LineZ, 30) = "=W" & LineZ & "-AA" & LineZ
                Ws.Cells(LineZ, 31) = "=Y" & LineZ & "-AA" & LineZ
                Ws.Cells(LineZ, 32) = oReader.Item("cc4")
                Ws.Cells(LineZ, 33) = oReader.Item("cc5")
                Ws.Cells(LineZ, 34) = oReader.Item("cc6")
                Ws.Cells(LineZ, 35) = "=IFERROR(AF" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 36) = "=IFERROR(AG" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 37) = "=IFERROR(AH" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 38) = "=AI" & LineZ & "-AK" & LineZ
                Ws.Cells(LineZ, 39) = oReader.Item("cc7")
                Ws.Cells(LineZ, 40) = oReader.Item("cc8")
                Ws.Cells(LineZ, 41) = oReader.Item("cc9")
                Ws.Cells(LineZ, 42) = "=IFERROR(AM" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 43) = "=IFERROR(AN" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 44) = "=IFERROR(AO" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 45) = "=AP" & LineZ & "-AR" & LineZ
                Ws.Cells(LineZ, 46) = oReader.Item("cc10")
                Ws.Cells(LineZ, 47) = oReader.Item("cc11")
                Ws.Cells(LineZ, 48) = oReader.Item("cc12")
                Ws.Cells(LineZ, 49) = "=IFERROR(AT" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 50) = "=IFERROR(AU" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 51) = "=IFERROR(AV" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 52) = "=AW" & LineZ & "-AY" & LineZ
                Ws.Cells(LineZ, 53) = "=IFERROR(AF" & LineZ & "/I" & LineZ & ",)"
                Ws.Cells(LineZ, 54) = "=IFERROR(AG" & LineZ & "/J" & LineZ & ",)"
                Ws.Cells(LineZ, 55) = "=IFERROR(AH" & LineZ & "/K" & LineZ & ",)"
                Ws.Cells(LineZ, 56) = "=IFERROR(AM" & LineZ & "/I" & LineZ & ",)"
                Ws.Cells(LineZ, 57) = "=IFERROR(AN" & LineZ & "/J" & LineZ & ",)"
                Ws.Cells(LineZ, 58) = "=IFERROR(AO" & LineZ & "/K" & LineZ & ",)"
                Ws.Cells(LineZ, 59) = "=IFERROR(AT" & LineZ & "/I" & LineZ & ",)"
                Ws.Cells(LineZ, 60) = "=IFERROR(AU" & LineZ & "/J" & LineZ & ",)"
                Ws.Cells(LineZ, 61) = "=IFERROR(AV" & LineZ & "/K" & LineZ & ",)"
                Ws.Cells(LineZ, 62) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",$T" & LineZ & "-$S" & LineZ & ")"
                Ws.Cells(LineZ, 63) = "=IFERROR(IF($T" & LineZ & ">$S" & LineZ & ","""",-(BC" & LineZ & "-BB" & LineZ & ")/$BJ" & LineZ & "),)"
                Ws.Cells(LineZ, 64) = "=IFERROR(IF($T" & LineZ & ">$S" & LineZ & ","""",-(BF" & LineZ & "-BE" & LineZ & ")/$BJ" & LineZ & "),)"
                Ws.Cells(LineZ, 65) = "=IFERROR(IF($T" & LineZ & ">$S" & LineZ & ","""",-(BI" & LineZ & "-BH" & LineZ & ")/$BJ" & LineZ & "),)"
                LineZ += 1

            End While
        End If
        oReader.Close()
        oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 1))
        oRng.EntireRow.RowHeight = 25.5
        oRng.Font.Bold = True
        Ws.Cells(LineZ, 5) = "合计" & Chr(10) & "Total"
        Ws.Cells(LineZ, 6) = "=SUM(F7:F" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 7) = "=SUM(G7:G" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 8) = "=SUM(H7:H" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 9) = "=SUM(I7:I" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 10) = "=SUM(J7:J" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 11) = "=SUM(K7:K" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 12) = "=SUM(L7:L" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 13) = "=SUM(M7:M" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 14) = "=SUM(N7:N" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 15) = "=SUM(O7:O" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 16) = "=SUM(P7:P" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 17) = "=SUM(Q7:Q" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 18) = "=IFERROR(O" & LineZ & "/I" & LineZ & ",)"
        Ws.Cells(LineZ, 19) = "=IFERROR(P" & LineZ & "/J" & LineZ & ",)"
        Ws.Cells(LineZ, 20) = "=IFERROR(Q" & LineZ & "/K" & LineZ & ",)"

        Ws.Cells(LineZ, 32) = "=SUM(AF7:AF" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 33) = "=SUM(AG7:AG" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 34) = "=SUM(AH7:AH" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 39) = "=SUM(AM7:AM" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 40) = "=SUM(AN7:AN" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 41) = "=SUM(AO7:AO" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 46) = "=SUM(AT7:AT" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 47) = "=SUM(AU7:AU" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 48) = "=SUM(AV7:AV" & LineZ - 1 & ")"

        Ws.Cells(LineZ, 53) = "=IFERROR(AF" & LineZ & "/I" & LineZ & ",)"
        Ws.Cells(LineZ, 54) = "=IFERROR(AG" & LineZ & "/J" & LineZ & ",)"
        Ws.Cells(LineZ, 55) = "=IFERROR(AH" & LineZ & "/K" & LineZ & ",)"
        Ws.Cells(LineZ, 56) = "=IFERROR(AM" & LineZ & "/I" & LineZ & ",)"
        Ws.Cells(LineZ, 57) = "=IFERROR(AN" & LineZ & "/J" & LineZ & ",)"
        Ws.Cells(LineZ, 58) = "=IFERROR(AO" & LineZ & "/K" & LineZ & ",)"
        Ws.Cells(LineZ, 59) = "=IFERROR(AT" & LineZ & "/I" & LineZ & ",)"
        Ws.Cells(LineZ, 60) = "=IFERROR(AU" & LineZ & "/J" & LineZ & ",)"
        Ws.Cells(LineZ, 61) = "=IFERROR(AV" & LineZ & "/K" & LineZ & ",)"
        Ws.Cells(LineZ, 62) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",$T" & LineZ & "-$S" & LineZ & ")"
        Ws.Cells(LineZ, 63) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",-(BC" & LineZ & "-BB" & LineZ & ")/$BJ" & LineZ & ")"
        Ws.Cells(LineZ, 64) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",-(BF" & LineZ & "-BE" & LineZ & ")/$BJ" & LineZ & ")"
        Ws.Cells(LineZ, 65) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",-(BI" & LineZ & "-BH" & LineZ & ")/$BJ" & LineZ & ")"



        ' 第四頁

        Ws = xWorkBook.Sheets(4)
        Ws.Activate()
        AdjustExcelFormat()
        LineZ = 7

        oCommand.CommandText = "select ogb04,ima02,ima021,ima25,t1,t2,t3,t4,t5,t6,round((c1.ccc23 * t1 / " & USDE & "),0) cc1,round((c2.ccc23 * t2 / " & USDE & "),0) cc2,round((c3.ccc23 * t3 / " & USDE & "),0) cc3,"
        oCommand.CommandText += "round((c1.ccc23a * t1 / " & USDE & "),0) cc4,round((c2.ccc23a * t2 / " & USDE & "),0) cc5,round((c3.ccc23a * t3 / " & USDE & "),0) cc6,"
        oCommand.CommandText += "round((c1.ccc23b * t1 / " & USDE & "),0) cc7,round((c2.ccc23b * t2 / " & USDE & "),0) cc8,round((c3.ccc23b * t3 / " & USDE & "),0) cc9,"
        oCommand.CommandText += "round(((c1.ccc23c + c1.ccc23d) * t1 / " & USDE & "),0) cc10,round(((c2.ccc23c + c2.ccc23d) * t2 / " & USDE & "),0) cc11,round(((c3.ccc23c + c3.ccc23d) * t3 / " & USDE & "),0) cc12 "
        oCommand.CommandText += "from ( select ogb04,ima02,ima021,ima25,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3, round(sum(t4),0) as t4, round(sum(t5),0) as t5,round(sum(t6),0) as t6 from ( "
        oCommand.CommandText += "select ogb04,ima02,ima021,ima25,(case when oga02 between to_date('" & sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ogb12 else 0 end) as t1,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ogb12 else 0 end) as t2,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ogb12 else 0 end) as t3,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oga23 = 'RMB' THEN ogb14 / " & USDE & " when oga23 = 'EUR' then ogb14 * " & EURTOUSD & " when oga23 = 'USD' then ogb14 end)  else 0 end) as t4,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oga23 = 'RMB' THEN ogb14 / " & USDE & " when oga23 = 'EUR' then ogb14 * " & EURTOUSD & " when oga23 = 'USD' then ogb14 end)  else 0 end) as t5,"
        oCommand.CommandText += "(case when oga02 between to_date('" & sTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oga23 = 'RMB' THEN ogb14 / " & USDE & " when oga23 = 'EUR' then ogb14 * " & EURTOUSD & " when oga23 = 'USD' then ogb14 end)  else 0 end) as t6 "
        oCommand.CommandText += "from ogb_file left join oga_file on ogb01 = oga01 left join ima_file on ogb04 = ima01 left join oea_file on ogb31 = oea01 where ogapost = 'Y' and oga02 between to_date('"
        oCommand.CommandText += sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ogb09 not in (select jce02 from jce_file) and ta_oea01 IS NULL "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select ohb04,ima02,ima021,ima25,(case when oha02 between to_date('" & sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ohb12 * -1 else 0 end) as t1,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ohb12 * -1 else 0 end) as t2,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then ohb12 * -1 else 0 end) as t3,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oha23 = 'RMB' THEN ohb14 / " & USDE & " * -1 when oha23 = 'EUR' then ohb14 * " & EURTOUSD & " * -1 when oha23 = 'USD' then ohb14 * -1 end)  else 0 end) as t4,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oha23 = 'RMB' THEN ohb14 / " & USDE & " * -1 when oha23 = 'EUR' then ohb14 * " & EURTOUSD & " * -1 when oha23 = 'USD' then ohb14 * -1 end)  else 0 end) as t5,"
        oCommand.CommandText += "(case when oha02 between to_date('" & sTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') then "
        oCommand.CommandText += "(case when oha23 = 'RMB' THEN ohb14 / " & USDE & " * -1 when oha23 = 'EUR' then ohb14 * " & EURTOUSD & " * -1 when oha23 = 'USD' then ohb14 * -1 end)  else 0 end) as t6 "
        oCommand.CommandText += "from ohb_file left join oha_file on ohb01 = oha01 left join ima_file on ohb04 = ima01 left join oea_file on ohb33 = oea01 where ohapost = 'Y' and oha02 between to_date('"
        oCommand.CommandText += sTime3.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eTime1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ohb09 not in (select jce02 from jce_file) and ta_oea01 IS NULL ) group by ogb04,ima02,ima021,ima25  ) ag "
        oCommand.CommandText += " left join ccc_file c1 on ag.ogb04 = c1.ccc01 and c1.ccc02 = " & tYear2 & " and c1.ccc03 = " & tMonth2
        oCommand.CommandText += " left join ccc_file c2 on ag.ogb04 = c2.ccc01 and c2.ccc02 = " & tYear1 & " and c2.ccc03 = " & tMonth1
        oCommand.CommandText += " left join ccc_file c3 on ag.ogb04 = c3.ccc01 and c3.ccc02 = " & tYear & " and c3.ccc03 = " & tMonth
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("ogb04")
                Ws.Cells(LineZ, 3) = oReader.Item("ima02")
                Ws.Cells(LineZ, 4) = oReader.Item("ima021")
                Ws.Cells(LineZ, 5) = oReader.Item("ima25")
                Ws.Cells(LineZ, 6) = oReader.Item("t1")
                Ws.Cells(LineZ, 7) = oReader.Item("t2")
                Ws.Cells(LineZ, 8) = oReader.Item("t3")
                Ws.Cells(LineZ, 9) = oReader.Item("t4")
                Ws.Cells(LineZ, 10) = oReader.Item("t5")
                Ws.Cells(LineZ, 11) = oReader.Item("t6")
                Ws.Cells(LineZ, 12) = oReader.Item("cc1")
                Ws.Cells(LineZ, 13) = oReader.Item("cc2")
                Ws.Cells(LineZ, 14) = oReader.Item("cc3")
                Ws.Cells(LineZ, 15) = "=I" & LineZ & "-L" & LineZ
                Ws.Cells(LineZ, 16) = "=J" & LineZ & "-M" & LineZ
                Ws.Cells(LineZ, 17) = "=K" & LineZ & "-N" & LineZ
                Ws.Cells(LineZ, 18) = "=IFERROR(O" & LineZ & "/I" & LineZ & ",)"
                Ws.Cells(LineZ, 19) = "=IFERROR(P" & LineZ & "/J" & LineZ & ",)"
                Ws.Cells(LineZ, 20) = "=IFERROR(Q" & LineZ & "/K" & LineZ & ",)"
                Ws.Cells(LineZ, 21) = "=IFERROR(I" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 22) = "=IFERROR(J" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 23) = "=IFERROR(K" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 24) = "=U" & LineZ & "-W" & LineZ
                Ws.Cells(LineZ, 25) = "=IFERROR(L" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 26) = "=IFERROR(M" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 27) = "=IFERROR(N" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 28) = "=U" & LineZ & "-Y" & LineZ
                Ws.Cells(LineZ, 29) = "=V" & LineZ & "-Z" & LineZ
                Ws.Cells(LineZ, 30) = "=W" & LineZ & "-AA" & LineZ
                Ws.Cells(LineZ, 31) = "=Y" & LineZ & "-AA" & LineZ
                Ws.Cells(LineZ, 32) = oReader.Item("cc4")
                Ws.Cells(LineZ, 33) = oReader.Item("cc5")
                Ws.Cells(LineZ, 34) = oReader.Item("cc6")
                Ws.Cells(LineZ, 35) = "=IFERROR(AF" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 36) = "=IFERROR(AG" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 37) = "=IFERROR(AH" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 38) = "=AI" & LineZ & "-AK" & LineZ
                Ws.Cells(LineZ, 39) = oReader.Item("cc7")
                Ws.Cells(LineZ, 40) = oReader.Item("cc8")
                Ws.Cells(LineZ, 41) = oReader.Item("cc9")
                Ws.Cells(LineZ, 42) = "=IFERROR(AM" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 43) = "=IFERROR(AN" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 44) = "=IFERROR(AO" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 45) = "=AP" & LineZ & "-AR" & LineZ
                Ws.Cells(LineZ, 46) = oReader.Item("cc10")
                Ws.Cells(LineZ, 47) = oReader.Item("cc11")
                Ws.Cells(LineZ, 48) = oReader.Item("cc12")
                Ws.Cells(LineZ, 49) = "=IFERROR(AT" & LineZ & "/F" & LineZ & ",)"
                Ws.Cells(LineZ, 50) = "=IFERROR(AU" & LineZ & "/G" & LineZ & ",)"
                Ws.Cells(LineZ, 51) = "=IFERROR(AV" & LineZ & "/H" & LineZ & ",)"
                Ws.Cells(LineZ, 52) = "=AW" & LineZ & "-AY" & LineZ
                Ws.Cells(LineZ, 53) = "=IFERROR(AF" & LineZ & "/I" & LineZ & ",)"
                Ws.Cells(LineZ, 54) = "=IFERROR(AG" & LineZ & "/J" & LineZ & ",)"
                Ws.Cells(LineZ, 55) = "=IFERROR(AH" & LineZ & "/K" & LineZ & ",)"
                Ws.Cells(LineZ, 56) = "=IFERROR(AM" & LineZ & "/I" & LineZ & ",)"
                Ws.Cells(LineZ, 57) = "=IFERROR(AN" & LineZ & "/J" & LineZ & ",)"
                Ws.Cells(LineZ, 58) = "=IFERROR(AO" & LineZ & "/K" & LineZ & ",)"
                Ws.Cells(LineZ, 59) = "=IFERROR(AT" & LineZ & "/I" & LineZ & ",)"
                Ws.Cells(LineZ, 60) = "=IFERROR(AU" & LineZ & "/J" & LineZ & ",)"
                Ws.Cells(LineZ, 61) = "=IFERROR(AV" & LineZ & "/K" & LineZ & ",)"
                Ws.Cells(LineZ, 62) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",$T" & LineZ & "-$S" & LineZ & ")"
                Ws.Cells(LineZ, 63) = "=IFERROR(IF($T" & LineZ & ">$S" & LineZ & ","""",-(BC" & LineZ & "-BB" & LineZ & ")/$BJ" & LineZ & "),)"
                Ws.Cells(LineZ, 64) = "=IFERROR(IF($T" & LineZ & ">$S" & LineZ & ","""",-(BF" & LineZ & "-BE" & LineZ & ")/$BJ" & LineZ & "),)"
                Ws.Cells(LineZ, 65) = "=IFERROR(IF($T" & LineZ & ">$S" & LineZ & ","""",-(BI" & LineZ & "-BH" & LineZ & ")/$BJ" & LineZ & "),)"
                LineZ += 1

            End While
            oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 1))
            oRng.EntireRow.RowHeight = 25.5
            oRng.Font.Bold = True
            Ws.Cells(LineZ, 5) = "合计" & Chr(10) & "Total"
            Ws.Cells(LineZ, 6) = "=SUM(F7:F" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 7) = "=SUM(G7:G" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 8) = "=SUM(H7:H" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 9) = "=SUM(I7:I" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 10) = "=SUM(J7:J" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 11) = "=SUM(K7:K" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 12) = "=SUM(L7:L" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 13) = "=SUM(M7:M" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 14) = "=SUM(N7:N" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 15) = "=SUM(O7:O" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 16) = "=SUM(P7:P" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 17) = "=SUM(Q7:Q" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 18) = "=IFERROR(O" & LineZ & "/I" & LineZ & ",)"
            Ws.Cells(LineZ, 19) = "=IFERROR(P" & LineZ & "/J" & LineZ & ",)"
            Ws.Cells(LineZ, 20) = "=IFERROR(Q" & LineZ & "/K" & LineZ & ",)"

            Ws.Cells(LineZ, 32) = "=SUM(AF7:AF" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 33) = "=SUM(AG7:AG" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 34) = "=SUM(AH7:AH" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 39) = "=SUM(AM7:AM" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 40) = "=SUM(AN7:AN" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 41) = "=SUM(AO7:AO" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 46) = "=SUM(AT7:AT" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 47) = "=SUM(AU7:AU" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 48) = "=SUM(AV7:AV" & LineZ - 1 & ")"

            Ws.Cells(LineZ, 53) = "=IFERROR(AF" & LineZ & "/I" & LineZ & ",)"
            Ws.Cells(LineZ, 54) = "=IFERROR(AG" & LineZ & "/J" & LineZ & ",)"
            Ws.Cells(LineZ, 55) = "=IFERROR(AH" & LineZ & "/K" & LineZ & ",)"
            Ws.Cells(LineZ, 56) = "=IFERROR(AM" & LineZ & "/I" & LineZ & ",)"
            Ws.Cells(LineZ, 57) = "=IFERROR(AN" & LineZ & "/J" & LineZ & ",)"
            Ws.Cells(LineZ, 58) = "=IFERROR(AO" & LineZ & "/K" & LineZ & ",)"
            Ws.Cells(LineZ, 59) = "=IFERROR(AT" & LineZ & "/I" & LineZ & ",)"
            Ws.Cells(LineZ, 60) = "=IFERROR(AU" & LineZ & "/J" & LineZ & ",)"
            Ws.Cells(LineZ, 61) = "=IFERROR(AV" & LineZ & "/K" & LineZ & ",)"
            Ws.Cells(LineZ, 62) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",$T" & LineZ & "-$S" & LineZ & ")"
            Ws.Cells(LineZ, 63) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",-(BC" & LineZ & "-BB" & LineZ & ")/$BJ" & LineZ & ")"
            Ws.Cells(LineZ, 64) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",-(BF" & LineZ & "-BE" & LineZ & ")/$BJ" & LineZ & ")"
            Ws.Cells(LineZ, 65) = "=IF($T" & LineZ & ">$S" & LineZ & ","""",-(BI" & LineZ & "-BH" & LineZ & ")/$BJ" & LineZ & ")"
        End If
        oReader.Close()
        
    End Sub
    Private Sub AdjustExcelFormat()
        Ws.Cells(6, 6) = tMonth2 & "月"
        Ws.Cells(6, 7) = tMonth1 & "月"
        Ws.Cells(6, 8) = tMonth & "月"
        Ws.Cells(6, 9) = tMonth2 & "月"
        Ws.Cells(6, 10) = tMonth1 & "月"
        Ws.Cells(6, 11) = tMonth & "月"
        Ws.Cells(6, 12) = tMonth2 & "月"
        Ws.Cells(6, 13) = tMonth1 & "月"
        Ws.Cells(6, 14) = tMonth & "月"
        Ws.Cells(6, 15) = tMonth2 & "月"
        Ws.Cells(6, 16) = tMonth1 & "月"
        Ws.Cells(6, 17) = tMonth & "月"
        Ws.Cells(6, 18) = tMonth2 & "月"
        Ws.Cells(6, 19) = tMonth1 & "月"
        Ws.Cells(6, 20) = tMonth & "月"
        Ws.Cells(6, 21) = tMonth2 & "月"
        Ws.Cells(6, 22) = tMonth1 & "月"
        Ws.Cells(6, 23) = tMonth & "月"
        Ws.Cells(6, 24) = tMonth2 & "月-" & tMonth & "月"
        Ws.Cells(6, 25) = tMonth2 & "月"
        Ws.Cells(6, 26) = tMonth1 & "月"
        Ws.Cells(6, 27) = tMonth & "月"
        Ws.Cells(6, 28) = tMonth2 & "月"
        Ws.Cells(6, 29) = tMonth1 & "月"
        Ws.Cells(6, 30) = tMonth & "月"
        Ws.Cells(6, 31) = tMonth2 & "月-" & tMonth & "月"
        Ws.Cells(6, 32) = tMonth2 & "月"
        Ws.Cells(6, 33) = tMonth1 & "月"
        Ws.Cells(6, 34) = tMonth & "月"
        Ws.Cells(6, 35) = tMonth2 & "月"
        Ws.Cells(6, 36) = tMonth1 & "月"
        Ws.Cells(6, 37) = tMonth & "月"
        Ws.Cells(6, 38) = tMonth2 & "月-" & tMonth & "月"
        Ws.Cells(6, 39) = tMonth2 & "月"
        Ws.Cells(6, 40) = tMonth1 & "月"
        Ws.Cells(6, 41) = tMonth & "月"
        Ws.Cells(6, 42) = tMonth2 & "月"
        Ws.Cells(6, 43) = tMonth1 & "月"
        Ws.Cells(6, 44) = tMonth & "月"
        Ws.Cells(6, 45) = tMonth2 & "月-" & tMonth & "月"
        Ws.Cells(6, 46) = tMonth2 & "月"
        Ws.Cells(6, 47) = tMonth1 & "月"
        Ws.Cells(6, 48) = tMonth & "月"
        Ws.Cells(6, 49) = tMonth2 & "月"
        Ws.Cells(6, 50) = tMonth1 & "月"
        Ws.Cells(6, 51) = tMonth & "月"
        Ws.Cells(6, 52) = tMonth2 & "月-" & tMonth & "月"
        Ws.Cells(6, 53) = tMonth2 & "月"
        Ws.Cells(6, 54) = tMonth1 & "月"
        Ws.Cells(6, 55) = tMonth & "月"
        Ws.Cells(6, 56) = tMonth2 & "月"
        Ws.Cells(6, 57) = tMonth1 & "月"
        Ws.Cells(6, 58) = tMonth & "月"
        Ws.Cells(6, 59) = tMonth2 & "月"
        Ws.Cells(6, 60) = tMonth1 & "月"
        Ws.Cells(6, 61) = tMonth & "月"
    End Sub
End Class