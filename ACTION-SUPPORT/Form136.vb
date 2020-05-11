Imports Microsoft.Office.Interop.Excel.XlFileFormat
Public Class Form136
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim Start1 As String = String.Empty
    Dim End1 As String = String.Empty
    Dim TotalPeriod As Int16 = 0
    Dim tDate1 As Date
    Dim tDate2 As Date
    Dim LineZ As Integer = 0
    Dim SC As String = String.Empty
    Dim SC1 As String = String.Empty
    Dim TT1 As Decimal = 0
    Dim TT2 As Decimal = 0
    Dim TT3 As Decimal = 0
    Dim TT4 As Decimal = 0
    Dim TT5 As Decimal = 0
    Dim TT6 As Decimal = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form136_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
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
        If Now.Month < 10 Then
            TextBox3.Text = Now.Year & "0" & Now.Month
            TextBox2.Text = Now.Year & "0" & Now.Month
        Else
            TextBox3.Text = Now.Year & Now.Month
            TextBox2.Text = Now.Year & Now.Month
        End If
        Label6.Text = 0
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Start1 = TextBox2.Text
        End1 = TextBox3.Text
        If String.IsNullOrEmpty(Start1) Or String.IsNullOrEmpty(End1) Then
            MsgBox("期间资料错误")
            Return
        End If
        If Len(Start1) <> 6 Or Len(End1) <> 6 Then
            MsgBox("月份资料为6码")
            Return
        End If
        If Conversion.Int(Start1) > Conversion.Int(End1) Then
            MsgBox("开时期间大于结束期间")
            Return
        End If
        TotalPeriod = (Conversion.Int(Strings.Left(End1, 4)) - Conversion.Int(Strings.Left(Start1, 4))) * 12
        TotalPeriod += Conversion.Int(Strings.Right(End1, 2))
        TotalPeriod -= Conversion.Int(Strings.Right(Start1, 2))
        TotalPeriod += 1
        If TotalPeriod > 12 Then
            MsgBox("超出12个月")
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
        SC = TextBox1.Text
        SC1 = TextBox4.Text
        Label6.Text = 0
        ' 20180827
        tDate1 = Convert.ToDateTime(Strings.Left(Start1, 4) & "/" & Strings.Right(Start1, 2) & "/01")
        tDate2 = Convert.ToDateTime(Strings.Left(End1, 4) & "/" & Strings.Right(End1, 2) & "/01").AddMonths(1).AddDays(-1)
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
        SaveFileDialog1.FileName = "采购入库金额单价采购量统计表"
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
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        AdjustExcelFormat()
        Ws.Name = "1 原表"
        Label6.Text = "PAGE 1"
        Label6.Refresh()
        oCommand.CommandText = "select distinct rvu04,rvu05 FROM rvu_file join rvv_file on rvv01=rvu01 LEFT OUTER JOIN pmc_file on pmc01 = rvu04 LEFT OUTER JOIN "
        oCommand.CommandText += "(SELECT rva01, rvb01, rvb02, rva05, rva06, rvb22, rvb07, rva07,rva09 FROM rva_file, rvb_file WHERE rva01 = rvb01 AND rvaconf = 'Y' ) ON rvb01 = rvv04 AND rvb02 = rvv05 "
        oCommand.CommandText += "LEFT OUTER join (SELECT pmm01, pmm20,pma02,pmm21,gec02,pmm22,pmm42, pmm12 FROM pmm_file left OUTER join pma_file ON pma01 = pmm20 left OUTER join gec_file ON gec01 = pmm21 AND gec011 = '1' ) ON pmm01 = rvv36 "
        oCommand.CommandText += "LEFT OUTER JOIN ima_file ON ima01 = rvv31 LEFT JOIN oga_file on oga99= rvu99 WHERE rvu01 = rvv01  and rvu00='1' AND rvuconf = 'Y' AND rva06 between to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        If Not String.IsNullOrEmpty(SC) Then
            oCommand.CommandText += " AND rvv31 LIKE '" & SC & "%' "
        End If
        If Not String.IsNullOrEmpty(SC1) Then
            oCommand.CommandText += " AND rvu04 LIKE '" & SC1 & "%' "
        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                TT1 = 0
                TT2 = 0
                TT3 = 0
                TT4 = 0
                TT5 = 0
                TT6 = 0
                oCommand2.CommandText = "select rva05,pmc03, rva06, rvv04, rva09,rvv01,rvv36, rvv31,rvv031,ima021, rvv35,rvb07,case when rvu00 in ('1') then rvv17 else 0 end as rvv17,"
                oCommand2.CommandText += "case when rvu00 in ('2') then rvv17 else 0 end as rvv171,rvv38t,pmm22,pmm42,(pmm42 * rvv39) as t1, (pmm42 * (rvv39t - rvv39)) as t2,"
                oCommand2.CommandText += "(pmm42 * rvv39t) as t3,rvv39, (rvv39t -rvv39) as t4,rvv39t FROM rvu_file join rvv_file on rvv01=rvu01 LEFT OUTER JOIN pmc_file on pmc01 = rvu04 LEFT OUTER JOIN "
                oCommand2.CommandText += "(SELECT rva01, rvb01, rvb02, rva05, rva06, rvb22, rvb07, rva07,rva09 FROM rva_file, rvb_file WHERE rva01 = rvb01 AND rvaconf = 'Y' ) ON rvb01 = rvv04 AND rvb02 = rvv05 "
                oCommand2.CommandText += "LEFT OUTER join (SELECT pmm01, pmm20,pma02,pmm21,gec02,pmm22,pmm42, pmm12 FROM pmm_file left OUTER join pma_file ON pma01 = pmm20 left OUTER join gec_file ON gec01 = pmm21 AND gec011 = '1' ) ON pmm01 = rvv36 "
                oCommand2.CommandText += "LEFT OUTER JOIN ima_file ON ima01 = rvv31 LEFT JOIN oga_file on oga99= rvu99 WHERE rvu01 = rvv01  and rvu00='1' AND rvuconf = 'Y' AND rva06 between to_date('"
                oCommand2.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
                If Not String.IsNullOrEmpty(SC) Then
                    oCommand2.CommandText += " AND rvv31 LIKE '" & SC & "%' "
                End If
                oCommand2.CommandText += " AND rvu04 = '" & oReader.Item("rvu04") & "' order by rva06"
                oReader2 = oCommand2.ExecuteReader()
                If oReader2.HasRows() Then
                    While oReader2.Read()
                        Ws.Cells(LineZ, 1) = oReader2.Item("rva05") & " " & oReader2.Item("pmc03")
                        Ws.Cells(LineZ, 2) = oReader2.Item("rva06")
                        Ws.Cells(LineZ, 3) = oReader2.Item("rvv04")
                        Ws.Cells(LineZ, 4) = oReader2.Item("rva09")
                        Ws.Cells(LineZ, 5) = oReader2.Item("rvv01")
                        Ws.Cells(LineZ, 6) = oReader2.Item("rvv36")
                        Ws.Cells(LineZ, 7) = oReader2.Item("rvv31")
                        Ws.Cells(LineZ, 8) = oReader2.Item("rvv031")
                        Ws.Cells(LineZ, 9) = oReader2.Item("ima021")
                        Ws.Cells(LineZ, 10) = oReader2.Item("rvv35")
                        Ws.Cells(LineZ, 11) = oReader2.Item("rvb07")
                        Ws.Cells(LineZ, 12) = oReader2.Item("rvv17")
                        Ws.Cells(LineZ, 13) = oReader2.Item("rvv171")
                        Ws.Cells(LineZ, 14) = oReader2.Item("rvv38t")
                        Ws.Cells(LineZ, 15) = oReader2.Item("pmm22")
                        Ws.Cells(LineZ, 16) = oReader2.Item("pmm42")
                        Ws.Cells(LineZ, 17) = oReader2.Item("t1")
                        Ws.Cells(LineZ, 18) = oReader2.Item("t2")
                        Ws.Cells(LineZ, 19) = oReader2.Item("t3")
                        Ws.Cells(LineZ, 20) = oReader2.Item("rvv39")
                        Ws.Cells(LineZ, 21) = oReader2.Item("t4")
                        Ws.Cells(LineZ, 22) = oReader2.Item("rvv39t")
                        TT1 += oReader2.Item("t1")
                        TT2 += oReader2.Item("t2")
                        TT3 += oReader2.Item("t3")
                        TT4 += oReader2.Item("rvv39")
                        TT5 += oReader2.Item("t4")
                        TT6 += oReader2.Item("rvv39t")

                        LineZ += 1
                    End While
                    Ws.Cells(LineZ, 1) = "合計:"
                    Ws.Cells(LineZ, 17) = TT1
                    Ws.Cells(LineZ, 18) = TT2
                    Ws.Cells(LineZ, 19) = TT3
                    Ws.Cells(LineZ, 20) = TT4
                    Ws.Cells(LineZ, 21) = TT5
                    Ws.Cells(LineZ, 22) = TT6
                    LineZ += 1
                End If
                oReader2.Close()
            End While
        End If
        oReader.Close()

        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustExcelFormat1()
        Ws.Name = "2 明细表"
        Label6.Text = "PAGE 2"
        Label6.Refresh()
        oCommand.CommandText = "select rva05,pmc03, rva06, rvv04, rva09,rvv01,rvv36, rvv31,rvv031,ima021, rvv35,rvb07,case when rvu00 in ('1') then rvv17 else 0 end as rvv17,"
        oCommand.CommandText += "case when rvu00 in ('2') then rvv17 else 0 end as rvv171,rvv38t,pmm22,pmm42,(pmm42 * rvv39) as t1, (pmm42 * (rvv39t - rvv39)) as t2,"
        oCommand.CommandText += "(pmm42 * rvv39t) as t3,rvv39, (rvv39t -rvv39) as t4,rvv39t,azf03 FROM rvu_file join rvv_file on rvv01=rvu01 LEFT OUTER JOIN pmc_file on pmc01 = rvu04 LEFT OUTER JOIN "
        oCommand.CommandText += "(SELECT rva01, rvb01, rvb02, rva05, rva06, rvb22, rvb07, rva07,rva09 FROM rva_file, rvb_file WHERE rva01 = rvb01 AND rvaconf = 'Y' ) ON rvb01 = rvv04 AND rvb02 = rvv05 "
        oCommand.CommandText += "LEFT OUTER join (SELECT pmm01, pmm20,pma02,pmm21,gec02,pmm22,pmm42, pmm12 FROM pmm_file left OUTER join pma_file ON pma01 = pmm20 left OUTER join gec_file ON gec01 = pmm21 AND gec011 = '1' ) ON pmm01 = rvv36 "
        oCommand.CommandText += "LEFT OUTER JOIN ima_file ON ima01 = rvv31 LEFT JOIN oga_file on oga99= rvu99 left join azf_file on ima11 = azf_file.azf01 and azf_file.azf02 = 'F' WHERE rvu01 = rvv01 and rvv31 <> 'MISC' and rvu00='1' AND rvuconf = 'Y' AND rva06 between to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        If Not String.IsNullOrEmpty(SC) Then
            oCommand.CommandText += " AND rvv31 LIKE '" & SC & "%' "
        End If
        If Not String.IsNullOrEmpty(SC1) Then
            oCommand.CommandText += " AND rvu04 LIKE '" & SC1 & "%' "
        End If
        oCommand.CommandText += " order by rva05 "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("rva05") & " " & oReader.Item("pmc03")
                Ws.Cells(LineZ, 2) = oReader.Item("rva06")
                Ws.Cells(LineZ, 3) = oReader.Item("rvv04")
                Ws.Cells(LineZ, 4) = oReader.Item("rva09")
                Ws.Cells(LineZ, 5) = oReader.Item("rvv01")
                Ws.Cells(LineZ, 6) = oReader.Item("rvv36")
                Ws.Cells(LineZ, 7) = oReader.Item("rvv31")
                Ws.Cells(LineZ, 8) = oReader.Item("rvv031")
                Ws.Cells(LineZ, 9) = oReader.Item("ima021")
                Ws.Cells(LineZ, 10) = oReader.Item("rvv35")
                Ws.Cells(LineZ, 11) = oReader.Item("rvb07")
                Ws.Cells(LineZ, 12) = oReader.Item("rvv17")
                Ws.Cells(LineZ, 13) = oReader.Item("rvv171")
                Ws.Cells(LineZ, 14) = oReader.Item("rvv38t")
                Ws.Cells(LineZ, 15) = oReader.Item("pmm22")
                Ws.Cells(LineZ, 16) = oReader.Item("pmm42")
                Ws.Cells(LineZ, 17) = oReader.Item("t1")
                Ws.Cells(LineZ, 18) = oReader.Item("t2")
                Ws.Cells(LineZ, 19) = oReader.Item("t3")
                Ws.Cells(LineZ, 20) = oReader.Item("rvv39")
                Ws.Cells(LineZ, 21) = oReader.Item("t4")
                Ws.Cells(LineZ, 22) = oReader.Item("rvv39t")
                Ws.Cells(LineZ, 23) = oReader.Item("azf03")

                LineZ += 1
            End While
        End If
        oReader.Close()


        ' 第三頁
        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        AdjustExcelFormat2()
        Ws.Name = "3 按料号汇总"
        Label6.Text = "PAGE 3"
        Label6.Refresh()
        oCommand.CommandText = "select rvv31,ima02,ima021,rvv35"
        For i As Int16 = 1 To 24 Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        oCommand.CommandText += ",sum(t1+ t2+t3+t4+t5 +t6+t7+t8+t9+t10+t11+t12) as ta,sum(t13+t14+t15+t16+t17+t18+t19+t20+t21+t22+t23+t24) as tt,imz02 from ( "
        oCommand.CommandText += "select rvv31,ima02,ima021,rvv35"
        For i As Int16 = 1 To 12 Step 1
            Dim TMonth As Int16 = Conversion.Int(Strings.Right(Start1, 2)) + i - 1
            Dim CT As String = String.Empty
            Dim CM As String = String.Empty
            If TMonth > 12 Then       
                CT = Conversion.Int(Strings.Left(Start1, 4)) + 1
                CM = TMonth - 12
            Else
                CT = Conversion.Int(Strings.Left(Start1, 4))
                CM = TMonth
            End If
            oCommand.CommandText += ",(case when year(rva06) = " & CT & " and month(rva06) = " & CM & " then rvv17 else 0 end) as t" & i
        Next
        For i As Int16 = 1 To 12 Step 1
            Dim TMonth As Int16 = Conversion.Int(Strings.Right(Start1, 2)) + i - 1
            Dim CT As String = String.Empty
            Dim CM As String = String.Empty
            If TMonth > 12 Then
                CT = Conversion.Int(Strings.Left(Start1, 4)) + 1
                CM = TMonth - 12
            Else
                CT = Conversion.Int(Strings.Left(Start1, 4))
                CM = TMonth
            End If
            oCommand.CommandText += ",(case when year(rva06) = " & CT & " and month(rva06) = " & CM & " then pmm42 * rvv39 else 0 end) as t" & i + 12
        Next
        oCommand.CommandText += ",imz02 FROM rvu_file join rvv_file on rvv01=rvu01 LEFT OUTER JOIN pmc_file on pmc01 = rvu04 LEFT OUTER JOIN "
        oCommand.CommandText += "(SELECT rva01, rvb01, rvb02, rva05, rva06, rvb22, rvb07, rva07,rva09 FROM rva_file ,  rvb_file  "
        oCommand.CommandText += "WHERE rva01 = rvb01 AND rvaconf = 'Y' )  ON  rvb01 = rvv04 AND rvb02 = rvv05 LEFT OUTER join (SELECT pmm01, pmm20,pma02,pmm21,gec02,pmm22,pmm42, pmm12 "
        oCommand.CommandText += "FROM pmm_file left OUTER join pma_file ON pma01 = pmm20 left OUTER join gec_file ON gec01 = pmm21 AND gec011 = '1' ) ON pmm01 = rvv36 "
        oCommand.CommandText += "LEFT OUTER JOIN ima_file ON ima01 = rvv31 LEFT JOIN oga_file on oga99= rvu99 left join imz_file on ima06 = imz01 WHERE rvu01 = rvv01 and rvv31 <> 'MISC'  and rvu00='1' "
        oCommand.CommandText += "AND rvuconf = 'Y' AND rva06 between to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        If Not String.IsNullOrEmpty(SC) Then
            oCommand.CommandText += " AND rvv31 LIKE '" & SC & "%' "
        End If
        If Not String.IsNullOrEmpty(SC1) Then
            oCommand.CommandText += " AND rvu04 LIKE '" & SC1 & "%' "
        End If
        oCommand.CommandText += ") group by rvv31,ima02,ima021,rvv35,imz02 order by tt desc"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 1 To oReader.FieldCount Step 1
                    Ws.Cells(LineZ, i) = oReader.Item(i - 1)
                Next
                LineZ += 1
            End While
        End If
        oReader.Close()

        ' 第四頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws.Activate()
        AdjustExcelFormat3()
        Ws.Name = "4 按厂商料号汇总"
        Label6.Text = "PAGE 4"
        Label6.Refresh()
        oCommand.CommandText = "select rvu04,rvu05,rvv31,ima02,ima021,rvv35,round(sum(t13+t14+t15+t16+t17+t18+t19+t20+t21+t22+t23+t24)/sum(t1+ t2+t3+t4+t5 +t6+t7+t8+t9+t10+t11+t12),2) as tn"
        For i As Int16 = 1 To 24 Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        oCommand.CommandText += ",sum(t1+ t2+t3+t4+t5 +t6+t7+t8+t9+t10+t11+t12) as ta,sum(t13+t14+t15+t16+t17+t18+t19+t20+t21+t22+t23+t24) as tt,imz02 from ( "
        oCommand.CommandText += "select rvu04,rvu05,rvv31,ima02,ima021,rvv35"
        For i As Int16 = 1 To 12 Step 1
            Dim TMonth As Int16 = Conversion.Int(Strings.Right(Start1, 2)) + i - 1
            Dim CT As String = String.Empty
            Dim CM As String = String.Empty
            If TMonth > 12 Then
                CT = Conversion.Int(Strings.Left(Start1, 4)) + 1
                CM = TMonth - 12
            Else
                CT = Conversion.Int(Strings.Left(Start1, 4))
                CM = TMonth
            End If
            oCommand.CommandText += ",(case when year(rva06) = " & CT & " and month(rva06) = " & CM & " then rvv17 else 0 end) as t" & i
        Next
        For i As Int16 = 1 To 12 Step 1
            Dim TMonth As Int16 = Conversion.Int(Strings.Right(Start1, 2)) + i - 1
            Dim CT As String = String.Empty
            Dim CM As String = String.Empty
            If TMonth > 12 Then
                CT = Conversion.Int(Strings.Left(Start1, 4)) + 1
                CM = TMonth - 12
            Else
                CT = Conversion.Int(Strings.Left(Start1, 4))
                CM = TMonth
            End If
            oCommand.CommandText += ",(case when year(rva06) = " & CT & " and month(rva06) = " & CM & " then pmm42 * rvv39 else 0 end) as t" & i + 12
        Next
        oCommand.CommandText += ",imz02 FROM rvu_file join rvv_file on rvv01=rvu01 LEFT OUTER JOIN pmc_file on pmc01 = rvu04 LEFT OUTER JOIN "
        oCommand.CommandText += "(SELECT rva01, rvb01, rvb02, rva05, rva06, rvb22, rvb07, rva07,rva09 FROM rva_file ,  rvb_file  "
        oCommand.CommandText += "WHERE rva01 = rvb01 AND rvaconf = 'Y' )  ON  rvb01 = rvv04 AND rvb02 = rvv05 LEFT OUTER join (SELECT pmm01, pmm20,pma02,pmm21,gec02,pmm22,pmm42, pmm12 "
        oCommand.CommandText += "FROM pmm_file left OUTER join pma_file ON pma01 = pmm20 left OUTER join gec_file ON gec01 = pmm21 AND gec011 = '1' ) ON pmm01 = rvv36 "
        oCommand.CommandText += "LEFT OUTER JOIN ima_file ON ima01 = rvv31 LEFT JOIN oga_file on oga99= rvu99 left join imz_file on ima06 = imz01 WHERE rvu01 = rvv01 and rvv31 <> 'MISC'  and rvu00='1' "
        oCommand.CommandText += "AND rvuconf = 'Y' AND rva06 between to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        If Not String.IsNullOrEmpty(SC) Then
            oCommand.CommandText += " AND rvv31 LIKE '" & SC & "%' "
        End If
        If Not String.IsNullOrEmpty(SC1) Then
            oCommand.CommandText += " AND rvu04 LIKE '" & SC1 & "%' "
        End If
        oCommand.CommandText += ") group by rvu04,rvu05,rvv31,ima02,ima021,rvv35,imz02 order by tt desc"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 1 To oReader.FieldCount Step 1
                    Ws.Cells(LineZ, i) = oReader.Item(i - 1)
                Next
                LineZ += 1
            End While
        End If
        oReader.Close()

        ' 第五頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws.Activate()
        AdjustExcelFormat4()
        Ws.Name = "5 按厂商汇总"
        Label6.Text = "PAGE 5"
        Label6.Refresh()
        oCommand.CommandText = "select rvu04,rvu05,pmy02"
        For i As Int16 = 1 To 12 Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        oCommand.CommandText += ",sum(t1+ t2+t3+t4+t5 +t6+t7+t8+t9+t10+t11+t12) as ta from ( "
        oCommand.CommandText += "select rvu04,rvu05,pmy02"
        
        For i As Int16 = 1 To 12 Step 1
            Dim TMonth As Int16 = Conversion.Int(Strings.Right(Start1, 2)) + i - 1
            Dim CT As String = String.Empty
            Dim CM As String = String.Empty
            If TMonth > 12 Then
                CT = Conversion.Int(Strings.Left(Start1, 4)) + 1
                CM = TMonth - 12
            Else
                CT = Conversion.Int(Strings.Left(Start1, 4))
                CM = TMonth
            End If
            oCommand.CommandText += ",(case when year(rva06) = " & CT & " and month(rva06) = " & CM & " then pmm42 * rvv39 else 0 end) as t" & i
        Next
        oCommand.CommandText += " FROM rvu_file join rvv_file on rvv01=rvu01 LEFT OUTER JOIN pmc_file on pmc01 = rvu04 LEFT OUTER JOIN "
        oCommand.CommandText += "(SELECT rva01, rvb01, rvb02, rva05, rva06, rvb22, rvb07, rva07,rva09 FROM rva_file ,  rvb_file  "
        oCommand.CommandText += "WHERE rva01 = rvb01 AND rvaconf = 'Y' )  ON  rvb01 = rvv04 AND rvb02 = rvv05 LEFT OUTER join (SELECT pmm01, pmm20,pma02,pmm21,gec02,pmm22,pmm42, pmm12 "
        oCommand.CommandText += "FROM pmm_file left OUTER join pma_file ON pma01 = pmm20 left OUTER join gec_file ON gec01 = pmm21 AND gec011 = '1' ) ON pmm01 = rvv36 "
        oCommand.CommandText += "LEFT OUTER JOIN ima_file ON ima01 = rvv31 LEFT JOIN oga_file on oga99= rvu99 left join pmy_file on substr(pmc01,1,1) = pmy01 WHERE rvu01 = rvv01 and rvv31 <> 'MISC'  and rvu00='1' "
        oCommand.CommandText += "AND rvuconf = 'Y' AND rva06 between to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        If Not String.IsNullOrEmpty(SC) Then
            oCommand.CommandText += " AND rvv31 LIKE '" & SC & "%' "
        End If
        If Not String.IsNullOrEmpty(SC1) Then
            oCommand.CommandText += " AND rvu04 LIKE '" & SC1 & "%' "
        End If
        oCommand.CommandText += ") group by rvu04,rvu05,pmy02 order by ta desc"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 1 To oReader.FieldCount Step 1
                    Ws.Cells(LineZ, i) = oReader.Item(i - 1)
                Next
                LineZ += 1
            End While
        End If
        oReader.Close()

        ' 第六頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws.Activate()
        AdjustExcelFormat5()
        Ws.Name = "6 按厂商类别汇总"
        Label6.Text = "PAGE 6"
        Label6.Refresh()
        oCommand.CommandText = "select pmy02"
        For i As Int16 = 1 To 12 Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        oCommand.CommandText += ",sum(t1+ t2+t3+t4+t5 +t6+t7+t8+t9+t10+t11+t12) as ta from ( "
        oCommand.CommandText += "select pmy02"

        For i As Int16 = 1 To 12 Step 1
            Dim TMonth As Int16 = Conversion.Int(Strings.Right(Start1, 2)) + i - 1
            Dim CT As String = String.Empty
            Dim CM As String = String.Empty
            If TMonth > 12 Then
                CT = Conversion.Int(Strings.Left(Start1, 4)) + 1
                CM = TMonth - 12
            Else
                CT = Conversion.Int(Strings.Left(Start1, 4))
                CM = TMonth
            End If
            oCommand.CommandText += ",(case when year(rva06) = " & CT & " and month(rva06) = " & CM & " then pmm42 * rvv39 else 0 end) as t" & i
        Next
        oCommand.CommandText += " FROM rvu_file join rvv_file on rvv01=rvu01 LEFT OUTER JOIN pmc_file on pmc01 = rvu04 LEFT OUTER JOIN "
        oCommand.CommandText += "(SELECT rva01, rvb01, rvb02, rva05, rva06, rvb22, rvb07, rva07,rva09 FROM rva_file ,  rvb_file  "
        oCommand.CommandText += "WHERE rva01 = rvb01 AND rvaconf = 'Y' )  ON  rvb01 = rvv04 AND rvb02 = rvv05 LEFT OUTER join (SELECT pmm01, pmm20,pma02,pmm21,gec02,pmm22,pmm42, pmm12 "
        oCommand.CommandText += "FROM pmm_file left OUTER join pma_file ON pma01 = pmm20 left OUTER join gec_file ON gec01 = pmm21 AND gec011 = '1' ) ON pmm01 = rvv36 "
        oCommand.CommandText += "LEFT OUTER JOIN ima_file ON ima01 = rvv31 LEFT JOIN oga_file on oga99= rvu99 left join pmy_file on substr(pmc01,1,1) = pmy01 WHERE rvu01 = rvv01 and rvv31 <> 'MISC'  and rvu00='1' "
        oCommand.CommandText += "AND rvuconf = 'Y' AND rva06 between to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        If Not String.IsNullOrEmpty(SC) Then
            oCommand.CommandText += " AND rvv31 LIKE '" & SC & "%' "
        End If
        If Not String.IsNullOrEmpty(SC1) Then
            oCommand.CommandText += " AND rvu04 LIKE '" & SC1 & "%' "
        End If
        oCommand.CommandText += ") group by pmy02 order by ta desc"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 1 To oReader.FieldCount Step 1
                    Ws.Cells(LineZ, i) = oReader.Item(i - 1)
                Next
                LineZ += 1
            End While
        End If
        oReader.Close()

    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.ColumnWidth = 21.11
        Ws.Cells(1, 1) = "廠商"
        Ws.Cells(1, 2) = "收貨日期"
        Ws.Cells(1, 3) = "收貨單號"
        Ws.Cells(1, 4) = "廠商送貨單號"
        Ws.Cells(1, 5) = "入庫/退貨單號"
        Ws.Cells(1, 6) = "採購單號"
        Ws.Cells(1, 7) = "料件編號"
        Ws.Cells(1, 8) = "品名"
        Ws.Cells(1, 9) = "規格"
        Ws.Cells(1, 10) = "單位"
        Ws.Cells(1, 11) = "收貨量"
        Ws.Cells(1, 12) = "入庫量"
        Ws.Cells(1, 13) = "驗退量"
        Ws.Cells(1, 14) = "含稅單價"
        Ws.Cells(1, 15) = "币别"
        Ws.Cells(1, 16) = "匯率"
        Ws.Cells(1, 17) = "未税金额(本幣)"
        Ws.Cells(1, 18) = "税额(本幣)"
        Ws.Cells(1, 19) = "含稅金額(本幣)"
        Ws.Cells(1, 20) = "未税金额(原幣)"
        Ws.Cells(1, 21) = "税额(原幣)"
        Ws.Cells(1, 22) = "含稅金額(原幣)"
        oRng = Ws.Range("G1", "G1")
        oRng.EntireColumn.NumberFormat = "@"
        oRng = Ws.Range("K1", "V1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00"
        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.ColumnWidth = 21.11
        Ws.Cells(1, 1) = "廠商"
        Ws.Cells(1, 2) = "收貨日期"
        Ws.Cells(1, 3) = "收貨單號"
        Ws.Cells(1, 4) = "廠商送貨單號"
        Ws.Cells(1, 5) = "入庫/退貨單號"
        Ws.Cells(1, 6) = "採購單號"
        Ws.Cells(1, 7) = "料件編號"
        Ws.Cells(1, 8) = "品名"
        Ws.Cells(1, 9) = "規格"
        Ws.Cells(1, 10) = "單位"
        Ws.Cells(1, 11) = "收貨量"
        Ws.Cells(1, 12) = "入庫量"
        Ws.Cells(1, 13) = "驗退量"
        Ws.Cells(1, 14) = "含稅單價"
        Ws.Cells(1, 15) = "币别"
        Ws.Cells(1, 16) = "匯率"
        Ws.Cells(1, 17) = "未税金额(本幣)"
        Ws.Cells(1, 18) = "税额(本幣)"
        Ws.Cells(1, 19) = "含稅金額(本幣)"
        Ws.Cells(1, 20) = "未税金额(原幣)"
        Ws.Cells(1, 21) = "税额(原幣)"
        Ws.Cells(1, 22) = "含稅金額(原幣)"
        Ws.Cells(1, 23) = "其它分群码三说明"
        oRng = Ws.Range("G1", "G1")
        oRng.EntireColumn.NumberFormat = "@"
        oRng = Ws.Range("K1", "V1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00"
        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.ColumnWidth = 21.11
        Ws.Cells(1, 1) = "料件编号"
        Ws.Cells(1, 2) = "品名"
        Ws.Cells(1, 3) = "規格"
        Ws.Cells(1, 4) = "單位"
        For i As Integer = 1 To 12 Step 1
            Dim TMonth As Int16 = Conversion.Int(Strings.Right(Start1, 2)) + i - 1
            If TMonth > 12 Then
                If TMonth - 12 < 10 Then
                    Ws.Cells(1, 4 + i) = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "/0" & TMonth - 12 & "入库量"
                Else
                    Ws.Cells(1, 4 + i) = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "/" & TMonth - 12 & "入库量"
                End If
            Else
                If TMonth < 10 Then
                    Ws.Cells(1, 4 + i) = Conversion.Int(Strings.Left(Start1, 4)) & "/0" & TMonth & "入库量"
                Else
                    Ws.Cells(1, 4 + i) = Conversion.Int(Strings.Left(Start1, 4)) & "/" & TMonth & "入库量"
                End If
            End If
        Next
        For i As Integer = 1 To 12 Step 1
            Dim TMonth As Int16 = Conversion.Int(Strings.Right(Start1, 2)) + i - 1
            If TMonth > 12 Then
                If TMonth - 12 < 10 Then
                    Ws.Cells(1, 16 + i) = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "/0" & TMonth - 12 & "未税金额(本幣)"
                Else
                    Ws.Cells(1, 16 + i) = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "/" & TMonth - 12 & "未税金额(本幣)"
                End If
            Else
                If TMonth < 10 Then
                    Ws.Cells(1, 16 + i) = Conversion.Int(Strings.Left(Start1, 4)) & "/0" & TMonth & "未税金额(本幣)"
                Else
                    Ws.Cells(1, 16 + i) = Conversion.Int(Strings.Left(Start1, 4)) & "/" & TMonth & "未税金额(本幣)"
                End If
            End If
        Next
        Ws.Cells(1, 29) = "入库量"
        Ws.Cells(1, 30) = "未税金额(本幣)"
        Ws.Cells(1, 31) = "分群码"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormat = "@"
        oRng = Ws.Range("E1", "AE1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00"
        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat3()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.ColumnWidth = 21.11
        Ws.Cells(1, 1) = "厂商编号"
        Ws.Cells(1, 2) = "厂商名称"
        Ws.Cells(1, 3) = "料件编号"
        Ws.Cells(1, 4) = "品名"
        Ws.Cells(1, 5) = "規格"
        Ws.Cells(1, 6) = "單位"
        Ws.Cells(1, 7) = "平均值 未税单价"
        For i As Integer = 1 To 12 Step 1
            Dim TMonth As Int16 = Conversion.Int(Strings.Right(Start1, 2)) + i - 1
            If TMonth > 12 Then
                If TMonth - 12 < 10 Then
                    Ws.Cells(1, 7 + i) = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "/0" & TMonth - 12 & "入库量"
                Else
                    Ws.Cells(1, 7 + i) = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "/" & TMonth - 12 & "入库量"
                End If
            Else
                If TMonth < 10 Then
                    Ws.Cells(1, 7 + i) = Conversion.Int(Strings.Left(Start1, 4)) & "/0" & TMonth & "入库量"
                Else
                    Ws.Cells(1, 7 + i) = Conversion.Int(Strings.Left(Start1, 4)) & "/" & TMonth & "入库量"
                End If
            End If
        Next
        For i As Integer = 1 To 12 Step 1
            Dim TMonth As Int16 = Conversion.Int(Strings.Right(Start1, 2)) + i - 1
            If TMonth > 12 Then
                If TMonth - 12 < 10 Then
                    Ws.Cells(1, 19 + i) = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "/0" & TMonth - 12 & "未税金额(本幣)"
                Else
                    Ws.Cells(1, 19 + i) = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "/" & TMonth - 12 & "未税金额(本幣)"
                End If
            Else
                If TMonth < 10 Then
                    Ws.Cells(1, 19 + i) = Conversion.Int(Strings.Left(Start1, 4)) & "/0" & TMonth & "未税金额(本幣)"
                Else
                    Ws.Cells(1, 19 + i) = Conversion.Int(Strings.Left(Start1, 4)) & "/" & TMonth & "未税金额(本幣)"
                End If
            End If
        Next
        Ws.Cells(1, 32) = "入库量"
        Ws.Cells(1, 33) = "未税金额(本幣)"
        Ws.Cells(1, 34) = "分群码"
        oRng = Ws.Range("A1", "C1")
        oRng.EntireColumn.NumberFormat = "@"
        oRng = Ws.Range("G1", "AG1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00"
        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat4()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.ColumnWidth = 21.11
        Ws.Cells(1, 1) = "厂商编号"
        Ws.Cells(1, 2) = "厂商名称"
        Ws.Cells(1, 3) = "厂商分类"

        For i As Integer = 1 To 12 Step 1
            Dim TMonth As Int16 = Conversion.Int(Strings.Right(Start1, 2)) + i - 1
            If TMonth > 12 Then
                If TMonth - 12 < 10 Then
                    Ws.Cells(1, 3 + i) = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "/0" & TMonth - 12 & "未税金额(本幣)"
                Else
                    Ws.Cells(1, 3 + i) = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "/" & TMonth - 12 & "未税金额(本幣)"
                End If
            Else
                If TMonth < 10 Then
                    Ws.Cells(1, 3 + i) = Conversion.Int(Strings.Left(Start1, 4)) & "/0" & TMonth & "未税金额(本幣)"
                Else
                    Ws.Cells(1, 3 + i) = Conversion.Int(Strings.Left(Start1, 4)) & "/" & TMonth & "未税金额(本幣)"
                End If
            End If
        Next
        Ws.Cells(1, 16) = "未税金额(本幣)"
        oRng = Ws.Range("A1", "C1")
        oRng.EntireColumn.NumberFormat = "@"
        oRng = Ws.Range("D1", "P1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00"
        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat5()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.ColumnWidth = 21.11
        Ws.Cells(1, 1) = "厂商类别"

        For i As Integer = 1 To 12 Step 1
            Dim TMonth As Int16 = Conversion.Int(Strings.Right(Start1, 2)) + i - 1
            If TMonth > 12 Then
                If TMonth - 12 < 10 Then
                    Ws.Cells(1, 1 + i) = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "/0" & TMonth - 12 & "未税金额(本幣)"
                Else
                    Ws.Cells(1, 1 + i) = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "/" & TMonth - 12 & "未税金额(本幣)"
                End If
            Else
                If TMonth < 10 Then
                    Ws.Cells(1, 1 + i) = Conversion.Int(Strings.Left(Start1, 4)) & "/0" & TMonth & "未税金额(本幣)"
                Else
                    Ws.Cells(1, 1 + i) = Conversion.Int(Strings.Left(Start1, 4)) & "/" & TMonth & "未税金额(本幣)"
                End If
            End If
        Next
        Ws.Cells(1, 14) = "未税金额(本幣)"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormat = "@"
        oRng = Ws.Range("B1", "N1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00"
        LineZ = 2
    End Sub
End Class