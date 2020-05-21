Public Class Form168
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
    Dim tTime As Date
    Dim fTime As Date
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim tWeek As Int16 = 0
    Dim fYear As Int16 = 0
    Dim fWeek As Int16 = 0
    Dim AzjYM As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form168_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
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
        tTime = Now
        'tTime = "2019/06/12"
        tYear = tTime.Year
        tMonth = tTime.Month
        oCommand.CommandText = "select nvl(azn05,1) from azn_file where azn01 = to_date('" & tTime.ToString("yyyy/MM/dd") & "','yyyy/MM/dd')"
        tWeek = oCommand.ExecuteScalar()
        ' 20190605 改為 22 週
        fYear = tYear
        fWeek = tWeek + 21
        If fWeek > 53 Then
            fYear = tYear + 1
            fWeek = fWeek - 53
        End If
        ' 確定最後的時間
        oCommand.CommandText = "select max(azn01) from azn_file where azn02  = " & fYear & " and azn05 = " & fWeek
        fTime = oCommand.ExecuteScalar()

        ' 確定匯率取數
        If tMonth < 10 Then
            AzjYM = tYear & "0" & tMonth
        Else
            AzjYM = tYear & tMonth
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
        SaveFileDialog1.FileName = "DAC Cashflow "
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
        Dim xPath As String = "C:\temp\DAC CashFlow_Forecast_Sample.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        AdjustExcelFormat()
        LineZ = 4

        ' 第四列
        oCommand.CommandText = "select "
        For i As Int16 = 1 To 22 Step 1
            oCommand.CommandText += "sum(t" & i & ") as t" & i & ","
        Next
        oCommand.CommandText += "1 from ( select nvl(sum(omc13),0) as t1,"
        For i As Int16 = 2 To 22 Step 1
            oCommand.CommandText += "0 as t" & i & ","
        Next
        oCommand.CommandText += "1 from oma_file left join omc_file on oma01 = omc01 where omaconf = 'Y' and oma03 not in ('D0001','D0002','D0003','D0005')  and oma00 like '1%' and omc13 > 0 and oma11 < to_date('"
        oCommand.CommandText += tTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select "
        For i As Int16 = 1 To 22 Step 1
            If tWeek < fWeek Then
                oCommand.CommandText += "(case when azn05 = " & tWeek + i - 1 & " then omc13 else 0 end) as t" & i & ","
            Else
                If tWeek + i - 1 > 53 Then
                    oCommand.CommandText += "(case when azn05 = " & tWeek + i - 1 - 53 & " then omc13 else 0 end) as t" & i & ","
                Else
                    oCommand.CommandText += "(case when azn05 = " & tWeek + i - 1 & " then omc13 else 0 end) as t" & i & ","
                End If
            End If
        Next
        oCommand.CommandText += "1 from oma_file left join omc_file on oma01 = omc01 left join azn_file on oma11 = azn01 where omaconf = 'Y' and oma03 not in ('D0001','D0002','D0003','D0005')  and oma00 like '1%' and omc13 > 0 and oma11 between to_date('"
        oCommand.CommandText += tTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & fTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') )"

        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 2 Step 1
                    Ws.Cells(4, 3 + i) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()

        ' 第五列
        oCommand.CommandText = "select "
        For i As Int16 = 1 To 22 Step 1
            oCommand.CommandText += "sum(t" & i & ") as t" & i & ","
        Next
        oCommand.CommandText += "1 from ( select nvl(sum(omc13),0) as t1,"
        For i As Int16 = 2 To 22 Step 1
            oCommand.CommandText += "0 as t" & i & ","
        Next
        oCommand.CommandText += "1 from oma_file left join omc_file on oma01 = omc01 where omaconf = 'Y' and oma03 in ('D0001','D0002','D0003','D0005')  and oma00 like '1%' and omc13 > 0 and oma11 < to_date('"
        oCommand.CommandText += tTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select "
        For i As Int16 = 1 To 22 Step 1
            If tWeek < fWeek Then
                oCommand.CommandText += "(case when azn05 = " & tWeek + i - 1 & " then omc13 else 0 end) as t" & i & ","
            Else
                If tWeek + i - 1 > 53 Then
                    oCommand.CommandText += "(case when azn05 = " & tWeek + i - 1 - 53 & " then omc13 else 0 end) as t" & i & ","
                Else
                    oCommand.CommandText += "(case when azn05 = " & tWeek + i - 1 & " then omc13 else 0 end) as t" & i & ","
                End If
            End If
        Next
        oCommand.CommandText += "1 from oma_file left join omc_file on oma01 = omc01 left join azn_file on oma11 = azn01 where omaconf = 'Y' and oma03 in ('D0001','D0002','D0003','D0005')  and oma00 like '1%' and omc13 > 0 and oma11 between to_date('"
        oCommand.CommandText += tTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & fTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select "
        For i As Int16 = 1 To 22 Step 1
            If tWeek > fWeek Then
                oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 & " then (case when t1 is null then t2 * tc_prm04 else t1 * tc_prm04 end) else 0 end),0) as t" & i & ","
            Else
                If tWeek + i - 1 > 53 Then
                    oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 - 53 & " then (case when t1 is null then t2 * tc_prm04 else t1 * tc_prm04 end) else 0 end),0) as t" & i & ","
                Else
                    oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 & " then (case when t1 is null then t2 * tc_prm04 else t1 * tc_prm04 end) else 0 end),0) as t" & i & ","
                End If
            End If
        Next
        oCommand.CommandText += "1 from ( select tc_prm01,tc_prm04, (D1 + 60) as D1, (tc_prl03 * tc_prl04 / 100 * az1.azj03) as t1,(avg(tc_bud12) * az2.azj03) as t2 from ( "
        oCommand.CommandText += "select tc_prm01,tc_prm04, D1, min(tc_prl02) D2 from ( select tc_prm01,tc_prm04,max(azn01) as d1 from tc_prm_file left join azn_file on azn02 = tc_prm02 and azn05 = tc_prm03 "
        oCommand.CommandText += "where tc_prm02 = " & tYear & " and tc_prm04 > 0 group by tc_prm01,tc_prm02,tc_prm03,tc_prm04 ) AB left join tc_prl_file on AB.tc_prm01 = tc_prl_file.tc_prl01 and tc_prl02 > AB.D1 "
        oCommand.CommandText += "group by AB.tc_prm01,AB.D1,Ab.tc_prm04 ) AC left join tc_prl_file on Ac.tc_prm01 = tc_prl_file.tc_prl01 and AC.D2 = tc_prl_file.tc_prl02 left join azj_file az1 on tc_prl06 = az1.azj01 and az1.azj02 = '" & AzjYM & "' "
        oCommand.CommandText += "left join tc_bud_file on tc_bud01 = 1 and tc_bud02 = 2019 and tc_bud04 = tc_prm01 left join azj_file az2 on tc_bud14 = az2.azj01 and az2.azj02 = '" & AzjYM & "' group by tc_prm01,tc_prm04, D1 + 60,(tc_prl03 * tc_prl04 / 100 * az1.azj03), az2.azj03 "
        oCommand.CommandText += ") AD  left join azn_file on AD.d1 = azn_file.azn01 where D1 between to_date('"
        oCommand.CommandText += tTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & fTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') )"

        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 2 Step 1
                    Ws.Cells(5, 3 + i) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()

        ' 第八列
        oCommand.CommandText = "select "
        For i As Int16 = 1 To 22 Step 1
            oCommand.CommandText += "sum(t" & i & ") as t" & i & ","
        Next
        oCommand.CommandText += "1 from ( select "
        For i As Int16 = 1 To 22 Step 1
            If tWeek > fWeek Then
                oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 & " then tc_ext05 else 0 end),0) as t" & i & ","
            Else
                If tWeek + i - 1 > 53 Then
                    oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 - 53 & " then tc_ext05 else 0 end),0) as t" & i & ","
                Else
                    oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 & " then tc_ext05 else 0 end),0) as t" & i & ","
                End If
            End If
        Next
        oCommand.CommandText += "1 from tc_ext_file left join azn_file on tc_ext01 = azn01 where tc_ext02 = 1 and tc_ext01 between to_date('"
        oCommand.CommandText += tTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & fTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') )"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 2 Step 1
                    Ws.Cells(8, 3 + i) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()


        ' 第九列 未完, 只計了 1項
        oCommand.CommandText = "select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12,sum(t13) as t13,sum(t14) as t14,sum(t15) as t15,sum(t16) as t16,sum(t17) as t17,sum(t18) as t18,sum(t19) as t19,sum(t20) as t20,sum(t21) as t21,sum(t22) as t22,1 from ( "
        oCommand.CommandText += "select nvl(sum(apc13),0) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16,0 as t17,0 as t18,0 as t19,0 as t20,0 as t21,0 as t22,1 from apa_file left join apc_file on apa01 = apc01 where apa41 = 'Y' and apa36 in ('002','003','015') and apc13 > 0 and apa12 < to_date('"
        oCommand.CommandText += tTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select nvl(sum(apc13) * -1,0) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16,0 as t17,0 as t18,0 as t19,0 as t20,0 as t21,0 as t22,1 from apa_file left join apc_file on apa01 = apc01 where apa41 = 'Y' and apaud04 = '11' and apc13 > 0 and apa12 < to_date('"
        oCommand.CommandText += tTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select "
        For i As Int16 = 1 To 22 Step 1
            If tWeek > fWeek Then
                oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 & " then apc13 else 0 end),0) as t" & i & ","
            Else
                If tWeek + i - 1 > 53 Then
                    oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 - 53 & " then apc13 else 0 end),0) as t" & i & ","
                Else
                    oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 & " then apc13 else 0 end),0) as t" & i & ","
                End If
            End If
        Next
        oCommand.CommandText += "1 from apa_file left join apc_file on apa01 = apc01 left join azn_file on apa12 = azn01 where apa41 = 'Y' and apa36 in ('002','003','015') and apc13 > 0 and apa12 between to_date('"
        oCommand.CommandText += tTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & fTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select "
        For i As Int16 = 1 To 22 Step 1
            If tWeek > fWeek Then
                oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 & " then apc13 * -1 else 0 end),0) as t" & i & ","
            Else
                If tWeek + i - 1 > 53 Then
                    oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 - 53 & " then apc13 * -1 else 0 end),0) as t" & i & ","
                Else
                    oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 & " then apc13 * -1 else 0 end),0) as t" & i & ","
                End If
            End If
        Next
        oCommand.CommandText += "1 from apa_file left join apc_file on apa01 = apc01 left join azn_file on apa12 = azn01 where apa41 = 'Y' and apaud04 = '11' and apc13 > 0 and apa12 between to_date('"
        oCommand.CommandText += tTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & fTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')  "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select "
        For i As Int16 = 1 To 22 Step 1
            If tWeek > fWeek Then
                oCommand.CommandText += "(case when azn05 = " & tWeek + i - 1 & " then qty * pmh12 * az1 else 0 end )as t" & i & ","
            Else
                If tWeek + i - 1 > 53 Then
                    oCommand.CommandText += "(case when azn05 = " & tWeek + i - 1 - 53 & " then qty * pmh12 * az1 else 0 end )as t" & i & ","
                Else
                    oCommand.CommandText += "(case when azn05 = " & tWeek + i - 1 & " then qty * pmh12 * az1 else 0 end )as t" & i & ","
                End If
            End If
        Next
        '        oCommand.CommandText += "1 from ( select pn,year1,week1,qty,pmh02,pmh12,pma08,max(azn01) as c1,max(azn01) + numtodsinterval(pma08,'day') as d1,pmh13 ,(case when pmh13 = 'RMB' then 1 else azj041 end) as az1 from dac_receive_plan left join ima_file on pn = ima01 "
        oCommand.CommandText += "1 from ( select pn,year1,week1,qty,pmh02,pmh12,pma08,max(azn01) as c1,to_date(Year(max(azn01)) || (case when month(max(azn01)) < 10 then '0' || month(max(azn01)) else to_char(month(max(azn01))) end) || 25,'yyyymmdd') + numtodsinterval(pma08,'day') as d1,pmh13 ,(case when pmh13 = 'RMB' then 1 else azj041 end) as az1 from dac_receive_plan left join ima_file on pn = ima01 "
        oCommand.CommandText += "left join pmh_file on ima532 = pmhdate and ima01 = pmh01 left join pmc_file on pmh02 = pmc01 left join pma_file on pmc17 = pma01 left join azn_file on year1 = azn02 and week1 = azn05 left join azj_file on pmh13 = azj01 and azj02 = '" & AzjYM & "' "
        oCommand.CommandText += "group by pn,year1,week1,qty,pmh02,pmh12,pma08,pmh13,azj041 ) AZ left join azn_file on AZ.d1 = azn01 )"

        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 2 Step 1
                    Ws.Cells(9, 3 + i) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()

        ' 第十列
        oCommand.CommandText = "select "
        For i As Int16 = 1 To 22 Step 1
            oCommand.CommandText += "sum(t" & i & ") as t" & i & ","
        Next
        oCommand.CommandText += "1 from ( select "
        For i As Int16 = 1 To 22 Step 1
            If tWeek > fWeek Then
                oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 & " then tc_ext05 else 0 end),0) as t" & i & ","
            Else
                If tWeek + i - 1 > 53 Then
                    oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 - 53 & " then tc_ext05 else 0 end),0) as t" & i & ","
                Else
                    oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 & " then tc_ext05 else 0 end),0) as t" & i & ","
                End If
            End If
        Next
        oCommand.CommandText += "1 from tc_ext_file left join azn_file on tc_ext01 = azn01 where tc_ext02 = 3 and tc_ext01 between to_date('"
        oCommand.CommandText += tTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & fTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') )"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 2 Step 1
                    Ws.Cells(10, 3 + i) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()

        '第十一列
        oCommand.CommandText = "select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12,sum(t13) as t13,sum(t14) as t14,sum(t15) as t15,sum(t16) as t16,sum(t17) as t17,sum(t18) as t18,sum(t19) as t19,sum(t20) as t20,sum(t21) as t21,sum(t22) as t22,1 from ( "
        oCommand.CommandText += "select nvl(sum(apc13),0) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16,0 as t17,0 as t18,0 as t19,0 as t20,0 as t21,0 as t22,1 from apa_file left join apc_file on apa01 = apc01 where apa41 = 'Y' and apa36 in ('002','003','015') and apc13 > 0 and apaud04 = '11' and apa12 < to_date('"
        oCommand.CommandText += tTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select "
        For i As Int16 = 1 To 22 Step 1
            If tWeek > fWeek Then
                oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 & " then apc13 else 0 end),0) as t" & i & ","
            Else
                If tWeek + i - 1 > 53 Then
                    oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 - 53 & " then apc13 else 0 end),0) as t" & i & ","
                Else
                    oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 & " then apc13 else 0 end),0) as t" & i & ","
                End If
            End If
        Next
        oCommand.CommandText += "1 from apa_file left join apc_file on apa01 = apc01 left join azn_file on apa12 = azn01 where apa41 = 'Y' and apa36 in ('002','003','015') and apc13 > 0 and apaud04 = '11' and apa12 between to_date('"
        oCommand.CommandText += tTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & fTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')  )"

        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 2 Step 1
                    Ws.Cells(11, 3 + i) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()

        ' 第十四列
        oCommand.CommandText = "select "
        For i As Int16 = 1 To 22 Step 1
            oCommand.CommandText += "sum(t" & i & ") as t" & i & ","
        Next
        oCommand.CommandText += "1 from ( select "
        For i As Int16 = 1 To 22 Step 1
            If tWeek > fWeek Then
                oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 & " then tc_ext05 else 0 end),0) as t" & i & ","
            Else
                If tWeek + i - 1 > 53 Then
                    oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 - 53 & " then tc_ext05 else 0 end),0) as t" & i & ","
                Else
                    oCommand.CommandText += "nvl((case when azn05 = " & tWeek + i - 1 & " then tc_ext05 else 0 end),0) as t" & i & ","
                End If
            End If
        Next
        oCommand.CommandText += "1 from tc_ext_file left join azn_file on tc_ext01 = azn01 where tc_ext02 = 2 and tc_ext01 between to_date('"
        oCommand.CommandText += tTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & fTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') )"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 2 Step 1
                    Ws.Cells(14, 3 + i) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()

    End Sub
    Private Sub AdjustExcelFormat()
        'oCommand.CommandText = "select distinct azn05 from azn_file where azn02 = 2019 and azn05 >= 23 order by azn05"
        'oReader = oCommand.ExecuteReader
        'If oReader.HasRows() Then
        'Dim WeekC As Decimal = 0
        'While oReader.Read()
        'Ws.Cells(1, 3 + WeekC) = "Week " & oReader.Item(0)
        'WeekC += 1
        'End While
        'End If
        'oReader.Close()
        If tWeek < fWeek Then
            For i As Int16 = 1 To 22 Step 1
                Ws.Cells(1, 2 + i) = "Week " & tWeek + i - 1
            Next
        Else
            For i As Int16 = 1 To 53 - tWeek + 1 Step 1
                Ws.Cells(1, 2 + i) = "Week " & tWeek + i - 1
            Next
            For i As Int16 = 1 To fWeek Step 1
                Ws.Cells(1, 2 + i + 53 - tWeek + 1) = "Week " & i
            Next
        End If
    End Sub
End Class