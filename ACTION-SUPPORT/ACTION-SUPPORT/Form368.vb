Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form368
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim xPath As String = String.Empty
    Dim Record1 As Boolean = False  ' 若 True 就不記錄
    Dim ReportType As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form186_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        Me.NumericUpDown1.Value = Today.Year
        Me.NumericUpDown2.Value = Today.Month
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        'If RadioButton1.Checked = True Then
        '    ReportType = 1
        'Else
        '    ReportType = 2
        'End If

        'If ReportType = 1 Then
        '    xPath = "C:\temp\STD_GM_Template.xlsx"
        'Else
        '    xPath = "C:\temp\ACT_GM_Template.xlsx"
        'End If

        xPath = "C:\temp\Std_Vs_Act_Template.xlsx"

        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
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
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        tYear = Me.NumericUpDown1.Value
        tMonth = Me.NumericUpDown2.Value
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        Dim Name1 As String = String.Empty
        'If ReportType = 1 Then
        '    Name1 = "标准毛利率报表"
        'Else
        '    Name1 = "实际毛利率报表"
        'End If

        Name1 = "Std Vs Act"

        SaveFileDialog1.FileName = Name1
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
        'If ReportType = 1 Then
        '    xPath = "C:\temp\STD_GM_Template.xlsx"
        'Else
        '    xPath = "C:\temp\ACT_GM_Template.xlsx"
        'End If

        xPath = "C:\temp\Std_Vs_Act_Template.xlsx"
        
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        'For i As Int16 = 1 To 12 Step 1
        '    Ws.Cells(2, 16 + i) = tYear & "/" & i & "/01"
        'Next
        LineZ = 6
        oCommand.CommandText = "select bma01,ima02,tqa02,ima25 from bma_file left join ima_file on bma01 = ima01 left join tqa_file on tqa03 = '2' and ima1005 = tqa01 where ima06 = '103' and bma10 = 2 and bmaacti = 'Y'"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Record1 = False
                DetailedData(tYear, tMonth, oReader.Item(0))
                DetailedData3(tYear, tMonth, oReader.Item(0))

                

                If Record1 = False Then
                    Ws.Cells(LineZ, 1) = oReader.Item(0)
                    Ws.Cells(LineZ, 2) = oReader.Item(1)
                    Ws.Cells(LineZ, 3) = oReader.Item(2)
                    Ws.Cells(LineZ, 4) = oReader.Item(3)
                    Ws.Cells(LineZ, 5) = "=IF(OR(Q" & LineZ & "=0,AC" & LineZ & "=0),0,IFERROR((AC" & LineZ & "-Q" & LineZ & ")/Q" & LineZ & ",))"
                    Ws.Cells(LineZ, 6) = "=IF(OR(R" & LineZ & "=0,AD" & LineZ & "=0),0,IFERROR((AD" & LineZ & "-R" & LineZ & ")/R" & LineZ & ",))"
                    Ws.Cells(LineZ, 7) = "=IF(OR(S" & LineZ & "=0,AE" & LineZ & "=0),0,IFERROR((AE" & LineZ & "-S" & LineZ & ")/S" & LineZ & ",))"
                    Ws.Cells(LineZ, 8) = "=IF(OR(T" & LineZ & "=0,AF" & LineZ & "=0),0,IFERROR((AF" & LineZ & "-T" & LineZ & ")/T" & LineZ & ",))"
                    Ws.Cells(LineZ, 9) = "=IF(OR(U" & LineZ & "=0,AG" & LineZ & "=0),0,IFERROR((AG" & LineZ & "-U" & LineZ & ")/U" & LineZ & ",))"
                    Ws.Cells(LineZ, 10) = "=IF(OR(V" & LineZ & "=0,AH" & LineZ & "=0),0,IFERROR((AH" & LineZ & "-V" & LineZ & ")/V" & LineZ & ",))"
                    Ws.Cells(LineZ, 11) = "=IF(OR(W" & LineZ & "=0,AI" & LineZ & "=0),0,IFERROR((AI" & LineZ & "-W" & LineZ & ")/W" & LineZ & ",))"
                    Ws.Cells(LineZ, 12) = "=IF(OR(X" & LineZ & "=0,AJ" & LineZ & "=0),0,IFERROR((AJ" & LineZ & "-X" & LineZ & ")/X" & LineZ & ",))"
                    Ws.Cells(LineZ, 13) = "=IF(OR(Y" & LineZ & "=0,AK" & LineZ & "=0),0,IFERROR((AK" & LineZ & "-Y" & LineZ & ")/Y" & LineZ & ",))"
                    Ws.Cells(LineZ, 14) = "=IF(OR(Z" & LineZ & "=0,AL" & LineZ & "=0),0,IFERROR((AL" & LineZ & "-Z" & LineZ & ")/Z" & LineZ & ",))"
                    Ws.Cells(LineZ, 15) = "=IF(OR(AA" & LineZ & "=0,AM" & LineZ & "=0),0,IFERROR((AM" & LineZ & "-AA" & LineZ & ")/AA" & LineZ & ",))"
                    Ws.Cells(LineZ, 16) = "=IF(OR(AB" & LineZ & "=0,AN" & LineZ & "=0),0,IFERROR((AN" & LineZ & "-AB" & LineZ & ")/AB" & LineZ & ",))"
                    DetailData2(tYear, tMonth, oReader.Item(0))
                    'Ws.Cells(LineZ, 18) = "=(SUM(E" & LineZ & ":P" & LineZ & ")-SUMIF(E" & LineZ & ":P" & LineZ & ",""1"",E" & LineZ & ":P" & LineZ & "))/(COUNTA(E" & LineZ & ":P" & LineZ & ")-COUNTIF(E" & LineZ & ":P" & LineZ & ",""=0"")-COUNTIF(E" & LineZ & ":P" & LineZ & ",""=1""))"
                    LineZ += 1
                    Label3.Text = LineZ
                    Label3.Refresh()
                End If
            End While
        End If
        oReader.Close()
        oRng = Ws.Range("A1", "AO1")
        oRng.EntireColumn.AutoFit()

        'For i As Int16 = 1 To 12 Step 1
        '    Ws.Cells(2, 28 + i) = tYear & "/" & i & "/01"
        'Next
        'LineZ = 6
        'oCommand.CommandText = "select bma01,ima02,tqa02,ima25 from bma_file left join ima_file on bma01 = ima01 left join tqa_file on tqa03 = '2' and ima1005 = tqa01 where ima06 = '103' and bma10 = 2 and bmaacti = 'Y'"
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        Record1 = False
        '        DetailedData3(tYear, tMonth, oReader.Item(0))
        '        If Record1 = False Then
        '            'Ws.Cells(LineZ, 1) = oReader.Item(0)
        '            'Ws.Cells(LineZ, 2) = oReader.Item(1)
        '            'Ws.Cells(LineZ, 3) = oReader.Item(2)
        '            'Ws.Cells(LineZ, 4) = oReader.Item(3)
        '            DetailData2(tYear, tMonth, oReader.Item(0))
        '            'Ws.Cells(LineZ, 18) = "=(SUM(E" & LineZ & ":P" & LineZ & ")-SUMIF(E" & LineZ & ":P" & LineZ & ",""1"",E" & LineZ & ":P" & LineZ & "))/(COUNTA(E" & LineZ & ":P" & LineZ & ")-COUNTIF(E" & LineZ & ":P" & LineZ & ",""=0"")-COUNTIF(E" & LineZ & ":P" & LineZ & ",""=1""))"
        '            LineZ += 1
        '            Label3.Text = LineZ
        '            Label3.Refresh()
        '        End If
        '    End While
        'End If
        'oReader.Close()
        'oRng = Ws.Range("A1", "R1")
        'oRng.EntireColumn.AutoFit()
    End Sub
    Private Sub DetailedData(ByVal Year1 As Int16, ByVal Month1 As Int16, ByVal ima01 As String)
        For i As Int16 = Month1 To 1 Step -1
            oCommand2.CommandText = "Select nvl(Round((stb07+stb08+stb09+stb09a) /ex1.er,4),0) from stb_file left join exchangeratebyyear ex1 on ex1.year1 = " & Year1 & " and ex1.currency = 'USD' "
            oCommand2.CommandText += "where stb01 = '" & ima01 & "' and stb02 = " & Year1 & " and stb03 = " & i

            'oCommand2.CommandText = "Select nvl(Round((stb07+stb08+stb09+stb09a) /ex1.er,4),0) from stb_file left join exchangeratebyyear ex1 on ex1.year1 = " & Year1 & " and ex1.currency = 'USD' "
            'oCommand2.CommandText += "where stb01 = '" & ima01 & "' and stb02 = " & Year1 & " and stb03 = " & i
            Dim STDCostUSD As Decimal = oCommand2.ExecuteScalar()
            If i = Month1 And STDCostUSD = 0 Then
                Record1 = True
                Exit For
            End If
            Dim TempDate As Date = Convert.ToDateTime(Year1 & "/" & i & "/01")
            Dim TempDate1 As Date = TempDate.AddMonths(1).AddDays(-1)
            oCommand2.CommandText = " Select (case when tc_prl06 ='USD' then tc_prl03 * tc_prl04 /100 else nvl(Round(tc_prl03 * tc_prl04 /100 * ex1.er / ex2.er,4),0) end) from tc_prl_file  left join exchangeratebyyear ex1 on ex1.year1 = " & Year1 & " and ex1.currency = tc_prl06 "
            oCommand2.CommandText += "left join exchangeratebyyear ex2 on ex2.year1 = " & Year1 & " and ex2.currency = 'USD' where tc_prl01  = '" & ima01 & "' and tc_prl02 >= to_date('" & TempDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') order by tc_prl02"
            oReader2 = oCommand2.ExecuteReader
            Dim SaleSRMB As Decimal = 0
            If oReader2.HasRows() Then
                oReader2.Read()
                SaleSRMB = oReader2.Item(0)
            Else
                If i = Month1 Then
                    Record1 = True
                    Exit For
                End If
            End If
            oReader2.Close()
            'Dim Perce1 As Decimal = Decimal.Round((SaleSRMB - STDCostUSD) / SaleSRMB, 4)
            Dim Perce1 As Decimal = Decimal.Round(STDCostUSD, 4)
            Ws.Cells(LineZ, i + 16) = Perce1
        Next
    End Sub
    Private Sub DetailedData3(ByVal Year1 As Int16, ByVal Month1 As Int16, ByVal ima01 As String)
        For i As Int16 = Month1 To 1 Step -1
            oCommand2.CommandText = "Select nvl(Round((ccc23) /ex1.er,4),0) from ccc_file left join exchangeratebyyear ex1 on ex1.year1 = " & Year1 & " and ex1.currency = 'USD' "
            oCommand2.CommandText += "where ccc01 = '" & ima01 & "' and ccc02 = " & Year1 & " and ccc03 = " & i

            'oCommand2.CommandText = "Select nvl(Round((stb07+stb08+stb09+stb09a) /ex1.er,4),0) from stb_file left join exchangeratebyyear ex1 on ex1.year1 = " & Year1 & " and ex1.currency = 'USD' "
            'oCommand2.CommandText += "where stb01 = '" & ima01 & "' and stb02 = " & Year1 & " and stb03 = " & i
            Dim STDCostUSD As Decimal = oCommand2.ExecuteScalar()
            If i = Month1 And STDCostUSD = 0 Then
                Record1 = True
                Exit For
            End If
            Dim TempDate As Date = Convert.ToDateTime(Year1 & "/" & i & "/01")
            Dim TempDate1 As Date = TempDate.AddMonths(1).AddDays(-1)
            oCommand2.CommandText = " Select (case when tc_prl06 ='USD' then tc_prl03 * tc_prl04 /100 else nvl(Round(tc_prl03 * tc_prl04 /100 * ex1.er / ex2.er,4),0) end) from tc_prl_file  left join exchangeratebyyear ex1 on ex1.year1 = " & Year1 & " and ex1.currency = tc_prl06 "
            oCommand2.CommandText += "left join exchangeratebyyear ex2 on ex2.year1 = " & Year1 & " and ex2.currency = 'USD' where tc_prl01  = '" & ima01 & "' and tc_prl02 >= to_date('" & TempDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') order by tc_prl02"
            oReader2 = oCommand2.ExecuteReader
            Dim SaleSRMB As Decimal = 0
            If oReader2.HasRows() Then
                oReader2.Read()
                SaleSRMB = oReader2.Item(0)
            Else
                If i = Month1 Then
                    Record1 = True
                    Exit For
                End If
            End If
            oReader2.Close()
            'Dim Perce1 As Decimal = Decimal.Round((SaleSRMB - STDCostUSD) / SaleSRMB, 4)
            Dim Perce1 As Decimal = Decimal.Round(STDCostUSD, 4)
            Ws.Cells(LineZ, i + 28) = Perce1
        Next
    End Sub
    Private Sub DetailData2(ByVal Year1 As Int16, ByVal Month1 As Int16, ByVal ima01 As String)
        oCommand2.CommandText = " Select nvl(Round(sum(ccc63)/ex1.er,4),0) from ccc_file  left join exchangeratebyyear ex1 on ex1.year1 = " & Year1 & " and ex1.currency = 'USD'  where ccc01 = '"
        oCommand2.CommandText += ima01 & "' and ccc02 = " & Year1 & " and ccc03 <= " & Month1 & " group by ex1.er"
        Dim SS As Decimal = oCommand2.ExecuteScalar()
        Ws.Cells(LineZ, 41) = SS
    End Sub
End Class