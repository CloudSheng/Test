Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form369
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim eYear As Int16 = 0
    Dim eMonth As Int16 = 0
    Dim LineZ As Integer = 0
    Dim TotalMonth As Int16 = 0
    Dim l_aah04_05 As Decimal = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form369_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Dim xPath As String = "C:\temp\DAC 部门费用总计 实际预算%.xlsx"
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
                'oCommand2.Connection = oConnection
                'oCommand2.CommandType = CommandType.Text
                'oCommand3.Connection = oConnection
                'oCommand3.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        tYear = Me.DateTimePicker1.Value.Year
        tMonth = Me.DateTimePicker1.Value.Month
        eYear = Me.DateTimePicker2.Value.Year
        eMonth = Me.DateTimePicker2.Value.Month
        If tYear <> eYear Then
            MsgBox("不同年度不能处理")
            Return
        End If
        If eMonth < tMonth Then
            MsgBox("月度有误")
            Return
        End If
        TotalMonth = eMonth - tMonth + 1

        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "DAC_部门费用_总计实际预算%"
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
        Dim xPath As String = "C:\temp\DAC 部门费用总计 实际预算%.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)

        'DAC 费用总计 实际预算% 表一 
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        LineZ = 6
        oCommand.CommandText = "select aao02,gem02,gem06"
        For i As Int16 = 1 To TotalMonth Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        For j As Int16 = 1 To TotalMonth Step 1
            oCommand.CommandText += ",sum(t" & 12 + j & ") as t" & 12 + j
        Next
        oCommand.CommandText += " from ( select aao02"
        For i As Int16 = 1 To TotalMonth Step 1
            oCommand.CommandText += ",sum(case when aao04 = " & i & " then aao05 - aao06 else 0 end) as t" & i
        Next
        For j As Int16 = 1 To TotalMonth Step 1
            oCommand.CommandText += ",0 as t" & 12 + j
        Next
        oCommand.CommandText += " from aao_file where aao03 = " & tYear & " and aao01 in ('5101','6601','6602','6604') and aao04 > 0 and aao02 <> 'D9999' group by aao02 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select tc_bud08"
        For i As Int16 = 1 To TotalMonth Step 1
            oCommand.CommandText += ",0"
        Next
        For j As Int16 = 1 To TotalMonth Step 1
            oCommand.CommandText += ",sum(case when tc_bud03 = " & j & " then tc_bud13 else 0 end) as t" & 12 + j
        Next
        oCommand.CommandText += " from tc_bud_file where tc_bud01 = 2 and tc_bud02 = " & tYear & " and (tc_bud07 like '5101%' or tc_bud07 like '6601%' or tc_bud07 like '6602%' or tc_bud07 like '6604%') and tc_bud08 <> 'D9999'  group by tc_bud08 ) left join gem_file on aao02 = gem01 group by aao02,gem02,gem06 order by aao02"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("aao02")
                Ws.Cells(LineZ, 2) = oReader.Item("gem02")
                Ws.Cells(LineZ, 3) = oReader.Item("gem06")
                For i As Int16 = 1 To TotalMonth Step 1
                    Ws.Cells(LineZ, 3 * i + 1) = oReader.Item(i + 2)
                Next
                For j As Int16 = 1 To TotalMonth Step 1
                    Ws.Cells(LineZ, 3 * j + 2) = oReader.Item(2 + TotalMonth + j)
                Next
                Dim AAA As String = String.Empty
                Dim BBB As String = String.Empty
                For k As Int16 = 1 To TotalMonth Step 1
                    Select Case k
                        Case 1
                            'AAA = "=D" & LineZ & "/E" & LineZ
                            AAA = "=(D" & LineZ & "-E" & LineZ & ")/E" & LineZ
                        Case 2
                            'AAA = "=G" & LineZ & "/H" & LineZ
                            AAA = "=(G" & LineZ & "-H" & LineZ & ")/H" & LineZ
                        Case 3
                            'AAA = "=J" & LineZ & "/K" & LineZ
                            AAA = "=(J" & LineZ & "-K" & LineZ & ")/K" & LineZ
                        Case 4
                            'AAA = "=M" & LineZ & "/N" & LineZ
                            AAA = "=(M" & LineZ & "-N" & LineZ & ")/N" & LineZ
                        Case 5
                            'AAA = "=P" & LineZ & "/Q" & LineZ
                            AAA = "=(P" & LineZ & "-Q" & LineZ & ")/Q" & LineZ
                        Case 6
                            'AAA = "=S" & LineZ & "/T" & LineZ
                            AAA = "=(S" & LineZ & "-T" & LineZ & ")/T" & LineZ
                        Case 7
                            'AAA = "=V" & LineZ & "/W" & LineZ
                            AAA = "=(V" & LineZ & "-W" & LineZ & ")/W" & LineZ
                        Case 8
                            'AAA = "=Y" & LineZ & "/Z" & LineZ
                            AAA = "=(Y" & LineZ & "-Z" & LineZ & ")/Z" & LineZ
                        Case 9
                            'AAA = "=AB" & LineZ & "/AC" & LineZ
                            AAA = "=(AB" & LineZ & "-AC" & LineZ & ")/AC" & LineZ
                        Case 10
                            'AAA = "=AE" & LineZ & "/AF" & LineZ
                            AAA = "=(AE" & LineZ & "-AF" & LineZ & ")/AF" & LineZ
                        Case 11
                            'AAA = "=AH" & LineZ & "/AI" & LineZ
                            AAA = "=(AH" & LineZ & "-AI" & LineZ & ")/AI" & LineZ
                        Case 12
                            'AAA = "=AK" & LineZ & "/AL" & LineZ
                            AAA = "=(AK" & LineZ & "-AL" & LineZ & ")/AL" & LineZ
                    End Select
                    Ws.Cells(LineZ, 3 * k + 3) = AAA
                Next
                'Ws.Cells(LineZ, 42) = "=AN" & LineZ & "/AO" & LineZ
                Ws.Cells(LineZ, 42) = "=(AN" & LineZ & "-AO" & LineZ & ")/AO" & LineZ
                Select Case TotalMonth
                    Case 1
                        AAA = "=D" & LineZ
                        BBB = "=E" & LineZ
                    Case 2
                        AAA = "=D" & LineZ & "+G" & LineZ
                        BBB = "=E" & LineZ & "+H" & LineZ
                    Case 3
                        AAA = "=D" & LineZ & "+G" & LineZ & "+J" & LineZ
                        BBB = "=E" & LineZ & "+H" & LineZ & "+K" & LineZ
                    Case 4
                        AAA = "=D" & LineZ & "+G" & LineZ & "+J" & LineZ & "+M" & LineZ
                        BBB = "=E" & LineZ & "+H" & LineZ & "+K" & LineZ & "+N" & LineZ
                    Case 5
                        AAA = "=D" & LineZ & "+G" & LineZ & "+J" & LineZ & "+M" & LineZ & "+P" & LineZ
                        BBB = "=E" & LineZ & "+H" & LineZ & "+K" & LineZ & "+N" & LineZ & "+Q" & LineZ
                    Case 6
                        AAA = "=D" & LineZ & "+G" & LineZ & "+J" & LineZ & "+M" & LineZ & "+P" & LineZ & "+S" & LineZ
                        BBB = "=E" & LineZ & "+H" & LineZ & "+K" & LineZ & "+N" & LineZ & "+Q" & LineZ & "+T" & LineZ
                    Case 7
                        AAA = "=D" & LineZ & "+G" & LineZ & "+J" & LineZ & "+M" & LineZ & "+P" & LineZ & "+S" & LineZ & "+V" & LineZ
                        BBB = "=E" & LineZ & "+H" & LineZ & "+K" & LineZ & "+N" & LineZ & "+Q" & LineZ & "+T" & LineZ & "+W" & LineZ
                    Case 8
                        AAA = "=D" & LineZ & "+G" & LineZ & "+J" & LineZ & "+M" & LineZ & "+P" & LineZ & "+S" & LineZ & "+V" & LineZ & "+Y" & LineZ
                        BBB = "=E" & LineZ & "+H" & LineZ & "+K" & LineZ & "+N" & LineZ & "+Q" & LineZ & "+T" & LineZ & "+W" & LineZ & "+Z" & LineZ
                    Case 9
                        AAA = "=D" & LineZ & "+G" & LineZ & "+J" & LineZ & "+M" & LineZ & "+P" & LineZ & "+S" & LineZ & "+V" & LineZ & "+Y" & LineZ & "+AB" & LineZ
                        BBB = "=E" & LineZ & "+H" & LineZ & "+K" & LineZ & "+N" & LineZ & "+Q" & LineZ & "+T" & LineZ & "+W" & LineZ & "+Z" & LineZ & "+AC" & LineZ
                    Case 10
                        AAA = "=D" & LineZ & "+G" & LineZ & "+J" & LineZ & "+M" & LineZ & "+P" & LineZ & "+S" & LineZ & "+V" & LineZ & "+Y" & LineZ & "+AB" & LineZ & "+AE" & LineZ
                        BBB = "=E" & LineZ & "+H" & LineZ & "+K" & LineZ & "+N" & LineZ & "+Q" & LineZ & "+T" & LineZ & "+W" & LineZ & "+Z" & LineZ & "+AC" & LineZ & "+AF" & LineZ
                    Case 11
                        AAA = "=D" & LineZ & "+G" & LineZ & "+J" & LineZ & "+M" & LineZ & "+P" & LineZ & "+S" & LineZ & "+V" & LineZ & "+Y" & LineZ & "+AB" & LineZ & "+AE" & LineZ & "+AH" & LineZ
                        BBB = "=E" & LineZ & "+H" & LineZ & "+K" & LineZ & "+N" & LineZ & "+Q" & LineZ & "+T" & LineZ & "+W" & LineZ & "+Z" & LineZ & "+AC" & LineZ & "+AF" & LineZ & "+AI" & LineZ
                    Case 12
                        AAA = "=D" & LineZ & "+G" & LineZ & "+J" & LineZ & "+M" & LineZ & "+P" & LineZ & "+S" & LineZ & "+V" & LineZ & "+Y" & LineZ & "+AB" & LineZ & "+AE" & LineZ & "+AH" & LineZ & "+AK" & LineZ
                        BBB = "=E" & LineZ & "+H" & LineZ & "+K" & LineZ & "+N" & LineZ & "+Q" & LineZ & "+T" & LineZ & "+W" & LineZ & "+Z" & LineZ & "+AC" & LineZ & "+AF" & LineZ & "+AI" & LineZ & "+AL" & LineZ
                End Select
                Ws.Cells(LineZ, 40) = AAA
                Ws.Cells(LineZ, 41) = BBB
                LineZ += 1
            End While
            For L As Int16 = 1 To TotalMonth Step 1
                Select Case L
                    Case 1
                        Ws.Cells(40, 4) = "=SUM(D6:D39)"
                        Ws.Cells(40, 5) = "=SUM(E6:E39)"
                    Case 2
                        Ws.Cells(40, 7) = "=SUM(G6:G39)"
                        Ws.Cells(40, 8) = "=SUM(H6:H39)"
                    Case 3
                        Ws.Cells(40, 10) = "=SUM(J6:J39)"
                        Ws.Cells(40, 11) = "=SUM(K6:K39)"
                    Case 4
                        Ws.Cells(40, 13) = "=SUM(M6:M39)"
                        Ws.Cells(40, 14) = "=SUM(N6:N39)"
                    Case 5
                        Ws.Cells(40, 16) = "=SUM(P6:P39)"
                        Ws.Cells(40, 17) = "=SUM(Q6:Q39)"
                    Case 6
                        Ws.Cells(40, 19) = "=SUM(S6:S39)"
                        Ws.Cells(40, 20) = "=SUM(T6:T39)"
                    Case 7
                        Ws.Cells(40, 22) = "=SUM(V6:V39)"
                        Ws.Cells(40, 23) = "=SUM(W6:W39)"
                    Case 8
                        Ws.Cells(40, 25) = "=SUM(Y6:Y39)"
                        Ws.Cells(40, 26) = "=SUM(Z6:Z39)"
                    Case 9
                        Ws.Cells(40, 28) = "=SUM(AB6:AB39)"
                        Ws.Cells(40, 29) = "=SUM(AC6:AC39)"
                    Case 10
                        Ws.Cells(40, 31) = "=SUM(AE6:AE39)"
                        Ws.Cells(40, 32) = "=SUM(AF6:AF39)"
                    Case 11
                        Ws.Cells(40, 34) = "=SUM(AH6:AH39)"
                        Ws.Cells(40, 35) = "=SUM(AI6:AI39)"
                    Case 12
                        Ws.Cells(40, 37) = "=SUM(AK6:AK39)"
                        Ws.Cells(40, 38) = "=SUM(AL6:AL39)"
                End Select
            Next
            Ws.Cells(40, 40) = "=SUM(AN6:AN39)"
            Ws.Cells(40, 41) = "=SUM(AO6:AO39)"
        End If
        oReader.Close()
        oRng = Ws.Range("A6", Ws.Cells(LineZ - 1, 42))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("F6", "F39")
        oRng.NumberFormatLocal = "0.00%"
        oRng = Ws.Range("I6", "I39")
        oRng.NumberFormatLocal = "0.00%"
        oRng = Ws.Range("L6", "L39")
        oRng.NumberFormatLocal = "0.00%"
        oRng = Ws.Range("O6", "O39")
        oRng.NumberFormatLocal = "0.00%"

        oRng = Ws.Range("R6", "R39")
        oRng.NumberFormatLocal = "0.00%"
        oRng = Ws.Range("U6", "U39")
        oRng.NumberFormatLocal = "0.00%"
        oRng = Ws.Range("X6", "X39")
        oRng.NumberFormatLocal = "0.00%"
        oRng = Ws.Range("AA6", "AA39")
        oRng.NumberFormatLocal = "0.00%"

        oRng = Ws.Range("AD6", "AD39")
        oRng.NumberFormatLocal = "0.00%"
        oRng = Ws.Range("AG6", "AG39")
        oRng.NumberFormatLocal = "0.00%"
        oRng = Ws.Range("AJ6", "AJ39")
        oRng.NumberFormatLocal = "0.00%"
        oRng = Ws.Range("AM6", "AM39")
        oRng.NumberFormatLocal = "0.00%"

        oRng = Ws.Range("AP6", "AP39")
        oRng.NumberFormatLocal = "0.00%"

        For M As Int16 = 1 To TotalMonth Step 1
            Select Case M
                Case 1
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 1 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(42, 4) = l_aah04_05
                    End If
                    oReader.Close()
                    'Ws.Cells(42, 5) = 16648075.1460975
                    Ws.Cells(43, 4) = "=SUM(D6:D39)/D42"
                    Ws.Cells(43, 5) = "=SUM(E6:E39)/E42"
                Case 2
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 2 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(42, 7) = l_aah04_05
                    End If
                    oReader.Close()
                    'Ws.Cells(42, 8) = 18577805.4844597
                    Ws.Cells(43, 7) = "=SUM(G6:G39)/G42"
                    Ws.Cells(43, 8) = "=SUM(H6:H39)/H42"
                Case 3
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 3 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(42, 10) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(43, 10) = "=SUM(J6:J39)/J42"
                    Ws.Cells(43, 11) = "=SUM(K6:K39)/K42"
                Case 4
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 4 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(42, 13) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(43, 13) = "=SUM(M6:M39)/M42"
                    Ws.Cells(43, 14) = "=SUM(N6:N39)/N42"
                Case 5
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 5 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(42, 16) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(43, 16) = "=SUM(P6:P39)/P42"
                    Ws.Cells(43, 17) = "=SUM(Q6:Q39)/Q42"
                Case 6
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 6 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(42, 19) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(43, 19) = "=SUM(S6:S39)/S42"
                    Ws.Cells(43, 20) = "=SUM(T6:T39)/T42"
                Case 7
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 7 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(42, 22) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(43, 22) = "=SUM(V6:V39)/V42"
                    Ws.Cells(43, 23) = "=SUM(W6:W39)/W42"
                Case 8
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 8 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(42, 25) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(43, 25) = "=SUM(Y6:Y39)/Y42"
                    Ws.Cells(43, 26) = "=SUM(Z6:Z39)/Z42"
                Case 9
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 9 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(42, 28) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(43, 28) = "=SUM(AB6:AB39)/AB42"
                    Ws.Cells(43, 29) = "=SUM(AC6:AC39)/AC42"
                Case 10
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 10 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(42, 31) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(43, 31) = "=SUM(AE6:AE39)/AE42"
                    Ws.Cells(43, 32) = "=SUM(AF6:AF39)/AF42"
                Case 11
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 11 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(42, 34) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(43, 34) = "=SUM(AH6:AH39)/AH42"
                    Ws.Cells(43, 35) = "=SUM(AI6:AI39)/AI42"
                Case 12
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 12 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(42, 37) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(43, 37) = "=SUM(AK6:AK39)/AK42"
                    Ws.Cells(43, 38) = "=SUM(AL6:AL39)/AL42"
            End Select
        Next
        oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
        oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aag01 in ('600101') "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            oReader.Read()
            l_aah04_05 = oReader.Item("aah04_05")
            If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
            Ws.Cells(42, 40) = l_aah04_05
        End If
        oReader.Close()
        Ws.Cells(43, 40) = "=SUM(AN6:AN39)/AN42"
        Ws.Cells(43, 41) = "=SUM(AO6:AO39)/AO42"

        oRng = Ws.Range("D43", "AP43")
        oRng.NumberFormatLocal = "0.00%"

        '收入达成% 表二
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        For N As Int16 = 1 To TotalMonth Step 1
            Select Case N
                Case 1
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 1 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(4, 3) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(7, 3) = "=C5*C6"
                    'Ws.Cells(9, 3) = "=C8-C6"
                    'Ws.Cells(10, 3) = "=C9/C6"
                    Ws.Cells(9, 3) = "=C8-C7"
                    Ws.Cells(10, 3) = "=C9/C7"
                Case 2
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 2 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(4, 4) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(7, 4) = "=D5*D6"
                    'Ws.Cells(9, 4) = "=D8-D6"
                    'Ws.Cells(10, 4) = "=D9/D6"
                    Ws.Cells(9, 4) = "=D8-D7"
                    Ws.Cells(10, 4) = "=D9/D7"
                Case 3
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 3 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(4, 5) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(7, 5) = "=E5*E6"
                    'Ws.Cells(9, 5) = "=E8-E6"
                    'Ws.Cells(10, 5) = "=E9/E6"
                    Ws.Cells(9, 5) = "=E8-E7"
                    Ws.Cells(10, 5) = "=E9/E7"
                Case 4
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 4 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(4, 6) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(7, 6) = "=F5*F6"
                    'Ws.Cells(9, 6) = "=F8-F6"
                    'Ws.Cells(10, 6) = "=F9/F6"
                    Ws.Cells(9, 6) = "=F8-F7"
                    Ws.Cells(10, 6) = "=F9/F7"
                Case 5
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 5 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(4, 7) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(7, 7) = "=G5*G6"
                    'Ws.Cells(9, 7) = "=G8-G6"
                    'Ws.Cells(10, 7) = "=G9/G6"
                    Ws.Cells(9, 7) = "=G8-G7"
                    Ws.Cells(10, 7) = "=G9/G7"
                Case 6
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 6 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(4, 8) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(7, 8) = "=H5*H6"
                    'Ws.Cells(9, 8) = "=H8-H6"
                    'Ws.Cells(10, 8) = "=H9/H6"
                    Ws.Cells(9, 8) = "=H8-H7"
                    Ws.Cells(10, 8) = "=H9/H7"
                Case 7
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 7 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(4, 9) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(7, 9) = "=I5*I6"
                    'Ws.Cells(9, 9) = "=I8-I6"
                    'Ws.Cells(10, 9) = "=I9/I6"
                    Ws.Cells(9, 9) = "=I8-I7"
                    Ws.Cells(10, 9) = "=I9/I7"
                Case 8
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 8 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(4, 10) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(7, 10) = "=J5*J6"
                    'Ws.Cells(9, 10) = "=J8-J6"
                    'Ws.Cells(10, 10) = "=J9/J6"
                    Ws.Cells(9, 10) = "=J8-J7"
                    Ws.Cells(10, 10) = "=J9/J7"
                Case 9
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 9 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(4, 11) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(7, 11) = "=K5*K6"
                    'Ws.Cells(9, 11) = "=K8-K6"
                    'Ws.Cells(10, 11) = "=K9/K6"
                    Ws.Cells(9, 11) = "=K8-K7"
                    Ws.Cells(10, 11) = "=K9/K7"
                Case 10
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 10 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(4, 12) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(7, 12) = "=L5*L6"
                    'Ws.Cells(9, 12) = "=L8-L6"
                    'Ws.Cells(10, 12) = "=L9/L6"
                    Ws.Cells(9, 12) = "=L8-L7"
                    Ws.Cells(10, 12) = "=L9/L7"
                Case 11
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 11 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(4, 13) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(7, 13) = "=M5*M6"
                    'Ws.Cells(9, 13) = "=M8-M6"
                    'Ws.Cells(10, 13) = "=M9/M6"
                    Ws.Cells(9, 13) = "=M8-M7"
                    Ws.Cells(10, 13) = "=M9/M7"
                Case 12
                    oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
                    oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aah03 = 12 AND aag01 in ('600101') "
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        oReader.Read()
                        l_aah04_05 = oReader.Item("aah04_05")
                        If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
                        Ws.Cells(4, 14) = l_aah04_05
                    End If
                    oReader.Close()
                    Ws.Cells(7, 14) = "=N5*N6"
                    'Ws.Cells(9, 14) = "=N8-N6"
                    'Ws.Cells(10, 14) = "=N9/N6"
                    Ws.Cells(9, 14) = "=N8-N7"
                    Ws.Cells(10, 14) = "=N9/N7"
            End Select
        Next
        oCommand.CommandText = "SELECT SUM(aah04-aah05) as aah04_05 FROM aag_file, aah_file WHERE aag03='2' AND aag01 = aah01 AND aag00 = aah00 "
        oCommand.CommandText += "  AND aah00 = '00' AND aah02 = " & tYear & " AND aag01 in ('600101') "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            oReader.Read()
            l_aah04_05 = oReader.Item("aah04_05")
            If l_aah04_05 < 0 Then l_aah04_05 = l_aah04_05 * -1
            Ws.Cells(4, 15) = l_aah04_05
        End If
        oReader.Close()
        Ws.Cells(7, 15) = "=O5*O6"
        'Ws.Cells(9, 15) = "=O8-O6"
        'Ws.Cells(10, 15) = "=O9/O6"
        Ws.Cells(9, 15) = "=O8-O7"
        Ws.Cells(10, 15) = "=O9/O7"

        oRng = Ws.Range("C10", "O10")
        oRng.NumberFormatLocal = "0.00%"

    End Sub
End Class