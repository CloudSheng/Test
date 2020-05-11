Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form161
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
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form161_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Dim xPath As String = "C:\temp\DAC  费用总计 实际 预算比较.xlsx"
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
        SaveFileDialog1.FileName = "部门费用及预算汇总表"
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
        Dim xPath As String = "C:\temp\DAC  费用总计 实际 预算比较.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
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
                            AAA = "=D" & LineZ & "-E" & LineZ
                        Case 2
                            AAA = "=G" & LineZ & "-H" & LineZ
                        Case 3
                            AAA = "=J" & LineZ & "-K" & LineZ
                        Case 4
                            AAA = "=M" & LineZ & "-N" & LineZ
                        Case 5
                            AAA = "=P" & LineZ & "-Q" & LineZ
                        Case 6
                            AAA = "=S" & LineZ & "-T" & LineZ
                        Case 7
                            AAA = "=V" & LineZ & "-W" & LineZ
                        Case 8
                            AAA = "=Y" & LineZ & "-Z" & LineZ
                        Case 9
                            AAA = "=AB" & LineZ & "-AC" & LineZ
                        Case 10
                            AAA = "=AE" & LineZ & "-AF" & LineZ
                        Case 11
                            AAA = "=AH" & LineZ & "-AI" & LineZ
                        Case 12
                            AAA = "=AK" & LineZ & "-AL" & LineZ
                    End Select
                    Ws.Cells(LineZ, 3 * k + 3) = AAA
                Next
                Ws.Cells(LineZ, 42) = "=AN" & LineZ & "-AO" & LineZ
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
        End If
        oReader.Close()
        oRng = Ws.Range("A6", Ws.Cells(LineZ - 1, 42))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
    End Sub
End Class