Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form162
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand3 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader3 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim pYear As Int16 = 0
    Dim tDate As Date
    Dim lYear As Int16 = 0
    Dim lMonth As Int16 = 0
    Dim LineZ As Integer = 0
    Dim DNP As String = String.Empty
    ' 2018/09/3
    Dim Ldate1 As Date
    Dim Ldate2 As Date
    ' 2018/09/04
    Dim TP As Decimal = 0
    Dim RS As Decimal = 0
    Dim X1 As Decimal = 0
    Dim X2 As Decimal = 0
    Dim X3 As Decimal = 0
    Dim X4 As Decimal = 0
    Dim X5 As Decimal = 0
    Dim X6 As Decimal = 0
    Dim X7 As Decimal = 0
    Dim X8 As Decimal = 0
    Dim X9 As Decimal = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form162_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        RS = 0
        If Me.RadioButton1.Checked Then
            RS = 1
        End If
        If Me.RadioButton2.Checked Then
            RS = 2
        End If
        If Me.RadioButton3.Checked Then
            RS = 3
        End If
        If Me.RadioButton4.Checked Then
            RS = 4
        End If
        If RS = 0 Then
            MsgBox("选择功能主管")
            Return
        End If
        tYear = Me.DateTimePicker1.Value.Year
        tMonth = Me.DateTimePicker1.Value.Month
        pYear = Me.DateTimePicker1.Value.AddYears(-1).Year
        tDate = Convert.ToDateTime(tYear & "/" & tMonth & "/01")
        pYear = tDate.AddYears(-1).Year
        lYear = Me.DateTimePicker1.Value.AddMonths(-1).Year
        lMonth = Me.DateTimePicker1.Value.AddMonths(-1).Month
        ' 2018/09/03
        Ldate1 = Convert.ToDateTime(tYear & "/01/01")
        Ldate2 = Convert.ToDateTime(tYear & "/" & tMonth & "/01").AddMonths(1).AddDays(-1)
        'ExportToExcel()
        'SaveExcel()
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        Select Case RS
            Case 1    ' 管理部主管
                xExcel = New Microsoft.Office.Interop.Excel.Application
                Dim xPath As String = "C:\temp\Exp-1.xlsx"
                If Not My.Computer.FileSystem.FileExists(xPath) Then
                    MsgBox("NO SAMPLE FILE")
                    Return
                End If
                xWorkBook = xExcel.Workbooks.Open(xPath)
                Ws = xWorkBook.Sheets(1)
                Ws.Activate()
                LineZ = 8
                AdjustExcelFormat()
                'oCommand.CommandText = "select distinct aao01,aag02,aao02 from aao_file left join aag_file on aao01 = aag01 where aao01 in ('510101','660101','660201','660401') and aao02 <> 'D9999'"
                oCommand.CommandText = "select distinct aao01,aag02,aao02 from ( select aao01,aag02,aao02 from aao_file left join aag_file on aao01 = aag01 where aao01 in ('510101','660101','660201','660401') and aao02 <> 'D9999' "
                oCommand.CommandText += "union all "
                oCommand.CommandText += "select tc_bud07,aag02,tc_bud08 from tc_bud_file left join aag_file on tc_bud07 = aag01 where tc_bud01 = 2 and tc_bud02 = 2019 and tc_bud07 in ('510101','660101','660201','660401') and tc_bud08 <> 'D9999' ) order by aao02,aao01"
                oReader = oCommand.ExecuteReader()
                If oReader.HasRows() Then
                    While oReader.Read()
                        X1 = Decimal.Round(GetLastYearSameMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X2 = Decimal.Round(GetLastMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X3 = Decimal.Round(GetThisYearSameMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X4 = GetThisYearSameMonthBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        X5 = Decimal.Round(GetLastYearBeforeMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X6 = Decimal.Round(GetThisYearBeforeMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X7 = GetThisYearBeforeMonthBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        X8 = Decimal.Round(GetLastYearNoMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X9 = GetThisYearBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        If X1 = 0 And X2 = 0 And X3 = 0 And X4 = 0 And X5 = 0 And X6 = 0 And X7 = 0 And X8 = 0 And X9 = 0 Then
                            Continue While
                        End If
                        Ws.Cells(LineZ, 2) = oReader.Item("aao01")
                        Ws.Cells(LineZ, 3) = oReader.Item("aag02")
                        Ws.Cells(LineZ, 4) = GetDepartNmae(oReader.Item("aao02"))
                        Ws.Cells(LineZ, 5) = GetDepartNmae("D0900")
                        Ws.Cells(LineZ, 6) = GetDepartBoss("D0900")
                        Ws.Cells(LineZ, 7) = X1
                        Ws.Cells(LineZ, 8) = X2
                        Ws.Cells(LineZ, 9) = X3
                        Ws.Cells(LineZ, 10) = X4
                        Ws.Cells(LineZ, 11) = "=I" & LineZ & "-J" & LineZ
                        Ws.Cells(LineZ, 12) = "=I" & LineZ & "-G" & LineZ
                        Ws.Cells(LineZ, 13) = "=I" & LineZ & "-H" & LineZ
                        Ws.Cells(LineZ, 14) = X5
                        Ws.Cells(LineZ, 15) = X6
                        Ws.Cells(LineZ, 16) = X7
                        Ws.Cells(LineZ, 17) = "=O" & LineZ & "-P" & LineZ
                        Ws.Cells(LineZ, 18) = "=O" & LineZ & "-N" & LineZ
                        Ws.Cells(LineZ, 19) = X8
                        Ws.Cells(LineZ, 20) = "=U" & LineZ & "-O" & LineZ
                        Ws.Cells(LineZ, 21) = X9
                        LineZ += 1
                    End While
                    Ws.Cells(LineZ, 6) = "总计"
                    Ws.Cells(LineZ, 7) = "=SUM(G8:G" & LineZ - 1 & ")"
                    ' 複制
                    oRng = Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 7))
                    oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 21)), Type:=xlFillDefault)
                End If
                oReader.Close()
                GG1() ' 劃線

                ' 第二頁
                Ws = xWorkBook.Sheets(2)
                Ws.Activate()
                LineZ = 8
                AdjustExcelFormat()
                'oCommand.CommandText = "select distinct aao01,aag02,aao02 from aao_file left join aag_file on aao01 = aag01 where aao01 in ('510104','660104','660204','660404') and aao02 <> 'D9999'"
                oCommand.CommandText = "select distinct aao01,aag02,aao02 from ( select aao01,aag02,aao02 from aao_file left join aag_file on aao01 = aag01 where aao01 in ('510104','660104','660204','660404') and aao02 <> 'D9999' "
                oCommand.CommandText += "union all "
                oCommand.CommandText += "select tc_bud07,aag02,tc_bud08 from tc_bud_file left join aag_file on tc_bud07 = aag01 where tc_bud01 = 2 and tc_bud02 = 2019 and tc_bud07 in ('510104','660104','660204','660404') and tc_bud08 <> 'D9999' ) order by aao02,aao01"
                oReader = oCommand.ExecuteReader()
                If oReader.HasRows() Then
                    While oReader.Read()
                        X1 = Decimal.Round(GetLastYearSameMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X2 = Decimal.Round(GetLastMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X3 = Decimal.Round(GetThisYearSameMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X4 = GetThisYearSameMonthBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        X5 = Decimal.Round(GetLastYearBeforeMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X6 = Decimal.Round(GetThisYearBeforeMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X7 = GetThisYearBeforeMonthBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        X8 = Decimal.Round(GetLastYearNoMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X9 = GetThisYearBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        If X1 = 0 And X2 = 0 And X3 = 0 And X4 = 0 And X5 = 0 And X6 = 0 And X7 = 0 And X8 = 0 And X9 = 0 Then
                            Continue While
                        End If
                        Ws.Cells(LineZ, 2) = oReader.Item("aao01")
                        Ws.Cells(LineZ, 3) = oReader.Item("aag02")
                        Ws.Cells(LineZ, 4) = GetDepartNmae(oReader.Item("aao02"))
                        Ws.Cells(LineZ, 5) = GetDepartNmae("D0900")
                        Ws.Cells(LineZ, 6) = GetDepartBoss("D0900")
                        Ws.Cells(LineZ, 7) = X1
                        Ws.Cells(LineZ, 8) = X2
                        Ws.Cells(LineZ, 9) = X3
                        Ws.Cells(LineZ, 10) = X4
                        Ws.Cells(LineZ, 11) = "=I" & LineZ & "-J" & LineZ
                        Ws.Cells(LineZ, 12) = "=I" & LineZ & "-G" & LineZ
                        Ws.Cells(LineZ, 13) = "=I" & LineZ & "-H" & LineZ
                        Ws.Cells(LineZ, 14) = X5
                        Ws.Cells(LineZ, 15) = X6
                        Ws.Cells(LineZ, 16) = X7
                        Ws.Cells(LineZ, 17) = "=O" & LineZ & "-P" & LineZ
                        Ws.Cells(LineZ, 18) = "=O" & LineZ & "-N" & LineZ
                        Ws.Cells(LineZ, 19) = X8
                        Ws.Cells(LineZ, 20) = "=U" & LineZ & "-O" & LineZ
                        Ws.Cells(LineZ, 21) = X9
                        LineZ += 1
                    End While
                    Ws.Cells(LineZ, 6) = "总计"
                    Ws.Cells(LineZ, 7) = "=SUM(G8:G" & LineZ - 1 & ")"
                    ' 複制
                    oRng = Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 7))
                    oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 21)), Type:=xlFillDefault)
                End If
                oReader.Close()
                GG1() ' 劃線

                ' 第三頁
                Ws = xWorkBook.Sheets(3)
                Ws.Activate()
                LineZ = 8
                AdjustExcelFormat()
                'oCommand.CommandText = "select distinct aao01,aag02,aao02 from aao_file left join aag_file on aao01 = aag01 where aao01 in ('510107','660112','660212','660418') and aao02 <> 'D9999'"
                oCommand.CommandText = "select distinct aao01,aag02,aao02 from ( select aao01,aag02,aao02 from aao_file left join aag_file on aao01 = aag01 where aao01 in ('510107','660112','660212','660418') and aao02 <> 'D9999' "
                oCommand.CommandText += "union all "
                oCommand.CommandText += "select tc_bud07,aag02,tc_bud08 from tc_bud_file left join aag_file on tc_bud07 = aag01 where tc_bud01 = 2 and tc_bud02 = 2019 and tc_bud07 in ('510107','660112','660212','660418') and tc_bud08 <> 'D9999' ) order by aao02,aao01"
                oReader = oCommand.ExecuteReader()
                If oReader.HasRows() Then
                    While oReader.Read()
                        X1 = Decimal.Round(GetLastYearSameMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X2 = Decimal.Round(GetLastMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X3 = Decimal.Round(GetThisYearSameMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X4 = GetThisYearSameMonthBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        X5 = Decimal.Round(GetLastYearBeforeMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X6 = Decimal.Round(GetThisYearBeforeMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X7 = GetThisYearBeforeMonthBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        X8 = Decimal.Round(GetLastYearNoMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X9 = GetThisYearBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        If X1 = 0 And X2 = 0 And X3 = 0 And X4 = 0 And X5 = 0 And X6 = 0 And X7 = 0 And X8 = 0 And X9 = 0 Then
                            Continue While
                        End If
                        Ws.Cells(LineZ, 2) = oReader.Item("aao01")
                        Ws.Cells(LineZ, 3) = oReader.Item("aag02")
                        Ws.Cells(LineZ, 4) = GetDepartNmae(oReader.Item("aao02"))
                        Ws.Cells(LineZ, 5) = GetDepartNmae("D0900")
                        Ws.Cells(LineZ, 6) = GetDepartBoss("D0900")
                        Ws.Cells(LineZ, 7) = X1
                        Ws.Cells(LineZ, 8) = X2
                        Ws.Cells(LineZ, 9) = X3
                        Ws.Cells(LineZ, 10) = X4
                        Ws.Cells(LineZ, 11) = "=I" & LineZ & "-J" & LineZ
                        Ws.Cells(LineZ, 12) = "=I" & LineZ & "-G" & LineZ
                        Ws.Cells(LineZ, 13) = "=I" & LineZ & "-H" & LineZ
                        Ws.Cells(LineZ, 14) = X5
                        Ws.Cells(LineZ, 15) = X6
                        Ws.Cells(LineZ, 16) = X7
                        Ws.Cells(LineZ, 17) = "=O" & LineZ & "-P" & LineZ
                        Ws.Cells(LineZ, 18) = "=O" & LineZ & "-N" & LineZ
                        Ws.Cells(LineZ, 19) = X8
                        Ws.Cells(LineZ, 20) = "=U" & LineZ & "-O" & LineZ
                        Ws.Cells(LineZ, 21) = X9
                        LineZ += 1
                    End While
                    Ws.Cells(LineZ, 6) = "总计"
                    Ws.Cells(LineZ, 7) = "=SUM(G8:G" & LineZ - 1 & ")"
                    ' 複制
                    oRng = Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 7))
                    oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 21)), Type:=xlFillDefault)
                End If
                oReader.Close()
                GG1() ' 劃線

                ' 第四頁
                Ws = xWorkBook.Sheets(4)
                Ws.Activate()
                LineZ = 8
                AdjustExcelFormat()
                'oCommand.CommandText = "select distinct aao01,aag02,aao02 from aao_file left join aag_file on aao01 = aag01 where aao01 in ('510122','660120','660218','660419') and aao02 <> 'D9999'"
                oCommand.CommandText = "select distinct aao01,aag02,aao02 from ( select aao01,aag02,aao02 from aao_file left join aag_file on aao01 = aag01 where aao01 in ('510122','660120','660218','660419') and aao02 <> 'D9999' "
                oCommand.CommandText += "union all "
                oCommand.CommandText += "select tc_bud07,aag02,tc_bud08 from tc_bud_file left join aag_file on tc_bud07 = aag01 where tc_bud01 = 2 and tc_bud02 = 2019 and tc_bud07 in ('510122','660120','660218','660419') and tc_bud08 <> 'D9999' ) order by aao02,aao01"
                oReader = oCommand.ExecuteReader()
                If oReader.HasRows() Then
                    While oReader.Read()
                        X1 = Decimal.Round(GetLastYearSameMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X2 = Decimal.Round(GetLastMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X3 = Decimal.Round(GetThisYearSameMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X4 = GetThisYearSameMonthBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        X5 = Decimal.Round(GetLastYearBeforeMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X6 = Decimal.Round(GetThisYearBeforeMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X7 = GetThisYearBeforeMonthBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        X8 = Decimal.Round(GetLastYearNoMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X9 = GetThisYearBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        If X1 = 0 And X2 = 0 And X3 = 0 And X4 = 0 And X5 = 0 And X6 = 0 And X7 = 0 And X8 = 0 And X9 = 0 Then
                            Continue While
                        End If
                        Ws.Cells(LineZ, 2) = oReader.Item("aao01")
                        Ws.Cells(LineZ, 3) = oReader.Item("aag02")
                        Ws.Cells(LineZ, 4) = GetDepartNmae(oReader.Item("aao02"))
                        Ws.Cells(LineZ, 5) = GetDepartNmae("D0900")
                        Ws.Cells(LineZ, 6) = GetDepartBoss("D0900")
                        Ws.Cells(LineZ, 7) = X1
                        Ws.Cells(LineZ, 8) = X2
                        Ws.Cells(LineZ, 9) = X3
                        Ws.Cells(LineZ, 10) = X4
                        Ws.Cells(LineZ, 11) = "=I" & LineZ & "-J" & LineZ
                        Ws.Cells(LineZ, 12) = "=I" & LineZ & "-G" & LineZ
                        Ws.Cells(LineZ, 13) = "=I" & LineZ & "-H" & LineZ
                        Ws.Cells(LineZ, 14) = X5
                        Ws.Cells(LineZ, 15) = X6
                        Ws.Cells(LineZ, 16) = X7
                        Ws.Cells(LineZ, 17) = "=O" & LineZ & "-P" & LineZ
                        Ws.Cells(LineZ, 18) = "=O" & LineZ & "-N" & LineZ
                        Ws.Cells(LineZ, 19) = X8
                        Ws.Cells(LineZ, 20) = "=U" & LineZ & "-O" & LineZ
                        Ws.Cells(LineZ, 21) = X9
                        LineZ += 1
                    End While
                    Ws.Cells(LineZ, 6) = "总计"
                    Ws.Cells(LineZ, 7) = "=SUM(G8:G" & LineZ - 1 & ")"
                    ' 複制
                    oRng = Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 7))
                    oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 21)), Type:=xlFillDefault)
                End If
                oReader.Close()
                GG1() ' 劃線
                TP = 4

                For i As Int16 = 1 To tMonth Step 1
                    If TP + i > 3 Then
                        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                    Else
                        Ws = xWorkBook.Sheets(TP + i)
                    End If
                    Ws.Activate()
                    AdjustExcelFormat1(i)
                    oCommand.CommandText = "select abb05,gem02,aba02,abb01,abb03,aag02,abb04,(case when abb06 = 1 then abb07 else abb07 * -1 end) as t1  from abb_file left join aba_file on abb01 = aba01 "
                    oCommand.CommandText += "left join aag_file on abb03 = aag01 left join gem_file on abb05 = gem01 where abapost = 'Y' and abb03 in ('510101','660101','660201','660401','510104','660104','660204','660404','510107','660112','660212','660418','510122','660120','660218','660419') AND abb05 <> 'D9999' "
                    oCommand.CommandText += "and aba03 = " & tYear & " and aba04 = " & i & " order by abb05"
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        While oReader.Read()
                            For j As Int16 = 0 To oReader.FieldCount - 1 Step 1
                                Ws.Cells(LineZ, j + 1) = oReader.Item(j)
                            Next
                            LineZ += 1
                        End While
                    End If
                    oReader.Close()
                Next

            Case 2
                xExcel = New Microsoft.Office.Interop.Excel.Application
                Dim xPath As String = "C:\temp\Exp-2.xlsx"
                If Not My.Computer.FileSystem.FileExists(xPath) Then
                    MsgBox("NO SAMPLE FILE")
                    Return
                End If
                xWorkBook = xExcel.Workbooks.Open(xPath)
                Ws = xWorkBook.Sheets(1)
                Ws.Activate()
                LineZ = 8
                AdjustExcelFormat()
                'oCommand.CommandText = "select distinct aao01,aag02,aao02 from aao_file left join aag_file on aao01 = aag01 where aao01 in ('660109','660209','660422') and aao02 <> 'D9999'"
                oCommand.CommandText = "select distinct aao01,aag02,aao02 from ( select aao01,aag02,aao02 from aao_file left join aag_file on aao01 = aag01 where aao01 in ('660109','660209','660422') and aao02 <> 'D9999' "
                oCommand.CommandText += "union all "
                oCommand.CommandText += "select tc_bud07,aag02,tc_bud08 from tc_bud_file left join aag_file on tc_bud07 = aag01 where tc_bud01 = 2 and tc_bud02 = 2019 and tc_bud07 in ('660109','660209','660422') and tc_bud08 <> 'D9999' ) order by aao02,aao01"
                oReader = oCommand.ExecuteReader()
                If oReader.HasRows() Then
                    While oReader.Read()
                        X1 = Decimal.Round(GetLastYearSameMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X2 = Decimal.Round(GetLastMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X3 = Decimal.Round(GetThisYearSameMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X4 = GetThisYearSameMonthBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        X5 = Decimal.Round(GetLastYearBeforeMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X6 = Decimal.Round(GetThisYearBeforeMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X7 = GetThisYearBeforeMonthBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        X8 = Decimal.Round(GetLastYearNoMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X9 = GetThisYearBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        If X1 = 0 And X2 = 0 And X3 = 0 And X4 = 0 And X5 = 0 And X6 = 0 And X7 = 0 And X8 = 0 And X9 = 0 Then
                            Continue While
                        End If
                        Ws.Cells(LineZ, 2) = oReader.Item("aao01")
                        Ws.Cells(LineZ, 3) = oReader.Item("aag02")
                        Ws.Cells(LineZ, 4) = GetDepartNmae(oReader.Item("aao02"))
                        Ws.Cells(LineZ, 5) = GetDepartNmae("D0210")
                        Ws.Cells(LineZ, 6) = GetDepartBoss("D0210")
                        Ws.Cells(LineZ, 7) = X1
                        Ws.Cells(LineZ, 8) = X2
                        Ws.Cells(LineZ, 9) = X3
                        Ws.Cells(LineZ, 10) = X4
                        Ws.Cells(LineZ, 11) = "=I" & LineZ & "-J" & LineZ
                        Ws.Cells(LineZ, 12) = "=I" & LineZ & "-G" & LineZ
                        Ws.Cells(LineZ, 13) = "=I" & LineZ & "-H" & LineZ
                        Ws.Cells(LineZ, 14) = X5
                        Ws.Cells(LineZ, 15) = X6
                        Ws.Cells(LineZ, 16) = X7
                        Ws.Cells(LineZ, 17) = "=O" & LineZ & "-P" & LineZ
                        Ws.Cells(LineZ, 18) = "=O" & LineZ & "-N" & LineZ
                        Ws.Cells(LineZ, 19) = X8
                        Ws.Cells(LineZ, 20) = "=U" & LineZ & "-O" & LineZ
                        Ws.Cells(LineZ, 21) = X9
                        LineZ += 1
                    End While
                    Ws.Cells(LineZ, 6) = "总计"
                    Ws.Cells(LineZ, 7) = "=SUM(G8:G" & LineZ - 1 & ")"
                    ' 複制
                    oRng = Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 7))
                    oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 21)), Type:=xlFillDefault)
                End If
                oReader.Close()
                GG1() ' 劃線

                ' 第二頁
                Ws = xWorkBook.Sheets(2)
                Ws.Activate()
                LineZ = 8
                AdjustExcelFormat()
                'oCommand.CommandText = "select distinct aao01,aag02,aao02 from aao_file left join aag_file on aao01 = aag01 where aao01 in ('510121') and aao02 <> 'D9999'"
                oCommand.CommandText = "select distinct aao01,aag02,aao02 from ( select aao01,aag02,aao02 from aao_file left join aag_file on aao01 = aag01 where aao01 in ('510121') and aao02 <> 'D9999' "
                oCommand.CommandText += "union all "
                oCommand.CommandText += "select tc_bud07,aag02,tc_bud08 from tc_bud_file left join aag_file on tc_bud07 = aag01 where tc_bud01 = 2 and tc_bud02 = 2019 and tc_bud07 in ('510121') and tc_bud08 <> 'D9999' ) order by aao02,aao01"
                oReader = oCommand.ExecuteReader()
                If oReader.HasRows() Then
                    While oReader.Read()
                        X1 = Decimal.Round(GetLastYearSameMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X2 = Decimal.Round(GetLastMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X3 = Decimal.Round(GetThisYearSameMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X4 = GetThisYearSameMonthBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        X5 = Decimal.Round(GetLastYearBeforeMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X6 = Decimal.Round(GetThisYearBeforeMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X7 = GetThisYearBeforeMonthBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        X8 = Decimal.Round(GetLastYearNoMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X9 = GetThisYearBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        If X1 = 0 And X2 = 0 And X3 = 0 And X4 = 0 And X5 = 0 And X6 = 0 And X7 = 0 And X8 = 0 And X9 = 0 Then
                            Continue While
                        End If
                        Ws.Cells(LineZ, 2) = oReader.Item("aao01")
                        Ws.Cells(LineZ, 3) = oReader.Item("aag02")
                        Ws.Cells(LineZ, 4) = GetDepartNmae(oReader.Item("aao02"))
                        Ws.Cells(LineZ, 5) = GetDepartNmae("D0210")
                        Ws.Cells(LineZ, 6) = GetDepartBoss("D0210")
                        Ws.Cells(LineZ, 7) = X1
                        Ws.Cells(LineZ, 8) = X2
                        Ws.Cells(LineZ, 9) = X3
                        Ws.Cells(LineZ, 10) = X4
                        Ws.Cells(LineZ, 11) = "=I" & LineZ & "-J" & LineZ
                        Ws.Cells(LineZ, 12) = "=I" & LineZ & "-G" & LineZ
                        Ws.Cells(LineZ, 13) = "=I" & LineZ & "-H" & LineZ
                        Ws.Cells(LineZ, 14) = X5
                        Ws.Cells(LineZ, 15) = X6
                        Ws.Cells(LineZ, 16) = X7
                        Ws.Cells(LineZ, 17) = "=O" & LineZ & "-P" & LineZ
                        Ws.Cells(LineZ, 18) = "=O" & LineZ & "-N" & LineZ
                        Ws.Cells(LineZ, 19) = X8
                        Ws.Cells(LineZ, 20) = "=U" & LineZ & "-O" & LineZ
                        Ws.Cells(LineZ, 21) = X9
                        LineZ += 1
                    End While
                    Ws.Cells(LineZ, 6) = "总计"
                    Ws.Cells(LineZ, 7) = "=SUM(G8:G" & LineZ - 1 & ")"
                    ' 複制
                    oRng = Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 7))
                    oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 21)), Type:=xlFillDefault)
                End If
                oReader.Close()
                GG1() ' 劃線

                TP = 2

                For i As Int16 = 1 To tMonth Step 1
                    If TP + i > 3 Then
                        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                    Else
                        Ws = xWorkBook.Sheets(TP + i)
                    End If
                    Ws.Activate()
                    AdjustExcelFormat1(i)
                    oCommand.CommandText = "select abb05,gem02,aba02,abb01,abb03,aag02,abb04,(case when abb06 = 1 then abb07 else abb07 * -1 end) as t1  from abb_file left join aba_file on abb01 = aba01 "
                    oCommand.CommandText += "left join aag_file on abb03 = aag01 left join gem_file on abb05 = gem01 where abapost = 'Y' and abb03 in ('660109','660209','660422', '510121') AND abb05 <> 'D9999' "
                    oCommand.CommandText += "and aba03 = " & tYear & " and aba04 = " & i & " order by abb05"
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        While oReader.Read()
                            For j As Int16 = 0 To oReader.FieldCount - 1 Step 1
                                Ws.Cells(LineZ, j + 1) = oReader.Item(j)
                            Next
                            LineZ += 1
                        End While
                    End If
                    oReader.Close()
                Next
            Case 3
                xExcel = New Microsoft.Office.Interop.Excel.Application
                Dim xPath As String = "C:\temp\Exp-3.xlsx"
                If Not My.Computer.FileSystem.FileExists(xPath) Then
                    MsgBox("NO SAMPLE FILE")
                    Return
                End If
                xWorkBook = xExcel.Workbooks.Open(xPath)
                Ws = xWorkBook.Sheets(1)
                Ws.Activate()
                LineZ = 8
                AdjustExcelFormat()
                'oCommand.CommandText = "select distinct aao01,aag02,aao02 from aao_file left join aag_file on aao01 = aag01 where aao01 in ('510112','660110','660210','660408') and aao02 <> 'D9999'"
                oCommand.CommandText = "select distinct aao01,aag02,aao02 from ( select aao01,aag02,aao02 from aao_file left join aag_file on aao01 = aag01 where aao01 in ('510112','660110','660210','660408') and aao02 <> 'D9999' "
                oCommand.CommandText += "union all "
                oCommand.CommandText += "select tc_bud07,aag02,tc_bud08 from tc_bud_file left join aag_file on tc_bud07 = aag01 where tc_bud01 = 2 and tc_bud02 = 2019 and tc_bud07 in ('510112','660110','660210','660408') and tc_bud08 <> 'D9999' ) order by aao02,aao01"
                oReader = oCommand.ExecuteReader()
                If oReader.HasRows() Then
                    While oReader.Read()
                        X1 = Decimal.Round(GetLastYearSameMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X2 = Decimal.Round(GetLastMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X3 = Decimal.Round(GetThisYearSameMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X4 = GetThisYearSameMonthBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        X5 = Decimal.Round(GetLastYearBeforeMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X6 = Decimal.Round(GetThisYearBeforeMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X7 = GetThisYearBeforeMonthBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        X8 = Decimal.Round(GetLastYearNoMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X9 = GetThisYearBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        If X1 = 0 And X2 = 0 And X3 = 0 And X4 = 0 And X5 = 0 And X6 = 0 And X7 = 0 And X8 = 0 And X9 = 0 Then
                            Continue While
                        End If
                        Ws.Cells(LineZ, 2) = oReader.Item("aao01")
                        Ws.Cells(LineZ, 3) = oReader.Item("aag02")
                        Ws.Cells(LineZ, 4) = GetDepartNmae(oReader.Item("aao02"))
                        Ws.Cells(LineZ, 5) = GetDepartNmae("D1592")
                        Ws.Cells(LineZ, 6) = GetDepartBoss("D1592")
                        Ws.Cells(LineZ, 7) = X1
                        Ws.Cells(LineZ, 8) = X2
                        Ws.Cells(LineZ, 9) = X3
                        Ws.Cells(LineZ, 10) = X4
                        Ws.Cells(LineZ, 11) = "=I" & LineZ & "-J" & LineZ
                        Ws.Cells(LineZ, 12) = "=I" & LineZ & "-G" & LineZ
                        Ws.Cells(LineZ, 13) = "=I" & LineZ & "-H" & LineZ
                        Ws.Cells(LineZ, 14) = X5
                        Ws.Cells(LineZ, 15) = X6
                        Ws.Cells(LineZ, 16) = X7
                        Ws.Cells(LineZ, 17) = "=O" & LineZ & "-P" & LineZ
                        Ws.Cells(LineZ, 18) = "=O" & LineZ & "-N" & LineZ
                        Ws.Cells(LineZ, 19) = X8
                        Ws.Cells(LineZ, 20) = "=U" & LineZ & "-O" & LineZ
                        Ws.Cells(LineZ, 21) = X9
                        LineZ += 1
                    End While
                    Ws.Cells(LineZ, 6) = "总计"
                    Ws.Cells(LineZ, 7) = "=SUM(G8:G" & LineZ - 1 & ")"
                    ' 複制
                    oRng = Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 7))
                    oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 21)), Type:=xlFillDefault)
                End If
                oReader.Close()
                GG1() ' 劃線

                ' 第二頁
                Ws = xWorkBook.Sheets(2)
                Ws.Activate()
                LineZ = 8
                AdjustExcelFormat()
                'oCommand.CommandText = "select distinct aao01,aag02,aao02 from aao_file left join aag_file on aao01 = aag01 where aao01 in ('510123','510124','660107','660207','660420') and aao02 <> 'D9999'"
                oCommand.CommandText = "select distinct aao01,aag02,aao02 from ( select aao01,aag02,aao02 from aao_file left join aag_file on aao01 = aag01 where aao01 in ('510123','510124','660107','660207','660420') and aao02 <> 'D9999' "
                oCommand.CommandText += "union all "
                oCommand.CommandText += "select tc_bud07,aag02,tc_bud08 from tc_bud_file left join aag_file on tc_bud07 = aag01 where tc_bud01 = 2 and tc_bud02 = 2019 and tc_bud07 in ('510123','510124','660107','660207','660420') and tc_bud08 <> 'D9999' ) order by aao02,aao01"
                oReader = oCommand.ExecuteReader()
                If oReader.HasRows() Then
                    While oReader.Read()
                        X1 = Decimal.Round(GetLastYearSameMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X2 = Decimal.Round(GetLastMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X3 = Decimal.Round(GetThisYearSameMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X4 = GetThisYearSameMonthBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        X5 = Decimal.Round(GetLastYearBeforeMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X6 = Decimal.Round(GetThisYearBeforeMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X7 = GetThisYearBeforeMonthBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        X8 = Decimal.Round(GetLastYearNoMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X9 = GetThisYearBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        If X1 = 0 And X2 = 0 And X3 = 0 And X4 = 0 And X5 = 0 And X6 = 0 And X7 = 0 And X8 = 0 And X9 = 0 Then
                            Continue While
                        End If
                        Ws.Cells(LineZ, 2) = oReader.Item("aao01")
                        Ws.Cells(LineZ, 3) = oReader.Item("aag02")
                        Ws.Cells(LineZ, 4) = GetDepartNmae(oReader.Item("aao02"))
                        Ws.Cells(LineZ, 5) = GetDepartNmae("D1592")
                        Ws.Cells(LineZ, 6) = GetDepartBoss("D1592")
                        Ws.Cells(LineZ, 7) = X1
                        Ws.Cells(LineZ, 8) = X2
                        Ws.Cells(LineZ, 9) = X3
                        Ws.Cells(LineZ, 10) = X4
                        Ws.Cells(LineZ, 11) = "=I" & LineZ & "-J" & LineZ
                        Ws.Cells(LineZ, 12) = "=I" & LineZ & "-G" & LineZ
                        Ws.Cells(LineZ, 13) = "=I" & LineZ & "-H" & LineZ
                        Ws.Cells(LineZ, 14) = X5
                        Ws.Cells(LineZ, 15) = X6
                        Ws.Cells(LineZ, 16) = X7
                        Ws.Cells(LineZ, 17) = "=O" & LineZ & "-P" & LineZ
                        Ws.Cells(LineZ, 18) = "=O" & LineZ & "-N" & LineZ
                        Ws.Cells(LineZ, 19) = X8
                        Ws.Cells(LineZ, 20) = "=U" & LineZ & "-O" & LineZ
                        Ws.Cells(LineZ, 21) = X9
                        LineZ += 1
                    End While
                    Ws.Cells(LineZ, 6) = "总计"
                    Ws.Cells(LineZ, 7) = "=SUM(G8:G" & LineZ - 1 & ")"
                    ' 複制
                    oRng = Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 7))
                    oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 21)), Type:=xlFillDefault)
                End If
                oReader.Close()
                GG1() ' 劃線

                TP = 2

                For i As Int16 = 1 To tMonth Step 1
                    If TP + i > 3 Then
                        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                    Else
                        Ws = xWorkBook.Sheets(TP + i)
                    End If
                    Ws.Activate()
                    AdjustExcelFormat1(i)
                    oCommand.CommandText = "select abb05,gem02,aba02,abb01,abb03,aag02,abb04,(case when abb06 = 1 then abb07 else abb07 * -1 end) as t1  from abb_file left join aba_file on abb01 = aba01 "
                    oCommand.CommandText += "left join aag_file on abb03 = aag01 left join gem_file on abb05 = gem01 where abapost = 'Y' and abb03 in ('510112','660110','660210','660408','510123','510124','660107','660207','660420') AND abb05 <> 'D9999' "
                    oCommand.CommandText += "and aba03 = " & tYear & " and aba04 = " & i & " order by abb05"
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        While oReader.Read()
                            For j As Int16 = 0 To oReader.FieldCount - 1 Step 1
                                Ws.Cells(LineZ, j + 1) = oReader.Item(j)
                            Next
                            LineZ += 1
                        End While
                    End If
                    oReader.Close()
                Next
            Case 4
                xExcel = New Microsoft.Office.Interop.Excel.Application
                Dim xPath As String = "C:\temp\Exp-4.xlsx"
                If Not My.Computer.FileSystem.FileExists(xPath) Then
                    MsgBox("NO SAMPLE FILE")
                    Return
                End If
                xWorkBook = xExcel.Workbooks.Open(xPath)
                Ws = xWorkBook.Sheets(1)
                Ws.Activate()
                LineZ = 8
                AdjustExcelFormat()
                'oCommand.CommandText = "select distinct aao01,aag02,aao02 from aao_file left join aag_file on aao01 = aag01 where aao01 in ('660406','510109') and aao02 <> 'D9999'"
                oCommand.CommandText = "select distinct aao01,aag02,aao02 from ( select aao01,aag02,aao02 from aao_file left join aag_file on aao01 = aag01 where aao01 in ('660406','510109') and aao02 <> 'D9999' "
                oCommand.CommandText += "union all "
                oCommand.CommandText += "select tc_bud07,aag02,tc_bud08 from tc_bud_file left join aag_file on tc_bud07 = aag01 where tc_bud01 = 2 and tc_bud02 = 2019 and tc_bud07 in ('660406','510109') and tc_bud08 <> 'D9999' ) order by aao02,aao01"
                oReader = oCommand.ExecuteReader()
                If oReader.HasRows() Then
                    While oReader.Read()
                        X1 = Decimal.Round(GetLastYearSameMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X2 = Decimal.Round(GetLastMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X3 = Decimal.Round(GetThisYearSameMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X4 = GetThisYearSameMonthBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        X5 = Decimal.Round(GetLastYearBeforeMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X6 = Decimal.Round(GetThisYearBeforeMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X7 = GetThisYearBeforeMonthBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        X8 = Decimal.Round(GetLastYearNoMonth(oReader.Item("aao01").ToString(), oReader.Item("aao02")), 0)
                        X9 = GetThisYearBudget(oReader.Item("aao01").ToString(), oReader.Item("aao02"))
                        If X1 = 0 And X2 = 0 And X3 = 0 And X4 = 0 And X5 = 0 And X6 = 0 And X7 = 0 And X8 = 0 And X9 = 0 Then
                            Continue While
                        End If
                        Ws.Cells(LineZ, 2) = oReader.Item("aao01")
                        Ws.Cells(LineZ, 3) = oReader.Item("aag02")
                        Ws.Cells(LineZ, 4) = GetDepartNmae(oReader.Item("aao02"))
                        Ws.Cells(LineZ, 5) = GetDepartNmae("D3100")
                        Ws.Cells(LineZ, 6) = GetDepartBoss("D3100")
                        Ws.Cells(LineZ, 7) = X1
                        Ws.Cells(LineZ, 8) = X2
                        Ws.Cells(LineZ, 9) = X3
                        Ws.Cells(LineZ, 10) = X4
                        Ws.Cells(LineZ, 11) = "=I" & LineZ & "-J" & LineZ
                        Ws.Cells(LineZ, 12) = "=I" & LineZ & "-G" & LineZ
                        Ws.Cells(LineZ, 13) = "=I" & LineZ & "-H" & LineZ
                        Ws.Cells(LineZ, 14) = X5
                        Ws.Cells(LineZ, 15) = X6
                        Ws.Cells(LineZ, 16) = X7
                        Ws.Cells(LineZ, 17) = "=O" & LineZ & "-P" & LineZ
                        Ws.Cells(LineZ, 18) = "=O" & LineZ & "-N" & LineZ
                        Ws.Cells(LineZ, 19) = X8
                        Ws.Cells(LineZ, 20) = "=U" & LineZ & "-O" & LineZ
                        Ws.Cells(LineZ, 21) = X9
                        LineZ += 1
                    End While
                    Ws.Cells(LineZ, 6) = "总计"
                    Ws.Cells(LineZ, 7) = "=SUM(G8:G" & LineZ - 1 & ")"
                    ' 複制
                    oRng = Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 7))
                    oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 21)), Type:=xlFillDefault)
                End If
                oReader.Close()
                GG1() ' 劃線

                TP = 1

                For i As Int16 = 1 To tMonth Step 1
                    If TP + i > 3 Then
                        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                    Else
                        Ws = xWorkBook.Sheets(TP + i)
                    End If
                    Ws.Activate()
                    AdjustExcelFormat1(i)
                    oCommand.CommandText = "select abb05,gem02,aba02,abb01,abb03,aag02,abb04,(case when abb06 = 1 then abb07 else abb07 * -1 end) as t1  from abb_file left join aba_file on abb01 = aba01 "
                    oCommand.CommandText += "left join aag_file on abb03 = aag01 left join gem_file on abb05 = gem01 where abapost = 'Y' and abb03 in ('660406','510109') AND abb05 <> 'D9999' "
                    oCommand.CommandText += "and aba03 = " & tYear & " and aba04 = " & i & " order by abb05"
                    oReader = oCommand.ExecuteReader()
                    If oReader.HasRows() Then
                        While oReader.Read()
                            For j As Int16 = 0 To oReader.FieldCount - 1 Step 1
                                Ws.Cells(LineZ, j + 1) = oReader.Item(j)
                            Next
                            LineZ += 1
                        End While
                    End If
                    oReader.Close()
                Next
        End Select



    End Sub
    Private Sub AdjustExcelFormat()
        Ws.Cells(7, 7) = pYear & "/" & tMonth & "/01"
        Ws.Cells(7, 8) = lYear & "/" & lMonth & "/01"
        Ws.Cells(7, 9) = tYear & "/" & tMonth & "/01"
        Ws.Cells(7, 10) = tYear & "/" & tMonth & "/01"
        Ws.Cells(7, 14) = "YTD " & pYear
        Ws.Cells(7, 15) = "YTD " & tYear
        Ws.Cells(7, 16) = "YTD " & tYear
        Ws.Cells(7, 19) = "Y" & pYear
        Ws.Cells(7, 20) = "Y" & tYear
        Ws.Cells(7, 21) = "Y" & tYear
    End Sub
    Private Function GetDepartNmae(ByVal gem01 As String)
        oCommand2.CommandText = "select gem02 from gem_file where gem01 = '" & gem01 & "'"
        Dim DN As String = oCommand2.ExecuteScalar()
        Return DN
    End Function
    Private Function GetDepartBoss(ByVal gem01 As String)
        oCommand2.CommandText = "select gem06 from gem_file where gem01 = '" & gem01 & "'"
        Dim DN As String = oCommand2.ExecuteScalar()
        Return DN
    End Function
    Private Function GetLastYearSameMonth(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
        oCommand2.CommandText += pYear & " and aao04 = " & tMonth
        Dim LYTM As Decimal = oCommand2.ExecuteScalar()
        Return LYTM
    End Function
    Private Function GetLastMonth(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
        oCommand2.CommandText += lYear & " and aao04 = " & lMonth
        Dim TYTM As Decimal = oCommand2.ExecuteScalar()
        Return TYTM
    End Function
    Private Function GetThisYearSameMonth(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
        oCommand2.CommandText += tYear & " and aao04 = " & tMonth
        Dim TYTM As Decimal = oCommand2.ExecuteScalar()
        Return TYTM
    End Function
    Private Function GetThisYearSameMonthBudget(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(tc_bud13),0) from tc_bud_file where tc_bud07 = '" & aag01 & "' and tc_bud08 = '" & gem01 & "' and tc_bud02 = "
        oCommand2.CommandText += tYear & " and tc_bud03 = " & tMonth
        Dim TYTMB As Decimal = Decimal.Round(oCommand2.ExecuteScalar(), 0)
        Return TYTMB
    End Function
    Private Function GetLastYearBeforeMonth(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
        oCommand2.CommandText += pYear & " and aao04 <= " & tMonth & " and aao04 > 0"
        Dim LYBM As Decimal = oCommand2.ExecuteScalar()
        Return LYBM
    End Function
    Private Function GetThisYearBeforeMonth(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
        oCommand2.CommandText += tYear & " and aao04 <= " & tMonth & " and aao04 > 0"
        Dim TYBM As Decimal = oCommand2.ExecuteScalar()
        Return TYBM
    End Function
    Private Function GetThisYearBeforeMonthBudget(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(tc_bud13),0) from tc_bud_file where tc_bud07 = '" & aag01 & "' and tc_bud08 = '" & gem01 & "' and tc_bud02 = "
        oCommand2.CommandText += tYear & " and tc_bud03 <= " & tMonth
        Dim TYBMB As Decimal = Decimal.Round(oCommand2.ExecuteScalar(), 0)
        Return TYBMB
    End Function
    Private Function GetLastYearNoMonth(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(aao05-aao06),0) from aao_file where aao01 = '" & aag01 & "' and aao02 = '" & gem01 & "' and aao03 = "
        oCommand2.CommandText += pYear.ToString() & " and aao04 > 0"
        Dim TYNM As Decimal = oCommand2.ExecuteScalar()
        Return TYNM
    End Function
    Private Function GetThisYearBudget(ByVal aag01 As String, ByVal gem01 As String)
        oCommand2.CommandText = "select nvl(sum(tc_bud13),0) from tc_bud_file where tc_bud07 = '" & aag01 & "' and tc_bud08 = '" & gem01 & "' and tc_bud02 = "
        oCommand2.CommandText += tYear.ToString()
        Dim TYTMB As Decimal = Decimal.Round(oCommand2.ExecuteScalar(), 0)
        Return TYTMB
    End Function
    Private Sub GG1()
        oRng = Ws.Range("B8", Ws.Cells(LineZ, 21))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
    End Sub
    Private Sub AdjustExcelFormat1(ByVal PN As String)
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = PN
        oRng = Ws.Range("A1", "H1")
        oRng.EntireColumn.ColumnWidth = 30
        oRng = Ws.Range("E1", "E1")
        oRng.EntireColumn.NumberFormat = "@"
        Ws.Cells(1, 1) = "部门编号"
        Ws.Cells(1, 2) = "部门名称"
        Ws.Cells(1, 3) = "凭证日期"
        Ws.Cells(1, 4) = "凭证编号"
        Ws.Cells(1, 5) = "科目编码"
        Ws.Cells(1, 6) = "科目名称"
        Ws.Cells(1, 7) = "摘要"
        Ws.Cells(1, 8) = "发生额"
        LineZ = 2
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Expense_MasterDepartment"
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
End Class