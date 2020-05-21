Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form123
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
    Dim xPath As String = String.Empty
    Dim gDatabase As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form123_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        tCurrency = Me.ComboBox1.SelectedItem.ToString()
        If String.IsNullOrEmpty(tCurrency) Then
            MsgBox("Currency Error")
            Return
        End If
        'Dim xPath As String = "C:\temp\IS - DAC（RMB）.xlsx"
        gDatabase = Me.ComboBox2.SelectedItem.ToString()
        Select Case gDatabase
            Case "DAC"
                oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
                If tCurrency = "USD" Then
                    xPath = "C:\Temp\IS-DAC（USD）.xlsx"
                Else
                    xPath = "C:\Temp\IS-DAC（RMB）.xlsx"
                End If
            Case "HAC"
                oConnection.ConnectionString = Module1.OpenOracleConnection("hkacttest")
                xPath = "C:\Temp\IS-HAC（USD）.xlsx"
            Case "BVI"
                oConnection.ConnectionString = Module1.OpenOracleConnection("action_bvi")
                xPath = "C:\Temp\IS-BVI（USD）.xlsx"
        End Select
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        
        If TextBox1.Text.Length < 6 Then
            MsgBox("ERROR")
            Return
        End If
        'gDatabase = Me.ComboBox2.SelectedItem.ToString()
        If String.IsNullOrEmpty(gDatabase) Then
            MsgBox("Database Error")
            Return
        End If
        'Select Case gDatabase
        '    Case "DAC"
        '        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        '    Case "HAC"
        '        oConnection.ConnectionString = Module1.OpenOracleConnection("hkacttest")
        '    Case "BVI"
        '        oConnection.ConnectionString = Module1.OpenOracleConnection("action_bvi")
        'End Select
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
        
        ' 確認 ExchangeRate
        If tCurrency = "USD" And gDatabase = "DAC" Then
            Dim CS As String = String.Empty
            If tMonth < 10 Then
                CS = tYear & "0" & tMonth
            Else
                CS = tYear & tMonth
            End If
            oCommand.CommandText = "SELECT nvl(AZJ041,0) FROM AZJ_FILE WHERE AZJ01  = 'USD' AND AZJ02 = '" & CS & "'"
            ExchangeRate = oCommand.ExecuteScalar()
            If ExchangeRate = 0 Then
                ExchangeRate = 1
            End If
            ExchangeRate1 = 6.85
        Else
            ExchangeRate = 1
            ExchangeRate1 = 1
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
        'Dim xPath As String = "C:\temp\IS - DAC（RMB）.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        LineZ = 7
        AdjustExcelFormat()
        DoInputData("600101", "600102", 1)
        LineZ += 1
        DoInputData("6051", "6099", 1)
        LineZ += 9
        DoInputData("640101", "6403", 0)
        LineZ += 4
        DoInputData("660101", "660199", 0)
        LineZ += 1
        DoInputData("660201", "660299", 0)
        LineZ += 1
        If gDatabase = "HAC" Then
            DoInputData("660301", "660311", 0)
        Else
            DoInputData("660301", "660303", 0)
        End If

        LineZ += 1
        DoInputData("660401", "660499", 0)
        LineZ += 1
        DoInputData("6701", "6701", 0)
        LineZ += 3
        DoInputData("6301", "6301", 1)
        LineZ += 2
        DoInputData("6711", "6711", 0)
        LineZ += 4
        DoInputData("6801", "6801", 0)

        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        LineZ = 7
        AdjustExcelFormat2()
        DoInputData1("600101", "600102", 1)
        LineZ += 1
        DoInputData1("6051", "6099", 1)
        LineZ += 9
        DoInputData1("640101", "6403", 0)
        LineZ += 4
        DoInputData1("660101", "660199", 0)
        LineZ += 1
        DoInputData1("660201", "660299", 0)
        LineZ += 1
        If gDatabase = "HAC" Then
            DoInputData1("660301", "660311", 0)
        Else
            DoInputData1("660301", "660303", 0)
        End If

        LineZ += 1
        DoInputData1("660401", "660499", 0)
        LineZ += 1
        DoInputData1("6701", "6701", 0)
        LineZ += 3
        DoInputData1("6301", "6301", 1)
        LineZ += 2
        DoInputData1("6711", "6711", 0)
        LineZ += 4
        DoInputData1("6801", "6801", 0)

        'Ws = xWorkBook.Sheets(3)
        'Ws.Activate()
        'LineZ = 7
        'AdjustExcelFormat3()
        ''DoInputData2("600101", "600102", 1)
        'DoInputData3()
        'LineZ += 1
        'DoInputData2("6051", "6099", 1)
        'LineZ += 9
        'DoInputData2("640101", "6403", 0)
        'LineZ += 4
        ''DoInputData2("660101", "660199", 0)
        'DoInputData2("660101", "66013102", 0)
        'LineZ += 1
        'DoInputData2("660201", "660299", 0)
        'LineZ += 1
        'If gDatabase = "HAC" Then
        '    DoInputData2("660301", "660311", 0)
        'Else
        '    DoInputData2("660301", "660303", 0)
        'End If

        'LineZ += 1
        ''DoInputData2("660401", "660499", 0)
        'DoInputData2("660401", "660420", 0)
        'LineZ += 4
        'DoInputData2("6301", "6301", 1)
        'LineZ += 2
        'DoInputData2("6711", "6711", 0)
        'LineZ += 4
        'DoInputData2("6801", "6801", 0)
    End Sub
    Private Sub AdjustExcelFormat()
        Select Case gDatabase
            Case "DAC"
                Ws.Cells(2, 1) = "Company Name：Dongguan Action Composite LTD. Co"
            Case "HAC"
                Ws.Cells(2, 1) = "Company Name：ACTION COMPOSITE TECHNOLOGY LIMITED"
            Case "BVI"
                Ws.Cells(2, 1) = "Company Name：ACTION COMPOSITES INTERNATIONAL LIMITED"
        End Select
        For i As Int16 = 1 To 12
            If i < 10 Then
                Ws.Cells(6, 2 * i + 1) = tYear & "/0" & i
            Else
                Ws.Cells(6, 2 * i + 1) = tYear & "/" & i
            End If
        Next
        Ws.Cells(3, 2) = tCurrency
    End Sub
    Private Sub AdjustExcelFormat2()
        Select Case gDatabase
            Case "DAC"
                Ws.Cells(2, 1) = "Company Name：Dongguan Action Composite LTD. Co"
            Case "HAC"
                Ws.Cells(2, 1) = "Company Name：ACTION COMPOSITE TECHNOLOGY LIMITED"
            Case "BVI"
                Ws.Cells(2, 1) = "Company Name：ACTION COMPOSITES INTERNATIONAL LIMITED"
        End Select
        If tMonth < 10 Then
            Ws.Cells(6, 3) = pYear & "/0" & tMonth
            Ws.Cells(6, 7) = tYear & "/0" & tMonth
            Ws.Cells(6, 9) = tYear & "/0" & tMonth
        Else
            Ws.Cells(6, 3) = pYear & "/" & tMonth
            Ws.Cells(6, 7) = tYear & "/" & tMonth
            Ws.Cells(6, 9) = tYear & "/0" & tMonth
        End If
        If pMonth < 10 Then
            Ws.Cells(6, 5) = lYear & "/0" & pMonth
        Else
            Ws.Cells(6, 5) = lYear & "/" & pMonth
        End If
        Ws.Cells(3, 2) = tCurrency
        Ws.Cells(6, 11) = tCurrency
        Ws.Cells(6, 12) = "YTD " & pYear
        Ws.Cells(6, 14) = "YTD " & tYear
        Ws.Cells(6, 16) = "YTD " & tYear
        Ws.Cells(6, 18) = tCurrency
        Ws.Cells(6, 19) = "Y" & pYear
        Ws.Cells(6, 20) = "Y" & tYear
        Ws.Cells(6, 21) = "Y" & tYear
    End Sub
    Private Sub AdjustExcelFormat3()
        Select Case gDatabase
            Case "DAC"
                Ws.Cells(2, 1) = "Company Name：Dongguan Action Composite LTD. Co"
            Case "HAC"
                Ws.Cells(2, 1) = "Company Name：ACTION COMPOSITE TECHNOLOGY LIMITED"
            Case "BVI"
                Ws.Cells(2, 1) = "Company Name：ACTION COMPOSITES INTERNATIONAL LIMITED"
        End Select
        For i As Int16 = 1 To 12
            If i < 10 Then
                Ws.Cells(6, 2 * i + 1) = tYear & "/0" & i
            Else
                Ws.Cells(6, 2 * i + 1) = tYear & "/" & i
            End If
        Next
        Ws.Cells(3, 2) = tCurrency
    End Sub
    Private Sub DoInputData(ByVal ACC1 As String, ByVal ACC2 As String, ByVal ACC3 As Int16)
        If tCurrency = "USD" And gDatabase = "DAC" Then
            oCommand.CommandText = "select "
            For i As Int16 = 1 To tMonth Step 1
                oCommand.CommandText += "nvl(sum(t" & i & "),0) as t" & i & ","
            Next
            oCommand.CommandText += "1 from ( select "
            For i As Int16 = 1 To tMonth Step 1
                oCommand.CommandText += "(case when aah03 = " & i & " then round((aah05 - aah04)/azj041,3) else 0 end ) as t" & i & ","
            Next
            oCommand.CommandText += "1 from aah_file,aag_file,azj_file where aah01 = aag01 and aag07 in ('2','3') and aah02 = " & tYear & " and aah03 <= " & tMonth & " and aah01 between '" & ACC1 & "' and '" & ACC2 & "' and azj01 = 'USD' and azj02 = aah02 || (case when length(aah03) < 2 then 0 end) || aah03 ) "
        Else
            oCommand.CommandText = "select "
            For i As Int16 = 1 To tMonth Step 1
                oCommand.CommandText += "nvl(sum(t" & i & "),0) as t" & i & ","
            Next
            oCommand.CommandText += "1 from ( select "
            For i As Int16 = 1 To tMonth Step 1
                oCommand.CommandText += "(case when aah03 = " & i & " then (aah05 - aah04) else 0 end ) as t" & i & ","
            Next
            oCommand.CommandText += "1 from aah_file,aag_file where aah01 = aag01 and aag07 in ('2','3') and aah02 = " & tYear & " and aah03 <= " & tMonth & " and aah01 between '" & ACC1 & "' and '" & ACC2 & "' ) "
        End If
        
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                For i As Int16 = 1 To tMonth Step 1
                    If ACC3 = 0 Then
                        Ws.Cells(LineZ, 2 * i + 1) = (oReader.Item(i - 1) * Decimal.MinusOne)
                    Else
                        Ws.Cells(LineZ, 2 * i + 1) = oReader.Item(i - 1)
                    End If

                Next
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Income_Statement" & tYear & "_" & gDatabase
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
    Private Sub DoInputData1(ByVal ACC1 As String, ByVal ACC2 As String, ByVal ACC3 As Int16)
        If tCurrency = "USD" And gDatabase = "DAC" Then
            oCommand.CommandText = "select nvl(sum(round((aah05-aah04)/azj041,3)),0) from aah_file,aag_file,azj_file where aah01 = aag01 and aag07 in ('2','3') and aah02 = " & pYear & " and aah03 = " & tMonth & " and aah01 between '" & ACC1 & "' and '" & ACC2 & "' and azj01 = 'USD' and azj02 = aah02 || (case when length(aah03) < 2 then 0 end) || aah03 "
        Else
            oCommand.CommandText = "select nvl(sum(aah05-aah04),0) from aah_file,aag_file where aah01 = aag01 and aag07 in ('2','3') and aah02 = " & pYear & " and aah03 = " & tMonth & " and aah01 between '" & ACC1 & "' and '" & ACC2 & "' "
        End If
        Dim BC1 As Decimal = oCommand.ExecuteScalar()
        If ACC3 = 0 Then
            Ws.Cells(LineZ, 3) = (BC1 * Decimal.MinusOne)
        Else
            Ws.Cells(LineZ, 3) = BC1
        End If
        If tCurrency = "USD" And gDatabase = "DAC" Then
            oCommand.CommandText = "select nvl(sum(round((aah05-aah04)/azj041,3)),0) from aah_file,aag_file,azj_file where aah01 = aag01 and aag07 in ('2','3') and aah02 = " & lYear & " and aah03 = " & pMonth & " and aah01 between '" & ACC1 & "' and '" & ACC2 & "' and azj01 = 'USD' and azj02 = aah02 || (case when length(aah03) < 2 then 0 end) || aah03 "
        Else
            oCommand.CommandText = "select nvl(sum(aah05-aah04),0) from aah_file,aag_file where aah01 = aag01 and aag07 in ('2','3') and aah02 = " & lYear & " and aah03 = " & pMonth & " and aah01 between '" & ACC1 & "' and '" & ACC2 & "' "
        End If

        Dim BC2 As Decimal = oCommand.ExecuteScalar()
        If ACC3 = 0 Then
            Ws.Cells(LineZ, 5) = (BC2 * Decimal.MinusOne)
        Else
            Ws.Cells(LineZ, 5) = BC2
        End If
        If tCurrency = "USD" And gDatabase = "DAC" Then
            oCommand.CommandText = "select nvl(sum(round((aah05-aah04)/azj041,3)),0) from aah_file,aag_file,azj_file where aah01 = aag01 and aag07 in ('2','3') and aah02 = " & tYear & " and aah03 = " & tMonth & " and aah01 between '" & ACC1 & "' and '" & ACC2 & "' and azj01 = 'USD' and azj02 = aah02 || (case when length(aah03) < 2 then 0 end) || aah03 "
        Else
            oCommand.CommandText = "select nvl(sum(aah05-aah04),0) from aah_file,aag_file where aah01 = aag01 and aag07 in ('2','3') and aah02 = " & tYear & " and aah03 = " & tMonth & " and aah01 between '" & ACC1 & "' and '" & ACC2 & "' "
        End If

        Dim BC3 As Decimal = oCommand.ExecuteScalar()
        If ACC3 = 0 Then
            Ws.Cells(LineZ, 7) = (BC3 * Decimal.MinusOne)
        Else
            Ws.Cells(LineZ, 7) = BC3
        End If
        'oCommand.CommandText = "select nvl(sum(tc_bud13),0)  from tc_bud_file where tc_bud01 = 2 and tc_bud02 = " & tYear & " and tc_bud03 = " & tMonth & " and tc_bud07 between '" & ACC1 & "' and '" & ACC2 & "' "
        'Dim BC4 As Decimal = oCommand.ExecuteScalar()
        'If ACC3 = 0 Then
        'Ws.Cells(LineZ, 9) = (BC4 * Decimal.MinusOne) / ExchangeRate
        'Else
        'Ws.Cells(LineZ, 9) = BC4 / ExchangeRate1
        'End If
        If tCurrency = "USD" And gDatabase = "DAC" Then
            oCommand.CommandText = "select nvl(sum(round((aah05-aah04)/azj041,3)),0) from aah_file,aag_file,azj_file where aah01 = aag01 and aag07 in ('2','3') and aah02 = " & pYear & " and aah03 <= " & tMonth & " and aah01 between '" & ACC1 & "' and '" & ACC2 & "' and azj01 = 'USD' and azj02 = aah02 || (case when length(aah03) < 2 then 0 end) || aah03 "
        Else
            oCommand.CommandText = "select nvl(sum(aah05-aah04),0) from aah_file,aag_file where aah01 = aag01 and aag07 in ('2','3') and aah02 = " & pYear & " and aah03 <= " & tMonth & " and aah01 between '" & ACC1 & "' and '" & ACC2 & "' "
        End If

        Dim BC5 As Decimal = oCommand.ExecuteScalar()
        If ACC3 = 0 Then
            Ws.Cells(LineZ, 12) = (BC5 * Decimal.MinusOne)
        Else
            Ws.Cells(LineZ, 12) = BC5
        End If
        If tCurrency = "USD" And gDatabase = "DAC" Then
            oCommand.CommandText = "select nvl(sum(round((aah05-aah04)/azj041,3)),0) from aah_file,aag_file,azj_file where aah01 = aag01 and aag07 in ('2','3') and aah02 = " & tYear & " and aah03 <= " & tMonth & " and aah01 between '" & ACC1 & "' and '" & ACC2 & "' and azj01 = 'USD' and azj02 = aah02 || (case when length(aah03) < 2 then 0 end) || aah03 "
        Else
            oCommand.CommandText = "select nvl(sum(aah05-aah04),0) from aah_file,aag_file where aah01 = aag01 and aag07 in ('2','3') and aah02 = " & tYear & " and aah03 <= " & tMonth & " and aah01 between '" & ACC1 & "' and '" & ACC2 & "' "
        End If

        Dim BC6 As Decimal = oCommand.ExecuteScalar()
        If ACC3 = 0 Then
            Ws.Cells(LineZ, 14) = (BC6 * Decimal.MinusOne)
        Else
            Ws.Cells(LineZ, 14) = BC6
        End If
        'oCommand.CommandText = "select nvl(sum(tc_bud13),0)  from tc_bud_file where tc_bud01 = 2 and tc_bud02 = " & tYear & " and tc_bud03 <= " & tMonth & " and tc_bud07 between '" & ACC1 & "' and '" & ACC2 & "' "
        'Dim BC7 As Decimal = oCommand.ExecuteScalar()
        'If ACC3 = 0 Then
        'Ws.Cells(LineZ, 16) = (BC7 * Decimal.MinusOne) / ExchangeRate
        'Else
        'Ws.Cells(LineZ, 16) = BC7 / ExchangeRate1
        'End If
        If tCurrency = "USD" And gDatabase = "DAC" Then
            oCommand.CommandText = "select nvl(sum(round((aah05-aah04)/azj041,3)),0) from aah_file,aag_file,AZJ_FILE where aah01 = aag01 and aag07 in ('2','3') and aah02 = " & pYear & " and aah01 between '" & ACC1 & "' and '" & ACC2 & "' and azj01 = 'USD' and azj02 = aah02 || (case when length(aah03) < 2 then 0 end) || aah03 "
        Else
            oCommand.CommandText = "select nvl(sum(aah05-aah04),0) from aah_file,aag_file where aah01 = aag01 and aag07 in ('2','3') and aah02 = " & pYear & " and aah01 between '" & ACC1 & "' and '" & ACC2 & "' "
        End If

        Dim BC8 As Decimal = oCommand.ExecuteScalar()
        If ACC3 = 0 Then
            Ws.Cells(LineZ, 19) = (BC8 * Decimal.MinusOne)
        Else
            Ws.Cells(LineZ, 19) = BC8
        End If
        'oCommand.CommandText = "select nvl(sum(tc_bud13),0)  from tc_bud_file where tc_bud01 = 2 and tc_bud02 = " & tYear & " and tc_bud07 between '" & ACC1 & "' and '" & ACC2 & "' "
        'Dim BC9 As Decimal = oCommand.ExecuteScalar()
        'If ACC3 = 0 Then
        'Ws.Cells(LineZ, 21) = (BC7 * Decimal.MinusOne) / ExchangeRate
        'Else
        'Ws.Cells(LineZ, 21) = BC9 / ExchangeRate1
        'End If
    End Sub
    Private Sub DoInputData2(ByVal ACC1 As String, ByVal ACC2 As String, ByVal ACC3 As Int16)

        oCommand.CommandText = "select "
        For i As Int16 = 1 To 12 Step 1
            oCommand.CommandText += "nvl(sum(t" & i & "),0) as t" & i & ","
        Next
        oCommand.CommandText += "1 from ( select "
        For i As Int16 = 1 To 12 Step 1
            oCommand.CommandText += "(case when tc_bud03 = " & i & " then tc_bud13 else 0 end ) as t" & i & ","
        Next
        oCommand.CommandText += "1 from tc_bud_file where tc_bud01 = 2 and tc_bud02 = " & tYear & " and tc_bud07 between '" & ACC1 & "' and '" & ACC2 & "' ) "

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                For i As Int16 = 1 To 12 Step 1
                    Ws.Cells(LineZ, 2 * i + 1) = oReader.Item(i - 1) / ExchangeRate1
                Next
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub DoInputData3()
        If tCurrency = "USD" Then
            oCommand.CommandText = "select "
            For i As Int16 = 1 To 12 Step 1
                oCommand.CommandText += "nvl(sum(t" & i & "),0) as t" & i & ","
            Next
            oCommand.CommandText += "1 from ( select "
            For i As Int16 = 1 To 12 Step 1
                oCommand.CommandText += "(case when tc_bud03 = " & i & " then (case when tc_bud14 = 'EUR' THEN round(tc_bud13 * 1.2,3) else tc_bud13 end) else 0 end ) as t" & i & ","
            Next
            oCommand.CommandText += "1 from tc_bud_file where tc_bud01 = 1 and tc_bud02 = " & tYear & " ) "
        Else
            oCommand.CommandText = "select "
            For i As Int16 = 1 To 12 Step 1
                oCommand.CommandText += "nvl(sum(t" & i & "),0) as t" & i & ","
            Next
            oCommand.CommandText += "1 from ( select "
            For i As Int16 = 1 To 12 Step 1
                oCommand.CommandText += "(case when tc_bud03 = " & i & " then (case when tc_bud14 = 'USD' then tc_bud13 * 6.85 else tc_bud13 * 8.22 end) else 0 end ) as t" & i & ","
            Next
            oCommand.CommandText += "1 from tc_bud_file where tc_bud01 = 1 and tc_bud02 = " & tYear & " and tc_bud07 not like '5101%' ) "
        End If
        
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                For i As Int16 = 1 To 12 Step 1
                    Ws.Cells(LineZ, 2 * i + 1) = oReader.Item(i - 1)
                Next
            End While
        End If
        oReader.Close()
    End Sub
End Class