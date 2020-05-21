Public Class Form155
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
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
    Dim gDatabase As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form155_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        If Today.Month < 10 Then
            TextBox1.Text = Today.Year & "0" & Today.Month
        Else
            TextBox1.Text = Today.Year & Today.Month
        End If
        mConnection.ConnectionString = Module1.OpenConnectionOfERPSUPPORT()
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mSQLS1.CommandTimeout = 600
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
                mSQLS2.CommandTimeout = 600
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Dim xPath As String = "C:\temp\ACAIS.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If

        If TextBox1.Text.Length < 6 Then
            MsgBox("ERROR")
            Return
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
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        Dim xPath As String = "C:\temp\ACAIS.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        LineZ = 7
        AdjustExcelFormat()
        'DoInputData("4005", "4061")
        DoInputData2("'4005','4030', '4040', '4061', '4020', '4041', '4042'", True)
        LineZ += 1
        DoInputData2("'4011','4012','4013','4014','4015','4016'", True)
        LineZ += 1
        DoInputData2("'4025', '4038', '4831', '4838', '4830', '4832', '4834', '4899', '4860', '4920','4833', '4836', '4839', '4810', '4910'", True)
        LineZ += 1
        DoInputData2("'4430','4405', '4440'", True)
        LineZ += 3
        DoInputData("Purc", "Purc", False)
        LineZ += 1
        DoInputData("B10", "B10", False)
        LineZ += 1
        DoInputData("B12", "B12", False)
        LineZ += 2
        DoInputData("Purc", "Purc", False)
        LineZ += 5
        DoInputData("Expe", "Expe", False)
        LineZ += 1
        DoInputData("fina", "fina", False)
        LineZ += 11
        DoInputData("8500", "8500", False)

        ' 第二頁新增 20190223
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        LineZ = 7
        AdjustExcelFormat1()
        DoInputData2("'4005','4030', '4040', '4061', '4020', '4041', '4042'", True, True)
        LineZ += 1
        DoInputData2("'4011','4012','4013','4014','4015','4016'", True, True)
        LineZ += 1
        DoInputData2("'4025', '4038', '4831', '4838', '4830', '4832', '4834', '4899', '4860', '4920','4833', '4836', '4839', '4810', '4910'", True, True)
        LineZ += 1
        DoInputData2("'4430','4405', '4440'", True, True)
        LineZ += 3
        DoInputData("Purc", "Purc", False, True)
        LineZ += 1
        DoInputData("B10", "B10", False, True)
        LineZ += 1
        DoInputData("B12", "B12", False, True)
        LineZ += 2
        DoInputData("Purc", "Purc", False, True)
        LineZ += 5
        DoInputData("Expe", "Expe", False, True)
        LineZ += 1
        DoInputData("fina", "fina", False, True)
        LineZ += 11
        DoInputData("8500", "8500", False, True)
    End Sub
    Private Sub AdjustExcelFormat()
        For i As Int16 = 1 To 12
            If i < 10 Then
                Ws.Cells(6, 2 * i + 1) = tYear & "/0" & i
            Else
                Ws.Cells(6, 2 * i + 1) = tYear & "/" & i
            End If
        Next
    End Sub
    Private Sub AdjustExcelFormat1()
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        For i As Int16 = 1 To 12
            If i < 10 Then
                Ws.Cells(6, 2 * i + 1) = tYear & "/0" & i
            Else
                Ws.Cells(6, 2 * i + 1) = tYear & "/" & i
            End If
        Next
        ' 加入匯率
        For i As Int16 = 1 To tMonth Step 1
            Dim AzjYM As String = String.Empty
            If i < 10 Then
                AzjYM = tYear & "0" & i
            Else
                AzjYM = tYear & i
            End If
            'oCommand.CommandText = "select nvl(azj07,1) from azj_file where  azj02 = '" & AzjYM & "' and azj01 = 'USD'"
            'Dim USDA As Decimal = oCommand.ExecuteScalar()
            oCommand.CommandText = "select nvl(azj03,1) from hkacttest.azj_file where  azj02 = '" & AzjYM & "' and azj01 = 'EUR'"
            Dim EURA As Decimal = oCommand.ExecuteScalar()
            'Dim ExchangeRateA As Decimal = Decimal.Round(EURA / USDA, 6)
            Ws.Cells(3, 2 * i + 1) = EURA
        Next
        oConnection.Close()

    End Sub
    Private Sub DoInputData(ByVal ACC1 As String, ByVal ACC2 As String, ByVal Keepit As Boolean)

        mSQLS1.CommandText = "select "
        For i As Int16 = 1 To tMonth Step 1
            mSQLS1.CommandText += "isnull(sum(t" & i & "),0) as t" & i & ","
        Next
        mSQLS1.CommandText += "1 from ( select "
        For i As Int16 = 1 To tMonth Step 1
            mSQLS1.CommandText += "(case when month1 = " & i & " then Amount1 else 0 end ) as t" & i & ","
        Next
        mSQLS1.CommandText += "1 AS NN from acais where year1 = " & tYear & " and month1 <= " & tMonth & " and Acc1 between '" & ACC1 & "' and '" & ACC2 & "' ) AS AB "


        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows Then
            While mSQLReader.Read()
                For i As Int16 = 1 To tMonth Step 1
                    If Keepit = True Then
                        Ws.Cells(LineZ, 2 * i + 1) = mSQLReader.Item(i - 1)
                    Else
                        Ws.Cells(LineZ, 2 * i + 1) = mSQLReader.Item(i - 1) * Decimal.MinusOne
                    End If

                Next
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub DoInputData1(ByVal ACC1 As String, ByVal ACC2 As String, ByVal ACC3 As String, ByVal ACC4 As String, ByVal Keepit As Boolean)

        mSQLS1.CommandText = "select "
        For i As Int16 = 1 To tMonth Step 1
            mSQLS1.CommandText += "isnull(sum(t" & i & "),0) as t" & i & ","
        Next
        mSQLS1.CommandText += "1 from ( select "
        For i As Int16 = 1 To tMonth Step 1
            mSQLS1.CommandText += "(case when month1 = " & i & " then Amount1 else 0 end ) as t" & i & ","
        Next
        mSQLS1.CommandText += "1 AS NN from acais where year1 = " & tYear & " and month1 <= " & tMonth & " and Acc1 between '" & ACC1 & "' and '" & ACC2 & "' AND Acc1 not between '"
        mSQLS1.CommandText += ACC3 & "' AND '" & ACC4 & "'   ) AS AB "


        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows Then
            While mSQLReader.Read()
                For i As Int16 = 1 To tMonth Step 1
                    If Keepit = True Then
                        Ws.Cells(LineZ, 2 * i + 1) = mSQLReader.Item(i - 1)
                    Else
                        Ws.Cells(LineZ, 2 * i + 1) = mSQLReader.Item(i - 1) * Decimal.MinusOne
                    End If
                Next
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Income_Statement ACA " & tYear & tMonth
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
        If mConnection.State = ConnectionState.Open Then
            Try
                mConnection.Close()
                Module1.KillExcelProcess(OldExcel)
                MsgBox("Finished")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub DoInputData2(ByVal ACC1 As String, ByVal Keepit As Boolean)

        mSQLS1.CommandText = "select "
        For i As Int16 = 1 To tMonth Step 1
            mSQLS1.CommandText += "isnull(sum(t" & i & "),0) as t" & i & ","
        Next
        mSQLS1.CommandText += "1 from ( select "
        For i As Int16 = 1 To tMonth Step 1
            mSQLS1.CommandText += "(case when month1 = " & i & " then Amount1 else 0 end ) as t" & i & ","
        Next
        mSQLS1.CommandText += "1 AS NN from acais where year1 = " & tYear & " and month1 <= " & tMonth & " and Acc1 in (" & ACC1 & ") ) AS AB "


        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows Then
            While mSQLReader.Read()
                For i As Int16 = 1 To tMonth Step 1
                    If Keepit = True Then
                        Ws.Cells(LineZ, 2 * i + 1) = mSQLReader.Item(i - 1)
                    Else
                        Ws.Cells(LineZ, 2 * i + 1) = mSQLReader.Item(i - 1) * Decimal.MinusOne
                    End If
                Next
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub DoInputData(ByVal ACC1 As String, ByVal ACC2 As String, ByVal Keepit As Boolean, ByVal USDA As Boolean)

        mSQLS1.CommandText = "select "
        For i As Int16 = 1 To tMonth Step 1
            mSQLS1.CommandText += "isnull(sum(t" & i & "),0) as t" & i & ","
        Next
        mSQLS1.CommandText += "1 from ( select "
        For i As Int16 = 1 To tMonth Step 1
            mSQLS1.CommandText += "(case when month1 = " & i & " then Amount1 else 0 end ) as t" & i & ","
        Next
        mSQLS1.CommandText += "1 AS NN from acais where year1 = " & tYear & " and month1 <= " & tMonth & " and Acc1 between '" & ACC1 & "' and '" & ACC2 & "' ) AS AB "


        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows Then
            While mSQLReader.Read()
                For i As Int16 = 1 To tMonth Step 1
                    Dim Position As String = String.Empty
                    Select Case i
                        Case 1
                            Position = "C3"
                        Case 2
                            Position = "E3"
                        Case 3
                            Position = "G3"
                        Case 4
                            Position = "I3"
                        Case 5
                            Position = "K3"
                        Case 6
                            Position = "M3"
                        Case 7
                            Position = "O3"
                        Case 8
                            Position = "Q3"
                        Case 9
                            Position = "S3"
                        Case 10
                            Position = "U3"
                        Case 11
                            Position = "W3"
                        Case 12
                            Position = "Y3"

                    End Select
                    If Keepit = True Then
                        Ws.Cells(LineZ, 2 * i + 1) = "=" & mSQLReader.Item(i - 1) & "*" & Position
                    Else
                        Ws.Cells(LineZ, 2 * i + 1) = "=" & mSQLReader.Item(i - 1) * Decimal.MinusOne & "*" & Position
                    End If

                Next
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub DoInputData1(ByVal ACC1 As String, ByVal ACC2 As String, ByVal ACC3 As String, ByVal ACC4 As String, ByVal Keepit As Boolean, ByVal USDA As Boolean)

        mSQLS1.CommandText = "select "
        For i As Int16 = 1 To tMonth Step 1
            mSQLS1.CommandText += "isnull(sum(t" & i & "),0) as t" & i & ","
        Next
        mSQLS1.CommandText += "1 from ( select "
        For i As Int16 = 1 To tMonth Step 1
            mSQLS1.CommandText += "(case when month1 = " & i & " then Amount1 else 0 end ) as t" & i & ","
        Next
        mSQLS1.CommandText += "1 AS NN from acais where year1 = " & tYear & " and month1 <= " & tMonth & " and Acc1 between '" & ACC1 & "' and '" & ACC2 & "' AND Acc1 not between '"
        mSQLS1.CommandText += ACC3 & "' AND '" & ACC4 & "'   ) AS AB "


        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows Then
            While mSQLReader.Read()
                For i As Int16 = 1 To tMonth Step 1
                    Dim Position As String = String.Empty
                    Select Case i
                        Case 1
                            Position = "C3"
                        Case 2
                            Position = "E3"
                        Case 3
                            Position = "G3"
                        Case 4
                            Position = "I3"
                        Case 5
                            Position = "K3"
                        Case 6
                            Position = "M3"
                        Case 7
                            Position = "O3"
                        Case 8
                            Position = "Q3"
                        Case 9
                            Position = "S3"
                        Case 10
                            Position = "U3"
                        Case 11
                            Position = "W3"
                        Case 12
                            Position = "Y3"

                    End Select
                    If Keepit = True Then
                        Ws.Cells(LineZ, 2 * i + 1) = "=" & mSQLReader.Item(i - 1) & "*" & Position
                    Else
                        Ws.Cells(LineZ, 2 * i + 1) = "=" & mSQLReader.Item(i - 1) * Decimal.MinusOne & "*" & Position
                    End If
                Next
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub DoInputData2(ByVal ACC1 As String, ByVal Keepit As Boolean, ByVal USDA As Boolean)

        mSQLS1.CommandText = "select "
        For i As Int16 = 1 To tMonth Step 1
            mSQLS1.CommandText += "isnull(sum(t" & i & "),0) as t" & i & ","
        Next
        mSQLS1.CommandText += "1 from ( select "
        For i As Int16 = 1 To tMonth Step 1
            mSQLS1.CommandText += "(case when month1 = " & i & " then Amount1 else 0 end ) as t" & i & ","
        Next
        mSQLS1.CommandText += "1 AS NN from acais where year1 = " & tYear & " and month1 <= " & tMonth & " and Acc1 in (" & ACC1 & ") ) AS AB "


        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows Then
            While mSQLReader.Read()
                For i As Int16 = 1 To tMonth Step 1
                    Dim Position As String = String.Empty
                    Select Case i
                        Case 1
                            Position = "C3"
                        Case 2
                            Position = "E3"
                        Case 3
                            Position = "G3"
                        Case 4
                            Position = "I3"
                        Case 5
                            Position = "K3"
                        Case 6
                            Position = "M3"
                        Case 7
                            Position = "O3"
                        Case 8
                            Position = "Q3"
                        Case 9
                            Position = "S3"
                        Case 10
                            Position = "U3"
                        Case 11
                            Position = "W3"
                        Case 12
                            Position = "Y3"

                    End Select
                    If Keepit = True Then
                        Ws.Cells(LineZ, 2 * i + 1) = "=" & mSQLReader.Item(i - 1) & "*" & Position
                    Else
                        Ws.Cells(LineZ, 2 * i + 1) = "=" & mSQLReader.Item(i - 1) * Decimal.MinusOne & "*" & Position
                    End If
                Next
            End While
        End If
        mSQLReader.Close()
    End Sub
End Class