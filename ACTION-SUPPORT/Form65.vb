Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form65
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim tYear As String = String.Empty
    Dim TotalWeek As Int16 = 0
    Dim DStartN As Date
    Dim DStartE As Date
    Dim LineZ As Integer = 0
    Dim ArrayX1 As String() = {"102020020028", "102010010019", "102010010027", "102020020022", "102010010002", "102010010034", "102020020008", "102010010012", "102010020012", _
                               "102010010039", "102010010040", "102010010041", "102020020014", "102020020016", "102010020002", "102010010013", "102010010009", "102010010007", _
                               "102010010011", "102020020009", "102010010023"}
    Dim ArrayX2 As String() = {"205000010003", "206020010001", "206030020005", "206020010002", "206010010002", "205000010001", "204000020011"}
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form65_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        Me.TextBox1.Text = Today.Year()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommander2.Connection = oConnection
                oCommander2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        tYear = Me.TextBox1.Text
        'DStartN = Me.TextBox1.Text & "/01/01"
        'DStartE = DStartN.AddYears(1).AddDays(-1)
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Label2.Text = "处理中"
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Label2.Text = "处理完毕"
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Main_Material_LostRate"
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
        Ws.Name = "碳布"
        AdjustExcelFormat()
        For i As Int16 = 0 To ArrayX1.Length - 1 Step 1
            oCommand.CommandText = "select ima02 FROM ima_file where ima01 = '" & ArrayX1(i).ToString() & "'"
            Dim l_ima02 As String = oCommand.ExecuteScalar()
            oCommand.CommandText = "select ima021 FROM ima_file where ima01 = '" & ArrayX1(i).ToString() & "'"
            Dim l_ima021 As String = oCommand.ExecuteScalar()
            oCommand.CommandText = "select ima25 FROM ima_file where ima01 = '" & ArrayX1(i).ToString() & "'"
            Dim l_ima25 As String = oCommand.ExecuteScalar()
            Ws.Cells(LineZ, 2) = 1 + 3 * i
            Ws.Cells(LineZ + 1, 2) = 2 + 3 * i
            Ws.Cells(LineZ + 2, 2) = 3 + 3 * i
            Ws.Cells(LineZ, 3) = ArrayX1(i).ToString()
            Ws.Cells(LineZ + 1, 3) = ArrayX1(i).ToString()
            Ws.Cells(LineZ + 2, 3) = ArrayX1(i).ToString()
            Ws.Cells(LineZ, 4) = l_ima02
            Ws.Cells(LineZ + 1, 4) = l_ima02
            Ws.Cells(LineZ + 2, 4) = l_ima02
            Ws.Cells(LineZ, 5) = l_ima021
            Ws.Cells(LineZ + 1, 5) = l_ima021
            Ws.Cells(LineZ + 2, 5) = l_ima021
            Ws.Cells(LineZ, 6) = l_ima25
            Ws.Cells(LineZ + 1, 6) = l_ima25
            Ws.Cells(LineZ + 2, 6) = l_ima25
            Ws.Cells(LineZ, 7) = "理論標準使用量"
            Ws.Cells(LineZ + 1, 7) = "实际領用量"
            Ws.Cells(LineZ + 2, 7) = "实际与标准用量对比"
            oRng = Ws.Range(Ws.Cells(LineZ + 2, 8), Ws.Cells(LineZ + 2, 7 + TotalWeek))
            'oRng.EntireRow.NumberFormatLocal = "0.00%"
            oRng.NumberFormatLocal = "0.00%'"
            For j As Int16 = 1 To TotalWeek Step 1
                DStartN = GDate1(j)
                DStartE = GDate2(j)
                oCommand.CommandText = "select nvl(round(sum((sfv09+sfvud07) * bmb06 /bmb07 * (1+bmb08 / 100)),3),0) from sfu_file,sfv_file,bmb_file where sfu01 = sfv01 and sfupost = 'Y' and sfu02 between to_date('"
                oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                oCommand.CommandText += DStartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and sfu04 = 'D3531' and sfv04 = bmb01 and bmb03 = '"
                oCommand.CommandText += ArrayX1(i).ToString() & "' and (bmb05 is null or bmb05 <= to_date('"
                oCommand.CommandText += DStartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd'))"
                Dim A1 As Decimal = oCommand.ExecuteScalar()
                Ws.Cells(LineZ, 7 + j) = A1
                Dim A2 As Decimal = GetA2(ArrayX1(i).ToString())
                Dim A3 As Decimal = GetA3(ArrayX1(i).ToString())
                Ws.Cells(LineZ + 1, 7 + j) = A2 - A3
                If A1 <> 0 Then
                    Ws.Cells(LineZ + 2, 7 + j) = ((A2 - A3) - A1) / A1
                End If
            Next
            LineZ += 3
        Next

        ' 第二頁

        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        Ws.Name = "油漆"
        AdjustExcelFormat1()
        For i As Int16 = 0 To ArrayX2.Length - 1 Step 1
            oCommand.CommandText = "select ima02 FROM ima_file where ima01 = '" & ArrayX2(i).ToString() & "'"
            Dim l_ima02 As String = oCommand.ExecuteScalar()
            oCommand.CommandText = "select ima021 FROM ima_file where ima01 = '" & ArrayX2(i).ToString() & "'"
            Dim l_ima021 As String = oCommand.ExecuteScalar()
            oCommand.CommandText = "select ima25 FROM ima_file where ima01 = '" & ArrayX2(i).ToString() & "'"
            Dim l_ima25 As String = oCommand.ExecuteScalar()
            Ws.Cells(LineZ, 2) = 1 + 3 * i
            Ws.Cells(LineZ + 1, 2) = 2 + 3 * i
            Ws.Cells(LineZ + 2, 2) = 3 + 3 * i
            Ws.Cells(LineZ, 3) = ArrayX1(i).ToString()
            Ws.Cells(LineZ + 1, 3) = ArrayX1(i).ToString()
            Ws.Cells(LineZ + 2, 3) = ArrayX1(i).ToString()
            Ws.Cells(LineZ, 4) = l_ima02
            Ws.Cells(LineZ + 1, 4) = l_ima02
            Ws.Cells(LineZ + 2, 4) = l_ima02
            Ws.Cells(LineZ, 5) = l_ima021
            Ws.Cells(LineZ + 1, 5) = l_ima021
            Ws.Cells(LineZ + 2, 5) = l_ima021
            Ws.Cells(LineZ, 6) = l_ima25
            Ws.Cells(LineZ + 1, 6) = l_ima25
            Ws.Cells(LineZ + 2, 6) = l_ima25
            Ws.Cells(LineZ, 7) = "理論標準使用量"
            Ws.Cells(LineZ + 1, 7) = "实际領用量"
            Ws.Cells(LineZ + 2, 7) = "实际与标准用量对比"
            oRng = Ws.Range(Ws.Cells(LineZ + 2, 8), Ws.Cells(LineZ + 2, 7 + TotalWeek))
            'oRng.EntireRow.NumberFormatLocal = "0.00%"
            oRng.NumberFormatLocal = "0.00%'"
            For j As Int16 = 1 To TotalWeek Step 1
                DStartN = GDate1(j)
                DStartE = GDate2(j)
                oCommand.CommandText = "select nvl(round(sum((sfv09+sfvud07) * bmb06 /bmb07 * (1+bmb08 / 100)),3),0) from sfu_file,sfv_file,bmb_file where sfu01 = sfv01 and sfupost = 'Y' and sfu02 between to_date('"
                oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                oCommand.CommandText += DStartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and sfu04 in ('D3561','D3563') and sfv04 = bmb01 and bmb03 = '"
                oCommand.CommandText += ArrayX2(i).ToString() & "' and (bmb05 is null or bmb05 <= to_date('"
                oCommand.CommandText += DStartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd'))"
                Dim A1 As Decimal = oCommand.ExecuteScalar()
                Ws.Cells(LineZ, 7 + j) = A1
                Dim A4 As Decimal = GetA4(ArrayX2(i).ToString())
                Ws.Cells(LineZ + 1, 7 + j) = A4
                If A1 <> 0 Then
                    Ws.Cells(LineZ + 2, 7 + j) = (A4 - A1) / A1
                End If
            Next
            LineZ += 3
        Next
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "B1")
        oRng.EntireColumn.ColumnWidth = 4
        oRng = Ws.Range("C1", "D1")
        oRng.EntireColumn.ColumnWidth = 17.11
        oRng = Ws.Range("E1", "E1")
        oRng.EntireColumn.ColumnWidth = 100
        oRng = Ws.Range("F1", "F1")
        oRng.EntireColumn.ColumnWidth = 10.44
        oRng = Ws.Range("G1", "G1")
        oRng.EntireColumn.ColumnWidth = 21.89
        oRng = Ws.Range(Ws.Cells(1, 1), Ws.Cells(1, 60))
        oRng.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A2", "A64")   '若有增加料號需調整此部份
        oRng.Merge()
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.NumberFormat = "@"
        Ws.Cells(1, 1) = "类别"
        Ws.Cells(1, 2) = "序号"
        Ws.Cells(1, 3) = "料号"
        Ws.Cells(1, 4) = "品名"
        Ws.Cells(1, 5) = "规格"
        Ws.Cells(1, 6) = "标准单位"
        oCommand.CommandText = "select count(distinct azn05) from azn_file where azn02 = " & tYear
        Try
            TotalWeek = oCommand.ExecuteScalar()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        For i As Int16 = 1 To TotalWeek Step 1
            Ws.Cells(1, 7 + i) = "W" & i
        Next
        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "B1")
        oRng.EntireColumn.ColumnWidth = 4
        oRng = Ws.Range("C1", "D1")
        oRng.EntireColumn.ColumnWidth = 17.11
        oRng = Ws.Range("E1", "E1")
        oRng.EntireColumn.ColumnWidth = 100
        oRng = Ws.Range("F1", "F1")
        oRng.EntireColumn.ColumnWidth = 10.44
        oRng = Ws.Range("G1", "G1")
        oRng.EntireColumn.ColumnWidth = 21.89
        oRng = Ws.Range(Ws.Cells(1, 1), Ws.Cells(1, 60))
        oRng.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A2", "A22")   '若有增加料號需調整此部份
        oRng.Merge()
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.NumberFormat = "@"
        Ws.Cells(1, 1) = "类别"
        Ws.Cells(1, 2) = "序号"
        Ws.Cells(1, 3) = "料号"
        Ws.Cells(1, 4) = "品名"
        Ws.Cells(1, 5) = "规格"
        Ws.Cells(1, 6) = "标准单位"
        oCommand.CommandText = "select count(distinct azn05) from azn_file where azn02 = " & tYear
        Try
            TotalWeek = oCommand.ExecuteScalar()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        For i As Int16 = 1 To TotalWeek Step 1
            Ws.Cells(1, 7 + i) = "W" & i
        Next
        LineZ = 2
    End Sub
    Private Function GDate1(ByVal gWeek As Int16)
        oCommander2.CommandText = "select min(azn01) from azn_file where azn02 = " & tYear & " and azn05 = " & gWeek
            Dim DN As Date = oCommander2.ExecuteScalar()
            Return DN
    End Function
    Private Function GDate2(ByVal gWeek As Int16)
        oCommander2.CommandText = "select max(azn01) from azn_file where azn02 = " & tYear & " and azn05 = " & gWeek
        Dim DN As Date = oCommander2.ExecuteScalar()
        Return DN
    End Function
    Private Function GetA2(ByVal tlf01 As String)
        oCommander2.CommandText = "SELECT nvl(sum(tlf10*tlf12),0) FROM tlf_file where tlf01 = '" & tlf01 & "' and tlf06 between to_date('"
        oCommander2.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommander2.CommandText += DStartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tlf907 = -1 and tlf902 = 'D353102'"
        Dim A2 As Decimal = oCommander2.ExecuteScalar()
        Return A2
    End Function
    Private Function GetA3(ByVal tlf01 As String)
        oCommander2.CommandText = "SELECT nvl(sum(tlf10*tlf12),0) FROM tlf_file where tlf01 = '" & tlf01 & "' and tlf06 between to_date('"
        oCommander2.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommander2.CommandText += DStartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tlf907 = -1 and tlf902 = 'D353102' and tlf13 = 'aimt324'"
        Dim A2 As Decimal = oCommander2.ExecuteScalar()
        Return A2
    End Function
    Private Function GetA4(ByVal tlf01 As String)
        oCommander2.CommandText = "SELECT nvl(sum(tlf10*tlf12),0) FROM tlf_file where tlf01 = '" & tlf01 & "' and tlf06 between to_date('"
        oCommander2.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommander2.CommandText += DStartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tlf907 = -1 and tlf902 in ('D356302','D356102')"
        Dim A2 As Decimal = oCommander2.ExecuteScalar()
        Return A2
    End Function
End Class