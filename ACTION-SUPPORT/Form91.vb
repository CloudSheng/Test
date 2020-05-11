Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form91
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim pYear As Int16 = 0
    Dim pMonth As Int16 = 0
    Dim aYear As Int16 = 0
    Dim aMonth As Int16 = 0
    Dim Start1 As String = String.Empty
    Dim End1 As String = String.Empty
    Dim Start2 As Date
    Dim End2 As Date
    Dim TotalPeriod As Int16 = 0
    Dim LineZ As Integer = 0
    Dim SC As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form91_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        If CheckBox1.Checked = False And CheckBox2.Checked = False And CheckBox3.Checked = False And CheckBox4.Checked = False Then
            MsgBox("输出资料有误")
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
        pYear = Strings.Left(Start1, 4)
        pMonth = Strings.Right(Start1, 2)
        tYear = Strings.Left(End1, 4)
        tMonth = Strings.Right(End1, 2)
        Start2 = Convert.ToDateTime(pYear & "/" & pMonth & "/01")
        End2 = Convert.ToDateTime(tYear & "/" & tMonth & "/01").AddMonths(1).AddDays(-1)
        SC = TextBox1.Text
        Label6.Text = 0
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "MISC_INOROUT_REPORT"
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
        oCommand.CommandText = "select tlf01,type,ima02,ima021,tlf11"
        For i As Int16 = 1 To TotalPeriod Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        oCommand.CommandText += " from ( "
        If CheckBox1.Checked = True Then
            oCommand.CommandText += "select tlf01,'1' as type,ima02,ima021,tlf11"
            For i As Int16 = 1 To TotalPeriod Step 1
                oCommand.CommandText += " ,(case when tlf06 between to_date('" & Start2.AddMonths(i - 1).ToString("yyyy/MM/dd")
                oCommand.CommandText += "','yyyy/mm/dd') and to_date('" & Start2.AddMonths(i).AddDays(-1).ToString("yyyy/MM/dd")
                oCommand.CommandText += "','yyyy/mm/dd') then sum(tlf10 * tlf12) else 0 end) as t" & i
            Next
            oCommand.CommandText += " from tlf_file,ima_file where tlf13 = 'aimt311'  and tlf06 between to_date('"
            oCommand.CommandText += Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            oCommand.CommandText += End2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tlf01 = ima01 and ima08 in ('P','S') "
            If Not String.IsNullOrEmpty(SC) Then
                oCommand.CommandText += " AND tlf01 = '" & SC & "' "
            End If
            oCommand.CommandText += " group by tlf01,tlf06,ima02,ima021,tlf11 "
        End If
        If CheckBox2.Checked = True Then
            If CheckBox1.Checked = True Then
                oCommand.CommandText += "union all "
            End If
            oCommand.CommandText += "select tlf01,'2' as type,ima02,ima021,tlf11"
            For i As Int16 = 1 To TotalPeriod Step 1
                aMonth = pMonth + i - 2
                aYear = pYear
                If aMonth = 0 Then
                    aMonth = 12
                    aYear -= 1
                End If
                If aMonth > 12 Then
                    aMonth -= 12
                    aYear += 1
                End If
                oCommand.CommandText += " ,(case when tlf06 between to_date('" & Start2.AddMonths(i - 1).ToString("yyyy/MM/dd")
                oCommand.CommandText += "','yyyy/mm/dd') and to_date('" & Start2.AddMonths(i).AddDays(-1).ToString("yyyy/MM/dd")
                oCommand.CommandText += "','yyyy/mm/dd') then (case when tlf21 is null then (select nvl(sum(ccc23 * tlf10 * tlf12 ),0) from ccc_file where ccc01 = tlf01 and ccc02 = "
                oCommand.CommandText += aYear & " and ccc03 = " & aMonth & " ) else sum(tlf21) end) else 0 end) as t" & i
            Next
            oCommand.CommandText += " from tlf_file,ima_file where tlf13 = 'aimt311' and tlf06 between to_date('"
            oCommand.CommandText += Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            oCommand.CommandText += End2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tlf01 = ima01 and ima08 in ('P','S') "
            If Not String.IsNullOrEmpty(SC) Then
                oCommand.CommandText += " AND tlf01 = '" & SC & "' "
            End If
            oCommand.CommandText += "group by tlf01,tlf06,tlf21,ima02,ima021,tlf11,tlf10,tlf12 "
        End If
        If CheckBox3.Checked = True Then
            If CheckBox1.Checked = True Or CheckBox2.Checked = True Then
                oCommand.CommandText += "union all "
            End If
            oCommand.CommandText += "select tlf01,'3' as type,ima02,ima021,tlf11"
            For i As Int16 = 1 To TotalPeriod Step 1
                oCommand.CommandText += " ,(case when tlf06 between to_date('" & Start2.AddMonths(i - 1).ToString("yyyy/MM/dd")
                oCommand.CommandText += "','yyyy/mm/dd') and to_date('" & Start2.AddMonths(i).AddDays(-1).ToString("yyyy/MM/dd")
                oCommand.CommandText += "','yyyy/mm/dd') then sum(tlf10 * tlf12) else 0 end) as t" & i
            Next
            oCommand.CommandText += " from tlf_file,ima_file where tlf13 = 'aimt312' and tlf06 between to_date('"
            oCommand.CommandText += Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            oCommand.CommandText += End2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tlf01 = ima01 and ima08 in ('P','S') "
            If Not String.IsNullOrEmpty(SC) Then
                oCommand.CommandText += " AND tlf01 = '" & SC & "' "
            End If
            oCommand.CommandText += " group by tlf01,tlf06,ima02,ima021,tlf11 "
        End If
        If CheckBox4.Checked = True Then
            If CheckBox1.Checked = True Or CheckBox2.Checked = True Or CheckBox3.Checked = True Then
                oCommand.CommandText += "union all "
            End If
            oCommand.CommandText += "select tlf01,'4' as type,ima02,ima021,tlf11"
            For i As Int16 = 1 To TotalPeriod Step 1
                aMonth = pMonth + i - 2
                aYear = pYear
                If aMonth = 0 Then
                    aMonth = 12
                    aYear -= 1
                End If
                If aMonth > 12 Then
                    aMonth -= 12
                    aYear += 1
                End If
                oCommand.CommandText += " ,(case when tlf06 between to_date('" & Start2.AddMonths(i - 1).ToString("yyyy/MM/dd")
                oCommand.CommandText += "','yyyy/mm/dd') and to_date('" & Start2.AddMonths(i).AddDays(-1).ToString("yyyy/MM/dd")
                oCommand.CommandText += "','yyyy/mm/dd') then (select nvl(sum(ccc23 * tlf10 * tlf12),0) from ccc_file where ccc01 = tlf01 and ccc02 = "
                oCommand.CommandText += aYear & " and ccc03 = " & aMonth & " ) else 0 end) as t" & i
            Next
            oCommand.CommandText += " from tlf_file,ima_file where tlf13 = 'aimt312' and tlf06 between to_date('"
            oCommand.CommandText += Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            oCommand.CommandText += End2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tlf01 = ima01 and ima08 in ('P','S') "
            If Not String.IsNullOrEmpty(SC) Then
                oCommand.CommandText += " AND tlf01 = '" & SC & "' "
            End If
            oCommand.CommandText += "group by tlf01,tlf06,tlf21,ima02,ima021,tlf11,tlf10,tlf12 "
        End If
        oCommand.CommandText += ") group by tlf01,type,ima02,ima021,tlf11"
        oReader = oCommand.ExecuteReader()
        Dim TR As Decimal = 0
        If oReader.HasRows Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("tlf01")
                Ws.Cells(LineZ, 2) = oReader.Item("ima02")
                Ws.Cells(LineZ, 3) = oReader.Item("ima021")
                Select Case oReader.Item("type")
                    Case 1
                        Ws.Cells(LineZ, 4) = "杂发数量"
                    Case 2
                        Ws.Cells(LineZ, 4) = "杂发金额"
                    Case 3
                        Ws.Cells(LineZ, 4) = "杂收数量"
                    Case 4
                        Ws.Cells(LineZ, 4) = "杂收金额"
                End Select
                Ws.Cells(LineZ, 5) = oReader.Item("tlf11")
                For i As Int16 = 1 To TotalPeriod Step 1
                    Ws.Cells(LineZ, 5 + i) = oReader.Item(4 + i)
                Next
                LineZ += 1
                TR += 1
                Label6.Text = TR
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.ColumnWidth = 17.44
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(1, 1) = "料号"
        Ws.Cells(1, 2) = "品名"
        Ws.Cells(1, 3) = "规格"
        Ws.Cells(1, 4) = "项目"
        Ws.Cells(1, 5) = "单位"
        For i As Integer = 1 To TotalPeriod Step 1
            Dim TMonth As Int16 = Conversion.Int(Strings.Right(Start1, 2)) + i - 1
            If TMonth > 12 Then
                If TMonth - 12 < 10 Then
                    Ws.Cells(1, 5 + i) = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "/0" & TMonth - 12
                Else
                    Ws.Cells(1, 5 + i) = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "/" & TMonth - 12
                End If
            Else
                If TMonth < 10 Then
                    Ws.Cells(1, 5 + i) = Conversion.Int(Strings.Left(Start1, 4)) & "/0" & TMonth
                Else
                    Ws.Cells(1, 5 + i) = Conversion.Int(Strings.Left(Start1, 4)) & "/" & TMonth
                End If
            End If
        Next
        LineZ = 2
    End Sub
End Class