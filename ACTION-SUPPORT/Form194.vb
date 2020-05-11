Public Class Form194
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
    Dim Start2 As Date
    Dim CharC As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form194_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        tYear = Me.NumericUpDown1.Value
        tMonth = Me.NumericUpDown2.Value
        Start2 = DateTimePicker1.Value
        CharC = TextBox1.Text
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "产品工单领料与BOM表单阶材料比较明细表"
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
        'oCommand.CommandText = "Select distinct ccg_2.ccg04,ima_1.ima02,ima_1.ima021,ima_1.ima55,ccg_2.bmb03,ima_2.ima08,ima_2.ima02,ima_2.ima021,ima_2.ima25,"
        'oCommand.CommandText += "(stb07 + stb08 + stb09 + stb09a) as stbtot from ( "
        'oCommand.CommandText += "Select distinct ccg04,bmb03 from ccg_file left join bmb_file on ccg04 = bmb01 and bmb04 <= to_date('"
        'oCommand.CommandText += Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and (bmb05 is null or bmb05 > to_date('"
        'oCommand.CommandText += Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')) where ccg02 = " & tYear & " and ccg03 = " & tMonth
        'If Not String.IsNullOrEmpty(CharC) Then
        '    oCommand.CommandText += " and ccg04 LIKE '%" & CharC & "%' "
        'End If
        'oCommand.CommandText += " union all "
        'oCommand.CommandText += "Select distinct ccg04,cch04 from cch_file left join ccg_file on cch01 = ccg01 where cch02 = " & tYear & " and cch03 = " & tMonth & " and cch04 <> ' DL+OH+SUB' "
        'If Not String.IsNullOrEmpty(CharC) Then
        '    oCommand.CommandText += " and ccg04 LIKE '%" & CharC & "%' "
        'End If
        'oCommand.CommandText += ") ccg_2 left join ima_file ima_1 on ccg04 = ima_1.ima01 left join ima_file ima_2 on bmb03 = ima_2.ima01 left join stb_file on stb01 = ima_2.ima01 and stb02 = " & tYear & " and stb03 = " & tMonth
        'oCommand.CommandText += " group by ccg_2.ccg04,ima_1.ima02,ima_1.ima021,ima_1.ima55,ccg_2.bmb03,ima_2.ima08,ima_2.ima02,ima_2.ima021,ima_2.ima25,(stb07 + stb08 + stb09 + stb09a) "

        oCommand.CommandText = "Select distinct ccg_2.ccg04,ima_1.ima02,ima_1.ima021,ima_1.ima55,ccg_2.bmb03,ima_2.ima08,ima_2.ima02,ima_2.ima021,ima_2.ima25,"
        oCommand.CommandText += "(stb07 + stb08 + stb09 + stb09a) as stbtot,ccg_2.bmb29 from ( "
        oCommand.CommandText += "Select distinct ccg04,bmb03,bmb29 from ccg_file left join bmb_file on ccg04 = bmb01 and bmb04 <= to_date('"
        oCommand.CommandText += Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and (bmb05 is null or bmb05 > to_date('"
        oCommand.CommandText += Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')) where ccg02 = " & tYear & " and ccg03 = " & tMonth
        If Not String.IsNullOrEmpty(CharC) Then
            oCommand.CommandText += " and ccg04 LIKE '%" & CharC & "%' "
        End If
        oCommand.CommandText += " union all "
        oCommand.CommandText += "Select distinct ccg04,cch04,(case when sfb95 is null then ' ' else sfb95 end) from cch_file left join ccg_file on cch01 = ccg01 left join sfb_file on ccg01 = sfb01 where cch02 = " & tYear & " and cch03 = " & tMonth & " and cch04 <> ' DL+OH+SUB' "
        If Not String.IsNullOrEmpty(CharC) Then
            oCommand.CommandText += " and ccg04 LIKE '%" & CharC & "%' "
        End If
        oCommand.CommandText += ") ccg_2 left join ima_file ima_1 on ccg04 = ima_1.ima01 left join ima_file ima_2 on bmb03 = ima_2.ima01 left join stb_file on stb01 = ima_2.ima01 and stb02 = " & tYear & " and stb03 = " & tMonth
        oCommand.CommandText += " group by ccg_2.ccg04,ima_1.ima02,ima_1.ima021,ima_1.ima55,ccg_2.bmb03,ima_2.ima08,ima_2.ima02,ima_2.ima021,ima_2.ima25,(stb07 + stb08 + stb09 + stb09a),bmb29 "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = tYear
                Ws.Cells(LineZ, 2) = tMonth
                Ws.Cells(LineZ, 3) = oReader.Item(0)
                Ws.Cells(LineZ, 4) = oReader.Item(1)
                Ws.Cells(LineZ, 5) = oReader.Item(2)
                Ws.Cells(LineZ, 6) = oReader.Item(3)
                Ws.Cells(LineZ, 7) = oReader.Item(4)
                Ws.Cells(LineZ, 8) = oReader.Item(5)
                Ws.Cells(LineZ, 9) = oReader.Item(6)
                Ws.Cells(LineZ, 10) = oReader.Item(7)
                Dim l_bmb10 As String = String.Empty
                l_bmb10 = Getbmb10(oReader.Item(0), oReader.Item(4), oReader.Item(10))
                Ws.Cells(LineZ, 11) = l_bmb10
                Dim l_ima25 As String = oReader.Item(8)
                Ws.Cells(LineZ, 12) = l_ima25
                Dim l_bmb10_fac As Decimal = Getbmb10_fac(oReader.Item(0), oReader.Item(4), oReader.Item(10))
                If l_bmb10 = "" Then
                    Ws.Cells(LineZ, 13) = 0
                    Ws.Cells(LineZ, 24) = 0
                Else
                    Ws.Cells(LineZ, 13) = l_bmb10_fac
                    Ws.Cells(LineZ, 24) = "=N" & LineZ & "/M" & LineZ & "*R" & LineZ
                End If
                Ws.Cells(LineZ, 14) = oReader.Item(9)
                Getccgcch(oReader.Item(0), oReader.Item(4), oReader.Item(10))
                Dim l_bmb06 As Decimal = Getbmb06(oReader.Item(0), oReader.Item(4), oReader.Item(10))
                Ws.Cells(LineZ, 18) = l_bmb06
                'Ws.Cells(LineZ, 15) = oReader.Item(10)
                'Ws.Cells(LineZ, 16) = oReader.Item(11)
                'If Not oReader.Item(10) = 0 Then
                '    Ws.Cells(LineZ, 17) = oReader.Item(11) / oReader.Item(10)
                'End If
                'If l_bmb10 = "" Then
                '    Ws.Cells(LineZ, 18) = 0
                'Else
                '    Ws.Cells(LineZ, 18) = oReader.Item(12)
                'End If
                Ws.Cells(LineZ, 19) = "=Q" & LineZ & "-R" & LineZ
                'Ws.Cells(LineZ, 20) = oReader.Item(13)
                'Ws.Cells(LineZ, 21) = oReader.Item(14)
                'Ws.Cells(LineZ, 22) = oReader.Item(15)
                'If Not oReader.Item(10) = 0 Then
                '    Ws.Cells(LineZ, 23) = oReader.Item(15) / oReader.Item(10)
                'End If
                'If l_bmb10 = "" Then
                '    Ws.Cells(LineZ, 24) = 0
                'Else
                '    If l_bmb10_fac <> 0 Then
                '        Ws.Cells(LineZ, 24) = oReader.Item(9) / l_bmb10_fac
                '    End If
                'End If
                Ws.Cells(LineZ, 25) = "=W" & LineZ & "-X" & LineZ
                Ws.Cells(LineZ, 26) = oReader.Item(10)
                LineZ += 1
                Label4.Text = LineZ
                Label4.Refresh()
            End While
        End If
        oReader.Close()
        oRng = Ws.Range("A1", "Y1")
        oRng.EntireColumn.AutoFit()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Cells(1, 1) = "年度"
        Ws.Cells(1, 2) = "月份"
        Ws.Cells(1, 3) = "主件料号"
        Ws.Cells(1, 4) = "品名"
        Ws.Cells(1, 5) = "规格"
        Ws.Cells(1, 6) = "生产单位"
        Ws.Cells(1, 7) = "元件料号"
        Ws.Cells(1, 8) = "来源码"
        Ws.Cells(1, 9) = "品名"
        Ws.Cells(1, 10) = "规格"
        Ws.Cells(1, 11) = "BOM表单位"
        Ws.Cells(1, 12) = "元件料号库存单位"
        Ws.Cells(1, 13) = "单位换算率"
        Ws.Cells(1, 14) = "单位标准总成本"
        Ws.Cells(1, 15) = "本期主件料号完工入库数量"
        Ws.Cells(1, 16) = "本期元件料号扣料数量"
        Ws.Cells(1, 17) = "单位实际用量"
        Ws.Cells(1, 18) = "单位标准用量"
        Ws.Cells(1, 19) = "单位用量差异"
        Ws.Cells(1, 20) = "下阶报废数量"
        Ws.Cells(1, 21) = "主件料号报废数量"
        Ws.Cells(1, 22) = "本期元件料号扣料金额"
        Ws.Cells(1, 23) = "单位实际成本"
        Ws.Cells(1, 24) = "单位标准成本"
        Ws.Cells(1, 25) = "单位成本差异"
        Ws.Cells(1, 26) = "主件料号特性代码"
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        oRng = Ws.Range("G1", "G1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        LineZ = 2
    End Sub
    Private Function Getbmb10(ByVal S1 As String, ByVal S2 As String, ByVal S99 As String)
        oCommand2.CommandText = "Select nvl(bmb10,'') FROM bmb_file WHERE bmb01 = '" & S1 & "' and bmb03 = '" & S2 & "' and bmb04 <= to_date('"
        oCommand2.CommandText += Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and (bmb05 is null or bmb05 > to_date('"
        oCommand2.CommandText += Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')) and bmb29 = '" & S99 & "'"
        Dim S3 As String = oCommand2.ExecuteScalar()
        Return S3
    End Function
    Private Function Getbmb10_fac(ByVal S1 As String, ByVal S2 As String, ByVal S99 As String)
        oCommand2.CommandText = "Select nvl(Round(1/bmb10_fac,2),0) FROM bmb_file WHERE bmb01 = '" & S1 & "' and bmb03 = '" & S2 & "' and bmb04 <= to_date('"
        oCommand2.CommandText += Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and (bmb05 is null or bmb05 > to_date('"
        oCommand2.CommandText += Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')) and bmb29 = '" & S99 & "'"
        Dim S3 As String = oCommand2.ExecuteScalar()
        Return S3
    End Function
    Private Function Getbmb06(ByVal S1 As String, ByVal S2 As String, ByVal S99 As String)
        oCommand2.CommandText = "Select nvl(Round(SUM(bmb06/bmb07 * (1+ bmb08/100)),8),0) FROM bmb_file WHERE bmb01 = '" & S1 & "' and bmb03 = '" & S2 & "' and bmb04 <= to_date('"
        oCommand2.CommandText += Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and (bmb05 is null or bmb05 > to_date('"
        oCommand2.CommandText += Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')) and bmb29 = '" & S99 & "'"
        Dim S3 As String = oCommand2.ExecuteScalar()
        Return S3
    End Function


    Private Sub Getccgcch(ByVal s1 As String, s2 As String, ByVal S3 As String)
        oCommand2.CommandText = "Select nvl(sum(ccg31),0),nvl(sum(cch31),0),nvl(sum(cch311),0),nvl(sum(ccg11+ccg21+ccg31-ccg91),0),nvl(sum(cch32),0) from ccg_file,cch_file,sfb_file where ccg01 = cch01 and ccg01 = sfb01 and cch01 = sfb01 and ccg02 = " & tYear & " and ccg03 = " & tMonth & " and cch02 = ccg02 and cch03 = ccg03 "
        oCommand2.CommandText += " and ccg04 = '" & s1 & "' AND cch04 = '" & s2 & "'"
        If S3 = " " Then
            oCommand2.CommandText += " and (sfb95 = ' ' or sfb95 is null) "
        Else
            oCommand2.CommandText += " and sfb95 = '" & S3 & "'"
        End If
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                Dim GGA As Decimal = oReader2.Item(0)
                Ws.Cells(LineZ, 15) = GGA
                Ws.Cells(LineZ, 16) = oReader2.Item(1)
                If GGA <> 0 Then
                    Ws.Cells(LineZ, 17) = "=P" & LineZ & "/O" & LineZ
                End If
                Ws.Cells(LineZ, 20) = oReader2.Item(2)
                Ws.Cells(LineZ, 21) = oReader2.Item(3)
                Ws.Cells(LineZ, 22) = oReader2.Item(4)
                If GGA <> 0 Then
                    Ws.Cells(LineZ, 23) = "=V" & LineZ & "/O" & LineZ
                End If

            End While
        End If
        oReader2.Close()
    End Sub
End Class