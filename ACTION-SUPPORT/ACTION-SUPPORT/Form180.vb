Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form180
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim TYear As Int16 = 0
    Dim Tweek As Int16 = 0
    Dim CYear1 As Int16 = 0
    Dim CWeek1 As Int16 = 0
    Dim CYear2 As Int16 = 0
    Dim CWeek2 As Int16 = 0
    Dim TotalWeek As Int16 = 0
    Dim SDate As Date
    Dim LineZ As Integer = 0
    Dim LineS1 As Int16 = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form180_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        TextBox1.Text = Now.Year
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
        oCommand.CommandText = "Select azn05 FROM azn_file where azn01 = to_date('" & Today.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        Dim TZ1 As Int16 = oCommand.ExecuteScalar()
        TextBox2.Text = TZ1
        TYear = TextBox1.Text
        Tweek = TextBox2.Text

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Dim xPath As String = "C:\temp\费用汇总周报表模板Template Rev2.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        ' 得到週數的前一週和年
        Tweek = TextBox2.Text
        CWeek2 = Tweek - 1
        'If CWeek2 = 0 Then
        '    CYear2 = TYear - 1
        '    CWeek2 = 53
        'Else
        '    CYear2 = TYear
        'End If
        If CWeek2 <= 0 Then
            CWeek2 = 1
        End If

        ' 得到前12週的值
        CWeek1 = CWeek2 - 11
        'If CWeek1 <= 0 Then
        '    CYear1 = CYear2 - 1
        '    CWeek1 = 53 + CWeek1
        'Else
        '    CYear1 = CYear2
        'End If
        If CWeek1 <= 0 Then
            CWeek1 = 1
        End If
        CYear1 = TYear
        CYear2 = TYear
        TotalWeek = CWeek2 - CWeek1 + 1
        SDate = Convert.ToDateTime(TYear & "/01/01")
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "费用汇总周报表"
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
        Dim xPath As String = "C:\temp\费用汇总周报表模板Template Rev2.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
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
        ' 先看日期
        oCommand.CommandText = "select min(azn01) from azn_file where azn02 = " & CYear1 & " and azn05 = " & CWeek1
        Dim MinD As Date = oCommand.ExecuteScalar()
        oCommand.CommandText = "select max(azn01) from azn_file where azn02 = " & CYear2 & " and azn05 = " & CWeek2
        Dim MaxD As Date = oCommand.ExecuteScalar()

        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        LineZ = 5

        For i As Int16 = TotalWeek To 1 Step -1
            Ws.Cells(4, 15 - i) = "W" & CWeek2 - i + 1
        Next

        'oCommand.CommandText = "Select aae02,sum(t1) as t1,sum(t2) as t2, sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 "
        'oCommand.CommandText += "from ( select aae01,aae02"
        oCommand.CommandText = "Select aae02"
        For i As Int16 = 1 To TotalWeek Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        oCommand.CommandText += ",sum(x1) as x1,sum(x2) as x2 from ( select aae01,aae02"
        'For i As Int16 = 0 To 11 Step 1
        '    Dim sWeek As Int16 = CWeek1 + i
        '    Dim sYear As Int16 = CYear1
        '    If sWeek > 53 Then
        '        sWeek = sWeek - 53
        '        sYear = sYear + 1
        '    End If
        '    Ws.Cells(4, 3 + i) = "W" & sWeek
        '    oCommand.CommandText += ",(Case when azn02 = " & sYear & " and azn05 = " & sWeek & "  then (case when abb06 = 1 then abb07 else abb07 * -1 end) else 0 end) as t" & i + 1
        'Next
        For i As Int16 = 1 To TotalWeek Step 1
            Dim sWeek As Int16 = CWeek1 + i - 1
            Dim sYear As Int16 = CYear1
            oCommand.CommandText += ",(Case when azn02 = " & sYear & " and azn05 = " & sWeek & "  then (case when abb06 = 1 then abb07 else abb07 * -1 end) else 0 end) as t" & i
        Next
        oCommand.CommandText += ",0 as x1,0 as x2 from aea_file left join abb_file on aea03 = abb01 and aea04 = abb02 and aea05 = abb03 left join aag_file on aea05 = aag01 left join aae_file on aag223 = aae01 "
        oCommand.CommandText += "left join azn_file on aea02 = azn01 where aag223 is not null and abb05 <> 'D9999' and aea02 between to_date('" & MinD.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += MaxD.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "Select aae01,aae02"
        For i As Int16 = 1 To TotalWeek Step 1
            oCommand.CommandText += ",0"
        Next
        oCommand.CommandText += ",sum(case when abb06 = 1 then abb07 else abb07 * -1 end),0 "
        oCommand.CommandText += "from aea_file left join abb_file on aea03 = abb01 and aea04 = abb02 and aea05 = abb03 left join aag_file on aea05 = aag01 left join aae_file on aag223 = aae01 "
        oCommand.CommandText += "where aag223 is not null and abb05 <> 'D9999' and aea02 between to_date('"
        oCommand.CommandText += SDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & MaxD.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') group by aae01,aae02 "
        oCommand.CommandText += "union all "
        If TYear = 2019 Then
            oCommand.CommandText += "Select aae01,aae02"
            For i As Int16 = 1 To TotalWeek Step 1
                oCommand.CommandText += ",0"
            Next
            oCommand.CommandText += ",0,sum(budget) from DAC_2019_Budget left join aag_file on acc1 = aag01 left join aae_file on aag223 = aae01 where aag223 is not null and month1 <= " & MaxD.Month & " group by aae01,aae02"
        Else
            oCommand.CommandText += "Select aae01,aae02"
            For i As Int16 = 1 To TotalWeek Step 1
                oCommand.CommandText += ",0"
            Next
            oCommand.CommandText += ",0,sum(tc_bud13) from tc_bud_file left join aag_file on tc_bud07 = aag01 left join aae_file on aag223 = aae01 where tc_bud01 = 2 and tc_bud02 = "
            oCommand.CommandText += TYear & " and tc_bud03 <= " & MaxD.Month & " and aag23 is not null and tc_bud08 <> 'D9999' group by aae01,aae02"
        End If
        oCommand.CommandText += " ) group by aae01,aae02 order by aae01"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            Dim SNS As Int16 = 1
            While oReader.Read()
                Ws.Cells(LineZ, 1) = SNS
                Ws.Cells(LineZ, 2) = oReader.Item("aae02")
                'For i As Int16 = 1 To 12 Step 1
                '    Ws.Cells(LineZ, i + 2) = oReader.Item(i)
                'Next
                For i As Int16 = TotalWeek To 1 Step -1
                    Ws.Cells(LineZ, 15 - i) = oReader.Item(oReader.FieldCount - 2 - i)
                Next
                Ws.Cells(LineZ, 15) = oReader.Item(oReader.FieldCount - 2)
                Ws.Cells(LineZ, 16) = oReader.Item(oReader.FieldCount - 1)
                Ws.Cells(LineZ, 17) = "=O" & LineZ & "-P" & LineZ
                Ws.Cells(LineZ, 18) = "=Q" & LineZ & "/P" & LineZ
                SNS += 1
                LineZ += 1
            End While
        End If
        oReader.Close()

        ' 加總
        oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 2))
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(LineZ, 1) = "Total"
        Ws.Cells(LineZ, 3) = "=SUM(C5:C" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 17)), Type:=xlFillDefault)
        Ws.Cells(LineZ, 18) = "=Q" & LineZ & "/P" & LineZ

        oRng = Ws.Range("A5", Ws.Cells(LineZ, 18))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous


        ' 第二頁
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        LineZ = 5

        For i As Int16 = TotalWeek To 1 Step -1
            Ws.Cells(4, 15 - i) = "W" & CWeek2 - i + 1
        Next

        oCommand.CommandText = "Select gem02"
        For i As Int16 = 1 To TotalWeek Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        oCommand.CommandText += ",sum(x1) as x1,sum(x2) as x2 from ( select abb05,gem02"

        For i As Int16 = 1 To TotalWeek Step 1
            Dim sWeek As Int16 = CWeek1 + i - 1
            Dim sYear As Int16 = CYear1
            oCommand.CommandText += ",(Case when azn02 = " & sYear & " and azn05 = " & sWeek & "  then (case when abb06 = 1 then abb07 else abb07 * -1 end) else 0 end) as t" & i
        Next
        oCommand.CommandText += ",0 as x1,0 as x2 from aea_file left join abb_file on aea03 = abb01 and aea04 = abb02 and aea05 = abb03 left join aag_file on aea05 = aag01 left join aae_file on aag223 = aae01 "
        oCommand.CommandText += "left join azn_file on aea02 = azn01 left join gem_file on abb05 = gem01 where aag223 is not null and abb05 <> 'D9999' and aea02 between to_date('" & MinD.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += MaxD.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "Select abb05,gem02"
        For i As Int16 = 1 To TotalWeek Step 1
            oCommand.CommandText += ",0"
        Next
        oCommand.CommandText += ",sum(case when abb06 = 1 then abb07 else abb07 * -1 end),0 "
        oCommand.CommandText += "from aea_file left join abb_file on aea03 = abb01 and aea04 = abb02 and aea05 = abb03 left join aag_file on aea05 = aag01 left join aae_file on aag223 = aae01 left join gem_file on abb05 = gem01 "
        oCommand.CommandText += "where aag223 is not null and abb05 <> 'D9999' and aea02 between to_date('"
        oCommand.CommandText += SDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & MaxD.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') group by abb05,gem02 "
        oCommand.CommandText += "union all "
        If TYear = 2019 Then
            oCommand.CommandText += "Select depno,gem02"
            For i As Int16 = 1 To TotalWeek Step 1
                oCommand.CommandText += ",0"
            Next
            oCommand.CommandText += ",0,sum(budget) from DAC_2019_Budget left join aag_file on acc1 = aag01 left join aae_file on aag223 = aae01 left join gem_file on depno = gem01 where aag223 is not null and month1 <= " & MaxD.Month & " group by depno,gem02"
        Else
            oCommand.CommandText += "Select tc_bud08,gem02"
            For i As Int16 = 1 To TotalWeek Step 1
                oCommand.CommandText += ",0"
            Next
            oCommand.CommandText += ",0,sum(tc_bud13) from tc_bud_file left join aag_file on tc_bud07 = aag01 left join aae_file on aag223 = aae01 left join gem_file on tc_bud08 = gem01 where tc_bud01 = 2 and tc_bud02 = "
            oCommand.CommandText += TYear & " and tc_bud03 <= " & MaxD.Month & " and aag23 is not null and tc_bud08 <> 'D9999' group by tc_bud08,gem02"
        End If
        oCommand.CommandText += " ) group by abb05,gem02 order by abb05"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            Dim SNS As Int16 = 1
            While oReader.Read()
                Ws.Cells(LineZ, 1) = SNS
                Ws.Cells(LineZ, 2) = oReader.Item("gem02")
                'For i As Int16 = 1 To 12 Step 1
                '    Ws.Cells(LineZ, i + 2) = oReader.Item(i)
                'Next
                For i As Int16 = TotalWeek To 1 Step -1
                    Ws.Cells(LineZ, 15 - i) = oReader.Item(oReader.FieldCount - 2 - i)
                Next
                Ws.Cells(LineZ, 15) = oReader.Item(oReader.FieldCount - 2)
                Ws.Cells(LineZ, 16) = oReader.Item(oReader.FieldCount - 1)
                Ws.Cells(LineZ, 17) = "=O" & LineZ & "-P" & LineZ
                Ws.Cells(LineZ, 18) = "=Q" & LineZ & "/P" & LineZ
                SNS += 1
                LineZ += 1
            End While
        End If
        oReader.Close()

        ' 加總
        oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 2))
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(LineZ, 1) = "Total"
        Ws.Cells(LineZ, 3) = "=SUM(C5:C" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 17)), Type:=xlFillDefault)
        Ws.Cells(LineZ, 18) = "=Q" & LineZ & "/P" & LineZ

        oRng = Ws.Range("A5", Ws.Cells(LineZ, 18))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous


        ' 第二頁  20191203 改為第三頁
        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        LineZ = 3
        oCommand.CommandText = "select aea05,aag02,aae02, gem02,azn05,aea02, aea03, abb04, (case when abb06 = 1 then abb07 else 0 end) as t1,(case when abb06 = 2 then abb07 else 0 end) as t2 "
        oCommand.CommandText += "from aea_file left join abb_file on aea03 = abb01 and aea04 = abb02 and aea05 = abb03 left join aag_file on aea05 = aag01 left join aae_file on aag223 = aae01 "
        oCommand.CommandText += "left join azn_file on aea02 = azn01 left join gem_file on abb05 = gem01 where aag223 is not null and aae01 <> '001' and abb05 <> 'D9999' and aea02 between to_date('" & MinD.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += MaxD.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') order by aea05"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = oReader.Item(i)
                Next
                LineZ += 1
                Label3.Text = LineZ
                Label3.Refresh()
            End While
        End If
        oReader.Close()
        oRng = Ws.Range("A3", Ws.Cells(LineZ - 1, 10))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
    End Sub
End Class