Imports Microsoft.Office.Interop.Excel.XlFileFormat
Public Class Form29
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim TYear As String = String.Empty
    Dim TMonth As String = String.Empty
    Dim DStartN As Date
    Dim DstartE As Date
    'Dim PYear As String = String.Empty
    'Dim PMonth As String = String.Empty
    Dim RecordA As Decimal = 0
    Dim RecordB As Decimal = 0
    'Dim RecordC As Decimal = 0
    'Dim RecordD As Decimal = 0
    'Dim RecordE As Decimal = 0
    'Dim RecordF As Decimal = 0
    'Dim RecordG As Decimal = 0
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")


    Private Sub Form29_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        'If Now.Month < 10 Then
        'TextBox1.Text = Now.Year & "0" & Now.Month
        'Else
        'TextBox1.Text = Now.Year & Now.Month
        TextBox1.Text = Now.Year
        'End If
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
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        'TYear = Strings.Left(TextBox1.Text, 4)
        TYear = TextBox1.Text
        DStartN = Convert.ToDateTime(TYear & "/01/01")
        DstartE = DStartN.AddMonths(1).AddDays(-1)
        BackgroundWorker1.RunWorkerAsync()
        'TMonth = Strings.Right(TextBox1.Text, 2)
        'If TMonth > 12 Or TMonth < 1 Then
        'MsgBox("Error Month Data")
        'Return
        'End If
        'If TMonth = 1 Then
        'PYear = TYear - 1
        'PMonth = 12
        'Else
        'PMonth = TMonth - 1
        'End If
        'oCommand.CommandText = "select nvl(abs(sum(aah04-aah05)),0) from aah_file where aah02 = " & PYear & " and aah03 between 0 and " & PMonth & " and aah01 between '220201' and '220203'"
        'RecordA = oCommand.ExecuteScalar()
        'oCommand.CommandText = "select nvl(abs(sum(aah04-aah05)),0) from aah_file where aah02 = " & TYear & " and aah03 between 0 and " & TMonth & " and aah01 between '220201' and '220203'"
        'RecordB = oCommand.ExecuteScalar()
        'RecordC = (RecordA / 12 + RecordB / 12) / 2
        'If RecordC = 0 Then
        '    MsgBox("RecordC equals 0")
        '    Return
        'End If
        'oCommand.CommandText = "select nvl(abs(sum(aah04-aah05)),0) from aah_file where aah02 = " & TYear & " and aah03 = " & TMonth & " and aah01 between '640101' and '640102'"
        'RecordD = oCommand.ExecuteScalar()
        'oCommand.CommandText = "select nvl(abs(sum(aah04)),0) from aah_file where aah02 = " & TYear & " and aah03 = " & TMonth & " and aah01 IN ('160101','180102','180103','180104','180106')"
        'RecordE = oCommand.ExecuteScalar()
        'RecordF = (RecordD + RecordE) / RecordC
        'If RecordF = 0 Then
        '    MsgBox("RecordF equals 0")
        '    Return
        'End If
        'RecordG = 360 / RecordF
        'xExcel = New Microsoft.Office.Interop.Excel.Application
        'xWorkBook = xExcel.Workbooks.Add()
        'Ws = xWorkBook.Sheets(1)
        'Ws.Activate()
        'Ws.Cells(1, 1) = RecordG

        'Ws = xWorkBook.Sheets(2)
        'Ws.Activate()
        'Ws.Cells(1, 1) = "厂商编号"
        'Ws.Cells(1, 2) = "厂商简称"
        'Ws.Cells(1, 3) = "付款条件"
        'LineZ = 2
        'oCommand.CommandText = "select pmc01,pmc03,pma02 from pmc_file,pma_File where pmc17 = pma01 and pmcacti = 'Y'"
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        Ws.Cells(LineZ, 1) = oReader.Item("pmc01")
        '        Ws.Cells(LineZ, 2) = oReader.Item("pmc03")
        '        Ws.Cells(LineZ, 3) = oReader.Item("pma02")
        '        LineZ += 1
        '    End While
        'End If
        'oReader.Close()
        'oConnection.Close()
        'SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Payment_Term"
        SaveFileDialog1.DefaultExt = ".xls"
        Dim SON As DialogResult = SaveFileDialog1.ShowDialog()
        If SON = DialogResult.OK Then
            Dim SFN As String = SaveFileDialog1.FileName
            Ws.SaveAs(SFN, XlFileFormat.xlExcel12)
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
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        AdjustExcelFormat1()
        For i As Integer = 1 To 12 Step 1
            ' 先取前20 供應商
            'oCommand.CommandText = "select * from ( "
            'oCommand.CommandText += "select sum(pmn88) as t1,pmm09 from pmm_file,pmn_file where pmm01 =  pmn01 and pmm18 = 'Y' and pmm04 between to_date('"
            'oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            'oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') group by pmm09 order by t1 desc ) where rownum <=20 order by t1 desc"
            oCommand.CommandText = "select * from ( "
            oCommand.CommandText += "select sum(apg05) as t1,apf03 from apf_file left join apg_file on apf01 = apg01 where apf02 between to_date('"
            oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and apf41 = 'Y'  and apf00 <> '34' group by apf03 order by t1 desc ) where rownum <=20 order by t1 desc"
            oReader = oCommand.ExecuteReader()
            RecordA = 0
            RecordB = 0
            If oReader.HasRows() Then
                While oReader.Read()
                    RecordA += vendorpaydate(oReader.Item("apf03"))
                    RecordB += standarddate(oReader.Item("apf03"))
                End While
            End If
            oReader.Close()
            'oCommand.CommandText = "select sum(pmn88 * 60) as t1 from pmm_file,pmn_file where pmm01 =  pmn01 and pmm18 = 'Y' and pmm04 between to_date('"
            'oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            'oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
            'RecordB = oCommand.ExecuteScalar()
            If RecordB <> 0 Then
                'Ws.Cells(LineZ, 1 + i) = (RecordA / 20) / RecordB '* 100
                Ws.Cells(LineZ, 1 + i) = RecordA / RecordB
            End If
            '處理完到下一個
            DstartE = DstartE.AddDays(1).AddMonths(1).AddDays(-1)
            'DstartE.AddMonths(1)
            'DstartE.AddDays(-1)
        Next
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "payment day"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 35
        Ws.Cells(1, 1) = "Dongguan Action Composites LTD Co."
        Ws.Cells(2, 1) = "top 20 vendors payment days KPI"
        Ws.Cells(3, 1) = "month"
        Ws.Cells(4, 1) = "payment days"
        For i As Int16 = 1 To 12 Step 1
            Ws.Cells(4, 1 + i) = GetMonthEnglish(i) & "-" & TYear
        Next
        oRng = Ws.Range("B5", "M5")
        oRng.NumberFormatLocal = "0.00%"
        LineZ = 5
    End Sub
    Private Function vendorpaydate(ByVal pmm09 As String)
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "select nvl(sum((apf02-apa02) * apg05),0) from apf_file left join apg_file on apf01 = apg01 "
        oCommander99.CommandText += "left join apa_file on apg04 = apa01 where apf02 between to_date('"
        oCommander99.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommander99.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and apf41 = 'Y' and apf00 <> '34' and apf03 = '" & pmm09 & "'"
        Dim MF As Decimal = oCommander99.ExecuteScalar()
        Return MF
    End Function
    Private Function standarddate(ByVal pmm09 As String)
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "select nvl(sum(60 * apg05),0) from apf_file left join apg_file on apf01 = apg01 "
        oCommander99.CommandText += "where apf02 between to_date('"
        oCommander99.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommander99.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and apf41 = 'Y' and apf00 <> '34' and apf03 = '" & pmm09 & "'"
        Dim MF As Decimal = oCommander99.ExecuteScalar()
        Return MF
    End Function
End Class