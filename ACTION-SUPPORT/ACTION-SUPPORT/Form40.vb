Imports Microsoft.Office.Interop.Excel.XlFileFormat
Public Class Form40
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
    Dim LastYear As Date
    Dim LastYearEnd As Date
    Dim RecordA As Decimal = 0
    Dim RecordB As Decimal = 0
    Dim LineZ As Integer = 0
    Dim ArrayX() As String = {"102010010027", "102010010002", "102010020012", "102020020009", "526000020005", "206010010002", "102010010034", "102010010022", "102010010019", "102010010009" + _
                              "202050020014", "515000020001", "203010020016", "508000020004", "206030020005", "205000010003", "102010010023", "102010010037", "102010020017", "201000020002" + _
                              "102010010003", "526000020006", "508000020009", "202050020016", "203010020018", "206020010001", "208000020010", "102010020012", "308040020027", "204000020002" + _
                              "527000020005", "315000020006", "204000020011", "203010020015", "203010020002", "201000020004", "204000020009", "207000010001", "315000020020", "204000020010" + _
                              "204000020035", "527000020001", "527000020006", "508000020012", "508000020013", "308040020003", "308040020031", "304000020001", "511000020003", "511000020001"}
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form40_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        TextBox1.Text = Now.Year
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
        TYear = TextBox1.Text
        DStartN = Convert.ToDateTime(TYear & "/01/01")
        DstartE = DStartN.AddMonths(1).AddDays(-1)
        LastYear = DStartN.AddYears(-1)
        LastYearEnd = DStartN.AddDays(-1)
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Payment_Cost_Index"
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
        Ws.Name = "直接材料"
        AdjustExcelFormat()
        For i As Integer = 1 To 12 Step 1
            RecordA = 0
            RecordB = 0
            For j As Integer = 0 To ArrayX.Length - 1 Step 1
                ' 分子 
                oCommand.CommandText = "select nvl(sum(pmn88),0) as t1 from pmm_file,pmn_file where pmm01 =  pmn01 and pmm18 = 'Y' and pmm04 between to_date('"
                oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmn04 = '"
                oCommand.CommandText += ArrayX(j).ToString() & "'"
                RecordA += oCommand.ExecuteScalar()

                ' 分母
                ' 要先算2015年最後一次單價
                Dim UnitPrice As Decimal = 0
                oCommand.CommandText = "select nvl(pmn31,0) from ( "
                oCommand.CommandText += "select pmn31 from pmm_file,pmn_file where pmm01 = pmn01 and pmm18 = 'Y' and pmm04 between to_date('"
                oCommand.CommandText += LastYear.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                oCommand.CommandText += LastYearEnd.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmn04 = '"
                oCommand.CommandText += ArrayX(j).ToString() & "' order by pmm04 desc ) where rownum = 1"
                UnitPrice = oCommand.ExecuteScalar()

                ' 算出分母
                oCommand.CommandText = "select nvl(sum(pmn20 * " & UnitPrice & "),0) as t1 from pmm_file,pmn_file where pmm01 =  pmn01 and pmm18 = 'Y' and pmm04 between to_date('"
                oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmn04 = '"
                oCommand.CommandText += ArrayX(j).ToString() & "'"
                RecordB += oCommand.ExecuteScalar()
            Next
            Ws.Cells(LineZ, 1 + i) = RecordA / RecordB
            '處理完到下一個
            DstartE = DstartE.AddDays(1).AddMonths(1).AddDays(-1)
        Next


        ' 第二頁
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        Ws.Name = "间接材料"
        AdjustExcelFormat()
        ' 回復日期
        DstartE = DStartN.AddMonths(1).AddDays(-1)
        For i As Integer = 1 To 12 Step 1
            RecordA = 0
            RecordB = 0
            ' 分子
            oCommand.CommandText = "select nvl(sum(pmn88),0) as t1 from pmm_file,pmn_file where pmm01 =  pmn01 and pmm18 = 'Y' and pmm04 between to_date('"
            oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmn04 not in (select distinct bmb03 from bmb_file) and pmn04 <> 'MISC' AND pmn04 not like '6%' and pmn04 not like '7%' and pmn04 not like '9%'"
            RecordA += oCommand.ExecuteScalar()

            ' 分母
            ' 要先算2015年最後一次單價
            oCommand.CommandText = "select distinct pmn04 from pmm_file,pmn_file where pmm01 =  pmn01 and pmm18 = 'Y' and pmm04 between to_date('"
            oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmn04 not in (select distinct bmb03 from bmb_file) and pmn04 <> 'MISC' AND pmn04 not like '6%' and pmn04 not like '7%' and pmn04 not like '9%'"
            oReader = oCommand.ExecuteReader()
            If oReader.HasRows() Then
                While oReader.Read()
                    Dim UnitPrice As Decimal = 0
                    oCommand.CommandText = "select nvl(pmn31,0) from ( "
                    oCommand.CommandText += "select pmn31 from pmm_file,pmn_file where pmm01 = pmn01 and pmm18 = 'Y' and pmm04 between to_date('"
                    oCommand.CommandText += LastYear.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                    oCommand.CommandText += LastYearEnd.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmn04 = '"
                    oCommand.CommandText += oReader.Item("pmn04") & "' order by pmm04 desc ) where rownum = 1"
                    UnitPrice = oCommand.ExecuteScalar()
                    ' 算出分母
                    oCommand.CommandText = "select nvl(sum(pmn20 * " & UnitPrice & "),0) as t1 from pmm_file,pmn_file where pmm01 =  pmn01 and pmm18 = 'Y' and pmm04 between to_date('"
                    oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                    oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmn04 = '"
                    oCommand.CommandText += oReader.Item("pmn04") & "'"
                    RecordB += oCommand.ExecuteScalar()
                End While
            End If
            oReader.Close()
            Ws.Cells(LineZ, 1 + i) = RecordA / RecordB
            '處理完到下一個
            DstartE = DstartE.AddDays(1).AddMonths(1).AddDays(-1)
        Next
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 35
        Ws.Cells(1, 1) = "Dongguan Action Composites LTD Co."
        Ws.Cells(2, 1) = "50 direct material purchase cost KPI"
        Ws.Cells(3, 1) = "month"
        Ws.Cells(4, 1) = "purchase cost KPI"
        For i As Int16 = 1 To 12 Step 1
            Ws.Cells(4, 1 + i) = GetMonthEnglish(i) & "-" & TYear
        Next
        oRng = Ws.Range("B5", "M5")
        oRng.NumberFormatLocal = "0.00%"
        LineZ = 5
    End Sub
End Class