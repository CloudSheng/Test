Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form93
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand3 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
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
    Dim TYM1 As String = String.Empty
    Dim PYM1 As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form93_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfERPSUPPORT()
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
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If Now.Month < 10 Then
            TextBox3.Text = Now.Year & "0" & Now.Month
            TextBox2.Text = Now.Year & "0" & Now.Month
        Else
            TextBox3.Text = Now.Year & Now.Month
            TextBox2.Text = Now.Year & Now.Month
        End If
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
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        pYear = Strings.Left(Start1, 4)
        pMonth = Strings.Right(Start1, 2)
        tYear = Strings.Left(End1, 4)
        tMonth = Strings.Right(End1, 2)

        If tMonth < 10 Then
            TYM1 = tYear & "0" & tMonth
        Else
            TYM1 = tYear & tMonth
        End If
        If pMonth < 10 Then
            PYM1 = pYear & "0" & pMonth
        Else
            PYM1 = pYear & pMonth
        End If
        Start2 = Convert.ToDateTime(pYear & "/" & pMonth & "/01")
        End2 = Convert.ToDateTime(tYear & "/" & tMonth & "/01").AddMonths(1).AddDays(-1)
        ExportToExcel()
        SaveExcel()
        'BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "项目案预算比较表"
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
                mConnection.Close()
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
        Ws.Name = "项目案 项目别"
        Ws.Activate()
        AdjustExcelFormat()
        oCommand.CommandText = "select distinct tc_bud10,tc_bud05,tc_bud06,tc_bud14 from hkacttest.tc_bud_file where tc_bud01 = '4' and  (case when tc_bud03 < 10 then tc_bud02 || '0' || tc_bud03 else tc_bud02 || tc_bud03 end) between '"
        oCommand.CommandText += PYM1 & "' AND '" & TYM1 & "'"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("tc_bud10")
                Ws.Cells(LineZ, 2) = oReader.Item("tc_bud05")
                Ws.Cells(LineZ, 3) = oReader.Item("tc_bud06")
                Ws.Cells(LineZ, 4) = oReader.Item("tc_bud14")
                Ws.Cells(LineZ, 5) = GetOrderAmount(oReader.Item("tc_bud10"))
                Ws.Cells(LineZ, 6) = GetSalesAct(oReader.Item("tc_bud10"), False)
                Ws.Cells(LineZ, 7) = GetMaterialAmount(oReader.Item("tc_bud10"), False)
                Ws.Cells(LineZ, 8) = GetLabor(oReader.Item("tc_bud10"), False)
                Ws.Cells(LineZ, 9) = GetTooling(oReader.Item("tc_bud10"), False)

                Ws.Cells(LineZ, 11) = "=SUM(G" & LineZ & ":J" & LineZ & ")"
                Ws.Cells(LineZ, 12) = "=F" & LineZ & "-K" & LineZ
                Ws.Cells(LineZ, 13) = GetSalesAct(oReader.Item("tc_bud10"), True)
                Ws.Cells(LineZ, 14) = GetMaterialAmount(oReader.Item("tc_bud10"), True)
                Ws.Cells(LineZ, 15) = GetLabor(oReader.Item("tc_bud10"), True)
                Ws.Cells(LineZ, 16) = GetTooling(oReader.Item("tc_bud10"), True)

                Ws.Cells(LineZ, 18) = "=SUM(N" & LineZ & ":Q" & LineZ & ")"
                Ws.Cells(LineZ, 19) = "=M" & LineZ & "-R" & LineZ
                LineZ += 1
            End While
            ' 加總
            Ws.Cells(LineZ, 6) = "=SUM(F7:F" & LineZ - 1 & ")"
            oRng = Ws.Range(Ws.Cells(LineZ, 6), Ws.Cells(LineZ, 6))
            oRng.AutoFill(Destination:=Ws.Range("F" & LineZ & ":S" & LineZ), Type:=xlFillDefault)
            oRng = Ws.Range("E7", Ws.Cells(LineZ, 19))
            oRng.NumberFormatLocal = "#,##0_ "
        End If
        oReader.Close()


        '第二頁
        Ws = xWorkBook.Sheets(2)
        Ws.Name = "项目案销售 客户别"
        Ws.Activate()
        AdjustExcelFormat1()
        oCommand.CommandText = "select tc_bud05,ofa032,tc_bud14,nvl(sum(ofb14),0) as t1,nvl(sum(tc_bud13),0) as t2 from hkacttest.tc_bud_file "
        oCommand.CommandText += "left join hkacttest.oeb_file on tc_bud10 = oeb41 left join hkacttest.ofb_file on ofb31 = oeb01 and ofb32 = oeb03 "
        oCommand.CommandText += "left join hkacttest.ofa_file on ofb01 = ofa01 and ofaconf = 'Y' and ofa02 between to_date('"
        oCommand.CommandText += Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += End2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') where tc_bud01 = '4' and  (case when tc_bud03 < 10 then tc_bud02 || '0' || tc_bud03 else tc_bud02 || tc_bud03 end) between '"
        oCommand.CommandText += PYM1 & "' AND '" & TYM1 & "' group by tc_bud05,ofa032,tc_bud14"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("tc_bud05")
                Ws.Cells(LineZ, 2) = oReader.Item("ofa032")
                Ws.Cells(LineZ, 3) = oReader.Item("tc_bud14")
                Ws.Cells(LineZ, 4) = oReader.Item("t1")
                Ws.Cells(LineZ, 5) = oReader.Item("t2")
                Ws.Cells(LineZ, 6) = "=D" & LineZ & "-E" & LineZ
                LineZ += 1
            End While
            Ws.Cells(LineZ, 2) = "Grand Total"
            Ws.Cells(LineZ, 4) = "=SUM(D6:D" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 5) = "=SUM(E6:E" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 6) = "=SUM(F6:F" & LineZ - 1 & ")"
            oRng = Ws.Range("D6", Ws.Cells(LineZ, 6))
            oRng.NumberFormatLocal = "#,##0_ "
        End If
        oReader.Close()

        '第三頁
        Ws = xWorkBook.Sheets(3)
        Ws.Name = "项目案销售 客户代表"
        Ws.Activate()
        AdjustExcelFormat2()
        oCommand.CommandText = "select tc_bud06,tc_bud14,nvl(sum(ofb14),0) as t1,nvl(sum(tc_bud13),0) as t2 from hkacttest.tc_bud_file "
        oCommand.CommandText += "left join hkacttest.oeb_file on tc_bud10 = oeb41 left join hkacttest.ofb_file on ofb31 = oeb01 and ofb32 = oeb03 "
        oCommand.CommandText += "left join hkacttest.ofa_file on ofb01 = ofa01 and ofaconf = 'Y' and ofa02 between to_date('"
        oCommand.CommandText += Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += End2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') where tc_bud01 = '4' and  (case when tc_bud03 < 10 then tc_bud02 || '0' || tc_bud03 else tc_bud02 || tc_bud03 end) between '"
        oCommand.CommandText += PYM1 & "' AND '" & TYM1 & "' group by tc_bud06,tc_bud14"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("tc_bud06")
                Ws.Cells(LineZ, 2) = oReader.Item("tc_bud14")
                Ws.Cells(LineZ, 3) = oReader.Item("t1")
                Ws.Cells(LineZ, 4) = oReader.Item("t2")
                Ws.Cells(LineZ, 5) = "=C" & LineZ & "-D" & LineZ
                LineZ += 1
            End While
            Ws.Cells(LineZ, 3) = "=SUM(C6:C" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 4) = "=SUM(D6:D" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 5) = "=SUM(E6:E" & LineZ - 1 & ")"
            oRng = Ws.Range("C6", Ws.Cells(LineZ, 5))
            oRng.NumberFormatLocal = "#,##0_ "
        End If
        oReader.Close()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.ColumnWidth = 17
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 40.89
        Ws.Cells(1, 1) = "Year:" & tYear
        Ws.Cells(2, 1) = "Month:" & Start2.ToString("yyyy/MM/dd") & "-" & End2.ToString("yyyy/MM/dd")
        Ws.Cells(3, 1) = "Action Composite Technology Limited"
        oRng = Ws.Range("A4", "S4")
        oRng.Interior.Color = Color.FromArgb(146, 208, 80)
        Ws.Cells(4, 1) = "Tooling overview  in USD"
        oRng = Ws.Range("A5", "A6")
        oRng.Merge()
        Ws.Cells(5, 1) = "Project no 项目号"
        oRng = Ws.Range("B5", "B6")
        oRng.Merge()
        Ws.Cells(5, 2) = "账款客户"
        oRng = Ws.Range("C5", "C6")
        oRng.Merge()
        Ws.Cells(5, 3) = "KAM 客户代表"
        oRng = Ws.Range("D5", "D6")
        oRng.Merge()
        Ws.Cells(5, 4) = "Currency"
        Ws.Cells(5, 5) = "Project order amt"
        Ws.Cells(6, 5) = "订单金额（原币）"
        Ws.Cells(5, 6) = "Current当期"
        Ws.Cells(5, 7) = "Current当期"
        Ws.Cells(5, 8) = "Current当期"
        Ws.Cells(5, 9) = "Current当期"
        Ws.Cells(5, 10) = "Current当期"
        Ws.Cells(5, 11) = "Current当期"
        Ws.Cells(5, 12) = "Current当期"
        Ws.Cells(6, 6) = "Sales-Actual"
        Ws.Cells(6, 7) = "Material"
        Ws.Cells(6, 8) = "Labor"
        Ws.Cells(6, 9) = "toolling"
        Ws.Cells(6, 10) = "other"
        Ws.Cells(6, 11) = "expenditure支出"
        Ws.Cells(6, 12) = "Margin"
        Ws.Cells(5, 13) = "YTD 历史累计"
        Ws.Cells(5, 14) = "YTD 历史累计"
        Ws.Cells(5, 15) = "YTD 历史累计"
        Ws.Cells(5, 16) = "YTD 历史累计"
        Ws.Cells(5, 17) = "YTD 历史累计"
        Ws.Cells(5, 18) = "YTD 历史累计"
        Ws.Cells(5, 19) = "YTD 历史累计"
        Ws.Cells(6, 13) = "Sales-Actual"
        Ws.Cells(6, 14) = "Material"
        Ws.Cells(6, 15) = "Labor"
        Ws.Cells(6, 16) = "toolling"
        Ws.Cells(6, 17) = "other"
        Ws.Cells(6, 18) = "expenditure支出"
        Ws.Cells(6, 19) = "Margin"
        oRng = Ws.Range("A5", "S6")
        oRng.HorizontalAlignment = xlCenter
        LineZ = 7
    End Sub
    Private Function GetOrderAmount(ByVal oea46 As String)
        oCommand2.CommandText = "SELECT nvl(sum(oea61),0) FROM hkacttest.oea_File where oea46 = '" & oea46 & "'"
        Dim OA As Decimal = oCommand2.ExecuteScalar()
        Return OA
    End Function
    Private Function GetMaterialAmount(ByVal tlf20 As String, ByVal History As Boolean)
        If History = True Then
            oCommand2.CommandText = "SELECT Round(nvl(sum(tlf21/azj041),0),3) FROM TLF_FILE,azj_file WHERE TLF20 = '"
            oCommand2.CommandText += tlf20 & "' AND TLF13 IN ('aimt301','aimt311') and azj01 = 'USD' and azj02 = year(tlf06) || (case when month(tlf06) < 10 then '0' || month(tlf06) else to_char(month(tlf06)) end) "
        Else
            oCommand2.CommandText = "SELECT Round(nvl(sum(tlf21/azj041),0),3) FROM TLF_FILE,azj_file WHERE TLF20 = '"
            oCommand2.CommandText += tlf20 & "' AND TLF13 IN ('aimt301','aimt311') and azj01 = 'USD' and azj02 = year(tlf06) || (case when month(tlf06) < 10 then '0' || month(tlf06) else to_char(month(tlf06)) end) "
            oCommand2.CommandText += "and tlf06 between to_date('" & Start2.ToString("yyyy/MM/dd")
            oCommand2.CommandText += "','yyyy/mm/dd') and to_date('" & End2.ToString("yyyy/MM/dd")
            oCommand2.CommandText += "','yyyy/mm/dd') "
        End If
        
        Try
            Dim MA As Decimal = oCommand2.ExecuteScalar()
            Return MA
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Private Function GetSalesAct(ByVal oga46 As String, ByVal History As Boolean)
        If History = True Then
            oCommand2.CommandText = "select nvl(sum(ofb14),0) from hkacttest.ofb_file,hkacttest.oeb_file,hkacttest.ofa_file where ofb01 = ofa01 and ofaconf = 'Y' and  ofb31 = oeb01 and ofb32 = oeb03 and oeb41 = '"
            oCommand2.CommandText += oga46 & "'"
        Else
            oCommand2.CommandText = "select nvl(sum(ofb14),0) from hkacttest.ofb_file,hkacttest.oeb_file,hkacttest.ofa_file where ofb01 = ofa01 and ofaconf = 'Y' and  ofb31 = oeb01 and ofb32 = oeb03 and oeb41 = '"
            oCommand2.CommandText += oga46 & "' and ofa02 between to_date('" & Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & End2.ToString("yyyy/MM/dd")
            oCommand2.CommandText += "','yyyy/mm/dd') "
        End If
        Dim SA As Decimal = oCommand2.ExecuteScalar()
        Return SA
    End Function
    Private Function GetLabor(ByVal ej As String, ByVal History As Boolean)
        Dim ToV As Decimal = 0
        If History = True Then
            mSQLS1.CommandText = "select year(edate) as t1,month(edate) as t2,sum(ehour * 18.25) as t3 from ProjectHR  where EProject = '"
            mSQLS1.CommandText += ej & "' group by year(edate),month(edate) order by t1,t2"
        Else
            mSQLS1.CommandText = "select year(edate) as t1,month(edate) as t2,sum(ehour * 18.25) as t3 from ProjectHR  where EProject = '"
            mSQLS1.CommandText += ej & "' and edate between '" & Start2.ToString("yyyy/MM/dd")
            mSQLS1.CommandText += "' and '" & End2.ToString("yyyy/MM/dd") & "' group by year(edate),month(edate) order by t1,t2"
        End If
        mSQLReader = mSQLS1.ExecuteReader()
        'oReader2 = oCommand2.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Dim Ex1 As Decimal = GetExchangeRate(mSQLReader.Item("t1"), mSQLReader.Item("t2"))
                ToV += mSQLReader.Item("t3") / Ex1
            End While
        End If
        mSQLReader.Close()
        Return ToV
    End Function
    Private Function GetExchangeRate(ByVal sYear As Integer, ByVal sMonth As Integer)
        Dim TYM2 As String = String.Empty
        If sMonth < 10 Then
            TYM2 = sYear & "0" & sMonth
        Else
            TYM2 = sYear & sMonth
        End If
        oCommand3.CommandText = "SELECT azj041 FROM azj_file WHERE azj01 = 'USD' AND azj02 = '" & TYM2 & "'"
        Dim EX1 As Decimal = oCommand3.ExecuteScalar()
        If EX1 = 0 Then
            Return 1
        Else
            Return EX1
        End If
    End Function
    Private Function GetTooling(ByVal tlf20 As String, ByVal History As Boolean)
        If History = True Then
            oCommand2.CommandText = "select Round(nvl(sum(pmn88 * pmm42 /azj041),0),3) from tlf_file,ima_file,azj_file,pmn_file,pmm_file  where tlf01 = ima01 and ima06 = '105' and  tlf20 = '"
            oCommand2.CommandText += tlf20 & "' and tlf907 = 1 and tlf13 = 'apmt150' and azj01 = 'USD' and azj02 = year(tlf06) || (case when month(tlf06) < 10 then '0' || month(tlf06) else to_char(month(tlf06)) end) "
            oCommand2.CommandText += "and tlf62 = pmn01 and tlf01 = pmn04 and tlf20 = pmn122 and pmn01 = pmm01 "
        Else
            oCommand2.CommandText = "select Round(nvl(sum(pmn88 * pmm42 /azj041),0),3) from tlf_file,ima_file,azj_file,pmn_file,pmm_file  where tlf01 = ima01 and ima06 = '105' and  tlf20 = '"
            oCommand2.CommandText += tlf20 & "' and tlf907 = 1 and tlf13 = 'apmt150' and azj01 = 'USD' and azj02 = year(tlf06) || (case when month(tlf06) < 10 then '0' || month(tlf06) else to_char(month(tlf06)) end) "
            oCommand2.CommandText += "and tlf62 = pmn01 and tlf01 = pmn04 and tlf20 = pmn122 and pmn01 = pmm01 and tlf06 between to_date('"
            oCommand2.CommandText += Start2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            oCommand2.CommandText += End2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        End If
        Dim ToV As Decimal = oCommand2.ExecuteScalar()
        Return ToV
    End Function
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.ColumnWidth = 17
        oRng = Ws.Range("A1", "B1")
        oRng.EntireColumn.ColumnWidth = 40.89
        Ws.Cells(1, 1) = "Year:" & tYear
        Ws.Cells(2, 1) = "Month:" & Start2.ToString("yyyy/MM/dd") & "-" & End2.ToString("yyyy/MM/dd")
        Ws.Cells(3, 1) = "Action Composite Technology Limited"
        oRng = Ws.Range("A4", "F4")
        oRng.Interior.Color = Color.FromArgb(146, 208, 80)
        Ws.Cells(4, 1) = "Project revenue comparision (Actual & Budget)"
        Ws.Cells(5, 1) = "账款客户"
        Ws.Cells(5, 2) = "收款客户"
        Ws.Cells(5, 3) = "Currency"
        Ws.Cells(5, 4) = "Sales-Actual"
        Ws.Cells(5, 5) = "sales-Budget"
        Ws.Cells(5, 6) = "实际-预算"
        
        oRng = Ws.Range("A5", "F5")
        oRng.HorizontalAlignment = xlCenter
        LineZ = 6
    End Sub
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.ColumnWidth = 17
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 40.89
        Ws.Cells(1, 1) = "Year:" & tYear
        Ws.Cells(2, 1) = "Month:" & Start2.ToString("yyyy/MM/dd") & "-" & End2.ToString("yyyy/MM/dd")
        Ws.Cells(3, 1) = "Action Composite Technology Limited"
        oRng = Ws.Range("A4", "E4")
        oRng.Merge()
        oRng.Interior.Color = Color.FromArgb(146, 208, 80)
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(4, 1) = "Project revenue comparision (Actual & Budget) in USD"
        Ws.Cells(5, 1) = "KAM 客户代表"
        Ws.Cells(5, 2) = "Currency"
        Ws.Cells(5, 3) = "Sales-Actual"
        Ws.Cells(5, 4) = "sales-Budget"
        Ws.Cells(5, 5) = "实际-预算"

        oRng = Ws.Range("A5", "E5")
        oRng.HorizontalAlignment = xlCenter
        LineZ = 6
    End Sub
End Class