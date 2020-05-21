Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form135
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
    Dim tDate As Date
    Dim tDate1 As Date
    Dim LineZ As Integer = 0
    Dim DNP As String = String.Empty
    Dim ExchangeRate1 As Decimal = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form135_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        tYear = Me.DateTimePicker1.Value.Year
        tMonth = Me.DateTimePicker1.Value.Month
        tDate1 = Convert.ToDateTime(tYear & "/01/01")
        BackgroundWorker1.RunWorkerAsync()
        'ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        
        tDate = Convert.ToDateTime(tYear & "/" & tMonth & "/01").AddMonths(1).AddDays(-1)
        ' 再確定哪些部門要做
        oCommand.CommandText = "select distinct ina04,gem02 from ina_file,inb_File,gem_file where ina01 = inb01 and ina04 = gem01 and inapost = 'Y' and year(ina02) = " & tYear
        oCommand.CommandText += " and ina02 <= to_date('" & tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and inb05 not in (select jce02 from jce_file) and ina00 in (1,2,5,6) order by ina04"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Label3.Text = oReader.Item("gem02")
                xExcel = New Microsoft.Office.Interop.Excel.Application
                xWorkBook = xExcel.Workbooks.Add()
                Ws = xWorkBook.Sheets(1)
                Ws.Name = "Summary Table"
                AdjustExcelFormat()
                oCommand2.CommandText = "select imz02"
                For i As Int16 = 1 To tMonth Step 1
                    oCommand2.CommandText += ",sum(t" & i & ") as t" & i
                Next
                oCommand2.CommandText += " from ( select imz02"
                For i As Int16 = 1 To tMonth Step 1
                    oCommand2.CommandText += ",(case when month(ina02) = " & i & " then round(inb09 * inb08_fac * ccc23 * -1,2) else 0 end ) as t" & i
                Next
                oCommand2.CommandText += " from ina_file,inb_File,ccc_file,ima_file,imz_file where ina01 = inb01 and inapost = 'Y' and year(ina02) = " & tYear
                oCommand2.CommandText += " and ina02 <= to_date('" & tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and inb05 not in (select jce02 from jce_file) "
                oCommand2.CommandText += "and ina00 in (1,2,5,6) and ina04 = '" & oReader.Item("ina04") & "' and inb04 = ccc01 and ccc02 = year(ina02) and ccc03 = month(ina02) and inb04 = ima01 and ccc01 = ima01 and ima06 = imz01 ) group by imz02"
                oReader2 = oCommand2.ExecuteReader()
                If oReader2.HasRows() Then
                    While oReader2.Read()
                        For i As Int16 = 1 To oReader2.FieldCount Step 1
                            Ws.Cells(LineZ, i) = oReader2.Item(i - 1)
                        Next
                        Ws.Cells(LineZ, 14) = "=SUM(B" & LineZ & ":M" & LineZ & ")"
                        LineZ += 1
                    End While
                    Ws.Cells(LineZ, 1) = "Total"
                    Ws.Cells(LineZ, 2) = "=SUM(B2:B" & LineZ - 1 & ")"
                    oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 2))
                    oRng.Interior.Color = Color.Yellow
                    oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 14)), Type:=xlFillDefault)
                End If
                oReader2.Close()
                ' 劃線
                oRng = Ws.Range("A1", Ws.Cells(LineZ, 14))
                oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
                oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
                oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
                oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
                oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

                ' 置右
                oRng = Ws.Range("B2", Ws.Cells(LineZ, 14))
                oRng.HorizontalAlignment = xlRight

                ' 從當年一月到指定月份  --明細
                For i As Int16 = 1 To tMonth Step 1
                    If i > 2 Then
                        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                    Else
                        Ws = xWorkBook.Sheets(i + 1)
                    End If
                    If i < 10 Then
                        Ws.Name = tYear & "0" & i
                    Else
                        Ws.Name = tYear & i
                    End If
                    Ws.Activate()
                    AdjustExcelFormat1()
                    oCommand2.CommandText = "SELECT tlf907,tlf14,azf1.azf03,tlf19,gem02,tlf11,tlf06,tlf026,tlf01,ima02,ima06,imz02,imz39,aag1.aag02,cxi04,aag2.aag02,"
                    '190515 add by Brady
                    'oCommand2.CommandText += "ima021,tlf902,imd02,ccc23,tlf10,ccc23a,ccc23b,ccc23c,ccc23d,ccc23e,ccc23f,ccc23g,ccc23h,(ccc23 * tlf10),ina07,inbud03,inbud04,'',ima11,azf2.azf03,tlf20 "
                    oCommand2.CommandText += "ima021,tlf902,imd02,ccc23,(tlf10 * inb08_fac),ccc23a,ccc23b,ccc23c,ccc23d,ccc23e,ccc23f,ccc23g,ccc23h,(ccc23 * tlf10 * inb08_fac),ina07,inbud03,inbud04,'',ima11,azf2.azf03,tlf20 "
                    '190515 add by Brady END
                    oCommand2.CommandText += "FROM tlf_file left join azf_file azf1 on tlf14 = azf1.azf01 and azf1.azf02 = '2' left join gem_file on tlf19 = gem01 "
                    oCommand2.CommandText += "LEFT JOIN inb_file ON tlf905 = inb01 AND tlf906 = inb03 LEFT JOIN ina_file ON tlf905 = ina01 LEFT JOIN (sfa_file LEFT JOIN sfb_File ON sfa01 = sfb01 AND sfb87 != 'X') "
                    oCommand2.CommandText += "ON inbud04 = sfb01 AND sfa27 = tlf01 LEFT JOIN ima_file on tlf01 = ima01 LEFT JOIN imz_file ON ima06 = imz01 LEFT JOIN aag_file aag1 on imz39 = aag1.aag01 "
                    oCommand2.CommandText += "LEFT JOIN cxi_file ON cxi01 = tlf19 and ima11 = cxi02 LEFT JOIN aag_file aag2 on cxi04 = aag2.aag01 LEFT JOIN imd_file on tlf902 = imd01 "
                    oCommand2.CommandText += "left join ccc_File on tlf01 = ccc01 and ccc02 = year(ina02) and ccc03 = month(ina02) left join azf_file azf2 on ima11 = azf2.azf01 and azf2.azf02 = 'F' "
                    oCommand2.CommandText += "WHERE tlf01 = ima01 AND (tlf13='aimt301' or tlf13='aimt311' OR  tlf13='aimt303' or tlf13='aimt313') and tlf907 = -1 AND tlf902 NOT IN (SELECT jce02 FROM jce_file) "
                    oCommand2.CommandText += "and tlf06 between to_date('" & tDate1.AddMonths(i - 1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & tDate1.AddMonths(i).AddDays(-1).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
                    oCommand2.CommandText += "and tlf19 = '" & oReader.Item("ina04") & "'"
                    oReader2 = oCommand2.ExecuteReader()
                    If oReader2.HasRows() Then
                        While oReader2.Read()
                            For j As Int16 = 1 To oReader2.FieldCount Step 1
                                If j = 1 Then
                                    Ws.Cells(LineZ, 1) = "-1:出库"
                                Else
                                    Ws.Cells(LineZ, j) = oReader2.Item(j - 1)
                                End If

                            Next
                            LineZ += 1
                        End While
                        ' 劃線
                        oRng = Ws.Range("A1", Ws.Cells(LineZ - 1, 37))
                        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
                        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
                        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
                        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
                        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
                    End If
                    oReader2.Close()
                Next

                SaveExcel(oReader.Item("gem02"), oReader.Item("ina04"))
            End While
        End If
        oReader.Close()
        oConnection.Close()
        MsgBox("Finished")
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 100
        Ws.Columns.EntireColumn.ColumnWidth = 13.33
        Ws.Columns.EntireRow.RowHeight = 37.5
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.VerticalAlignment = xlCenter
        Ws.Columns.EntireColumn.Font.Name = "Arial"
        'Ws.Columns.EntireColumn.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        Ws.Columns.EntireColumn.NumberFormat = "#,##0.00_ ;-#,##0.00 "
        oRng = Ws.Range("A1", "N1")
        oRng.Font.Bold = True
        oRng = Ws.Range("B1", "M1")
        oRng.NumberFormat = "[$-en-US]mmm-yy;@"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireRow.Font.Bold = True
        Ws.Cells(1, 1) = "Description"

        For i As Int16 = 1 To 12 Step 1
            Ws.Cells(1, 1 + i) = tDate1.AddMonths(i - 1)
        Next
        Ws.Cells(1, 14) = "Total"
        LineZ = 2
    End Sub
    Private Sub SaveExcel(ByVal gem02 As String, ByVal gem01 As String)
        Dim SS As String = String.Empty
        If tMonth < 10 Then
            SS = "0" & tMonth
        Else
            SS = tMonth
        End If
        Dim SFN As String = "S:\A02_Finance_財務部\FN32-外挂报表\杂发费用明细表\" & tYear & SS & "\" & gem02 & ".xlsx"
        'Dim SFN As String = "C:\TEMP\" & tYear & SS & "_" & gem02 & ".xlsx"
        Ws.SaveAs(SFN, XlFileFormat.xlOpenXMLWorkbook)
        xWorkBook.Saved = True
        xWorkBook.Close()
        xExcel.Quit()
        'If oConnection.State = ConnectionState.Open Then
        Try
            'oConnection.Close()
            Module1.KillExcelProcess(OldExcel)
            MailSend(SFN, gem01, gem02)
            'MsgBox("Finished")
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        'End If
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 100
        Ws.Columns.EntireColumn.ColumnWidth = 17.44
        oRng = Ws.Range("I1", "I1")
        oRng.EntireColumn.NumberFormat = "@"
        oRng = Ws.Range("V1", "AD1")
        oRng.EntireColumn.NumberFormat = "#,##0.00_ ;-#,##0.00 "
        Ws.Cells(1, 1) = "入出库码"
        Ws.Cells(1, 2) = "异动原因"
        Ws.Cells(1, 3) = "说明内容"
        Ws.Cells(1, 4) = "部门编号"
        Ws.Cells(1, 5) = "部门名称"
        Ws.Cells(1, 6) = "单位"
        Ws.Cells(1, 7) = "单据日期"
        Ws.Cells(1, 8) = "单据编号"
        Ws.Cells(1, 9) = "料件编号"
        Ws.Cells(1, 10) = "品名"
        Ws.Cells(1, 11) = "分群码"
        Ws.Cells(1, 12) = "说明"
        Ws.Cells(1, 13) = "料件所属会计科目"
        Ws.Cells(1, 14) = "存货科目名称"
        Ws.Cells(1, 15) = "对方科目"
        Ws.Cells(1, 16) = "对方科目名称"
        Ws.Cells(1, 17) = "规格"
        Ws.Cells(1, 18) = "仓库"
        Ws.Cells(1, 19) = "仓库名称"
        Ws.Cells(1, 20) = "单价"
        Ws.Cells(1, 21) = "数量"
        Ws.Cells(1, 22) = "材料金额"
        Ws.Cells(1, 23) = "人工金额"
        Ws.Cells(1, 24) = "制造费用一"
        Ws.Cells(1, 25) = "加工费用"
        Ws.Cells(1, 26) = "制造费用二"
        Ws.Cells(1, 27) = "制造费用三"
        Ws.Cells(1, 28) = "制造费用四"
        Ws.Cells(1, 29) = "制造费用五"
        Ws.Cells(1, 30) = "总金额"
        Ws.Cells(1, 31) = "单头备注"
        Ws.Cells(1, 32) = "单身备注"
        Ws.Cells(1, 33) = "工单单号(财务)"
        Ws.Cells(1, 34) = "产出量"
        Ws.Cells(1, 35) = "其他分群码 三"
        Ws.Cells(1, 36) = "说明内容"
        Ws.Cells(1, 37) = "项目号码"
        LineZ = 2
    End Sub

    Public Sub MailSend(ByVal FileName As String, ByVal gem01 As String, ByVal gem02 As String)
        Dim MS As New System.Net.Mail.MailMessage
        Dim MA As New System.Net.Mail.MailAddress("action.server@action-composites.com.cn")
        MS.From = MA
        MS.Subject = "物料杂发数量及金额明细表-" & gem02
        'Dim mConnectionBuilder As New SqlClient.SqlConnectionStringBuilder
        Dim mConnection As New SqlClient.SqlConnection
        Dim mSQLS1 As New SqlClient.SqlCommand
        Dim mSQLReader As SqlClient.SqlDataReader
        mConnection.ConnectionString = Module1.OpenConnectionOfERPSUPPORT()
        If mConnection.State <> ConnectionState.Open Then
            mConnection.Open()
            mSQLS1.Connection = mConnection
            mSQLS1.CommandType = CommandType.Text
            mSQLS1.CommandTimeout = 600
        End If
        mSQLS1.CommandText = "SELECT * FROM MAILLIST WHERE CC = 'TO' AND ProgramName = 'DepartmentMISC' And DepartmentCode = '" & gem01 & "'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                MS.To.Add(mSQLReader.Item("MailAddress"))
            End While
        End If
        mSQLReader.Close()

        mSQLS1.CommandText = "SELECT * FROM MAILLIST WHERE CC = 'CC' AND ProgramName = 'DepartmentMISC' And DepartmentCode = '" & gem01 & "'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                MS.CC.Add(mSQLReader.Item("MailAddress"))
            End While
        End If
        mSQLReader.Close()

        mSQLS1.CommandText = "SELECT * FROM MAILLIST WHERE CC = 'BCC' AND ProgramName = 'DepartmentMISC' And DepartmentCode = '" & gem01 & "'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                MS.Bcc.Add(mSQLReader.Item("MailAddress"))
            End While
        End If

        mSQLReader.Close()
        mConnection.Close()
        MS.IsBodyHtml = True
        Dim MAM As New System.Net.Mail.Attachment(FileName)
        MS.Attachments.Add(MAM)
        ' 信件做好了
        Dim SMT As New System.Net.Mail.SmtpClient("smtp.action-composites.com.cn")
        SMT.UseDefaultCredentials = True
        'SMT.PickupDirectoryLocation = "C:\temp\ab"
        Dim UAP As New System.Net.NetworkCredential()
        UAP.UserName = "action.server@action-composites.com.cn"
        UAP.Password = "action@2017"
        SMT.Credentials = UAP

        Try
            SMT.Send(MS)
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
    End Sub
End Class