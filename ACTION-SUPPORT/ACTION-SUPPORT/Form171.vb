Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form171
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
    Dim LineZ As Integer = 0
    Dim tTime As Date
    Dim tYear As Int16 = 0
    Dim pYear As Int16 = 0
    Dim fTime As Date
    Dim AzjYM As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form171_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        NumericUpDown1.Value = Now.Year()
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
        mConnection.ConnectionString = Module1.OpenConnectionOfERPSUPPORT()
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        tYear = NumericUpDown1.Value
        pYear = tYear - 1
        tTime = Convert.ToDateTime(tYear & "/01/01")
        fTime = Convert.ToDateTime(tYear & "/12/31")


        mSQLS1.CommandText = "DELETE USD_ExchangeRate"
        Try
            mSQLS1.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

        oCommand.CommandText = "select AZJ041 from azj_file WHERE azj01 = 'USD' AND AZJ02 LIKE '" & tYear & "%' ORDER BY AZJ02"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            Dim M1 As Int16 = 1
            While oReader.Read()

                mSQLS1.CommandText = "INSERT INTO USD_ExchangeRate VALUES (" & M1 & "," & oReader.Item("azj041") & ")"
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
                M1 += 1
            End While
        End If
        oReader.Close()

        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Project P&L Report"
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
        If mConnection.State = ConnectionState.Open Then
            mConnection.Close()
        End If
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        Dim xPath As String = "C:\temp\Project P&L Sample.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        AdjustExcelFormat()
        LineZ = 4

        oCommand.CommandText = "select pja01,pja02,(case when pjaud02 = 1 then 'US/JP' when pjaud02 = 2 then 'China' when pjaud02 = 3 then 'Europe' when pjaud02 = 4 then 'Vietnam' else pjaud02 end) as x1 ,pjaud04,"
        oCommand.CommandText += "nvl((pjaud08 * 0.1479),0) as d1,nvl((pjaud07 * 35 * 0.1479),0) as d2, nvl((pja33 + pja34) * 0.1479,0) as d3,(pjaud09 / 100) as pjaud09,Moldpaid,RD, MC, ACAcost from pja_file "
        oCommand.CommandText += "left join dac2018_fixproject on pja01 = projectcode order by pja01"
        oReader = oCommand.ExecuteReader()
        Dim CopyReader1 As Oracle.ManagedDataAccess.Client.OracleDataReader = oCommand.ExecuteReader()
        Dim CopyReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader = oCommand.ExecuteReader()
        Dim CopyReader3 As Oracle.ManagedDataAccess.Client.OracleDataReader = oCommand.ExecuteReader()

        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("pja01")
                Ws.Cells(LineZ, 2) = oReader.Item("pja02")
                Ws.Cells(LineZ, 3) = oReader.Item("X1")
                Ws.Cells(LineZ, 6) = oReader.Item("pjaud04")
                Ws.Cells(LineZ, 10) = oReader.Item("D1")
                Ws.Cells(LineZ, 11) = oReader.Item("D2")
                Ws.Cells(LineZ, 12) = oReader.Item("D3")
                'Ws.Cells(LineZ, 14) = "=SUM(J" & LineZ & ":M" & LineZ & ")"
                Ws.Cells(LineZ, 15) = oReader.Item("pjaud09")
                Ws.Cells(LineZ, 16) = oReader.Item("MoldPaid")
                Ws.Cells(LineZ, 17) = oReader.Item("RD")
                Ws.Cells(LineZ, 18) = oReader.Item("MC")
                Ws.Cells(LineZ, 19) = oReader.Item("ACACost")
                'Ws.Cells(LineZ, 20) = "=SUM(P" & LineZ & ":S" & LineZ & ")"
                GetMoldFee(oReader.Item("pja01"))
                GetRDHour(oReader.Item("pja01"))
                GetMCData(oReader.Item("pja01"))
                LineZ += 1
                Me.Label2.Text = "Page1 " & LineZ
            End While
            Ws.Cells(LineZ, 1) = "Sub-Total"
            Ws.Cells(LineZ, 9) = "=SUBTOTAL(9,I4:I" & LineZ - 1 & ")"
            oRng = Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9))
            oRng.AutoFill(Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 85)), Type:=xlFillDefault)

            oRng = Ws.Range("A4", Ws.Cells(LineZ, 85))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        oReader.Close()

        ' 第二頁完工比例
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustExcelFormat2()
        LineZ = 3
        If CopyReader1.HasRows() Then
            While CopyReader1.Read()
                Ws.Cells(LineZ, 1) = CopyReader1.Item("pja01")
                Ws.Cells(LineZ, 2) = CopyReader1.Item("pja02")
                Ws.Cells(LineZ, 3) = CopyReader1.Item("X1")
                Ws.Cells(LineZ, 4) = CopyReader1.Item("pjaud04")
                Ws.Cells(LineZ, 6) = CopyReader1.Item("pjaud09")
                Ws.Cells(LineZ, 7) = "=IFERROR(('2019'!Q" & LineZ + 1 & "+'2019'!V" & LineZ + 1 & ")/'2019'!K" & LineZ + 1 & ",)"
                Ws.Cells(LineZ, 8) = "=IFERROR(('2019'!$Q" & LineZ + 1 & "+'2019'!$V" & LineZ + 1 & "+'2019'!$AA" & LineZ + 1 & ")/'2019'!$K" & LineZ + 1 & ",)"
                Ws.Cells(LineZ, 9) = "=IFERROR(('2019'!$Q" & LineZ + 1 & "+'2019'!$V" & LineZ + 1 & "+'2019'!$AA" & LineZ + 1 & "+'2019'!$AF" & LineZ + 1 & ")/'2019'!$K" & LineZ + 1 & ",)"
                Ws.Cells(LineZ, 10) = "=IFERROR(('2019'!$Q" & LineZ + 1 & "+'2019'!$V" & LineZ + 1 & "+'2019'!$AA" & LineZ + 1 & "+'2019'!$AF" & LineZ + 1 & "+'2019'!$AK" & LineZ + 1 & ")/'2019'!$K" & LineZ + 1 & ",)"
                Ws.Cells(LineZ, 11) = "=IFERROR(('2019'!$Q" & LineZ + 1 & "+'2019'!$V" & LineZ + 1 & "+'2019'!$AA" & LineZ + 1 & "+'2019'!$AF" & LineZ + 1 & "+'2019'!$AK" & LineZ + 1 & "+'2019'!$AP" & LineZ + 1 & ")/'2019'!$K" & LineZ + 1 & ",)"
                Ws.Cells(LineZ, 12) = "=IFERROR(('2019'!$Q" & LineZ + 1 & "+'2019'!$V" & LineZ + 1 & "+'2019'!$AA" & LineZ + 1 & "+'2019'!$AF" & LineZ + 1 & "+'2019'!$AK" & LineZ + 1 & "+'2019'!$AP" & LineZ + 1 & "+'2019'!$AU" & LineZ + 1 & ")/'2019'!$K" & LineZ + 1 & ",)"
                Ws.Cells(LineZ, 13) = "=IFERROR(('2019'!$Q" & LineZ + 1 & "+'2019'!$V" & LineZ + 1 & "+'2019'!$AA" & LineZ + 1 & "+'2019'!$AF" & LineZ + 1 & "+'2019'!$AK" & LineZ + 1 & "+'2019'!$AP" & LineZ + 1 & "+'2019'!$AU" & LineZ + 1 & "+'2019'!$AZ" & LineZ + 1 & ")/'2019'!$K" & LineZ + 1 & ",)"
                Ws.Cells(LineZ, 14) = "=IFERROR(('2019'!$Q" & LineZ + 1 & "+'2019'!$V" & LineZ + 1 & "+'2019'!$AA" & LineZ + 1 & "+'2019'!$AF" & LineZ + 1 & "+'2019'!$AK" & LineZ + 1 & "+'2019'!$AP" & LineZ + 1 & "+'2019'!$AU" & LineZ + 1 & "+'2019'!$AZ" & LineZ + 1 & "+'2019'!$BE" & LineZ + 1 & ")/'2019'!$K" & LineZ + 1 & ",)"
                Ws.Cells(LineZ, 15) = "=IFERROR(('2019'!$Q" & LineZ + 1 & "+'2019'!$V" & LineZ + 1 & "+'2019'!$AA" & LineZ + 1 & "+'2019'!$AF" & LineZ + 1 & "+'2019'!$AK" & LineZ + 1 & "+'2019'!$AP" & LineZ + 1 & "+'2019'!$AU" & LineZ + 1 & "+'2019'!$AZ" & LineZ + 1 & "+'2019'!$BE" & LineZ + 1 & "+'2019'!$BJ" & LineZ + 1 & ")/'2019'!$K" & LineZ + 1 & ",)"
                Ws.Cells(LineZ, 16) = "=IFERROR(('2019'!$Q" & LineZ + 1 & "+'2019'!$V" & LineZ + 1 & "+'2019'!$AA" & LineZ + 1 & "+'2019'!$AF" & LineZ + 1 & "+'2019'!$AK" & LineZ + 1 & "+'2019'!$AP" & LineZ + 1 & "+'2019'!$AU" & LineZ + 1 & "+'2019'!$AZ" & LineZ + 1 & "+'2019'!$BE" & LineZ + 1 & "+'2019'!$BJ" & LineZ + 1 & "+'2019'!$BO" & LineZ + 1 & ")/'2019'!$K" & LineZ + 1 & ",)"
                Ws.Cells(LineZ, 17) = "=IFERROR(('2019'!$Q" & LineZ + 1 & "+'2019'!$V" & LineZ + 1 & "+'2019'!$AA" & LineZ + 1 & "+'2019'!$AF" & LineZ + 1 & "+'2019'!$AK" & LineZ + 1 & "+'2019'!$AP" & LineZ + 1 & "+'2019'!$AU" & LineZ + 1 & "+'2019'!$AZ" & LineZ + 1 & "+'2019'!$BE" & LineZ + 1 & "+'2019'!$BJ" & LineZ + 1 & "+'2019'!$BO" & LineZ + 1 & "+'2019'!$BT" & LineZ + 1 & ")/'2019'!$K" & LineZ + 1 & ",)"
                Ws.Cells(LineZ, 18) = "=IFERROR(('2019'!$Q" & LineZ + 1 & "+'2019'!$V" & LineZ + 1 & "+'2019'!$AA" & LineZ + 1 & "+'2019'!$AF" & LineZ + 1 & "+'2019'!$AK" & LineZ + 1 & "+'2019'!$AP" & LineZ + 1 & "+'2019'!$AU" & LineZ + 1 & "+'2019'!$AZ" & LineZ + 1 & "+'2019'!$BE" & LineZ + 1 & "+'2019'!$BJ" & LineZ + 1 & "+'2019'!$BO" & LineZ + 1 & "+'2019'!$BT" & LineZ + 1 & "+'2019'!$BY" & LineZ + 1 & ")/'2019'!$K" & LineZ + 1 & ",)"
                Ws.Cells(LineZ, 19) = "=IFERROR(('2019'!Q" & LineZ + 1 & "+'2019'!CD" & LineZ + 1 & ")/'2019'!K" & LineZ + 1 & ",)"
                LineZ += 1
                Me.Label2.Text = "Page2 " & LineZ
            End While
            Ws.Cells(LineZ, 1) = "Sub-Total"
            Ws.Cells(LineZ, 7) = "=SUBTOTAL(101,G3:G" & LineZ - 1 & ")"
            oRng = Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 7))
            oRng.AutoFill(Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ, 19)), Type:=xlFillDefault)

            oRng = Ws.Range("A3", Ws.Cells(LineZ, 19))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        CopyReader1.Close()


        ' 第三頁收入
        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        AdjustExcelFormat3()
        LineZ = 3
        If CopyReader2.HasRows() Then
            While CopyReader2.Read()
                Ws.Cells(LineZ, 1) = CopyReader2.Item("pja01")
                Ws.Cells(LineZ, 2) = CopyReader2.Item("pja02")
                Ws.Cells(LineZ, 3) = CopyReader2.Item("X1")
                Ws.Cells(LineZ, 4) = CopyReader2.Item("pjaud04")
                Ws.Cells(LineZ, 6) = "='2019'!$I" & LineZ + 1 & "*(完工比例!G" & LineZ & "-完工比例!$F" & LineZ & ")"
                Ws.Cells(LineZ, 7) = "='2019'!$I" & LineZ + 1 & "*(完工比例!H" & LineZ & "-完工比例!$F" & LineZ & ")"
                Ws.Cells(LineZ, 8) = "='2019'!$I" & LineZ + 1 & "*(完工比例!I" & LineZ & "-完工比例!$F" & LineZ & ")"
                Ws.Cells(LineZ, 9) = "='2019'!$I" & LineZ + 1 & "*(完工比例!J" & LineZ & "-完工比例!$F" & LineZ & ")"
                Ws.Cells(LineZ, 10) = "='2019'!$I" & LineZ + 1 & "*(完工比例!K" & LineZ & "-完工比例!$F" & LineZ & ")"
                Ws.Cells(LineZ, 11) = "='2019'!$I" & LineZ + 1 & "*(完工比例!L" & LineZ & "-完工比例!$F" & LineZ & ")"
                Ws.Cells(LineZ, 12) = "='2019'!$I" & LineZ + 1 & "*(完工比例!M" & LineZ & "-完工比例!$F" & LineZ & ")"
                Ws.Cells(LineZ, 13) = "='2019'!$I" & LineZ + 1 & "*(完工比例!N" & LineZ & "-完工比例!$F" & LineZ & ")"
                Ws.Cells(LineZ, 14) = "='2019'!$I" & LineZ + 1 & "*(完工比例!O" & LineZ & "-完工比例!$F" & LineZ & ")"
                Ws.Cells(LineZ, 15) = "='2019'!$I" & LineZ + 1 & "*(完工比例!P" & LineZ & "-完工比例!$F" & LineZ & ")"
                Ws.Cells(LineZ, 16) = "='2019'!$I" & LineZ + 1 & "*(完工比例!Q" & LineZ & "-完工比例!$F" & LineZ & ")"
                Ws.Cells(LineZ, 17) = "='2019'!$I" & LineZ + 1 & "*(完工比例!R" & LineZ & "-完工比例!$F" & LineZ & ")"
                Ws.Cells(LineZ, 18) = "='2019'!$I" & LineZ + 1 & "*(完工比例!S" & LineZ & "-完工比例!$F" & LineZ & ")"
                LineZ += 1
                Me.Label2.Text = "Page3 " & LineZ
            End While
            Ws.Cells(LineZ, 1) = "Sub-Total"
            Ws.Cells(LineZ, 6) = "=SUBTOTAL(101,F3:F" & LineZ - 1 & ")"
            oRng = Ws.Range(Ws.Cells(LineZ, 6), Ws.Cells(LineZ, 6))
            oRng.AutoFill(Ws.Range(Ws.Cells(LineZ, 6), Ws.Cells(LineZ, 18)), Type:=xlFillDefault)

            oRng = Ws.Range("A3", Ws.Cells(LineZ, 18))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        CopyReader2.Close()


        ' 第四頁成本
        Ws = xWorkBook.Sheets(4)
        Ws.Activate()
        AdjustExcelFormat3()
        LineZ = 3
        If CopyReader3.HasRows() Then
            While CopyReader3.Read()
                Ws.Cells(LineZ, 1) = CopyReader3.Item("pja01")
                Ws.Cells(LineZ, 2) = CopyReader3.Item("pja02")
                Ws.Cells(LineZ, 3) = CopyReader3.Item("X1")
                Ws.Cells(LineZ, 4) = CopyReader3.Item("pjaud04")
                Ws.Cells(LineZ, 6) = "='2019'!$N" & LineZ + 1 & "*完工比例!G" & LineZ
                Ws.Cells(LineZ, 7) = "='2019'!$N" & LineZ + 1 & "*完工比例!H" & LineZ
                Ws.Cells(LineZ, 8) = "='2019'!$N" & LineZ + 1 & "*完工比例!I" & LineZ
                Ws.Cells(LineZ, 9) = "='2019'!$N" & LineZ + 1 & "*完工比例!J" & LineZ
                Ws.Cells(LineZ, 10) = "='2019'!$N" & LineZ + 1 & "*完工比例!K" & LineZ
                Ws.Cells(LineZ, 11) = "='2019'!$N" & LineZ + 1 & "*完工比例!L" & LineZ
                Ws.Cells(LineZ, 12) = "='2019'!$N" & LineZ + 1 & "*完工比例!M" & LineZ
                Ws.Cells(LineZ, 13) = "='2019'!$N" & LineZ + 1 & "*完工比例!N" & LineZ
                Ws.Cells(LineZ, 14) = "='2019'!$N" & LineZ + 1 & "*完工比例!O" & LineZ
                Ws.Cells(LineZ, 15) = "='2019'!$N" & LineZ + 1 & "*完工比例!P" & LineZ
                Ws.Cells(LineZ, 16) = "='2019'!$N" & LineZ + 1 & "*完工比例!Q" & LineZ
                Ws.Cells(LineZ, 17) = "='2019'!$N" & LineZ + 1 & "*完工比例!R" & LineZ
                Ws.Cells(LineZ, 18) = "='2019'!$N" & LineZ + 1 & "*完工比例!S" & LineZ
                LineZ += 1
                Me.Label2.Text = "Page4 " & LineZ
            End While
            Ws.Cells(LineZ, 1) = "Sub-Total"
            Ws.Cells(LineZ, 6) = "=SUBTOTAL(101,F3:F" & LineZ - 1 & ")"
            oRng = Ws.Range(Ws.Cells(LineZ, 6), Ws.Cells(LineZ, 6))
            oRng.AutoFill(Ws.Range(Ws.Cells(LineZ, 6), Ws.Cells(LineZ, 18)), Type:=xlFillDefault)

            oRng = Ws.Range("A3", Ws.Cells(LineZ, 18))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        CopyReader3.Close()

    End Sub
    Private Sub AdjustExcelFormat()
        Ws.Cells(2, 16) = "YTD " & pYear & "/12/31 cost incurred"
        For i As Int16 = 1 To 12 Step 1
            Ws.Cells(2, 21 + (i - 1) * 5) = tTime.AddMonths(i - 1).ToString("yyyy/MM/dd") & "-" & tTime.AddMonths(i).AddDays(-1).ToString("yyyy/MM/dd")
        Next
        Ws.Cells(2, 81) = tYear
    End Sub
    Private Sub GetMoldFee(ByVal pja01 As String)
        oCommand2.CommandText = "select "
        For j As Int16 = 1 To 12 Step 1
            oCommand2.CommandText += "(case when month(tlf06) = " & j & " then round(sum(rvv39 * pmm42 / azj041),3) else 0 end) as t" & j & ","
        Next
        oCommand2.CommandText += "1 from TLF_FILE left join rvv_file on tlf905=rvv01 and tlf906 = rvv02 left join pmm_file on rvv36 = pmm01 left join pmn_file on pmm01 = pmn01 and rvv37 = pmn02 "
        oCommand2.CommandText += "left join azj_file on azj01 = 'USD' AND azj02 = year(tlf06) || (case when length(month(tlf06)) = 1 then '0' || to_char(month(tlf06)) else to_char(month(tlf06)) end) "
        oCommand2.CommandText += "where tlf13 = 'apmt150' and tlf01 like '7%' and tlf20 = '" & pja01 & "' and tlf06 between to_date('"
        oCommand2.CommandText += tTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand2.CommandText += fTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmnud04 = '1' group by tlf06 order by tlf20"
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                For j As Int16 = 1 To 12 Step 1
                    Ws.Cells(LineZ, 21 + (j - 1) * 5) = oReader2.Item(j - 1)
                Next
            End While
        Else
            For j As Int16 = 1 To 12 Step 1
                Ws.Cells(LineZ, 21 + (j - 1) * 5) = 0
            Next
        End If
        oReader2.Close()
    End Sub

    Private Sub GetRDHour(ByVal pja01 As String)
        mSQLS1.CommandText = "select "
        For j As Int16 = 1 To 12 Step 1
            mSQLS1.CommandText += "round((case when month(edate) = " & j & " then sum(ehour * 35 / rate1) end),3) as t" & j & ","
        Next
        mSQLS1.CommandText += "1 from ProjectHR left join USD_ExchangeRate on month(edate) = Month1 where EProject = '" & pja01 & "'  and edate between '"
        mSQLS1.CommandText += tTime.ToString("yyyy/MM/dd") & "' and '" & fTime.ToString("yyyy/MM/dd") & "' group by month(edate),rate1"
        mSQLReader = mSQLS1.ExecuteReader
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                For j As Int16 = 1 To 12 Step 1
                    Ws.Cells(LineZ, 22 + (j - 1) * 5) = mSQLReader.Item(j - 1)
                Next
            End While
        Else
            For j As Int16 = 1 To 12 Step 1
                Ws.Cells(LineZ, 22 + (j - 1) * 5) = 0
            Next
        End If
        mSQLReader.Close()
    End Sub

    Private Sub GetMCData(ByVal pja01 As String)
        oCommand2.CommandText = "select "
        For j As Int16 = 1 To 12 Step 1
            oCommand2.CommandText += "sum(t" & j & ") as t" & j & ","
        Next
        oCommand2.CommandText += "1 from ( select "
        For j As Int16 = 1 To 12 Step 1
            oCommand2.CommandText += "(case when month(tlf06) = " & j & " then round(nvl(sum(tlf10 * tlf12 * ccc23 / azj041),0),3) else 0 end ) as t" & j & ","
        Next
        oCommand2.CommandText += "1 from pja_file left join tlf_file on pja01 = tlf20 left join gem_file on tlf19 = gem01 left join ima_file on tlf01 = ima01 "
        oCommand2.CommandText += "left join ina_file on tlf905 = ina01 left join inb_file on tlf905 = inb01 and ina01 = inb01 and tlf906 = inb03 "
        oCommand2.CommandText += "left join ccc_File on tlf01 = ccc01 and ccc02 = year(ina02) and ccc03 = month(ina02) left join azj_file on azj01 = 'USD' AND azj02 = year(tlf06) || (case when length(month(tlf06)) = 1 then '0' || to_char(month(tlf06)) else to_char(month(tlf06)) end) "
        oCommand2.CommandText += "where tlf20 is not null and tlf20 <> ' ' and tlf06 between to_date('"
        oCommand2.CommandText += tTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand2.CommandText += fTime.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ima06 <> '105' and tlf907 = -1 and tlf13 like 'aimt3%' and tlf19 in ('D2300','D3100') AND pja01 = '"
        oCommand2.CommandText += pja01 & "' group by month(tlf06) )"
        oReader2 = oCommand2.ExecuteReader
        If oReader2.HasRows() Then
            While oReader2.Read()
                For j As Int16 = 1 To 12 Step 1
                    Ws.Cells(LineZ, 23 + (j - 1) * 5) = oReader2.Item(j - 1)
                Next
            End While
        Else
            For j As Int16 = 1 To 12 Step 1
                Ws.Cells(LineZ, 23 + (j - 1) * 5) = 0
            Next
        End If
        oReader2.Close()
    End Sub

    Private Sub AdjustExcelFormat2()
        For i As Int16 = 1 To 12 Step 1
            Ws.Cells(1, 7 + (i - 1)) = tTime.ToString("yyyy/MM/dd") & "-" & tTime.AddMonths(i).AddDays(-1).ToString("yyyy/MM/dd")
        Next
        Ws.Cells(1, 19) = tYear
    End Sub
    Private Sub AdjustExcelFormat3()
        For i As Int16 = 1 To 12 Step 1
            Ws.Cells(1, 6 + (i - 1)) = tTime.ToString("yyyy/MM/dd") & "-" & tTime.AddMonths(i).AddDays(-1).ToString("yyyy/MM/dd")
        Next
        Ws.Cells(1, 18) = tYear
    End Sub
End Class