Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form109
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim TYear As String = String.Empty
    Dim TMonth As String = String.Empty
    Dim CYear As String = String.Empty
    Dim CMonth As String = String.Empty
    Dim g_pja01 As String = String.Empty
    Dim g_ima01 As String = String.Empty
    Dim g_imd01 As String = String.Empty
    Dim g_ima06 As String = String.Empty
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form109_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If Now.Month < 10 Then
            TextBox1.Text = Now.Year & "0" & Now.Month
        Else
            TextBox1.Text = Now.Year & Now.Month
        End If
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
                oCommand2.Connection = oConnection
                oCommand2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        TYear = Strings.Left(TextBox1.Text, 4)
        TMonth = Strings.Right(TextBox1.Text, 2)
        ' add by cloud 20171221
        g_ima01 = String.Empty
        g_imd01 = String.Empty
        g_ima06 = String.Empty
        If Not String.IsNullOrEmpty(TextBox2.Text) Then
            g_ima01 = TextBox2.Text
        End If
        If Not String.IsNullOrEmpty(TextBox3.Text) Then
            g_imd01 = TextBox3.Text
        End If
        If Not String.IsNullOrEmpty(TextBox4.Text) Then
            g_ima06 = TextBox4.Text
        End If
        'ExportToExcel()
        'SaveExcel()
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Aging-Cost"
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
        AdjustExcelFormat1()
        Dim TotalNumber As Integer = 0
        oCommand.CommandText = "select img01,ima02,ima021,img02,imd02,ima06,imz02,img09,img10,nvl(ccc23,0) as t2,img37,floor(sysdate -img37) as t1 from img_File "
        oCommand.CommandText += "left join ima_File on img01 = ima01 left join imd_file on img02 = imd01 left join imz_file on ima06 = imz01 left join ccc_File on img01 = ccc01 and ima01 = ccc01 and ccc02 = "
        oCommand.CommandText += TYear & " and ccc03 = " & TMonth & " where img10 <> 0 "
        If Not String.IsNullOrEmpty(g_ima01) Then
            oCommand.CommandText += " AND img01 like '" & g_ima01 & "%' "
        End If
        If Not String.IsNullOrEmpty(g_imd01) Then
            oCommand.CommandText += " AND imd01 like '" & g_imd01 & "%' "
        End If
        If Not String.IsNullOrEmpty(g_ima06) Then
            oCommand.CommandText += " AND ima06 = '" & g_ima06 & "' "
        End If
        oCommand.CommandText += " and img02 not in (select jce02 from jce_file) "
        oReader = oCommand.ExecuteReader()
        oReader2 = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("img01")
                Ws.Cells(LineZ, 2) = oReader.Item("ima02")
                Ws.Cells(LineZ, 3) = oReader.Item("ima021")
                Ws.Cells(LineZ, 4) = oReader.Item("img02")
                Ws.Cells(LineZ, 5) = oReader.Item("imd02")
                Ws.Cells(LineZ, 6) = oReader.Item("ima06")
                Ws.Cells(LineZ, 7) = oReader.Item("imz02")
                Ws.Cells(LineZ, 8) = oReader.Item("img09")
                Ws.Cells(LineZ, 9) = oReader.Item("img10")
                Ws.Cells(LineZ, 10) = oReader.Item("t2")
                Ws.Cells(LineZ, 11) = Decimal.Round(oReader.Item("img10") * oReader.Item("t2"), 2)
                Ws.Cells(LineZ, 12) = oReader.Item("img37")
                Ws.Cells(LineZ, 13) = Now.ToString("yyyy/MM/dd")
                Ws.Cells(LineZ, 14) = oReader.Item("t1")
                Select Case oReader.Item("t1")
                    Case 0 To 30
                        Ws.Cells(LineZ, 15) = Decimal.Round(oReader.Item("img10") * oReader.Item("t2"), 2)
                    Case 31 To 60
                        Ws.Cells(LineZ, 16) = Decimal.Round(oReader.Item("img10") * oReader.Item("t2"), 2)
                    Case 61 To 90
                        Ws.Cells(LineZ, 17) = Decimal.Round(oReader.Item("img10") * oReader.Item("t2"), 2)
                    Case 91 To 120
                        Ws.Cells(LineZ, 18) = Decimal.Round(oReader.Item("img10") * oReader.Item("t2"), 2)
                    Case 121 To 180
                        Ws.Cells(LineZ, 19) = Decimal.Round(oReader.Item("img10") * oReader.Item("t2"), 2)
                    Case Is > 180
                        Ws.Cells(LineZ, 20) = Decimal.Round(oReader.Item("img10") * oReader.Item("t2"), 2)
                End Select
                LineZ += 1
                Try
                    TotalNumber += 1
                    Label2.Text = TotalNumber
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End While
        End If
        oReader.Close()
        '最上方的處理
        Ws.Cells(1, 9) = "=SUM(I4:I" & LineZ - 1 & ")"
        Ws.Cells(1, 11) = "=SUM(K4:K" & LineZ - 1 & ")"
        Ws.Cells(1, 15) = "=SUM(O4:O" & LineZ - 1 & ")"
        Ws.Cells(1, 16) = "=SUM(P4:P" & LineZ - 1 & ")"
        Ws.Cells(1, 17) = "=SUM(Q4:Q" & LineZ - 1 & ")"
        Ws.Cells(1, 18) = "=SUM(R4:R" & LineZ - 1 & ")"
        Ws.Cells(1, 19) = "=SUM(S4:S" & LineZ - 1 & ")"
        Ws.Cells(1, 20) = "=SUM(T4:T" & LineZ - 1 & ")"

        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustExcelFormat2()
        If oReader2.HasRows() Then
            While oReader2.Read()
                Ws.Cells(LineZ, 1) = oReader2.Item("img01")
                Ws.Cells(LineZ, 2) = oReader2.Item("ima02")
                Ws.Cells(LineZ, 3) = oReader2.Item("ima021")
                Ws.Cells(LineZ, 4) = oReader2.Item("img02")
                Ws.Cells(LineZ, 5) = oReader2.Item("imd02")
                Ws.Cells(LineZ, 6) = oReader2.Item("ima06")
                Ws.Cells(LineZ, 7) = oReader2.Item("imz02")
                Ws.Cells(LineZ, 8) = oReader2.Item("img09")
                Ws.Cells(LineZ, 9) = oReader2.Item("img10")
                Ws.Cells(LineZ, 10) = oReader2.Item("img37")
                Ws.Cells(LineZ, 11) = Now.ToString("yyyy/MM/dd")
                Ws.Cells(LineZ, 12) = oReader2.Item("t1")
                Select Case oReader2.Item("t1")
                    Case 0 To 30
                        Ws.Cells(LineZ, 13) = oReader2.Item("img10")
                    Case 31 To 60
                        Ws.Cells(LineZ, 14) = oReader2.Item("img10")
                    Case 61 To 90
                        Ws.Cells(LineZ, 15) = oReader2.Item("img10")
                    Case 91 To 120
                        Ws.Cells(LineZ, 16) = oReader2.Item("img10")
                    Case 121 To 180
                        Ws.Cells(LineZ, 17) = oReader2.Item("img10")
                    Case Is > 180
                        Ws.Cells(LineZ, 18) = oReader2.Item("img10")
                End Select
                LineZ += 1
                TotalNumber += 1
                Label2.Text = TotalNumber
            End While
        End If
        oReader2.Close()
        '最上方的處理
        Ws.Cells(1, 9) = "=SUM(I4:I" & LineZ - 1 & ")"
        Ws.Cells(1, 13) = "=SUM(M4:M" & LineZ - 1 & ")"
        Ws.Cells(1, 14) = "=SUM(N4:N" & LineZ - 1 & ")"
        Ws.Cells(1, 15) = "=SUM(O4:O" & LineZ - 1 & ")"
        Ws.Cells(1, 16) = "=SUM(P4:P" & LineZ - 1 & ")"
        Ws.Cells(1, 17) = "=SUM(Q4:Q" & LineZ - 1 & ")"
        Ws.Cells(1, 18) = "=SUM(R4:R" & LineZ - 1 & ")"

        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        AdjustExcelFormat3()
        ' 103 庫存商品
        Ws.Cells(5, 2) = "库存商品"
        Ws.Cells(5, 3) = "FG-DAC"
        oCommand.CommandText = "Select sum(Case when t1 <= 30 then t2 * img10 end) as v1 ,sum(Case when t1 > 30 and t1 <= 60 then t2 * img10 end) as v2 ,sum(case when t1 > 60 and t1 <= 90 then t2 * img10 end) as v3,"
        oCommand.CommandText += "sum(case when t1 > 90 and t1 <= 120 then t2 * img10 end) as v4,sum(case when t1 > 120 and t1 <= 180 then t2 * img10 end) as v5,sum(case when t1 > 180 then t2 * img10 end) as v6 "
        oCommand.CommandText += "from ( select img01,ima02,ima021,img02,imd02,ima06,imz02,img09,img10,nvl(ccc23,0) as t2,img37,floor(sysdate -img37) as t1 from img_File "
        oCommand.CommandText += "left join ima_File on img01 = ima01 left join imd_file on img02 = imd01 left join imz_file on ima06 = imz01 left join ccc_File on img01 = ccc01 and ima01 = ccc01 and ccc02 = "
        oCommand.CommandText += TYear & " and ccc03 = " & TMonth & " where img10 <> 0 and ima06 = '103' "
        If Not String.IsNullOrEmpty(g_ima01) Then
            oCommand.CommandText += " AND img01 like '" & g_ima01 & "%' "
        End If
        If Not String.IsNullOrEmpty(g_imd01) Then
            oCommand.CommandText += " AND imd01 like '" & g_imd01 & "%' "
        End If
        If Not String.IsNullOrEmpty(g_ima06) Then
            oCommand.CommandText += " AND ima06 = '" & g_ima06 & "' "
        End If
        oCommand.CommandText += " and img02 not in (select jce02 from jce_file) ) "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(5, 4 + i) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()
        Ws.Cells(5, 10) = "=SUM(D5:I5)"
        ' 102
        Ws.Cells(6, 2) = "半成品"
        Ws.Cells(6, 3) = "WIP-DAC"
        oCommand.CommandText = "Select sum(Case when t1 <= 30 then t2 * img10 end) as v1 ,sum(Case when t1 > 30 and t1 <= 60 then t2 * img10 end) as v2 ,sum(case when t1 > 60 and t1 <= 90 then t2 * img10 end) as v3,"
        oCommand.CommandText += "sum(case when t1 > 90 and t1 <= 120 then t2 * img10 end) as v4,sum(case when t1 > 120 and t1 <= 180 then t2 * img10 end) as v5,sum(case when t1 > 180 then t2 * img10 end) as v6 "
        oCommand.CommandText += "from ( select img01,ima02,ima021,img02,imd02,ima06,imz02,img09,img10,nvl(ccc23,0) as t2,img37,floor(sysdate -img37) as t1 from img_File "
        oCommand.CommandText += "left join ima_File on img01 = ima01 left join imd_file on img02 = imd01 left join imz_file on ima06 = imz01 left join ccc_File on img01 = ccc01 and ima01 = ccc01 and ccc02 = "
        oCommand.CommandText += TYear & " and ccc03 = " & TMonth & " where img10 <> 0 and ima06 = '102' "
        If Not String.IsNullOrEmpty(g_ima01) Then
            oCommand.CommandText += " AND img01 like '" & g_ima01 & "%' "
        End If
        If Not String.IsNullOrEmpty(g_imd01) Then
            oCommand.CommandText += " AND imd01 like '" & g_imd01 & "%' "
        End If
        If Not String.IsNullOrEmpty(g_ima06) Then
            oCommand.CommandText += " AND ima06 = '" & g_ima06 & "' "
        End If
        oCommand.CommandText += " and img02 not in (select jce02 from jce_file) ) "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(6, 4 + i) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()
        Ws.Cells(6, 10) = "=SUM(D6:I6)"
        '101, 104, 106
        Ws.Cells(7, 2) = "物料"
        Ws.Cells(7, 3) = "Material-DAC"
        oCommand.CommandText = "Select sum(Case when t1 <= 30 then t2 * img10 end) as v1 ,sum(Case when t1 > 30 and t1 <= 60 then t2 * img10 end) as v2 ,sum(case when t1 > 60 and t1 <= 90 then t2 * img10 end) as v3,"
        oCommand.CommandText += "sum(case when t1 > 90 and t1 <= 120 then t2 * img10 end) as v4,sum(case when t1 > 120 and t1 <= 180 then t2 * img10 end) as v5,sum(case when t1 > 180 then t2 * img10 end) as v6 "
        oCommand.CommandText += "from ( select img01,ima02,ima021,img02,imd02,ima06,imz02,img09,img10,nvl(ccc23,0) as t2,img37,floor(sysdate -img37) as t1 from img_File "
        oCommand.CommandText += "left join ima_File on img01 = ima01 left join imd_file on img02 = imd01 left join imz_file on ima06 = imz01 left join ccc_File on img01 = ccc01 and ima01 = ccc01 and ccc02 = "
        oCommand.CommandText += TYear & " and ccc03 = " & TMonth & " where img10 <> 0 and ima06 IN ('101','104', '106') "
        If Not String.IsNullOrEmpty(g_ima01) Then
            oCommand.CommandText += " AND img01 like '" & g_ima01 & "%' "
        End If
        If Not String.IsNullOrEmpty(g_imd01) Then
            oCommand.CommandText += " AND imd01 like '" & g_imd01 & "%' "
        End If
        If Not String.IsNullOrEmpty(g_ima06) Then
            oCommand.CommandText += " AND ima06 = '" & g_ima06 & "' "
        End If
        oCommand.CommandText += " and img02 not in (select jce02 from jce_file) ) "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(7, 4 + i) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()
        Ws.Cells(7, 10) = "=SUM(D7:I7)"
        ' 加總
        Ws.Cells(8, 3) = "Total"
        Ws.Cells(8, 4) = "=SUM(D5:D7)"
        Ws.Cells(8, 5) = "=SUM(E5:E7)"
        Ws.Cells(8, 6) = "=SUM(F5:F7)"
        Ws.Cells(8, 7) = "=SUM(G5:G7)"
        Ws.Cells(8, 8) = "=SUM(H5:H7)"
        Ws.Cells(8, 9) = "=SUM(I5:I7)"
        Ws.Cells(8, 10) = "=SUM(J5:J7)"
        ' 劃線
        oRng = Ws.Range("B4", "J8")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        ' 調整大小
        oRng = Ws.Range("A1", "J1")
        oRng.EntireColumn.AutoFit()

    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "inventory aging-amount"
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 25.89
        Ws.Cells(1, 1) = "币别：RMB"
        Ws.Cells(1, 8) = "sub-total"
        Ws.Cells(1, 14) = "sub-total"
        Ws.Cells(2, 1) = "料件编号"
        Ws.Cells(2, 2) = "品名"
        Ws.Cells(2, 3) = "规格"
        Ws.Cells(2, 4) = "仓库"
        Ws.Cells(2, 5) = "仓库名称"
        Ws.Cells(2, 6) = "分群码"
        Ws.Cells(2, 7) = "分群说明"
        Ws.Cells(2, 8) = "库存单位"
        Ws.Cells(2, 9) = "库存数量"
        Ws.Cells(2, 10) = "单价"
        Ws.Cells(2, 11) = "金额"
        Ws.Cells(2, 12) = "呆滞日期"
        Ws.Cells(2, 13) = "查询日期"
        Ws.Cells(2, 14) = "差异天数"
        Ws.Cells(2, 15) = "<=30"
        Ws.Cells(2, 16) = "<=60"
        Ws.Cells(2, 17) = "<=90"
        Ws.Cells(2, 18) = "<=120"
        Ws.Cells(2, 19) = "<=180"
        Ws.Cells(2, 20) = ">180"
        Ws.Cells(3, 1) = "part no."
        Ws.Cells(3, 2) = "part name"
        Ws.Cells(3, 3) = "Sepc."
        Ws.Cells(3, 4) = "WH code"
        Ws.Cells(3, 5) = "WH Name"
        Ws.Cells(3, 6) = "M. group code"
        Ws.Cells(3, 7) = "M. group name"
        Ws.Cells(3, 8) = "inventory unit"
        Ws.Cells(3, 9) = "qty"
        Ws.Cells(3, 10) = "price"
        Ws.Cells(3, 11) = "amount"
        Ws.Cells(3, 12) = "last moving dated"
        Ws.Cells(3, 13) = "Inquery date"
        Ws.Cells(3, 14) = "age by days"
        Ws.Cells(3, 15) = "<=30"
        Ws.Cells(3, 16) = "<=60"
        Ws.Cells(3, 17) = "<=90"
        Ws.Cells(3, 18) = "<=120"
        Ws.Cells(3, 19) = "<=180"
        Ws.Cells(3, 20) = ">180"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormatLocal = "@"

        LineZ = 4
    End Sub
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "inventory aging-qty"
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 25.89
        Ws.Cells(1, 1) = "币别：RMB"
        Ws.Cells(1, 8) = "sub-total"
        Ws.Cells(1, 12) = "sub-total"
        Ws.Cells(2, 1) = "料件编号"
        Ws.Cells(2, 2) = "品名"
        Ws.Cells(2, 3) = "规格"
        Ws.Cells(2, 4) = "仓库"
        Ws.Cells(2, 5) = "仓库名称"
        Ws.Cells(2, 6) = "分群码"
        Ws.Cells(2, 7) = "分群说明"
        Ws.Cells(2, 8) = "库存单位"
        Ws.Cells(2, 9) = "库存数量"
        Ws.Cells(2, 10) = "呆滞日期"
        Ws.Cells(2, 11) = "查询日期"
        Ws.Cells(2, 12) = "差异天数"
        Ws.Cells(2, 13) = "<=30"
        Ws.Cells(2, 14) = "<=60"
        Ws.Cells(2, 15) = "<=90"
        Ws.Cells(2, 16) = "<=120"
        Ws.Cells(2, 17) = "<=180"
        Ws.Cells(2, 18) = ">180"
        Ws.Cells(3, 1) = "part no."
        Ws.Cells(3, 2) = "part name"
        Ws.Cells(3, 3) = "Sepc."
        Ws.Cells(3, 4) = "WH code"
        Ws.Cells(3, 5) = "WH Name"
        Ws.Cells(3, 6) = "M. group code"
        Ws.Cells(3, 7) = "M. group name"
        Ws.Cells(3, 8) = "inventory unit"
        Ws.Cells(3, 9) = "qty"
        Ws.Cells(3, 10) = "last moving dated"
        Ws.Cells(3, 11) = "Inquery date"
        Ws.Cells(3, 12) = "age by days"
        Ws.Cells(3, 13) = "<=30"
        Ws.Cells(3, 14) = "<=60"
        Ws.Cells(3, 15) = "<=90"
        Ws.Cells(3, 16) = "<=120"
        Ws.Cells(3, 17) = "<=180"
        Ws.Cells(3, 18) = ">180"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormatLocal = "@"

        LineZ = 4
    End Sub
    Private Sub AdjustExcelFormat3()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "Total Amount"
        Ws.Rows.EntireRow.RowHeight = 20
        Ws.Rows.EntireRow.NumberFormatLocal = "#,##0.00_ "
        oRng = Ws.Range("B2", "B2")
        oRng.EntireRow.RowHeight = 40
        oRng = Ws.Range("B2", "J2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(2, 2) = "DAC 库龄表" & Chr(10) & "DAC inventory aging-amount"
        Ws.Cells(3, 2) = "币别：RMB"
        Ws.Cells(4, 2) = "项目"
        Ws.Cells(4, 3) = "Item"
        Ws.Cells(4, 4) = "<=30"
        Ws.Cells(4, 5) = "<=60"
        Ws.Cells(4, 6) = "<=90"
        Ws.Cells(4, 7) = "<=120"
        Ws.Cells(4, 8) = "<=180"
        Ws.Cells(4, 9) = ">180"
        Ws.Cells(4, 10) = "Total"

        LineZ = 5
    End Sub
End Class