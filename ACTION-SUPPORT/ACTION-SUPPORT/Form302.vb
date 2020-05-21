Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel

Public Class Form302
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oSQLReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oSQLS1 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form302_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If

        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()

        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oSQLS1.Connection = oConnection
                oSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        BackgroundWorker1.RunWorkerAsync()
        'ExportToExcel()
        'oConnection.Close()

        'oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        'If oConnection.State <> ConnectionState.Open Then
        'Try
        'oConnection.Open()
        'oCommand.Connection = oConnection
        'oCommand.CommandType = CommandType.Text
        'oSQLS1.Connection = oConnection
        'oSQLS1.CommandType = CommandType.Text
        'Catch ex As Exception
        'MsgBox(ex.Message)
        'End Try
        'End If
        'ExportToExcel_1()

        'SaveExcel()
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
        ExportToExcel_1()
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub

    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "单阶材料用途查询"
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

        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        Ws.Name = "不含元件下階料號"
        AdjustExcelFormat()
        oCommand.CommandText = "select bmb03,bma01,ima02,ima021,ima06,ima08,bmb04,bmb05,bmb07,bmb06,bmb08,ima44,bmb10,round((bmb06/bmb07)*(bmb08/100),8) as B1, round((bmb06/bmb07)+(bmb06/bmb07)*(bmb08/100),8) as B2,bma06  from bmb_file,ima_file,bma_file "
        oCommand.CommandText += " where bmb03=ima01 and bmb01=bma01 and imaacti='Y' and  bmaacti='Y' and bma10='2' and  bmb05 is  null and ima06 in('101','106','104') and ima08 in('P','S','D') order by bmb03 "

        oSQLReader = oCommand.ExecuteReader()
        If oSQLReader.HasRows() Then
            While oSQLReader.Read()
                Ws.Cells(LineZ, 1) = oSQLReader.Item("bmb03")
                Ws.Cells(LineZ, 2) = oSQLReader.Item("bma01")
                Ws.Cells(LineZ, 3) = oSQLReader.Item("ima02")
                Ws.Cells(LineZ, 4) = oSQLReader.Item("ima021")
                Ws.Cells(LineZ, 5) = oSQLReader.Item("ima06")
                Ws.Cells(LineZ, 6) = oSQLReader.Item("ima08")
                Ws.Cells(LineZ, 7) = oSQLReader.Item("bmb04")
                Ws.Cells(LineZ, 8) = oSQLReader.Item("bmb05")
                Ws.Cells(LineZ, 9) = oSQLReader.Item("bmb07")
                Ws.Cells(LineZ, 10) = oSQLReader.Item("bmb06")
                Ws.Cells(LineZ, 11) = oSQLReader.Item("bmb08")
                Ws.Cells(LineZ, 12) = oSQLReader.Item("ima44")
                Ws.Cells(LineZ, 13) = oSQLReader.Item("bmb10")
                Ws.Cells(LineZ, 14) = oSQLReader.Item("B1")
                Ws.Cells(LineZ, 15) = oSQLReader.Item("B2")
                Ws.Cells(LineZ, 16) = oSQLReader.Item("bma06")
                LineZ += 1

            End While
        End If
        oSQLReader.Close()
        Ws.Cells(1, 2) = Now()                     '200506 add by Brady
        'Try
        'oCommand.ExecuteNonQuery()
        'Catch ex As Exception
        'MsgBox(ex.Message())
        'Return
        'End Try
    End Sub

    Private Sub ExportToExcel_1()

        ' 第二頁 (顯示元件下階料号)    
        Ws = xWorkBook.Sheets(2)
        Ws.Name = "含元件下階料號"
        Ws.Activate()
        AdjustExcelFormat1()

        '200506 add by Brady
        ''190301 add by Brady
        ''oCommand.CommandText = "select b.bmb03 as bmb03_b,c.bmb03 as bmb03_c,a.bma01 as bma01_a,ima02,ima021,ima06,ima08,c.bmb04 as bmb04_c,c.bmb05 as bmb05_c, "
        ''oCommand.CommandText += "c.bmb07 as bmb07_c, c.bmb06 as bmb06_c, c.bmb08 as bmb08_c, ima44, "
        ''oCommand.CommandText += "c.bmb10 as bmb10_c,round((c.bmb06/c.bmb07)*(c.bmb08/100),8) as B1,round((c.bmb06/c.bmb07)+(c.bmb06/c.bmb07)*(c.bmb08/100),8) as B2,a.bma06 as bma06_a "
        ''oCommand.CommandText += "from bmb_file c left outer join (select bma01,bmb03 from bma_file,bmb_file "
        ''oCommand.CommandText += "where bmaacti='Y' and bma10='2' and bma01 = bmb01 and bmb05 is null) b on b.bma01 = c.bmb03, ima_file,bma_file a "
        ''oCommand.CommandText += "where c.bmb03=ima01 and imaacti='Y' and c.bmb01=a.bma01 and  a.bmaacti='Y' and a.bma10='2' "
        ''oCommand.CommandText += "and  c.bmb05 is  null and ima06 in('101','102','106','104') and ima08 in('P','S','D','M') and length(c.bmb03) <= 12 "
        ''oCommand.CommandText += "order by c.bmb03,a.bma01"
        'oCommand.CommandText = "select b.bmb03 as bmb03_b,c.bmb03 as bmb03_c,a.bma01 as bma01_a,ima02,ima021,ima06,ima08,c.bmb04 as bmb04_c,c.bmb05 as bmb05_c, "
        'oCommand.CommandText += "c.bmb07 as bmb07_c, c.bmb06 as bmb06_c, c.bmb08 as bmb08_c, ima44, ima46, ima48, "
        'oCommand.CommandText += "c.bmb10 as bmb10_c,round((c.bmb06/c.bmb07)*(c.bmb08/100),8) as B1,round((c.bmb06/c.bmb07)+(c.bmb06/c.bmb07)*(c.bmb08/100),8) as B2,a.bma06 as bma06_a "
        'oCommand.CommandText += "from bmb_file c left outer join (select bma01,bmb03 from bma_file,bmb_file "
        'oCommand.CommandText += "where bmaacti='Y' and bma10='2' and bma01 = bmb01 and bmb05 is null) b on b.bma01 = c.bmb03, ima_file,bma_file a "
        'oCommand.CommandText += "where c.bmb03=ima01 and imaacti='Y' and c.bmb01=a.bma01 and  a.bmaacti='Y' and a.bma10='2' "
        'oCommand.CommandText += "and  c.bmb05 is  null and ima06 in('101','102','106','104') and ima08 in('P','S','D','M') and length(c.bmb03) <= 12 "
        'oCommand.CommandText += "order by c.bmb03,a.bma01"
        ''190301 add by Brady END        
        oCommand.CommandText = "select b.bmb03 as bmb03_b,c.bmb03 as bmb03_c,a.bma01 as bma01_a,ima02,ima021,ima06,ima08,c.bmb04 as bmb04_c,c.bmb05 as bmb05_c, "
        oCommand.CommandText += "c.bmb07 as bmb07_c, c.bmb06 as bmb06_c, c.bmb08 as bmb08_c, ima44, ima46, ima48, "
        oCommand.CommandText += "c.bmb10 as bmb10_c,round((c.bmb06/c.bmb07)*(c.bmb08/100),8) as B1,round((c.bmb06/c.bmb07)+(c.bmb06/c.bmb07)*(c.bmb08/100),8) as B2,a.bma06 as bma06_a "
        oCommand.CommandText += "from bmb_file c left outer join (select bma01,bmb03 from bma_file,bmb_file "
        oCommand.CommandText += "where bmaacti='Y' and bma10='2' and bma01 = bmb01 and bma06 = bmb29 and bmb05 is null) b on b.bma01 = c.bmb03, ima_file,bma_file a "
        oCommand.CommandText += "where c.bmb03=ima01 and imaacti='Y' and c.bmb01=a.bma01 and  a.bmaacti='Y' and a.bma10='2' and a.bma06 = c.bmb29 "
        oCommand.CommandText += "and  c.bmb05 is  null and ima06 in('101','102','106','104') and ima08 in('P','S','D','M') and length(c.bmb03) <= 12 "
        oCommand.CommandText += "order by c.bmb03,a.bma01"
        '200506 add by Brady END


        oSQLReader = oCommand.ExecuteReader()
        If oSQLReader.HasRows() Then
            While oSQLReader.Read()
                Ws.Cells(LineZ, 1) = oSQLReader.Item("bmb03_b")
                Ws.Cells(LineZ, 2) = oSQLReader.Item("bmb03_c")
                Ws.Cells(LineZ, 3) = oSQLReader.Item("bma01_a")
                Ws.Cells(LineZ, 4) = oSQLReader.Item("ima02")
                Ws.Cells(LineZ, 5) = oSQLReader.Item("ima021")
                Ws.Cells(LineZ, 6) = oSQLReader.Item("ima06")
                Ws.Cells(LineZ, 7) = oSQLReader.Item("ima08")
                Ws.Cells(LineZ, 8) = oSQLReader.Item("bmb04_c")
                Ws.Cells(LineZ, 9) = oSQLReader.Item("bmb05_c")
                Ws.Cells(LineZ, 10) = oSQLReader.Item("bmb07_c")
                Ws.Cells(LineZ, 11) = oSQLReader.Item("bmb06_c")
                Ws.Cells(LineZ, 12) = oSQLReader.Item("bmb08_c")
                Ws.Cells(LineZ, 13) = oSQLReader.Item("ima44")
                Ws.Cells(LineZ, 14) = oSQLReader.Item("bmb10_c")
                Ws.Cells(LineZ, 15) = oSQLReader.Item("B1")
                Ws.Cells(LineZ, 16) = oSQLReader.Item("B2")
                Ws.Cells(LineZ, 17) = oSQLReader.Item("bma06_a")
                Ws.Cells(LineZ, 18) = oSQLReader.Item("ima46")       '190301 add by Brady
                Ws.Cells(LineZ, 19) = oSQLReader.Item("ima48")       '190301 add by Brady
                LineZ += 1

            End While
        End If
        oSQLReader.Close()
        Ws.Cells(1, 2) = Now()                                       '200506 add by Brady
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
            Return
        End Try

    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 65
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 20
        '200506 add by Brady
        'Ws.Cells(1, 1) = "元件料号"
        'Ws.Cells(1, 2) = "主件料号"
        'Ws.Cells(1, 3) = "品名"
        'Ws.Cells(1, 4) = "规格"
        'Ws.Cells(1, 5) = "分群码"
        'Ws.Cells(1, 6) = "来源码"
        'Ws.Cells(1, 7) = "生效日期"
        'Ws.Cells(1, 8) = "失效日期"
        'Ws.Cells(1, 9) = "主件底数"
        'Ws.Cells(1, 10) = "组成用量"
        'Ws.Cells(1, 11) = "损耗率"
        'Ws.Cells(1, 12) = "采购单位"
        'Ws.Cells(1, 13) = "发料单位"
        'Ws.Cells(1, 14) = "单位损耗量"
        'Ws.Cells(1, 15) = "实际单位用量QPA"
        'Ws.Cells(1, 16) = "BOM特性代码"
        'oRng = Ws.Range("N1", "Q1")
        'oRng.EntireColumn.NumberFormatLocal = "#,##0.00000000_ "
        'LineZ = 2
        Ws.Cells(1, 1) = "报表打印时间"
        Ws.Cells(2, 1) = "元件料号"
        Ws.Cells(2, 2) = "主件料号"
        Ws.Cells(2, 3) = "品名"
        Ws.Cells(2, 4) = "规格"
        Ws.Cells(2, 5) = "分群码"
        Ws.Cells(2, 6) = "来源码"
        Ws.Cells(2, 7) = "生效日期"
        Ws.Cells(2, 8) = "失效日期"
        Ws.Cells(2, 9) = "主件底数"
        Ws.Cells(2, 10) = "组成用量"
        Ws.Cells(2, 11) = "损耗率"
        Ws.Cells(2, 12) = "采购单位"
        Ws.Cells(2, 13) = "发料单位"
        Ws.Cells(2, 14) = "单位损耗量"
        Ws.Cells(2, 15) = "实际单位用量QPA"
        Ws.Cells(2, 16) = "BOM特性代码"
        oRng = Ws.Range("N2", "Q2")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00000000_ "
        LineZ = 3
        '200506 add by Brady END
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 65
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 20
        '200506 add by Brady
        'Ws.Cells(1, 1) = "元件下階料号"
        'Ws.Cells(1, 2) = "元件料号"
        'Ws.Cells(1, 3) = "主件料号"
        'Ws.Cells(1, 4) = "品名"
        'Ws.Cells(1, 5) = "规格"
        'Ws.Cells(1, 6) = "分群码"
        'Ws.Cells(1, 7) = "来源码"
        'Ws.Cells(1, 8) = "生效日期"
        'Ws.Cells(1, 9) = "失效日期"
        'Ws.Cells(1, 10) = "主件底数"
        'Ws.Cells(1, 11) = "组成用量"
        'Ws.Cells(1, 12) = "损耗率"
        'Ws.Cells(1, 13) = "采购单位"
        'Ws.Cells(1, 14) = "发料单位"
        'Ws.Cells(1, 15) = "单位损耗量"
        'Ws.Cells(1, 16) = "实际单位用量QPA"
        'Ws.Cells(1, 17) = "BOM特性代码"
        'Ws.Cells(1, 18) = "最少采购数量"                      '190301 add by Brady
        'Ws.Cells(1, 19) = "交货前置期（天）"                  '190301 add by Brady
        'oRng = Ws.Range("O1", "R1")
        'oRng.EntireColumn.NumberFormatLocal = "#,##0.00000000_ "
        'LineZ = 2
        Ws.Cells(1, 1) = "报表打印时间"
        Ws.Cells(2, 1) = "元件下階料号"
        Ws.Cells(2, 2) = "元件料号"
        Ws.Cells(2, 3) = "主件料号"
        Ws.Cells(2, 4) = "品名"
        Ws.Cells(2, 5) = "规格"
        Ws.Cells(2, 6) = "分群码"
        Ws.Cells(2, 7) = "来源码"
        Ws.Cells(2, 8) = "生效日期"
        Ws.Cells(2, 9) = "失效日期"
        Ws.Cells(2, 10) = "主件底数"
        Ws.Cells(2, 11) = "组成用量"
        Ws.Cells(2, 12) = "损耗率"
        Ws.Cells(2, 13) = "采购单位"
        Ws.Cells(2, 14) = "发料单位"
        Ws.Cells(2, 15) = "单位损耗量"
        Ws.Cells(2, 16) = "实际单位用量QPA"
        Ws.Cells(2, 17) = "BOM特性代码"
        Ws.Cells(2, 18) = "最少采购数量"                      '190301 add by Brady
        Ws.Cells(2, 19) = "交货前置期（天）"                  '190301 add by Brady
        oRng = Ws.Range("O2", "R2")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00000000_ "
        LineZ = 3
        '200506 add by Brady END
    End Sub


End Class