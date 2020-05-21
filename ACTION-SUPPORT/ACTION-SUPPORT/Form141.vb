Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form141
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim tDate1 As Date
    Dim tDate2 As Date
    Dim ArrayS1() As String = {"102010010069", "102010010070", "102010010019", "102010010034", "102010010039", "102010010063", "102010010062", "102010010073", "102020020022"}
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form141_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
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
        tDate1 = Me.DateTimePicker1.Value
        tDate2 = Me.DateTimePicker2.Value
        If tDate2 < tDate1 Then
            MsgBox("Date Error")
            Return
        End If
        BackgroundWorker1.RunWorkerAsync()
        'ExportToExcel()
        'SaveExcel()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "纱料周用量统计表"
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
        For i As Int16 = 0 To ArrayS1.Length - 1 Step 1
            If i > 2 Then
                Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
            Else
                Ws = xWorkBook.Sheets(i + 1)
            End If
            Ws.Activate()
            'oCommand.CommandText = "select ima02 from ima_file where ima01 = '" & ArrayS1(i).ToString() & "'"
            'Dim WSN As String = oCommand.ExecuteScalar()
            'AdjustExcelFormat(WSN)
            AdjustExcelFormat(ArrayS1(i).ToString())
            Ws.Cells(2, 1) = "料号：" & ArrayS1(i).ToString()
            oCommand.CommandText = "select sfb05,sum(sfe16) as t1 from sfe_file left join sfb_file on sfe01 = sfb01 where sfe04 between to_date('"
            oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') and to_date('"
            oCommand.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') and sfe07 ='"
            oCommand.CommandText += ArrayS1(i).ToString() & "' group by sfb05"
            oReader = oCommand.ExecuteReader()
            If oReader.HasRows() Then
                While oReader.Read()
                    Ws.Cells(LineZ, 1) = oReader.Item("sfb05")
                    Ws.Cells(LineZ, 2) = oReader.Item("t1")
                    oCommand2.CommandText = "select nvl(sum(sfv09),0) as t1 from sfu_file left join sfv_file on sfu01 = sfv01 left join sfb_file on sfv11 = sfb01 where sfu02 between to_date('"
                    oCommand2.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') and to_date('"
                    oCommand2.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') and sfupost = 'Y' and sfb05 = '"
                    oCommand2.CommandText += oReader.Item("sfb05") & "'"
                    Dim AA As Decimal = oCommand2.ExecuteScalar()
                    Ws.Cells(LineZ, 4) = AA
                    LineZ += 1
                End While
                If LineZ > 4 Then
                    oRng = Ws.Range("A4", Ws.Cells(LineZ - 1, 4))
                    oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
                    oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
                    oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
                    oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
                    oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                    oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
                End If
            End If
            oReader.Close()
        Next
    End Sub
    Private Sub AdjustExcelFormat(ByVal hh As String)
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.HorizontalAlignment = xlcenter
        Ws.Name = hh
        Ws.Columns.ColumnWidth = 20.6
        Ws.Rows.RowHeight = 27
        oRng = Ws.Range("A1", "D1")
        oRng.Merge()
        oRng = Ws.Range("A2", "B2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlLeft
        oRng = Ws.Range("C2", "D2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlLeft
        Ws.Cells(1, 1) = "纱料用量统计表"
        Ws.Cells(2, 3) = "裁纱日期："
        Ws.Cells(3, 1) = "半成品料号"
        Ws.Cells(3, 2) = "标准单位用量/sqm"
        Ws.Cells(3, 3) = "实际单位用量/sqm"
        Ws.Cells(3, 4) = "半成品数量/pcs"

        oRng = Ws.Range("A1", "D3")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        LineZ = 4
    End Sub
End Class