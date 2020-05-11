Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form159
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form159_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Dim xPath As String = "C:\temp\AC-R-11-03 Key Spare Part List for production machine.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If

        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        Dim xPath As String = "C:\temp\AC-R-11-03 Key Spare Part List for production machine.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        LineZ = 4
        oCommand.CommandText = "select critical_part.ima01,ima02,ima021,ima25,ima27,nvl(sum(img10),0) as img10, ima49  from critical_part left join ima_file on critical_part.ima01 = ima_file.ima01 "
        oCommand.CommandText += "left join img_file on critical_part.ima01 =img01 and img10 > 0 group by critical_part.ima01,ima02,ima021,ima25,ima27,ima49"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            Dim LineX As Integer = 1
            While oReader.Read()
                Ws.Cells(LineZ, 1) = LineX
                Ws.Cells(LineZ, 2) = oReader.Item("ima01")
                Ws.Cells(LineZ, 3) = oReader.Item("ima02")
                Ws.Cells(LineZ, 4) = oReader.Item("ima021")
                Ws.Cells(LineZ, 8) = oReader.Item("ima25")
                Ws.Cells(LineZ, 10) = oReader.Item("ima27")
                Ws.Cells(LineZ, 12) = oReader.Item("img10")
                Ws.Cells(LineZ, 13) = "=L" & LineZ & "-J" & LineZ
                Ws.Cells(LineZ, 14) = oReader.Item("ima49")
                LineX += 1
                LineZ += 1
            End While
        End If
        oReader.Close()
        oRng = Ws.Range("A4", Ws.Cells(LineZ - 1, 17))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "AC-R-11-03 Key Spare Part List for production machine " & Now.ToString("yyMMdd")
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
End Class