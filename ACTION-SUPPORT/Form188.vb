Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form188
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim mSQLReader2 As SqlClient.SqlDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form188_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfRDMes()
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Dim xPath As String = "C:\temp\RD_SN_Template.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        Dim xPath As String = "C:\temp\RD_SN_Template.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Open(XPath)
        Ws = xWorkBook.Sheets(1)
        'Ws.Name = "SN"
        LineZ = 2
        Dim SSA As String = String.Empty
        mSQLS1.CommandText = "Select tracking.sn, station, timein from Tracking left join sn on Tracking.sn = sn.sn where sn.updatedstation <> '9999' Group by Tracking.sn,Tracking.timein,Tracking.station order by sn ,station"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                If SSA <> mSQLReader.Item(0) Then
                    LineZ += 1
                    Ws.Cells(LineZ, 1) = mSQLReader.Item(0)
                    SSA = mSQLReader.Item(0)
                End If
                Select Case mSQLReader.Item(1)
                    Case "0080"
                        Ws.Cells(LineZ, 2) = mSQLReader.Item(2)
                    Case "0112"
                        Ws.Cells(LineZ, 3) = mSQLReader.Item(2)
                    Case "0180"
                        Ws.Cells(LineZ, 4) = mSQLReader.Item(2)
                    Case "0330"
                        Ws.Cells(LineZ, 5) = mSQLReader.Item(2)
                    Case "0390"
                        Ws.Cells(LineZ, 6) = mSQLReader.Item(2)
                    Case "0395"
                        Ws.Cells(LineZ, 7) = mSQLReader.Item(2)
                    Case "0380"
                        Ws.Cells(LineZ, 8) = mSQLReader.Item(2)
                    Case "0400"
                        Ws.Cells(LineZ, 9) = mSQLReader.Item(2)
                    Case "0478"
                        Ws.Cells(LineZ, 10) = mSQLReader.Item(2)
                    Case "0490"
                        Ws.Cells(LineZ, 11) = mSQLReader.Item(2)
                    Case "0405"
                        Ws.Cells(LineZ, 12) = mSQLReader.Item(2)
                    Case "0455"
                        Ws.Cells(LineZ, 13) = mSQLReader.Item(2)
                    Case "0629"
                        Ws.Cells(LineZ, 14) = mSQLReader.Item(2)
                    Case "0650"
                        Ws.Cells(LineZ, 15) = mSQLReader.Item(2)
                    Case "0665"
                        Ws.Cells(LineZ, 16) = mSQLReader.Item(2)
                    Case "0670"
                        Ws.Cells(LineZ, 17) = mSQLReader.Item(2)
                    Case "0673"
                        Ws.Cells(LineZ, 18) = mSQLReader.Item(2)
                    Case "0690"
                        Ws.Cells(LineZ, 19) = mSQLReader.Item(2)
                    Case "0730"
                        Ws.Cells(LineZ, 20) = mSQLReader.Item(2)
                    Case Else
                        Continue While
                End Select
            End While
            ' 劃線
            oRng = Ws.Range("A3", Ws.Cells(LineZ, 20))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        mSQLReader.Close()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "RD_SN_Record"
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
        If mConnection.State = ConnectionState.Open Then
            Try
                mConnection.Close()
                Module1.KillExcelProcess(OldExcel)
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    
End Class