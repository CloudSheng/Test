Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form18
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader

    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form18_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
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
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Block_Report"
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
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        LineZ = 2
        'mSQLS1.CommandText = "SELECT lot.model,model.modelname,sn.SN,failure.failtime,failure.failstation,station.stationname_cn ,failure.defect,defect.desc_th ,failure.users,users.name  FROM SN "
        'mSQLS1.CommandText += "LEFT JOIN lot on sn.lot = lot.lot LEFT JOIN model on lot.model = model.model LEFT JOIN FAILURE ON SN.SN = FAILURE.SN AND failure.rework = 'BLCK' LEFT JOIN station on failure.failstation = station.station "
        'mSQLS1.CommandText += "left join defect on failure.defect = defect.defect left join users on failure.users = users.id "
        mSQLS1.CommandText = "SELECT lot.model,model.modelname,sn.SN,sn.updatedstation,sn.updatedtime,sn.remark FROM SN LEFT JOIN lot on sn.lot = lot.lot LEFT JOIN model on lot.model = model.model "
        mSQLS1.CommandText += "WHERE SN.BLOCK = 'Y'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("modelname")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("sn")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("updatedtime")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("updatedstation") '& " " & mSQLReader.Item("stationname_cn")
                Ws.Cells(LineZ, 6) = mSQLReader.Item("remark")
                'Ws.Cells(LineZ, 6) = mSQLReader.Item("defect") & " " & mSQLReader.Item("desc_th")
                'Ws.Cells(LineZ, 7) = mSQLReader.Item("users") & " " & mSQLReader.Item("name")
                LineZ += 1
            End While
        End If
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "隔离明細"
        Ws.Columns.EntireColumn.ColumnWidth = 30
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Cells(1, 1) = "Part No."
        Ws.Cells(1, 2) = "Part Description"
        Ws.Cells(1, 3) = "SN"
        Ws.Cells(1, 4) = "LOCK Time"
        Ws.Cells(1, 5) = "LOCK  Station"
        Ws.Cells(1, 6) = "Remark"
        Dim oRange As Microsoft.Office.Interop.Excel.Range = Ws.Range("E1", "E1")
        oRange.EntireColumn.NumberFormatLocal = "@"
        'Ws.Cells(1, 6) = "Defect"
        'Ws.Cells(1, 7) = "By"

    End Sub
End Class