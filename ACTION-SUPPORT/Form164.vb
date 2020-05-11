Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel
Public Class Form164
    Dim mConnection As New SqlClient.SqlConnection
    Dim mConnection2 As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form164_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
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
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
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
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
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
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "MES标签信息与ERP规格检查报表"
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
        oCommand.CommandText = "select ima01,ima021 from ima_file where imaacti = 'Y' and ima08 = 'M' and ima06 = '103' and ima021 is not null and ima25 = 'PCS'"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                mSQLS1.CommandText = "SELECT COUNT(MODEL) FROM model_paravalue where VALUE = '" & oReader.Item("ima01") & "' and parameter = 'ERP PN'"
                Dim CS As Int16 = mSQLS1.ExecuteScalar()
                If CS = 0 Then
                    Continue While
                End If
                mSQLS1.CommandText = "select model_type,p2.value from model left join model_paravalue p2 on model.model = p2.model and p2.parameter = 'Customer PN' where model.model in ( "
                mSQLS1.CommandText += "SELECT MODEL FROM model_paravalue where VALUE = '" & oReader.Item("ima01") & "' and parameter = 'ERP PN' ) "
                mSQLReader = mSQLS1.ExecuteReader()
                If mSQLReader.HasRows() Then
                    mSQLReader.Read()
                    Ws.Cells(LineZ, 1) = oReader.Item("ima01")
                    Ws.Cells(LineZ, 2) = oReader.Item("ima021")
                    Ws.Cells(LineZ, 3) = mSQLReader.Item("model_type")
                    Ws.Cells(LineZ, 4) = mSQLReader.Item("value")
                    Ws.Cells(LineZ, 5) = "=IF(B" & LineZ & "=D" & LineZ & ",""通过"",""不一致"")"
                    If Not IsDBNull(mSQLReader.Item("value")) Then
                        If oReader.Item("ima021") = mSQLReader.Item("value") Then
                            oRng = Ws.Range(Ws.Cells(LineZ, 5), Ws.Cells(LineZ, 5))
                            oRng.Interior.Color = Color.Green
                        Else
                            oRng = Ws.Range(Ws.Cells(LineZ, 5), Ws.Cells(LineZ, 5))
                            oRng.Interior.Color = Color.Red
                        End If
                    Else
                        oRng = Ws.Range(Ws.Cells(LineZ, 5), Ws.Cells(LineZ, 5))
                        oRng.Interior.Color = Color.Red
                    End If
                    
                    LineZ += 1
                End If
                mSQLReader.Close()
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 30
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "E1")
        oRng.Merge()
        oRng.Font.Size = 22
        Ws.Cells(1, 1) = "MES产品标签与ERP客户产品编号对应一览表"
        Ws.Cells(2, 1) = "ERP料件编号"
        Ws.Cells(2, 2) = "规格"
        Ws.Cells(2, 3) = "客户名称"
        Ws.Cells(2, 4) = "产品编号信息"
        Ws.Cells(2, 5) = "判定"
        LineZ = 3
    End Sub
End Class