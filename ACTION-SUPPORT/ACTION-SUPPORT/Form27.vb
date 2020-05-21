Imports Microsoft.Office.Interop.Excel.XlFileFormat
Public Class Form27
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim ArrayX1 As String() = {}
    Dim Position As Int16 = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form27_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
                MsgBox(ex.Message())
            End Try
        End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "MES_ERP_LIST"
        SaveFileDialog1.DefaultExt = ".xls"
        Dim SON As DialogResult = SaveFileDialog1.ShowDialog()
        If SON = DialogResult.OK Then
            Dim SFN As String = SaveFileDialog1.FileName
            Ws.SaveAs(SFN, XlFileFormat.xlExcel12)
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
    Private Sub ExportToExcel()
        mSQLS1.CommandText = "select station from station where station not in ('8888','9999','BLCK','SCRP')"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Array.Resize(ArrayX1, ArrayX1.Length + 1)
                ArrayX1.SetValue(mSQLReader.Item("station").ToString, Position)
                Position += 1
            End While
        End If
        mSQLReader.Close()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        Dim CheckFormat As String = String.Empty
        LineZ = 2
        mSQLS1.CommandText = "select model_station_paravalue.model,station,cf01,value from model_station_paravalue "
        mSQLS1.CommandText += "left join model_paravalue on parameter = 'ERP PN' and model_station_paravalue.model = model_paravalue.model "
        mSQLS1.CommandText += "where profilename = 'ERP' order by model,station"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                If String.IsNullOrEmpty(CheckFormat) Then
                    CheckFormat = mSQLReader.Item("model").ToString()
                    Ws.Cells(LineZ, 1) = mSQLReader.Item("value")
                    Ws.Cells(LineZ, 2) = mSQLReader.Item("model")
                    oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 2))
                    oRng.Interior.Color = Color.LightBlue
                End If
                If CheckFormat <> mSQLReader.Item("model").ToString() Then
                    LineZ += 1
                    CheckFormat = mSQLReader.Item("model").ToString()
                    Ws.Cells(LineZ, 1) = mSQLReader.Item("value")
                    Ws.Cells(LineZ, 2) = mSQLReader.Item("model")
                    oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 2))
                    oRng.Interior.Color = Color.LightBlue
                End If
                Position = Array.IndexOf(ArrayX1, mSQLReader.Item("station").ToString())
                Ws.Cells(LineZ, 3 + Position) = mSQLReader.Item("cf01")
            End While
        End If
        mSQLReader.Close()

        ' add by cloud 20180105
        Ws = xWorkBook.Sheets(2)
        AdjustExcelFormat1()
        mSQLS1.CommandText = "select model,model_station_paravalue.station,station.stationname_cn,cf01 from model_station_paravalue left join station on model_station_paravalue.station = station.station  where profilename = 'ERP' and cf01 is not null and cf01 <> ''"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("station")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("stationname_cn")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("cf01")
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
    End Sub

    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.ColumnWidth = 15
        oRng = Ws.Range("A1", "B1")
        oRng.EntireColumn.ColumnWidth = 40
        oRng.Interior.Color = Color.LightBlue
        Ws.Cells(1, 1) = "ERP_PN"
        Ws.Cells(1, 2) = "MODEL"
        For i As Integer = 0 To ArrayX1.Length - 1 Step 1
            Ws.Cells(1, 3 + i) = "'" & ArrayX1(i).ToString()
        Next
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.ColumnWidth = 15
        oRng = Ws.Range("A1", "B1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(1, 1) = "产品型号"
        Ws.Cells(1, 2) = "工站"
        Ws.Cells(1, 3) = "工站名称"
        Ws.Cells(1, 4) = "ERP料号"
        LineZ = 2
    End Sub
End Class