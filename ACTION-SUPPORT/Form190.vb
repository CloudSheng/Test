Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form190
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim tYear As Int16 = 0
    Dim xPath As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        
        xPath = "C:\temp\Template_Form190.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If

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
        tYear = Me.NumericUpDown1.Value

        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Form190_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        Me.NumericUpDown1.Value = Today.Year

    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        
        SaveFileDialog1.FileName = "电费占产品销售收入"
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
        xPath = "C:\temp\Template_Form190.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        For i As Int16 = 1 To 12 Step 1
            Ws.Cells(4, 1 + i) = tYear & "/" & i & "/01"
        Next
        LineZ = 5
        oCommand.CommandText = "Select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ("
        oCommand.CommandText += "Select (case when month(apa02) = 1 then sum(apa31) else 0 end) as t1,(case when month(apa02) = 2 then sum(apa31) else 0 end) as t2,(case when month(apa02) = 3 then sum(apa31) else 0 end) as t3,"
        oCommand.CommandText += "(case when month(apa02) = 4 then sum(apa31) else 0 end) as t4,(case when month(apa02) = 5 then sum(apa31) else 0 end) as t5,(case when month(apa02) = 6 then sum(apa31) else 0 end) as t6,"
        oCommand.CommandText += "(case when month(apa02) = 7 then sum(apa31) else 0 end) as t7,(case when month(apa02) = 8 then sum(apa31) else 0 end) as t8,(case when month(apa02) = 9 then sum(apa31) else 0 end) as t9,"
        oCommand.CommandText += "(case when month(apa02) = 10 then sum(apa31) else 0 end) as t10,(case when month(apa02) = 11 then sum(apa31) else 0 end) as t11,(case when month(apa02) = 12 then sum(apa31) else 0 end) as t12 "
        oCommand.CommandText += "from ( Select apa31,apa02 from apa_file where apa05 in ('615142','615049') and year(apa02) = " & tYear & " ) group by apa02 )"
        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount() - 1 Step 1
                    Ws.Cells(LineZ, i + 2) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()
        LineZ += 1
        oCommand.CommandText = "Select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ("
        oCommand.CommandText += "Select (case when aah03 = 1 then sum(c1) else 0 end) as t1,(case when aah03 = 2 then sum(c1) else 0 end) as t2,(case when aah03 = 3 then sum(c1) else 0 end) as t3,"
        oCommand.CommandText += "(case when aah03 = 4 then sum(c1) else 0 end) as t4,(case when aah03 = 5 then sum(c1) else 0 end) as t5,(case when aah03 = 6 then sum(c1) else 0 end) as t6,"
        oCommand.CommandText += "(case when aah03 = 7 then sum(c1) else 0 end) as t7,(case when aah03 = 8 then sum(c1) else 0 end) as t8,(case when aah03 = 9 then sum(c1) else 0 end) as t9,"
        oCommand.CommandText += "(case when aah03 = 10 then sum(c1) else 0 end) as t10,(case when aah03 = 11 then sum(c1) else 0 end) as t11,(case when aah03 = 12 then sum(c1) else 0 end) as t12 "
        oCommand.CommandText += "from ( Select (aah05 -aah04) as c1, aah03  from aah_file where aah01 in ('600101','600102') and aah02 = " & tYear & " ) group by aah03 )"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount() - 1 Step 1
                    Ws.Cells(LineZ, i + 2) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()
    End Sub
End Class