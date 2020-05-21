Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form303
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim Start1 As String = String.Empty
    Dim End1 As String = String.Empty
    Dim TotalPeriod As Int16 = 0
    Dim LineZ As Integer = 0
    Dim SC As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form303_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
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
        If Now.Month < 10 Then
            TextBox3.Text = Now.Year & "0" & Now.Month
            TextBox2.Text = Now.Year & "0" & Now.Month
        Else
            TextBox3.Text = Now.Year & Now.Month
            TextBox2.Text = Now.Year & Now.Month
        End If
        Label6.Text = 0
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Start1 = TextBox2.Text
        End1 = TextBox3.Text
        If String.IsNullOrEmpty(Start1) Or String.IsNullOrEmpty(End1) Then
            MsgBox("期间资料错误")
            Return
        End If
        If Len(Start1) <> 6 Or Len(End1) <> 6 Then
            MsgBox("月份资料为6码")
            Return
        End If
        If Conversion.Int(Start1) > Conversion.Int(End1) Then
            MsgBox("开时期间大于结束期间")
            Return
        End If
        If CheckBox1.Checked = False And CheckBox2.Checked = False Then
            MsgBox("输出资料有误")
            Return
        End If
        TotalPeriod = (Conversion.Int(Strings.Left(End1, 4)) - Conversion.Int(Strings.Left(Start1, 4))) * 12
        TotalPeriod += Conversion.Int(Strings.Right(End1, 2))
        TotalPeriod -= Conversion.Int(Strings.Right(Start1, 2))
        TotalPeriod += 1
        If TotalPeriod > 12 Then
            MsgBox("超出12个月")
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
        SC = TextBox1.Text
        Label6.Text = 0
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "采购&委外料件月入库数量"
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
        AdjustExcelFormat()
        oCommand.CommandText = "select ccc01,type,ima02,ima021,ima25"
        For i As Int16 = 1 To TotalPeriod Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        oCommand.CommandText += " from ( "
        If CheckBox1.Checked = True Then
            oCommand.CommandText += "select ccc01,'1' as type,ima02,ima021,ima25"
            For i As Int16 = 1 To TotalPeriod Step 1
                Dim TMonth As Int16 = Conversion.Int(Strings.Right(Start1, 2)) + i - 1
                Dim CT As String = String.Empty
                If TMonth > 12 Then
                    If TMonth - 12 < 10 Then
                        CT = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "0" & TMonth - 12
                    Else
                        CT = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "" & TMonth - 12
                    End If
                Else
                    If TMonth < 10 Then
                        CT = Conversion.Int(Strings.Left(Start1, 4)) & "0" & TMonth
                    Else
                        CT = Conversion.Int(Strings.Left(Start1, 4)) & TMonth
                    End If
                End If
                oCommand.CommandText += " ,(case when ccc02 || (case when length(ccc03) = 1 then 0 || ccc03 else to_char(ccc03) end)  = '"
                oCommand.CommandText += CT & "' then ccc211 else 0 end) as t" & i
            Next
            oCommand.CommandText += " from ccc_file,ima_file where ccc01 = ima01 and ima08 in ('P') and  ccc02 || (case when length(ccc03) = 1 then 0 || ccc03 else to_char(ccc03) end) between '"
            oCommand.CommandText += Start1 & "' and '" & End1 & "' "
            If Not String.IsNullOrEmpty(SC) Then
                oCommand.CommandText += " AND ccc01 = '" & SC & "' "
            End If
            
        End If


        If CheckBox2.Checked = True Then
            If CheckBox1.Checked = True Then
                oCommand.CommandText += "union all "
            End If
            oCommand.CommandText += "select ccc01,'2' as type,ima02,ima021,ima25"
            For i As Int16 = 1 To TotalPeriod Step 1
                Dim TMonth As Int16 = Conversion.Int(Strings.Right(Start1, 2)) + i - 1
                Dim CT As String = String.Empty
                If TMonth > 12 Then
                    If TMonth - 12 < 10 Then
                        CT = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "0" & TMonth - 12
                    Else
                        CT = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "" & TMonth - 12
                    End If
                Else
                    If TMonth < 10 Then
                        CT = Conversion.Int(Strings.Left(Start1, 4)) & "0" & TMonth
                    Else
                        CT = Conversion.Int(Strings.Left(Start1, 4)) & TMonth
                    End If
                End If
                oCommand.CommandText += " ,(case when ccc02 || (case when length(ccc03) = 1 then 0 || ccc03 else to_char(ccc03) end)  = '"
                oCommand.CommandText += CT & "' then ccc212 else 0 end) as t" & i
            Next
            oCommand.CommandText += " from ccc_file,ima_file where ccc01 = ima01 and ima08 in ('S') and  ccc02 || (case when length(ccc03) = 1 then 0 || ccc03 else to_char(ccc03) end) between '"
            oCommand.CommandText += Start1 & "' and '" & End1 & "' "
            If Not String.IsNullOrEmpty(SC) Then
                oCommand.CommandText += " AND ccc01 = '" & SC & "' "
            End If
            
        End If

        oCommand.CommandText += " ) group by ccc01,type,ima02,ima021,ima25 order by ccc01,type "
        oReader = oCommand.ExecuteReader()
        Dim TR As Decimal = 0
        If oReader.HasRows Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("ccc01")
                Ws.Cells(LineZ, 2) = oReader.Item("ima02")
                Ws.Cells(LineZ, 3) = oReader.Item("ima021")
                Select Case oReader.Item("type")
                    Case 1
                        Ws.Cells(LineZ, 4) = "P:采购料件"
                    Case 2
                        Ws.Cells(LineZ, 4) = "S:委外加工料件"

                End Select

                Ws.Cells(LineZ, 5) = oReader.Item("ima25")

                For i As Int16 = 1 To TotalPeriod Step 1
                    Ws.Cells(LineZ, 5 + i) = oReader.Item(4 + i)
                Next
                LineZ += 1
                TR += 1
                Label6.Text = TR
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.ColumnWidth = 20
        Ws.Cells(1, 1) = "Part no.料件编号"
        Ws.Cells(1, 2) = "Part_N 品名"
        Ws.Cells(1, 3) = "Spec.规格"
        Ws.Cells(1, 4) = "来源码"
        Ws.Cells(1, 5) = "库存单位"
        For i As Integer = 1 To TotalPeriod Step 1
            Dim TMonth As Int16 = Conversion.Int(Strings.Right(Start1, 2)) + i - 1
            If TMonth > 12 Then
                If TMonth - 12 < 10 Then
                    Ws.Cells(1, 5 + i) = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "/0" & TMonth - 12
                Else
                    Ws.Cells(1, 5 + i) = Conversion.Int(Strings.Left(Start1, 4)) + 1 & "/" & TMonth - 12
                End If
            Else
                If TMonth < 10 Then
                    Ws.Cells(1, 5 + i) = Conversion.Int(Strings.Left(Start1, 4)) & "/0" & TMonth
                Else
                    Ws.Cells(1, 5 + i) = Conversion.Int(Strings.Left(Start1, 4)) & "/" & TMonth
                End If
            End If
        Next
        LineZ = 2
    End Sub
End Class