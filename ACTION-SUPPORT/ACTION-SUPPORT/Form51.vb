Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form51
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim DStartN As Date
    Dim DstartE As Date
    Dim AllM As Long = 0
    Dim OCC01 As String = String.Empty
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form51_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        oCommand.CommandText = "SELECT OCC01,OCC02 FROM OCC_FILE WHERE OCCACTI = 'Y'"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Me.ComboBox1.Items.Add(oReader.Item(0).ToString() & "|" & oReader.Item(1).ToString())
            End While
        End If
        oReader.Close()
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
        DStartN = DateTimePicker1.Value
        DstartE = DateTimePicker2.Value
        If DstartE < DStartN Then
            MsgBox("Date Error")
            Return
        End If
        AllM = DateDiff(DateInterval.Month, DStartN, DstartE) + 1
        ' 客户
        If Not IsNothing(ComboBox1.SelectedItem) Then
            OCC01 = ComboBox1.SelectedItem.ToString()
            Dim stCount As Int16 = Strings.InStr(OCC01, "|")
            If stCount > 0 Then
                OCC01 = Strings.Left(OCC01, stCount - 1)
            End If
        End If
        'MsgBox(OCC01)
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "DAC_SALES_OVERVIEW"
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
        Ws.Name = "DAC monthly sales overview"
        AdjustExcelFormat()
        oCommand.CommandText = "select distinct oga03,oga032 from oga_file where oga02 between to_date('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ogapost = 'Y' " 'order by oga03"
        If Not String.IsNullOrEmpty(OCC01) Then
            oCommand.CommandText += " AND oga03 = '" & OCC01 & "' "
        End If
        oCommand.CommandText += " order by oga03"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("oga032")
                Dim D1 As Date = DStartN
                Dim D2 As Date = Convert.ToDateTime(DStartN.Year & "/" & DStartN.Month & "/01").AddMonths(1).AddDays(-1)
                For i As Integer = 1 To AllM Step 1
                    If i = AllM Then
                        D2 = DstartE
                    End If
                    oCommand2.CommandText = "select nvl(round(sum(ogb14 * oga24),2),0) from oga_file,ogb_file where oga01 = ogb01 and oga02 between to_date('"
                    oCommand2.CommandText += D1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                    oCommand2.CommandText += D2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ogapost = 'Y' and oga03 = '" & oReader.Item("oga03") & "'"
                    Dim GetSum As Decimal = oCommand2.ExecuteScalar()
                    If GetSum > 0 Then
                        Ws.Cells(LineZ, i + 1) = GetSum
                    End If
                    If i = 1 Then
                        D1 = D2.AddDays(1)
                    Else
                        D1 = D1.AddMonths(1)
                    End If
                    D2 = D1.AddMonths(1).AddDays(-1)
                Next
                LineZ += 1
            End While
        End If
        oReader.Close()
        ' 加總
        Ws.Cells(LineZ, 1) = "sum"
        Ws.Cells(LineZ, 2) = "=SUM(B2:B" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 2))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 1 + AllM)), Type:=xlFillDefault)
        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 1 + AllM))
        oRng.Interior.Color = Color.Yellow

        ' 第二頁
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        Ws.Name = "DAC monthly returned overview"
        AdjustExcelFormat()
        oCommand.CommandText = "select distinct oha03,oha032 from oha_file where oha02 between to_date('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ohapost = 'Y' " 'order by oga03"
        If Not String.IsNullOrEmpty(OCC01) Then
            oCommand.CommandText += " AND oha03 = '" & OCC01 & "' "
        End If
        oCommand.CommandText += " order by oha03"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("oha032")
                Dim D1 As Date = DStartN
                Dim D2 As Date = Convert.ToDateTime(DStartN.Year & "/" & DStartN.Month & "/01").AddMonths(1).AddDays(-1)
                For i As Integer = 1 To AllM Step 1
                    If i = AllM Then
                        D2 = DstartE
                    End If
                    oCommand2.CommandText = "select nvl(round(sum(ohb14 * oha24 * -1),2),0) from oha_file,ohb_file where oha01 = ohb01 and oha02 between to_date('"
                    oCommand2.CommandText += D1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                    oCommand2.CommandText += D2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ohapost = 'Y' and oha03 = '" & oReader.Item("oha03") & "'"
                    Dim GetSum As Decimal = oCommand2.ExecuteScalar()
                    If GetSum < 0 Then
                        Ws.Cells(LineZ, i + 1) = GetSum
                    End If
                    If i = 1 Then
                        D1 = D2.AddDays(1)
                    Else
                        D1 = D1.AddMonths(1)
                    End If
                    D2 = D1.AddMonths(1).AddDays(-1)
                Next
                LineZ += 1
            End While
        End If
        oReader.Close()
        ' 加總
        Ws.Cells(LineZ, 1) = "sum"
        Ws.Cells(LineZ, 2) = "=SUM(B2:B" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 2))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 1 + AllM)), Type:=xlFillDefault)
        oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ, 1 + AllM))
        oRng.Interior.Color = Color.Yellow
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75

        Ws.Cells(1, 1) = "customer"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 30
        Dim TempMonth As Date = DStartN
        For i As Integer = 1 To AllM Step 1
            Ws.Cells(1, 1 + i) = DStartN.AddMonths(i - 1).ToString("yyyy/MM")
        Next
        oRng = Ws.Range(Ws.Cells(1, 2), Ws.Cells(1, 1 + AllM))
        oRng.EntireColumn.ColumnWidth = 23
        oRng.EntireColumn.NumberFormatLocal = "_ [$￥-804]* #,##0.00_ ;_ [$￥-804]* -#,##0.00_ ;_ [$￥-804]* ""-""??_ ;_ @_ "
        oRng.NumberFormatLocal = "@"
        LineZ = 2
    End Sub
End Class