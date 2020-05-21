Public Class Form170
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim Linez As Integer = 0
    Dim Year1 As Int16 = 0
    Dim Week1 As Int16 = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form170_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            xExcel = New Microsoft.Office.Interop.Excel.Application
            xWorkBook = xExcel.Workbooks.Open(ExcelPath)
            Ws = xWorkBook.Sheets(1)
            Linez = 3
        Else
            Return
        End If

        oCommand.CommandText = "truncate table dac_receive_plan"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

        Me.ProgressBar1.Value = 1
        oCommand.CommandText = "select azn02 from azn_file where azn01 = to_date('" & Today.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        Year1 = oCommand.ExecuteScalar()
        oCommand.CommandText = "select azn05 from azn_file where azn01 = to_date('" & Today.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        Week1 = oCommand.ExecuteScalar()

        Dim BB As Integer = Ws.UsedRange.Rows.Count
        Me.ProgressBar1.Maximum = BB

        Dim CC As Integer = Ws.UsedRange.Columns.Count
        For i As Integer = 3 To BB Step 1
            ProgressBar1.Value += 1
            oRng = Ws.Range(Ws.Cells(i, 2), Ws.Cells(i, 2))
            If Not IsDBNull(oRng.Value) Or IsNothing(oRng.Value) Then
                Dim PN As String = oRng.Value
                For j As Integer = 7 To 28 Step 1
                    oRng = Ws.Range(Ws.Cells(i, j), Ws.Cells(i, j))
                    If Not (IsDBNull(oRng.Value) Or IsNothing(oRng.Value)) Then
                        Dim Week2 As Int16 = Week1 + j - 7
                        Dim Year2 As Int16 = Year1
                        If Week2 > 52 Then
                            Week2 = Week2 - 52
                            Year2 = Year1 + 1
                        End If
                        oCommand.CommandText = "INSERT INTO dac_receive_plan VALUES ('" & PN & "'," & Year2 & "," & Week2 & "," & oRng.Value & ")"
                        Try
                            oCommand.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.Message())
                            Return
                        End Try
                    End If
                Next
            Else
                Continue For
            End If

        Next
        xWorkBook.Close()
        xExcel.Quit()
        Module1.KillExcelProcess(OldExcel)
        MsgBox("Done")

    End Sub
End Class