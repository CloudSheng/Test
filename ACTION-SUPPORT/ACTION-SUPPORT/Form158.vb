Public Class Form158
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim LineS1 As Int16 = 0
    Dim tYear As Int16 = 0
    Dim pYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim pMonth As Int16 = 0
    Dim lYear As Int16 = 0
    Dim tCurrency As String = String.Empty
    Dim ExchangeRate As Decimal = 0
    Dim ExchangeRate1 As Decimal = 0
    Dim gDatabase As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form158_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        If Today.Month < 10 Then
            TextBox1.Text = Today.Year & "0" & Today.Month
        Else
            TextBox1.Text = Today.Year & Today.Month
        End If
        mConnection.ConnectionString = Module1.OpenConnectionOfERPSUPPORT()
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mSQLS1.CommandTimeout = 600
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
                mSQLS2.CommandTimeout = 600
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
        Dim xPath As String = "C:\temp\ACAIS2.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If

        If TextBox1.Text.Length < 6 Then
            MsgBox("ERROR")
            Return
        End If
        tYear = Strings.Left(Me.TextBox1.Text, 4)
        pYear = tYear - 1
        tMonth = Strings.Right(Me.TextBox1.Text, 2)
        'pMonth = tMonth - 1
        'If pMonth = 0 Then
        'pMonth = 12
        'lYear = tYear - 1
        'Else
        'lYear = tYear
        'End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        Dim xPath As String = "C:\temp\ACAIS2.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        LineZ = 5
        AdjustExcelFormat()
        DoInputData("B2 :")
        LineZ += 1
        DoInputData("4005")
        LineZ += 1
        DoInputData("4030")
        LineZ += 1
        DoInputData("4040")
        LineZ += 1
        DoInputData("4061")
        LineZ += 1
        DoInputData("B3 :")
        LineZ += 1
        DoInputData("4011")
        LineZ += 1
        DoInputData("4013")
        LineZ += 1
        DoInputData("4015")
        LineZ += 1
        DoInputData("B4 :")
        LineZ += 1
        DoInputData("4025")
        LineZ += 1
        DoInputData("4038")
        LineZ += 1
        DoInputData("4831")
        LineZ += 1
        DoInputData("4838")
        LineZ += 1
        DoInputData("B5 :")
        LineZ += 1
        DoInputData("4830")
        LineZ += 1
        DoInputData("4832")
        LineZ += 1
        DoInputData("4833")
        LineZ += 1
        DoInputData("4834")
        LineZ += 1
        DoInputData("4899")
        LineZ += 1
        DoInputData("B6 :")
        LineZ += 1
        DoInputData("4430")
        LineZ += 1
        DoInputData("4405")
        LineZ += 1
        DoInputData("4440")
        LineZ += 1
        DoInputData("B9 :")
        LineZ += 1
        DoInputData("4860")
        LineZ += 1
        DoInputData("4920")
        LineZ += 1
        DoInputData("Purc")
        LineZ += 2
        'DoInputData("B2 :")
        'LineZ += 1
        DoInputData("B10 ")
        LineZ += 1
        DoInputData("B101")
        LineZ += 1
        DoInputData("5300")
        LineZ += 1
        DoInputData("B102")
        LineZ += 1
        DoInputData("5001")
        LineZ += 1
        DoInputData("5100")
        LineZ += 1
        DoInputData("5101")
        LineZ += 1
        DoInputData("5102")
        LineZ += 1
        DoInputData("5104")
        LineZ += 1
        DoInputData("5398")
        LineZ += 1
        DoInputData("5399")
        LineZ += 1
        DoInputData("B103")
        LineZ += 1
        DoInputData("5702")
        LineZ += 1
        DoInputData("5703")
        LineZ += 1
        DoInputData("5704")
        LineZ += 1
        DoInputData("5707")
        LineZ += 1
        DoInputData("5711")
        LineZ += 1
        DoInputData("B104")
        LineZ += 1
        DoInputData("5721")
        LineZ += 1
        DoInputData("5722")
        LineZ += 1
        DoInputData("5723")
        LineZ += 1
        DoInputData("5725")
        LineZ += 1
        DoInputData("5750")
        LineZ += 1
        DoInputData("B12 ")
        LineZ += 1
        DoInputData("5015")
        LineZ += 1
        DoInputData("Expe")
        LineZ += 1
        DoInputData("B13 ")
        LineZ += 1
        DoInputData("B131")
        LineZ += 1
        DoInputData("6200")
        LineZ += 1
        DoInputData("6201")
        LineZ += 1
        DoInputData("6211")
        LineZ += 1
        DoInputData("6212")
        LineZ += 1
        DoInputData("6300")
        LineZ += 1
        DoInputData("B133")
        LineZ += 1
        DoInputData("6407")
        LineZ += 1
        DoInputData("B134")
        LineZ += 1
        DoInputData("6500")
        LineZ += 1
        DoInputData("6501")
        LineZ += 1
        DoInputData("6660")
        LineZ += 1
        DoInputData("6670")
        LineZ += 1
        DoInputData("6680")
        LineZ += 1
        DoInputData("B135")
        LineZ += 1
        DoInputData("6700")
        LineZ += 1
        DoInputData("6740")
        LineZ += 1
        DoInputData("B14 ")
        LineZ += 1
        DoInputData("B141")
        LineZ += 1
        DoInputData("7030")
        LineZ += 1
        DoInputData("7035")
        LineZ += 1
        DoInputData("B15 ")
        LineZ += 1
        DoInputData("B151")
        LineZ += 1
        DoInputData("7100")
        LineZ += 2
        DoInputData("B16 ")
        LineZ += 1
        DoInputData("B161")
        LineZ += 1
        DoInputData("7205")
        LineZ += 1
        DoInputData("7215")
        LineZ += 1
        DoInputData("7230")
        LineZ += 1
        DoInputData("B162")
        LineZ += 1
        DoInputData("7301")
        LineZ += 1
        DoInputData("7302")
        LineZ += 1
        DoInputData("7306")
        LineZ += 1
        DoInputData("7309")
        LineZ += 1
        DoInputData("7310")
        LineZ += 1
        DoInputData("7311")
        LineZ += 1
        DoInputData("7312")
        LineZ += 1
        DoInputData("7313")
        LineZ += 1
        DoInputData("7319")
        LineZ += 1
        DoInputData("B163")
        LineZ += 1
        DoInputData("7340")
        LineZ += 1
        DoInputData("7343")
        LineZ += 1
        DoInputData("7350")
        LineZ += 1
        DoInputData("7351")
        LineZ += 1
        DoInputData("7352")
        LineZ += 1
        DoInputData("7353")
        LineZ += 1
        DoInputData("7354")
        LineZ += 1
        DoInputData("7356")
        LineZ += 1
        DoInputData("B164")
        LineZ += 1
        DoInputData("7370")
        LineZ += 1
        DoInputData("7395")
        LineZ += 1
        DoInputData("7380")
        LineZ += 1
        DoInputData("B166")
        LineZ += 1
        DoInputData("7600")
        LineZ += 1
        DoInputData("7651")
        LineZ += 1
        DoInputData("7652")
        LineZ += 1
        DoInputData("7671")
        LineZ += 1
        DoInputData("7672")
        LineZ += 1
        DoInputData("7690")
        LineZ += 1
        DoInputData("B167")
        LineZ += 1
        DoInputData("7700")
        LineZ += 1
        DoInputData("B168")
        LineZ += 1
        DoInputData("7400")
        LineZ += 1
        DoInputData("7450")
        LineZ += 1
        DoInputData("7451")
        LineZ += 1
        DoInputData("7480")
        LineZ += 1
        DoInputData("7206")
        LineZ += 1
        DoInputData("B169")
        LineZ += 1
        DoInputData("7740")
        LineZ += 1
        DoInputData("7741")
        LineZ += 1
        DoInputData("7755")
        LineZ += 1
        DoInputData("7760")
        LineZ += 1
        DoInputData("7761")
        LineZ += 1
        DoInputData("7762")
        LineZ += 1
        DoInputData("7782")
        LineZ += 1
        DoInputData("7785")
        LineZ += 1
        DoInputData("7790")
        LineZ += 1
        DoInputData("7795")
        LineZ += 1
        DoInputData("7801")
        LineZ += 1
        DoInputData("7802")
        LineZ += 1
        DoInputData("7806")
        LineZ += 1
        DoInputData("7817")
        LineZ += 1
        DoInputData("7840")
        LineZ += 1
        DoInputData("7870")
        LineZ += 1
        DoInputData("7880")
        LineZ += 1
        DoInputData("7890")
        LineZ += 1
        DoInputData("7990")
        LineZ += 1
        DoInputData("B170")
        LineZ += 1
        DoInputData("B171")
        LineZ += 1
        DoInputData("7320")
        LineZ += 1
        DoInputData("7321")
        LineZ += 1
        DoInputData("7322")
        LineZ += 1
        DoInputData("7323")
        LineZ += 1
        DoInputData("7324")
        LineZ += 1
        DoInputData("7325")
        LineZ += 1
        DoInputData("7326")
        LineZ += 1
        DoInputData("7327")
        LineZ += 1
        DoInputData("7329")
        LineZ += 1
        DoInputData("B172")
        LineZ += 1
        DoInputData("7330")
        LineZ += 1
        DoInputData("7331")
        LineZ += 1
        DoInputData("7332")
        LineZ += 1
        DoInputData("7333")
        LineZ += 1
        DoInputData("7334")
        LineZ += 1
        DoInputData("7335")
        LineZ += 1
        DoInputData("7336")
        LineZ += 1
        DoInputData("7337")
        LineZ += 1
        DoInputData("B173")
        LineZ += 1
        DoInputData("7440")
        LineZ += 1
        DoInputData("7441")
        LineZ += 1
        DoInputData("7442")
        LineZ += 1
        DoInputData("7443")
        LineZ += 1
        DoInputData("7444")
        LineZ += 1
        DoInputData("7445")
        LineZ += 1
        DoInputData("7446")
        LineZ += 1
        DoInputData("7447")
        LineZ += 1
        DoInputData("fina")
        LineZ += 1
        DoInputData("F6 :")
        LineZ += 1
        DoInputData("8320")
        LineZ += 1
        DoInputData("Othe")
        LineZ += 1
        DoInputData("S1 :")
        LineZ += 1
        DoInputData("8500")
        LineZ += 1
        DoInputData("Prof")
    End Sub
    Private Sub AdjustExcelFormat()
        If tMonth < 10 Then
            Ws.Cells(3, 4) = tYear & "/0" & tMonth
            Ws.Cells(3, 6) = pYear & "/0" & tMonth
        Else
            Ws.Cells(3, 4) = tYear & "/" & tMonth
            Ws.Cells(3, 6) = pYear & "/" & tMonth
        End If

    End Sub
    Private Sub DoInputData(ByVal ACC1 As String)

        mSQLS1.CommandText = "select isnull(Amount1,0) from acais where year1 = " & tYear & " and month1 = " & tMonth & " and acc1 = '" & ACC1 & "'"
        Dim V1 As Decimal = mSQLS1.ExecuteScalar()
        Ws.Cells(LineZ, 4) = V1

        mSQLS1.CommandText = "select isnull(Amount1,0) from acais where year1 = " & pYear & " and month1 = " & tMonth & " and acc1 = '" & ACC1 & "'"
        Dim V2 As Decimal = mSQLS1.ExecuteScalar()
        Ws.Cells(LineZ, 6) = V2

    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "P&L detail ledger " & tYear & tMonth
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
                MsgBox("Finished")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
End Class