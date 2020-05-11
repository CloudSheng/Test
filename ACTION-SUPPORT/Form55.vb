Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form55
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim mSQLReader2 As SqlClient.SqlDataReader
    Dim TYear As String = String.Empty
    Dim TMonth As String = String.Empty
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim ptime As String = String.Empty
    Dim MaxDetailCount As Int16 = 0
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim PageCount As Integer = 0
    Dim DateCount As Integer = 0
    Dim FirstDate As Date
    Dim ModelCheck As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form55_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        If Now.Month > 9 Then
            TextBox1.Text = Now.Year & Now.Month
        Else
            TextBox1.Text = Now.Year & "0" & Now.Month
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
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        TYear = Strings.Left(TextBox1.Text, 4)
        TMonth = Strings.Right(TextBox1.Text, 2)
        TimeS1 = Convert.ToDateTime(TYear & "/" & TMonth & "/01")
        TimeS2 = TimeS1.AddMonths(1).AddDays(-1)
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        mSQLS2.CommandText = "DROP TABLE temp_report"
        Try
            mSQLS2.ExecuteNonQuery()
        Catch ex As Exception

        End Try
        mSQLS2.CommandText = "CREATE TABLE temp_report(CustomerName nvarchar(30),ModelID nvarchar(30),SN nvarchar(30),"
        mSQLS2.CommandText += "DayPlus INT,UserID nvarchar(7))"
        Try
            mSQLS2.ExecuteNonQuery()
        Catch ex As Exception

        End Try
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        AdjustExcelFormat()
        mSQLS2.CommandText = "select sn,modelid,UserID,CustomerName,Convert(nvarchar(100),Timeout,21) as Timeout from efd.dbo.tracking where TimeOUT between '"
        mSQLS2.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS2.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and StationId in ('0630') order by CustomerName "
        mSQLReader2 = mSQLS2.ExecuteReader()
        If mSQLReader2.HasRows() Then
            While mSQLReader2.Read()
                mSQLS1.CommandText = "select TOP 1 result from Efd.dbo.Tracking where stationid = '0640' and sn = '"
                mSQLS1.CommandText += mSQLReader2.Item("sn") & "' and timeout > '" & mSQLReader2.Item("timeout") & "' order by timeout "
                Dim CheckFlag As String = mSQLS1.ExecuteScalar()
                If IsDBNull(CheckFlag) Or CheckFlag = "F" Then

                Else
                    ' INSERT INTO TEMP
                    InsertIntoTemp(mSQLReader2.Item("CustomerName"), mSQLReader2.Item("modelid"), mSQLReader2.Item("sn"), mSQLReader2.Item("Timeout"), mSQLReader2.Item("UserID"))
                End If

            End While
        End If
        mSQLReader2.Close()
        mSQLS2.CommandText = "SELECT CustomerName,UserID,ModelID,DayPlus,Count(SN) as t1 FROM temp_report group by CustomerName,UserID,ModelID,DayPlus Order by CustomerName,DayPlus,ModelID"
        mSQLReader2 = mSQLS2.ExecuteReader()
        If mSQLReader2.HasRows() Then
            While mSQLReader2.Read()
                Ws.Cells(LineZ, 1) = mSQLReader2.Item("UserID")
                Ws.Cells(LineZ, 2) = mSQLReader2.Item("CustomerName")
                Ws.Cells(LineZ, 3) = mSQLReader2.Item("ModelID")
                Ws.Cells(LineZ, 5) = GetIE(mSQLReader2.Item("ModelID"))
                Ws.Cells(LineZ, 6) = mSQLReader2.Item("t1")
                LineZ += 1
            End While
        End If
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        xExcel.ActiveWindow.DisplayGridlines = False
        Ws.Columns.EntireColumn.ColumnWidth = 15
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.WrapText = True
        Ws.Cells(1, 1) = "姓名"
        Ws.Cells(1, 2) = "客户"
        Ws.Cells(1, 3) = "产品型号"
        Ws.Cells(1, 4) = "ERP料号"
        Ws.Cells(1, 5) = "ERP IE工时"
        Ws.Cells(1, 6) = "月产出数量总计"
        LineZ = 2
    End Sub
    Private Sub InsertIntoTemp(ByVal c1 As String, ByVal c2 As String, c3 As String, c4 As Date, c5 As String)
        Dim DP As Integer = Decimal.Ceiling((c4 - TimeS1).TotalDays)
        Dim mSQLS4 As New SqlClient.SqlCommand
        mSQLS4.Connection = mConnection
        mSQLS4.CommandType = CommandType.Text
        mSQLS4.CommandText = "INSERT INTO temp_report VALUES ('" & c1 & "','" & c2 & "','" & c3 & "'," & DP & ",'" & c5 & "')"
        Try
            mSQLS4.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
    End Sub
    Private Function GetIE(ByVal ModelID As String)
        Dim mSQLS4 As New SqlClient.SqlCommand
        mSQLS4.Connection = mConnection
        mSQLS4.CommandType = CommandType.Text
        mSQLS4.CommandText = "select TaktTimeVal from efd.dbo.TaktTime where StationGroupId = 2 and ModelID = '" & ModelID & "'"
        Dim DC As Decimal = mSQLS4.ExecuteScalar()
        If IsDBNull(DC) Then
            DC = 0
        End If
        Return DC
    End Function
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        'If HaveReport > 0 Then
        SaveExcel()
        'End If
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Polish_MonthEffiency"
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
End Class