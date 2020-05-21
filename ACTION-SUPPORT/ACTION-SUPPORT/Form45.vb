Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form45
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim mSQLReader2 As SqlClient.SqlDataReader
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

    Private Sub Form45_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        ptime = Today.AddDays(-1).ToString("yyyy/MM/dd")
        ptime = ptime & " 08:00:00"
        Me.DateTimePicker1.Value = Convert.ToDateTime(ptime)
        Me.DateTimePicker2.Value = Convert.ToDateTime(ptime).AddDays(1).AddSeconds(-1)
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
        TimeS1 = DateTimePicker1.Value
        TimeS2 = DateTimePicker2.Value
        DateCount = Decimal.Ceiling((TimeS2 - TimeS1).TotalDays)
        If DateCount > 7 Or DateCount <= 0 Then
            MsgBox("时间设定有误")
            Return
        End If
        FirstDate = TimeS1.Date

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
        mSQLS1.CommandText = "SELECT distinct userid FROM EFD.dbo.Tracking Where TimeOUT Between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' AND StationId = '0630' Order by userid "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                PageCount += 1
                If PageCount > 3 Then
                    Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
                Else
                    Ws = xWorkBook.Sheets(PageCount)
                End If
                Ws.Activate()
                AdjustExcelFormat()
                Ws.Cells(3, 1) = mSQLReader.Item("userid")
                Dim UN As String = GetUserName(mSQLReader.Item("userid"))
                Ws.Cells(3, 2) = UN
                Ws.Name = mSQLReader.Item("userid") & UN
                ' 開始處理主要資料
                GetDetailData(mSQLReader.Item("userid"))
            End While
        End If
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        'If HaveReport > 0 Then
        SaveExcel()
        'End If
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Polish_DailyEffiency"
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
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        xExcel.ActiveWindow.DisplayGridlines = False
        Ws.Columns.EntireColumn.ColumnWidth = 15
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.WrapText = True
        oRng = Ws.Range("C1", "G1")
        oRng.Merge()
        oRng = Ws.Range("C2", "G2")
        oRng.Merge()
        oRng = Ws.Range("A4", "A4")
        oRng.EntireRow.RowHeight = 66
        oRng = Ws.Range("A1", "AH1")
        oRng.EntireColumn.ColumnWidth = 18.75
        Ws.Cells(1, 3) = TimeS1.ToShortDateString() & "-" & TimeS2.ToShortDateString()
        Ws.Cells(2, 3) = "Polish Section"
        oRng = Ws.Range("A3", "B3")
        oRng.Font.Color = Color.Blue
        oRng.NumberFormatLocal = "@"
        oRng = Ws.Range("C3", "E3")
        oRng.Interior.Color = Color.Red
        oRng.Font.Color = Color.White
        Ws.Cells(4, 1) = "客户 Name"
        Ws.Cells(4, 2) = "产品型号" & Chr(10) & "Product Model"
        Ws.Cells(4, 3) = "难度系数等级" & Chr(10) & "Degree of difficulty"
        Ws.Cells(3, 3) = "本周實際效率:"
        Ws.Cells(3, 5) = "本周KPA效率:"
        Ws.Cells(4, 4) = "積分參數" & Chr(10) & "Ratio of difficulty"
        Ws.Cells(4, 5) = "工時(分)" & Chr(10) & "IE Takttime(m)"
        'Ws.Cells(4, 6) = "日產能/人" & Chr(10) & "Daily output pcs/worker"
        FirstDate = TimeS1.Date
        For i As Integer = 1 To DateCount Step 1
            Ws.Cells(3, 6 + (i - 1) * 4) = FirstDate
            oRng = Ws.Range(Ws.Cells(3, 6 + (i - 1) * 4), Ws.Cells(3, 7 + (i - 1) * 4))
            oRng.Interior.Color = Color.LightGray
            Ws.Cells(3, 8 + (i - 1) * 4) = "上班時數:"
            Ws.Cells(3, 9 + (i - 1) * 4) = 11
            oRng = Ws.Range(Ws.Cells(3, 9 + (i - 1) * 4), Ws.Cells(3, 9 + (i - 1) * 4))
            oRng.Interior.Color = Color.Yellow
            Ws.Cells(4, 6 + (i - 1) * 4) = "產出" & Chr(10) & "Actual output pcs/worker"
            Ws.Cells(4, 7 + (i - 1) * 4) = "IE 工时" & Chr(10) & "IE hours"
            Ws.Cells(4, 8 + (i - 1) * 4) = "實際效率" & Chr(10) & "IE efficiency"
            Ws.Cells(4, 9 + (i - 1) * 4) = "KPA 效率" & Chr(10) & "KPA efficiency"
            oRng = Ws.Range(Ws.Cells(3, 8 + (i - 1) * 4), Ws.Cells(3, 9 + (i - 1) * 4))
            oRng.EntireColumn.NumberFormatLocal = "0.0000%"
            oRng = Ws.Range(Ws.Cells(3, 9 + (i - 1) * 4), Ws.Cells(3, 9 + (i - 1) * 4))
            oRng.NumberFormatLocal = "0.00_ "
            FirstDate = FirstDate.AddDays(1)
        Next
        LineZ = 5
    End Sub
    Private Function GetUserName(ByVal Userid As String)
        mSQLS2.CommandText = "SELECT name FROM users WHERE id = '" & Userid & "'"
        Dim UN As String = mSQLS2.ExecuteScalar()
        Return UN
    End Function
    Private Sub GetDetailData(ByVal UserID As String)
        Dim mSQLS3 As New SqlClient.SqlCommand

        mSQLS3.Connection = mConnection
        mSQLS3.CommandType = CommandType.Text

        mSQLS2.CommandText = "select sn,modelid,CustomerName,Convert(nvarchar(100),Timeout,21) as Timeout from efd.dbo.tracking where TimeOUT between '"
        mSQLS2.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS2.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and StationId in ('0630','0635') and userid = '"
        mSQLS2.CommandText += UserID & "' order by CustomerName "
        mSQLReader2 = mSQLS2.ExecuteReader()
        If mSQLReader2.HasRows() Then
            While mSQLReader2.Read()
                mSQLS3.CommandText = "select TOP 1 result from Efd.dbo.Tracking where stationid = '0640' and sn = '"
                mSQLS3.CommandText += mSQLReader2.Item("sn") & "' and timeout > '" & mSQLReader2.Item("timeout") & "' order by timeout "
                Dim CheckFlag As String = mSQLS3.ExecuteScalar()
                If IsDBNull(CheckFlag) Or CheckFlag = "F" Then

                Else
                    ' INSERT INTO TEMP
                    InsertIntoTemp(mSQLReader2.Item("CustomerName"), mSQLReader2.Item("modelid"), mSQLReader2.Item("sn"), mSQLReader2.Item("Timeout"), UserID)
                End If

            End While
        End If
        mSQLReader2.Close()

        ' 寫完之後, 讀入Excel
        ' 先清空 ModelCheck
        ModelCheck = String.Empty
        mSQLS2.CommandText = "SELECT CustomerName,ModelID,DayPlus,Count(SN) as t1 FROM temp_report where UserID = '" & UserID & "' group by CustomerName,ModelID,DayPlus Order by CustomerName,DayPlus,ModelID"
        mSQLReader2 = mSQLS2.ExecuteReader()
        If mSQLReader2.HasRows() Then
            While mSQLReader2.Read()
                If String.IsNullOrEmpty(ModelCheck) Then
                    ModelCheck = mSQLReader2.Item("ModelID")
                End If
                If ModelCheck <> mSQLReader2.Item("ModelID") Then
                    LineZ += 1
                End If
                Ws.Cells(LineZ, 1) = mSQLReader2.Item("CustomerName")
                Ws.Cells(LineZ, 2) = mSQLReader2.Item("ModelID")
                Dim DC As Decimal = GetDifficult(mSQLReader2.Item("ModelID"))
                Ws.Cells(LineZ, 3) = DC
                'Ws.Cells(LineZ, 3)
                Dim Ra As Decimal = GetRatio(mSQLReader2.Item("ModelID"))
                Ws.Cells(LineZ, 4) = Ra
                Dim PositionCheck As Integer = mSQLReader2.Item("DayPlus")
                Dim IETime As Decimal = GetIE(mSQLReader2.Item("ModelID"))
                Ws.Cells(LineZ, 5) = IETime
                Ws.Cells(LineZ, 6 + (PositionCheck - 1) * 4) = mSQLReader2.Item("t1")
                Ws.Cells(LineZ, 7 + (PositionCheck - 1) * 4) = Decimal.Round(mSQLReader2.Item("t1") * IETime / 60, 2)
                Select Case PositionCheck
                    Case 1
                        Ws.Cells(LineZ, 8) = "=G" & LineZ & "/I3"
                        Ws.Cells(LineZ, 9) = "=H" & LineZ & "*D" & LineZ
                    Case 2
                        Ws.Cells(LineZ, 12) = "=K" & LineZ & "/M3"
                        Ws.Cells(LineZ, 13) = "=L" & LineZ & "*D" & LineZ
                    Case 3
                        Ws.Cells(LineZ, 16) = "=O" & LineZ & "/Q3"
                        Ws.Cells(LineZ, 17) = "=P" & LineZ & "*D" & LineZ
                    Case 4
                        Ws.Cells(LineZ, 20) = "=S" & LineZ & "/U3"
                        Ws.Cells(LineZ, 21) = "=T" & LineZ & "*D" & LineZ
                    Case 5
                        Ws.Cells(LineZ, 24) = "=W" & LineZ & "/Y3"
                        Ws.Cells(LineZ, 25) = "=X" & LineZ & "*D" & LineZ
                    Case 6
                        Ws.Cells(LineZ, 28) = "=AA" & LineZ & "/AC3"
                        Ws.Cells(LineZ, 29) = "=AB" & LineZ & "*D" & LineZ
                    Case 7
                        Ws.Cells(LineZ, 32) = "=AE" & LineZ & "/AG3"
                        Ws.Cells(LineZ, 33) = "=AF" & LineZ & "*D" & LineZ
                End Select
            End While
        End If
        mSQLReader2.Close()
        FillExcelLine()
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
    Private Sub FillExcelLine()
        oRng = Ws.Range(Ws.Cells(3, 1), Ws.Cells(LineZ, 9 + (DateCount - 1) * 4))
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
    End Sub
    Private Function GetDifficult(ByVal ModelID As String)
        Dim mSQLS4 As New SqlClient.SqlCommand
        mSQLS4.Connection = mConnection
        mSQLS4.CommandType = CommandType.Text
        mSQLS4.CommandText = "select Degree from efd.dbo.TaktTime where StationGroupId = 2 and ModelID = '" & ModelID & "'"
        Dim DC As Decimal = mSQLS4.ExecuteScalar()
        If IsDBNull(DC) Then
            DC = 0
        End If
        Return DC
    End Function
    Private Function GetRatio(ByVal ModelID As String)
        Dim mSQLS4 As New SqlClient.SqlCommand
        mSQLS4.Connection = mConnection
        mSQLS4.CommandType = CommandType.Text
        mSQLS4.CommandText = "select Ratio from efd.dbo.TaktTime where StationGroupId = 2 and ModelID = '" & ModelID & "'"
        Dim DC As Decimal = mSQLS4.ExecuteScalar()
        If IsDBNull(DC) Then
            DC = 0
        End If
        Return DC
    End Function
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
End Class