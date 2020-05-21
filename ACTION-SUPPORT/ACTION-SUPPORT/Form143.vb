Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.XlChartType
Public Class Form143
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim LineMax As Int16 = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form143_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfERPSUPPORT
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Dim xPath As String = "C:\temp\Project_RD_Sample.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
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

        TimeS1 = Convert.ToDateTime(DateTimePicker1.Value.Year & "/" & DateTimePicker1.Value.Month & "/01")
        TimeS2 = TimeS1.AddMonths(1).AddDays(-1)
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "RD人工工时报表"
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
        MsgBox("Finished")
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        Dim xPath As String = "C:\temp\Project_RD_Sample.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        LineZ = 7
        mSQLS1.CommandText = "select modelid,sum(ehour) as t1 from ProjectHR  where EDepartNo like '%PSE%' and edate between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd") & "' and '" & TimeS2.ToString("yyyy/MM/dd") & "' group by modelid"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 2) = mSQLReader.Item("modelid")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("t1")
                LineZ += 1
            End While
            LineMax = LineZ - 1
        End If
        If LineMax <> 0 Then
            oRng = Ws.Range("A7", Ws.Cells(LineMax, 1))
            oRng.Merge()
            Ws.Cells(7, 1) = "PSE All"
            Ws.Cells(5, 3) = "=SUM(C7:C" & LineMax & ")"
            oRng = Ws.Range("A7", Ws.Cells(LineMax, 3))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        mSQLReader.Close()
        ' 第1階-第五階
        For i As Int16 = 1 To 5 Step 1
            LineZ = 7
            mSQLS1.CommandText = "select modelid,sum(ehour) as t1 from ProjectHR  where EDepartNo like '%PSE%' and eap = " & i & " and edate between '"
            mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd") & "' and '" & TimeS2.ToString("yyyy/MM/dd") & "' group by modelid"
            mSQLReader = mSQLS1.ExecuteReader()
            If mSQLReader.HasRows() Then
                While mSQLReader.Read()
                    Ws.Cells(LineZ, 3 * i + 2) = mSQLReader.Item("modelid")
                    Ws.Cells(LineZ, 3 * i + 3) = mSQLReader.Item("t1")
                    LineZ += 1
                End While
            End If
            If LineZ <> 7 Then
                oRng = Ws.Range(Ws.Cells(7, 3 * i + 1), Ws.Cells(LineZ - 1, 3 * i + 1))
                oRng.Merge()
                Select Case i
                    Case 1
                        Ws.Cells(7, 4) = "第一阶段"
                        Ws.Cells(5, 6) = "=SUM(F7:F" & LineZ - 1 & ")"
                    Case 2
                        Ws.Cells(7, 7) = "第二阶段"
                        Ws.Cells(5, 9) = "=SUM(I7:I" & LineZ - 1 & ")"
                    Case 3
                        Ws.Cells(7, 10) = "第三阶段"
                        Ws.Cells(5, 12) = "=SUM(L7:L" & LineZ - 1 & ")"
                    Case 4
                        Ws.Cells(7, 13) = "第四阶段"
                        Ws.Cells(5, 15) = "=SUM(O7:O" & LineZ - 1 & ")"
                    Case 5
                        Ws.Cells(7, 16) = "第五阶段"
                        Ws.Cells(5, 18) = "=SUM(R7:R" & LineZ - 1 & ")"
                End Select
                
                oRng = Ws.Range(Ws.Cells(7, i * 3 + 1), Ws.Cells(LineZ - 1, i * 3 + 3))
                oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
                oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
                oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
                oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
                oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
            End If
            mSQLReader.Close()

        Next
        ' 第10階
        LineZ = 7
        mSQLS1.CommandText = "select modelid,sum(ehour) as t1 from ProjectHR  where EDepartNo like '%PSE%' and eap = 10 and edate between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd") & "' and '" & TimeS2.ToString("yyyy/MM/dd") & "' group by modelid"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 20) = mSQLReader.Item("modelid")
                Ws.Cells(LineZ, 21) = mSQLReader.Item("t1")
                LineZ += 1
            End While
        End If
        If LineMax <> 0 Then
            oRng = Ws.Range("S7", Ws.Cells(LineZ - 1, 19))
            oRng.Merge()
            Ws.Cells(7, 19) = "其他"
            Ws.Cells(5, 21) = "=SUM(U7:U" & LineZ - 1 & ")"
            oRng = Ws.Range("S7", Ws.Cells(LineZ - 1, 21))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        mSQLReader.Close()

        '劃圖
        ' 圖1
        Dim YB As Excel.Chart = Ws.Shapes.AddChart(xlPie, 0, 15.5 * LineMax, 500, 500).Chart
        oRng = Ws.Range("B7", Ws.Cells(LineMax, 3))
        YB.SetSourceData(oRng)
        'MsgBox(YB.HasLegend)
        YB.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementLegendNone)
        YB.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowPercent, False, Type.Missing, False, False, True, False, True, False, Type.Missing)

        '第二頁
        '
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        LineZ = 7
        mSQLS1.CommandText = "select modelid,sum(ehour) as t1 from ProjectHR  where EDepartNo like '%CAD%' and edate between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd") & "' and '" & TimeS2.ToString("yyyy/MM/dd") & "' group by modelid"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 2) = mSQLReader.Item("modelid")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("t1")
                LineZ += 1
            End While
            LineMax = LineZ - 1
        End If
        If LineMax <> 0 Then
            oRng = Ws.Range("A7", Ws.Cells(LineMax, 1))
            oRng.Merge()
            Ws.Cells(7, 1) = "All"
            Ws.Cells(5, 3) = "=SUM(C7:C" & LineMax & ")"
            oRng = Ws.Range("A7", Ws.Cells(LineMax, 3))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        mSQLReader.Close()
        ' 第1階-第五階
        For i As Int16 = 1 To 5 Step 1
            LineZ = 7
            mSQLS1.CommandText = "select modelid,sum(ehour) as t1 from ProjectHR  where EDepartNo like '%CAD%' and eap = " & i & " and edate between '"
            mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd") & "' and '" & TimeS2.ToString("yyyy/MM/dd") & "' group by modelid"
            mSQLReader = mSQLS1.ExecuteReader()
            If mSQLReader.HasRows() Then
                While mSQLReader.Read()
                    Ws.Cells(LineZ, 3 * i + 2) = mSQLReader.Item("modelid")
                    Ws.Cells(LineZ, 3 * i + 3) = mSQLReader.Item("t1")
                    LineZ += 1
                End While
            End If
            If LineZ <> 7 Then
                oRng = Ws.Range(Ws.Cells(7, 3 * i + 1), Ws.Cells(LineZ - 1, 3 * i + 1))
                oRng.Merge()
                Select Case i
                    Case 1
                        Ws.Cells(7, 4) = "第一阶段"
                        Ws.Cells(5, 6) = "=SUM(F7:F" & LineZ - 1 & ")"
                    Case 2
                        Ws.Cells(7, 7) = "第二阶段"
                        Ws.Cells(5, 9) = "=SUM(I7:I" & LineZ - 1 & ")"
                    Case 3
                        Ws.Cells(7, 10) = "第三阶段"
                        Ws.Cells(5, 12) = "=SUM(L7:L" & LineZ - 1 & ")"
                    Case 4
                        Ws.Cells(7, 13) = "第四阶段"
                        Ws.Cells(5, 15) = "=SUM(O7:O" & LineZ - 1 & ")"
                    Case 5
                        Ws.Cells(7, 16) = "第五阶段"
                        Ws.Cells(5, 18) = "=SUM(R7:R" & LineZ - 1 & ")"
                End Select

                oRng = Ws.Range(Ws.Cells(7, i * 3 + 1), Ws.Cells(LineZ - 1, i * 3 + 3))
                oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
                oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
                oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
                oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
                oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
            End If
            mSQLReader.Close()

        Next
        ' 第10階
        LineZ = 7
        mSQLS1.CommandText = "select modelid,sum(ehour) as t1 from ProjectHR  where EDepartNo like '%CAD%' and eap = 10 and edate between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd") & "' and '" & TimeS2.ToString("yyyy/MM/dd") & "' group by modelid"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 20) = mSQLReader.Item("modelid")
                Ws.Cells(LineZ, 21) = mSQLReader.Item("t1")
                LineZ += 1
            End While
        End If
        If LineMax <> 0 Then
            oRng = Ws.Range("S7", Ws.Cells(LineZ - 1, 19))
            oRng.Merge()
            Ws.Cells(7, 19) = "其他"
            Ws.Cells(5, 21) = "=SUM(U7:U" & LineZ - 1 & ")"
            oRng = Ws.Range("S7", Ws.Cells(LineZ - 1, 21))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        mSQLReader.Close()



        '第三頁
        '
        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        LineZ = 7
        mSQLS1.CommandText = "select modelid,sum(ehour) as t1 from ProjectHR  where EDepartNo like '%PET%' and edate between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd") & "' and '" & TimeS2.ToString("yyyy/MM/dd") & "' group by modelid"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 2) = mSQLReader.Item("modelid")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("t1")
                LineZ += 1
            End While
            LineMax = LineZ - 1
        End If
        If LineMax <> 0 Then
            oRng = Ws.Range("A7", Ws.Cells(LineMax, 1))
            oRng.Merge()
            Ws.Cells(7, 1) = "All"
            Ws.Cells(5, 3) = "=SUM(C7:C" & LineMax & ")"
            oRng = Ws.Range("A7", Ws.Cells(LineMax, 3))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        mSQLReader.Close()
        ' 第1階-第五階
        For i As Int16 = 1 To 5 Step 1
            LineZ = 7
            mSQLS1.CommandText = "select modelid,sum(ehour) as t1 from ProjectHR  where EDepartNo like '%PET%' and eap = " & i & " and edate between '"
            mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd") & "' and '" & TimeS2.ToString("yyyy/MM/dd") & "' group by modelid"
            mSQLReader = mSQLS1.ExecuteReader()
            If mSQLReader.HasRows() Then
                While mSQLReader.Read()
                    Ws.Cells(LineZ, 3 * i + 2) = mSQLReader.Item("modelid")
                    Ws.Cells(LineZ, 3 * i + 3) = mSQLReader.Item("t1")
                    LineZ += 1
                End While
            End If
            If LineZ <> 7 Then
                oRng = Ws.Range(Ws.Cells(7, 3 * i + 1), Ws.Cells(LineZ - 1, 3 * i + 1))
                oRng.Merge()
                Select Case i
                    Case 1
                        Ws.Cells(7, 4) = "第一阶段"
                        Ws.Cells(5, 6) = "=SUM(F7:F" & LineZ - 1 & ")"
                    Case 2
                        Ws.Cells(7, 7) = "第二阶段"
                        Ws.Cells(5, 9) = "=SUM(I7:I" & LineZ - 1 & ")"
                    Case 3
                        Ws.Cells(7, 10) = "第三阶段"
                        Ws.Cells(5, 12) = "=SUM(L7:L" & LineZ - 1 & ")"
                    Case 4
                        Ws.Cells(7, 13) = "第四阶段"
                        Ws.Cells(5, 15) = "=SUM(O7:O" & LineZ - 1 & ")"
                    Case 5
                        Ws.Cells(7, 16) = "第五阶段"
                        Ws.Cells(5, 18) = "=SUM(R7:R" & LineZ - 1 & ")"
                End Select

                oRng = Ws.Range(Ws.Cells(7, i * 3 + 1), Ws.Cells(LineZ - 1, i * 3 + 3))
                oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
                oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
                oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
                oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
                oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
            End If
            mSQLReader.Close()

        Next
        ' 第10階
        LineZ = 7
        mSQLS1.CommandText = "select modelid,sum(ehour) as t1 from ProjectHR  where EDepartNo like '%PET%' and eap = 10 and edate between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd") & "' and '" & TimeS2.ToString("yyyy/MM/dd") & "' group by modelid"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 20) = mSQLReader.Item("modelid")
                Ws.Cells(LineZ, 21) = mSQLReader.Item("t1")
                LineZ += 1
            End While
        End If
        If LineMax <> 0 Then
            oRng = Ws.Range("S7", Ws.Cells(LineZ - 1, 19))
            oRng.Merge()
            Ws.Cells(7, 19) = "其他"
            Ws.Cells(5, 21) = "=SUM(U7:U" & LineZ - 1 & ")"
            oRng = Ws.Range("S7", Ws.Cells(LineZ - 1, 21))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        mSQLReader.Close()

        '第四頁
        Ws = xWorkBook.Sheets(4)
        Ws.Activate()
        LineZ = 7
        mSQLS1.CommandText = "select modelid,sum(ehour) as t1 from ProjectHR  where edate between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd") & "' and '" & TimeS2.ToString("yyyy/MM/dd") & "' group by modelid"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 2) = mSQLReader.Item("modelid")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("t1")
                LineZ += 1
            End While
            LineMax = LineZ - 1
        End If
        If LineMax <> 0 Then
            oRng = Ws.Range("A7", Ws.Cells(LineMax, 1))
            oRng.Merge()
            Ws.Cells(7, 1) = "All"
            Ws.Cells(5, 3) = "=SUM(C7:C" & LineMax & ")"
            oRng = Ws.Range("A7", Ws.Cells(LineMax, 3))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        mSQLReader.Close()
        ' 第1階-第五階
        For i As Int16 = 1 To 5 Step 1
            LineZ = 7
            mSQLS1.CommandText = "select modelid,sum(ehour) as t1 from ProjectHR  where eap = " & i & " and edate between '"
            mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd") & "' and '" & TimeS2.ToString("yyyy/MM/dd") & "' group by modelid"
            mSQLReader = mSQLS1.ExecuteReader()
            If mSQLReader.HasRows() Then
                While mSQLReader.Read()
                    Ws.Cells(LineZ, 3 * i + 2) = mSQLReader.Item("modelid")
                    Ws.Cells(LineZ, 3 * i + 3) = mSQLReader.Item("t1")
                    LineZ += 1
                End While
            End If
            If LineZ <> 7 Then
                oRng = Ws.Range(Ws.Cells(7, 3 * i + 1), Ws.Cells(LineZ - 1, 3 * i + 1))
                oRng.Merge()
                Select Case i
                    Case 1
                        Ws.Cells(7, 4) = "第一阶段"
                        Ws.Cells(5, 6) = "=SUM(F7:F" & LineZ - 1 & ")"
                    Case 2
                        Ws.Cells(7, 7) = "第二阶段"
                        Ws.Cells(5, 9) = "=SUM(I7:I" & LineZ - 1 & ")"
                    Case 3
                        Ws.Cells(7, 10) = "第三阶段"
                        Ws.Cells(5, 12) = "=SUM(L7:L" & LineZ - 1 & ")"
                    Case 4
                        Ws.Cells(7, 13) = "第四阶段"
                        Ws.Cells(5, 15) = "=SUM(O7:O" & LineZ - 1 & ")"
                    Case 5
                        Ws.Cells(7, 16) = "第五阶段"
                        Ws.Cells(5, 18) = "=SUM(R7:R" & LineZ - 1 & ")"
                End Select

                oRng = Ws.Range(Ws.Cells(7, i * 3 + 1), Ws.Cells(LineZ - 1, i * 3 + 3))
                oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
                oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
                oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
                oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
                oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
            End If
            mSQLReader.Close()

        Next
        ' 第10階
        LineZ = 7
        mSQLS1.CommandText = "select modelid,sum(ehour) as t1 from ProjectHR  where eap = 10 and edate between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd") & "' and '" & TimeS2.ToString("yyyy/MM/dd") & "' group by modelid"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 20) = mSQLReader.Item("modelid")
                Ws.Cells(LineZ, 21) = mSQLReader.Item("t1")
                LineZ += 1
            End While
        End If
        If LineMax <> 0 Then
            oRng = Ws.Range("S7", Ws.Cells(LineZ - 1, 19))
            oRng.Merge()
            Ws.Cells(7, 19) = "其他"
            Ws.Cells(5, 21) = "=SUM(U7:U" & LineZ - 1 & ")"
            oRng = Ws.Range("S7", Ws.Cells(LineZ - 1, 21))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        mSQLReader.Close()
    End Sub
End Class