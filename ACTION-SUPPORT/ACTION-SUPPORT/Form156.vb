Public Class Form156
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim TotalFile As Int16 = 0
    Dim Linez As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            xExcel = New Microsoft.Office.Interop.Excel.Application
            xWorkBook = xExcel.Workbooks.Open(ExcelPath)
            Ws = xWorkBook.Sheets(1)
            Linez = 6
        Else
            Return
        End If
        Dim PP As Boolean = False
        mConnection.ConnectionString = Module1.OpenConnectionOfERPSUPPORT()
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If

        oRng = Ws.Range("K4", "K4")
        Dim Date1 As String = oRng.Value
        Dim Month1 As String = Strings.Mid(Date1, 4, 2)
        Dim Month2 As Int16 = Convert.ToInt16(Month1)
        Dim Year1 As String = Strings.Right(Date1, 4)
        Dim Year2 As Integer = Convert.ToInt64(Year1)

        mSQLS1.CommandText = "DELETE ACABS WHERE year1 = " & Year2 & " AND month1 = " & Month2
        Try
            mSQLS1.ExecuteNonQuery()
        Catch ex As Exception

        End Try


        Dim BB As Integer = Ws.UsedRange.Rows.Count
        For i As Integer = 13 To BB Step 1
            ' B段
            oRng = Ws.Range("B" & i, "B" & i)
            Dim SectorB As String = oRng.Value
            If Not IsNothing(SectorB) And SectorB = "Profit Period" And PP = False Then
                PP = True
                Dim Acc1 As String = Strings.Left(SectorB, 4)
                If Not String.IsNullOrEmpty(Acc1) Then
                    oRng = Ws.Range("M" & i, "M" & i)
                    Dim Num1 As Decimal = 0
                    If IsNumeric(oRng.Value2) Then
                        Num1 = oRng.Value2
                        mSQLS1.CommandText = "INSERT INTO ACABS VALUES (" & Year2 & "," & Month2 & ",'" & Acc1 & "'," & Num1 & ")"
                        Try
                            mSQLS1.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.Message())
                        End Try
                    End If
                End If
            End If

            'C 段

            oRng = Ws.Range("C" & i, "C" & i)
            Dim SectorC As String = oRng.Value
            If Not IsNothing(SectorC) Then
                Dim Acc1 As String = Strings.Left(SectorC, 4)
                If Not String.IsNullOrEmpty(Acc1) Then
                    oRng = Ws.Range("M" & i, "M" & i)
                    Dim Num1 As Decimal = 0
                    If IsNumeric(oRng.Value2) Then
                        Num1 = oRng.Value2
                        mSQLS1.CommandText = "INSERT INTO ACABS VALUES (" & Year2 & "," & Month2 & ",'" & Acc1 & "'," & Num1 & ")"
                        Try
                            mSQLS1.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.Message())
                        End Try
                    End If
                End If
            End If

            'D 段
            oRng = Ws.Range("D" & i, "D" & i)
            Dim SectorD As String = oRng.Value
            If Not IsNothing(SectorD) Then
                Dim Acc1 As String = Strings.Left(SectorD, 4)
                If Not String.IsNullOrEmpty(Acc1) Then
                    oRng = Ws.Range("M" & i, "M" & i)
                    Dim Num1 As Decimal = 0
                    If IsNumeric(oRng.Value2) Then
                        Num1 = oRng.Value2
                        mSQLS1.CommandText = "INSERT INTO ACABS VALUES (" & Year2 & "," & Month2 & ",'" & Acc1 & "'," & Num1 & ")"
                        Try
                            mSQLS1.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.Message())
                        End Try
                    End If
                End If
            End If

            'E 段
            oRng = Ws.Range("E" & i, "E" & i)
            Dim SectorE As String = oRng.Value
            If Not IsNothing(SectorE) Then
                Dim Acc1 As String = Strings.Left(SectorE, 4)
                If Not String.IsNullOrEmpty(Acc1) Then
                    oRng = Ws.Range("M" & i, "M" & i)
                    Dim Num1 As Decimal = 0
                    If IsNumeric(oRng.Value2) Then
                        Num1 = oRng.Value2
                        mSQLS1.CommandText = "INSERT INTO ACABS VALUES (" & Year2 & "," & Month2 & ",'" & Acc1 & "'," & Num1 & ")"
                        Try
                            mSQLS1.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.Message())
                        End Try
                    End If
                End If
            End If

            'F 段
            oRng = Ws.Range("F" & i, "F" & i)
            Dim SectorF As String = oRng.Value
            If Not IsNothing(SectorF) Then
                Dim Acc1 As String = Strings.Left(SectorF, 4)
                If Not String.IsNullOrEmpty(Acc1) Then
                    oRng = Ws.Range("M" & i, "M" & i)
                    Dim Num1 As Decimal = 0
                    If IsNumeric(oRng.Value2) Then
                        Num1 = oRng.Value2
                        mSQLS1.CommandText = "INSERT INTO ACABS VALUES (" & Year2 & "," & Month2 & ",'" & Acc1 & "'," & Num1 & ")"
                        Try
                            mSQLS1.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.Message())
                        End Try
                    End If
                End If
            End If

            ' G段
            oRng = Ws.Range("G" & i, "G" & i)
            Dim SectorG As String = oRng.Value
            If Not IsNothing(SectorG) Then
                Dim Acc1 As String = Strings.Left(SectorG, 4)
                If Not String.IsNullOrEmpty(Acc1) Then
                    oRng = Ws.Range("M" & i, "M" & i)
                    Dim Num1 As Decimal = 0
                    If IsNumeric(oRng.Value2) Then
                        Num1 = oRng.Value2
                        mSQLS1.CommandText = "INSERT INTO ACABS VALUES (" & Year2 & "," & Month2 & ",'" & Acc1 & "'," & Num1 & ")"
                        Try
                            mSQLS1.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.Message())
                        End Try
                    End If
                End If
            End If


            ' H段
            oRng = Ws.Range("H" & i, "H" & i)
            Dim SectorH As String = oRng.Value
            If Not IsNothing(SectorH) Then
                Dim Acc1 As String = Strings.Left(SectorH, 4)
                If Not String.IsNullOrEmpty(Acc1) Then
                    oRng = Ws.Range("M" & i, "M" & i)
                    Dim Num1 As Decimal = 0
                    If IsNumeric(oRng.Value2) Then
                        Num1 = oRng.Value2
                        mSQLS1.CommandText = "INSERT INTO ACABS VALUES (" & Year2 & "," & Month2 & ",'" & Acc1 & "'," & Num1 & ")"
                        Try
                            mSQLS1.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.Message())
                        End Try
                    End If
                End If
            End If
        Next
        xWorkBook.Close()
        xExcel.Quit()
        Module1.KillExcelProcess(OldExcel)
        MsgBox("DONE")
    End Sub
End Class