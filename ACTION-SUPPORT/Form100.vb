Imports Microsoft.Office.Interop.Excel.XlFileFormat
Public Class Form100
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    'Dim oCell As Microsoft.Office.Interop.Excel.
    Dim TotalFile As Int16 = 0
    Dim Linez As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Dim xPath As String = My.Computer.FileSystem.CurrentDirectory
        'MsgBox(xPath)
        'xExcel = New Microsoft.Office.Interop.Excel.Application
        'xWorkBook = xExcel.Workbooks.Add()
        'Ws = xWorkBook.Sheets(1)
        'Ws.Activate()
        'Dim xPath1 As String = "C:\Users\Cloud\Pictures"
        'Dim AllFile As String() = My.Computer.FileSystem.GetFiles(xPath1).ToArray()
        'For i As Int16 = 0 To AllFile.Count - 1 Step 1
        '    If (AllFile(i).EndsWith("jpg") Or AllFile(i).EndsWith("JPG")) Then
        '        TotalFile += 1
        '        Dim LeftSide As Int16 = Decimal.Remainder(TotalFile, 2)
        '        Dim TopSide As Int16 = Decimal.Truncate(TotalFile / 2)
        '        If LeftSide = 1 Then
        '            Ws.Shapes.AddPicture(AllFile(i), Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, TopSide * 240, 340, 240)
        '        Else
        '            Ws.Shapes.AddPicture(AllFile(i), Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 380, TopSide * 240, 340, 240)
        '        End If
        '    End If
        'Next
        'Dim FileName1 As String = "C:\temp\aaaa.xlsx"
        ''Dim FileName1 As String = "c:\temp\" & D1.Day & "-Receive"
        'Try
        '    Ws.SaveAs(FileName1, xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing)
        '    xWorkBook.Saved = True
        '    xWorkBook.Close()
        '    xExcel.Quit()
        'Catch ex As Exception
        '    MsgBox(ex.Message())
        'End Try

        'Try
        '    Module1.KillExcelProcess(OldExcel)
        'Catch ex As Exception

        'End Try
        MsgBox(Strings.Chr(78))
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            xExcel = New Microsoft.Office.Interop.Excel.Application
            xWorkBook = xExcel.Workbooks.Open(ExcelPath)
            Ws = xWorkBook.Sheets(1)
            LineZ = 6
        End If
        'oRng = Ws.Range("C14", "J14")
        Dim BB As Integer = Ws.UsedRange.Rows.Count
        For i As Integer = 15 To BB Step 1
            Dim oRng2 As Microsoft.Office.Interop.Excel.Range
            oRng2 = Ws.Range(Ws.Cells(i, 4), Ws.Cells(i, 4))
            If Not String.IsNullOrEmpty(oRng2.Value2) Then
                Dim ACCNO As String = Strings.Left(oRng2.Value2.ToString(), 4)
                MsgBox(ACCNO)
                Dim oRng3 As Microsoft.Office.Interop.Excel.Range = Ws.Range(Ws.Cells(i, 10), Ws.Cells(i, 10))
                Dim Money1 As Decimal = oRng3.Value
                MsgBox(Money1)
            End If
        Next

        'For Each c In oRng
        'If Not String.IsNullOrEmpty(c.value) Then
        'MsgBox(c.row)
        'MsgBox(c.column)
        'End If

        'Next c
    End Sub
End Class