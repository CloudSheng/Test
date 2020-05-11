Public Class Form152
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            xExcel = New Microsoft.Office.Interop.Excel.Application
            xWorkBook = xExcel.Workbooks.Add()
            Ws = xWorkBook.Sheets(1)
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
    End Sub
End Class