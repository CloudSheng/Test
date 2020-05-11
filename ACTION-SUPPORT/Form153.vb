Public Class Form153
    'Dim xPPT As Microsoft.Office.Interop.PowerPoint.Application
    'Dim xWorkBook As Microsoft.Office.Interop.PowerPoint.Presentations
    'Dim xSlide As Microsoft.Office.Interop.PowerPoint.Slides
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            '      xPPT = New Microsoft.Office.Interop.PowerPoint.Application
            '    xWorkBook = xPPT.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoFalse)
            '  Dim XX As Microsoft.Office.Interop.PowerPoint.CustomLayout = xWorkBook.SlideMaster.CustomLayouts().Item(1)
            ' xSlide = xWorkBook.Slides.AddSlide(1, XX)
            '            Dim FirstText As Microsoft.Office.Interop.PowerPoint.
            ' xWorkBook.SaveAs("c:\Temp\aa.pptx", PowerPoint.PpSaveAsFileType.ppSaveAsDefault)
            ' xWorkBook.Close()
            ' xPPT.Quit()
            GC.Collect()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
    End Sub
End Class