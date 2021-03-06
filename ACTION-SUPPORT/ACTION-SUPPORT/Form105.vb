﻿Imports Microsoft.Office.Interop.Excel.XlFileFormat
Public Class Form105
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Me.Label1.Text = 0
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Form105_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        mSQLS1.CommandText = "select sn.sn,lot.model,sn.updatedstation,datediff(Hour, updatedtime ,getdate()) from sn "
        mSQLS1.CommandText += "left join lot on sn.lot = lot.lot where sn.updatedstation in ('0493','0395','0193','0180','0605','0165','0673','0567','0455','0633','0629','0112','0435','0405','0478') "
        mSQLS1.CommandText += "and (topreworkstation is null or topreworkstation = '') "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows Then
            While mSQLReader.Read
                For i As Int16 = 0 To mSQLReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = mSQLReader.Item(i)
                Next
                LineZ += 1
                Label1.Text = LineZ - 1
            End While
        End If
        mSQLReader.Close()

    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.ColumnWidth = 23
        Ws.Name = "Detail"
        Ws.Cells(1, 1) = "SN"
        Ws.Cells(1, 2) = "型号"
        Ws.Cells(1, 3) = "目前所在工站"
        Ws.Cells(1, 4) = "停留时间（H)"
        LineZ = 2
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.NumberFormatLocal = "@"
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "接收站库龄表"
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