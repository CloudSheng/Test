Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel

Public Class Form355
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand3 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim DBC As String = String.Empty
    Dim LineZ As Integer = 0
    Dim LineY As Integer = 0
    Dim DNP As String = String.Empty
    Dim OCC01 As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Dim SaveFileDialog1 As New SaveFileDialog
    Dim l_tc_xma01 As String = String.Empty
    Dim l_tc_xma13 As String = String.Empty
    Dim l_tc_xma19 As String = String.Empty
    Dim l_tc_xma32 As String = String.Empty
    Dim l_tc_xma42 As String = String.Empty
    Dim l_tc_xma34 As String = String.Empty
    Dim ll_tc_xma19 As String = String.Empty
    Dim l_ofa23 As String = String.Empty
    Private Sub Form355_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'If Me.BackgroundWorker1.IsBusy() Then
        'MsgBox("处理中，请等待")
        'Return
        'End If        

        Dim xPath As String = "C:\temp\import_declaration_manual_demo.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If

        DBC = "actiontest"
        oConnection.ConnectionString = Module1.OpenOracleConnection(DBC)
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommand2.Connection = oConnection
                oCommand2.CommandType = CommandType.Text
                oCommand3.Connection = oConnection
                oCommand3.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

        If Not String.IsNullOrEmpty(TextBox1.Text) Then
            l_tc_xma01 = TextBox1.Text
        End If

        'xExcel = New Microsoft.Office.Interop.Excel.Application
        'xWorkBook = xExcel.Workbooks.Add()
        ExportToExcel()
        oConnection.Close()

        SaveExcel()
    End Sub

    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        Dim xPath As String = "C:\temp\import_declaration_manual_demo.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)

        '8.1新版进口货物报关单
        Ws = xWorkBook.Sheets(1)
        oCommand.CommandText = "select tc_xmaud03,tc_xma11,tc_xma15,tc_xma28,tc_xma20,tc_xmaud02,tc_xma12,tc_xma16, "
        oCommand.CommandText += "      tc_xma17,tc_xma05,tc_xma07,tc_xma41,tc_xma13,tc_xma18,tc_xma14,tc_xma06,tc_xma08, "
        oCommand.CommandText += "      tc_xma30,tc_xma22,tc_xma10,tc_xma25,tc_xma21,tc_xma31,tc_xma33,tc_xma24, "
        oCommand.CommandText += "      tc_xma19,tc_xma32,tc_xma42,tc_xma34 "
        oCommand.CommandText += " from tc_xma_file "
        oCommand.CommandText += "where tc_xma01 = '" & l_tc_xma01 & "' "
        LineY = 5
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineY + 2, 1) = oReader.Item("tc_xmaud03")
                Ws.Cells(LineY + 6, 1) = oReader.Item("tc_xma11")
                Ws.Cells(LineY + 8, 1) = oReader.Item("tc_xma15")
                Ws.Cells(LineY, 5) = oReader.Item("tc_xma28")
                Ws.Cells(LineY + 2, 5) = oReader.Item("tc_xma20")
                Ws.Cells(LineY + 4, 5) = oReader.Item("tc_xmaud02")
                Ws.Cells(LineY + 6, 5) = oReader.Item("tc_xma12")
                Ws.Cells(LineY + 8, 5) = oReader.Item("tc_xma16")
                Ws.Cells(LineY + 8, 6) = oReader.Item("tc_xma17")
                Ws.Cells(LineY, 7) = oReader.Item("tc_xma05")
                Ws.Cells(LineY + 2, 7) = oReader.Item("tc_xma07")
                Ws.Cells(LineY + 4, 7) = oReader.Item("tc_xma41")
                Ws.Cells(LineY + 6, 7) = oReader.Item("tc_xma13")
                Ws.Cells(LineY + 8, 7) = oReader.Item("tc_xma18")
                Ws.Cells(LineY + 8, 8) = oReader.Item("tc_xma14")
                Ws.Cells(LineY, 9) = oReader.Item("tc_xma06")
                Ws.Cells(LineY + 2, 9) = oReader.Item("tc_xma08")
                Ws.Cells(LineY + 6, 9) = oReader.Item("tc_xma30")
                Ws.Cells(LineY + 8, 9) = oReader.Item("tc_xma22")
                Ws.Cells(LineY, 10) = oReader.Item("tc_xma10")
                Ws.Cells(LineY + 2, 10) = oReader.Item("tc_xma25")
                Ws.Cells(LineY + 4, 10) = oReader.Item("tc_xma21")
                Ws.Cells(LineY + 6, 10) = oReader.Item("tc_xma31")
                Ws.Cells(LineY + 8, 10) = oReader.Item("tc_xma33")
                Ws.Cells(LineY + 8, 13) = oReader.Item("tc_xma24")
                l_tc_xma19 = oReader.Item("tc_xma19")
                l_tc_xma32 = oReader.Item("tc_xma32")
                l_tc_xma42 = oReader.Item("tc_xma42")
                l_tc_xma34 = oReader.Item("tc_xma34")
            End While
        End If
        oReader.Close()

        'oCommand.CommandText = "select ofa0451,ofa0457 "
        'oCommand.CommandText += " from ofa_file,tc_xma_file "
        'oCommand.CommandText += "where ofa01 = tc_xma29 and tc_xma01 = '" & l_tc_xma01 & "' "
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        Ws.Cells(LineY + 4, 1) = oReader.Item("ofa0451")
        '        Ws.Cells(LineY + 3, 3) = oReader.Item("ofa0457")
        '    End While
        'End If
        'oReader.Close()

        oCommand.CommandText = "select tc_xmbud01,tc_xmbud10,tc_xmb03,tc_xmb04,tc_xmb13,tc_xmb16,tc_xmb15,tc_xmb05,tc_xmb02  "
        oCommand.CommandText += " from tc_xmb_file "
        oCommand.CommandText += "where tc_xmb01 = '" & l_tc_xma01 & "' "
        oCommand.CommandText += "order by tc_xmb02 "
        LineZ = 28
        Dim l_tc_xmbud01 As String = String.Empty
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                'l_tc_xmbud01 = l_tc_xmbud01
                l_tc_xmbud01 += " " & oReader.Item("tc_xmbud01")

                'Ws.Cells(LineZ, 1) = oReader.Item("tc_xmbud10")
                Ws.Cells(LineZ, 1) = oReader.Item("tc_xmb02")
                Ws.Cells(LineZ + 2, 1) = oReader.Item("tc_xmbud10")
                Ws.Cells(LineZ, 2) = oReader.Item("tc_xmb03")
                Ws.Cells(LineZ, 3) = oReader.Item("tc_xmb04")
                Ws.Cells(LineZ, 5) = oReader.Item("tc_xmb13")
                Ws.Cells(LineZ, 6) = "千克"
                Ws.Cells(LineZ, 7) = oReader.Item("tc_xmb16")
                Ws.Cells(LineZ + 1, 7) = oReader.Item("tc_xmb15")
                Ws.Cells(LineZ + 2, 7) = l_tc_xma19
                Ws.Cells(LineZ, 8) = oReader.Item("tc_xmb05")
                Ws.Cells(LineZ, 9) = l_tc_xma32
                Ws.Cells(LineZ, 10) = l_tc_xma42
                Ws.Cells(LineZ, 13) = l_tc_xma34
                LineZ += 3
            End While
        End If
        oReader.Close()
        Ws.Cells(17, 1) = l_tc_xmbud01

        '装箱单
        Ws = xWorkBook.Sheets(2)
        oCommand.CommandText = "select tc_xmaud03,tc_xma38,tc_xma10,tc_xma11,tc_xma13,tc_xma12,tc_xma15,tc_xma16,tc_xma17,tc_xma18,tc_xma19 "
        oCommand.CommandText += " from tc_xma_file "
        oCommand.CommandText += "where tc_xma01 = '" & l_tc_xma01 & "' "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(3, 3) = oReader.Item("tc_xmaud03")
                Ws.Cells(3, 8) = oReader.Item("tc_xma38")
                Ws.Cells(4, 8) = oReader.Item("tc_xma10")
                Ws.Cells(6, 3) = oReader.Item("tc_xma11")
                Ws.Cells(7, 3) = oReader.Item("tc_xma13")
                Ws.Cells(7, 5) = oReader.Item("tc_xma12")
                Ws.Cells(8, 2) = oReader.Item("tc_xma16")
                Ws.Cells(8, 4) = oReader.Item("tc_xma15")
                Ws.Cells(8, 7) = oReader.Item("tc_xma17")
                Ws.Cells(8, 9) = oReader.Item("tc_xma18")
                Ws.Cells(12, 7) = oReader.Item("tc_xma19")
                Ws.Cells(12, 8) = oReader.Item("tc_xma19")
                l_tc_xma19 = oReader.Item("tc_xma19")
            End While
        End If
        oReader.Close()

        oCommand.CommandText = "select tc_xmb02,tc_xmbud10,tc_xmb03,tc_xmb04,tc_xmb13,tc_xmb16,tc_xmb15,tc_xmb05 "
        oCommand.CommandText += " from tc_xmb_file "
        oCommand.CommandText += "where tc_xmb01 = '" & l_tc_xma01 & "' "
        oCommand.CommandText += "order by tc_xmb02 "
        LineZ = 13
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                'Ws.Cells(LineZ, 1) = oReader.Item("tc_xmbud10")
                Ws.Cells(LineZ, 1) = oReader.Item("tc_xmb02")
                Ws.Cells(LineZ, 2) = oReader.Item("tc_xmb03")
                Ws.Cells(LineZ, 3) = oReader.Item("tc_xmb04")
                Ws.Cells(LineZ, 5) = oReader.Item("tc_xmb13")
                Ws.Cells(LineZ, 6) = "千克"
                Ws.Cells(LineZ, 7) = oReader.Item("tc_xmb16")
                Ws.Cells(LineZ, 8) = oReader.Item("tc_xmb15")
                Ws.Cells(LineZ, 9) = oReader.Item("tc_xmb05")
                LineZ += 1
            End While
        End If
        oReader.Close()
        Ws.Cells(24, 5) = "=SUM(E13:E" & LineZ - 1 & ")"
        Ws.Cells(24, 8) = "=SUM(H13:H" & LineZ - 1 & ")"
        

        '发票
        Ws = xWorkBook.Sheets(3)
        oCommand.CommandText = "select tc_xmaud03,tc_xma29,tc_xma02,tc_xma19 "
        oCommand.CommandText += " from tc_xma_file "
        oCommand.CommandText += "where tc_xma01 = '" & l_tc_xma01 & "' "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(3, 3) = oReader.Item("tc_xmaud03")
                Ws.Cells(3, 8) = oReader.Item("tc_xma29")
                Ws.Cells(4, 8) = oReader.Item("tc_xma02")
                Ws.Cells(8, 7) = oReader.Item("tc_xma19")
                Ws.Cells(8, 8) = oReader.Item("tc_xma19")
                l_tc_xma19 = oReader.Item("tc_xma19")
            End While
        End If
        oReader.Close()

        oCommand.CommandText = "select tc_xmb02,tc_xmbud10,tc_xmb04,tc_xmb13,tc_xmb16,tc_xmb15,tc_xmb05 "
        oCommand.CommandText += " from tc_xmb_file "
        oCommand.CommandText += "where tc_xmb01 = '" & l_tc_xma01 & "' "
        oCommand.CommandText += "order by tc_xmb02 "
        LineZ = 9
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                'Ws.Cells(LineZ, 1) = oReader.Item("tc_xmbud10")
                Ws.Cells(LineZ, 1) = oReader.Item("tc_xmb02")
                Ws.Cells(LineZ, 3) = oReader.Item("tc_xmb04")
                Ws.Cells(LineZ, 5) = oReader.Item("tc_xmb13")
                Ws.Cells(LineZ, 6) = "千克"
                Ws.Cells(LineZ, 7) = oReader.Item("tc_xmb16")
                Ws.Cells(LineZ, 8) = oReader.Item("tc_xmb15")
                Ws.Cells(LineZ, 9) = oReader.Item("tc_xmb05")
                LineZ += 1
            End While
        End If
        oReader.Close()
        Ws.Cells(16, 8) = "=SUM(H9:H" & LineZ - 1 & ")"

    End Sub

    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "import_declaration_manual"
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
        If oConnection.State = ConnectionState.Open Then
            Try
                oConnection.Close()
                Module1.KillExcelProcess(OldExcel)
                MsgBox("Finished")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
End Class