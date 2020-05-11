Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel

Public Class Form354
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
    Dim l_tc_xmb02 As String = String.Empty
    Dim ll_tc_xma19 As String = String.Empty
    Dim l_ofa23 As String = String.Empty
    Private Sub Form354_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'If Me.BackgroundWorker1.IsBusy() Then
        'MsgBox("处理中，请等待")
        'Return
        'End If        

        Dim xPath As String = "C:\temp\export_declaration_demo.xlsx"
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
        Dim xPath As String = "C:\temp\export_declaration_demo.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)

        '8.1新版出口货物报关单
        Ws = xWorkBook.Sheets(1)
        oCommand.CommandText = "select tc_xma03,tc_xma04,tc_xma05,tc_xma06,tc_xma07,tc_xma08,tc_xma09,tc_xma10,tc_xma11, "
        oCommand.CommandText += "      tc_xma12,tc_xma13,tc_xma14,tc_xma15,tc_xma16,tc_xma17,tc_xma18,tc_xma19,tc_xma20, "
        oCommand.CommandText += "      tc_xma21,tc_xma22,tc_xma23,tc_xma24,tc_xma25,tc_xma26,tc_xma27,tc_xma28,tc_xmaud02 "
        oCommand.CommandText += " from tc_xma_file "
        oCommand.CommandText += "where tc_xma01 = '" & l_tc_xma01 & "' "
        LineY = 3
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                'Ws.Cells(LineY, 3) = oReader.Item("tc_xma03")
                Ws.Cells(LineY, 9) = oReader.Item("tc_xma04")
                Ws.Cells(LineY + 2, 7) = oReader.Item("tc_xma05")
                'Ws.Cells(LineY + 2, 9) = oReader.Item("tc_xma06")
                Ws.Cells(LineY + 4, 7) = oReader.Item("tc_xma07")
                Ws.Cells(LineY + 4, 9) = oReader.Item("tc_xma08")
                Ws.Cells(LineY + 2, 10) = oReader.Item("tc_xma10")
                Ws.Cells(LineY + 6, 5) = oReader.Item("tc_xmaud02")
                Ws.Cells(LineY + 8, 1) = oReader.Item("tc_xma11")
                'Ws.Cells(LineY + 8, 5) = oReader.Item("tc_xma12")
                Ws.Cells(LineY + 8, 7) = oReader.Item("tc_xma13")
                l_tc_xma13 = oReader.Item("tc_xma13")
                Ws.Cells(LineY + 10, 8) = oReader.Item("tc_xma14")
                Ws.Cells(LineY + 10, 1) = oReader.Item("tc_xma15")
                Ws.Cells(LineY + 10, 5) = oReader.Item("tc_xma16")
                Ws.Cells(LineY + 10, 6) = oReader.Item("tc_xma17")
                Ws.Cells(LineY + 10, 7) = oReader.Item("tc_xma18")
                l_tc_xma19 = oReader.Item("tc_xma19")
            End While
        End If
        oReader.Close()

        oCommand.CommandText = "select ofa0451,ofa0457 "
        oCommand.CommandText += " from ofa_file,tc_xma_file "
        oCommand.CommandText += "where ofa01 = tc_xma29 and tc_xma01 = '" & l_tc_xma01 & "' "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineY + 4, 1) = oReader.Item("ofa0451")
                Ws.Cells(LineY + 3, 3) = oReader.Item("ofa0457")
            End While
        End If
        oReader.Close()

        oCommand.CommandText = "select tc_xmb02,tc_xmbud10,tc_xmb03,tc_xmb04,tc_xmb05,tc_xmb06,tc_xmb07,tc_xmb08,tc_xmb09, "
        oCommand.CommandText += "      tc_xmb10,tc_xmb11,tc_xmb12,tc_xmb13,tc_xmb14,tc_xmb15,tc_xmb16,tc_xmb17,tc_xmb18,tc_xmbud02,tc_xmb19 "
        oCommand.CommandText += " from tc_xmb_file "
        oCommand.CommandText += "where tc_xmb01 = '" & l_tc_xma01 & "' "
        oCommand.CommandText += "order by tc_xmb02 "
        LineZ = 21
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("tc_xmb02")
                l_tc_xmb02 = oReader.Item("tc_xmb02")                '190104 add by Brady
                Ws.Cells(LineZ + 2, 1) = oReader.Item("tc_xmbud10")
                Ws.Cells(LineZ, 2) = oReader.Item("tc_xmb03")
                Ws.Cells(LineZ, 3) = oReader.Item("tc_xmb04")
                Ws.Cells(LineZ + 2, 4) = oReader.Item("tc_xmb11")
                Ws.Cells(LineZ + 2, 5) = oReader.Item("tc_xmb12")
                Ws.Cells(LineZ, 5) = oReader.Item("tc_xmb13")
                Ws.Cells(LineZ, 6) = "千克"
                Ws.Cells(LineZ, 7) = oReader.Item("tc_xmb16")
                Ws.Cells(LineZ + 1, 7) = oReader.Item("tc_xmb15")
                Ws.Cells(LineZ + 2, 7) = l_tc_xma19
                Ws.Cells(LineZ, 8) = "中国"
                Ws.Cells(LineZ + 1, 8) = "（142）"
                Ws.Cells(LineZ, 9) = l_tc_xma13
                Ws.Cells(LineZ, 10) = "东莞"
                Ws.Cells(LineZ, 13) = "全免"
                LineZ += 3
                If LineZ = 39 Then Exit While
            End While
        End If
        oReader.Close()

        'oCommand.CommandText = "select tc_xma03,tc_xma04,tc_xma05,tc_xma06,tc_xma07,tc_xma08,tc_xma09,tc_xma10,tc_xma11, "
        'oCommand.CommandText += "      tc_xma12,tc_xma13,tc_xma14,tc_xma15,tc_xma16,tc_xma17,tc_xma18,tc_xma19,tc_xma20, "
        'oCommand.CommandText += "      tc_xma21,tc_xma22,tc_xma23,tc_xma24,tc_xma25,tc_xma26,tc_xma27,tc_xma28 "
        'oCommand.CommandText += " from tc_xma_file "
        'oCommand.CommandText += "where tc_xma01 = '" & l_tc_xma01 & "' "
        'LineY = 48
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        Ws.Cells(LineY, 3) = oReader.Item("tc_xma03")
        '        Ws.Cells(LineY, 9) = oReader.Item("tc_xma04")
        '        Ws.Cells(LineY + 2, 7) = oReader.Item("tc_xma05")
        '        Ws.Cells(LineY + 2, 9) = oReader.Item("tc_xma06")
        '        Ws.Cells(LineY + 4, 7) = oReader.Item("tc_xma07")
        '        Ws.Cells(LineY + 4, 9) = oReader.Item("tc_xma08")
        '        Ws.Cells(LineY + 2, 10) = oReader.Item("tc_xma10")
        '        Ws.Cells(LineY + 8, 1) = oReader.Item("tc_xma11")
        '        Ws.Cells(LineY + 8, 5) = oReader.Item("tc_xma12")
        '        Ws.Cells(LineY + 8, 7) = oReader.Item("tc_xma13")
        '        Ws.Cells(LineY + 10, 8) = oReader.Item("tc_xma14")
        '        Ws.Cells(LineY + 10, 1) = oReader.Item("tc_xma15")
        '        Ws.Cells(LineY + 10, 5) = oReader.Item("tc_xma16")
        '        Ws.Cells(LineY + 10, 6) = oReader.Item("tc_xma17")
        '    End While
        'End If
        'oReader.Close()

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

        oCommand.CommandText = "select tc_xmb02,tc_xmbud10,tc_xmb03,tc_xmb04,tc_xmb05,tc_xmb06,tc_xmb07,tc_xmb08,tc_xmb09, "
        oCommand.CommandText += "      tc_xmb10,tc_xmb11,tc_xmb12,tc_xmb13,tc_xmb14,tc_xmb15,tc_xmb16,tc_xmb17,tc_xmb18,tc_xmbud02,tc_xmb19 "
        oCommand.CommandText += " from tc_xmb_file "
        oCommand.CommandText += "where tc_xmb01 = '" & l_tc_xma01 & "' "
        '190104 add by Brady
        'oCommand.CommandText += "  and tc_xmb02 > 6 order by tc_xmb02 "
        oCommand.CommandText += "  and tc_xmb02 > " & l_tc_xmb02 & "order by tc_xmb02 "
        '190104 add by Brady END
        'LineZ = 66
        LineZ = 50
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("tc_xmb02")
                l_tc_xmb02 = oReader.Item("tc_xmb02")                '190104 add by Brady
                Ws.Cells(LineZ + 2, 1) = oReader.Item("tc_xmbud10")
                Ws.Cells(LineZ, 2) = oReader.Item("tc_xmb03")
                Ws.Cells(LineZ, 3) = oReader.Item("tc_xmb04")
                Ws.Cells(LineZ + 2, 4) = oReader.Item("tc_xmb11")
                Ws.Cells(LineZ + 2, 5) = oReader.Item("tc_xmb12")
                Ws.Cells(LineZ, 5) = oReader.Item("tc_xmb13")
                Ws.Cells(LineZ, 6) = "千克"
                Ws.Cells(LineZ, 7) = oReader.Item("tc_xmb16")
                Ws.Cells(LineZ + 1, 7) = oReader.Item("tc_xmb15")
                Ws.Cells(LineZ + 2, 7) = l_tc_xma19
                Ws.Cells(LineZ, 8) = "中国"
                Ws.Cells(LineZ + 1, 8) = "（142）"
                Ws.Cells(LineZ, 9) = l_tc_xma13
                Ws.Cells(LineZ, 10) = "东莞"
                Ws.Cells(LineZ, 13) = "全免"
                LineZ += 3
                If LineZ = 92 Then Exit While
            End While
        End If
        oReader.Close()

        'oCommand.CommandText = "select tc_xma03,tc_xma04,tc_xma05,tc_xma06,tc_xma07,tc_xma08,tc_xma09,tc_xma10,tc_xma11, "
        'oCommand.CommandText += "      tc_xma12,tc_xma13,tc_xma14,tc_xma15,tc_xma16,tc_xma17,tc_xma18,tc_xma19,tc_xma20, "
        'oCommand.CommandText += "      tc_xma21,tc_xma22,tc_xma23,tc_xma24,tc_xma25,tc_xma26,tc_xma27,tc_xma28 "
        'oCommand.CommandText += " from tc_xma_file "
        'oCommand.CommandText += "where tc_xma01 = '" & l_tc_xma01 & "' "
        'LineY = 93
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        Ws.Cells(LineY, 3) = oReader.Item("tc_xma03")
        '        Ws.Cells(LineY, 9) = oReader.Item("tc_xma04")
        '        Ws.Cells(LineY + 2, 7) = oReader.Item("tc_xma05")
        '        Ws.Cells(LineY + 2, 9) = oReader.Item("tc_xma06")
        '        Ws.Cells(LineY + 4, 7) = oReader.Item("tc_xma07")
        '        Ws.Cells(LineY + 4, 9) = oReader.Item("tc_xma08")
        '        Ws.Cells(LineY + 2, 10) = oReader.Item("tc_xma10")
        '        Ws.Cells(LineY + 8, 1) = oReader.Item("tc_xma11")
        '        Ws.Cells(LineY + 8, 5) = oReader.Item("tc_xma12")
        '        Ws.Cells(LineY + 8, 7) = oReader.Item("tc_xma13")
        '        Ws.Cells(LineY + 10, 8) = oReader.Item("tc_xma14")
        '        Ws.Cells(LineY + 10, 1) = oReader.Item("tc_xma15")
        '        Ws.Cells(LineY + 10, 5) = oReader.Item("tc_xma16")
        '        Ws.Cells(LineY + 10, 6) = oReader.Item("tc_xma17")
        '    End While
        'End If
        'oReader.Close()

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

        oCommand.CommandText = "select tc_xmb02,tc_xmbud10,tc_xmb03,tc_xmb04,tc_xmb05,tc_xmb06,tc_xmb07,tc_xmb08,tc_xmb09, "
        oCommand.CommandText += "      tc_xmb10,tc_xmb11,tc_xmb12,tc_xmb13,tc_xmb14,tc_xmb15,tc_xmb16,tc_xmb17,tc_xmb18,tc_xmbud02,tc_xmb19 "
        oCommand.CommandText += " from tc_xmb_file "
        oCommand.CommandText += "where tc_xmb01 = '" & l_tc_xma01 & "' "
        '190104 add by Brady
        'oCommand.CommandText += "  and tc_xmb02 > 20 order by tc_xmb02 "
        oCommand.CommandText += "  and tc_xmb02 > " & l_tc_xmb02 & "order by tc_xmb02 "
        '190104 add by Brady END
        LineZ = 98
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("tc_xmb02")
                Ws.Cells(LineZ + 2, 1) = oReader.Item("tc_xmbud10")
                Ws.Cells(LineZ, 2) = oReader.Item("tc_xmb03")
                Ws.Cells(LineZ, 3) = oReader.Item("tc_xmb04")
                Ws.Cells(LineZ + 2, 4) = oReader.Item("tc_xmb11")
                Ws.Cells(LineZ + 2, 5) = oReader.Item("tc_xmb12")
                Ws.Cells(LineZ, 5) = oReader.Item("tc_xmb13")
                Ws.Cells(LineZ, 6) = "千克"
                Ws.Cells(LineZ, 7) = oReader.Item("tc_xmb16")
                Ws.Cells(LineZ + 1, 7) = oReader.Item("tc_xmb15")
                Ws.Cells(LineZ + 2, 7) = l_tc_xma19
                Ws.Cells(LineZ, 8) = "中国"
                Ws.Cells(LineZ + 1, 8) = "（142）"
                Ws.Cells(LineZ, 9) = l_tc_xma13
                Ws.Cells(LineZ, 10) = "东莞"
                Ws.Cells(LineZ, 13) = "全免"
                LineZ += 3
                'If LineZ = 129 Then Exit While
            End While
        End If
        oReader.Close()

        'oCommand.CommandText = "select tc_xma03,tc_xma04,tc_xma05,tc_xma06,tc_xma07,tc_xma08,tc_xma09,tc_xma10,tc_xma11, "
        'oCommand.CommandText += "      tc_xma12,tc_xma13,tc_xma14,tc_xma15,tc_xma16,tc_xma17,tc_xma18,tc_xma19,tc_xma20, "
        'oCommand.CommandText += "      tc_xma21,tc_xma22,tc_xma23,tc_xma24,tc_xma25,tc_xma26,tc_xma27,tc_xma28 "
        'oCommand.CommandText += " from tc_xma_file "
        'oCommand.CommandText += "where tc_xma01 = '" & l_tc_xma01 & "' "
        'LineY = 138
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        Ws.Cells(LineY, 3) = oReader.Item("tc_xma03")
        '        Ws.Cells(LineY, 9) = oReader.Item("tc_xma04")
        '        Ws.Cells(LineY + 2, 7) = oReader.Item("tc_xma05")
        '        Ws.Cells(LineY + 2, 9) = oReader.Item("tc_xma06")
        '        Ws.Cells(LineY + 4, 7) = oReader.Item("tc_xma07")
        '        Ws.Cells(LineY + 4, 9) = oReader.Item("tc_xma08")
        '        Ws.Cells(LineY + 2, 10) = oReader.Item("tc_xma10")
        '        Ws.Cells(LineY + 8, 1) = oReader.Item("tc_xma11")
        '        Ws.Cells(LineY + 8, 5) = oReader.Item("tc_xma12")
        '        Ws.Cells(LineY + 8, 7) = oReader.Item("tc_xma13")
        '        Ws.Cells(LineY + 10, 8) = oReader.Item("tc_xma14")
        '        Ws.Cells(LineY + 10, 1) = oReader.Item("tc_xma15")
        '        Ws.Cells(LineY + 10, 5) = oReader.Item("tc_xma16")
        '        Ws.Cells(LineY + 10, 6) = oReader.Item("tc_xma17")
        '    End While
        'End If
        'oReader.Close()

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

        'oCommand.CommandText = "select tc_xmb02,tc_xmbud10,tc_xmb03,tc_xmb04,tc_xmb05,tc_xmb06,tc_xmb07,tc_xmb08,tc_xmb09, "
        'oCommand.CommandText += "      tc_xmb10,tc_xmb11,tc_xmb12,tc_xmb13,tc_xmb14,tc_xmb15,tc_xmb16,tc_xmb17,tc_xmb18,tc_xmbud02,tc_xmb19 "
        'oCommand.CommandText += " from tc_xmb_file "
        'oCommand.CommandText += "where tc_xmb01 = '" & l_tc_xma01 & "' "
        'oCommand.CommandText += "  and tc_xmb02 > 18 order by tc_xmb02 "
        'LineZ = 156
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        Ws.Cells(LineZ, 1) = oReader.Item("tc_xmb02")
        '        Ws.Cells(LineZ + 2, 1) = oReader.Item("tc_xmbud10")
        '        Ws.Cells(LineZ, 2) = oReader.Item("tc_xmb03")
        '        Ws.Cells(LineZ, 3) = oReader.Item("tc_xmb04")
        '        Ws.Cells(LineZ + 2, 4) = oReader.Item("tc_xmb11")
        '        Ws.Cells(LineZ + 2, 5) = oReader.Item("tc_xmb12")
        '        Ws.Cells(LineZ, 5) = oReader.Item("tc_xmb13")
        '        Ws.Cells(LineZ, 6) = "千克"
        '        Ws.Cells(LineZ, 7) = oReader.Item("tc_xmb16")
        '        Ws.Cells(LineZ + 1, 7) = oReader.Item("tc_xmb15")
        '        Ws.Cells(LineZ + 2, 7) = "美元"
        '        Ws.Cells(LineZ, 8) = "中国"
        '        Ws.Cells(LineZ + 1, 8) = "（142）"
        '        Ws.Cells(LineZ, 9) = l_tc_xma13
        '        Ws.Cells(LineZ, 10) = "东莞"
        '        Ws.Cells(LineZ, 13) = "全免"
        '        LineZ += 3
        '        If LineZ = 174 Then Exit While
        '    End While
        'End If
        'oReader.Close()

        'oCommand.CommandText = "select tc_xma03,tc_xma04,tc_xma05,tc_xma06,tc_xma07,tc_xma08,tc_xma09,tc_xma10,tc_xma11, "
        'oCommand.CommandText += "      tc_xma12,tc_xma13,tc_xma14,tc_xma15,tc_xma16,tc_xma17,tc_xma18,tc_xma19,tc_xma20, "
        'oCommand.CommandText += "      tc_xma21,tc_xma22,tc_xma23,tc_xma24,tc_xma25,tc_xma26,tc_xma27,tc_xma28 "
        'oCommand.CommandText += " from tc_xma_file "
        'oCommand.CommandText += "where tc_xma01 = '" & l_tc_xma01 & "' "
        'LineY = 183
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        Ws.Cells(LineY, 3) = oReader.Item("tc_xma03")
        '        Ws.Cells(LineY, 9) = oReader.Item("tc_xma04")
        '        Ws.Cells(LineY + 2, 7) = oReader.Item("tc_xma05")
        '        Ws.Cells(LineY + 2, 9) = oReader.Item("tc_xma06")
        '        Ws.Cells(LineY + 4, 7) = oReader.Item("tc_xma07")
        '        Ws.Cells(LineY + 4, 9) = oReader.Item("tc_xma08")
        '        Ws.Cells(LineY + 2, 10) = oReader.Item("tc_xma10")
        '        Ws.Cells(LineY + 8, 1) = oReader.Item("tc_xma11")
        '        Ws.Cells(LineY + 8, 5) = oReader.Item("tc_xma12")
        '        Ws.Cells(LineY + 8, 7) = oReader.Item("tc_xma13")
        '        Ws.Cells(LineY + 10, 8) = oReader.Item("tc_xma14")
        '        Ws.Cells(LineY + 10, 1) = oReader.Item("tc_xma15")
        '        Ws.Cells(LineY + 10, 5) = oReader.Item("tc_xma16")
        '        Ws.Cells(LineY + 10, 6) = oReader.Item("tc_xma17")
        '    End While
        'End If
        'oReader.Close()

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

        'oCommand.CommandText = "select tc_xmb02,tc_xmbud10,tc_xmb03,tc_xmb04,tc_xmb05,tc_xmb06,tc_xmb07,tc_xmb08,tc_xmb09, "
        'oCommand.CommandText += "      tc_xmb10,tc_xmb11,tc_xmb12,tc_xmb13,tc_xmb14,tc_xmb15,tc_xmb16,tc_xmb17,tc_xmb18,tc_xmbud02,tc_xmb19 "
        'oCommand.CommandText += " from tc_xmb_file "
        'oCommand.CommandText += "where tc_xmb01 = '" & l_tc_xma01 & "' "
        'oCommand.CommandText += "  and tc_xmb02 > 24 order by tc_xmb02 "
        'LineZ = 201
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        Ws.Cells(LineZ, 1) = oReader.Item("tc_xmb02")
        '        Ws.Cells(LineZ + 2, 1) = oReader.Item("tc_xmbud10")
        '        Ws.Cells(LineZ, 2) = oReader.Item("tc_xmb03")
        '        Ws.Cells(LineZ, 3) = oReader.Item("tc_xmb04")
        '        Ws.Cells(LineZ + 2, 4) = oReader.Item("tc_xmb11")
        '        Ws.Cells(LineZ + 2, 5) = oReader.Item("tc_xmb12")
        '        Ws.Cells(LineZ, 5) = oReader.Item("tc_xmb13")
        '        Ws.Cells(LineZ, 6) = "千克"
        '        Ws.Cells(LineZ, 7) = oReader.Item("tc_xmb16")
        '        Ws.Cells(LineZ + 1, 7) = oReader.Item("tc_xmb15")
        '        Ws.Cells(LineZ + 2, 7) = "美元"
        '        Ws.Cells(LineZ, 8) = "中国"
        '        Ws.Cells(LineZ + 1, 8) = "（142）"
        '        Ws.Cells(LineZ, 9) = l_tc_xma13
        '        Ws.Cells(LineZ, 10) = "东莞"
        '        Ws.Cells(LineZ, 13) = "全免"
        '        LineZ += 3
        '        If LineZ = 219 Then Exit While
        '    End While
        'End If
        'oReader.Close()

        '装箱单
        Ws = xWorkBook.Sheets(2)
        oCommand.CommandText = "select tc_xma11,tc_xma14,tc_xma29||','||tc_xma30||','||tc_xma31||','||tc_xma32||','||tc_xma33||','||tc_xma34||','||tc_xma35||','||tc_xma36 as tc_xma29_36, "
        oCommand.CommandText += "      ofa0451,ofa0452,ofa0453,ofa0455 "
        oCommand.CommandText += " from ofa_file,tc_xma_file "
        oCommand.CommandText += "where ofa01 = tc_xma29 and tc_xma01 = '" & l_tc_xma01 & "' "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(8, 3) = oReader.Item("tc_xma29_36")
                Ws.Cells(9, 3) = Today.AddDays(0).ToString("yyyy/MM/dd")
                Ws.Cells(10, 3) = oReader.Item("tc_xma11")
                'Ws.Cells(11, 3) = oReader.Item("ofa0451")
                'Ws.Cells(12, 3) = oReader.Item("ofa0452")
                'Ws.Cells(13, 3) = oReader.Item("ofa0453")
                'Ws.Cells(14, 3) = oReader.Item("ofa0455")
                Ws.Cells(15, 3) = oReader.Item("tc_xma14")
            End While
        End If
        oReader.Close()

        oCommand.CommandText = "select tc_xmb02,tc_xmb04,tc_xmb05,tc_xmb06,tc_xmb07,tc_xmb08,tc_xmb09, "
        oCommand.CommandText += "      tc_xmb10,tc_xmb11,tc_xmb12,tc_xmb13,tc_xmb14,tc_xmb15,tc_xmb16,tc_xmb17,tc_xmb18,tc_xmbud02,tc_xmb19 "
        oCommand.CommandText += " from tc_xmb_file "
        oCommand.CommandText += "where tc_xmb01 = '" & l_tc_xma01 & "' "
        oCommand.CommandText += "order by tc_xmb02 "
        LineZ = 20
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("tc_xmbud02")
                Ws.Cells(LineZ, 3) = oReader.Item("tc_xmb04")
                Ws.Cells(LineZ, 4) = oReader.Item("tc_xmb11")
                Ws.Cells(LineZ, 5) = oReader.Item("tc_xmb12")
                Ws.Cells(LineZ, 6) = oReader.Item("tc_xmb13")
                Ws.Cells(LineZ, 7) = oReader.Item("tc_xmb17")
                Ws.Cells(LineZ, 8) = oReader.Item("tc_xmb19")
                LineZ += 1
            End While
        End If
        oReader.Close()
        Ws.Cells(54, 4) = "=SUM(D20:D" & LineZ - 1 & ")"
        Ws.Cells(54, 6) = "=SUM(F20:F" & LineZ - 1 & ")"
        Ws.Cells(54, 7) = "=SUM(G20:G" & LineZ - 1 & ")"
        Ws.Cells(54, 8) = "=SUM(H20:H" & LineZ - 1 & ")"

        '发票
        Ws = xWorkBook.Sheets(3)
        oCommand.CommandText = "select tc_xma11,tc_xma14,tc_xma29||','||tc_xma30||','||tc_xma31||','||tc_xma32||','||tc_xma33||','||tc_xma34||','||tc_xma35||','||tc_xma36 as tc_xma29_36, "
        oCommand.CommandText += "      ofa0451,ofa0452,ofa0453,ofa0455,ofa23 "
        oCommand.CommandText += " from ofa_file,tc_xma_file "
        oCommand.CommandText += "where ofa01 = tc_xma29 and tc_xma01 = '" & l_tc_xma01 & "' "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(8, 3) = oReader.Item("tc_xma29_36")
                Ws.Cells(9, 3) = Today.AddDays(0).ToString("yyyy/MM/dd")
                Ws.Cells(10, 3) = oReader.Item("tc_xma11")
                'Ws.Cells(12, 3) = oReader.Item("ofa0451")
                'Ws.Cells(13, 3) = oReader.Item("ofa0452")
                'Ws.Cells(14, 3) = oReader.Item("ofa0453")
                'Ws.Cells(15, 3) = oReader.Item("ofa0455")
                Ws.Cells(16, 3) = oReader.Item("tc_xma14")
                l_ofa23 = oReader.Item("ofa23")
            End While
        End If
        oReader.Close()

        oCommand.CommandText = "select tc_xmb02,tc_xmb04,tc_xmb05,tc_xmb06,tc_xmb07,tc_xmb08,tc_xmb09, "
        oCommand.CommandText += "      tc_xmb10,tc_xmb11,tc_xmb12,tc_xmb13,tc_xmb14,tc_xmb15,tc_xmb16,tc_xmb17,tc_xmb18,tc_xmbud02,tc_xmb19 "
        oCommand.CommandText += " from tc_xmb_file "
        oCommand.CommandText += "where tc_xmb01 = '" & l_tc_xma01 & "' "
        oCommand.CommandText += "order by tc_xmb02 "
        LineZ = 20
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item("tc_xmb04")
                Ws.Cells(LineZ, 3) = oReader.Item("tc_xmb11")
                Ws.Cells(LineZ, 4) = oReader.Item("tc_xmb12")
                Ws.Cells(LineZ, 5) = oReader.Item("tc_xmb13")
                Ws.Cells(LineZ, 6) = oReader.Item("tc_xmb16")
                Ws.Cells(LineZ, 7) = oReader.Item("tc_xmb15")
                LineZ += 1
            End While
        End If
        oReader.Close()
        Ws.Cells(54, 3) = "=SUM(C20:D" & LineZ - 1 & ")"
        Ws.Cells(54, 5) = "=SUM(E20:E" & LineZ - 1 & ")"
        Ws.Cells(54, 6) = l_ofa23
        Ws.Cells(54, 7) = "=SUM(G20:G" & LineZ - 1 & ")"


        '申报要素
        Ws = xWorkBook.Sheets(4)
        oCommand.CommandText = "select tc_xmb02,tc_xmbud10,tc_xmb03,tc_xmb04,tc_xmb05,tc_xmb06,tc_xmb07,tc_xmb08,tc_xmb09, "
        oCommand.CommandText += "      tc_xmb10,tc_xmb11,tc_xmb12,tc_xmb13,tc_xmb14,tc_xmb15,tc_xmb16,tc_xmb17,tc_xmb18,tc_xmbud02,tc_xmb19,tc_xmbud01 "
        oCommand.CommandText += " from tc_xmb_file "
        oCommand.CommandText += "where tc_xmb01 = '" & l_tc_xma01 & "' "
        oCommand.CommandText += "order by tc_xmb02 "
        LineZ = 7
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("tc_xmb02")
                l_tc_xmb02 = oReader.Item("tc_xmb02")                '190104 add by Brady
                Ws.Cells(LineZ, 2) = oReader.Item("tc_xmb03")
                Ws.Cells(LineZ, 3) = oReader.Item("tc_xmb04")
                Ws.Cells(LineZ, 4) = oReader.Item("tc_xmb05")
                Ws.Cells(LineZ, 5) = oReader.Item("tc_xmb06")
                Ws.Cells(LineZ, 6) = oReader.Item("tc_xmb07")
                Ws.Cells(LineZ, 7) = oReader.Item("tc_xmb08")
                Ws.Cells(LineZ, 8) = oReader.Item("tc_xmb09")
                Ws.Cells(LineZ + 1, 5) = oReader.Item("tc_xmbud01")
                LineZ += 2
                If LineZ = 37 Then Exit While
            End While
        End If
        oReader.Close()

        oCommand.CommandText = "select tc_xmb02,tc_xmbud10,tc_xmb03,tc_xmb04,tc_xmb05,tc_xmb06,tc_xmb07,tc_xmb08,tc_xmb09, "
        oCommand.CommandText += "      tc_xmb10,tc_xmb11,tc_xmb12,tc_xmb13,tc_xmb14,tc_xmb15,tc_xmb16,tc_xmb17,tc_xmb18,tc_xmbud02,tc_xmb19,tc_xmbud01 "
        oCommand.CommandText += " from tc_xmb_file "
        oCommand.CommandText += "where tc_xmb01 = '" & l_tc_xma01 & "' "
        '190104 add by Brady
        'oCommand.CommandText += "  and tc_xmb02 > 15 order by tc_xmb02 "
        oCommand.CommandText += "  and tc_xmb02 > " & l_tc_xmb02 & "order by tc_xmb02 "
        '190104 add by Brady END
        LineZ = 44
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("tc_xmb02")
                l_tc_xmb02 = oReader.Item("tc_xmb02")                '190104 add by Brady
                Ws.Cells(LineZ, 2) = oReader.Item("tc_xmb03")
                Ws.Cells(LineZ, 3) = oReader.Item("tc_xmb04")
                Ws.Cells(LineZ, 4) = oReader.Item("tc_xmb05")
                Ws.Cells(LineZ, 5) = oReader.Item("tc_xmb06")
                Ws.Cells(LineZ, 6) = oReader.Item("tc_xmb07")
                Ws.Cells(LineZ, 7) = oReader.Item("tc_xmb08")
                Ws.Cells(LineZ, 8) = oReader.Item("tc_xmb09")
                Ws.Cells(LineZ + 1, 5) = oReader.Item("tc_xmbud01")
                LineZ += 2
                If LineZ = 74 Then Exit While
            End While
        End If
        oReader.Close()

        oCommand.CommandText = "select tc_xmb02,tc_xmbud10,tc_xmb03,tc_xmb04,tc_xmb05,tc_xmb06,tc_xmb07,tc_xmb08,tc_xmb09, "
        oCommand.CommandText += "      tc_xmb10,tc_xmb11,tc_xmb12,tc_xmb13,tc_xmb14,tc_xmb15,tc_xmb16,tc_xmb17,tc_xmb18,tc_xmbud02,tc_xmb19,tc_xmbud01 "
        oCommand.CommandText += " from tc_xmb_file "
        oCommand.CommandText += "where tc_xmb01 = '" & l_tc_xma01 & "' "
        '190104 add by Brady
        'oCommand.CommandText += "  and tc_xmb02 > 30 order by tc_xmb02 "
        oCommand.CommandText += "  and tc_xmb02 > " & l_tc_xmb02 & "order by tc_xmb02 "
        '190104 add by Brady END
        LineZ = 80
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("tc_xmb02")
                Ws.Cells(LineZ, 2) = oReader.Item("tc_xmb03")
                Ws.Cells(LineZ, 3) = oReader.Item("tc_xmb04")
                Ws.Cells(LineZ, 4) = oReader.Item("tc_xmb05")
                Ws.Cells(LineZ, 5) = oReader.Item("tc_xmb06")
                Ws.Cells(LineZ, 6) = oReader.Item("tc_xmb07")
                Ws.Cells(LineZ, 7) = oReader.Item("tc_xmb08")
                Ws.Cells(LineZ, 8) = oReader.Item("tc_xmb09")
                Ws.Cells(LineZ + 1, 5) = oReader.Item("tc_xmbud01")
                LineZ += 2
                If LineZ = 110 Then Exit While
            End While
        End If
        oReader.Close()

        '联邦报关单
        Ws = xWorkBook.Sheets(5)
        oCommand.CommandText = "select tc_xma03,tc_xma04,tc_xma05,tc_xma06,tc_xma07,tc_xma08,tc_xma09,tc_xma10,tc_xma11, "
        oCommand.CommandText += "      tc_xma12,tc_xma13,tc_xma14,tc_xma15,tc_xma16,tc_xma17,tc_xma18,tc_xma19,tc_xma20, "
        oCommand.CommandText += "      tc_xma21,tc_xma22,tc_xma23,tc_xma24,tc_xma25,tc_xma26,tc_xma27,tc_xma28,tc_xmaud02, "
        oCommand.CommandText += "      ofa0451,ofa0457"
        oCommand.CommandText += " from tc_xma_file,ofa_file "
        oCommand.CommandText += "where ofa01 = tc_xma29 and tc_xma01 = '" & l_tc_xma01 & "' "
        LineY = 4
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineY, 14) = oReader.Item("tc_xma10")
                Ws.Cells(LineY + 2, 1) = oReader.Item("ofa0451")
                Ws.Cells(LineY + 2, 6) = oReader.Item("ofa0457")
                Ws.Cells(LineY + 2, 16) = oReader.Item("tc_xma08")
                Ws.Cells(LineY + 4, 11) = oReader.Item("tc_xmaud02")
                Ws.Cells(LineY + 6, 1) = oReader.Item("tc_xma11")
                Ws.Cells(LineY + 6, 11) = oReader.Item("tc_xma13")
                Ws.Cells(LineY + 6, 15) = oReader.Item("tc_xma21")
                Ws.Cells(LineY + 8, 4) = oReader.Item("tc_xma16")
                Ws.Cells(LineY + 8, 6) = oReader.Item("tc_xma17")
                Ws.Cells(LineY + 8, 8) = oReader.Item("tc_xma18")
                Ws.Cells(LineY + 8, 13) = oReader.Item("tc_xma22")
                Ws.Cells(LineY + 8, 15) = oReader.Item("tc_xma23")
                Ws.Cells(LineY + 8, 17) = oReader.Item("tc_xma24")
            End While
        End If
        oReader.Close()

        oCommand.CommandText = "select tc_xmb02,tc_xmbud10,tc_xmb03,tc_xmb04,tc_xmb05,tc_xmb06,tc_xmb07,tc_xmb08,tc_xmb09, "
        oCommand.CommandText += "      tc_xmb10,tc_xmb11,tc_xmb12,tc_xmb13,tc_xmb14,tc_xmb15,tc_xmb16,tc_xmb17,tc_xmb18,tc_xmbud02,tc_xmb19 "
        oCommand.CommandText += " from tc_xmb_file "
        oCommand.CommandText += "where tc_xmb01 = '" & l_tc_xma01 & "' "
        oCommand.CommandText += "order by tc_xmb02 "
        LineZ = 16
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("tc_xmb02")
                Ws.Cells(LineZ, 2) = oReader.Item("tc_xmbud10")
                Ws.Cells(LineZ, 3) = oReader.Item("tc_xmb03")
                Ws.Cells(LineZ, 4) = oReader.Item("tc_xmb04")
                Ws.Cells(LineZ, 5) = oReader.Item("tc_xmb08")
                Ws.Cells(LineZ, 6) = oReader.Item("tc_xmb05")
                Ws.Cells(LineZ, 7) = oReader.Item("tc_xmb06")
                Ws.Cells(LineZ, 8) = oReader.Item("tc_xmb09")
                Ws.Cells(LineZ, 9) = oReader.Item("tc_xmb07")
                Ws.Cells(LineZ, 12) = oReader.Item("tc_xmb13")
                Ws.Cells(LineZ, 13) = oReader.Item("tc_xmb13")
                Ws.Cells(LineZ, 14) = oReader.Item("tc_xmb14")
                Ws.Cells(LineZ, 15) = "中国"
                Ws.Cells(LineZ, 16) = oReader.Item("tc_xmb15")
                Ws.Cells(LineZ, 17) = l_tc_xma19
                Ws.Cells(LineZ, 18) = oReader.Item("tc_xmb16")
                Ws.Cells(LineZ, 19) = "东莞"
                LineZ += 1
            End While
        End If
        oReader.Close()

        '凤岗装箱单
        Ws = xWorkBook.Sheets(6)
        oCommand.CommandText = "select tc_xma03,tc_xma04,tc_xma05,tc_xma06,tc_xma07,tc_xma08,tc_xma09,tc_xma10,tc_xma11, "
        oCommand.CommandText += "      tc_xma12,tc_xma13,tc_xma14,tc_xma15,tc_xma16,tc_xma17,tc_xma18,tc_xma19,tc_xma20, "
        oCommand.CommandText += "      tc_xma21,tc_xma22,tc_xma23,tc_xma24,tc_xma25,tc_xma26,tc_xma27,tc_xma28,tc_xmaud02 "
        oCommand.CommandText += " from tc_xma_file "
        oCommand.CommandText += "where tc_xma01 = '" & l_tc_xma01 & "' "
        LineY = 3
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineY, 7) = oReader.Item("tc_xma10")
                Ws.Cells(LineY + 2, 3) = oReader.Item("tc_xma11")
                Ws.Cells(LineY + 3, 3) = oReader.Item("tc_xma13")
                Ws.Cells(LineY + 3, 6) = oReader.Item("tc_xma13")
                Ws.Cells(LineY + 5, 3) = oReader.Item("tc_xma16")
                Ws.Cells(LineY + 5, 7) = oReader.Item("tc_xma17")
                Ws.Cells(LineY + 5, 9) = oReader.Item("tc_xma18")
                Ws.Cells(LineY + 9, 8) = oReader.Item("tc_xma19")
                Ws.Cells(LineY + 9, 9) = oReader.Item("tc_xma19")
            End While
        End If
        oReader.Close()

        oCommand.CommandText = "select tc_xmb02,tc_xmbud10,tc_xmb03,tc_xmb04,tc_xmb05,tc_xmb06,tc_xmb07,tc_xmb08,tc_xmb09, "
        oCommand.CommandText += "      tc_xmb10,tc_xmb11,tc_xmb12,tc_xmb13,tc_xmb14,tc_xmb15,tc_xmb16,tc_xmb17,tc_xmb18,tc_xmbud02,tc_xmb19 "
        oCommand.CommandText += " from tc_xmb_file "
        oCommand.CommandText += "where tc_xmb01 = '" & l_tc_xma01 & "' "
        oCommand.CommandText += "order by tc_xmb02 "
        LineZ = 13
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("tc_xmb02")
                Ws.Cells(LineZ, 2) = oReader.Item("tc_xmbud10")
                Ws.Cells(LineZ, 3) = oReader.Item("tc_xmb03")
                Ws.Cells(LineZ, 4) = oReader.Item("tc_xmb04")
                Ws.Cells(LineZ, 6) = oReader.Item("tc_xmb13")
                Ws.Cells(LineZ, 7) = oReader.Item("tc_xmb11")
                Ws.Cells(LineZ, 8) = oReader.Item("tc_xmb16")
                Ws.Cells(LineZ, 9) = oReader.Item("tc_xmb15")
                LineZ += 1
            End While
        End If
        oReader.Close()
        Ws.Cells(47, 9) = "=SUM(I13:I" & LineZ - 1 & ")"

    End Sub

    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "export_declaration"
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