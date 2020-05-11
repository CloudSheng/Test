Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel

Public Class Form362
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
    Dim t_ofa01_1 As String = String.Empty
    Dim t_ofa01_2 As String = String.Empty
    Dim t_ofa01_3 As String = String.Empty
    Dim t_ofa01_4 As String = String.Empty
    Dim t_ofa01_5 As String = String.Empty
    Dim t_ofa01_6 As String = String.Empty
    Dim t_ofa01_7 As String = String.Empty
    Dim t_ofa01_8 As String = String.Empty
    Dim t_ofa01_9 As String = String.Empty
    Dim t_ofa01_10 As String = String.Empty
    Dim l_ima021 As String = String.Empty
    Dim l_ofa02 As Date
    Dim l_ofaud05 As String = String.Empty
    Dim l_ogd12b As String = String.Empty
    Dim l_ogd12e As String = String.Empty

    Dim t_ogd15t As Double = 0


    Private Sub Form354_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'If Me.BackgroundWorker1.IsBusy() Then
        'MsgBox("处理中，请等待")
        'Return
        'End If        

        Dim xPath As String = "C:\temp\Seafreight_SI_sample.xlsx"
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
            t_ofa01_1 = TextBox1.Text
        End If
        If Not String.IsNullOrEmpty(TextBox2.Text) Then
            t_ofa01_2 = TextBox2.Text
        End If
        If Not String.IsNullOrEmpty(TextBox3.Text) Then
            t_ofa01_3 = TextBox3.Text
        End If
        If Not String.IsNullOrEmpty(TextBox4.Text) Then
            t_ofa01_4 = TextBox4.Text
        End If
        If Not String.IsNullOrEmpty(TextBox5.Text) Then
            t_ofa01_5 = TextBox5.Text
        End If
        If Not String.IsNullOrEmpty(TextBox6.Text) Then
            t_ofa01_6 = TextBox6.Text
        End If
        If Not String.IsNullOrEmpty(TextBox7.Text) Then
            t_ofa01_7 = TextBox7.Text
        End If
        If Not String.IsNullOrEmpty(TextBox8.Text) Then
            t_ofa01_8 = TextBox8.Text
        End If
        If Not String.IsNullOrEmpty(TextBox9.Text) Then
            t_ofa01_9 = TextBox9.Text
        End If
        If Not String.IsNullOrEmpty(TextBox10.Text) Then
            t_ofa01_10 = TextBox10.Text
        End If

        'xExcel = New Microsoft.Office.Interop.Excel.Application
        'xWorkBook = xExcel.Workbooks.Add()
        ExportToExcel()
        oConnection.Close()

        SaveExcel()
    End Sub

    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        Dim xPath As String = "C:\temp\Seafreight_SI_sample.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)

        '艾可迅补料
        Ws = xWorkBook.Sheets(1)

        '190815 add by Brady 修正抓取 ta_obk15的取值邏輯
        ''190705 add by Brady
        ''oCommand.CommandText = " select ogd12b,ogd12e,ofb06,ogd10,ogd15t,ogd16t "
        ''oCommand.CommandText += "  from ofa_file,ofb_file,ogd_file "
        ''oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "','" & t_ofa01_9 & "','" & t_ofa01_10 & "') "
        ''oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 "
        ''oCommand.CommandText += " order by ogd12b "
        ''oCommand.CommandText = " select ogd12b,ogd12e,ta_obk15,ogd10,ogd15t,ogd16t "                                                            '190814 mark by Brady
        'oCommand.CommandText = " select ogd12b,ogd12e,NVL(ta_obk15,' ') as ta_obk15,ogd10,NVL(ogd15t,0) as ogd15t,NVL(ogd16t,0) as ogd16t "      '190814 add by Brady 修正 ogd15t,ogd16t不可為空值的Bug
        'oCommand.CommandText += "  from ogd_file,ofa_file,ofb_file,obk_file "
        'oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "','" & t_ofa01_9 & "','" & t_ofa01_10 & "') "
        'oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 "
        'oCommand.CommandText += "   and obk02 = ofa04 and obk01 = ofb04 and obk03 = ofb11 and obkacti = 'Y' "
        'oCommand.CommandText += " order by ogd12b "
        ''190705 add by Brady END
        'LineZ = 28
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        Ws.Cells(LineZ, 3) = "8708299000"
        '        l_ogd12b = oReader.Item("ogd12b")
        '        l_ogd12e = oReader.Item("ogd12e")
        '        Ws.Cells(LineZ, 4) = l_ogd12b + "-" + l_ogd12e
        '        '190704 add by Brady
        '        'Ws.Cells(LineZ, 5) = oReader.Item("ofb06")
        '        Ws.Cells(LineZ, 5) = oReader.Item("ta_obk15")
        '        '190704 add by Brady END
        '        Ws.Cells(LineZ, 6) = oReader.Item("ogd10")
        '        Ws.Cells(LineZ, 7) = oReader.Item("ogd15t")
        '        Ws.Cells(LineZ, 8) = oReader.Item("ogd16t")
        '        LineZ += 1
        '    End While
        'End If
        'oReader.Close()
        oCommand.CommandText = " select ogd12b,ogd12e,ogd10,NVL(ogd15t,0) as ogd15t,NVL(ogd16t,0) as ogd16t,ofa04,ofb04,ofb11 "
        oCommand.CommandText += "  from ogd_file,ofa_file,ofb_file "
        oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "','" & t_ofa01_9 & "','" & t_ofa01_10 & "') "
        oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 "
        oCommand.CommandText += " order by ogd12b "
        LineZ = 28
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 3) = "8708299000"
                l_ogd12b = oReader.Item("ogd12b")
                l_ogd12e = oReader.Item("ogd12e")
                Ws.Cells(LineZ, 4) = l_ogd12b + "-" + l_ogd12e
                Dim l_ofa04 As String = String.Empty
                Dim l_ofb04 As String = String.Empty
                Dim l_ofb11 As String = String.Empty
                l_ofa04 = oReader.Item("ofa04")
                l_ofb04 = oReader.Item("ofb04")
                l_ofb11 = oReader.Item("ofb11")
                oCommand2.CommandText = " select NVL(ta_obk15,' ') as ta_obk15 from obk_file "
                oCommand2.CommandText += " where obk02 = '" & l_ofa04 & "' and obk01 = '" & l_ofb04 & "' and obk03 = '" & l_ofb11 & "'"
                oCommand2.CommandText += "   and obkacti = 'Y' "
                oReader2 = oCommand2.ExecuteReader()
                If oReader2.HasRows() Then
                    While oReader2.Read()
                        Ws.Cells(LineZ, 5) = oReader2.Item("ta_obk15")
                    End While
                End If
                oReader2.Close()
                Ws.Cells(LineZ, 6) = oReader.Item("ogd10")
                Ws.Cells(LineZ, 7) = oReader.Item("ogd15t")
                Ws.Cells(LineZ, 8) = oReader.Item("ogd16t")
                LineZ += 1
            End While
        End If
        oReader.Close()
        '190815 add by Brady END

        '190909 add by Brady 
        'Ws.Cells(92, 6) = "=SUM(F28:F" & LineZ - 1 & ")"
        'Ws.Cells(92, 7) = "=SUM(G28:G" & LineZ - 1 & ")"
        'Ws.Cells(92, 8) = "=SUM(H28:H" & LineZ - 1 & ")"
        LineZ += 1
        Ws.Cells(LineZ, 5) = "Total :"
        Ws.Cells(LineZ, 6) = "=SUM(F28:F" & LineZ - 2 & ")"
        Ws.Cells(LineZ, 7) = "=SUM(G28:G" & LineZ - 2 & ")"
        Ws.Cells(LineZ, 8) = "=SUM(H28:H" & LineZ - 2 & ")"
        LineZ += 2
        Ws.Cells(LineZ, 1) = "Container size:"
        Ws.Cells(LineZ + 1, 1) = "Container no.:"
        Ws.Cells(LineZ + 2, 1) = "Seal no.:"
        Ws.Cells(LineZ + 4, 1) = "Shipping Term (incoterms)(交货条款） :EX-WORK"
        '190909 add by Brady END

    End Sub

    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Seafreight_SI"
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