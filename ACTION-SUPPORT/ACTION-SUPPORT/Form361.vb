Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel

Public Class Form361
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
    Dim t_ogd12b As Integer = 0             '200104 add by Brady
    Dim t_ogd12e As Integer = 0             '200104 add by Brady
    Dim old_ogd12e As Integer = 0           '200113 add by Brady
    Dim l_cnt As Integer = 0                '200113 add by Brady
    Dim t_ogd15t As Double = 0
    

    Private Sub Form354_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'If Me.BackgroundWorker1.IsBusy() Then
        'MsgBox("处理中，请等待")
        'Return
        'End If        

        Dim xPath As String = "C:\temp\Trainfreight_SI_sample.xlsx"
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
        Dim xPath As String = "C:\temp\Trainfreight_SI_sample.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)

        '191030 add by Brady
        ''PACKLING LIST
        'Ws = xWorkBook.Sheets(1)
        ''oCommand.CommandText = " select ima021,ofa02,ofaud05 "                      '190814 mark by Brady
        'oCommand.CommandText = " select ima021,ofa02,NVL(ofaud05,' ') as ofaud05 "   '190814 add by Brady 修正 ofaud05不可為空值的Bug
        'oCommand.CommandText += "  from ofa_file,ofb_file,ogd_file,ima_file "
        'oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "','" & t_ofa01_9 & "','" & t_ofa01_10 & "') "
        'oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 and ofb04 = ima01 "
        'oCommand.CommandText += " order by ogd12b "
        'LineZ = 4
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        l_ima021 = oReader.Item("ima021")
        '        Ws.Cells(LineZ, 6) = l_ima021
        '        l_ofa02 = oReader.Item("ofa02")
        '        Ws.Cells(LineZ + 2, 6) = l_ofa02
        '        l_ofaud05 = oReader.Item("ofaud05")
        '        Ws.Cells(LineZ + 4, 6) = l_ofaud05
        '        Exit While
        '    End While
        'End If
        'oReader.Close()
        ''190815 add by Brady 修正抓取 ta_obk15的取值邏輯
        ' ''190704 add by Brady CS告知要修正 [Description] 欄位的取值邏輯
        '' ''190626 add by Brady CS告知要修正 [PCS][NetWeight(KGS)] 兩個欄位的取值邏輯
        '' ''oCommand.CommandText = " select ofb06,ogd12b,ogd12e,ogd09,ogd14,ogd15t,ofa01 "
        ' ''oCommand.CommandText = " select ofb06,ogd12b,ogd12e,ogd13,ogd14t,ogd15t,ofa01 "
        '' ''190626 add by Brady END
        ' ''oCommand.CommandText += "  from ofa_file,ofb_file,ogd_file,ima_file "
        ' ''oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "') "
        ' ''oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 and ofb04 = ima01 "
        ' ''oCommand.CommandText += " order by ogd12b "
        ' ''oCommand.CommandText = " select ta_obk15,ogd12b,ogd12e,ogd13,ogd14t,ogd15t,ofa01 "                                        '190814 mark by Brady
        ''oCommand.CommandText = " select NVL(ta_obk15,' ') as ta_obk15,ogd12b,ogd12e,ogd13,ogd14t,NVL(ogd15t,0) as ogd15t,ofa01 "   '190814 add by Brady 修正 ogd15t不可為空值的Bug
        ''oCommand.CommandText += "  from ofa_file,ofb_file,ogd_file,ima_file,obk_file "
        ''oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "','" & t_ofa01_9 & "','" & t_ofa01_10 & "') "
        ''oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 and ofb04 = ima01 "
        ''oCommand.CommandText += "   and obk02 = ofa04 and obk01 = ofb04 and obk03 = ofb11 and obkacti = 'Y' "
        ''oCommand.CommandText += " order by ogd12b "
        ' ''190704 add by Brady END
        ''LineZ = 11
        ''oReader = oCommand.ExecuteReader()
        ''If oReader.HasRows() Then
        ''    While oReader.Read()
        ''        Ws.Cells(LineZ, 1) = "870829900"
        ''        '190704 add by Brady CS告知要修正 [Description] 欄位的取值邏輯
        ''        'Ws.Cells(LineZ, 2) = oReader.Item("ofb06")
        ''        Ws.Cells(LineZ, 2) = oReader.Item("ta_obk15")
        ''        '190704 add by Brady END
        ''        l_ogd12b = oReader.Item("ogd12b")
        ''        l_ogd12e = oReader.Item("ogd12e")
        ''        Ws.Cells(LineZ, 3) = l_ogd12b + "-" + l_ogd12e
        ''        '190626 add by Brady CS告知要修正 [PCS][NetWeight(KGS)] 兩個欄位的取值邏輯
        ''        'Ws.Cells(LineZ, 4) = oReader.Item("ogd09")
        ''        'Ws.Cells(LineZ, 5) = oReader.Item("ogd14")
        ''        Ws.Cells(LineZ, 4) = oReader.Item("ogd13")
        ''        Ws.Cells(LineZ, 5) = oReader.Item("ogd14t")
        ''        '190626 add by Brady END 
        ''        Ws.Cells(LineZ, 6) = oReader.Item("ogd15t")
        ''        t_ogd15t = t_ogd15t + oReader.Item("ogd15t")
        ''        Ws.Cells(LineZ, 7) = oReader.Item("ofa01")
        ''        LineZ += 1
        ''        If LineZ = 51 Then
        ''            LineZ = 56
        ''            Ws.Cells(LineZ, 6) = l_ima021
        ''            Ws.Cells(LineZ + 2, 6) = l_ofa02
        ''            Ws.Cells(LineZ + 4, 6) = l_ofaud05
        ''            LineZ = 63
        ''        End If
        ''    End While
        ''End If
        ''oReader.Close()
        'oCommand.CommandText = " select ogd12b,ogd12e,ogd13,ogd14t,NVL(ogd15t,0) as ogd15t,ofa01,ofa04,ofb04,ofb11 "
        'oCommand.CommandText += "  from ofa_file,ofb_file,ogd_file,ima_file "
        'oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "','" & t_ofa01_9 & "','" & t_ofa01_10 & "') "
        'oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 and ofb04 = ima01 "
        'oCommand.CommandText += " order by ogd12b "
        'LineZ = 11
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        Ws.Cells(LineZ, 1) = "870829900"
        '        Dim l_ofa04 As String = String.Empty
        '        Dim l_ofb04 As String = String.Empty
        '        Dim l_ofb11 As String = String.Empty
        '        l_ofa04 = oReader.Item("ofa04")
        '        l_ofb04 = oReader.Item("ofb04")
        '        l_ofb11 = oReader.Item("ofb11")
        '        oCommand2.CommandText = " select NVL(ta_obk15,' ') as ta_obk15 from obk_file "
        '        oCommand2.CommandText += " where obk02 = '" & l_ofa04 & "' and obk01 = '" & l_ofb04 & "' and obk03 = '" & l_ofb11 & "'"
        '        oCommand2.CommandText += "   and obkacti = 'Y' "
        '        oReader2 = oCommand2.ExecuteReader()
        '        If oReader2.HasRows() Then
        '            While oReader2.Read()
        '                Ws.Cells(LineZ, 2) = oReader2.Item("ta_obk15")
        '            End While
        '        End If
        '        oReader2.Close()
        '        l_ogd12b = oReader.Item("ogd12b")
        '        l_ogd12e = oReader.Item("ogd12e")
        '        Ws.Cells(LineZ, 3) = l_ogd12b + "-" + l_ogd12e
        '        Ws.Cells(LineZ, 4) = oReader.Item("ogd13")
        '        Ws.Cells(LineZ, 5) = oReader.Item("ogd14t")
        '        Ws.Cells(LineZ, 6) = oReader.Item("ogd15t")
        '        t_ogd15t = t_ogd15t + oReader.Item("ogd15t")
        '        Ws.Cells(LineZ, 7) = oReader.Item("ofa01")
        '        LineZ += 1
        '        If LineZ = 51 Then
        '            LineZ = 56
        '            Ws.Cells(LineZ, 6) = l_ima021
        '            Ws.Cells(LineZ + 2, 6) = l_ofa02
        '            Ws.Cells(LineZ + 4, 6) = l_ofaud05
        '            LineZ = 63
        '        End If
        '    End While
        'End If
        'oReader.Close()
        ''190815 add by Brady END
        'If LineZ < 51 Then
        '    Ws.Cells(52, 3) = l_ogd12e
        '    Ws.Cells(52, 4) = "=SUM(D11:D" & LineZ - 1 & ")"
        '    Ws.Cells(52, 6) = "=SUM(F11:F" & LineZ - 1 & ")"
        'Else
        '    Ws.Cells(99, 3) = l_ogd12e
        '    Ws.Cells(99, 4) = "=SUM(D11:D" & LineZ - 1 & ")"
        '    Ws.Cells(99, 6) = t_ogd15t

        'End If

        ''INVOICE-USD
        'Ws = xWorkBook.Sheets(2)
        'oCommand.CommandText = " select ima021,ofa02,NVL(ofaud05,' ') as ofaud05 "
        'oCommand.CommandText += "  from ofa_file,ofb_file,ogd_file,ima_file "
        'oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "','" & t_ofa01_9 & "','" & t_ofa01_10 & "') "
        'oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 and ofb04 = ima01 "
        'oCommand.CommandText += " order by ogd12b "
        'LineZ = 4
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        l_ima021 = oReader.Item("ima021")
        '        Ws.Cells(LineZ, 5) = l_ima021
        '        l_ofa02 = oReader.Item("ofa02")
        '        Ws.Cells(LineZ + 2, 5) = l_ofa02
        '        l_ofaud05 = oReader.Item("ofaud05")
        '        Ws.Cells(LineZ + 4, 5) = l_ofaud05
        '        Exit While
        '    End While
        'End If
        'oReader.Close()
        ''190815 add by Brady 修正抓取 ta_obk15的取值邏輯
        ' ''190704 add by Brady CS告知要修正 [Description] 欄位的取值邏輯
        '' ''190626 add by Brady CS告知要修正 [PCS] 欄位的取值邏輯
        '' ''oCommand.CommandText = " select ofb06,ogd09,ofb13,ofb14,ofa01 "
        ' ''oCommand.CommandText = " select ofb06,ogd13,ofb13,ofb14,ofa01 "
        '' ''190626 add by Brady END
        ' ''oCommand.CommandText += "  from ofa_file,ofb_file,ogd_file,ima_file "
        ' ''oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "') "
        ' ''oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 and ofb04 = ima01 and ofa23 = 'USD' "
        ' ''oCommand.CommandText += " order by ogd12b "        
        ''oCommand.CommandText = " select NVL(ta_obk15,' ') as ta_obk15,ogd13,ofb13,ofb14,ofa01 "
        ''oCommand.CommandText += "  from ofa_file,ofb_file,ogd_file,ima_file,obk_file "
        ''oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "','" & t_ofa01_9 & "','" & t_ofa01_10 & "') "
        ''oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 and ofb04 = ima01 and ofa23 = 'USD' "
        ''oCommand.CommandText += "   and obk02 = ofa04 and obk01 = ofb04 and obk03 = ofb11 and obkacti = 'Y' "
        ''oCommand.CommandText += " order by ogd12b "
        ' ''190704 add by Brady END
        ''LineZ = 11
        ''oReader = oCommand.ExecuteReader()
        ''If oReader.HasRows() Then
        ''    While oReader.Read()
        ''        Ws.Cells(LineZ, 1) = "870829900"
        ''        '190704 add by Brady CS告知要修正 [Description] 欄位的取值邏輯
        ''        'Ws.Cells(LineZ, 2) = oReader.Item("ofb06")
        ''        Ws.Cells(LineZ, 2) = oReader.Item("ta_obk15")
        ''        '190704 add by Brady END
        ''        '190626 add by Brady CS告知要修正 [PCS] 欄位的取值邏輯
        ''        'Ws.Cells(LineZ, 3) = oReader.Item("ogd09")
        ''        Ws.Cells(LineZ, 3) = oReader.Item("ogd13")
        ''        '190626 add by Brady END
        ''        Ws.Cells(LineZ, 4) = oReader.Item("ofb13")
        ''        Ws.Cells(LineZ, 5) = oReader.Item("ofb14")
        ''        Ws.Cells(LineZ, 6) = oReader.Item("ofa01")
        ''        LineZ += 1
        ''    End While
        ''End If
        ''oReader.Close()          
        'oCommand.CommandText = " select ogd13,ofb13,ofb14,ofa01,ofa04,ofb04,ofb11 "
        'oCommand.CommandText += "  from ofa_file,ofb_file,ogd_file,ima_file "
        'oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "','" & t_ofa01_9 & "','" & t_ofa01_10 & "') "
        'oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 and ofb04 = ima01 and ofa23 = 'USD' "
        'oCommand.CommandText += " order by ogd12b "
        'LineZ = 11
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        Ws.Cells(LineZ, 1) = "870829900"
        '        Dim l_ofa04 As String = String.Empty
        '        Dim l_ofb04 As String = String.Empty
        '        Dim l_ofb11 As String = String.Empty
        '        l_ofa04 = oReader.Item("ofa04")
        '        l_ofb04 = oReader.Item("ofb04")
        '        l_ofb11 = oReader.Item("ofb11")
        '        oCommand2.CommandText = " select NVL(ta_obk15,' ') as ta_obk15 from obk_file "
        '        oCommand2.CommandText += " where obk02 = '" & l_ofa04 & "' and obk01 = '" & l_ofb04 & "' and obk03 = '" & l_ofb11 & "'"
        '        oCommand2.CommandText += "   and obkacti = 'Y' "
        '        oReader2 = oCommand2.ExecuteReader()
        '        If oReader2.HasRows() Then
        '            While oReader2.Read()
        '                Ws.Cells(LineZ, 2) = oReader2.Item("ta_obk15")
        '            End While
        '        End If
        '        oReader2.Close()
        '        Ws.Cells(LineZ, 3) = oReader.Item("ogd13")
        '        Ws.Cells(LineZ, 4) = oReader.Item("ofb13")
        '        Ws.Cells(LineZ, 5) = oReader.Item("ofb14")
        '        Ws.Cells(LineZ, 6) = oReader.Item("ofa01")
        '        LineZ += 1
        '    End While
        'End If
        'oReader.Close()
        ''190815 add by Brady END
        'Ws.Cells(53, 3) = "=SUM(C11:C" & LineZ - 1 & ")"
        'Ws.Cells(53, 5) = "=SUM(E11:E" & LineZ - 1 & ")"

        ''INVOICE-EUR
        'Ws = xWorkBook.Sheets(3)
        'oCommand.CommandText = " select ima021,ofa02,NVL(ofaud05,' ') as ofaud05"
        'oCommand.CommandText += "  from ofa_file,ofb_file,ogd_file,ima_file "
        'oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "','" & t_ofa01_9 & "','" & t_ofa01_10 & "') "
        'oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 and ofb04 = ima01 "
        'oCommand.CommandText += " order by ogd12b "
        'LineZ = 4
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        l_ima021 = oReader.Item("ima021")
        '        Ws.Cells(LineZ, 5) = l_ima021
        '        l_ofa02 = oReader.Item("ofa02")
        '        Ws.Cells(LineZ + 2, 5) = l_ofa02
        '        l_ofaud05 = oReader.Item("ofaud05")
        '        Ws.Cells(LineZ + 4, 5) = l_ofaud05
        '        Exit While
        '    End While
        'End If
        'oReader.Close()
        ''190815 add by Brady 修正抓取 ta_obk15的取值邏輯
        ' ''190704 add by Brady CS告知要修正 [Description] 欄位的取值邏輯
        '' ''190626 add by Brady CS告知要修正 [PCS] 欄位的取值邏輯
        '' ''oCommand.CommandText = " select ofb06,ogd09,ofb13,ofb14,ofa01 "
        ' ''oCommand.CommandText = " select ofb06,ogd13,ofb13,ofb14,ofa01 "
        '' ''190626 add by Brady END
        ' ''oCommand.CommandText += "  from ofa_file,ofb_file,ogd_file,ima_file "
        ' ''oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "') "
        ' ''oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 and ofb04 = ima01 and ofa23 = 'EUR' "
        ' ''oCommand.CommandText += " order by ogd12b "        
        ''oCommand.CommandText = " select NVL(ta_obk15,' ') as ta_obk15,ogd13,ofb13,ofb14,ofa01 "
        ''oCommand.CommandText += "  from ofa_file,ofb_file,ogd_file,ima_file,obk_file "
        ''oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "','" & t_ofa01_9 & "','" & t_ofa01_10 & "') "
        ''oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 and ofb04 = ima01 and ofa23 = 'EUR' "
        ''oCommand.CommandText += "   and obk02 = ofa04 and obk01 = ofb04 and obk03 = ofb11 and obkacti = 'Y' "
        ''oCommand.CommandText += " order by ogd12b "
        ' ''190704 add by Brady END
        ''LineZ = 11
        ''oReader = oCommand.ExecuteReader()
        ''If oReader.HasRows() Then
        ''    While oReader.Read()
        ''        Ws.Cells(LineZ, 1) = "870829900"
        ''        '190704 add by Brady CS告知要修正 [Description] 欄位的取值邏輯
        ''        'Ws.Cells(LineZ, 2) = oReader.Item("ofb06")
        ''        Ws.Cells(LineZ, 2) = oReader.Item("ta_obk15")
        ''        '190704 add by Brady END
        ''        '190626 add by Brady CS告知要修正 [PCS] 欄位的取值邏輯
        ''        'Ws.Cells(LineZ, 3) = oReader.Item("ogd09")
        ''        Ws.Cells(LineZ, 3) = oReader.Item("ogd13")
        ''        '190626 add by Brady END
        ''        Ws.Cells(LineZ, 4) = oReader.Item("ofb13")
        ''        Ws.Cells(LineZ, 5) = oReader.Item("ofb14")
        ''        Ws.Cells(LineZ, 6) = oReader.Item("ofa01")
        ''        LineZ += 1
        ''    End While
        ''End If
        ''oReader.Close()            
        'oCommand.CommandText = " select ogd13,ofb13,ofb14,ofa01,ofa04,ofb04,ofb11 "
        'oCommand.CommandText += "  from ofa_file,ofb_file,ogd_file,ima_file "
        'oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "','" & t_ofa01_9 & "','" & t_ofa01_10 & "') "
        'oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 and ofb04 = ima01 and ofa23 = 'EUR' "
        'oCommand.CommandText += " order by ogd12b "
        'LineZ = 11
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        Ws.Cells(LineZ, 1) = "870829900"
        '        Dim l_ofa04 As String = String.Empty
        '        Dim l_ofb04 As String = String.Empty
        '        Dim l_ofb11 As String = String.Empty
        '        l_ofa04 = oReader.Item("ofa04")
        '        l_ofb04 = oReader.Item("ofb04")
        '        l_ofb11 = oReader.Item("ofb11")
        '        oCommand2.CommandText = " select NVL(ta_obk15,' ') as ta_obk15 from obk_file "
        '        oCommand2.CommandText += " where obk02 = '" & l_ofa04 & "' and obk01 = '" & l_ofb04 & "' and obk03 = '" & l_ofb11 & "'"
        '        oCommand2.CommandText += "   and obkacti = 'Y' "
        '        oReader2 = oCommand2.ExecuteReader()
        '        If oReader2.HasRows() Then
        '            While oReader2.Read()
        '                Ws.Cells(LineZ, 2) = oReader2.Item("ta_obk15")
        '            End While
        '        End If
        '        oReader2.Close()
        '        Ws.Cells(LineZ, 3) = oReader.Item("ogd13")
        '        Ws.Cells(LineZ, 4) = oReader.Item("ofb13")
        '        Ws.Cells(LineZ, 5) = oReader.Item("ofb14")
        '        Ws.Cells(LineZ, 6) = oReader.Item("ofa01")
        '        LineZ += 1
        '    End While
        'End If
        'oReader.Close()
        ''190815 add by Brady END
        'Ws.Cells(53, 3) = "=SUM(C11:C" & LineZ - 1 & ")"
        'Ws.Cells(53, 5) = "=SUM(E11:E" & LineZ - 1 & ")"

        'PACKLING LIST
        Ws = xWorkBook.Sheets(1)
        oCommand.CommandText = " select ima021,ofa02,NVL(ofaud05,' ') as ofaud05 "
        oCommand.CommandText += "  from ofa_file,ofb_file,ogd_file,ima_file "
        oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "','" & t_ofa01_9 & "','" & t_ofa01_10 & "') "
        oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 and ofb04 = ima01 "
        oCommand.CommandText += " order by ogd12b "
        LineZ = 4
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                l_ima021 = oReader.Item("ima021")
                Ws.Cells(LineZ, 7) = l_ima021
                l_ofa02 = oReader.Item("ofa02")
                Ws.Cells(LineZ + 2, 7) = l_ofa02
                l_ofaud05 = oReader.Item("ofaud05")
                Ws.Cells(LineZ + 4, 7) = l_ofaud05
                Exit While
            End While
        End If
        oReader.Close()
        'oCommand.CommandText = " select ogd12b,ogd12e,ogd13,ogd14t,NVL(ogd15t,0) as ogd15t,ofa01,ofa04,ofb04,ofb11 "   '200224 mark by Brady 
        'oCommand.CommandText = " select ogd12b,ogd12e,ogd13,ogd14t,NVL(ogd15t,0) as ogd15t,ofa01,ofa04,ofb04,ofb11,imaud01 "   '200325 mark by Brady  '200224 add by Brady
        oCommand.CommandText = " select ogd12b,ogd12e,ogd13,ogd14t,NVL(ogd15t,0) as ogd15t,ofa01,ofa04,ofb04,ofb11,imaud01,ogd10 "   '200325 add by Brady
        oCommand.CommandText += "  from ofa_file,ofb_file,ogd_file,ima_file "
        oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "','" & t_ofa01_9 & "','" & t_ofa01_10 & "') "
        oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 and ofb04 = ima01 "
        oCommand.CommandText += " order by ogd12b "
        LineZ = 11
        l_cnt = 0        '200113 add by Brady
        old_ogd12e = 0   '200113 add by Brady
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = "870829900"
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
                        Ws.Cells(LineZ, 2) = oReader2.Item("ta_obk15")
                    End While
                End If
                oReader2.Close()

                '200104 add by Brady
                If LineZ = 11 Then
                    t_ogd12b = oReader.Item("ogd12b")
                End If
                t_ogd12e = oReader.Item("ogd12e")
                If old_ogd12e <> t_ogd12e Then
                    l_cnt += 1
                    old_ogd12e = t_ogd12e
                End If
                '200104 add by Brady END

                l_ogd12b = oReader.Item("ogd12b")
                l_ogd12e = oReader.Item("ogd12e")

                '200224 add by Brady
                ''Ws.Cells(LineZ, 3) = l_ogd12b + "-" + l_ogd12e    '200113 mark by Brady
                'Ws.Cells(LineZ, 3) = l_ogd12b                      '200113 add by Brady
                'Ws.Cells(LineZ, 4) = oReader.Item("ogd13")
                'Ws.Cells(LineZ, 5) = oReader.Item("ogd14t")
                'Ws.Cells(LineZ, 6) = oReader.Item("ogd15t")
                't_ogd15t = t_ogd15t + oReader.Item("ogd15t")
                'Ws.Cells(LineZ, 7) = oReader.Item("ofa01")
                Ws.Cells(LineZ, 3) = oReader.Item("imaud01")
                'Ws.Cells(LineZ, 4) = l_ogd12b                  '200325 mark by Brady
                Ws.Cells(LineZ, 4) = oReader.Item("ogd10")      '200325 add by Brady
                Ws.Cells(LineZ, 5) = oReader.Item("ogd13")
                Ws.Cells(LineZ, 6) = oReader.Item("ogd14t")
                Ws.Cells(LineZ, 7) = oReader.Item("ogd15t")
                t_ogd15t = t_ogd15t + oReader.Item("ogd15t")
                Ws.Cells(LineZ, 8) = oReader.Item("ofa01")
                '200224 add by Brady END

                LineZ += 1
            End While
        End If
        oReader.Close()
        '200224 add by Brady
        ''191204 add by Brady
        ''Ws.Cells(150, 3) = l_ogd12e
        ''Ws.Cells(150, 4) = "=SUM(D11:D" & LineZ - 1 & ")"
        ''Ws.Cells(150, 6) = "=SUM(F11:F" & LineZ - 1 & ")"
        ''Ws.Cells(200, 3) = l_ogd12e   '200104 mark by Brady
        ''Ws.Cells(200, 3) = t_ogd12e - t_ogd12b + 1    '200113 mark by Brady  '200104 add by Brady
        'Ws.Cells(200, 3) = l_cnt                       '200113 add by Brady
        'Ws.Cells(200, 4) = "=SUM(D11:D" & LineZ - 1 & ")"
        'Ws.Cells(200, 6) = "=SUM(F11:F" & LineZ - 1 & ")"
        ''191204 add by Brady END
        Ws.Cells(200, 4) = l_cnt
        Ws.Cells(200, 5) = "=SUM(E11:E" & LineZ - 1 & ")"
        Ws.Cells(200, 7) = "=SUM(G11:G" & LineZ - 1 & ")"
        '200224 add by Brady END

        'INVOICE-USD
        Ws = xWorkBook.Sheets(2)
        oCommand.CommandText = " select ima021,ofa02,NVL(ofaud05,' ') as ofaud05 "
        oCommand.CommandText += "  from ofa_file,ofb_file,ogd_file,ima_file "
        oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "','" & t_ofa01_9 & "','" & t_ofa01_10 & "') "
        oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 and ofb04 = ima01 "
        oCommand.CommandText += " order by ogd12b "
        LineZ = 4
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                l_ima021 = oReader.Item("ima021")
                Ws.Cells(LineZ, 6) = l_ima021
                l_ofa02 = oReader.Item("ofa02")
                Ws.Cells(LineZ + 2, 6) = l_ofa02
                l_ofaud05 = oReader.Item("ofaud05")
                Ws.Cells(LineZ + 4, 6) = l_ofaud05
                Exit While
            End While
        End If
        oReader.Close()
        'oCommand.CommandText = " select ogd13,ogbud07 as ofb13,ogd13*ogbud07 as ofb14,ofa01,ofa04,ofb04,ofb11 "         '200224 mark by Brady
        oCommand.CommandText = " select ogd13,ogbud07 as ofb13,ogd13*ogbud07 as ofb14,ofa01,ofa04,ofb04,ofb11,imaud01 "  '200224 add by Brady
        oCommand.CommandText += "  from ofa_file,ofb_file,ogd_file,ima_file "
        oCommand.CommandText += "      ,ogb_file "
        oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "','" & t_ofa01_9 & "','" & t_ofa01_10 & "') "
        oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 and ofb04 = ima01 and ofa23 = 'USD' "
        oCommand.CommandText += "   and ogb01 = ofa011 and ogb03 = ofb03 "
        oCommand.CommandText += " order by ogd12b "
        LineZ = 11
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = "870829900"
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
                        Ws.Cells(LineZ, 2) = oReader2.Item("ta_obk15")
                    End While
                End If
                oReader2.Close()
                '200224 add by Brady
                'Ws.Cells(LineZ, 3) = oReader.Item("ogd13")
                'Ws.Cells(LineZ, 4) = oReader.Item("ofb13")
                'Ws.Cells(LineZ, 5) = oReader.Item("ofb14")
                'Ws.Cells(LineZ, 6) = oReader.Item("ofa01")
                Ws.Cells(LineZ, 3) = oReader.Item("imaud01")
                Ws.Cells(LineZ, 4) = oReader.Item("ogd13")
                Ws.Cells(LineZ, 5) = oReader.Item("ofb13")
                Ws.Cells(LineZ, 6) = oReader.Item("ofb14")
                Ws.Cells(LineZ, 7) = oReader.Item("ofa01")
                '200224 add by Brady END
                LineZ += 1
            End While
        End If
        oReader.Close()
        '200224 add by Brady
        ''191204 add by Brady
        ''Ws.Cells(100, 3) = "=SUM(C11:C" & LineZ - 1 & ")"
        ''Ws.Cells(100, 5) = "=SUM(E11:E" & LineZ - 1 & ")"
        'Ws.Cells(152, 3) = "=SUM(C11:C" & LineZ - 1 & ")"
        'Ws.Cells(152, 5) = "=SUM(E11:E" & LineZ - 1 & ")"
        ''191204 add by Brady END
        Ws.Cells(152, 4) = "=SUM(D11:D" & LineZ - 1 & ")"
        Ws.Cells(152, 6) = "=SUM(F11:F" & LineZ - 1 & ")"
        '200224 add by Brady END

        'INVOICE-EUR
        Ws = xWorkBook.Sheets(3)
        oCommand.CommandText = " select ima021,ofa02,NVL(ofaud05,' ') as ofaud05"
        oCommand.CommandText += "  from ofa_file,ofb_file,ogd_file,ima_file "
        oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "','" & t_ofa01_9 & "','" & t_ofa01_10 & "') "
        oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 and ofb04 = ima01 "
        oCommand.CommandText += " order by ogd12b "
        LineZ = 4
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                l_ima021 = oReader.Item("ima021")
                Ws.Cells(LineZ, 6) = l_ima021
                l_ofa02 = oReader.Item("ofa02")
                Ws.Cells(LineZ + 2, 6) = l_ofa02
                l_ofaud05 = oReader.Item("ofaud05")
                Ws.Cells(LineZ + 4, 6) = l_ofaud05
                Exit While
            End While
        End If
        oReader.Close()
        'oCommand.CommandText = " select ogd13,ogbud07 as ofb13,ogd13*ogbud07 as ofb14,ofa01,ofa04,ofb04,ofb11 "              '200224 mark by Brady
        oCommand.CommandText = " select ogd13,ogbud07 as ofb13,ogd13*ogbud07 as ofb14,ofa01,ofa04,ofb04,ofb11,imaud01 "       '200224 add by Brady
        oCommand.CommandText += "  from ofa_file,ofb_file,ogd_file,ima_file "
        oCommand.CommandText += "      ,ogb_file "
        oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "','" & t_ofa01_9 & "','" & t_ofa01_10 & "') "
        oCommand.CommandText += "   and ofa01 = ofb01 and ofa011 = ogd01 and ofb03 = ogd03 and ofb04 = ima01 and ofa23 = 'EUR' "
        oCommand.CommandText += "   and ogb01 = ofa011 and ogb03 = ofb03 "
        oCommand.CommandText += " order by ogd12b "
        LineZ = 11
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = "870829900"
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
                        Ws.Cells(LineZ, 2) = oReader2.Item("ta_obk15")
                    End While
                End If
                oReader2.Close()
                '200224 add by Brady
                'Ws.Cells(LineZ, 3) = oReader.Item("ogd13")
                'Ws.Cells(LineZ, 4) = oReader.Item("ofb13")
                'Ws.Cells(LineZ, 5) = oReader.Item("ofb14")
                'Ws.Cells(LineZ, 6) = oReader.Item("ofa01")
                Ws.Cells(LineZ, 3) = oReader.Item("imaud01")
                Ws.Cells(LineZ, 4) = oReader.Item("ogd13")
                Ws.Cells(LineZ, 5) = oReader.Item("ofb13")
                Ws.Cells(LineZ, 6) = oReader.Item("ofb14")
                Ws.Cells(LineZ, 7) = oReader.Item("ofa01")
                '200224 add by Brady END
                LineZ += 1
            End While
        End If
        oReader.Close()
        '200224 add by Brady
        ''191204 add by Brady
        ''Ws.Cells(100, 3) = "=SUM(C11:C" & LineZ - 1 & ")"
        ''Ws.Cells(100, 5) = "=SUM(E11:E" & LineZ - 1 & ")"
        'Ws.Cells(152, 3) = "=SUM(C11:C" & LineZ - 1 & ")"
        'Ws.Cells(152, 5) = "=SUM(E11:E" & LineZ - 1 & ")"
        ''191204 add by Brady END       
        Ws.Cells(152, 4) = "=SUM(D11:D" & LineZ - 1 & ")"
        Ws.Cells(152, 6) = "=SUM(F11:F" & LineZ - 1 & ")"
        '200224 add by Brady END

        '191030 add by Brady END

    End Sub

    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Trainfreight_SI"
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