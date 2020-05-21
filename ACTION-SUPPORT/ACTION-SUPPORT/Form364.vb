Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel

Public Class Form364
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
    Dim l_gate As String = String.Empty
    Dim l_ofa0451 As String = String.Empty
    Dim l_ofa0452 As String = String.Empty
    Dim l_ofa0453 As String = String.Empty
    Dim l_ofa04 As String = String.Empty
    Dim l_ogb04 As String = String.Empty
    Dim l_ogb11 As String = String.Empty
    Dim t_ogd15t As Double = 0
    Dim l_ta_obk15 As String = String.Empty


    Private Sub Form364_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'If Me.BackgroundWorker1.IsBusy() Then
        'MsgBox("处理中，请等待")
        'Return
        'End If        

        Dim xPath As String = "C:\temp\VDA_lable_sample.xlsx"
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
        Dim xPath As String = "C:\temp\VDA_lable_sample.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)

        l_gate = "DONGGUAN ACTION COMPOSITES LTD CO.No.8,Long Kou Road,Shi Gu Village,TangXia Town,523729, Dong Guan City,China "
        Ws = xWorkBook.Sheets(1)
        'oCommand.CommandText = "select ofa0451,ofa0452,ofa0453,ofa01,to_char(ofa02,'YYYY/MM/DD') as ofa02,ogb04,ogb11,NVL(ogb12,0) as ogb12,NVL(ogd12b,0) as ogd12b,ofa04 "
        oCommand.CommandText = "select NVL(ofa0451,' ') as ofa0451,NVL(ofa0452,' ') as ofa0452,NVL(ofa0453,' ') as ofa0453,ofa01,to_char(ofa02,'YYYY/MM/DD') as ofa02,NVL(ogb04,' ') as ogb04,NVL(ogb11,' ') as ogb11,NVL(ogd09,0) as ogd09,NVL(ogd12b,0) as ogd12b,NVL(ofa04,' ') as ofa04 "
        oCommand.CommandText += "  from ofa_file,ogb_file,ogd_file "
        oCommand.CommandText += " where ofa01 in ('" & t_ofa01_1 & "','" & t_ofa01_2 & "','" & t_ofa01_3 & "','" & t_ofa01_4 & "','" & t_ofa01_5 & "','" & t_ofa01_6 & "','" & t_ofa01_7 & "','" & t_ofa01_8 & "','" & t_ofa01_9 & "','" & t_ofa01_10 & "') "
        oCommand.CommandText += "   and ofa011 = ogb01 and ogd01 = ogb01 and ogd03 = ogb03 "
        oCommand.CommandText += " order by ofa01,ogd12b "
        LineZ = 2
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                l_ofa0451 = oReader.Item("ofa0451")
                l_ofa0452 = oReader.Item("ofa0452")
                l_ofa0453 = oReader.Item("ofa0453")
                Ws.Cells(LineZ, 1) = l_ofa0451 + "," + l_ofa0452 + "," + l_ofa0453
                Ws.Cells(LineZ, 2) = l_gate
                Ws.Cells(LineZ, 3) = oReader.Item("ofa01")
                Ws.Cells(LineZ, 4) = oReader.Item("ofa02")

                l_ofa04 = oReader.Item("ofa04")
                l_ogb04 = oReader.Item("ogb04")
                l_ogb11 = oReader.Item("ogb11")

                Ws.Cells(LineZ, 5) = l_ogb04
                Ws.Cells(LineZ, 6) = l_ogb11

                oCommand2.CommandText = " select NVL(ta_obk15,' ') as ta_obk15 from obk_file "
                oCommand2.CommandText += " where obk02 = '" & l_ofa04 & "' and obk01 = '" & l_ogb04 & "' and obk03 = '" & l_ogb11 & "'"
                'oCommand2.CommandText += "   and obkacti = 'Y' and rownum = 1 "
                oCommand2.CommandText += "   and obkacti = 'Y' "
                oReader2 = oCommand2.ExecuteReader()
                If oReader2.HasRows() Then
                    While oReader2.Read()
                        'l_ta_obk15 = oReader2.Item("ta_obk15")
                        Ws.Cells(LineZ, 7) = oReader2.Item("ta_obk15")
                        Exit While
                    End While
                End If
                oReader2.Close()

                'Ws.Cells(LineZ, 7) = l_ta_obk15
                Ws.Cells(LineZ, 8) = oReader.Item("ogd09")
                Ws.Cells(LineZ, 9) = oReader.Item("ogd12b")
                LineZ += 1
            End While
        End If
        Ws.Columns.EntireColumn.WrapText = False       '190910 add by Brady
        oReader.Close()
    End Sub

    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "VDA_lable"
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