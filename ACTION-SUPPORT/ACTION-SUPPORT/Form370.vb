Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel

Public Class Form370
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
    Dim LineZ As Integer = 0
    Dim DBC As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Dim SaveFileDialog1 As New SaveFileDialog
    Dim t_oga01_1 As String = String.Empty
    Dim l_ogb05 As String = String.Empty
    Dim l_ged01 As String = String.Empty
    Dim l_ogb04 As String = String.Empty


    Private Sub Form370_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'If Me.BackgroundWorker1.IsBusy() Then
        'MsgBox("处理中，请等待")
        'Return
        'End If        

        Dim xPath As String = "C:\temp\EDI_Template.xlsx"
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
            t_oga01_1 = TextBox1.Text
        End If

        'xExcel = New Microsoft.Office.Interop.Excel.Application
        'xWorkBook = xExcel.Workbooks.Add()
        ExportToExcel()
        oConnection.Close()

        SaveExcel()
    End Sub

    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        Dim xPath As String = "C:\temp\EDI_Template.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)

        Ws = xWorkBook.Sheets(1)
        'oCommand2.CommandText = "select count(*) from ogb_file "
        'oCommand2.CommandText += " where ogb01 in ('" & t_oga01_1 & "') "
        'Dim ogb_cnt As Decimal = oCommand2.ExecuteScalar()

        ''oCommand.CommandText = "select NVL(ofa0451,' ') as ofa0451,NVL(ofa0452,' ') as ofa0452,NVL(ofa0453,' ') as ofa0453,ofa01,to_char(ofa02,'YYYY/MM/DD') as ofa02,NVL(ogb04,' ') as ogb04,NVL(ogb11,' ') as ogb11,NVL(ogd09,0) as ogd09,NVL(ogd12b,0) as ogd12b,NVL(ofa04,' ') as ofa04 "
        'oCommand.CommandText = " select OGA00,	OGA01,	OGA011,	OGA02,	OGA021,	OGA022,	OGA03,	OGA032,	OGA033,	OGA04,	OGA044,	OGA05,	OGA06,	OGA07,	OGA08,	OGA09,	"
        'oCommand.CommandText += "   OGA10,	OGA11,	OGA12,	OGA13,	OGA14,	OGA15,	OGA16,	OGA161,	OGA162,	OGA163,	OGA17,	OGA18,	OGA19,	OGA20,	OGA21,	OGA211,	OGA212, "
        'oCommand.CommandText += "	OGA213,	OGA23,	OGA24,	OGA25,	OGA26,	OGA27,	OGA28,	OGA29,	OGA30,	OGA31,	OGA32,	OGA33,	OGA34,	OGA35,	OGA36,	OGA37,	OGA38, "
        'oCommand.CommandText += "   OGA39,	OGA40,	OGA41,	OGA42,	OGA43,	OGA44,	OGA45,	OGA46,	OGA47,	OGA48,	OGA49,	OGA50,	OGA501,	OGA51,	OGA511,	OGA52,	OGA53, "
        'oCommand.CommandText += "   OGA54,	OGA99,	OGA901,	OGA902,	OGA903,	OGA904,	OGA905,	OGA906,	OGA907,	OGA908,	OGA909,	OGA910,	OGA911,	OGACONF,	OGAPOST,	OGAPRSW, "
        'oCommand.CommandText += "  	OGAUSER,	OGAGRUP,	OGAMODU,	OGADATE,	OGA55,	OGAMKSG,	OGA65,	OGA66,	OGA67,	OGA1001,	OGA1002,	OGA1003,	OGA1004, "
        'oCommand.CommandText += "   OGA1005,	OGA1006,	OGA1007,	OGA1008,	OGA1009,	OGA1010,	OGA1011,	OGA1012,	OGA1013,	OGA1014,	OGA1015,	OGA1016, "
        'oCommand.CommandText += "   OGA68,	OGASPC,	OGA69,	OGA912,	OGA913,	OGA914,	OGA70,	OGAUD01,	OGAUD02,	OGAUD03,	OGAUD04,	OGAUD05,	OGAUD06,	OGAUD07, "
        'oCommand.CommandText += "   OGAUD08,	OGAUD09,	OGAUD10,	OGAUD11,	OGAUD12,	OGAUD13,	OGAUD14,	OGAUD15,	OGA83,	OGA84,	OGA85,	OGA86,	OGA87, "
        'oCommand.CommandText += "   OGA88,	OGA89,	OGA90,	OGA91,	OGA92,	OGA93,	OGA94,	OGA95,	OGA96,	OGA97,	OGACOND,	OGACONU,	OGAPLANT,	OGALEGAL,	OGA71, "
        'oCommand.CommandText += " 	OGAORIU,	OGAORIG,	OGACONT,	OGA98,	OGA72,	OGA57 "
        'oCommand.CommandText += "  from oga_file "
        'oCommand.CommandText += " where oga01 in ('" & t_oga01_1 & "') "
        'LineZ = 2
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        For z As Int16 = 1 To ogb_cnt Step 1
        '            For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
        '                Ws.Cells(LineZ, i + 1) = oReader.Item(i)
        '            Next
        '            LineZ += 1
        '        Next
        '    End While
        'End If
        'Ws.Columns.EntireColumn.WrapText = False
        'oReader.Close()

        'oCommand.CommandText = " select OGB01,	OGB03,	OGB04,	OGB05,	OGB05_FAC,	OGB06,	OGB07,	OGB08,	OGB09,	OGB091,	OGB092,	OGB11,	OGB12,	OGB13,	OGB14,	OGB14T,	OGB15, "
        'oCommand.CommandText += "   OGB15_FAC,	OGB16,	OGB17,	OGB18,	OGB19,	OGB20,	OGB21,	OGB22,	OGB31,	OGB32,	OGB60,	OGB63,	OGB64,	OGB901,	OGB902,	OGB903,	OGB904,	OGB905, "
        'oCommand.CommandText += "   OGB906,	OGB907,	OGB908,	OGB909,	OGB910,	OGB911,	OGB912,	OGB913,	OGB914,	OGB915,	OGB916,	OGB917,	OGB65,	OGB1001,	OGB1002,	OGB1005,	OGB1007, "
        'oCommand.CommandText += "   OGB1008,	OGB1009,	OGB1010,	OGB1011,	OGB1003,	OGB1004,	OGB1006,	OGB1012,	OGB930,	OGB1013,	OGB1014,	OGB41,	OGB42,	OGB43, "
        'oCommand.CommandText += "   OGB931,	OGB932,	OGBUD01,	OGBUD02,	OGBUD03,	OGBUD04,	OGBUD05,	OGBUD06,	OGBUD07,	OGBUD08,	OGBUD09,	OGBUD10,	OGBUD11,	OGBUD12, "
        'oCommand.CommandText += "   OGBUD13,	OGBUD14,	OGBUD15,	OGB44,	OGB45,	OGB46,	OGB47,	OGBPLANT,	OGBLEGAL,	OGB48,	OGB49,	OGB37,	OGB40 "
        'oCommand.CommandText += "  from ogb_file "
        'oCommand.CommandText += " where ogb01 in ('" & t_oga01_1 & "') "
        'LineZ = 2
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
        '            Ws.Cells(LineZ, i + 157) = oReader.Item(i)
        '        Next
        '        LineZ += 1
        '    End While
        'End If
        'Ws.Columns.EntireColumn.WrapText = False
        'oReader.Close()

        oCommand.CommandText = " select unique ogd03, ogd15t,	ogd14t,	ogd10,substr(ogb31,3,12) as ogb31,	ogb11,	ogb04,	ogb12,ogb11, ogb04,	ogb12,	ogb03,	ogb092,	ogbud13,	ogd13,	ogd12b,	ogd12e, NVL(ogb05,' ') as ogb05 "
        oCommand.CommandText += "  from ogd_file,ogb_file "
        oCommand.CommandText += " where ogd01 in ('" & t_oga01_1 & "') and ogd01 = ogb01 and ogd03 = ogb03 "
        oCommand.CommandText += "  order by ogd03 "
        LineZ = 2
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 5) = oReader.Item("ogd15t")
                Ws.Cells(LineZ, 6) = oReader.Item("ogd14t")
                Ws.Cells(LineZ, 7) = oReader.Item("ogd10")
                Ws.Cells(LineZ, 11) = oReader.Item("ogb31")
                Ws.Cells(LineZ, 12) = oReader.Item("ogb11")
                l_ogb04 = oReader.Item("ogb04")
                If l_ogb04 = "126AA0609001066" Then
                    l_ogb04 = "AA0609"
                End If
                Ws.Cells(LineZ, 13) = l_ogb04
                Ws.Cells(LineZ, 14) = oReader.Item("ogb12")
                Ws.Cells(LineZ, 16) = oReader.Item("ogb03")
                Ws.Cells(LineZ, 17) = oReader.Item("ogb092")
                Ws.Cells(LineZ, 18) = oReader.Item("ogbud13")
                Ws.Cells(LineZ, 19) = oReader.Item("ogb03")
                Ws.Cells(LineZ, 20) = oReader.Item("ogd13")
                Ws.Cells(LineZ, 21) = oReader.Item("ogd12b")
                Ws.Cells(LineZ, 22) = oReader.Item("ogd12e")
                l_ogb05 = oReader.Item("ogb05")
                Select Case l_ogb05
                    Case "KG"
                        Ws.Cells(LineZ, 15) = "KG"
                    Case "L"
                        Ws.Cells(LineZ, 15) = "L"
                    Case "M"
                        Ws.Cells(LineZ, 15) = "M"
                    Case "M²"
                        Ws.Cells(LineZ, 15) = "M2"
                    Case "SET"
                        Ws.Cells(LineZ, 15) = "SA"
                    Case "PCS"
                        Ws.Cells(LineZ, 15) = "ST"
                End Select
                LineZ += 1
            End While
        End If
        Ws.Columns.EntireColumn.WrapText = False
        oReader.Close()

        oCommand2.CommandText = "select count(*) from ogd_file "
        oCommand2.CommandText += " where ogd01 in ('" & t_oga01_1 & "') "
        Dim ogb_cnt As Decimal = oCommand2.ExecuteScalar()

        oCommand.CommandText = " select oga69, substr(ofa47,1,8) as ofa47, substr(ged02,1,14) as ged02, oga02, substr(oga01,1,8) as oga01, NVL(ged01,' ') as ged01	"
        oCommand.CommandText += "  from oga_file,ofa_file left join ged_file on ofa43 = ged01 "
        oCommand.CommandText += " where oga01 in ('" & t_oga01_1 & "') and oga01 = ofa011 "
        LineZ = 2
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For z As Int16 = 1 To ogb_cnt Step 1
                    Ws.Cells(LineZ, 1) = oReader.Item("oga69")
                    Ws.Cells(LineZ, 2) = oReader.Item("ofa47")
                    Ws.Cells(LineZ, 3) = oReader.Item("ged02")
                    Ws.Cells(LineZ, 4) = oReader.Item("oga02")
                    Ws.Cells(LineZ, 8) = oReader.Item("oga01")
                    Ws.Cells(LineZ, 9) = oReader.Item("oga02")
                    l_ged01 = oReader.Item("ged01")
                    Select Case l_ged01
                        Case "06"
                            Ws.Cells(LineZ, 10) = "8"
                        Case "07"
                            Ws.Cells(LineZ, 10) = "8"
                        Case "24"
                            Ws.Cells(LineZ, 10) = "8"
                        Case "01"
                            Ws.Cells(LineZ, 10) = "10"
                        Case "02"
                            Ws.Cells(LineZ, 10) = "10"
                        Case "03"
                            Ws.Cells(LineZ, 10) = "10"
                        Case "04"
                            Ws.Cells(LineZ, 10) = "10"
                        Case "05"
                            Ws.Cells(LineZ, 10) = "10"
                        Case "09"
                            Ws.Cells(LineZ, 10) = "10"
                        Case "10"
                            Ws.Cells(LineZ, 10) = "10"
                        Case "12"
                            Ws.Cells(LineZ, 10) = "10"
                        Case "18"
                            Ws.Cells(LineZ, 10) = "10"
                        Case "20"
                            Ws.Cells(LineZ, 10) = "10"
                        Case "26"
                            Ws.Cells(LineZ, 10) = "10"
                        Case "27"
                            Ws.Cells(LineZ, 10) = "10"
                        Case "28"
                            Ws.Cells(LineZ, 10) = "10"
                        Case "8"
                            Ws.Cells(LineZ, 10) = "11"
                        Case "13"
                            Ws.Cells(LineZ, 10) = "11"
                        Case "14"
                            Ws.Cells(LineZ, 10) = "11"
                        Case "15"
                            Ws.Cells(LineZ, 10) = "11"
                        Case "16"
                            Ws.Cells(LineZ, 10) = "11"
                        Case "17"
                            Ws.Cells(LineZ, 10) = "11"
                        Case "19"
                            Ws.Cells(LineZ, 10) = "11"
                        Case "21"
                            Ws.Cells(LineZ, 10) = "11"
                        Case "22"
                            Ws.Cells(LineZ, 10) = "11"
                        Case "23"
                            Ws.Cells(LineZ, 10) = "11"
                        Case "25"
                            Ws.Cells(LineZ, 10) = "11"
                        Case "11"
                            Ws.Cells(LineZ, 10) = "20"
                    End Select
                    LineZ += 1
                Next
            End While
        End If
        Ws.Columns.EntireColumn.WrapText = False
        oReader.Close()

    End Sub

    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "EDI_shipping"
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