Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form21
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim Vi As Int16 = 0
    Dim Vj As Integer = 0
    Dim LineZ As Integer = 0
    Dim LineX As Integer = 0
    Dim DS As Data.DataSet = New DataSet()
    Dim ArrayS1 As String() = {"102010010002", "102010010027", "102010010037", "102010010022", "102010010034", _
                               "102010020012", "102010010012", "102020020009", "102020020028", "102020020014", _
                               "102020020016", "102010020002", "204000020011", "205000010001", "205000010003", _
                               "206010010002", "206030020005", "206020010002", "206020010001", "102010010023", _
                               "102010010009", "102020020008", "102010010007", "102010010011", "102010010013", _
                               "102010010019", "102020020022", "102010010039", "102010010040", "102010010041", _
                               "203010020024", "203010020016", "203010020027", "203010020002", "203010020013", _
                               "203010020014", "203010020007", "203010020008", "203010020015", "203010020025", _
                               "203010020018", "102010010070", "102010010069", "102010010056", "102010010062", _
                               "102010010063", "102010010047", "102010010048", "102010010043", "102020020003", _
                               "102010010060", "203010010031", "203010020030", "102010010046", "102010010073", _
                               "102010010074"}
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form21_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.GroupBox2.Enabled = False
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString = "SELECT ERPPN,W1,W2,W3,W4,W5,W6,W7,W8,W9,W10,W11,W12,W13,W14,W15,W16,W17,W18,W19,W20,W21,W22,W23,W24,"
            ExcelString += "W25,W26,W27,W28,W29,W30,W31,W32,W33,W34,W35,W36,W37,W38,W39,W40,W41,W42,W43,W44,W45,W46,W47,W48,W49,W50,W51,W52,W53 FROM [sheet1$]"
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)
            Try
                ExcelAdapater.Fill(DS, "tables1")
            Catch ex As Exception
                MsgBox(ex.Message())
                Label1.Text = "读取失败"
                Me.GroupBox2.Enabled = False
                Return
            End Try
            Label1.Text = "已读入"
            Me.GroupBox2.Enabled = True
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If DS.Tables("tables1").Rows.Count = 0 Then
            MsgBox("无资料可处理")
            Return
        End If
        Me.Label3.Text = 0
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommander2.Connection = oConnection
                oCommander2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        DropTempTable()
        oCommand.CommandText = "CREATE TABLE aaaa_temp(bmb01 varchar2(40),bmb03 varchar2(40),bmb06 number(16,8),bmb10 varchar2(4),month1 number(5))"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
            Return
        End Try
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub DropTempTable()
        oCommand.CommandText = "DROP TABLE aaaa_temp"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub ExpandBOM(ByVal ERPPN As String, ByVal Q1 As Decimal)
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "SELECT bmb01,bmb03,Round(bmb06/bmb07*(1+bmb08/100),8) as t1,bmb10 FROM bmb_file,bma_file,ima_file where bmb01 = bma01 and bma10 = 2 and bmb01 = '"
        oCommander99.CommandText += ERPPN & "' and bmb05 is null and bma06 = ima910 and bmb29 = bma06 and bmb01 = ima01"
        Dim oReader99 As Oracle.ManagedDataAccess.Client.OracleDataReader = oCommander99.ExecuteReader()
        If oReader99.HasRows() Then
            While oReader99.Read()
                Dim T1 As Double = oReader99.Item("t1") * Q1
                If ArrayS1.Contains(oReader99.Item("bmb03").ToString()) Then
                    ' Insert into aaaa_temp
                    InsertIntoTempDB(oReader99.Item("bmb01"), oReader99.Item("bmb03"), oReader99.Item("bmb10"), T1)
                    Continue While
                Else
                    Call ExpandBOM(oReader99.Item("bmb03"), T1)
                End If
            End While
        End If
        oReader99.Close()
    End Sub
    Private Sub InsertIntoTempDB(ByVal ERPPN1 As String, ByVal ERPPN2 As String, ByVal Unit As String, ByVal Q2 As Decimal)
        Dim oCommander98 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander98.Connection = oConnection
        oCommander98.CommandType = CommandType.Text
        oCommander98.CommandText = "INSERT INTO aaaa_temp VALUES ('" & ERPPN1 & "','" & ERPPN2 & "',"
        oCommander98.CommandText += Q2 & ",'" & Unit & "'," & Me.Vi & ")"

        Try
            oCommander98.ExecuteNonQuery()
            Me.Label3.Text += 1
        Catch ex As Exception
            MsgBox(ex.Message())
            Return
        End Try
    End Sub
    Private Sub InsertTemp()
        For Me.Vi = 1 To 53 Step 1
            For Me.Vj = 0 To DS.Tables("tables1").Rows.Count - 1 Step 1
                If Not String.IsNullOrEmpty(DS.Tables("tables1").Rows(Me.Vj).Item(Me.Vi).ToString()) Then
                    If DS.Tables("tables1").Rows(Me.Vj).Item(Me.Vi) <> 0 Then
                        ExpandBOM(DS.Tables("tables1").Rows(Me.Vj).Item(0), DS.Tables("tables1").Rows(Me.Vj).Item(Me.Vi))
                    End If
                End If
            Next
        Next
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        InsertTemp()
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        oCommand.CommandText = "SELECT count(*) FROM aaaa_temp"
        Dim THAS As Integer = oCommand.ExecuteScalar()
        If THAS > 0 Then
            xExcel = New Microsoft.Office.Interop.Excel.Application
            xWorkBook = xExcel.Workbooks.Add()
            Ws = xWorkBook.Sheets(1)
            'AdjustExcelFormat()
            Ws.Activate()
            AdjustExcelFormat()
            oCommand.CommandText = "SELECT bmb03,ima02,ima021,month1,bmb10,sum(bmb06) as c1 from aaaa_temp, ima_file where bmb03 = ima01 "
            oCommand.CommandText += "group by bmb03,ima02,ima021,bmb10,month1 order by bmb03"
            oReader = oCommand.ExecuteReader()
            If oReader.HasRows() Then
                Dim DChar As String = String.Empty
                While oReader.Read()
                    If String.IsNullOrEmpty(DChar) Then
                        DChar = oReader.Item("bmb03")
                        Ws.Cells(LineX, 1) = "'" & DChar
                        Ws.Cells(LineX, 2) = oReader.Item("ima02")
                        Ws.Cells(LineX, 3) = oReader.Item("ima021")
                        Ws.Cells(LineX, 4) = oReader.Item("bmb10")
                    End If
                    If oReader.Item("bmb03") <> DChar Then
                        DChar = oReader.Item("bmb03")
                        LineX += 1
                        Ws.Cells(LineX, 1) = "'" & DChar
                        Ws.Cells(LineX, 2) = oReader.Item("ima02")
                        Ws.Cells(LineX, 3) = oReader.Item("ima021")
                        Ws.Cells(LineX, 4) = oReader.Item("bmb10")
                    End If
                    Ws.Cells(LineX, 4 + oReader.Item("month1")) = oReader.Item("c1")
                End While
            End If
            oReader.Close()
        End If
        ' 20151126 ADD
        LineX += 1
        For Me.Vj = 0 To DS.Tables("tables1").Rows.Count - 1 Step 1
            oCommand.CommandText = "select count(*) from bma_file where bma10 = 2 and bma01 = '" & DS.Tables("tables1").Rows(Me.Vj).Item(0) & "'"
            Dim HASBOM As Int16 = oCommand.ExecuteScalar()
            If HASBOM = 0 Then
                Ws.Cells(LineX, 1) = "'" & DS.Tables("tables1").Rows(Me.Vj).Item(0)
                Ws.Cells(LineX, 2) = "NO BOM"
                LineX += 1
            End If
        Next
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Simulation_MRP"
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
                DropTempTable()
                oConnection.Close()
                Module1.KillExcelProcess(OldExcel)
                MsgBox("Finished")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "数据"
        Ws.Columns.EntireColumn.ColumnWidth = 20
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Cells(1, 1) = "料件编号"
        Ws.Cells(1, 2) = "品名"
        Ws.Cells(1, 3) = "规格"
        Ws.Cells(1, 4) = "发料单位"
        Ws.Cells(1, 5) = "W1"
        Ws.Cells(1, 6) = "W2"
        Ws.Cells(1, 7) = "W3"
        Ws.Cells(1, 8) = "W4"
        Ws.Cells(1, 9) = "W5"
        Ws.Cells(1, 10) = "W6"
        Ws.Cells(1, 11) = "W7"
        Ws.Cells(1, 12) = "W8"
        Ws.Cells(1, 13) = "W9"
        Ws.Cells(1, 14) = "W10"
        Ws.Cells(1, 15) = "W11"
        Ws.Cells(1, 16) = "W12"
        Ws.Cells(1, 17) = "W13"
        Ws.Cells(1, 18) = "W14"
        Ws.Cells(1, 19) = "W15"
        Ws.Cells(1, 20) = "W16"
        Ws.Cells(1, 21) = "W17"
        Ws.Cells(1, 22) = "W18"
        Ws.Cells(1, 23) = "W19"
        Ws.Cells(1, 24) = "W20"
        Ws.Cells(1, 25) = "W21"
        Ws.Cells(1, 26) = "W22"
        Ws.Cells(1, 27) = "W23"
        Ws.Cells(1, 28) = "W24"
        Ws.Cells(1, 29) = "W25"
        Ws.Cells(1, 30) = "W26"
        Ws.Cells(1, 31) = "W27"
        Ws.Cells(1, 32) = "W28"
        Ws.Cells(1, 33) = "W29"
        Ws.Cells(1, 34) = "W30"
        Ws.Cells(1, 35) = "W31"
        Ws.Cells(1, 36) = "W32"
        Ws.Cells(1, 37) = "W33"
        Ws.Cells(1, 38) = "W34"
        Ws.Cells(1, 39) = "W35"
        Ws.Cells(1, 40) = "W36"
        Ws.Cells(1, 41) = "W37"
        Ws.Cells(1, 42) = "W38"
        Ws.Cells(1, 43) = "W39"
        Ws.Cells(1, 44) = "W40"
        Ws.Cells(1, 45) = "W41"
        Ws.Cells(1, 46) = "W42"
        Ws.Cells(1, 47) = "W43"
        Ws.Cells(1, 48) = "W44"
        Ws.Cells(1, 49) = "W45"
        Ws.Cells(1, 50) = "W46"
        Ws.Cells(1, 51) = "W47"
        Ws.Cells(1, 52) = "W48"
        Ws.Cells(1, 53) = "W49"
        Ws.Cells(1, 54) = "W50"
        Ws.Cells(1, 55) = "W51"
        Ws.Cells(1, 56) = "W52"
        Ws.Cells(1, 57) = "W53"
        LineX = 2
    End Sub
End Class