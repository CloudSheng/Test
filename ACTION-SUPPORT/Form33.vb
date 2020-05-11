Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form33
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim Vi As Int16 = 0
    Dim Vj As Integer = 0
    Dim LineZ As Integer = 0
    Dim LineX As Integer = 0
    Dim DS As Data.DataSet = New DataSet()
    Dim DStartN As Date
    Dim CC As Integer = 0
    Dim PaNext As Integer = 0
    Dim DW1 As Integer = 0
    Dim DW2 As Integer = 0
    Dim DW3 As Integer = 0
    Dim DW4 As Integer = 0
    Dim SC As Decimal = 0
    Dim DRate As Decimal = 0
    Dim DDate As Date
    Dim SC2 As Decimal = 0
    Dim SD1 As Decimal = 0
    Dim SD2 As Decimal = 0
    Dim SD3 As Decimal = 0
    Dim SD4 As Decimal = 0
    Dim SD5 As Decimal = 0
    Dim SD6 As Decimal = 0
    Dim SD7 As Decimal = 0
    Dim SD8 As Decimal = 0
    Dim SD9 As Decimal = 0
    Dim SD10 As Decimal = 0
    Dim SD11 As Decimal = 0
    Dim SD12 As Decimal = 0
    Dim SD13 As Decimal = 0
    Dim SD14 As Decimal = 0
    Dim SD15 As Decimal = 0
    Dim SD16 As Decimal = 0
    Dim SD17 As Decimal = 0
    Dim SD18 As Decimal = 0
    Dim SD19 As Decimal = 0
    Dim SD20 As Decimal = 0
    Dim SD21 As Decimal = 0
    Dim SD22 As Decimal = 0
    Dim SD23 As Decimal = 0
    Dim SD24 As Decimal = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form33_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.GroupBox2.Enabled = False
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
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
        DStartN = Today()
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
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString = "SELECT 料件编号,品名,规格,发料单位, W1, W2, W3, W4, W5, W6, W7, W8, W9, W10, W11, W12, W13, W14, W15, W16, W17, W18, W19, W20, W21, W22, W23, W24, "
            ExcelString += "W25,W26,W27,W28,W29,W30,W31,W32,W33,W34,W35,W36,W37,W38,W39,W40,W41,W42,W43,W44,W45,W46,W47,W48,W49,W50,W51,W52,W53 FROM [数据$]"
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
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        'AdjustExcelFormat()
        Ws.Activate()
        AdjustExcelFormat()
        '先訂位
        DW1 = GetAzn02(DStartN)
        DW2 = GetAzn05(DStartN)
        ' 開始處理DS
        For Me.Vj = 0 To DS.Tables("tables1").Rows.Count - 1 Step 1
            Ws.Cells(LineZ, 1) = DS.Tables("tables1").Rows(Me.Vj).Item(0)
            SC = 0
            SC2 = 0
            SC = GetPrice(DS.Tables("tables1").Rows(Me.Vj).Item(0))
            Ws.Cells(LineZ, 2) = DS.Tables("tables1").Rows(Me.Vj).Item(1)
            Ws.Cells(LineZ, 3) = DS.Tables("tables1").Rows(Me.Vj).Item(2)
            Ws.Cells(LineZ, 4) = DS.Tables("tables1").Rows(Me.Vj).Item(3)
            For Me.Vi = 1 To 53 Step 1
                If IsDBNull(DS.Tables("tables1").Rows(Me.Vj).Item((3 + Me.Vi))) Then
                    Continue For
                End If
                If DS.Tables("tables1").Rows(Me.Vj).Item((3 + Me.Vi)) = 0 Then
                    Continue For
                End If
                ' 先取該週最後一天
                DDate = GetLastDay(DW1, Me.Vi)
                ' 再取該週的匯率
                DRate = GetRate(DDate)
                If Me.Vi - 2 <= DW2 Then  '-2週後 小於等於啟始週數的, 要加總
                    SC2 += DS.Tables("tables1").Rows(Me.Vj).Item((3 + Me.Vi)) * SC / DRate
                    Ws.Cells(LineZ, 5) = SC2
                Else
                    Ws.Cells(LineZ, 5 - 2 + Me.Vi - DW2) = DS.Tables("tables1").Rows(Me.Vj).Item((3 + Me.Vi)) * SC / DRate
                End If
            Next
            LineZ += 1
            Label3.Text = LineZ
        Next
        ' 加總
        Ws.Cells(LineZ, 1) = "合计"
        Ws.Cells(LineZ, 5) = "=SUM(E2:E" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 5), Ws.Cells(LineZ, 5))
        oRng.AutoFill(Destination:=Ws.Range("E" & LineZ & ":DB" & LineZ), Type:=xlFillDefault)

        ' 處理第二頁, 月資料
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustExcelFormat1()
        '先訂位
        DW1 = DStartN.Year
        DW2 = DStartN.Month
        ' 開始處理 DS
        For Me.Vj = 0 To DS.Tables("tables1").Rows.Count - 1 Step 1
            Ws.Cells(LineZ, 1) = DS.Tables("tables1").Rows(Me.Vj).Item(0)
            SC = 0
            SC2 = 0
            CleanSD() '月變數歸0
            SC = GetPrice(DS.Tables("tables1").Rows(Me.Vj).Item(0))
            Ws.Cells(LineZ, 2) = DS.Tables("tables1").Rows(Me.Vj).Item(1)
            Ws.Cells(LineZ, 3) = DS.Tables("tables1").Rows(Me.Vj).Item(2)
            Ws.Cells(LineZ, 4) = DS.Tables("tables1").Rows(Me.Vj).Item(3)
            For Me.Vi = 1 To 53 Step 1
                If IsDBNull(DS.Tables("tables1").Rows(Me.Vj).Item((3 + Me.Vi))) Then
                    Continue For
                End If
                If DS.Tables("tables1").Rows(Me.Vj).Item((3 + Me.Vi)) = 0 Then
                    Continue For
                End If
                ' 先取該週最後一天
                If Me.Vi > 2 Then
                    DDate = GetLastDay(DW1, Me.Vi - 2)
                Else
                    DDate = GetLastDay(DW1, 1)
                End If
                ' 再取該週的匯率
                DRate = GetRate(DDate)
                ' 取年月
                DW3 = DDate.Year
                DW4 = DDate.Month
                If DW4 <= DW2 Then  '小於等於啟始週數的, 要加總
                    SC2 += DS.Tables("tables1").Rows(Me.Vj).Item((3 + Me.Vi)) * SC / DRate
                    Ws.Cells(LineZ, 5) = SC2
                Else
                    Select Case DW4
                        Case 1
                            SD1 += DS.Tables("tables1").Rows(Me.Vj).Item((3 + Me.Vi)) * SC / DRate
                            Ws.Cells(LineZ, 5 + DW4 - DW2) = SD1
                        Case 2
                            SD2 += DS.Tables("tables1").Rows(Me.Vj).Item((3 + Me.Vi)) * SC / DRate
                            Ws.Cells(LineZ, 5 + DW4 - DW2) = SD2
                        Case 3
                            SD3 += DS.Tables("tables1").Rows(Me.Vj).Item((3 + Me.Vi)) * SC / DRate
                            Ws.Cells(LineZ, 5 + DW4 - DW2) = SD3
                        Case 4
                            SD4 += DS.Tables("tables1").Rows(Me.Vj).Item((3 + Me.Vi)) * SC / DRate
                            Ws.Cells(LineZ, 5 + DW4 - DW2) = SD4
                        Case 5
                            SD5 += DS.Tables("tables1").Rows(Me.Vj).Item((3 + Me.Vi)) * SC / DRate
                            Ws.Cells(LineZ, 5 + DW4 - DW2) = SD5
                        Case 6
                            SD6 += DS.Tables("tables1").Rows(Me.Vj).Item((3 + Me.Vi)) * SC / DRate
                            Ws.Cells(LineZ, 5 + DW4 - DW2) = SD6
                        Case 7
                            SD7 += DS.Tables("tables1").Rows(Me.Vj).Item((3 + Me.Vi)) * SC / DRate
                            Ws.Cells(LineZ, 5 + DW4 - DW2) = SD7
                        Case 8
                            SD8 += DS.Tables("tables1").Rows(Me.Vj).Item((3 + Me.Vi)) * SC / DRate
                            Ws.Cells(LineZ, 5 + DW4 - DW2) = SD8
                        Case 9
                            SD9 += DS.Tables("tables1").Rows(Me.Vj).Item((3 + Me.Vi)) * SC / DRate
                            Ws.Cells(LineZ, 5 + DW4 - DW2) = SD9
                        Case 10
                            SD10 += DS.Tables("tables1").Rows(Me.Vj).Item((3 + Me.Vi)) * SC / DRate
                            Ws.Cells(LineZ, 5 + DW4 - DW2) = SD10
                        Case 11
                            SD11 += DS.Tables("tables1").Rows(Me.Vj).Item((3 + Me.Vi)) * SC / DRate
                            Ws.Cells(LineZ, 5 + DW4 - DW2) = SD11
                        Case 12
                            SD12 += DS.Tables("tables1").Rows(Me.Vj).Item((3 + Me.Vi)) * SC / DRate
                            Ws.Cells(LineZ, 5 + DW4 - DW2) = SD12
                    End Select
                End If
            Next
            LineZ += 1
            Label3.Text = LineZ + PaNext
        Next
        ' 加總
        Ws.Cells(LineZ, 1) = "合计"
        Ws.Cells(LineZ, 5) = "=SUM(E2:E" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 5), Ws.Cells(LineZ, 5))
        oRng.AutoFill(Destination:=Ws.Range("E" & LineZ & ":DB" & LineZ), Type:=xlFillDefault)
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "AP_Simulation"
        SaveFileDialog1.DefaultExt = ".xls"
        Dim SON As DialogResult = SaveFileDialog1.ShowDialog()
        If SON = DialogResult.OK Then
            Dim SFN As String = SaveFileDialog1.FileName
            Ws.SaveAs(SFN, XlFileFormat.xlExcel12)
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
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "周数据"
        Ws.Columns.EntireColumn.ColumnWidth = 20
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Cells(1, 1) = "料件编号"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(1, 2) = "品名"
        Ws.Cells(1, 3) = "规格"
        Ws.Cells(1, 4) = "发料单位"
        oCommand.CommandText = "SELECT distinct azn02,azn05 FROM AZN_FILE WHERE AZN01 >= TO_DATE('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') order by azn02,azn05"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            CC = 0
            While oReader.Read()
                Ws.Cells(1, 5 + CC) = oReader.Item("azn02") & "W" & oReader.Item("azn05")
                CC += 1
            End While
        End If
        oReader.Close()
        LineZ = 2
    End Sub
    Private Function GetAzn02(ByVal eDate As Date)
        Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander2.Connection = oConnection
        oCommander2.CommandType = CommandType.Text
        oCommander2.CommandText = "SELECT azn02 FROM azn_file where azn01 = TO_DATE('" & eDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        Dim ADW1 As Integer = oCommander2.ExecuteScalar()
        Return ADW1
    End Function
    Private Function GetAzn05(ByVal eDate As Date)
        Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander2.Connection = oConnection
        oCommander2.CommandType = CommandType.Text
        oCommander2.CommandText = "SELECT azn05 FROM azn_file where azn01 = TO_DATE('" & eDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        Dim ADW2 As Integer = oCommander2.ExecuteScalar()
        Return ADW2
    End Function
    Private Function GetPrice(ByVal ima01 As String)
        Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander2.Connection = oConnection
        oCommander2.CommandType = CommandType.Text
        oCommander2.CommandText = "select ima53 from ima_file where ima01 = '" & ima01 & "'"
        Dim ADW3 As Integer = oCommander2.ExecuteScalar()
        Return ADW3
    End Function
    Private Function GetRate(ByVal eDate As Date)
        Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander2.Connection = oConnection
        oCommander2.CommandType = CommandType.Text
        Dim MM As String = String.Empty
        MM = eDate.Month
        If Strings.Len(MM) = 1 Then
            MM = "0" & MM
        End If
        MM = eDate.Year & MM
        oCommander2.CommandText = "select azj04 from azj_file where azj01 = 'USD' AND azj02 = '" & MM & "'"
        Dim MK As Integer = oCommander2.ExecuteScalar()
        If IsDBNull(MK) Or MK = 0 Then
            Dim SX As String = String.Empty
            SX = DStartN.Month()
            If Strings.Len(SX) = 1 Then
                SX = "0" & SX
            End If
            SX = DStartN.Year & SX
            oCommander2.CommandText = "select azj04 from azj_file where azj01 = 'USD' AND azj02 = '" & SX & "'"
            MK = oCommander2.ExecuteScalar()
            If IsDBNull(MK) Or MK = 0 Then
                MK = 1
            End If
        End If
        Return MK
    End Function
    Private Function GetLastDay(ByVal azn02 As Integer, ByVal azn05 As Integer)
        Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander2.Connection = oConnection
        oCommander2.CommandType = CommandType.Text
        oCommander2.CommandText = "SELECT max(azn01) FROM azn_file where azn02 = " & azn02 & " and azn05 = " & azn05
        Dim ADW4 As Date = oCommander2.ExecuteScalar()
        Return ADW4
    End Function
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "月数据"
        Ws.Columns.EntireColumn.ColumnWidth = 20
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Cells(1, 1) = "料件编号"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(1, 2) = "品名"
        Ws.Cells(1, 3) = "规格"
        Ws.Cells(1, 4) = "发料单位"
        oCommand.CommandText = "SELECT distinct azn02,azn04 FROM AZN_FILE WHERE AZN01 >= TO_DATE('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') order by azn02,azn04"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            CC = 0
            While oReader.Read()
                Dim MX As String = oReader.Item("azn04")
                If Strings.Len(MX) = 1 Then
                    MX = "0" & MX
                End If
                Ws.Cells(1, 5 + CC) = oReader.Item("azn02") & MX
                CC += 1
            End While
        End If
        oReader.Close()
        PaNext = LineZ
        LineZ = 2
    End Sub
    Private Sub CleanSD()
        SD1 = 0
        SD2 = 0
        SD3 = 0
        SD4 = 0
        SD5 = 0
        SD6 = 0
        SD7 = 0
        SD8 = 0
        SD9 = 0
        SD10 = 0
        SD11 = 0
        SD12 = 0
        SD13 = 0
        SD14 = 0
        SD15 = 0
        SD16 = 0
        SD17 = 0
        SD18 = 0
        SD19 = 0
        SD20 = 0
        SD21 = 0
        SD22 = 0
        SD23 = 0
        SD24 = 0
    End Sub

End Class