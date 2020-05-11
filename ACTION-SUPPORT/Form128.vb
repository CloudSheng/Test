Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form128
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
    Dim LineS1 As Int16 = 0
    Dim tDate As Date
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim tWeek As Int16 = 0
    Dim MaxWeek As Int16 = 0
    Dim ColumnIndex As Int16 = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
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
        tDate = DateTimePicker1.Value
        tYear = DateTimePicker1.Value.Year
        tMonth = DateTimePicker1.Value.Month
        oCommand.CommandText = "SELECT azn05 FROM azn_file WHERE azn01 = to_date('" & DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        tWeek = oCommand.ExecuteScalar()
        oCommand.CommandText = "select max(azn05) from azn_file where azn02 = " & tYear
        MaxWeek = oCommand.ExecuteScalar()
        BackgroundWorker1.RunWorkerAsync()
        'ExportToExcel()
        'SaveExcel()
    End Sub

    Private Sub Form128_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        AdjustExcelFormat1()
        Ws.Cells(2, 2) = Getaah("100101", "100211")
        Ws.Cells(2, 4) = "=B2"
        Ws.Cells(2, 5) = "=D9"
        oRng = Ws.Range("E2", "E2")
        oRng.AutoFill(Destination:=Ws.Range("E2", Ws.Cells(2, 4 + ColumnIndex)), Type:=xlFillDefault)
        Ws.Cells(3, 3) = GetAR()
        Ws.Cells(8, 3) = GetAP()

        oCommand.CommandText = "select azn05,sum(t1) as t1 from ( select azn05,sum(omc13) as t1 from oma_file,omc_file,azn_file where oma01 = omc01 and oma11 > to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and omc13 > 0 and oma11 = azn01 and year(oma11) = " & tYear & " group by azn05 "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select azn05,sum(tc_cif_04*oeb13*oea24) from tc_cif_file,oeb_file,azn_file,oea_file where tc_cif_01 = oeb01 and tc_cif_02 = oeb03 and oeb70 = 'N' and tc_cif_05 > to_date('"
        oCommand.CommandText += tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_05 = azn01 and oeb01 = oea01 and year(tc_cif_05) = " & tYear & " group by azn05 ) group by azn05 order by azn05"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Dim TempWeek As Decimal = oReader.Item("azn05")
                Ws.Cells(3, 4 + TempWeek - tWeek) = oReader.Item("t1")
            End While
        End If
        oReader.Close()

        oCommand.CommandText = "select azn05,sum(tc_ext05 * -1) as t1 from tc_ext_file,azn_File where tc_ext02 = 1 and tc_ext01 > to_date('" & tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_ext01 =azn01 and year(tc_ext01) = " & tYear & " and azn05 > " & tWeek & " group by azn05 order by azn05"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Dim TempWeek As Decimal = oReader.Item("azn05")
                Ws.Cells(4, 4 + TempWeek - tWeek) = oReader.Item("t1")
            End While
        End If
        oReader.Close()

        oCommand.CommandText = "select azn05,sum(tc_ext05 * -1) as t1 from tc_ext_file,azn_File where tc_ext02 = 2 and tc_ext01 > to_date('" & tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_ext01 =azn01 and year(tc_ext01) = " & tYear & " group by azn05 order by azn05"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Dim TempWeek As Decimal = oReader.Item("azn05")
                Ws.Cells(5, 4 + TempWeek - tWeek) = oReader.Item("t1")
            End While
        End If
        oReader.Close()

        oCommand.CommandText = "select azn05,sum(tc_ext05 * -1) as t1 from tc_ext_file,azn_File where tc_ext02 = 3 and tc_ext01 > to_date('" & tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_ext01 =azn01 and year(tc_ext01) = " & tYear & " group by azn05 order by azn05"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Dim TempWeek As Decimal = oReader.Item("azn05")
                Ws.Cells(6, 4 + TempWeek - tWeek) = oReader.Item("t1")
            End While
        End If
        oReader.Close()

        oCommand.CommandText = "select azn05,sum(tc_ext05 * -1) as t1 from tc_ext_file,azn_File where tc_ext02 = 4 and tc_ext01 > to_date('" & tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_ext01 =azn01 and year(tc_ext01) = " & tYear & " group by azn05 order by azn05"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Dim TempWeek As Decimal = oReader.Item("azn05")
                Ws.Cells(7, 4 + TempWeek - tWeek) = oReader.Item("t1")
            End While
        End If
        oReader.Close()

        'Ws.Cells(9, 4) = "=SUM(D2:D8)+C3"
        Ws.Cells(9, 5) = "=SUM(E2:E8)"
        oRng = Ws.Range("E9", "E9")
        oRng.Interior.Color = Color.FromArgb(141, 180, 226)
        'oRng.NumberFormatLocal = "¥#,##0_);[红色](¥#,##0)"
        oRng.AutoFill(Destination:=Ws.Range("E9", Ws.Cells(9, 4 + ColumnIndex)), Type:=xlFillDefault)

        oRng = Ws.Range("B2", Ws.Cells(9, 4 + ColumnIndex))
        oRng.NumberFormatLocal = "#,##0.00_);[红色](#,##0.00)"
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        'Ws.Columns.Font.Name = "Arial"
        Ws.Name = "Cash"

        oRng = Ws.Range("A2", "B2")
        oRng.EntireColumn.ColumnWidth = 32.5

        oRng = Ws.Range("C1", "BD1")
        oRng.EntireColumn.ColumnWidth = 9.89

        'Ws.Cells(1, 1) = "收支周别"
        Ws.Cells(1, 1) = "Currency"
        Ws.Cells(1, 2) = "RMB"
        Ws.Cells(1, 3) = "W0(前期未收/付款余额）"
        ColumnIndex = 0
        For i As Int16 = tWeek To MaxWeek Step 1
            Ws.Cells(1, 4 + ColumnIndex) = "W" & i
            ColumnIndex += 1
        Next
        oRng = Ws.Range("C1", Ws.Cells(1, 4 + ColumnIndex))
        oRng.Interior.Color = Color.MistyRose

        Ws.Cells(2, 1) = "DAC bank balance "
        Ws.Cells(3, 1) = "HK AR Collected"
        Ws.Cells(4, 1) = "DAC Fixed Exp"
        Ws.Cells(5, 1) = "DAC Capex"
        Ws.Cells(6, 1) = "DAC non-bom material"
        Ws.Cells(7, 1) = "DAC Molds & Jigs"
        Ws.Cells(8, 1) = "DAC Material"
        Ws.Cells(9, 1) = "Ending balance"
        Ws.Cells(9, 4) = "=C3+C8+SUM(D2:D8)"

        'oRng = Ws.Range("C9", "C9")
        'oRng.AutoFill(Destination:=Ws.Range("C9", Ws.Cells(9, 4 + ColumnIndex)), Type:=xlFillDefault)


        'oRng = Ws.Range("D9", "O41")
        'oRng.NumberFormatLocal = "#,##0,"

        LineZ = 2
    End Sub
    Private Function Getaah(ByVal aah01_s As String, ByVal aah01_e As String)
        oCommand.CommandText = "select nvl(sum(aah04-aah05),0) from aah_file,aag_file where aah02 = " & tYear & " and aah03 <= " & tMonth & " and aah01 between '" & aah01_s & "' and '" & aah01_e & "' and aah01 = aag01 and aag07 in ('2','3')"
        Dim aahA As Decimal = oCommand.ExecuteScalar()
        Return aahA
    End Function
    Private Function GetAR()
        oCommand.CommandText = "select nvl(sum(omc13),0) from oma_file,omc_file where oma01 = omc01 and oma11 <= to_date('" & tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and omc13 > 0"
        Dim DAR As Decimal = oCommand.ExecuteScalar
        Return DAR
    End Function
    Private Function GetAP()
        oCommand.CommandText = "select nvl(sum(apc13),0) from apa_file,apc_file where apa01 = apc01 and apa12 <= to_date('" & tDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and apc13 > 0 and apa63 = 1"
        Dim DAR As Decimal = oCommand.ExecuteScalar
        DAR = DAR * Decimal.MinusOne
        Return DAR
    End Function
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "CashFlow_Forecast"
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