Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form81
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim Start1 As String = String.Empty
    Dim End1 As String = String.Empty
    Dim TotalPeriod As Int16 = 0
    Dim LineZ As Integer = 0
    Dim SC As String = String.Empty
    Dim TR As Decimal = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form81_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommand2.Connection = oConnection
                oCommand2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        oCommand.CommandText = "SELECT count(DISTINCT bmb03) FROM BMB_FILE,BMA_FILE,IMA_FILE WHERE BMB01 = BMA01 AND BMA10 = '2' AND BMB03 = IMA01 AND IMA06 NOT IN ('102','103') and bmb05 is null and bmaacti = 'Y'"
        TR = oCommand.ExecuteScalar()
        If TR > 0 Then
            Me.ProgressBar1.Maximum = TR
            Me.ProgressBar1.Value = 0
            BackgroundWorker1.RunWorkerAsync()
            'ExportToExcel()
            'SaveExcel()
        Else
            MsgBox("无资料")
            Return
        End If
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Purpose of Material"
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
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        AdjustExcelFormat()
        oCommand.CommandText = "SELECT distinct bmb03,ima02,ima021,ima25,ima44,ima44_fac,(ima48+ima49+ima491+ima50) as t1,ima46,ima45 FROM BMB_FILE,BMA_FILE,IMA_FILE WHERE BMB01 = BMA01 AND BMA10 = '2' AND BMB03 = IMA01 AND IMA06 NOT IN ('102','103') and bmb05 is null and bmaacti = 'Y' "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("bmb03")
                Ws.Cells(LineZ, 2) = oReader.Item("ima02")
                Ws.Cells(LineZ, 3) = oReader.Item("ima021")
                Ws.Cells(LineZ, 5) = oReader.Item("ima25")
                Ws.Cells(LineZ, 6) = oReader.Item("ima44_fac")
                Ws.Cells(LineZ, 7) = oReader.Item("ima44")
                Ws.Cells(LineZ, 8) = oReader.Item("t1")
                Ws.Cells(LineZ, 9) = oReader.Item("ima46")
                Ws.Cells(LineZ, 10) = oReader.Item("ima45")
                GetCustomer(oReader.Item("bmb03"))
                GetMasterItem(oReader.Item("bmb03"))
                GetSector(oReader.Item("bmb03"))
                GetAffectDate(oReader.Item("bmb03"))
                LineZ += 1
                Me.ProgressBar1.Value += 1
            End While
        End If
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 13.75
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(1, 1) = "元件料号"
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 7.5
        Ws.Cells(1, 2) = "品名"
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.ColumnWidth = 61
        Ws.Cells(1, 3) = "规格"
        oRng = Ws.Range("D1", "D1")
        oRng.EntireColumn.ColumnWidth = 10.75
        Ws.Cells(1, 4) = "生效日期"
        oRng = Ws.Range("E1", "J1")
        oRng.EntireColumn.ColumnWidth = 8.75
        Ws.Cells(1, 5) = "库存单位"
        Ws.Cells(1, 6) = "换算率"
        Ws.Cells(1, 7) = "采购单位"
        Ws.Cells(1, 8) = "前置期"
        Ws.Cells(1, 9) = "MOQ"
        Ws.Cells(1, 10) = "MPQ"
        oRng = Ws.Range("K1", "K1")
        oRng.EntireColumn.ColumnWidth = 15.88
        Ws.Cells(1, 11) = "产品客户代码"
        oRng = Ws.Range("L1", "L1")
        oRng.EntireColumn.ColumnWidth = 23.38
        Ws.Cells(1, 12) = "主件简号"
        oRng = Ws.Range("M1", "M1")
        oRng.EntireColumn.ColumnWidth = 18.38
        Ws.Cells(1, 13) = "应用生产工序"
        oRng = Ws.Range("A1", "M1")
        oRng.Interior.Color = Color.LightBlue

        LineZ = 2
    End Sub
    Private Sub GetCustomer(ByVal bmb03 As String)
        Dim CS1 As String = String.Empty
        oCommand2.CommandText = "select distinct substr(bmb01,4,2) as c1 from bmb_file,bma_file where bmb01 = bma01 and  bmb03 = '" & bmb03 & "' and bmb05 is null and bma10 = '2' and bmaacti = 'Y'"
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                CS1 += oReader2.Item("c1") & "|"
            End While
            CS1 = CS1.Remove(CS1.Length() - 1, 1)
            Ws.Cells(LineZ, 11) = CS1
        End If
        oReader2.Close()
    End Sub
    Private Sub GetMasterItem(ByVal bmb03 As String)
        Dim CS1 As String = String.Empty
        oCommand2.CommandText = "select distinct substr(bmb01,4,6) as c1 from bmb_file,bma_file where bmb01 = bma01 and  bmb03 = '" & bmb03 & "' and bmb05 is null and bma10 = '2' and bmaacti = 'Y'"
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                CS1 += oReader2.Item("c1") & "|"
            End While
            CS1 = CS1.Remove(CS1.Length() - 1, 1)
            Ws.Cells(LineZ, 12) = CS1
        End If
        oReader2.Close()
    End Sub
    Private Sub GetSector(ByVal bmb03 As String)
        Dim CS1 As String = String.Empty
        oCommand2.CommandText = "select distinct  (case when substr(bmb01,length(bmb01),1) = 'A' then substr(bmb01,length(bmb01)-2,3) else "
        oCommand2.CommandText += "substr(bmb01,length(bmb01)-1,2) end) as c1 from bmb_file,bma_file where bmb01 = bma01 and bma10 = '2' and  bmb03 = '"
        oCommand2.CommandText += bmb03 & "' and bmaacti = 'Y'"
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                Select Case oReader2.Item("c1")
                    Case "31"
                        CS1 += "裁纱" & "|"
                    Case "32"
                        CS1 += "预型" & "|"
                    Case "32A"
                        CS1 += "二次预型" & "|"
                    Case "35"
                        CS1 += "成型" & "|"
                    Case "35A"
                        CS1 += "二次成型" & "|"
                    Case "36"
                        CS1 += "CNC" & "|"
                    Case "36A"
                        CS1 += "二次CNC" & "|"
                    Case "61"
                        CS1 += "补土" & "|"
                    Case "61A"
                        CS1 += "二次补土" & "|"
                    Case "63"
                        CS1 += "涂装" & "|"
                    Case "63A"
                        CS1 += "二次涂装" & "|"
                    Case "64"
                        CS1 += "胶合" & "|"
                    Case "64A"
                        CS1 += "二次胶合" & "|"
                    Case "65"
                        CS1 += "抛光" & "|"
                    Case "65A"
                        CS1 += "二次抛光" & "|"
                    Case "66"
                        CS1 += "包装" & "|"
                End Select
            End While
            If CS1.Length > 0 Then
                CS1 = CS1.Remove(CS1.Length() - 1, 1)
            End If
            Ws.Cells(LineZ, 13) = CS1
        End If
        oReader2.Close()
    End Sub
    Private Sub GetAffectDate(ByVal bmb03 As String)
        Dim CS1 As String = String.Empty
        oCommand2.CommandText = "select max(bmb04) as c2,min(bmb04) as c1 from bmb_file,bma_file where bmb01 = bma01 and  bmb03 = '" & bmb03 & "' and bmb05 is null and bma10 = '2' and bmaacti = 'Y'"
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            oReader2.Read()
            If oReader2.Item("c1") = oReader2.Item("c2") Then
                CS1 = oReader2.Item("c2")
            Else
                CS1 = oReader2.Item("c1") & "|" & oReader2.Item("c2")
            End If
            Ws.Cells(LineZ, 4) = CS1
        End If
        oReader2.Close()
    End Sub
End Class