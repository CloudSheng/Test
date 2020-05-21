Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form77
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim tModel_type As String
    Dim tModel As String
    Dim tLot As String
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form77_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        BindModel_Type()
        Dim Model_Type As String = String.Empty
        BindModel(Model_Type)
        BindLot(Model_Type)
    End Sub
    Private Sub BindModel_Type()
        Me.ComboBox1.Items.Clear()
        mSQLS1.CommandText = "SELECT * FROM model_type WHERE model_type <> 'Action'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox1.Items.Add(mSQLReader.Item(0).ToString())
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub BindModel(ByVal Models1 As String)
        Me.ComboBox2.Items.Clear()
        mSQLS1.CommandText = "select distinct lot.model,model.modelname  from lot,model " _
                          & " where lot.model = model.model and model.model_type <> 'Action'"
        If Not String.IsNullOrEmpty(Models1) Then
            mSQLS1.CommandText += " AND model.model_type = '" & Models1 & "'"
        End If
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox2.Items.Add(mSQLReader.Item(0).ToString() & "|" & mSQLReader.Item(1).ToString())
            End While
        End If
        mSQLReader.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim model_type As String = ComboBox1.Items(ComboBox1.SelectedIndex).ToString()
        BindModel(model_type)
        BindLot(model_type)
    End Sub

    Private Sub BindLot(ByVal Models1 As String)
        Me.ComboBox3.Items.Clear()
        mSQLS1.CommandText = "select distinct lot.lot from lot,model " _
                          & " where lot.model = model.model and model.model_type <> 'Action'"
        If Not String.IsNullOrEmpty(Models1) Then
            mSQLS1.CommandText += " AND model.model_type = '" & Models1 & "'"
        End If
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox3.Items.Add(mSQLReader.Item(0).ToString())
            End While
        End If
        mSQLReader.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        ProgressBar1.Value = 0
        tModel_type = String.Empty
        tModel = String.Empty
        tLot = String.Empty
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If Not IsNothing(ComboBox1.SelectedItem) Then
            tModel_type = ComboBox1.SelectedItem.ToString()
        End If
        If Not IsNothing(ComboBox2.SelectedItem) Then
            tModel = ComboBox2.SelectedItem.ToString()
            Dim stCount As Int16 = Strings.InStr(tModel, "|")
            If stCount > 0 Then
                tModel = Strings.Left(tModel, stCount - 1)
            End If
        End If
        If Not IsNothing(ComboBox3.SelectedItem) Then
            tLot = ComboBox3.SelectedItem.ToString()
        End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        mSQLS1.CommandText = "SELECT COUNT(MODEL) FROM ( " _
            & "select model,modelname,AA,count(*) as Count1 from ( SELECT m.model,m.modelname," _
           & "(case when t.station in ('0405') then '10405' " _
            & "when t.station in ('0410') then '20410' " _
            & "when t.station in ('0415') then '30415' " _
            & "when t.station = '0440' then '40440' " _
            & "when t.station = '0435' then '50435' " _
            & "when t.station = '0460' then '60460' " _
            & "when t.station = '0540' then '70540' " _
            & "when t.station = '0570' then '80570' " _
            & "when t.station = '0583' then '90583' " _
            & "when t.station = '0417' then 'A0417' " _
            & "when t.station = '0445' then 'B0445' " _
            & "when t.station = '0465' then 'C0465' " _
            & "when t.station = '0545' then 'D0545' " _
            & "when t.station = '0575' then 'E0575' " _
            & "when t.station = '0584' then 'F0584' " _
            & "when t.station in ('0418','0420') then 'G0418' " _
            & "when t.station = '0450' then 'H0450' " _
            & "when t.station = '0455' then 'I0455' " _
            & "when t.station = '0470' then 'J0470' " _
            & "when t.station = '0550' then 'K0550' " _
            & "when t.station = '0567' then 'L0567' " _
            & "when t.station = '0560' then 'M0560' " _
            & "when t.station = '0563' then 'N0563' " _
            & "when t.station = '0580' then 'O0580' " _
            & "when t.station = '0585' then 'P0585' " _
            & "when t.station = '0587' then 'Q0587' " _
            & "when t.station = '0590' then 'R0590' " _
            & "when t.station = '0595' then 'S0595' " _
            & "when t.station = '0591' then 'T0591' " _
            & "when t.station = '0592' then 'U0592' " _
            & "else '' end) as AA ,c.value " _
            & "FROM lot l JOIN model m ON l.model=m.model JOIN sn s ON l.lot=s.lot JOIN station t ON t.station=case when (s.topreworkstation is not null and s.topreworkstation<>'') then s.topreworkstation else s.updatedstation end " _
            & "LEFT JOIN model_paravalue c on m.model = c.model and c.parameter = 'ERP PN'" _
            & "WHERE t.station  <> '9999' and t.station between '0405' and '0592'  "
        If Not String.IsNullOrEmpty(tModel_type) Then
            mSQLS1.CommandText += "and m.model_type like '" & tModel_type & "' "
        End If
        If Not String.IsNullOrEmpty(tModel) Then
            mSQLS1.CommandText += "and l.model like '" & tModel & "' "
        End If
        If Not String.IsNullOrEmpty(tLot) Then
            mSQLS1.CommandText += "and l.lot like '" & tLot & "' "
        End If
        mSQLS1.CommandText += ")  as B  GROUP BY model,modelname,AA ) AS C WHERE 1 =1"
        Dim HaveReport As Integer = mSQLS1.ExecuteScalar()
        If HaveReport = 0 Then
            MsgBox("没有资料，请重选条件")
            Return
        End If
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        Dim CheckFormat As String = String.Empty
        'LineZ = 4
        Me.ProgressBar1.Maximum = HaveReport
        mSQLS1.CommandText = "select model,modelname,AA,count(*) as Count1,Value from ( SELECT m.model,m.modelname," _
           & "(case when t.station in ('0405') then '10405' " _
            & "when t.station in ('0410') then '20410' " _
            & "when t.station in ('0415') then '30415' " _
            & "when t.station = '0440' then '40440' " _
            & "when t.station = '0435' then '50435' " _
            & "when t.station = '0460' then '60460' " _
            & "when t.station = '0540' then '70540' " _
            & "when t.station = '0570' then '80570' " _
            & "when t.station = '0583' then '90583' " _
            & "when t.station = '0417' then 'A0417' " _
            & "when t.station = '0445' then 'B0445' " _
            & "when t.station = '0465' then 'C0465' " _
            & "when t.station = '0545' then 'D0545' " _
            & "when t.station = '0575' then 'E0575' " _
            & "when t.station = '0584' then 'F0584' " _
            & "when t.station in ('0418','0420') then 'G0418' " _
            & "when t.station = '0450' then 'H0450' " _
            & "when t.station = '0455' then 'I0455' " _
            & "when t.station = '0470' then 'J0470' " _
            & "when t.station = '0550' then 'K0550' " _
            & "when t.station = '0567' then 'L0567' " _
            & "when t.station = '0560' then 'M0560' " _
            & "when t.station = '0563' then 'N0563' " _
            & "when t.station = '0580' then 'O0580' " _
            & "when t.station = '0585' then 'P0585' " _
            & "when t.station = '0587' then 'Q0587' " _
            & "when t.station = '0590' then 'R0590' " _
            & "when t.station = '0595' then 'S0595' " _
            & "when t.station = '0591' then 'T0591' " _
            & "when t.station = '0592' then 'U0592' " _
            & "else '' end) as AA ,c.value " _
            & "FROM lot l JOIN model m ON l.model=m.model JOIN sn s ON l.lot=s.lot JOIN station t ON t.station=case when (s.topreworkstation is not null and s.topreworkstation<>'') then s.topreworkstation else s.updatedstation end " _
            & "LEFT JOIN model_paravalue c on m.model = c.model and c.parameter = 'ERP PN'" _
            & "WHERE t.station  <> '9999' and t.station between '0405' and '0592' "
        If Not String.IsNullOrEmpty(tModel_type) Then
            mSQLS1.CommandText += "and m.model_type like '" & tModel_type & "' "
        End If
        If Not String.IsNullOrEmpty(tModel) Then
            mSQLS1.CommandText += "and l.model like '" & tModel & "' "
        End If
        If Not String.IsNullOrEmpty(tLot) Then
            mSQLS1.CommandText += "and l.lot like '" & tLot & "' "
        End If
        mSQLS1.CommandText += ")  as B  GROUP BY model,modelname,AA,Value  order by model,modelname,AA"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                If String.IsNullOrEmpty(CheckFormat) Then
                    SumRow()
                    CheckFormat = mSQLReader.Item("model")
                    Ws.Cells(LineZ, 1) = mSQLReader.Item("Value")
                    Ws.Cells(LineZ, 2) = CheckFormat
                    Ws.Cells(LineZ, 3) = mSQLReader.Item("modelname")
                    oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 3))
                    oRng.Interior.Color = Color.LightBlue
                End If
                If Not CheckFormat = mSQLReader.Item("model") Then
                    SumRow()
                    LineZ += 1
                    Ws.Cells(LineZ, 1) = mSQLReader.Item("Value")
                    Ws.Cells(LineZ, 2) = mSQLReader.Item("model")
                    Ws.Cells(LineZ, 3) = mSQLReader.Item("modelname")
                    oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 3))
                    oRng.Interior.Color = Color.LightBlue
                    CheckFormat = mSQLReader("model")
                End If
                Dim CheckPosition As String = Strings.Left(mSQLReader.Item("AA").ToString, 1)
                        Select CheckPosition
                    Case "1"
                        Ws.Cells(LineZ, 4) = mSQLReader.Item("Count1")
                    Case "2"
                        Ws.Cells(LineZ, 5) = mSQLReader.Item("Count1")
                    Case "3"
                        Ws.Cells(LineZ, 6) = mSQLReader.Item("Count1")
                    Case "4"
                        Ws.Cells(LineZ, 7) = mSQLReader.Item("Count1")
                        '            Case "Z"
                        '                Ws.Cells(LineZ, 8) = mSQLReader.Item("Count1")
                    Case "5"
                        Ws.Cells(LineZ, 8) = mSQLReader.Item("Count1")
                    Case "6"
                        Ws.Cells(LineZ, 9) = mSQLReader.Item("Count1")
                    Case "7"
                        Ws.Cells(LineZ, 10) = mSQLReader.Item("Count1")
                    Case "8"
                        Ws.Cells(LineZ, 11) = mSQLReader.Item("Count1")
                    Case "9"
                        Ws.Cells(LineZ, 12) = mSQLReader.Item("Count1")
                    Case "A"
                        Ws.Cells(LineZ, 13) = mSQLReader.Item("Count1")
                    Case "B"
                        Ws.Cells(LineZ, 14) = mSQLReader.Item("Count1")
                    Case "C"
                        Ws.Cells(LineZ, 15) = mSQLReader.Item("Count1")
                        'Case "Y"
                        '   Ws.Cells(LineZ, 16) = mSQLReader.Item("Count1")
                    Case "D"
                        Ws.Cells(LineZ, 16) = mSQLReader.Item("Count1")
                    Case "E"
                        Ws.Cells(LineZ, 17) = mSQLReader.Item("Count1")
                    Case "F"
                        Ws.Cells(LineZ, 18) = mSQLReader.Item("Count1")
                    Case "G"
                        Ws.Cells(LineZ, 19) = mSQLReader.Item("Count1")
                    Case "H"
                        Ws.Cells(LineZ, 20) = mSQLReader.Item("Count1")
                    Case "I"
                        Ws.Cells(LineZ, 21) = mSQLReader.Item("Count1")
                    Case "J"
                        Ws.Cells(LineZ, 22) = mSQLReader.Item("Count1")
                    Case "K"
                        Ws.Cells(LineZ, 23) = mSQLReader.Item("Count1")
                    Case "L"
                        Ws.Cells(LineZ, 24) = mSQLReader.Item("Count1")
                    Case "M"
                        Ws.Cells(LineZ, 25) = mSQLReader.Item("Count1")
                    Case "N"
                        Ws.Cells(LineZ, 26) = mSQLReader.Item("Count1")
                    Case "O"
                        Ws.Cells(LineZ, 27) = mSQLReader.Item("Count1")
                    Case "P"
                        Ws.Cells(LineZ, 28) = mSQLReader.Item("Count1")
                    Case "Q"
                        Ws.Cells(LineZ, 29) = mSQLReader.Item("Count1")
                    Case "R"
                        Ws.Cells(LineZ, 30) = mSQLReader.Item("Count1")
                    Case "S"
                        Ws.Cells(LineZ, 31) = mSQLReader.Item("Count1")
                    Case "T"
                        Ws.Cells(LineZ, 32) = mSQLReader.Item("Count1")
                    Case "U"
                        Ws.Cells(LineZ, 33) = mSQLReader.Item("Count1")
                        '            Case "V"
                        '                Ws.Cells(LineZ, 36) = mSQLReader.Item("Count1")
                        '            Case "W"
                        '                Ws.Cells(LineZ, 37) = mSQLReader.Item("Count1")
                        '            Case "X"
                        '                Ws.Cells(LineZ, 38) = mSQLReader.Item("Count1")
                End Select
                ProgressBar1.Value += 1
            End While
            SumRow()
        End If
        SumColumn()
        mSQLReader.Close()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "C1")
        oRng.EntireColumn.ColumnWidth = 40
        oRng = Ws.Range("D1", "AP1")
        oRng.EntireColumn.ColumnWidth = 15
        oRng = Ws.Range("A2", "AL3")
        oRng.Interior.Color = Color.LightBlue
        Ws.Cells(1, 3) = "合计数量："
        Ws.Cells(2, 1) = "ERP 料号"
        Ws.Cells(2, 2) = "产品名称"
        Ws.Cells(2, 3) = "产品名称"
        Ws.Cells(2, 4) = "补土接收站"
        Ws.Cells(2, 5) = "补土1"
        Ws.Cells(2, 6) = "点补"
        Ws.Cells(2, 7) = "补土2"
        Ws.Cells(2, 8) = "二次补土接收站"
        Ws.Cells(2, 9) = "补土3"
        Ws.Cells(2, 10) = "补土4"
        Ws.Cells(2, 11) = "补土5"
        Ws.Cells(2, 12) = "补土6"
        Ws.Cells(2, 13) = "清洁1"
        Ws.Cells(2, 14) = "清洁2"
        Ws.Cells(2, 15) = "清洁3"
        Ws.Cells(2, 16) = "清洁4"
        Ws.Cells(2, 17) = "清洁5"
        Ws.Cells(2, 18) = "清洁6"
        Ws.Cells(2, 19) = "涂装1"
        Ws.Cells(2, 20) = "涂装2"
        Ws.Cells(2, 21) = "涂装接收站"
        Ws.Cells(2, 22) = "涂装3"
        Ws.Cells(2, 23) = "涂装4"
        Ws.Cells(2, 24) = "二次涂装接收站"
        Ws.Cells(2, 25) = "贴水标"
        Ws.Cells(2, 26) = "水标检验"
        Ws.Cells(2, 27) = "涂装5"
        Ws.Cells(2, 28) = "涂装6"
        Ws.Cells(2, 29) = "涂装后称重"
        Ws.Cells(2, 30) = "涂装检验"
        Ws.Cells(2, 31) = "涂装检验1"
        Ws.Cells(2, 32) = "尺寸自检"
        Ws.Cells(2, 33) = "尺寸检验"
        Ws.Cells(2, 34) = "产品加总"
        LineZ = 3
    End Sub
    Private Sub SumRow()
        Ws.Cells(LineZ, 34) = "=SUM(D" & LineZ & ":AG" & LineZ & ")"
    End Sub
    Private Sub SumColumn()
        Ws.Cells(1, 4) = "=SUM(D4:D" & LineZ & ")"
        oRng = Ws.Range("D1", "D1")
        oRng.AutoFill(Destination:=Ws.Range("D1", "AH1"), Type:=xlFillDefault)
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "PAINTING_WIP_STATUS"
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
        If mConnection.State = ConnectionState.Open Then
            Try
                mConnection.Close()
                Module1.KillExcelProcess(OldExcel)
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
End Class