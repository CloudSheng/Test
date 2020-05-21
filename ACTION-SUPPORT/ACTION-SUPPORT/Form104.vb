Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form104
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

    Private Sub Form104_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
           & "(case when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0478') then '10405' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0605') then '20410' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0400') then '30415' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0480') then '40440' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0610') then '50435' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0650') then '60460' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0490') then '70540' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0620') then '80570' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0492') then '90583' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0405') then 'A0417' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0435') then 'B0445' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0629') then 'C0465' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0633') then 'D0545' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0673') then 'E0575' " _
            & "when s.topreworkstation = '0480' then 'F0584' " _
            & "else '' end) as AA ,c.value " _
            & "FROM lot l JOIN model m ON l.model=m.model JOIN sn s ON l.lot=s.lot " _
            & "LEFT JOIN model_paravalue c on m.model = c.model and c.parameter = 'ERP PN'" _
            & "WHERE s.updatedstation <> '9999' "
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
           & "(case when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0478') then '10405' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0605') then '20410' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0400') then '30415' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0480') then '40440' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0610') then '50435' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0650') then '60460' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0490') then '70540' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0620') then '80570' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0492') then '90583' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0405') then 'A0417' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0435') then 'B0445' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0629') then 'C0465' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0633') then 'D0545' " _
            & "when (s.topreworkstation is null or s.topreworkstation = '') and s.updatedstation = ('0673') then 'E0575' " _
            & "when s.topreworkstation = '0480' then 'F0584' " _
            & "else '' end) as AA ,c.value " _
            & "FROM lot l JOIN model m ON l.model=m.model JOIN sn s ON l.lot=s.lot " _
            & "LEFT JOIN model_paravalue c on m.model = c.model and c.parameter = 'ERP PN'" _
            & "WHERE s.updatedstation <> '9999' "
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
                    CheckFormat = mSQLReader.Item("model")
                    Ws.Cells(LineZ, 1) = mSQLReader.Item("Value")
                    Ws.Cells(LineZ, 2) = CheckFormat
                    Ws.Cells(LineZ, 3) = mSQLReader.Item("modelname")
                    oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 3))
                    oRng.Interior.Color = Color.LightBlue
                End If
                If Not CheckFormat = mSQLReader.Item("model") Then
                    LineZ += 1
                    Ws.Cells(LineZ, 1) = mSQLReader.Item("Value")
                    Ws.Cells(LineZ, 2) = mSQLReader.Item("model")
                    Ws.Cells(LineZ, 3) = mSQLReader.Item("modelname")
                    oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 3))
                    oRng.Interior.Color = Color.LightBlue
                    CheckFormat = mSQLReader("model")
                End If
                Dim CheckPosition As String = Strings.Left(mSQLReader.Item("AA").ToString, 1)
                Select Case CheckPosition
                    Case "1"
                        Ws.Cells(LineZ, 4) = mSQLReader.Item("Count1")
                    Case "2"
                        Ws.Cells(LineZ, 5) = mSQLReader.Item("Count1")
                    Case "3"
                        Ws.Cells(LineZ, 6) = mSQLReader.Item("Count1")
                    Case "4"
                        Ws.Cells(LineZ, 7) = mSQLReader.Item("Count1")
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
                        Ws.Cells(LineZ, 18) = mSQLReader.Item("Count1")
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
                        Ws.Cells(LineZ, 13) = mSQLReader.Item("Count1")
                End Select
                ProgressBar1.Value += 1
            End While
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
        Ws.Cells(2, 4) = "胶合接收 478"
        Ws.Cells(2, 5) = "二次胶合接收 605"
        Ws.Cells(2, 6) = "喷砂 400"
        Ws.Cells(2, 7) = "胶合1 480"
        Ws.Cells(2, 8) = "胶合2 610"
        Ws.Cells(2, 9) = "装配 650"
        Ws.Cells(2, 10) = "胶合检验1 490"
        Ws.Cells(2, 11) = "胶合检验2 620"
        Ws.Cells(2, 12) = "后硬化 492"
        Ws.Cells(2, 13) = "胶合段返工品"
        Ws.Cells(2, 14) = "补土接收 405"
        Ws.Cells(2, 15) = "二次补土接收 435"
        Ws.Cells(2, 16) = "抛光接收 629"
        Ws.Cells(2, 17) = "二次抛光接收 633"
        Ws.Cells(2, 18) = "包装接收 673"
        LineZ = 3
    End Sub

    Private Sub SumColumn()
        Ws.Cells(1, 4) = "=SUM(D4:D" & LineZ & ")"
        oRng = Ws.Range("D1", "D1")
        oRng.AutoFill(Destination:=Ws.Range("D1", "R1"), Type:=xlFillDefault)
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Gluing_WIP_STATUS"
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