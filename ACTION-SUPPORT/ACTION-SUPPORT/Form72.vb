Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form72
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

    Private Sub Form72_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        '  & "when t.station in ('0405','0410','0415','0417','0418','0420','0430','0440','0445','0450','0455') then '7補土' " _
        '  & "when t.station in ('0460','0465','0470','0475','0540','0545','0550','0560','0563','0567','0570','0575','0580','0583','0584','0585','0587','0590','0592','0595') then '8涂裝' " _
        mSQLS1.CommandText = "SELECT COUNT(MODEL) FROM ( " _
        & "select model,modelname,AA,count(*) as Count1,Value from ( SELECT m.model,m.modelname," _
         & "(case when t.station in ('0055','0080','0100','0110','0111') then '1裁纱' " _
           & "when t.station in ('0112','0113') then '2备料' " _
           & "when t.station in ('0130','0140','0150','0151','0160','0170','0172','0175','0177','0180','0193') then '3预型' " _
           & "when t.station in ('0165','0173','0174','0175','0190','0195','0200','0210','0215','0220','0223','0225','0230','0231','0240','0250','0255','0260','0280','0300','0315','0316','0320','0321','0325','0326','0330','0331','0333','0390','0395') then '4成型' " _
           & "when t.station in ('0335','0340','0350','0360','0370','0380','0390','0495','0500','0510','0520','0530') then '5CNC' " _
           & "when t.station in ('0400','0435','0478','0480','0490','0492','0493','0605','0610','0620','0623','0627') then '6胶合' " _
           & "when t.station in ('0629','0630','0633','0635','0640','0645') then '9拋光' " _
           & "when t.station in ('0642','0649','0650','0652','0657','0658','0659','0660','0665','0666','0667','0670','0673') then 'A待包裝' " _
           & "when t.station in ('0675','0680','0690') then 'B已包裝' " _
           & "when t.station in ('0642','0649','0650','0652','0657','0658','0659','0660','0665','0666','0667','0670','0673','0675','0680','0690') then 'C包裝' " _
           & "when t.station in ('0405') then 'D底漆防漆' " _
           & "when t.station in ('0410','0415','0417','0440','0475') then 'E底漆研磨' " _
           & "when t.station in ('0418','0420','0430','0445','0450','0455') then 'F底漆涂装' " _
           & "when t.station in ('0460','0465','0540','0545','0567','0570','0575','0583','0584') then 'G涂装研磨' " _
           & "when t.station in ('0470','0550','0560','0563','0580','0585','0587','0590','0591','0592','0595') then 'H面漆涂装' else '~XX' end) as AA ,c.value " _
           & "FROM lot l JOIN model m ON l.model=m.model JOIN sn s ON l.lot=s.lot JOIN station t ON t.station=case when (s.topreworkstation is not null and s.topreworkstation<>'') then s.topreworkstation else s.updatedstation end " _
           & "LEFT JOIN model_paravalue c on m.model = c.model and c.parameter = 'ERP PN'" _
           & "WHERE t.station  <> '9999' "
        If Not String.IsNullOrEmpty(tModel_type) Then
            mSQLS1.CommandText += "and m.model_type like '" & tModel_type & "' "
        End If
        If Not String.IsNullOrEmpty(tModel) Then
            mSQLS1.CommandText += "and l.model like '" & tModel & "' "
        End If
        If Not String.IsNullOrEmpty(tLot) Then
            mSQLS1.CommandText += "and l.lot like '" & tLot & "' "
        End If
        mSQLS1.CommandText += ")  as B  GROUP BY model,modelname,AA,value ) AS C WHERE 1 =1"
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
        Me.ProgressBar1.Maximum = HaveReport
        mSQLS1.CommandText = "select model,modelname,AA,count(*) as Count1,Value from ( SELECT m.model,m.modelname," _
          & "(case when t.station in ('0055','0080','0100','0110','0111') then '1裁纱' " _
           & "when t.station in ('0112','0113') then '2备料' " _
           & "when t.station in ('0130','0140','0150','0151','0160','0170','0172','0175','0177','0180','0193') then '3预型' " _
           & "when t.station in ('0165','0173','0174','0175','0190','0195','0200','0210','0215','0220','0223','0225','0230','0231','0240','0250','0255','0260','0280','0300','0315','0316','0320','0321','0325','0326','0330','0331','0333','0390','0395') then '4成型' " _
           & "when t.station in ('0335','0340','0350','0360','0370','0380','0390','0495','0500','0510','0520','0530') then '5CNC' " _
           & "when t.station in ('0400','0435','0478','0480','0490','0492','0493','0605','0610','0620','0623','0627') then '6胶合' " _
          & "when t.station in ('0629','0630','0633','0635','0640','0645') then '9拋光' " _
           & "when t.station in ('0642','0649','0650','0652','0657','0658','0659','0660','0665','0666','0667','0670','0673') then 'A待包裝' " _
           & "when t.station in ('0675','0680','0690') then 'B已包裝' " _
           & "when t.station in ('0642','0649','0650','0652','0657','0658','0659','0660','0665','0666','0667','0670','0673','0675','0680','0690') then 'C包裝' " _
           & "when t.station in ('0405') then 'D底漆防漆' " _
           & "when t.station in ('0410','0415','0417','0440','0475') then 'E底漆研磨' " _
           & "when t.station in ('0418','0420','0430','0445','0450','0455') then 'F底漆涂装' " _
           & "when t.station in ('0460','0465','0540','0545','0567','0570','0575','0583','0584') then 'G涂装研磨' " _
           & "when t.station in ('0470','0550','0560','0563','0580','0585','0587','0590','0591','0592','0595') then 'H面漆涂装' else '~XX' end) as AA ,c.value " _
           & "FROM lot l JOIN model m ON l.model=m.model JOIN sn s ON l.lot=s.lot JOIN station t ON t.station=case when (s.topreworkstation is not null and s.topreworkstation<>'') then s.topreworkstation else s.updatedstation end " _
           & "LEFT JOIN model_paravalue c on m.model = c.model and c.parameter = 'ERP PN'" _
           & "WHERE t.station  <> '9999' "
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
                        'Case "7"
                        '   Ws.Cells(LineZ, 10) = mSQLReader.Item("Count1")
                        'Case "8"
                        '   Ws.Cells(LineZ, 11) = mSQLReader.Item("Count1")
                    Case "9"
                        Ws.Cells(LineZ, 12) = mSQLReader.Item("Count1")
                    Case "C"
                        Ws.Cells(LineZ, 13) = mSQLReader.Item("Count1")
                    Case "A"
                        Ws.Cells(LineZ, 14) = mSQLReader.Item("Count1")
                    Case "B"
                        Ws.Cells(LineZ, 15) = mSQLReader.Item("Count1")
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
                End Select
                ProgressBar1.Value += 1
            End While
            SumRow()
        End If
        mSQLReader.Close()
        ' 全部COLUMN 加總  20160427
        SumColumn()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "WIP_STATUS2"
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

    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "C1")
        oRng.EntireColumn.ColumnWidth = 40
        oRng = Ws.Range("D1", "T1")
        oRng.EntireColumn.ColumnWidth = 15
        oRng = Ws.Range("A2", "T2")
        oRng.Interior.Color = Color.LightBlue
        Ws.Cells(1, 3) = "数量合计"
        Ws.Cells(2, 1) = "ERP 料号"
        Ws.Cells(2, 2) = "产品名称"
        Ws.Cells(2, 3) = "产品名称"
        Ws.Cells(2, 4) = "标签"
        Ws.Cells(2, 5) = "裁紗"
        Ws.Cells(2, 6) = "预型"
        Ws.Cells(2, 7) = "成型"
        Ws.Cells(2, 8) = "CNC"
        Ws.Cells(2, 9) = "胶合"
        Ws.Cells(2, 10) = "补土"
        Ws.Cells(2, 11) = "涂装"
        Ws.Cells(2, 12) = "抛光"
        Ws.Cells(2, 13) = "包装合计"
        Ws.Cells(2, 14) = "待包装"
        Ws.Cells(2, 15) = "已包装"
        Ws.Cells(2, 16) = "底漆防漆"
        Ws.Cells(2, 17) = "底漆研磨"
        Ws.Cells(2, 18) = "底漆涂装"
        Ws.Cells(2, 19) = "涂装研磨"
        Ws.Cells(2, 20) = "面漆涂装"
        LineZ = 3
    End Sub
    Private Sub SumColumn()
        Ws.Cells(1, 4) = "=SUBTOTAL(9,D3:D" & LineZ & ")"
        oRng = Ws.Range("D1", "D1")
        oRng.AutoFill(Destination:=Ws.Range("D1", "T1"), Type:=xlFillDefault)
    End Sub
    Private Sub SumRow()
        Ws.Cells(LineZ, 13) = "=SUM(N" & LineZ & ":O" & LineZ & ")"
        Ws.Cells(LineZ, 10) = "=SUM(P" & LineZ & ":R" & LineZ & ")"
        Ws.Cells(LineZ, 11) = "=SUM(S" & LineZ & ":T" & LineZ & ")"
    End Sub
End Class