Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form2
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

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
           & "(case when t.station in ('0055','0080','0100','0110','0111') and s.block = 'N' then '1裁纱' " _
            & "when t.station in ('0112','0113') and s.block = 'N' then '2备料' " _
            & "when t.station in ('0150','0151','0165','0170','0172','0175','0177') and s.block = 'N' then '3预型' " _
            & "when (t.station between '0180' and '0329' or t.station in ('0173','0174')) and s.block = 'N' then '4成型' " _
            & "when t.station in ('0330','0331','0333') and s.block = 'N' then 'Z成型检验' " _
            & "when t.station in ('0390','0395') and s.block = 'N' then '5成型固化' " _
            & "when t.station in ('0335','0340','0350','0360','0370','0493','0495','0500','0510','0520') and s.block = 'N' then '6CNC' " _
            & "when t.station in ('0380','0385','0530') and s.block = 'N' then '7CNC检验' " _
            & "when t.station = '0400' and s.block = 'N' then '8喷砂' " _
            & "when t.station in ('0478','0480','0485') and s.block = 'N' then '9胶合1' " _
            & "when t.station in ('0605','0610','0611','0623') and s.block = 'N' then 'A胶合2' " _
            & "when t.station in ('0560','0563') and s.block = 'N' then 'B贴水标' " _
            & "when t.station in ('0490','0491','0492') and s.block = 'N' then 'C胶合检验1' " _
            & "when t.station in ('0620','0627') and s.block = 'N' then 'Y胶合检验2' " _
            & "when t.station in ('0405','0410') and s.block = 'N' then 'D补土1' " _
            & "when t.station in ('0415','0417') and s.block = 'N' then 'E点补' " _
            & "when t.station in ('0435','0440') and s.block = 'N' then 'F补土2' " _
            & "when t.station in ('0460') and s.block = 'N' then 'G补土3' " _
            & "when t.station in ('0540') and s.block = 'N' then 'H补土4' " _
            & "when t.station in ('0570') and s.block = 'N' then 'I补土5' " _
            & "when t.station in ('0583') and s.block = 'N' then 'J补土6' " _
            & "when t.station in ('0418','0420','0430','0441','0455') and s.block = 'N' then 'K涂装1' " _
            & "when t.station in ('0445','0450','0567') and s.block = 'N' then 'L涂装2' " _
            & "when t.station in ('0461','0465','0470','0475') and s.block = 'N' then 'M涂装3' " _
            & "when t.station in ('0541','0545','0550') and s.block = 'N' then 'N涂装4' " _
            & "when t.station in ('0575','0580') and s.block = 'N' then 'O涂装5' " _
            & "when t.station in ('0584','0585') and s.block = 'N' then 'P涂装6' " _
            & "when t.station in ('0587','0590','0591','0592','0595','0600') and s.block = 'N' then 'Q涂装完成' " _
            & "when t.station in ('0629','0630','0633','0635') and s.block = 'N' then 'R抛光' " _
            & "when t.station in ('0640','0642','0645','0657') and s.block = 'N' then 'S抛光检验' " _
            & "when t.station in ('0649','0650','0652','0658','0659','0665','0666','0667','0668','0669','0670','0673','0674') and s.block = 'N' then 'T静置' " _
            & "when t.station in ('0675','0680','0690') and s.block = 'N' then 'U包装' " _
            & "when t.station between '0720' and '0730' and s.block = 'N' then 'V成品' " _
            & "when t.station = 'BLCK' and s.block = 'N' then '~BLCK' " _
            & "when s.block = 'Y' then 'W隔離' " _
            & "when t.station = '0799' and s.block = 'N' then 'X0799' " _
            & "else '' end) as AA ,c.value " _
            & "FROM lot l JOIN model m ON l.model=m.model JOIN sn s ON l.lot=s.lot JOIN station t ON t.station=case when (s.topreworkstation is not null and s.topreworkstation<>'') then s.topreworkstation else s.updatedstation end " _
            & "LEFT JOIN model_paravalue c on m.model = c.model and c.parameter = 'ERP PN'" _
            & "WHERE t.station  <> '9999' " 'and l.remark not like '%Training%' "
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
        LineZ = 4
        Me.ProgressBar1.Maximum = HaveReport
        mSQLS1.CommandText = "select model,modelname,AA,count(*) as Count1,Value from ( SELECT m.model,m.modelname," _
            & "(case when t.station in ('0055','0080','0100','0110','0111') and s.block = 'N' then '1裁纱' " _
            & "when t.station in ('0112','0113') and s.block = 'N' then '2备料' " _
            & "when t.station in ('0150','0151','0165','0170','0172','0175','0177') and s.block = 'N' then '3预型' " _
            & "when (t.station between '0180' and '0329' or t.station in ('0173','0174')) and s.block = 'N' then '4成型' " _
            & "when t.station in ('0330','0331','0333') and s.block = 'N' then 'Z成型检验' " _
            & "when t.station in ('0390','0395') and s.block = 'N' then '5成型固化' " _
            & "when t.station in ('0335','0340','0350','0360','0370','0493','0495','0500','0510','0520') and s.block = 'N' then '6CNC' " _
            & "when t.station in ('0380','0385','0530') and s.block = 'N' then '7CNC检验' " _
            & "when t.station = '0400' and s.block = 'N' then '8喷砂' " _
            & "when t.station in ('0478','0480','0485') and s.block = 'N' then '9胶合1' " _
            & "when t.station in ('0605','0610','0611','0623') and s.block = 'N' then 'A胶合2' " _
            & "when t.station in ('0560','0563') and s.block = 'N' then 'B贴水标' " _
            & "when t.station in ('0490','0491','0492') and s.block = 'N' then 'C胶合检验1' " _
            & "when t.station in ('0620','0627') and s.block = 'N' then 'Y胶合检验2' " _
            & "when t.station in ('0405','0410') and s.block = 'N' then 'D补土1' " _
            & "when t.station in ('0415','0417') and s.block = 'N' then 'E点补' " _
            & "when t.station in ('0435','0440') and s.block = 'N' then 'F补土2' " _
            & "when t.station in ('0460') and s.block = 'N' then 'G补土3' " _
            & "when t.station in ('0540') and s.block = 'N' then 'H补土4' " _
            & "when t.station in ('0570') and s.block = 'N' then 'I补土5' " _
            & "when t.station in ('0583') and s.block = 'N' then 'J补土6' " _
            & "when t.station in ('0418','0420','0430','0441','0455') and s.block = 'N' then 'K涂装1' " _
            & "when t.station in ('0445','0450','0567') and s.block = 'N' then 'L涂装2' " _
            & "when t.station in ('0461','0465','0470','0475') and s.block = 'N' then 'M涂装3' " _
            & "when t.station in ('0541','0545','0550') and s.block = 'N' then 'N涂装4' " _
            & "when t.station in ('0575','0580') and s.block = 'N' then 'O涂装5' " _
            & "when t.station in ('0584','0585') and s.block = 'N' then 'P涂装6' " _
            & "when t.station in ('0587','0590','0591','0592','0595','0600') and s.block = 'N' then 'Q涂装完成' " _
            & "when t.station in ('0629','0630','0633','0635') and s.block = 'N' then 'R抛光' " _
            & "when t.station in ('0640','0642','0645','0657') and s.block = 'N' then 'S抛光检验' " _
            & "when t.station in ('0649','0650','0652','0658','0659','0665','0666','0667','0668','0669','0670','0673','0674') and s.block = 'N' then 'T静置' " _
            & "when t.station in ('0670','0675','0680','0690') and s.block = 'N' then 'U包装' " _
            & "when t.station between '0720' and '0730' and s.block = 'N' then 'V成品' " _
            & "when t.station = 'BLCK' and s.block = 'N' then '~BLCK' " _
            & "when s.block = 'Y' then 'W隔離' " _
            & "when t.station = '0799' and s.block = 'N' then 'X0799' " _
            & "else '~XX' end) as AA ,c.value " _
            & "FROM lot l JOIN model m ON l.model=m.model JOIN sn s ON l.lot=s.lot JOIN station t ON t.station=case when (s.topreworkstation is not null and s.topreworkstation<>'') then s.topreworkstation else s.updatedstation end " _
            & "LEFT JOIN model_paravalue c on m.model = c.model and c.parameter = 'ERP PN'" _
            & "WHERE t.station  <> '9999' " ' and l.remark not like '%Training%' "
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
                    Case "Z"
                        Ws.Cells(LineZ, 8) = mSQLReader.Item("Count1")
                    Case "5"
                        Ws.Cells(LineZ, 9) = mSQLReader.Item("Count1")
                    Case "6"
                        Ws.Cells(LineZ, 10) = mSQLReader.Item("Count1")
                    Case "7"
                        Ws.Cells(LineZ, 11) = mSQLReader.Item("Count1")
                    Case "8"
                        Ws.Cells(LineZ, 12) = mSQLReader.Item("Count1")
                    Case "9"
                        Ws.Cells(LineZ, 13) = mSQLReader.Item("Count1")
                    Case "A"
                        Ws.Cells(LineZ, 14) = mSQLReader.Item("Count1")
                    Case "B"
                        Ws.Cells(LineZ, 28) = mSQLReader.Item("Count1")
                    Case "C"
                        Ws.Cells(LineZ, 15) = mSQLReader.Item("Count1")
                    Case "Y"
                        Ws.Cells(LineZ, 16) = mSQLReader.Item("Count1")
                    Case "D"
                        Ws.Cells(LineZ, 17) = mSQLReader.Item("Count1")
                    Case "E"
                        Ws.Cells(LineZ, 18) = mSQLReader.Item("Count1")
                    Case "F"
                        Ws.Cells(LineZ, 19) = mSQLReader.Item("Count1")
                    Case "G"
                        Ws.Cells(LineZ, 20) = mSQLReader.Item("Count1")
                    Case "H"
                        Ws.Cells(LineZ, 21) = mSQLReader.Item("Count1")
                    Case "I"
                        Ws.Cells(LineZ, 22) = mSQLReader.Item("Count1")
                    Case "J"
                        Ws.Cells(LineZ, 23) = mSQLReader.Item("Count1")
                    Case "K"
                        Ws.Cells(LineZ, 24) = mSQLReader.Item("Count1")
                    Case "L"
                        Ws.Cells(LineZ, 25) = mSQLReader.Item("Count1")
                    Case "M"
                        Ws.Cells(LineZ, 26) = mSQLReader.Item("Count1")
                    Case "N"
                        Ws.Cells(LineZ, 27) = mSQLReader.Item("Count1")
                    Case "O"
                        Ws.Cells(LineZ, 29) = mSQLReader.Item("Count1")
                    Case "P"
                        Ws.Cells(LineZ, 30) = mSQLReader.Item("Count1")
                    Case "Q"
                        Ws.Cells(LineZ, 31) = mSQLReader.Item("Count1")
                    Case "R"
                        Ws.Cells(LineZ, 32) = mSQLReader.Item("Count1")
                    Case "S"
                        Ws.Cells(LineZ, 33) = mSQLReader.Item("Count1")
                    Case "T"
                        Ws.Cells(LineZ, 34) = mSQLReader.Item("Count1")
                    Case "U"
                        Ws.Cells(LineZ, 35) = mSQLReader.Item("Count1")
                    Case "V"
                        Ws.Cells(LineZ, 36) = mSQLReader.Item("Count1")
                    Case "W"
                        Ws.Cells(LineZ, 37) = mSQLReader.Item("Count1")
                    Case "X"
                        Ws.Cells(LineZ, 39) = mSQLReader.Item("Count1")
                    Case "~"
                        Ws.Cells(LineZ, 38) = mSQLReader.Item("Count1")
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
        oRng = Ws.Range("D1", "AQ1")
        oRng.EntireColumn.ColumnWidth = 15
        oRng = Ws.Range("A2", "AM3")
        oRng.Interior.Color = Color.LightBlue
        Ws.Cells(1, 3) = "合计数量："
        'Ws.Cells(1, 1) = "WIP_STATUS"
        Ws.Cells(2, 1) = "ERP 料号"
        Ws.Cells(3, 1) = "ERP PN"
        'Ws.Cells(1, 2) = "ERP PN"
        Ws.Cells(2, 2) = "产品名称"
        Ws.Cells(3, 2) = "Product name"
        Ws.Cells(2, 3) = "产品名称"
        Ws.Cells(3, 3) = "WIP_Product Description"
        'Ws.Cells(1, 3) = "裁纱 Cutting"
        'Ws.Cells(1, 4) = "备料 Prepreg"
        Ws.Cells(2, 4) = "标签"
        Ws.Cells(3, 4) = "Label"
        'Ws.Cells(1, 5) = "预型 Layup"
        Ws.Cells(2, 5) = "裁紗/备料"
        Ws.Cells(3, 5) = "Cutting/Prepreg"
        'Ws.Cells(1, 6) = "成型 Molding"
        Ws.Cells(2, 6) = "预型"
        Ws.Cells(3, 6) = "Layup"
        'Ws.Cells(1, 7) = "后加温 Solidify"
        Ws.Cells(2, 7) = "成型"
        Ws.Cells(3, 7) = "Molding"
        'Ws.Cells(1, 8) = "CNC"
        Ws.Cells(2, 8) = "成型检验"
        Ws.Cells(3, 8) = "Molding Inspection"
        Ws.Cells(2, 9) = "成型固化"
        Ws.Cells(3, 9) = "POST CURING"
        'Ws.Cells(1, 9) = "喷砂 Sand blasting"
        Ws.Cells(2, 10) = "CNC"
        Ws.Cells(3, 10) = "CNC"
        'Ws.Cells(1, 10) = "胶合1 Gluing 1"
        Ws.Cells(2, 11) = "CNC 检验"
        Ws.Cells(3, 11) = "CNC Inspection"
        'Ws.Cells(1, 11) = "胶合2 Gluing 2"
        Ws.Cells(2, 12) = "喷砂"
        Ws.Cells(3, 12) = "Sand blasting"
        'Ws.Cells(1, 12) = "胶合3 Gluing 3"
        Ws.Cells(2, 13) = "胶合1"
        Ws.Cells(3, 13) = "Gluing 1"
        'Ws.Cells(1, 13) = "补土1 Sanding 1"
        Ws.Cells(2, 14) = "胶合2&3"
        Ws.Cells(3, 14) = "Gluing 2&3"
        'Ws.Cells(1, 14) = "补土2 Sanding 2"
        'Ws.Cells(2, 15) = "胶合3"
        'Ws.Cells(3, 15) = "Gluing 3"
        'Ws.Cells(1, 15) = "补土3 Sanding 3"
        Ws.Cells(2, 15) = "胶合检验1"
        Ws.Cells(3, 15) = "Inspection Gluing 1"
        Ws.Cells(2, 16) = "胶合检验2&3"
        Ws.Cells(3, 16) = "Inspection Gluing 2&3"
        'Ws.Cells(1, 16) = "补土4 Sanding 4"
        Ws.Cells(2, 17) = "补土1"
        Ws.Cells(3, 17) = "Sanding 1"
        'Ws.Cells(1, 17) = "涂装1 Painting 1"
        Ws.Cells(2, 18) = "点补&清洁1"
        Ws.Cells(3, 18) = "Replenish&CleaningⅠ"
        'Ws.Cells(1, 18) = "涂装2 Painting 2"
        Ws.Cells(2, 19) = "补土2"
        Ws.Cells(3, 19) = "Sanding 2"
        'Ws.Cells(1, 19) = "涂装3 Painting 3"
        Ws.Cells(2, 20) = "补土3"
        Ws.Cells(3, 20) = "Sanding 3"
        'Ws.Cells(1, 20) = "涂装4 Painting 4"
        Ws.Cells(2, 21) = "补土4"
        Ws.Cells(3, 21) = "Sanding 4"
        'Ws.Cells(1, 21) = "抛光 Polishing"
        Ws.Cells(2, 22) = "补土5"
        Ws.Cells(3, 22) = "Sanding 5"
        'Ws.Cells(1, 22) = "静置 Waiting"
        Ws.Cells(2, 23) = "补土6"
        Ws.Cells(3, 23) = "Sanding 6"
        'Ws.Cells(1, 23) = "包装 Packing"
        Ws.Cells(2, 24) = "涂装1"
        Ws.Cells(3, 24) = "Painting 1"
        'Ws.Cells(1, 24) = "成品 FG"
        Ws.Cells(2, 25) = "涂装2"
        Ws.Cells(3, 25) = "Painting 2"
        'Ws.Cells(1, 25) = "加总 Total"
        Ws.Cells(2, 26) = "涂装3"
        Ws.Cells(3, 26) = "Painting 3"
        Ws.Cells(2, 27) = "涂装4"
        Ws.Cells(3, 27) = "Painting 4"
        Ws.Cells(2, 28) = "贴水标"
        Ws.Cells(3, 28) = "Apply Decal"
        Ws.Cells(2, 29) = "涂装5"
        Ws.Cells(3, 29) = "Painting 5"
        Ws.Cells(2, 30) = "涂装6"
        Ws.Cells(3, 30) = "Painting 6"
        Ws.Cells(2, 31) = "涂装检验"
        Ws.Cells(3, 31) = "Painting Finish"
        Ws.Cells(2, 32) = "抛光"
        Ws.Cells(3, 32) = "Polishing"
        Ws.Cells(2, 33) = "抛光检验"
        Ws.Cells(3, 33) = "Polishing Inspection"
        Ws.Cells(2, 34) = "静置/组配"
        Ws.Cells(3, 34) = "Waiting & Assembly"
        Ws.Cells(2, 35) = "包装"
        Ws.Cells(3, 35) = "Packing"
        Ws.Cells(2, 36) = "成品仓"
        Ws.Cells(3, 36) = "FG"
        Ws.Cells(2, 37) = "隔離/待判"
        Ws.Cells(3, 37) = "Block / Hold"
        Ws.Cells(2, 38) = "呆滞品"
        Ws.Cells(3, 38) = "Dead stock"
        Ws.Cells(3, 39) = "B FG"
        Ws.Cells(2, 39) = "0799站"
        Ws.Cells(2, 40) = "产品加总"
        Ws.Cells(3, 40) = "Total WIP"
        Ws.Cells(2, 41) = "总数量"
        Ws.Cells(3, 41) = "Total Qty"
        Ws.Cells(2, 42) = "胶合后"
        Ws.Cells(3, 42) = "Gluing After"
        Ws.Cells(2, 43) = "涂装后包装前"
        Ws.Cells(3, 43) = "After Paint Before Pack"
    End Sub


    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "WIP_STATUS"
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
    Private Sub SumRow()
        Ws.Cells(LineZ, 40) = "=SUM(H" & LineZ & ":AM" & LineZ & ")"
        Ws.Cells(LineZ, 41) = "=SUM(E" & LineZ & ":AM" & LineZ & ")"
        Ws.Cells(LineZ, 42) = "=SUM(N" & LineZ & ":AM" & LineZ & ")"
        Ws.Cells(LineZ, 43) = "=SUM(AE" & LineZ & ":AH" & LineZ & ")" '+SUM(AA" & LineZ & ":AD" & LineZ & ")"
    End Sub
    Private Sub SumColumn()
        Ws.Cells(1, 4) = "=SUM(D4:D" & LineZ & ")"
        oRng = Ws.Range("D1", "D1")
        oRng.AutoFill(Destination:=Ws.Range("D1", "AQ1"), Type:=xlFillDefault)
    End Sub
End Class