Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form112
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim tStation1 As String
    Dim ptime As String = String.Empty
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form112_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        ptime = Today.AddDays(-1).ToString("yyyy/MM/dd")
        ptime = ptime & " 08:00:00"
        Me.DateTimePicker1.Value = Convert.ToDateTime(ptime)
        Me.DateTimePicker2.Value = Convert.ToDateTime(ptime).AddDays(1).AddSeconds(-1)
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
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString = "SELECT Week,ERP_PN,PART,WIP_Product_Description,Weekly_Plan_Qty,Last_week_difference_Qty FROM [Sheet1$]"
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)
            Dim DS As Data.DataSet = New DataSet()
            Try
                ExcelAdapater.Fill(DS, "table1")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            Me.DataGridView1.DataSource = DS.Tables("table1")
            CheckStatus()
        End If
    End Sub
    Private Sub CheckStatus()
        If Me.DataGridView1.Rows.Count > 0 Then
            Me.Button2.Enabled = True
            Me.Label3.Text = "已读取来源档"
        Else
            Me.Button2.Enabled = False
            Me.Label3.Text = "未读取来源档"
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If Me.DataGridView1.Rows.Count <= 0 Then
            MsgBox("资料有误")
            Return
        End If
        Dim xPath As String = "C:\temp\Weekly schedule follow up.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If

        TimeS2 = DateTimePicker1.Value
        TimeS1 = DateTimePicker2.Value
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Weekly schedule follow up-MPL"
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
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        Dim xPath As String = "C:\temp\Weekly schedule follow up.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        LineZ = 4
        For i As Integer = 0 To DataGridView1.Rows.Count - 1 Step 1
            If Not IsDBNull(DataGridView1.Rows(i).Cells("ERP_PN").Value) Then
                Ws.Cells(LineZ, 1) = DataGridView1.Rows(i).Cells("Week").Value
                Ws.Cells(LineZ, 2) = DataGridView1.Rows(i).Cells("ERP_PN").Value
                Ws.Cells(LineZ, 3) = DataGridView1.Rows(i).Cells("PART").Value
                Ws.Cells(LineZ, 4) = DataGridView1.Rows(i).Cells("WIP_Product_Description").Value
                Ws.Cells(LineZ, 5) = DataGridView1.Rows(i).Cells("Weekly_Plan_Qty").Value
                Ws.Cells(LineZ, 44) = DataGridView1.Rows(i).Cells("Last_week_difference_Qty").Value
                mSQLS1.CommandText = "select value,model,modelname,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,"
                mSQLS1.CommandText += "sum(s1) as s1,sum(s2) as s2,sum(s3) as s3,sum(s4) as s4,sum(s5) as s5,sum(s6) as s6,sum(s7) as s7,sum(s8) as s8,sum(w1) as w1,sum(w2) as w2,sum(w3) as w3,sum(w4) as w4,"
                mSQLS1.CommandText += "sum(w5) as w5,sum(w6) as w6,sum(w7) as w7,sum(w8) as w8,sum(w9) as w9,sum(w10) as w10,sum(w11) as w11,sum(w12) as w12 from ( "
                mSQLS1.CommandText += "select c.value ,lot.model,model.modelname ,(case when station in ('0150','0151') then 1 else 0 end) as t1, (case when station in ('0330','0331') then 1 else 0 end) as t2,"
                mSQLS1.CommandText += "(case when station in ('0380') then 1 else 0 end) as t3,(case when station in ('0480') then 1 else 0 end) as t4,(case when station in ('0410') then 1 else 0 end) as t5,"
                mSQLS1.CommandText += "(case when station in ('0590') then 1 else 0 end) as t6,(case when station in ('0640') then 1 else 0 end) as t7,(case when station in ('0675') then 1 else 0 end) as t8,"
                mSQLS1.CommandText += "0 as s1,0 as s2,0 as s3,0 as s4,0 as s5,0 as s6,0 as s7,0 as s8,0 as w1,0 as w2,0 as w3,0 as w4,0 as w5,0 as w6, 0 as w7,0 as w8,0 as w9,0 as w10,0 as w11,0 as w12 from tracking "
                mSQLS1.CommandText += "left join lot on tracking.lot = lot.lot left join model on lot.model = model.model LEFT JOIN model_paravalue c on model.model = c.model and c.parameter = 'ERP PN' "
                mSQLS1.CommandText += "where timeout between '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                mSQLS1.CommandText += "and lot.model = '" & DataGridView1.Rows(i).Cells("PART").Value & "' "
                mSQLS1.CommandText += "and station in ('0150','0151','0330','0331','0380','0480','0410','0590','0640','0675') and result = 'P' "
                mSQLS1.CommandText += "union all "
                mSQLS1.CommandText += "select c.value ,lot.model,model.modelname, (case when station in ('0150','0151') then 1 else 0 end) as t1,(case when station in ('0330','0331') then 1 else 0 end) as t2,"
                mSQLS1.CommandText += "(case when station in ('0380') then 1 else 0 end) as t3,(case when station in ('0480') then 1 else 0 end) as t4,(case when station in ('0410') then 1 else 0 end) as t5,"
                mSQLS1.CommandText += "(case when station in ('0590') then 1 else 0 end) as t6,(case when station in ('0640') then 1 else 0 end) as t7,(case when station in ('0675') then 1 else 0 end) as t8,"
                mSQLS1.CommandText += "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 from scrap_tracking left join lot on scrap_tracking.lot = lot.lot left join model on lot.model = model.model "
                mSQLS1.CommandText += "LEFT JOIN model_paravalue c on model.model = c.model and c.parameter = 'ERP PN' "
                mSQLS1.CommandText += "where timeout between '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                mSQLS1.CommandText += "and lot.model = '" & DataGridView1.Rows(i).Cells("PART").Value & "' and station in ('0150','0151','0330','0331','0380','0480','0410','0590','0640','0675') and result = 'P' "
                mSQLS1.CommandText += "union all "
                mSQLS1.CommandText += "select c.value ,lot.model,model.modelname,(case when station in ('0150','0151') then 1 else 0 end) as t1, (case when station in ('0330','0331') then 1 else 0 end) as t2,"
                mSQLS1.CommandText += "(case when station in ('0380') then 1 else 0 end) as t3,(case when station in ('0480') then 1 else 0 end) as t4,(case when station in ('0410') then 1 else 0 end) as t5,"
                mSQLS1.CommandText += "(case when station in ('0590') then 1 else 0 end) as t6,(case when station in ('0640') then 1 else 0 end) as t7,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 from tracking_dup "
                mSQLS1.CommandText += "left join lot on tracking_dup.lot = lot.lot left join model on lot.model = model.model LEFT JOIN model_paravalue c on model.model = c.model and c.parameter = 'ERP PN' "
                mSQLS1.CommandText += "where timeout between '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                mSQLS1.CommandText += "and lot.model = '" & DataGridView1.Rows(i).Cells("PART").Value & "' and station in ('0150','0151','0330','0331','0380','0480','0410','0590','0640','0675') and result = 'P' "
                mSQLS1.CommandText += "union all "
                mSQLS1.CommandText += "select c.value , lot.model,model.modelname , 0,0,0,0,0,0,0,0,(case when updatedstation in ('0150','0151') then 1 else 0 end ),(case when updatedstation in ('0330','0331') then 1 else 0 end ),"
                mSQLS1.CommandText += "(case when updatedstation in ('0380') then 1 else 0 end ),(case when updatedstation in ('0490') then 1 else 0 end ),(case when updatedstation in ('0430','0475') then 1 else 0 end ),"
                mSQLS1.CommandText += "(case when updatedstation in ('0590') then 1 else 0 end ),(case when updatedstation in ('0640') then 1 else 0 end ),(case when updatedstation in ('0670') then 1 else 0 end ),"
                mSQLS1.CommandText += "0,0,0,0,0,0,0,0,0,0,0,0 from scrap left join scrap_sn on scrap.sn =scrap_sn.sn left join lot on scrap_sn.lot = lot.lot left join model on lot.model = model.model "
                mSQLS1.CommandText += "LEFT JOIN model_paravalue c on model.model = c.model and c.parameter = 'ERP PN' "
                mSQLS1.CommandText += "where scrap.datetime between '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                mSQLS1.CommandText += "and lot.model = '" & DataGridView1.Rows(i).Cells("PART").Value & "' "
                mSQLS1.CommandText += "and scrap_sn.updatedstation in ('0150','0151','0330','0331','0380','0490','0430','0475','0590','0640','0670') "
                mSQLS1.CommandText += "union all "
                mSQLS1.CommandText += "select c.value,l.model,m.modelname , 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,(case when t.station in ('0150','0151','0165','0170','0175') then 1 else 0 end),"
                mSQLS1.CommandText += "(case when t.station between '0180' and '0333' or t.station in ('0390','0395') then 1 else 0 end ),(case when t.station in ('0335','0340','0350','0360','0370','0493','0495','0500','0510','0520','0380','0530') then 1 else 0 end),"
                mSQLS1.CommandText += "(case when t.station in ('0400','0478','0480','0605','0610','0623','0490','0492','0620','0627') then 1 else 0 end ),(case when t.station in ('0405','0410','0415','0417','0435','0440','0460','0540','0570','0583') then 1 else 0 end ),"
                mSQLS1.CommandText += "(case when t.station in ('0418','0420','0455','0430','0445','0450','0567','0465','0470','0475','0545','0550','0560','0563','0575','0580','0584','0585','0587','0590','0591','0592','0595','0600') then 1 else 0 end ),"
                mSQLS1.CommandText += "(case when t.station in ('0629','0630','0633','0635','0640','0642','0645','0657') then 1 else 0 end ),"
                mSQLS1.CommandText += "(case when t.station in ('0666','0667') then 1 else 0 end ),"
                mSQLS1.CommandText += "(case when t.station in ('0649','0668') then 1 else 0 end ),"
                mSQLS1.CommandText += "(case when t.station in ('0650','0652','0658','0659','0660','0665','0670','0673') then 1 else 0 end ),"
                mSQLS1.CommandText += "(case when t.station in ('0675','0680','0690') then 1 else 0 end ),"
                mSQLS1.CommandText += "(case when t.station in ('0720','0730') then 1 else 0 end ) FROM lot l JOIN model m ON l.model=m.model JOIN sn s ON l.lot=s.lot JOIN station t ON t.station=case when (s.topreworkstation is not null and s.topreworkstation<>'') then s.topreworkstation else s.updatedstation end "
                mSQLS1.CommandText += "LEFT JOIN model_paravalue c on m.model = c.model and c.parameter = 'ERP PN' WHERE t.station  <> '9999' and l.remark not like '%Training%' and l.model = '" & DataGridView1.Rows(i).Cells("PART").Value & "' ) AS AB group by value,model,modelname order by model"
                mSQLReader = mSQLS1.ExecuteReader()
                If mSQLReader.HasRows() Then
                    While mSQLReader.Read()
                        Ws.Cells(LineZ, 6) = mSQLReader.Item("t1")
                        Ws.Cells(LineZ, 7) = mSQLReader.Item("t2")
                        Ws.Cells(LineZ, 8) = mSQLReader.Item("t3")
                        Ws.Cells(LineZ, 9) = mSQLReader.Item("t4")
                        Ws.Cells(LineZ, 10) = mSQLReader.Item("t5")
                        Ws.Cells(LineZ, 11) = mSQLReader.Item("t6")
                        Ws.Cells(LineZ, 12) = mSQLReader.Item("t7")
                        Ws.Cells(LineZ, 13) = mSQLReader.Item("t8")
                        Ws.Cells(LineZ, 14) = mSQLReader.Item("s1")
                        Ws.Cells(LineZ, 15) = mSQLReader.Item("s2")
                        Ws.Cells(LineZ, 16) = mSQLReader.Item("s3")
                        Ws.Cells(LineZ, 17) = mSQLReader.Item("s4")
                        Ws.Cells(LineZ, 18) = mSQLReader.Item("s5")
                        Ws.Cells(LineZ, 19) = mSQLReader.Item("s6")
                        Ws.Cells(LineZ, 20) = mSQLReader.Item("s7")
                        Ws.Cells(LineZ, 21) = mSQLReader.Item("s8")
                        Ws.Cells(LineZ, 30) = mSQLReader.Item("w1")
                        Ws.Cells(LineZ, 31) = mSQLReader.Item("w2")
                        Ws.Cells(LineZ, 32) = mSQLReader.Item("w3")
                        Ws.Cells(LineZ, 33) = mSQLReader.Item("w4")
                        Ws.Cells(LineZ, 34) = mSQLReader.Item("w5")
                        Ws.Cells(LineZ, 35) = mSQLReader.Item("w6")
                        Ws.Cells(LineZ, 36) = mSQLReader.Item("w7")
                        Ws.Cells(LineZ, 37) = mSQLReader.Item("w8")
                        Ws.Cells(LineZ, 38) = mSQLReader.Item("w9")
                        Ws.Cells(LineZ, 39) = mSQLReader.Item("w10")
                        Ws.Cells(LineZ, 40) = mSQLReader.Item("w11")
                        Ws.Cells(LineZ, 41) = mSQLReader.Item("w12")
                        LineZ += 1
                    End While
                End If
                mSQLReader.Close()
            End If
        Next
    End Sub
End Class