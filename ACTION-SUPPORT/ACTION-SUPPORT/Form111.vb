Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form111
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim ptime As String = String.Empty
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim tModel As String
    Dim DefectCode As String() = {}
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form111_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
                mSQLS1.CommandTimeout = 600
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        Dim Model_Type As String = String.Empty
        BindModel(Model_Type)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Dim xPath As String = "C:\temp\产出与质量状态统计表.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mSQLS1.CommandTimeout = 600
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If

        TimeS1 = DateTimePicker1.Value
        TimeS2 = DateTimePicker2.Value

        tModel = String.Empty
        If Not IsNothing(ComboBox2.SelectedItem) Then
            tModel = ComboBox2.SelectedItem.ToString()
        End If
        BackgroundWorker1.RunWorkerAsync()
        'ExportToExcel()
        'SaveExcel()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "产出与质量状态统计表"
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
        Dim xPath As String = "C:\temp\产出与质量状态统计表.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        AdjustmentExcelFormat()
        LineZ = 4

        mSQLS1.CommandText = "select value,model,modelname,sum(s1) as s1,sum(s2) as s2,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4 "
        For i As Int16 = 0 To DefectCode.Length - 1 Step 1
            mSQLS1.CommandText += ",sum(d" & i & ") as d" & i
        Next
        mSQLS1.CommandText += " from ( "
        mSQLS1.CommandText += "select value,lot.model,model.modelname,0 as s1,0 as s2,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation in ('0050','0055','0080','0090','0100','0110','0111','0112','0120','0130','0140','0150','0165','0170','0175','0177','0180','0190','0193','0195','0200','0210','0215','0220','0223','0225','0230','0315','0320','0325','0330') then 1 else 0 end ) as t1,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation in ('0113','0151','0160','0172','0173','0174','0231','0240','0250','0255','0260','0280','0300','0316','0321','0326','0331','0333') then 1 else 0 end ) as t2,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation in ('0335','0340','0350','0360','0370','0380','0395','0400','0405','0410','0415','0417','0418','0420','0430','0435','0440','0445','0450','0475','0478','0480','0490','0492','0493','0495','0500','0510','0520','0530','0605','0610','0620','0623','0627') then 1 else 0 end) as t3,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation in ('0460','0465','0470','0540','0545','0550','0560','0563','0567','0570','0575','0580','0583','0584','0585','0587','0590','0591','0592','0595','0629','0630','0633','0635','0640','0642','0645','0649','0650','0652','0657','0658','0659','0660','0665','0666','0667','0668','0670','0673','0675','0680','0690') then 1 else 0 end ) as t4 "
        For i As Int16 = 0 To DefectCode.Length - 1 Step 1
            mSQLS1.CommandText += ",(case when scrap.defect = '" & DefectCode(i).ToString() & "' then 1 else 0 end ) as d" & i
        Next
        mSQLS1.CommandText += " from scrap left join scrap_sn on scrap.sn = scrap_sn.sn left join lot on scrap.lot = lot.lot "
        mSQLS1.CommandText += "left join model on lot.model = model.model left join model_paravalue on lot.model = model_paravalue.model  and model_paravalue.parameter = 'ERP PN' "
        mSQLS1.CommandText += "where scrap.datetime between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select value,lot.model,model.modelname,(case when station = '0330' then 1 else 0 end) as s1,(case when station = '0331' then 1 else 0 end) as s2,0,0,0,0"
        For i As Int16 = 0 To DefectCode.Length - 1 Step 1
            mSQLS1.CommandText += ",0"
        Next
        mSQLS1.CommandText += " from tracking left join lot on tracking.lot = lot.lot left join model on lot.model = model.model "
        mSQLS1.CommandText += "left join model_paravalue on lot.model = model_paravalue.model  and model_paravalue.parameter = 'ERP PN' "
        mSQLS1.CommandText += "where tracking.timeout between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and tracking.station in ('0330','0331') "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select value,lot.model,model.modelname,(case when station = '0330' then 1 else 0 end) as s1,(case when station = '0331' then 1 else 0 end) as s2,0,0,0,0"
        For i As Int16 = 0 To DefectCode.Length - 1 Step 1
            mSQLS1.CommandText += ",0"
        Next
        mSQLS1.CommandText += " from tracking_dup left join lot on tracking_dup.lot = lot.lot left join model on lot.model = model.model "
        mSQLS1.CommandText += "left join model_paravalue on lot.model = model_paravalue.model  and model_paravalue.parameter = 'ERP PN' "
        mSQLS1.CommandText += "where tracking_dup.timeout between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and tracking_dup.station in ('0330','0331') "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select value,lot.model,model.modelname,(case when station = '0330' then 1 else 0 end) as s1,(case when station = '0331' then 1 else 0 end) as s2,0,0,0,0"
        For i As Int16 = 0 To DefectCode.Length - 1 Step 1
            mSQLS1.CommandText += ",0"
        Next
        mSQLS1.CommandText += " from scrap_tracking left join lot on scrap_tracking.lot = lot.lot left join model on lot.model = model.model "
        mSQLS1.CommandText += "left join model_paravalue on lot.model = model_paravalue.model  and model_paravalue.parameter = 'ERP PN' "
        mSQLS1.CommandText += "where scrap_tracking.timein between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and scrap_tracking.station in ('0330','0331') ) AS AB "
        'If Not String.IsNullOrEmpty(tModel_type) Then
        'mSQLS1.CommandText += "and m.model_type like '" & tModel_type & "' "
        'End If
        If Not String.IsNullOrEmpty(tModel) Then
            mSQLS1.CommandText += " WHERE model like '" & tModel & "' "
        End If
        mSQLS1.CommandText += " group by value,model,modelname"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("value")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("modelname")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("s1")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("s2")
                Ws.Cells(LineZ, 6) = mSQLReader.Item("t1")
                Ws.Cells(LineZ, 7) = mSQLReader.Item("t2")
                Ws.Cells(LineZ, 9) = mSQLReader.Item("t3")
                Ws.Cells(LineZ, 10) = mSQLReader.Item("t4")
                For i As Int16 = 0 To DefectCode.Length - 1 Step 1
                    If mSQLReader.Item(9 + i) <> 0 Then
                        Ws.Cells(LineZ, 13 + i) = mSQLReader.Item(9 + i)
                    End If
                Next
                LineZ += 1
                Me.Label3.Text = LineZ
            End While
        End If
        mSQLReader.Close()
        ' 第二頁
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustmentExcelFormat1()
        LineZ = 4
        mSQLS1.CommandText = "select value,model,modelname,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,"
        mSQLS1.CommandText += "sum(f1) as f1,sum(f2) as f2,sum(f3) as f3,sum(f4) as f4,sum(f5) as f5,sum(f6) as f6,sum(f7) as f7 "
        For i As Int16 = 0 To DefectCode.Length - 1 Step 1
            mSQLS1.CommandText += ",sum(d" & i & ") as d" & i
        Next
        mSQLS1.CommandText += " from ( select value,lot.model,model.modelname,(case when station = '0490' then 1 else 0 end) as t1,"
        mSQLS1.CommandText += "(case when station = '0590' then 1 else 0 end) as t2,(case when station = '0640' then 1 else 0 end) as t3,"
        mSQLS1.CommandText += "(case when station = '0670' then 1 else 0 end) as t4,0 as f1,0 as f2,0 as f3,0 as f4,0 as f5,0 as f6,0 as f7"
        For i As Int16 = 0 To DefectCode.Length - 1 Step 1
            mSQLS1.CommandText += ",0 as d" & i
        Next
        mSQLS1.CommandText += " from tracking left join lot on tracking.lot = lot.lot left join model on lot.model = model.model "
        mSQLS1.CommandText += "left join model_paravalue on lot.model = model_paravalue.model  and model_paravalue.parameter = 'ERP PN' "
        mSQLS1.CommandText += "where tracking.timeout between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS1.CommandText += "and tracking.station in ('0490','0590','0640','0670') "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select value,lot.model,model.modelname,(case when station = '0490' then 1 else 0 end) as t1,"
        mSQLS1.CommandText += "(case when station = '0590' then 1 else 0 end) as t2,(case when station = '0640' then 1 else 0 end) as t3,"
        mSQLS1.CommandText += "(case when station = '0670' then 1 else 0 end) as t4,0,0,0,0,0,0,0"
        For i As Int16 = 0 To DefectCode.Length - 1 Step 1
            mSQLS1.CommandText += ",0"
        Next
        mSQLS1.CommandText += "from tracking_dup left join lot on tracking_dup.lot = lot.lot left join model on lot.model = model.model "
        mSQLS1.CommandText += "left join model_paravalue on lot.model = model_paravalue.model  and model_paravalue.parameter = 'ERP PN' "
        mSQLS1.CommandText += "where tracking_dup.timeout between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS1.CommandText += "and tracking_dup.station in ('0490','0590','0640','0670') "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select value,lot.model,model.modelname,(case when station = '0490' then 1 else 0 end) as t1,"
        mSQLS1.CommandText += "(case when station = '0590' then 1 else 0 end) as t2,(case when station = '0640' then 1 else 0 end) as t3,"
        mSQLS1.CommandText += "(case when station = '0670' then 1 else 0 end) as t4,0,0,0,0,0,0,0"
        For i As Int16 = 0 To DefectCode.Length - 1 Step 1
            mSQLS1.CommandText += ",0"
        Next
        mSQLS1.CommandText += "from scrap_tracking left join lot on scrap_tracking.lot = lot.lot left join model on lot.model = model.model "
        mSQLS1.CommandText += "left join model_paravalue on lot.model = model_paravalue.model  and model_paravalue.parameter = 'ERP PN' "
        mSQLS1.CommandText += "where scrap_tracking.timeout between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS1.CommandText += "and scrap_tracking.station in ('0490','0590','0640','0670') "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select value,lot.model,model.modelname,0,0,0,0,(case when failstation in ('0490','0620','0627') then 1 else 0 end),"
        mSQLS1.CommandText += "(case when failstation in ('0590','0595','0640','0659','0670') and rework in ('0478','0480','0605') then 1 else 0 end ),(case when failstation in ('0590','0595') then 1 else 0 end ),"
        mSQLS1.CommandText += "(case when failstation in ('0475','0590','0595','0640','0645','0659','0670') and rework in ('0410','0440','0455','0460','0540','0567','0570','0583') then 1 else 0 end),(case when failstation in ('0640','0645') then 1 else 0 end),"
        mSQLS1.CommandText += "(case when failstation in ('0659','0670') and rework in ('0629','0630','0633') then 1 else 0 end),(case when failstation in ('0670') then 1 else 0 end )"
        For i As Int16 = 0 To DefectCode.Length - 1 Step 1
            mSQLS1.CommandText += ",(case when failure.defect = '" & DefectCode(i).ToString() & "' then 1 else 0 end ) as d" & i
        Next
        mSQLS1.CommandText += " from failure left join lot on failure.lot = lot.lot left join model on lot.model = model.model "
        mSQLS1.CommandText += "left join model_paravalue on lot.model = model_paravalue.model  and model_paravalue.parameter = 'ERP PN' "
        mSQLS1.CommandText += "where failtime between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS1.CommandText += "and rework <> 'SCRP' "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select value,lot.model,model.modelname,0,0,0,0,(case when failstation in ('0490','0620','0627') then 1 else 0 end),"
        mSQLS1.CommandText += "(case when failstation in ('0590','0595','0640','0659','0670') and rework in ('0478','0480','0605') then 1 else 0 end ),(case when failstation in ('0590','0595') then 1 else 0 end ),"
        mSQLS1.CommandText += "(case when failstation in ('0475','0590','0595','0640','0645','0659','0670') and rework in ('0410','0440','0455','0460','0540','0567','0570','0583') then 1 else 0 end),(case when failstation in ('0640','0645') then 1 else 0 end),"
        mSQLS1.CommandText += "(case when failstation in ('0659','0670') and rework in ('0629','0630','0633') then 1 else 0 end),(case when failstation in ('0670') then 1 else 0 end )"
        For i As Int16 = 0 To DefectCode.Length - 1 Step 1
            mSQLS1.CommandText += ",(case when scrap_failure.defect = '" & DefectCode(i).ToString() & "' then 1 else 0 end ) as d" & i
        Next
        mSQLS1.CommandText += " from scrap_failure left join lot on scrap_failure.lot = lot.lot left join model on lot.model = model.model "
        mSQLS1.CommandText += "left join model_paravalue on lot.model = model_paravalue.model  and model_paravalue.parameter = 'ERP PN' "
        mSQLS1.CommandText += "where failtime between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS1.CommandText += "and rework <> 'SCRP' ) AS ab "
        If Not String.IsNullOrEmpty(tModel) Then
            mSQLS1.CommandText += " WHERE model like '" & tModel & "' "
        End If
        mSQLS1.CommandText += " group by value,model,modelname"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("value")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("modelname")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("t1")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("f1")
                Ws.Cells(LineZ, 6) = mSQLReader.Item("f2")
                Ws.Cells(LineZ, 8) = mSQLReader.Item("t2")
                Ws.Cells(LineZ, 9) = mSQLReader.Item("f3")
                Ws.Cells(LineZ, 10) = mSQLReader.Item("f4")
                Ws.Cells(LineZ, 12) = mSQLReader.Item("t3")
                Ws.Cells(LineZ, 13) = mSQLReader.Item("f5")
                Ws.Cells(LineZ, 14) = mSQLReader.Item("f6")
                Ws.Cells(LineZ, 16) = mSQLReader.Item("t4")
                Ws.Cells(LineZ, 17) = mSQLReader.Item("f7")
                For i As Int16 = 0 To DefectCode.Length - 1 Step 1
                    If mSQLReader.Item(14 + i) <> 0 Then
                        Ws.Cells(LineZ, 20 + i) = mSQLReader.Item(14 + i)
                    End If
                Next
                LineZ += 1
                Me.Label3.Text = LineZ
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub AdjustmentExcelFormat()
        mSQLS1.CommandText = "select count(*) from  defect"
        Dim DL As Int16 = mSQLS1.ExecuteScalar()
        Array.Resize(DefectCode, DL)
        mSQLS1.CommandText = "select * from defect order by defect "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            Dim ColumnX As Int16 = 0
            While mSQLReader.Read()
                Ws.Cells(1, 13 + ColumnX) = mSQLReader.Item("defect").ToString()
                Ws.Cells(3, 13 + ColumnX) = mSQLReader.Item("desc_en")
                DefectCode(ColumnX) = mSQLReader.Item("defect")
                ColumnX += 1
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub AdjustmentExcelFormat1()
        mSQLS1.CommandText = "select count(*) from  defect"
        Dim DL As Int16 = mSQLS1.ExecuteScalar()
        Array.Clear(DefectCode, 0, DL)
        Array.Resize(DefectCode, DL)
        mSQLS1.CommandText = "select * from defect order by defect "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            Dim ColumnX As Int16 = 0
            While mSQLReader.Read()
                Ws.Cells(1, 20 + ColumnX) = mSQLReader.Item("defect").ToString()
                Ws.Cells(3, 20 + ColumnX) = mSQLReader.Item("desc_en")
                DefectCode(ColumnX) = mSQLReader.Item("defect")
                ColumnX += 1
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
                Me.ComboBox2.Items.Add(mSQLReader.Item(0).ToString())
            End While
        End If
        mSQLReader.Close()
    End Sub
End Class