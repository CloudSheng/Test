Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form137
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim ptime As String = String.Empty
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    'Dim tModel_type = String.Empty
    Dim tModel As String = String.Empty
    Dim LineZ As Integer = 0

    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form137_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        'BindModel_Type()
        BindModel(tModel)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Dim xPath As String = "C:\temp\DAC产品直通率报表.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If

        'tModel_type = String.Empty
        tModel = String.Empty
        'If Not IsNothing(ComboBox1.SelectedItem) Then
        'tModel_type = ComboBox1.SelectedItem.ToString()
        'End If
        If Not IsNothing(ComboBox2.SelectedItem) Then
            tModel = ComboBox2.SelectedItem.ToString()
            Dim stCount As Int16 = Strings.InStr(tModel, "|")
            If stCount > 0 Then
                tModel = Strings.Left(tModel, stCount - 1)
            End If
        End If
        TimeS1 = DateTimePicker1.Value
        TimeS2 = DateTimePicker2.Value
        BackgroundWorker1.RunWorkerAsync()
        'ExportToExcel()
        'SaveExcel()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "DAC产品直通率报表"
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
                'mConnection.Close()
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
        Dim xPath As String = "C:\temp\DAC产品直通率报表.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        LineZ = 6
        Ws.Cells(3, 1) = "取数日期/时间：" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "-" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss")

        mSQLS1.CommandText = "select model,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,"
        mSQLS1.CommandText += "sum(t12) as t12,sum(t13) as t13,sum(t14) as t14,sum(t15) as t15,sum(t16) as t16,sum(t17) as t17,sum(t18) as t18,sum(t19) as t19,sum(t20) as t20,sum(t21) as t21,sum(t22) as t22,"
        mSQLS1.CommandText += "sum(t23) as t23,sum(t24) as t24,sum(t25) as t25,sum(t26) as t26,sum(t27) as t27,sum(t28) as t28,sum(t29) as t29,sum(t30) as t30,sum(t31) as t31,sum(t32) as t32,sum(t33) as t33,"
        mSQLS1.CommandText += "sum(t34) as t34,sum(t35) as t35,sum(t36) as t36,sum(t37) as t37,sum(t38) as t38,sum(t39) as t39,sum(t40) as t40,sum(t41) as t41,sum(t42) as t42,sum(t43) as t43,sum(t44) as t44,"
        mSQLS1.CommandText += "sum(t45) as t45,sum(t46) as t46,sum(t47) as t47,sum(t48) as t48,sum(t49) as t49,sum(t50) as t50,sum(t51) as t51,sum(t52) as t52,sum(t53) as t53,sum(t54) as t54,sum(t55) as t55,sum(t56) as t56 "
        mSQLS1.CommandText += "from ( select model,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,"
        mSQLS1.CommandText += "sum(t12) as t12,sum(t13) as t13,sum(t14) as t14,sum(t15) as t15,sum(t16) as t16,sum(t17) as t17,sum(t18) as t18,sum(t19) as t19,sum(t20) as t20,sum(t21) as t21,sum(t22) as t22,"
        mSQLS1.CommandText += "sum(t23) as t23,sum(t24) as t24,sum(t25) as t25,sum(t26) as t26,sum(t27) as t27,sum(t28) as t28,0 as t29,0 as t30,0 as t31,0 as t32,0 as t33,0 as t34,0 as t35,0 as t36,0 as t37,0 as t38,0 as t39,0 as t40,0 as t41,0 as t42,"
        mSQLS1.CommandText += "0 as t43,0 as t44,0 as t45,0 as t46,0 as t47,0 as t48,0 as t49,0 as t50,0 as t51,0 as t52,0 as t53,0 as t54,0 as t55,0 as t56 "
        mSQLS1.CommandText += " from ( select model,(case when station = '0670' then count(sn) else 0 end ) as t1,(case when station = '0145' then count(sn) else 0 end ) as t2,"
        mSQLS1.CommandText += "(case when station = '0172' then count(sn) else 0 end ) as t3,(case when station = '0174' then count(sn) else 0 end ) as t4,"
        mSQLS1.CommandText += "(case when station = '0331' then count(sn) else 0 end ) as t5,(case when station = '0330' then count(sn) else 0 end ) as t6,"
        mSQLS1.CommandText += "(case when station = '0380' then count(sn) else 0 end ) as t7,(case when station = '0385' then count(sn) else 0 end ) as t8,"
        mSQLS1.CommandText += "(case when station = '0530' then count(sn) else 0 end ) as t9,(case when station = '0430' then count(sn) else 0 end ) as t10,"
        mSQLS1.CommandText += "(case when station = '0441' then count(sn) else 0 end ) as t11,(case when station = '0461' then count(sn) else 0 end ) as t12,"
        mSQLS1.CommandText += "(case when station = '0475' then count(sn) else 0 end ) as t13,(case when station = '0541' then count(sn) else 0 end ) as t14,"
        mSQLS1.CommandText += "(case when station = '0563' then count(sn) else 0 end ) as t15,(case when station = '0590' then count(sn) else 0 end ) as t16,"
        mSQLS1.CommandText += "(case when station = '0592' then count(sn) else 0 end ) as t17,(case when station = '0595' then count(sn) else 0 end ) as t18,"
        mSQLS1.CommandText += "(case when station = '0490' then count(sn) else 0 end ) as t19,(case when station = '0491' then count(sn) else 0 end ) as t20,"
        mSQLS1.CommandText += "(case when station = '0620' then count(sn) else 0 end ) as t21,(case when station = '0627' then count(sn) else 0 end ) as t22,"
        mSQLS1.CommandText += "(case when station = '0640' then count(sn) else 0 end ) as t23,(case when station = '0645' then count(sn) else 0 end ) as t24,"
        mSQLS1.CommandText += "(case when station = '0667' then count(sn) else 0 end ) as t25,(case when station = '0660' then count(sn) else 0 end ) as t26,"
        mSQLS1.CommandText += "(case when station = '0674' then count(sn) else 0 end ) as t27,(case when station = '0659' then count(sn) else 0 end ) as t28"
        mSQLS1.CommandText += " from ( select distinct sn, lot.model,station from ( select tracking.lot,tracking.sn,station from tracking where tracking.timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and station in ('0670','0145','0172','0174','0331','0330','0380','0385','0530','0430','0441','0461','0475','0541','0563','0590','0592','0595','0490','0491','0620','0627','0640','0645','0667','0660','0674','0659') "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select tracking_dup.lot,tracking_dup.sn,station from tracking_dup left join lot on tracking_dup.lot = lot.lot where tracking_dup.timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and station in ('0670','0145','0172','0174','0331','0330','0380','0385','0530','0430','0441','0461','0475','0541','0563','0590','0592','0595','0490','0491','0620','0627','0640','0645','0667','0660','0674','0659') "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select scrap_tracking.lot,scrap_tracking.sn,station from scrap_tracking where scrap_tracking.timein between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and station in ('0670','0145','0172','0174','0331','0330','0380','0385','0530','0430','0441','0461','0475','0541','0563','0590','0592','0595','0490','0491','0620','0627','0640','0645','0667','0660','0674','0659') "
        mSQLS1.CommandText += ") as Ab left join lot on ab.lot = lot.lot ) as Ac group by ac.model,ac.station ) as AD group by ad.model "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select lot.model,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0, "
        mSQLS1.CommandText += "(case when station = '0670' then count(sn) else 0 end ) as t29,(case when station = '0145' then count(sn) else 0 end ) as t30,"
        mSQLS1.CommandText += "(case when station = '0172' then count(sn) else 0 end ) as t31,(case when station = '0174' then count(sn) else 0 end ) as t32,"
        mSQLS1.CommandText += "(case when station = '0331' then count(sn) else 0 end ) as t33,(case when station = '0330' then count(sn) else 0 end ) as t34,"
        mSQLS1.CommandText += "(case when station = '0380' then count(sn) else 0 end ) as t35,(case when station = '0385' then count(sn) else 0 end ) as t36,"
        mSQLS1.CommandText += "(case when station = '0530' then count(sn) else 0 end ) as t37,(case when station = '0430' then count(sn) else 0 end ) as t38,"
        mSQLS1.CommandText += "(case when station = '0441' then count(sn) else 0 end ) as t39,(case when station = '0461' then count(sn) else 0 end ) as t40,"
        mSQLS1.CommandText += "(case when station = '0475' then count(sn) else 0 end ) as t41,(case when station = '0541' then count(sn) else 0 end ) as t42,"
        mSQLS1.CommandText += "(case when station = '0563' then count(sn) else 0 end ) as t43,(case when station = '0590' then count(sn) else 0 end ) as t44,"
        mSQLS1.CommandText += "(case when station = '0592' then count(sn) else 0 end ) as t45,(case when station = '0595' then count(sn) else 0 end ) as t46,"
        mSQLS1.CommandText += "(case when station = '0490' then count(sn) else 0 end ) as t47,(case when station = '0491' then count(sn) else 0 end ) as t48,"
        mSQLS1.CommandText += "(case when station = '0620' then count(sn) else 0 end ) as t49,(case when station = '0627' then count(sn) else 0 end ) as t50,"
        mSQLS1.CommandText += "(case when station = '0640' then count(sn) else 0 end ) as t51,(case when station = '0645' then count(sn) else 0 end ) as t52,"
        mSQLS1.CommandText += "(case when station = '0667' then count(sn) else 0 end ) as t53,(case when station = '0660' then count(sn) else 0 end ) as t54,"
        mSQLS1.CommandText += "(case when station = '0674' then count(sn) else 0 end ) as t55,(case when station = '0659' then count(sn) else 0 end ) as t56 "
        mSQLS1.CommandText += "from tracking left join lot on tracking.lot = lot.lot where tracking.timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and station in ('0670','0145','0172','0174','0331','0330','0380','0385','0530','0430','0441','0461','0475','0541','0563','0590','0592','0595','0490','0491','0620','0627','0640','0645','0667','0660','0674','0659') "
        mSQLS1.CommandText += "and result = 'P' group by model,station ) as AE "
        'If Not String.IsNullOrEmpty(tModel_type) Then
        'mSQLS1.CommandText += "and model.model_type like '" & tModel_type & "' "
        'End If
        If Not String.IsNullOrEmpty(tModel) Then
            mSQLS1.CommandText += " WHERE model = '" & tModel & "' "
        End If
        mSQLS1.CommandText += " group by model"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            Dim i As Int16 = 1
            While mSQLReader.Read()
                Ws.Cells(LineZ, 2) = i
                Ws.Cells(LineZ, 3) = mSQLReader.Item("model")
                For j As Int16 = 1 To 28 Step 1
                    Ws.Cells(LineZ, 3 * j + 2) = mSQLReader.Item(j)
                    Ws.Cells(LineZ, 3 * j + 3) = mSQLReader.Item(28 + j)
                    If mSQLReader.Item(j) = 0 Then
                        Ws.Cells(LineZ, 3 * j + 4) = 1
                    Else
                        Ws.Cells(LineZ, 3 * j + 4) = Decimal.Divide(mSQLReader.Item(28 + j), mSQLReader.Item(j))
                    End If
                Next
                Ws.Cells(LineZ, 4) = "=G" & LineZ & "*J" & LineZ & "*M" & LineZ & "*P" & LineZ & "*S" & LineZ & "*V" & LineZ & "*Y" & LineZ & "*AB" & LineZ & "*AE" & LineZ & "*AH" & LineZ & "*AK" & LineZ & "*AN" & LineZ & "*AQ" & LineZ & "*AT" & LineZ & "*AW" & LineZ & "*AZ" & LineZ & "*BC" & LineZ & "*BF" & LineZ & "*BI" & LineZ & "*BL" & LineZ & "*BO" & LineZ & "*BR" & LineZ & "*BU" & LineZ & "*BX" & LineZ & "*CA" & LineZ & "*CD" & LineZ & "*CG" & LineZ & "*CJ" & LineZ
                LineZ += 1
                i += 1
            End While

            ' 加總
            Ws.Cells(LineZ, 3) = "报表期间直通率 FTT Rate"
            Ws.Cells(LineZ, 4) = "=G" & LineZ & "*J" & LineZ & "*M" & LineZ & "*P" & LineZ & "*S" & LineZ & "*V" & LineZ & "*Y" & LineZ & "*AB" & LineZ & "*AE" & LineZ & "*AH" & LineZ & "*AK" & LineZ & "*AN" & LineZ & "*AQ" & LineZ & "*AT" & LineZ & "*AW" & LineZ & "*AZ" & LineZ & "*BC" & LineZ & "*BF" & LineZ & "*BI" & LineZ & "*BL" & LineZ & "*BO" & LineZ & "*BR" & LineZ & "*BU" & LineZ & "*BX" & LineZ & "*CA" & LineZ & "*CD" & LineZ & "*CG" & LineZ & "*CJ" & LineZ
            Ws.Cells(LineZ, 5) = "=SUM(E6:E" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 6) = "=SUM(F6:F" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 7) = "=IFERROR(F" & LineZ & "/E" & LineZ & ",1)"
            oRng = Ws.Range(Ws.Cells(LineZ, 5), Ws.Cells(LineZ, 7))
            oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 5), Ws.Cells(LineZ, 88)), Type:=xlFillDefault)

            oRng = Ws.Range("B6", Ws.Cells(LineZ, 88))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        mSQLReader.Close()


    End Sub
    
    Private Function Tqty(ByVal model As String, stations As String)
        mSQLS2.CommandText = "select sum(t1 + t2) from ( select count(distinct sn) as t1,0  as t2 from ( "
        mSQLS2.CommandText += "select sn from tracking left join lot on tracking.lot = lot.lot where tracking.station in (" & stations & ") and lot.model = '" & model & "' and tracking.timeout between '"
        mSQLS2.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS2.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS2.CommandText += "union all "
        mSQLS2.CommandText += "select sn from tracking_dup left join lot on tracking_dup.lot = lot.lot where tracking_dup.station in (" & stations & ") and lot.model = '" & model & "' and tracking_dup.timeout between '"
        mSQLS2.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS2.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS2.CommandText += "union all "
        mSQLS2.CommandText += "select sn from scrap_tracking left join lot on scrap_tracking.lot = lot.lot where scrap_tracking.station in (" & stations & ") and lot.model = '" & model & "' and scrap_tracking.timeout between '"
        mSQLS2.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS2.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' ) as ab "
        mSQLS2.CommandText += "union all "
        mSQLS2.CommandText += "select 0,count(distinct(scrap_sn.sn)) from scrap_sn left join scrap on scrap_sn.sn = scrap.sn left join lot on scrap.lot = lot.lot where scrap_sn.updatedstation in (" & stations & ") and lot.model = '" & model & "' and scrap.datetime between '"
        mSQLS2.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS2.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' ) as AC"
        Dim Sqty As Decimal = mSQLS2.ExecuteScalar()
        Return Sqty
    End Function
    Private Function ScQty(ByVal model As String, stations As String)
        mSQLS2.CommandText = "select count(distinct(scrap_sn.sn)) from scrap_sn left join scrap on scrap_sn.sn = scrap.sn left join lot on scrap.lot = lot.lot where scrap_sn.updatedstation in (" & stations & ") and lot.model = '" & model & "' and scrap.defect not in ('DJ01','DJ02','DL02','DL03','DL04','DL07','DL08') and scrap.datetime between '"
        mSQLS2.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS2.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        Dim Sqty As Decimal = mSQLS2.ExecuteScalar()
        Return Sqty
    End Function
    Private Function FailQty(ByVal model As String, stations As String)
        mSQLS2.CommandText = "select count(sn)  from failure left join lot on failure.lot = lot.lot where rework not in ('BLCK','SCRP') and failtime between '"
        mSQLS2.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '"
        mSQLS2.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and failstation in (" & stations & ")  and lot.model = '" & model & "'"
        Dim Sqty As Decimal = mSQLS2.ExecuteScalar()
        Return Sqty
    End Function
    'Private Sub BindModel_Type()
    '    Me.ComboBox1.Items.Clear()
    '    mSQLS1.CommandText = "SELECT * FROM model_type WHERE model_type <> 'Action'"
    '    mSQLReader = mSQLS1.ExecuteReader()
    '    If mSQLReader.HasRows() Then
    '        While mSQLReader.Read()
    '            Me.ComboBox1.Items.Add(mSQLReader.Item(0).ToString())
    '        End While
    '    End If
    '    mSQLReader.Close()
    'End Sub
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
    'Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)
    '    Dim model_type As String = ComboBox1.Items(ComboBox1.SelectedIndex).ToString()
    '    BindModel(model_type)
    'End Sub
End Class