Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form47
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim tModel_type As String
    Dim tModel As String
    Dim tStation As String
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form47_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        tModel_type = String.Empty
        tModel = String.Empty
        tStation = String.Empty
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
            'Else
            'MsgBox("请选择产品")
            'Return
        End If
        If Not IsNothing(ComboBox3.SelectedItem) Then
            tStation = ComboBox3.SelectedItem.ToString()
        End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        mSQLS1.CommandText = "select count(sn) from sn left join lot on sn.lot = lot.lot left join station on sn.updatedstation = station.station "
        'mSQLS1.CommandText += "where lot.model = '" & tModel & "' and sn.updatedstation in ('0590','0630','0635','0640','0642','0645','0650','0657','0665','0670','0675','0680','0690','0730') and sn.topreworkstation is null "
        'mSQLS1.CommandText += "where sn.updatedstation in ('0590','0629','0630','0635','0640','0642','0645','0650','0657','0658','0659','0665','0670','0673','0675','0680','0690','0730') and sn.topreworkstation is null "

        '191109 add by Brady
        'mSQLS1.CommandText += "where sn.updatedstation in ('0590','0625','0629','0630','0635','0640','0642','0645','0650','0657','0658','0659','0665','0669','0670','0673','0675','0680','0690','0730') and sn.topreworkstation is null "
        mSQLS1.CommandText += "where sn.updatedstation in ('0590','0625','0629','0630','0635','0640','0642','0645','0650','0657','0658','0659','0665','0669','0670','0673','0675','0680','0690','0730','0587','0674','0669') and sn.topreworkstation is null "
        '191109 add by Brady END

        If Not String.IsNullOrEmpty(tModel) Then
            mSQLS1.CommandText += " and lot.model = '" & tModel & "'"
        End If
        If Not String.IsNullOrEmpty(tStation) Then
            mSQLS1.CommandText += " and sn.updatedstation = '" & tStation & "'"
        End If
        Dim TotC As Decimal = mSQLS1.ExecuteScalar()
        Me.Label4.Text = TotC

        mSQLS1.CommandText = "select sn.lot,sn.sn,sn.updatedstation,station.stationname from sn left join lot on sn.lot = lot.lot left join station on sn.updatedstation = station.station "
        'mSQLS1.CommandText += "where lot.model = '" & tModel & "' and sn.updatedstation in ('0590','0630','0635','0640','0642','0645','0650','0657','0665','0670','0675','0680','0690','0730') and sn.topreworkstation is null "
        'mSQLS1.CommandText += "where sn.updatedstation in ('0590','0629','0630','0635','0640','0642','0645','0650','0657','0658','0659','0665','0670','0673','0674','0675','0680','0690','0730') and sn.topreworkstation is null "

        '191109 add by Brady
        'mSQLS1.CommandText += "where sn.updatedstation in ('0590','0625','0629','0630','0635','0640','0642','0645','0650','0657','0658','0659','0665','0669','0670','0673','0674','0675','0680','0690','0730') and sn.topreworkstation is null "
        mSQLS1.CommandText += "where sn.updatedstation in ('0590','0625','0629','0630','0635','0640','0642','0645','0650','0657','0658','0659','0665','0669','0670','0673','0674','0675','0680','0690','0730','0587','0674','0669') and sn.topreworkstation is null "
        '191109 add by Brady END

        If Not String.IsNullOrEmpty(tModel) Then
            mSQLS1.CommandText += " and lot.model = '" & tModel & "'"
        End If
        If Not String.IsNullOrEmpty(tStation) Then
            mSQLS1.CommandText += " and sn.updatedstation = '" & tStation & "'"
        End If
        Dim nowC As Decimal = 1
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("sn")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("updatedstation") & " " & mSQLReader.Item("stationname")
                Ws.Cells(LineZ, 3) = Get0665Hours(mSQLReader.Item("lot"), mSQLReader.Item("sn"))
                LineZ += 1
                nowC += 1
                Label5.Text = nowC
            End While
        End If
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        oRng = Ws.Range("A1", "C1")
        oRng.EntireColumn.ColumnWidth = 40
        oRng.Merge()
        Ws.Cells(1, 1) = tModel & " 产品静置剩余时间列表"
        Ws.Cells(2, 3) = "报表生成时间：" & Now.ToString("yyyy-MM-dd HH:mm")
        Ws.Cells(3, 1) = "产品系列号"
        Ws.Cells(3, 2) = "当前工站"
        Ws.Cells(3, 3) = "静置剩余时间（H)"
        LineZ = 4
    End Sub
    Private Function Get0665Hours(ByVal lot As String, ByVal sn As String)
        Dim mSQLS2 As New SqlClient.SqlCommand
        mSQLS2.Connection = mConnection
        mSQLS2.CommandType = CommandType.Text
        mSQLS2.CommandText = "SELECT value FROM LOT LEFT JOIN model_paravalue ON lot.model = model_paravalue.model and parameter = 'Polish72' "
        mSQLS2.CommandText += "WHERE LOT = '" & lot & "'"
        Dim P72 As String = mSQLS2.ExecuteScalar()
        If P72 = "Y" Then
            mSQLS2.CommandText = "select isnull(72 -"
        Else
            mSQLS2.CommandText = "select isnull(36 -"
        End If

        mSQLS2.CommandText += " (datediff(HH,MAX(t1),getdate())),0) from ( "
        mSQLS2.CommandText += "select isnull(timeout,timein) as t1 from tracking,station where tracking.station = station.station and tracking.sn = '"
        mSQLS2.CommandText += sn & "' and tracking.lot = '" & lot & "' and station.stationname like 'paint%' "
        mSQLS2.CommandText += "union all "
        mSQLS2.CommandText += "select isnull(timeout,timein) as t1 from tracking_dup,station where tracking_dup.station = station.station and tracking_dup.sn = '"
        mSQLS2.CommandText += sn & "' and tracking_dup.lot = '" & lot & "' and station.stationname like 'paint%' ) as tt"
        Dim LastHour As Decimal = mSQLS2.ExecuteScalar()
        If LastHour < 0 Then
            Return 0
        Else
            Return LastHour
        End If
    End Function
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Wait_Hours_Report"
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
End Class