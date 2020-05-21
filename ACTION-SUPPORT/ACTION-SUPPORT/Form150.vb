Public Class Form150
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim mSQLReader2 As SqlClient.SqlDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim l_customer As String = String.Empty
    Dim l_model As String = String.Empty
    Dim l_lot As String = String.Empty
    Dim l_Status As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfRDMes()
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
        BindModel_Type()
        BindModel()
        BindLot()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Dim xPath As String = "C:\temp\RD_Sample_Template.xlsx"
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
        l_customer = String.Empty
        l_model = String.Empty
        l_lot = String.Empty
        l_Status = String.Empty
        If Not IsNothing(ComboBox1.SelectedItem) Then
            l_customer = Me.ComboBox1.SelectedItem.ToString()
        End If
        If Not IsNothing(ComboBox2.SelectedItem) Then
            l_model = Me.ComboBox2.SelectedItem.ToString()
        End If
        If Not IsNothing(ComboBox3.SelectedItem) Then
            l_lot = Me.ComboBox3.SelectedItem.ToString()
        End If
        If Not IsNothing(ComboBox4.SelectedItem) Then
            l_Status = Me.ComboBox4.SelectedItem.ToString()
        End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BindModel_Type()
        Me.ComboBox1.Items.Clear()
        mSQLS1.CommandText = "select distinct model_type from model "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox1.Items.Add(mSQLReader.Item(0).ToString())
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub BindModel()
        Me.ComboBox2.Items.Clear()
        mSQLS1.CommandText = "SELECT model FROM model "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox2.Items.Add(mSQLReader.Item(0).ToString())
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub Bindlot()
        Me.ComboBox3.Items.Clear()
        mSQLS1.CommandText = "SELECT lot FROM lot "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox3.Items.Add(mSQLReader.Item(0).ToString())
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        Dim xPath As String = "C:\temp\RD_Sample_Template.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        LineZ = 6

        mSQLS1.CommandText = "select model.model_type, m1.value as v3,lot.model,l2.value as v1,lot.qty,l1.value as v2, lot.lot, lot.remark ,lot.users, users.name , m2.value as v4 ,lot.datetime , lot.prefix , lot.route, lot.status "
        mSQLS1.CommandText += "from lot left join model on lot.model = model.model left join model_paravalue m1 on lot.model = m1.model and m1.parameter = 'Product Group' "
        mSQLS1.CommandText += "left join lot_paravalue l1 on lot.lot = l1.lot and l1.parameter = 'Shipping Add' left join lot_paravalue l2 on lot.lot = l2.lot and l2.parameter = 'Customer order' "
        mSQLS1.CommandText += "left join users on lot.users = users.id left join model_paravalue m2 on lot.model = m2.model and m2.parameter = 'ERP PN' WHERE 1 =1 "
        If Not String.IsNullOrEmpty(l_customer) Then
            mSQLS1.CommandText += " AND model.model_type = '" & l_customer & "' "
        End If
        If Not String.IsNullOrEmpty(l_model) Then
            mSQLS1.CommandText += " AND lot.model = '" & l_model & "' "
        End If
        If Not String.IsNullOrEmpty(l_lot) Then
            mSQLS1.CommandText += " AND lot.lot = '" & l_lot & "' "
        End If
        If Not String.IsNullOrEmpty(l_Status) Then
            mSQLS1.CommandText += " AND lot.status = '" & l_Status & "' "
        End If
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            Dim it As Int16 = 1
            While mSQLReader.Read
                Ws.Cells(LineZ, 1) = it
                Ws.Cells(LineZ, 2) = mSQLReader.Item("model_type")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("v3")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 10) = mSQLReader.Item("v1")
                Ws.Cells(LineZ, 16) = mSQLReader.Item("qty")
                Ws.Cells(LineZ, 18) = mSQLReader.Item("v2")
                Ws.Cells(LineZ, 20) = mSQLReader.Item("lot")
                Ws.Cells(LineZ, 21) = mSQLReader.Item("remark")
                Ws.Cells(LineZ, 22) = mSQLReader.Item("users") & " " & mSQLReader.Item("name")
                Ws.Cells(LineZ, 24) = mSQLReader.Item("v4")
                Ws.Cells(LineZ, 25) = mSQLReader.Item("datetime")
                Ws.Cells(LineZ, 29) = mSQLReader.Item("v4")
                Ws.Cells(LineZ, 30) = mSQLReader.Item("v2")
                Ws.Cells(LineZ, 35) = mSQLReader.Item("lot")
                Ws.Cells(LineZ, 36) = mSQLReader.Item("prefix")
                Ws.Cells(LineZ, 37) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 38) = mSQLReader.Item("route")
                Ws.Cells(LineZ, 39) = mSQLReader.Item("qty")
                Ws.Cells(LineZ, 41) = mSQLReader.Item("status")
                Ws.Cells(LineZ, 42) = mSQLReader.Item("remark")
                Ws.Cells(LineZ, 43) = mSQLReader.Item("datetime")
                Ws.Cells(LineZ, 44) = mSQLReader.Item("users") & " " & mSQLReader.Item("name")

                mSQLS2.CommandText = "select count(distinct sn) as t1 from ( select sn from tracking where tracking.station = '0730' and tracking.timeout is not null and tracking.lot = '"
                mSQLS2.CommandText += mSQLReader.Item("lot") & "' "
                mSQLS2.CommandText += "union all select sn from tracking_dup where tracking_dup.station = '0730' and tracking_dup.timeout is not null and tracking_dup.lot = '"
                mSQLS2.CommandText += mSQLReader.Item("lot") & "' "
                mSQLS2.CommandText += "union all select sn from scrap_tracking where scrap_tracking.station = '0730' and scrap_tracking.timeout is not null and scrap_tracking.lot = '"
                mSQLS2.CommandText += mSQLReader.Item("lot") & "' ) as AB"
                Dim AB1 As Decimal = mSQLS2.ExecuteScalar()
                Ws.Cells(LineZ, 15) = AB1
                mSQLS2.CommandText = "select isnull(count(*),0) as t1 from sn where sn.lot = '" & mSQLReader.Item("lot") & "'"
                Dim AB2 As Decimal = mSQLS2.ExecuteScalar()
                Ws.Cells(LineZ, 17) = AB2
                Ws.Cells(LineZ, 40) = AB2
                mSQLS2.CommandText = "select min(timein) from ( select timein from tracking where tracking.station = '0080'  and tracking.lot = '" & mSQLReader.Item("lot") & "' "
                mSQLS2.CommandText += "union all "
                mSQLS2.CommandText += "select timein from scrap_tracking where scrap_tracking.station = '0080'  and scrap_tracking.lot = '" & mSQLReader.Item("lot") & "' ) as AB"
                If Not IsDBNull(mSQLS2.ExecuteScalar()) Then
                    Dim AB3 As Date = mSQLS2.ExecuteScalar()
                    Ws.Cells(LineZ, 26) = AB3.ToString("yyyy/MM/dd")
                End If
                'Ws.Cells(LineZ, 27) = DateDiff(DateInterval.Day, mSQLReader.Item("datetime"), AB3) - 1
                Ws.Cells(LineZ, 27) = "=Z" & LineZ & "-Y" & LineZ & "-1"

                mSQLS2.CommandText = "select max(timeout) as timeout from ( Select timeout from tracking where tracking.station = '0730' AND tracking.timeout is not null and tracking.lot = '"
                mSQLS2.CommandText += mSQLReader.Item("lot") & "' "
                mSQLS2.CommandText += "union all Select timeout from tracking_dup where tracking_dup.station = '0730' AND tracking_dup.timeout is not null and tracking_dup.lot = '"
                mSQLS2.CommandText += mSQLReader.Item("lot") & "' "
                mSQLS2.CommandText += "union all Select timeout from scrap_tracking where scrap_tracking.station = '0730' AND scrap_tracking.timeout is not null and scrap_tracking.lot = '"
                mSQLS2.CommandText += mSQLReader.Item("lot") & "' ) as ab "

                If Not IsDBNull(mSQLS2.ExecuteScalar()) Then
                    Dim AB5 As Date = mSQLS2.ExecuteScalar()
                    Ws.Cells(LineZ, 31) = AB5.ToString("yyyy/MM/dd")
                End If
                Ws.Cells(LineZ, 32) = "=AE" & LineZ & "-AD" & LineZ

                Ws.Cells(LineZ, 34) = "=AE" & LineZ & "-Z" & LineZ

                mSQLS2.CommandText = "select isnull(count(*),0) as t1 from scrap where scrap.lot = '" & mSQLReader.Item("lot") & "'"
                Dim AB4 As Decimal = mSQLS2.ExecuteScalar()
                Ws.Cells(LineZ, 46) = AB4
                Ws.Cells(LineZ, 45) = AB2 + AB4
                If AB2 + AB4 <> 0 Then
                    Ws.Cells(LineZ, 47) = Decimal.Round((AB4 / (AB2 + AB4)), 2)
                Else
                    Ws.Cells(LineZ, 47) = 0
                End If
                mSQLS2.CommandText = "select defect.desc_th , count(sn) as t1 from scrap left join defect on scrap.defect = defect.defect  where scrap.lot = '" & mSQLReader.Item("lot") & "' group by defect.desc_th "
                mSQLReader2 = mSQLS2.ExecuteReader()
                If mSQLReader2.HasRows() Then
                    Dim BC As String = String.Empty
                    While mSQLReader2.Read()
                        BC += mSQLReader2.Item("t1") & "PCS" & mSQLReader2.Item("desc_th") & Chr(10)
                    End While
                    Ws.Cells(LineZ, 48) = BC
                End If
                mSQLReader2.Close()

                it += 1
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "RD_Sample_Report"
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