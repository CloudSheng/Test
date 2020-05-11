Public Class Form7
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString = "SELECT * FROM [sheet1$]"
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)
            Dim DS As Data.DataSet = New DataSet()
            Try
                ExcelAdapater.Fill(DS, "table1")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            Me.DataGridView1.DataSource = DS.Tables("table1")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Me.DataGridView1.Rows.Count = 0 Then
            MsgBox("无资料可处理")
            Return
        End If
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
        For i As Integer = 0 To DataGridView1.Rows.Count - 1 Step 1
            mSQLS1.CommandText = "SELECT COUNT(*) FROM sn where sn ='"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "'"
            Dim C1 As Int16 = mSQLS1.ExecuteScalar()
            If C1 = 0 Then
                MsgBox(DataGridView1.Rows(i).Cells(1).Value & "ERROR , NO THIS SN")
                Exit For
            End If
            If C1 > 1 Then
                MsgBox("ERROR , 2 OR MORE SN ")
                Exit For
            End If
            'mConnection.BeginTransaction()
            mSQLS1.CommandText = "INSERT INTO scrap_sn(lot, sn, register ,lasttimein,laststation,lasttimeout,currentstation,lastresult,failcount,block,shipped,remark,updatedstation)select lot, sn, register ,lasttimein,laststation,lasttimeout,currentstation,lastresult,failcount,block,shipped,remark,updatedstation from sn where sn = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                '   mConnection.BeginTransaction.Rollback()
                Exit For
            End Try
            mSQLS1.CommandText = "delete sn where sn = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                '  mConnection.BeginTransaction.Rollback()
                Exit For
            End Try
            mSQLS1.CommandText = "SELECT COUNT(*) FROM tracking where sn ='"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
            Dim C2 As Int16 = mSQLS1.ExecuteScalar()
            If C2 > 0 Then
                mSQLS1.CommandText = "insert into scrap_tracking(lot, sn, timein ,timeout,result,fresh,station,pc,users,lead,qa,timeout1,timeout2,timeout3,users1,users2,users3,transferby1,transferby2,transferby3) select lot, sn, timein ,timeout,result,fresh,station,pc,users,lead,qa,timeout1,timeout2,timeout3,users1,users2,users3,transferby1,transferby2,transferby3 from tracking where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    ' mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
                mSQLS1.CommandText = "delete tracking where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    'mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
            End If
            mSQLS1.CommandText = "SELECT COUNT(*) FROM comp_tracking where sn ='"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
            Dim C3 As Int16 = mSQLS1.ExecuteScalar()
            If C3 > 0 Then
                mSQLS1.CommandText = "insert into scrap_comp_tracking (lot, sn, pn, seq, datetime, datecode, lotcode, batch, supplier, pnsn) select lot, sn, pn, seq, datetime, datecode, lotcode, batch, supplier, pnsn from comp_tracking where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    ' mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
                mSQLS1.CommandText = "delete comp_tracking where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    'mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
            End If
            mSQLS1.CommandText = "SELECT COUNT(*) FROM paravalue where sn ='"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
            Dim C4 As Int16 = mSQLS1.ExecuteScalar()
            If C4 > 0 Then
                mSQLS1.CommandText = "insert into scrap_paravalue(lot, sn, station,parameter,value) select lot, sn, station,parameter,value from paravalue where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    ' mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
                mSQLS1.CommandText = "delete paravalue where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    '  mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
            End If
            mSQLS1.CommandText = "SELECT COUNT(*) FROM failure where sn ='"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
            Dim C5 As Int16 = mSQLS1.ExecuteScalar()
            If C5 > 0 Then
                mSQLS1.CommandText = "insert into scrap_failure(lot, sn, failtime,failstation,defect,defect_remk,rework,rework_remk,reworktime,users,pn,pnstation,pnseq,datecode,lotcode,batch,supplier,pnsn) select lot, sn, failtime,failstation,defect,defect_remk,rework,rework_remk,reworktime,users,pn,pnstation,pnseq,datecode,lotcode,batch,supplier,pnsn from failure where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    '   mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
                mSQLS1.CommandText = "delete failure where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    '  mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
            End If
            mSQLS1.CommandText = "SELECT COUNT(*) FROM box_detail where sn ='"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
            Dim C6 As Int16 = mSQLS1.ExecuteScalar()
            If C6 > 0 Then
                mSQLS1.CommandText = "insert into scrap_box_detail(lot, sn, boxid ,seq_number,datetime) select lot, sn, boxid ,seq_number,datetime from box_detail where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    '   mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
                mSQLS1.CommandText = "delete box_detail where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    '   mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
            End If
            mSQLS1.CommandText = "insert into scrap(sn, lot,datetime,users,defect,cause,remark,stock_location) values('"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "','"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "','"
            mSQLS1.CommandText += Now.ToString("yyyy/MM/dd HH:mm:ss") & "','99018','" & TextBox1.Text & "','SCRP','" & TextBox2.Text & "','SCRP')"
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                '  mConnection.BeginTransaction.Rollback()
                Exit For
            End Try
            '   mConnection.BeginTransaction.Commit()
        Next
        MsgBox("处理完毕")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Me.DataGridView1.Rows.Count = 0 Then
            MsgBox("无资料可处理")
            Return
        End If
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
        For i As Integer = 0 To DataGridView1.Rows.Count - 1 Step 1
            mSQLS1.CommandText = "SELECT COUNT(*) FROM scrap_sn where sn ='"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "'"
            Dim C1 As Int16 = mSQLS1.ExecuteScalar()
            If C1 = 0 Then
                MsgBox("ERROR , NO THIS SN" & DataGridView1.Rows(i).Cells(1).Value)
                Exit For
            End If
            If C1 > 1 Then
                MsgBox("ERROR , 2 OR MORE SN ")
                Exit For
            End If
            'mConnection.BeginTransaction()
            mSQLS1.CommandText = "insert into sn(lot, sn, register ,lasttimein,laststation,lasttimeout,currentstation,lastresult,failcount,block,shipped,remark,updatedstation)select lot, sn, register ,lasttimein,laststation,lasttimeout,currentstation,lastresult,failcount,block,shipped,remark,updatedstation from scrap_sn where sn = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                '   mConnection.BeginTransaction.Rollback()
                Exit For
            End Try
            mSQLS1.CommandText = "delete scrap_sn where sn = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                '  mConnection.BeginTransaction.Rollback()
                Exit For
            End Try
            mSQLS1.CommandText = "SELECT COUNT(*) FROM scrap_tracking where sn ='"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
            Dim C2 As Int16 = mSQLS1.ExecuteScalar()
            If C2 > 0 Then
                mSQLS1.CommandText = "insert into tracking(lot, sn, timein ,timeout,result,fresh,station,pc,users,lead,qa,timeout1,timeout2,timeout3,users1,users2,users3,transferby1,transferby2,transferby3) select lot, sn, timein ,timeout,result,fresh,station,pc,users,lead,qa,timeout1,timeout2,timeout3,users1,users2,users3,transferby1,transferby2,transferby3 from scrap_tracking where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    ' mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
                mSQLS1.CommandText = "delete scrap_tracking where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    'mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
            End If
            mSQLS1.CommandText = "SELECT COUNT(*) FROM scrap_comp_tracking where sn ='"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
            Dim C3 As Int16 = mSQLS1.ExecuteScalar()
            If C3 > 0 Then
                mSQLS1.CommandText = "insert into comp_tracking (lot, sn, pn, seq, datetime, datecode, lotcode, batch, supplier, pnsn) select lot, sn, pn, seq, datetime, datecode, lotcode, batch, supplier, pnsn from scrap_comp_tracking where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    ' mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
                mSQLS1.CommandText = "delete scrap_comp_tracking where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    'mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
            End If
            mSQLS1.CommandText = "SELECT COUNT(*) FROM scrap_paravalue where sn ='"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
            Dim C4 As Int16 = mSQLS1.ExecuteScalar()
            If C4 > 0 Then
                mSQLS1.CommandText = "insert into paravalue(lot, sn, station,parameter,value) select lot, sn, station,parameter,value from scrap_paravalue where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    ' mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
                mSQLS1.CommandText = "delete scrap_paravalue where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    '  mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
            End If
            mSQLS1.CommandText = "SELECT COUNT(*) FROM scrap_failure where sn ='"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
            Dim C5 As Int16 = mSQLS1.ExecuteScalar()
            If C5 > 0 Then
                mSQLS1.CommandText = "insert into failure(lot, sn, failtime,failstation,defect,defect_remk,rework,rework_remk,reworktime,users,pn,pnstation,pnseq,datecode,lotcode,batch,supplier,pnsn) select lot, sn, failtime,failstation,defect,defect_remk,rework,rework_remk,reworktime,users,pn,pnstation,pnseq,datecode,lotcode,batch,supplier,pnsn from scrap_failure where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    '   mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
                mSQLS1.CommandText = "delete scrap_failure where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    '  mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
            End If
            mSQLS1.CommandText = "SELECT COUNT(*) FROM scrap_box_detail where sn ='"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
            Dim C6 As Int16 = mSQLS1.ExecuteScalar()
            If C6 > 0 Then
                mSQLS1.CommandText = "insert into box_detail(lot, sn, boxid ,seq_number,datetime) select lot, sn, boxid ,seq_number,datetime from scrap_box_detail where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    '   mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
                mSQLS1.CommandText = "delete scrap_failure where sn = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
                mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    '   mConnection.BeginTransaction.Rollback()
                    Exit For
                End Try
            End If
            mSQLS1.CommandText = "delete scrap where sn = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(1).Value & "' and lot = '"
            mSQLS1.CommandText += DataGridView1.Rows(i).Cells(0).Value & "' "
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                '  mConnection.BeginTransaction.Rollback()
                Exit For
            End Try
            '   mConnection.BeginTransaction.Commit()
        Next
        MsgBox("处理完毕")
    End Sub
End Class