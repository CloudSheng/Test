Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form189
    Dim Ds As New DataSet()
    Dim Sda As New SqlClient.SqlDataAdapter
    Dim conn As New SqlClient.SqlConnection()
    Dim conn1 As New SqlClient.SqlConnection()
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS11 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form189_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conn.ConnectionString = Module1.OpenConnectionOfERPSUPPORT()
        conn1.ConnectionString = Module1.OpenConnectionOfERPSUPPORT()
        If conn.State <> ConnectionState.Open Then
            Try
                conn.Open()
                mSQLS1.Connection = conn
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If conn1.State <> ConnectionState.Open Then
            Try
                conn1.Open()
                mSQLS11.Connection = conn1
                mSQLS11.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If conn.State <> ConnectionState.Open Then
            Try
                conn.Open()
                mSQLS1.Connection = conn
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString = "SELECT * FROM [S2_Machine list$]"
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)

            Try
                ExcelAdapater.Fill(Ds, "table1")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            LoadData()
            Dim Tran1 As SqlClient.SqlTransaction = conn.BeginTransaction()
            mSQLS1.Transaction = Tran1
            Dim TR As Decimal = 0
            Dim EMS As String = String.Empty

            mSQLS1.CommandText = "DELETE IES2"
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try

            mSQLS1.CommandText = "DELETE IEList_S2"
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try

            For i As Int16 = 0 To Ds.Tables("table1").Rows.Count - 1 Step 1
                EMS = String.Empty
                TR += 1
                '' 先檢查
                'mSQLS1.CommandText = "Select Count(*) from IQMES3.dbo.z_ms_equipment where equipment_id = '" & Ds.Tables("table1").Rows(i).Item(0) & "'"
                'Dim EQH As Int16 = mSQLS1.ExecuteScalar()
                'If EQH = 0 Then
                '    EMS = "12001 - 设备未登记到MES系统中（A列设备编号须在MES系统设备清单中）"
                '    Label2.Text = "错误"
                '    Label2.ForeColor = Color.Red
                '    Label2.Refresh()
                '    mSQLS11.CommandText = "INSERT INTO IEList_S2 VALUES ('" & Ds.Tables("table1").Rows(i).Item(0) & "','" & Ds.Tables("table1").Rows(i).Item(1) & "','N')"
                '    Try
                '        mSQLS11.ExecuteNonQuery()
                '    Catch ex As Exception
                '        MsgBox(ex.Message())
                '    End Try
                '    mSQLS11.CommandText = "UPDATE IEE1 SET Result = 'NG' Where ErrorCode = '12001' "
                '    Try
                '        mSQLS11.ExecuteNonQuery()
                '    Catch ex As Exception
                '        MsgBox(ex.Message())
                '    End Try
                '    Tran1.Rollback()
                '    Exit For
                'End If

                'If String.IsNullOrEmpty(Ds.Tables("table1").Rows(i).Item(2).ToString()) Or String.IsNullOrEmpty(Ds.Tables("table1").Rows(i).Item(3).ToString()) _
                '    Or String.IsNullOrEmpty(Ds.Tables("table1").Rows(i).Item(4).ToString()) Or String.IsNullOrEmpty(Ds.Tables("table1").Rows(i).Item(3).ToString()) _
                '    Or String.IsNullOrEmpty(Ds.Tables("table1").Rows(i).Item(8).ToString()) Or String.IsNullOrEmpty(Ds.Tables("table1").Rows(i).Item(9).ToString()) _
                '    Or String.IsNullOrEmpty(Ds.Tables("table1").Rows(i).Item(10).ToString()) Or String.IsNullOrEmpty(Ds.Tables("table1").Rows(i).Item(11).ToString()) Then

                '    EMS = "12002 - 设备参数缺失：C-E列及I-L列不可为空值"
                '    Label2.Text = "错误"
                '    Label2.ForeColor = Color.Red
                '    Label2.Refresh()
                '    mSQLS11.CommandText = "UPDATE IEE1 SET Result = 'NG' Where ErrorCode = '12002' "
                '    Try
                '        mSQLS11.ExecuteNonQuery()
                '    Catch ex As Exception
                '        MsgBox(ex.Message())
                '    End Try
                '    Tran1.Rollback()
                '    Exit For
                'End If

                'If Ds.Tables("table1").Rows(i).Item(10) >= Ds.Tables("table1").Rows(i).Item(11) Then
                '    EMS = "12003 - 设备参数错误：生效周数必须小于失效周数"
                '    Label2.Text = "错误" & EMS
                '    Label2.ForeColor = Color.Red
                '    Label2.Refresh()
                '    mSQLS11.CommandText = "UPDATE IEE1 SET Result = 'NG' Where ErrorCode = '12003' "
                '    Try
                '        mSQLS11.ExecuteNonQuery()
                '    Catch ex As Exception
                '        MsgBox(ex.Message())
                '    End Try
                '    Tran1.Rollback()
                '    Exit For
                'End If

                'If Ds.Tables("table1").Rows(i).Item(2).ToString().Contains("Autoclave") And IsDBNull(Ds.Tables("table1").Rows(i).Item(12)) Then
                '    EMS = "12004 - 设备参数缺失：类别中含Autoclave的设备不可缺省容量（重量）参数"
                '    Label2.Text = "错误"
                '    Label2.ForeColor = Color.Red
                '    Label2.Refresh()
                '    mSQLS11.CommandText = "UPDATE IEE1 SET Result = 'NG' Where ErrorCode = '12004' "
                '    Try
                '        mSQLS11.ExecuteNonQuery()
                '    Catch ex As Exception
                '        MsgBox(ex.Message())
                '    End Try
                '    Tran1.Rollback()
                '    Exit For
                'End If

                mSQLS1.CommandText = "INSERT INTO IES2 VALUES ('" & Ds.Tables("table1").Rows(i).Item(0) & "','" & Ds.Tables("table1").Rows(i).Item(1) & "','"
                mSQLS1.CommandText += Ds.Tables("table1").Rows(i).Item(2) & "'," & Ds.Tables("table1").Rows(i).Item(3) & "," & Ds.Tables("table1").Rows(i).Item(4) & ","
                mSQLS1.CommandText += Ds.Tables("table1").Rows(i).Item(5) & "," & Ds.Tables("table1").Rows(i).Item(6) & "," & Ds.Tables("table1").Rows(i).Item(7) & ","
                mSQLS1.CommandText += Ds.Tables("table1").Rows(i).Item(8) & "," & Ds.Tables("table1").Rows(i).Item(9) & ",'" & Ds.Tables("table1").Rows(i).Item(10) & "','"
                mSQLS1.CommandText += Ds.Tables("table1").Rows(i).Item(11) & "',"
                If IsDBNull(Ds.Tables("table1").Rows(i).Item(12)) Then
                    mSQLS1.CommandText += "NULL"
                Else
                    mSQLS1.CommandText += Convert.ToString(Ds.Tables("table1").Rows(i).Item(12))
                End If

                mSQLS1.CommandText += ",'" & Ds.Tables("table1").Rows(i).Item(13) & "')"
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    EMS = ex.Message()
                    Tran1.Rollback()
                    Exit For
                End Try

            Next
            If Not String.IsNullOrEmpty(EMS) Then
                MsgBox(EMS)
                Ds.Tables("table1").Clear()
                Label2.Text = "未开始"
                Label2.ForeColor = Color.Black
                Label2.Refresh()
            Else
                Tran1.Commit()
                Label2.Text = "共汇入" & TR & "笔"
                Label2.Refresh()
            End If
        End If
    End Sub
    Private Sub LoadData()
        'Sda = New SqlClient.SqlDataAdapter("select * from IES5", conn)
        'Sda.Fill(Ds)
        Me.DataGridView1.DataSource = Ds.Tables("table1")
        Me.DataGridView1.Columns(0).HeaderText = "设备编号"
        Me.DataGridView1.Columns(1).HeaderText = "名称"
        Me.DataGridView1.Columns(2).HeaderText = "类别"
        'Me.DataGridView1.AutoResizeColumns()
        Me.DataGridView1.Columns(3).HeaderText = "工作平台数量"
        'Me.DataGridView1.Columns(3).Width = 130
        Me.DataGridView1.Columns(4).HeaderText = "容量（尺寸）"
        Me.DataGridView1.Columns(5).HeaderText = "Availability"
        Me.DataGridView1.Columns(5).DefaultCellStyle.Format = "0.0%"
        Me.DataGridView1.Columns(6).HeaderText = "Performance"
        Me.DataGridView1.Columns(6).DefaultCellStyle.Format = "0.0%"
        Me.DataGridView1.Columns(7).HeaderText = "Quality"
        Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "0.0%"
        Me.DataGridView1.Columns(8).HeaderText = "OEE"
        Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "0.0%"
        Me.DataGridView1.Columns(9).HeaderText = "标准产能"
        Me.DataGridView1.Columns(9).DefaultCellStyle.Format = "0.0"
        Me.DataGridView1.Columns(10).HeaderText = "生效周数"
        Me.DataGridView1.Columns(11).HeaderText = "失效周数"
        Me.DataGridView1.Columns(12).HeaderText = "容量（重量）"
        Me.DataGridView1.Columns(13).HeaderText = "备注"
        Me.DataGridView1.AutoResizeColumns()
        'Me.DataGridView1.ColumnHeadersHeight = 10
        'Me.DataGridView1.AutoResizeColumns()
        Me.DataGridView1.Show()
    End Sub
End Class