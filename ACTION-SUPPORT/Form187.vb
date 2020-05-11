﻿Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form187
    Dim Ds As New DataSet()
    Dim Sda As New SqlClient.SqlDataAdapter
    Dim conn As New SqlClient.SqlConnection()
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    'Dim DS As Data.DataSet = New DataSet()
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form187_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conn.ConnectionString = Module1.OpenConnectionOfERPSUPPORT()
        If conn.State <> ConnectionState.Open Then
            Try
                conn.Open()
                mSQLS1.Connection = conn
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcelS4()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcelS4()
    End Sub
    Private Sub ExportToExcelS4()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Name = "S4_Template"
        Ws.Activate()
        AdjustExcelFormat1()
        If conn.State <> ConnectionState.Open Then
            Try
                conn.Open()
                mSQLS1.Connection = conn
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        mSQLS1.CommandText = "Select * from IES4"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                For i As Int16 = 0 To mSQLReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = mSQLReader.Item(i)
                Next
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
        oRng = Ws.Range("A1", "H1")
        oRng.EntireColumn.AutoFit()

        oRng = Ws.Range("A1", Ws.Cells(LineZ - 1, 8))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
    End Sub
    Private Sub AdjustExcelFormat1()
        'xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Cells(1, 1) = "料件编号"
        Ws.Cells(1, 2) = "工段"
        Ws.Cells(1, 3) = "品名"
        'Ws.Cells(1, 3) = "默认设备"
        Ws.Cells(1, 4) = "工艺分类"
        Ws.Cells(1, 5) = "N1_新增涂装工时"
        Ws.Cells(1, 6) = "T2_检验工时"
        Ws.Cells(1, 7) = "T5_设备周期时间"
        Ws.Cells(1, 8) = "变更日期"
        LineZ = 2
    End Sub
    Private Sub SaveExcelS4()
        SaveFileDialog1.FileName = "S4_Additional ST_Updating history"
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
            Dim ExcelString = "SELECT * FROM [S4_Template$]"
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
            Dim Tran2 As Oracle.ManagedDataAccess.Client.OracleTransaction = oConnection.BeginTransaction()
            oCommand.Transaction = Tran2
            For i As Int16 = 0 To Ds.Tables("table1").Rows.Count - 1 Step 1
                EMS = String.Empty
                TR += 1
                ' 先檢查
                If Ds.Tables("table1").Rows(i).Item(3) = "NA" Then
                    EMS = "14001 - 工艺分类不可为NA"
                    Label2.Text = "错误"
                    Label2.ForeColor = Color.Red
                    Label2.Refresh()
                    Tran1.Rollback()
                    Tran2.Rollback()
                    Exit For
                End If

                For j As Int16 = 0 To Ds.Tables("table1").Columns.Count - 1 Step 1
                    If String.IsNullOrEmpty(Ds.Tables("table1").Rows(i).Item(j).ToString()) Then
                        EMS = "14002 - 所有项目不可为空值"
                        Label2.Text = "错误"
                        Label2.ForeColor = Color.Red
                        Label2.Refresh()
                        Tran1.Rollback()
                        Tran2.Rollback()
                        GoTo Error1
                        'Exit For
                    End If
                Next

                If Not String.IsNullOrEmpty(EMS) Then
                    Exit For
                End If
                If IsNumeric(Ds.Tables("table1").Rows(i).Item(5)) Then
                    If Not (Ds.Tables("table1").Rows(i).Item(1).ToString() = "Prepreg" Or Ds.Tables("table1").Rows(i).Item(1).ToString() = "Layup") And Ds.Tables("table1").Rows(i).Item(5) = 0 Then
                        EMS = "14003 - 除工段为Prepreg和Layup的料件外，其他料件T2_检验工时录入时不可为0值"
                        Label2.Text = "错误"
                        Label2.ForeColor = Color.Red
                        Label2.Refresh()
                        Tran1.Rollback()
                        Tran2.Rollback()
                        Exit For
                    End If
                End If
                
                If IsNumeric(Ds.Tables("table1").Rows(i).Item(4)) Then
                    If Ds.Tables("table1").Rows(i).Item(1).ToString() = "Painting" And Ds.Tables("table1").Rows(i).Item(4) = 0 Then
                        EMS = "14004 - 工段为Painting的料件N1_新增涂装工时录入时不可为0值"
                        Label2.Text = "错误"
                        Label2.ForeColor = Color.Red
                        Label2.Refresh()
                        Tran1.Rollback()
                        Tran2.Rollback()
                        Exit For
                    End If
                End If
                
                If IsNumeric(Ds.Tables("table1").Rows(i).Item(6)) Then
                    If (Ds.Tables("table1").Rows(i).Item(1).ToString() = "Prepreg" Or Ds.Tables("table1").Rows(i).Item(1).ToString() = "Molding" Or Ds.Tables("table1").Rows(i).Item(1).ToString() = "Painting") And Ds.Tables("table1").Rows(i).Item(6) = 0 Then
                        EMS = "14005 - 工段为Prepreg、Molding、Painting的料件T5_设备周期时间录入时不可为0值"
                        Label2.Text = "错误"
                        Label2.ForeColor = Color.Red
                        Label2.Refresh()
                        Tran1.Rollback()
                        Tran2.Rollback()
                        Exit For
                    End If
                End If
                

                mSQLS1.CommandText = "INSERT INTO IES4 VALUES ('" & Ds.Tables("table1").Rows(i).Item(0) & "','" & Ds.Tables("table1").Rows(i).Item(1) & "','"
                mSQLS1.CommandText += Ds.Tables("table1").Rows(i).Item(2) & "','" & Ds.Tables("table1").Rows(i).Item(3) & "','" & Ds.Tables("table1").Rows(i).Item(4) & "','"
                mSQLS1.CommandText += Ds.Tables("table1").Rows(i).Item(5) & "','" & Ds.Tables("table1").Rows(i).Item(6) & "','" & Ds.Tables("table1").Rows(i).Item(7) & "')"
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    EMS = ex.Message()
                    Tran1.Rollback()
                    Tran2.Rollback()
                    Exit For
                End Try

                Dim U1 As Boolean = True
                Dim U2 As Boolean = True
                Dim U3 As Boolean = True

                If Ds.Tables("table1").Rows(i).Item(4) = "NA" Then
                    U1 = False
                End If
                If Ds.Tables("table1").Rows(i).Item(5) = "NA" Then
                    U2 = False
                End If
                If Ds.Tables("table1").Rows(i).Item(6) = "NA" Then
                    U3 = False
                End If

                Dim S1 As String = String.Empty
                Dim S2 As String = String.Empty
                Dim S3 As String = String.Empty

                If Ds.Tables("table1").Rows(i).Item(4) = "/" Then
                    S1 = 0
                Else
                    S1 = Ds.Tables("table1").Rows(i).Item(4)
                End If
                If Ds.Tables("table1").Rows(i).Item(5) = "/" Then
                    S2 = 0
                Else
                    S2 = Ds.Tables("table1").Rows(i).Item(5)
                End If
                If Ds.Tables("table1").Rows(i).Item(6) = "/" Then
                    S3 = 0
                Else
                    S3 = Ds.Tables("table1").Rows(i).Item(6)
                End If

                oCommand.CommandText = "UPDATE tc_imf_file SET tc_imf02 = '" & Ds.Tables("table1").Rows(i).Item(3) & "' "
                If U1 = True Then
                    oCommand.CommandText += ",tc_imf03 = '" & S1 & "' "
                End If
                If U2 = True Then
                    oCommand.CommandText += ",tc_imf04 = " & S2
                End If
                If U3 = True Then
                    oCommand.CommandText += ",tc_imf10 = " & S3
                End If
                oCommand.CommandText += " WHERE tc_imf01 = '" & Ds.Tables("table1").Rows(i).Item(0) & "' "
                Try
                    oCommand.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    Tran1.Rollback()
                    Tran2.Rollback()
                End Try

            Next
Error1:
            If Not String.IsNullOrEmpty(EMS) Then
                MsgBox(EMS)
                Ds.Tables("table1").Clear()
                Label2.Text = "未开始"
                Label2.ForeColor = Color.Black
                Label2.Refresh()
            Else
                Tran1.Commit()
                Tran2.Commit()
                Label2.Text = "共汇入" & TR & "笔"
                Label2.Refresh()
            End If



            'LoadData()
        End If
    End Sub
    Private Sub LoadData()
        'Sda = New SqlClient.SqlDataAdapter("select * from IES5", conn)
        'Sda.Fill(Ds)
        Me.DataGridView1.DataSource = Ds.Tables("table1")
        Me.DataGridView1.Columns(0).HeaderText = "料件编号"
        Me.DataGridView1.Columns(1).HeaderText = "工段"
        Me.DataGridView1.Columns(2).HeaderText = "品名"
        'Me.DataGridView1.Columns(2).HeaderText = "默认设备"
        Me.DataGridView1.AutoResizeColumns()
        Me.DataGridView1.Columns(3).HeaderText = "工艺分类"
        Me.DataGridView1.Columns(3).Width = 130
        Me.DataGridView1.Columns(4).HeaderText = "N1_新增涂装工时"
        Me.DataGridView1.Columns(4).Width = 160
        Me.DataGridView1.Columns(5).HeaderText = "T2_检验工时"
        Me.DataGridView1.Columns(5).Width = 130
        Me.DataGridView1.Columns(6).HeaderText = "T5_设备周期时间"
        Me.DataGridView1.Columns(6).Width = 160
        Me.DataGridView1.Columns(7).HeaderText = "变更日期"
        Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "yyyy/MM/dd"
        Me.DataGridView1.Columns(7).Width = 100
        'Me.DataGridView1.ColumnHeadersHeight = 10
        'Me.DataGridView1.AutoResizeColumns()
        Me.DataGridView1.Show()
    End Sub
End Class