Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form113
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim ptime As String = String.Empty
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim tYear As String = String.Empty
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form113_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        ptime = Today.AddDays(-1).ToString("yyyy/MM/dd")
        ptime = ptime & " 08:00:00"
        Me.DateTimePicker2.Value = Convert.ToDateTime(ptime)
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
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
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Dim xPath As String = "C:\temp\Shipping plan management.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommand2.Connection = oConnection
                oCommand2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        tYear = Me.DateTimePicker2.Value.Year
        TimeS1 = Me.DateTimePicker2.Value
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        CreateTempTable()
        ProcessData()
        ExportToExcel()
    End Sub
    Private Sub CreateTempTable()
        oCommand.CommandText = "DROP TABLE ship_temp"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception

        End Try
        oCommand.CommandText = "CREATE TABLE ship_temp (eType number(2), Customer varchar2(40), PN varchar2(40), ima02 varchar2(255), ima021 varchar2(255), MESMODEL varchar2(20), ima08 varchar2(1), "
        oCommand.CommandText += "ima25 varchar2(4), start1 number(15,3) "
        For i As Int16 = 1 To 53 Step 1
            oCommand.CommandText += ", w" & i & " number(15,3)"
        Next
        oCommand.CommandText += " )"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

    End Sub
    Private Sub ProcessData()
        ' 年度預測計劃
        Me.Label1.Text = "Proceess 1"
        oCommand.CommandText = "select 1,tqa02,tc_bud04,ima02,ima021,'',ima08,ima25,tc_bud02,tc_bud03,tc_bud11 from tc_bud_file "
        oCommand.CommandText += "left join ima_file on tc_bud04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = 2 where tc_bud01 = 1 and tc_bud02 = " & tYear
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Dim Mods As Decimal = 0
                Mods = Decimal.Remainder(oReader.Item("tc_bud11"), 4)
                Dim sMonth As Decimal = oReader.Item("tc_bud03")
                Dim sYear As Decimal = oReader.Item("tc_bud02")
                oCommand2.CommandText = "select azn05 from azn_file where azn02 = " & oReader.Item("tc_bud02") & " and azn04= " & oReader.Item("tc_bud03") & " order by azn05"
                Dim firstweek As Decimal = oCommand2.ExecuteScalar()
                For i = 1 To 4 Step 1
                    Dim Quantity1 As Decimal = 0
                    If i = 1 And Mods <> 0 Then
                        Quantity1 = Decimal.Floor(oReader.Item("tc_bud11") / 4) + Mods
                    Else
                        Quantity1 = Decimal.Divide((oReader.Item("tc_bud11") - Mods), 4)
                    End If
                    oCommand2.CommandText = "INSERT INTO ship_temp (eType, Customer, PN, ima02,ima021,ima08,ima25,w" & firstweek + i - 1 & ") VALUES (1,'"
                    oCommand2.CommandText += oReader.Item("tqa02") & "','" & oReader.Item("tc_bud04") & "','" & oReader.Item("ima02") & "','" & oReader.Item("ima021") & "','" & oReader.Item("ima08") & "','" & oReader.Item("ima25") & "'," & Quantity1 & ")"
                    Try
                        oCommand2.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                Next
            End While
        End If
        oReader.Close()
        '' 訂單
        'Me.Label1.Text = "Proceess 2"
        'oCommand.CommandText = "select 2,tqa02,oeb04,ima02,ima021,'',ima08,ima25,azn05,oeb12 from oeb_file left join oea_file on oeb01 = oea01 "
        'oCommand.CommandText += "left join ima_file on oeb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = 2 left join azn_file on oea02 = azn01 "
        'oCommand.CommandText += "where  oeaconf = 'Y' and oea02 > to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd')"
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        oCommand2.CommandText = "INSERT INTO ship_temp (eType, Customer, PN, ima02,ima021,ima08,ima25,w" & oReader.Item("azn05") & ") VALUES (2,'"
        '        oCommand2.CommandText += oReader.Item("tqa02") & "','" & oReader.Item("oeb04") & "','" & oReader.Item("ima02") & "','" & oReader.Item("ima021") & "','" & oReader.Item("ima08") & "','" & oReader.Item("ima25") & "'," & oReader.Item("oeb12") & ")"
        '        Try
        '            oCommand2.ExecuteNonQuery()
        '        Catch ex As Exception
        '            MsgBox(ex.Message())
        '        End Try
        '    End While
        'End If
        'oReader.Close()
        ' 多交期資料
        Me.Label1.Text = "Proceess 3"
        oCommand.CommandText = "select 3,tqa02,oeb04,ima02,ima021,'',ima08,ima25,azn05,tc_cif_04 from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
        oCommand.CommandText += "left join ima_file on oeb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = 2 left join azn_file on tc_cif_05 = azn01 "
        oCommand.CommandText += "where tc_cif_05 > to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd')"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                oCommand2.CommandText = "INSERT INTO ship_temp (eType, Customer, PN, ima02,ima021,ima08,ima25,w" & oReader.Item("azn05") & ") VALUES (3,'"
                oCommand2.CommandText += oReader.Item("tqa02") & "','" & oReader.Item("oeb04") & "','" & oReader.Item("ima02") & "','" & oReader.Item("ima021") & "','" & oReader.Item("ima08") & "','" & oReader.Item("ima25") & "'," & oReader.Item("tc_cif_04") & ")"
                Try
                    oCommand2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End While
        End If
        oReader.Close()
        ' 入庫數量
        Me.Label1.Text = "Proceess 4"
        oCommand.CommandText = "select 6,tqa02,tlf01,ima02,ima021,'',ima08,ima25,azn05,(tlf10*tlf12) as t1 from tlf_file left join ima_file on tlf01 = ima01 "
        oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = 2 left join azn_file on tlf06  = azn01 "
        oCommand.CommandText += "where tlf06 > to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd') and tlf13 = 'aimt324' and tlf902 = 'D146103'"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                oCommand2.CommandText = "INSERT INTO ship_temp (eType, Customer, PN, ima02,ima021,ima08,ima25,w" & oReader.Item("azn05") & ") VALUES (6,'"
                oCommand2.CommandText += oReader.Item("tqa02") & "','" & oReader.Item("tlf01") & "','" & oReader.Item("ima02") & "','" & oReader.Item("ima021") & "','" & oReader.Item("ima08") & "','" & oReader.Item("ima25") & "'," & oReader.Item("t1") & ")"
                Try
                    oCommand2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End While
        End If
        oReader.Close()
        ' 出貨單
        Me.Label1.Text = "Proceess 5"
        oCommand.CommandText = "select 7,tqa02,ogb04,ima02,ima021,'',ima08,ima25,azn05,ogb12 from ogb_file left join oga_file on ogb01 = oga01 "
        oCommand.CommandText += "left join ima_file on ogb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = 2 left join azn_file on oga02 = azn01 "
        oCommand.CommandText += "where  ogaconf = 'Y' and ogapost = 'Y' and oga02 > to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd')"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                oCommand2.CommandText = "INSERT INTO ship_temp (eType, Customer, PN, ima02,ima021,ima08,ima25,w" & oReader.Item("azn05") & ") VALUES (7,'"
                oCommand2.CommandText += oReader.Item("tqa02") & "','" & oReader.Item("ogb04") & "','" & oReader.Item("ima02") & "','" & oReader.Item("ima021") & "','" & oReader.Item("ima08") & "','" & oReader.Item("ima25") & "'," & oReader.Item("ogb12") & ")"
                Try
                    oCommand2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End While
        End If
        oReader.Close()
        ' 期初  訂單餘量
        Me.Label1.Text = "Proceess 6"
        oCommand.CommandText = "select 12,tqa02,oeb04,ima02,ima021,'',ima08,ima25,sum(oeb12-oeb24) as t1 from oeb_file left join oea_file on oeb01 = oea01 "
        oCommand.CommandText += "left join ima_file on oeb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = 2 where oeaconf = 'Y' and oeb04 not like 'S%' and ima25 = 'PCS' and ima06 = '103' and ima01 like '%66' and imaacti = 'Y' "
        oCommand.CommandText += "and (oeb12 - oeb24) > 0 and oeb70 = 'N' group by tqa02,oeb04,ima02,ima021,ima08,ima25 order by oeb04"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                oCommand2.CommandText = "INSERT INTO ship_temp (eType, Customer, PN, ima02,ima021,ima08,ima25,start1) VALUES (12,'"
                oCommand2.CommandText += oReader.Item("tqa02") & "','" & oReader.Item("oeb04") & "','" & oReader.Item("ima02") & "','" & oReader.Item("ima021") & "','" & oReader.Item("ima08") & "','" & oReader.Item("ima25") & "'," & oReader.Item("t1") & ")"
                Try
                    oCommand2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End While
        End If
        oReader.Close()
        ' 期初 庫存??
        Me.Label1.Text = "Proceess 7"
        oCommand.CommandText = "select 11, tqa02,imk01,ima02,ima021,'',ima08,ima25,sum(imk09) as t1 from imk_file "
        oCommand.CommandText += "left join ima_file on imk01 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = 2 "
        oCommand.CommandText += "where imk05 = 2018 and imk06 = 2 and imaacti = 'Y' and ima01 not like 'S%' and ima25 = 'PCS' and ima06 = '103' and ima01 like '%66' "
        oCommand.CommandText += "and imk02 = 'D146103' and imk09 <> 0 group by tqa02,imk01,ima02,ima021,ima08,ima25"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                oCommand2.CommandText = "INSERT INTO ship_temp (eType, Customer, PN, ima02,ima021,ima08,ima25,start1) VALUES (11,'"
                oCommand2.CommandText += oReader.Item("tqa02") & "','" & oReader.Item("imk01") & "','" & oReader.Item("ima02") & "','" & oReader.Item("ima021") & "','" & oReader.Item("ima08") & "','" & oReader.Item("ima25") & "'," & oReader.Item("t1") & ")"
                Try
                    oCommand2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End While
        End If
        oReader.Close()
        ' 期初 -  shiptemp2 生產欠數
        Me.Label1.Text = "Proceess 8"
        oCommand.CommandText = "select 8, tqa02,pn,ima02,ima021,'',ima08,ima25,ship_temp2.q1 from ship_temp2 left join ima_file on pn = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = 2"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                oCommand2.CommandText = "INSERT INTO ship_temp (eType, Customer, PN, ima02,ima021,ima08,ima25,start1) VALUES (8,'"
                oCommand2.CommandText += oReader.Item("tqa02") & "','" & oReader.Item("pn") & "','" & oReader.Item("ima02") & "','" & oReader.Item("ima021") & "','" & oReader.Item("ima08") & "','" & oReader.Item("ima25") & "'," & oReader.Item("q1") & ")"
                Try
                    oCommand2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End While
        End If
        oReader.Close()
        Me.Label1.Text = "Idle"
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        Dim xPath As String = "C:\temp\Shipping plan management.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        'AdjustmentExcelFormat()
        LineZ = 4
        oCommand.CommandText = "select etype,customer,pn,ima02,ima021,mesmodel,ima08,ima25,start1"
        For i As Int16 = 1 To 53 Step 1
            oCommand.CommandText += ",sum(w" & i & ") as w" & i
        Next
        oCommand.CommandText += " from ship_temp group by etype,customer,pn,ima02,ima021,mesmodel,ima08,ima25,start1 order by pn,etype"
        oReader = oCommand.ExecuteReader()

        Dim CheckPN As String = String.Empty
        Dim IniRows As Int16 = 1
        If oReader.HasRows() Then
            While oReader.Read()
                If CheckPN <> oReader.Item("pn") Then
                    IniRows = 1
                    CheckPN = oReader.Item("pn")
                End If
                Dim NowRows As Int16 = mSQLReader.Item("eType")
                For i As Int16 = IniRows To NowRows Step 1
                    Select Case i
                        Case 1
                            Ws.Cells(LineZ, 1) = 1
                            Ws.Cells(LineZ, 2) = oReader.Item("customer")
                            Ws.Cells(LineZ, 3) = oReader.Item("pn")
                            Ws.Cells(LineZ, 4) = oReader.Item("ima02")
                            Ws.Cells(LineZ, 5) = oReader.Item("ima021")
                            Ws.Cells(LineZ, 6) = "待處理"
                            Dim l_ima08 As String = String.Empty
                            Select Case oReader.Item("ima08")
                                Case ""
                            End Select
                    End Select
                Next
            End While
        End If
        oReader.Close()
    End Sub
End Class