Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports System.Drawing
Public Class Form121
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand3 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
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
    Dim LineS1 As Int16 = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Dim ArrayS1 As String() = {"101AA0101031066", "101AA0102031066", "101AA0103031066", "101AA0105031066", "101AA0106031066", _
                               "101AA0107031066", "101AA0108031066", "101AA0109031066", "101AA0110031066", "101AA0115031066", _
                               "101AA0801031066", "101AA0901031066", "101AB0201041066", "101AB0209041066", "101AB0222031066", _
                               "101AB0223031066", "101AB0224031066", "101AB0225031066", "101AF0203031066", "101AH0403031066", _
                               "101AH0517031066", "101AH0707031066", "101AH0708031066", "101AH0710031066", "101AH0711031066", _
                               "101AH0801031066", "101AH0803031066", "101AH0807032066", "101AH0808032066", "101AH0905031066", _
                               "101AH0906031066", "101AL0103031066", "101AL0104031066", "101AL0302031066", "101AN0101031066", _
                               "101AP0101031066", "101AQ0301031066", "101AQ0303031066", "101AQ0305031066", "101AQ0306031066", _
                               "101AQ0307031066", "101BG0102031066", "104AB0104031066", "108AB0113030066", "121AH0305031066", _
                               "121AH0320031066", "125AH0902051066", "125AH0902061066", "127AL0301041066", "128AP0102031066", _
                               "128BG0101031066", "143AP0110031066", "144BG0103031066", "147AB0702031066", "147AB0704031066"}

    Private Sub Form121_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Dim xPath As String = "C:\temp\Demand and schedule.xlsx"
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
                oCommand3.Connection = oConnection
                oCommand3.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
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
        BackgroundWorker1.RunWorkerAsync()
        'CreateTempTable()
        'ProcessData()
        'ExportToExcel()
        'SaveExcel()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        CreateTempTable()
        ProcessData()
        ExportToExcel()
    End Sub
    Private Sub CreateTempTable()
        Me.Label2.Text = "DROP TABLE"
        oCommand.CommandText = "DROP TABLE ship_temp3"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception

        End Try
        Me.Label2.Text = "CREATE TABLE"
        oCommand.CommandText = "CREATE TABLE ship_temp3 (eType number(2), Customer varchar2(40), PN varchar2(40), ima02 varchar2(255), MES_MODEL varchar2(30) "
        For i As Int16 = 20 To 178 Step 1
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
        Me.Label2.Text = "临时资料处理"

        '銷售訂單
        'oCommand.CommandText = "select 1,tqa02,oeb04,ima02,tc_azn02,tc_azn05,oeb12 from oeb_file left join ima_file on oeb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = 2 "
        'oCommand.CommandText += "left join oea_file on oeb01 = oea01  left join tc_azn_file on oea02 = tc_azn01 where oea02 between to_date('2018/05/12','yyyy/MM/dd') and to_date('2021/5/7','yyyy/MM/dd') and ima06 = '103' and ima08 ='M' "
        'oCommand.CommandText += "and oeb04 not like 'S%' and oeb04 not like 'A%' and oeaconf = 'Y' and substr(ima01,10,1) <> 'A' and ima55 IN ('PCS','SET') "
        oCommand.CommandText = "select 1,tqa02,ERPPN,ima02,tc_azn02,tc_azn05,Quantity from ship_temp99 left join ima_file on ERPPN = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = 2 left join tc_azn_file on RecordDate = tc_azn01 where RecordDate between to_date('2018/05/12','yyyy/MM/dd') and to_date('2021/5/7','yyyy/MM/dd') and ima06 = '103' and ima08 ='M' and ERPPN not like 'S%' and ERPPN not like 'A%' and ima55 IN ('PCS','SET') "
        oReader = oCommand.ExecuteReader()
        Dim Section1C As Decimal = 0
        If oReader.HasRows Then
            While oReader.Read()
                Section1C += 1
                Me.Label2.Text = "Sales Order " & Section1C
                Me.Label2.Refresh()
                If IsDBNull(oReader.Item("ERPPN")) Then
                    Continue While
                End If
                mSQLS1.CommandText = "select model from model_paravalue where parameter = 'ERP PN'  AND VALUE = '" & oReader.Item("ERPPN") & "'"
                Dim l_Model As String = String.Empty
                l_Model = mSQLS1.ExecuteScalar()
                Dim l_week As Decimal = 0
                Dim l_year As Decimal = 0
                l_year = oReader.Item("tc_azn02")
                l_week = 53 * (l_year - 2018) + oReader.Item("tc_azn05")
                If ArrayS1.Contains(oReader.Item("ERPPN").ToString()) Then
                    mSQLS1.CommandText = "SELECT * FROM ERPSUPPORT.dbo.SETVSPCS WHERE [SET] = '" & oReader.Item("ERPPN") & "'"
                    mSQLReader = mSQLS1.ExecuteReader
                    If mSQLReader.HasRows() Then
                        While mSQLReader.Read()
                            mSQLS2.CommandText = "select model from model_paravalue where parameter = 'ERP PN'  AND VALUE = '" & mSQLReader.Item("PCS") & "'"
                            l_Model = mSQLS2.ExecuteScalar()
                            oCommand2.CommandText = "SELECT ima02 FROM ima_file WHERE  ima01 = '" & mSQLReader.Item("PCS") & "'"
                            Dim l_ima02 As String = String.Empty
                            l_ima02 = oCommand2.ExecuteScalar()
                            oCommand2.CommandText = "INSERT INTO ship_temp3 (eType, Customer, PN, ima02, MES_MODEL ,w" & l_week & ") VALUES (1,'"
                            oCommand2.CommandText += oReader.Item("tqa02") & "','" & mSQLReader.Item("PCS") & "','" & l_ima02 & "','" & l_Model & "'," & oReader.Item("Quantity") & ")"
                            Try
                                oCommand2.ExecuteNonQuery()
                            Catch ex As Exception
                                MsgBox(ex.Message())
                            End Try
                        End While
                    End If
                    mSQLReader.Close()
                Else
                    oCommand2.CommandText = "INSERT INTO ship_temp3 (eType, Customer, PN, ima02, MES_MODEL ,w" & l_week & ") VALUES (1,'"
                    oCommand2.CommandText += oReader.Item("tqa02") & "','" & oReader.Item("ERPPN") & "','" & oReader.Item("ima02") & "','" & l_Model & "'," & oReader.Item("Quantity") & ")"
                    Try
                        oCommand2.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End While
        End If
        oReader.Close()

        ' 多交期資料
        oCommand.CommandText = "select 2,tqa02,oeb04,ima02,tc_azn02,tc_azn05,tc_cif_04 from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
        oCommand.CommandText += "left join ima_file on oeb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = 2 left join tc_azn_file on tc_cif_05 = tc_azn01 "
        oCommand.CommandText += "where tc_cif_05 between to_date('2018/05/12','yyyy/MM/dd') and to_date('2021/5/7','yyyy/MM/dd') and oeb70 = 'N' and tc_cif_01 not like 'FC%' "
        oReader = oCommand.ExecuteReader()
        Dim Section2C As Decimal = 0
        If oReader.HasRows() Then
            While oReader.Read()
                Section2C += 1
                Me.Label2.Text = "Customer Demand " & Section2C
                Me.Label2.Refresh()
                If IsDBNull(oReader.Item("oeb04")) Then
                    Continue While
                End If
                mSQLS1.CommandText = "select model from model_paravalue where parameter = 'ERP PN'  AND VALUE = '" & oReader.Item("oeb04") & "'"
                Dim l_Model As String = String.Empty
                l_Model = mSQLS1.ExecuteScalar()
                Dim l_week As Decimal = 0
                Dim l_year As Decimal = 0
                l_year = oReader.Item("tc_azn02")
                l_week = 53 * (l_year - 2018) + oReader.Item("tc_azn05")
                If ArrayS1.Contains(oReader.Item("oeb04").ToString()) Then
                    mSQLS1.CommandText = "SELECT * FROM ERPSUPPORT.dbo.SETVSPCS WHERE [SET] = '" & oReader.Item("oeb04") & "'"
                    mSQLReader = mSQLS1.ExecuteReader
                    If mSQLReader.HasRows() Then
                        While mSQLReader.Read()
                            mSQLS2.CommandText = "select model from model_paravalue where parameter = 'ERP PN'  AND VALUE = '" & mSQLReader.Item("PCS") & "'"
                            l_Model = mSQLS2.ExecuteScalar()
                            oCommand2.CommandText = "SELECT ima02 FROM ima_file WHERE  ima01 = '" & mSQLReader.Item("PCS") & "'"
                            Dim l_ima02 As String = String.Empty
                            l_ima02 = oCommand2.ExecuteScalar()
                            oCommand2.CommandText = "INSERT INTO ship_temp3 (eType, Customer, PN, ima02, MES_MODEL ,w" & l_week & ") VALUES (2,'"
                            oCommand2.CommandText += oReader.Item("tqa02") & "','" & mSQLReader.Item("PCS") & "','" & l_ima02 & "','" & l_Model & "'," & oReader.Item("tc_cif_04") & ")"
                            Try
                                oCommand2.ExecuteNonQuery()
                            Catch ex As Exception
                                MsgBox(ex.Message())
                            End Try
                        End While
                    End If
                    mSQLReader.Close()
                Else
                    oCommand2.CommandText = "INSERT INTO ship_temp3 (eType, Customer, PN, ima02, MES_MODEL ,w" & l_week & ") VALUES (2,'"
                    oCommand2.CommandText += oReader.Item("tqa02") & "','" & oReader.Item("oeb04") & "','" & oReader.Item("ima02") & "','" & l_Model & "'," & oReader.Item("tc_cif_04") & ")"
                    Try
                        oCommand2.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End While
        End If
        oReader.Close()

        ' 預測訂單
        oCommand.CommandText = "select 3,tqa02,ta_opd14,ima02,tc_azn02,tc_azn05,tc_cif_04 from tc_cif_file left join opd_file on tc_cif_01 = opd01 and tc_cif_02 = opd05 "
        oCommand.CommandText += "left join ima_file on ta_opd14 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = 2 left join tc_azn_file on tc_cif_05 = tc_azn01 "
        oCommand.CommandText += "where tc_cif_05 between to_date('2018/05/12','yyyy/MM/dd') and to_date('2021/5/7','yyyy/MM/dd') and tc_cif_01 like 'FC%' "
        oReader = oCommand.ExecuteReader()
        Dim Section3C As Decimal = 0
        If oReader.HasRows() Then
            While oReader.Read()
                Section3C += 1
                Me.Label2.Text = "Forecast Order " & Section3C
                Me.Label2.Refresh()
                If IsDBNull(oReader.Item("ta_opd14")) Then
                    Continue While
                End If
                mSQLS1.CommandText = "select model from model_paravalue where parameter = 'ERP PN'  AND VALUE = '" & oReader.Item("ta_opd14") & "'"
                Dim l_Model As String = String.Empty
                l_Model = mSQLS1.ExecuteScalar()
                Dim l_week As Decimal = 0
                Dim l_year As Decimal = 0
                l_year = oReader.Item("tc_azn02")
                l_week = 53 * (l_year - 2018) + oReader.Item("tc_azn05")
                If ArrayS1.Contains(oReader.Item("ta_opd14").ToString()) Then
                    mSQLS1.CommandText = "SELECT * FROM ERPSUPPORT.dbo.SETVSPCS WHERE [SET] = '" & oReader.Item("ta_opd14") & "'"
                    mSQLReader = mSQLS1.ExecuteReader
                    If mSQLReader.HasRows() Then
                        While mSQLReader.Read()
                            mSQLS2.CommandText = "select model from model_paravalue where parameter = 'ERP PN'  AND VALUE = '" & mSQLReader.Item("PCS") & "'"
                            l_Model = mSQLS2.ExecuteScalar()
                            oCommand2.CommandText = "SELECT ima02 FROM ima_file WHERE  ima01 = '" & mSQLReader.Item("PCS") & "'"
                            Dim l_ima02 As String = String.Empty
                            l_ima02 = oCommand2.ExecuteScalar()
                            oCommand2.CommandText = "INSERT INTO ship_temp3 (eType, Customer, PN, ima02, MES_MODEL ,w" & l_week & ") VALUES (3,'"
                            oCommand2.CommandText += oReader.Item("tqa02") & "','" & mSQLReader.Item("PCS") & "','" & l_ima02 & "','" & l_Model & "'," & oReader.Item("tc_cif_04") & ")"
                            Try
                                oCommand2.ExecuteNonQuery()
                            Catch ex As Exception
                                MsgBox(ex.Message())
                            End Try
                        End While
                    End If
                    mSQLReader.Close()
                Else
                    oCommand2.CommandText = "INSERT INTO ship_temp3 (eType, Customer, PN, ima02, MES_MODEL ,w" & l_week & ") VALUES (3,'"
                    oCommand2.CommandText += oReader.Item("tqa02") & "','" & oReader.Item("ta_opd14") & "','" & oReader.Item("ima02") & "','" & l_Model & "'," & oReader.Item("tc_cif_04") & ")"
                    Try
                        oCommand2.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End While
        End If
        oReader.Close()

        '入庫計劃
        oCommand.CommandText = "select 5,tqa02,tc_prk01,ima02,tc_prk02,tc_prk03,tc_prk04 from tc_prk_file left join ima_file on tc_prk01 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = 2  "
        oCommand.CommandText += "where tc_prk02 between 2018 and 2021 and ima06 = '103' and ima08 ='M' and tc_prk01 not like 'S%' and tc_prk01 not like 'A%' and ima55 IN ('PCS','SET') "
        oReader = oCommand.ExecuteReader()
        Dim Section5C As Decimal = 0
        If oReader.HasRows Then
            While oReader.Read()
                If oReader.Item("tc_prk02") = 2018 And oReader.Item("tc_prk03") < 20 Then
                    Continue While
                End If
                Section5C += 1
                Me.Label2.Text = "Production Plan " & Section5C
                Me.Label2.Refresh()
                If IsDBNull(oReader.Item("tc_prk01")) Then
                    Continue While
                End If
                mSQLS1.CommandText = "select model from model_paravalue where parameter = 'ERP PN'  AND VALUE = '" & oReader.Item("tc_prk01") & "'"
                Dim l_Model As String = String.Empty
                l_Model = mSQLS1.ExecuteScalar()
                Dim l_week As Decimal = 0
                Dim l_year As Decimal = 0
                l_year = oReader.Item("tc_prk02")
                l_week = 53 * (l_year - 2018) + oReader.Item("tc_prk03")
                oCommand2.CommandText = "INSERT INTO ship_temp3 (eType, Customer, PN, ima02, MES_MODEL ,w" & l_week & ") VALUES (5,'"
                oCommand2.CommandText += oReader.Item("tqa02") & "','" & oReader.Item("tc_prk01") & "','" & oReader.Item("ima02") & "','" & l_Model & "'," & oReader.Item("tc_prk04") & ")"
                Try
                    oCommand2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End While
        End If
        oReader.Close()

        ' 實際入庫
        oCommand.CommandText = "select 6,tqa02,tlf01,ima02,tc_azn02,tc_azn05,sum(tlf10 * tlf12 * tlf907) as t1  from tlf_file "
        oCommand.CommandText += "left join ima_file on tlf01 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = 2  "
        oCommand.CommandText += "left join tc_azn_file on tlf06 = tc_azn01 where tlf06 between to_date('2018/05/12','yyyy/mm/dd') and to_date('2021/5/7','yyyy/mm/dd') and ima06 = '103' and ima08 ='M' "
        oCommand.CommandText += "and ima01 not like 'S%' and ima01 not like 'A%' and ima55 IN ('PCS','SET') and tlf13 = 'aimt324' and tlf902 = 'D146103' group by tqa02,tlf01,ima02,tc_azn02,tc_azn05"
        oReader = oCommand.ExecuteReader()
        Dim Section6C As Decimal = 0
        If oReader.HasRows() Then
            While oReader.Read()
                Section6C += 1
                Me.Label2.Text = "Actual Production " & Section6C
                Me.Label2.Refresh()
                If IsDBNull(oReader.Item("tlf01")) Then
                    Continue While
                End If
                mSQLS1.CommandText = "select model from model_paravalue where parameter = 'ERP PN'  AND VALUE = '" & oReader.Item("tlf01") & "'"
                Dim l_Model As String = String.Empty
                l_Model = mSQLS1.ExecuteScalar()
                Dim l_week As Decimal = 0
                Dim l_year As Decimal = 0
                l_year = oReader.Item("tc_azn02")
                l_week = 53 * (l_year - 2018) + oReader.Item("tc_azn05")
                If ArrayS1.Contains(oReader.Item("tlf01").ToString()) Then
                    mSQLS1.CommandText = "SELECT * FROM ERPSUPPORT.dbo.SETVSPCS WHERE [SET] = '" & oReader.Item("tlf01") & "'"
                    mSQLReader = mSQLS1.ExecuteReader
                    If mSQLReader.HasRows() Then
                        While mSQLReader.Read()
                            mSQLS2.CommandText = "select model from model_paravalue where parameter = 'ERP PN'  AND VALUE = '" & mSQLReader.Item("PCS") & "'"
                            l_Model = mSQLS2.ExecuteScalar()
                            oCommand2.CommandText = "SELECT ima02 FROM ima_file WHERE  ima01 = '" & mSQLReader.Item("PCS") & "'"
                            Dim l_ima02 As String = String.Empty
                            l_ima02 = oCommand2.ExecuteScalar()
                            oCommand2.CommandText = "INSERT INTO ship_temp3 (eType, Customer, PN, ima02, MES_MODEL ,w" & l_week & ") VALUES (6,'"
                            oCommand2.CommandText += oReader.Item("tqa02") & "','" & mSQLReader.Item("PCS") & "','" & l_ima02 & "','" & l_Model & "'," & oReader.Item("t1") & ")"
                            Try
                                oCommand2.ExecuteNonQuery()
                            Catch ex As Exception
                                MsgBox(ex.Message())
                            End Try
                        End While
                    End If
                    mSQLReader.Close()
                Else
                    oCommand2.CommandText = "INSERT INTO ship_temp3 (eType, Customer, PN, ima02, MES_MODEL ,w" & l_week & ") VALUES (6,'"
                    oCommand2.CommandText += oReader.Item("tqa02") & "','" & oReader.Item("tlf01") & "','" & oReader.Item("ima02") & "','" & l_Model & "'," & oReader.Item("t1") & ")"
                    Try
                        oCommand2.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End While
        End If
        oReader.Close()

        ' 實際出貨
        oCommand.CommandText = "select 7,tqa02,ogb04,ima02,tc_azn02,tc_azn05,ogb12 from ogb_file left join oga_file on ogb01 = oga01 "
        oCommand.CommandText += "left join ima_file on ogb04 = ima01 left join tqa_file on ima1005 = tqa01 and tqa03 = 2 left join tc_azn_file on oga02 = tc_azn01 "
        oCommand.CommandText += "where  ogaconf = 'Y' and ima06 = '103' and ima08 ='M' and ima01 not like 'S%' and ima01 not like 'A%' and ima55 IN ('PCS','SET') and ogapost = 'Y' and oga02 between to_date('2018/5/12','yyyy/MM/dd') and  to_date('2021/5/7','yyyy/MM/dd')"
        oReader = oCommand.ExecuteReader()
        Dim Section7C As Decimal = 0
        If oReader.HasRows() Then
            While oReader.Read()
                Section7C += 1
                Me.Label2.Text = "Actual Shipment " & Section7C
                Me.Label2.Refresh()
                If IsDBNull(oReader.Item("ogb04")) Then
                    Continue While
                End If
                mSQLS1.CommandText = "select model from model_paravalue where parameter = 'ERP PN'  AND VALUE = '" & oReader.Item("ogb04") & "'"
                Dim l_Model As String = String.Empty
                l_Model = mSQLS1.ExecuteScalar()
                Dim l_week As Decimal = 0
                Dim l_year As Decimal = 0
                l_year = oReader.Item("tc_azn02")
                l_week = 53 * (l_year - 2018) + oReader.Item("tc_azn05")
                If ArrayS1.Contains(oReader.Item("ogb04").ToString()) Then
                    mSQLS1.CommandText = "SELECT * FROM ERPSUPPORT.dbo.SETVSPCS WHERE [SET] = '" & oReader.Item("ogb04") & "'"
                    mSQLReader = mSQLS1.ExecuteReader
                    If mSQLReader.HasRows() Then
                        While mSQLReader.Read()
                            mSQLS2.CommandText = "select model from model_paravalue where parameter = 'ERP PN'  AND VALUE = '" & mSQLReader.Item("PCS") & "'"
                            l_Model = mSQLS2.ExecuteScalar()
                            oCommand2.CommandText = "SELECT ima02 FROM ima_file WHERE  ima01 = '" & mSQLReader.Item("PCS") & "'"
                            Dim l_ima02 As String = String.Empty
                            l_ima02 = oCommand2.ExecuteScalar()
                            oCommand2.CommandText = "INSERT INTO ship_temp3 (eType, Customer, PN, ima02, MES_MODEL ,w" & l_week & ") VALUES (7,'"
                            oCommand2.CommandText += oReader.Item("tqa02") & "','" & mSQLReader.Item("PCS") & "','" & l_ima02 & "','" & l_Model & "'," & oReader.Item("ogb12") & ")"
                            Try
                                oCommand2.ExecuteNonQuery()
                            Catch ex As Exception
                                MsgBox(ex.Message())
                            End Try
                        End While
                    End If
                    mSQLReader.Close()
                Else
                    oCommand2.CommandText = "INSERT INTO ship_temp3 (eType, Customer, PN, ima02, MES_MODEL ,w" & l_week & ") VALUES (7,'"
                    oCommand2.CommandText += oReader.Item("tqa02") & "','" & oReader.Item("ogb04") & "','" & oReader.Item("ima02") & "','" & l_Model & "'," & oReader.Item("ogb12") & ")"
                    Try
                        oCommand2.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End While
        End If
        oReader.Close()

        ' 期初庫存
        oCommand.CommandText = "select 8,tqa02,tc_ini02,ima02,tc_ini03 from tc_ini_file left join ima_file on tc_ini02 = ima01 "
        oCommand.CommandText += "left join tqa_file on ima1005 = tqa01 and tqa03 = 2 where  tc_ini01 = 2"
        oReader = oCommand.ExecuteReader()
        Dim Section8C As Decimal = 0
        If oReader.HasRows() Then
            While oReader.Read()
                Section8C += 1
                Me.Label2.Text = "Stock " & Section8C
                Me.Label2.Refresh()
                If IsDBNull(oReader.Item("tc_ini02")) Then
                    Continue While
                End If
                mSQLS1.CommandText = "select model from model_paravalue where parameter = 'ERP PN'  AND VALUE = '" & oReader.Item("tc_ini02") & "'"
                Dim l_Model As String = String.Empty
                l_Model = mSQLS1.ExecuteScalar()
                oCommand2.CommandText = "INSERT INTO ship_temp3 (eType, Customer, PN, ima02, MES_MODEL ) VALUES (8,'"
                oCommand2.CommandText += oReader.Item("tqa02") & "','" & oReader.Item("tc_ini02") & "','" & oReader.Item("ima02") & "','" & l_Model & "')"
                Try
                    oCommand2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End While
        End If
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        Dim xPath As String = "C:\temp\Demand and schedule.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        LineZ = 4

        oCommand.CommandText = "select distinct pn,customer,ima02,MES_MODEL from ship_temp3 order by MES_MODEL"
        oReader = oCommand.ExecuteReader()
        Dim TotalRow As Int16 = 0
        If oReader.HasRows() Then
            While oReader.Read()
                TotalRow += 1
                For i As Int16 = 1 To 11 Step 1
                    If i < 4 Or (i > 4 And i < 8) Then
                        oCommand2.CommandText = "select etype"
                        For j As Int16 = 20 To 178 Step 1
                            oCommand2.CommandText += ",sum(w" & j & ") as w" & j
                        Next
                        oCommand2.CommandText += " from ship_temp3 where etype = " & i & " and pn = '" & oReader.Item("pn") & "' group by etype"

                    End If

                    Ws.Cells(LineZ, 1) = TotalRow
                    Ws.Cells(LineZ, 2) = oReader.Item("customer")
                    Ws.Cells(LineZ, 3) = oReader.Item("pn")
                    Ws.Cells(LineZ, 4) = oReader.Item("ima02")
                    Ws.Cells(LineZ, 6) = oReader.Item("MES_MODEL")

                    Select Case i
                        Case 1
                            Ws.Cells(LineZ, 5) = "Sales order"
                            Ws.Cells(LineZ, 7) = "销售订单"
                            Ws.Cells(LineZ, 8) = 0
                        Case 2
                            Ws.Cells(LineZ, 5) = "Customer demand"
                            Ws.Cells(LineZ, 7) = "客户需求"
                            Ws.Cells(LineZ, 8) = 0
                        Case 3
                            Ws.Cells(LineZ, 5) = "Forecast Order"
                            Ws.Cells(LineZ, 7) = "預測订单"
                            Ws.Cells(LineZ, 8) = 0
                        Case 4
                            Ws.Cells(LineZ, 5) = "Shipping Backlog"
                            Ws.Cells(LineZ, 7) = "需求差异"
                            oCommand3.CommandText = "select nvl(tc_ini03,0) from tc_ini_file where tc_ini01 = 1 and tc_ini02 = '" & oReader.Item("pn") & "'"
                            Dim l_short As Decimal = oCommand3.ExecuteScalar()
                            Ws.Cells(LineZ, 8) = l_short
                            Ws.Cells(LineZ, 9) = "=H" & LineZ & "+I" & LineZ + 3 & "-I" & LineZ - 2 & "-I" & LineZ - 1
                            oRng = Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9))
                            oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 167)), Type:=xlFillDefault)
                        Case 5
                            Ws.Cells(LineZ, 5) = "Production completion plan"
                            Ws.Cells(LineZ, 7) = "入庫計划"
                            Ws.Cells(LineZ, 8) = 0
                        Case 6
                            Ws.Cells(LineZ, 5) = "Actual Production"
                            Ws.Cells(LineZ, 7) = "实际入庫"
                            Ws.Cells(LineZ, 8) = 0
                        Case 7
                            Ws.Cells(LineZ, 5) = "Actual Shipment"
                            Ws.Cells(LineZ, 7) = "实际出货"
                            Ws.Cells(LineZ, 8) = 0
                        Case 8
                            Ws.Cells(LineZ, 5) = "Production Backlog"
                            Ws.Cells(LineZ, 7) = "生产欠数"
                            oCommand3.CommandText = "select nvl(tc_ini03,0) from tc_ini_file where tc_ini01 = 4 and tc_ini02 = '" & oReader.Item("pn") & "'"
                            Dim l_short As Decimal = oCommand3.ExecuteScalar()
                            Ws.Cells(LineZ, 8) = l_short
                            Ws.Cells(LineZ, 9) = "=H" & LineZ & "+I" & LineZ - 7 & "-I" & LineZ - 2
                            oRng = Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9))
                            oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 167)), Type:=xlFillDefault)
                        Case 9
                            Ws.Cells(LineZ, 5) = "Stock FG  Q'ty"
                            Ws.Cells(LineZ, 7) = "成品庫存"
                            oCommand3.CommandText = "select nvl(tc_ini03,0) from tc_ini_file where tc_ini01 = 2 and tc_ini02 = '" & oReader.Item("pn") & "'"
                            Dim l_short As Decimal = oCommand3.ExecuteScalar()
                            Ws.Cells(LineZ, 8) = l_short
                            Ws.Cells(LineZ, 9) = "=H" & LineZ & "+I" & LineZ - 3 & "-I" & LineZ - 2
                            oRng = Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9))
                            oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 167)), Type:=xlFillDefault)
                        Case 10
                            Ws.Cells(LineZ, 5) = "Order  remaining"
                            Ws.Cells(LineZ, 7) = "订单余量"
                            oCommand3.CommandText = "select nvl(tc_ini03,0) from tc_ini_file where tc_ini01 = 3 and tc_ini02 = '" & oReader.Item("pn") & "'"
                            Dim l_short As Decimal = oCommand3.ExecuteScalar()
                            Ws.Cells(LineZ, 8) = l_short
                            Ws.Cells(LineZ, 9) = "=H" & LineZ & "+I" & LineZ - 9 & "-I" & LineZ - 3
                            oRng = Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9))
                            oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 167)), Type:=xlFillDefault)
                        Case 11
                            Ws.Cells(LineZ, 5) = "Plan & Actual differences"
                            Ws.Cells(LineZ, 7) = "计划差异"
                            Ws.Cells(LineZ, 9) = "=H" & LineZ & "+I" & LineZ - 5 & "-I" & LineZ - 6
                            oRng = Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9))
                            oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 167)), Type:=xlFillDefault)
                    End Select
                    If i < 4 Or (i > 4 And i < 8) Then
                        oReader2 = oCommand2.ExecuteReader()
                        If oReader2.HasRows() Then
                            While oReader2.Read()
                                For k As Integer = 1 To oReader2.FieldCount - 1 Step 1
                                    If Not IsDBNull(oReader2.Item(k)) Then
                                        Ws.Cells(LineZ, 8 + k) = oReader2.Item(k)
                                    End If
                                Next
                            End While
                        Else
                        End If
                        oReader2.Close()

                    End If
                    If Decimal.Remainder(TotalRow, 2) = 1 Then
                        oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 1))
                        oRng.EntireRow.Interior.Color = Color.FromArgb(217, 217, 217)
                    End If
                    LineZ += 1
                    Me.Label2.Text = "制作报表中" & LineZ
                    Me.Label2.Refresh()
                Next
            End While
        End If
        oReader.Close()
        oRng = Ws.Range("A4", "JM" & LineZ - 1)
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng.ShrinkToFit = True
        oRng.NumberFormat = "0_ ;[Red]-0 "

    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "需求和进度" & Now.ToString("yyyyMMddHHmm")
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
        If oConnection.State = ConnectionState.Open Then
            Try
                oConnection.Close()
                mConnection.Close()
                Module1.KillExcelProcess(OldExcel)
                MsgBox("Finished")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
End Class