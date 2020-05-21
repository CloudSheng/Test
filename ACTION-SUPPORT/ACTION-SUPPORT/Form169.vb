Public Class Form169
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim DocDate As Date
    Dim PaperDate As Date
    Dim oma01 As String = String.Empty
    Dim omb03 As Integer = 0
    Dim apa01 As String = String.Empty
    Dim apb02 As Integer = 0
    Dim Apa12 As Date
    Dim CompareSign2 As String = String.Empty
    'Dim Hac_Invoice_No As String = String.Empty       '191205 add by Brady
    'Dim Vac_Invoice_No As String = String.Empty       '191205 add by Brady

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString As New OleDb.OleDbCommand
            ExcelString.CommandText = "SELECT * FROM [Sheet1$] WHERE new_add_item = 'Y'"
            ExcelString.Connection = Excelconn
            Dim ExcelDataReader As OleDb.OleDbDataReader = ExcelString.ExecuteReader
            If ExcelDataReader.HasRows() Then
                Dim Tran1 As Oracle.ManagedDataAccess.Client.OracleTransaction = oConnection.BeginTransaction()
                oCommand.Transaction = Tran1
                oCommander2.Transaction = Tran1
                DocDate = "2019/01/01"
                CompareSign2 = ""
                'Hac_Invoice_No = " "                  '191205 add by Brady
                'Vac_Invoice_No = " "                  '191205 add by Brady
                While ExcelDataReader.Read()
                    If IsDBNull(ExcelDataReader.Item(2)) Then
                        Continue While
                    End If
                    Dim l_FOC As String = String.Empty
                    If Not IsDBNull(ExcelDataReader.Item(24)) Then
                        l_FOC = ExcelDataReader.Item(24)
                    End If
                    Dim CompareSign As String = ExcelDataReader.Item(2) & ExcelDataReader.Item(3) & ExcelDataReader.Item(6)
                    'MsgBox(ExcelDataReader.Item(1))
                    If IsDBNull(DocDate) Or CompareSign <> CompareSign2 Then  ' 建立表頭
                        DocDate = Convert.ToDateTime(ExcelDataReader.Item(2))
                        CompareSign2 = ExcelDataReader.Item(2) & ExcelDataReader.Item(3) & ExcelDataReader.Item(6)
                        PaperDate = DocDate
                        oma01 = String.Empty
                        oma01 = Getoma01()
                        If l_FOC <> "FOC" Then
                            'If (IsDBNull(Hac_Invoice_No) Or Hac_Invoice_No <> Convert.ToString(ExcelDataReader.Item(6))) Then              '191205 add by Brady
                            oCommand.CommandText = "INSERT INTO oma_file VALUES ('14','" & oma01 & "',to_date('" & DocDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),'D0003','Austria Action','D0003',NULL,'Action Composites GmbH','Action Composites GmbH','1',NULL,"
                            '200102 add by Brady
                            'oCommand.CommandText += "'N','2',to_date('" & DocDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),NULL,to_date('" & DocDate.AddDays(120).ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),to_date('" & DocDate.AddDays(120).ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),"
                            oCommand.CommandText += "'N','2',to_date('" & DocDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),'" & ExcelDataReader.Item(6) & "',to_date('" & DocDate.AddDays(120).ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),to_date('" & DocDate.AddDays(120).ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),"
                            '200102 add by Brady
                            oCommand.CommandText += "'002','HK00001','D0180',NULL,0,100,0,'1','XX','3',NULL,NULL,NULL,'112202',NULL,'Y','S02',0,'C','N','" & ExcelDataReader.Item(19) & "',1,'001',NULL,'11',NULL,NULL,NULL,NULL,NULL,NULL,NULL,'Y'," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(17) & ",0,0"
                            oCommand.CommandText += "," & ExcelDataReader.Item(17) & ",0," & ExcelDataReader.Item(17) & ",0," & ExcelDataReader.Item(17) & ",0," & ExcelDataReader.Item(17) & ",0,1," & ExcelDataReader.Item(17) & ",0," & ExcelDataReader.Item(17) & ",1," & ExcelDataReader.Item(17) & ",NULL,NULL,NULL,'N','N',NULL,"
                            oCommand.CommandText += "0,'HK00001','D0180','HK00001',to_date('" & DocDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),'0','N','1',0,0,'HKACTTEST',NULL,NULL,NULL,'D0003','Austria Action',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'2','HKACTTTEST',"
                            oCommand.CommandText += "NULL,'HK00001','D0180',NULL,NULL,0,0,1,NULL)"

                            Try
                                oCommand.ExecuteNonQuery()
                            Catch ex As Exception
                                Tran1.Rollback()
                                MsgBox(ex.Message())
                                Exit While
                            End Try
                            omb03 = 1
                            '191205 add by Brady
                            'oCommand.CommandText = "INSERT INTO omb_file VALUES ('14','" & oma01 & "',1,NULL,'" & ExcelDataReader.Item(13) & "','" & ExcelDataReader.Item(7) & "'," & ExcelDataReader.Item(12) & "," & ExcelDataReader.Item(15) & "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(15)
                            'omb01-omb15
                            oCommand.CommandText = "INSERT INTO omb_file VALUES ('14','" & oma01 & "',1,'" & ExcelDataReader.Item(27) & "','" & ExcelDataReader.Item(13) & "','" & ExcelDataReader.Item(7) & "'," & ExcelDataReader.Item(12) & "," & ExcelDataReader.Item(15) & "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(15)
                            '191205 add by Brady END

                            '200321 add by Brady
                            'oCommand.CommandText += "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(15) & "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(17) & ",NULL,NULL,NULL,0,0,NULL,NULL,NULL,NULL,'603','99','N',NULL,NULL,NULL,'" & ExcelDataReader.Item(10) & "',"
                            'omb16-ombud01
                            oCommand.CommandText += "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(15) & "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(17) & ",NULL,NULL,'600101',0,0,NULL,NULL,NULL,NULL,'603','99','N',NULL,NULL,NULL,'" & ExcelDataReader.Item(10) & "',"
                            '200321 add by Brady END

                            oCommand.CommandText += "NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'HKACTTTEST','HKACTTEST',NULL,NULL,NULL)"
                            Try
                                oCommand.ExecuteNonQuery()
                            Catch ex As Exception
                                Tran1.Rollback()
                                MsgBox(ex.Message())
                                Exit While
                            End Try

                            oCommand.CommandText = "INSERT INTO omc_file VALUES ('" & oma01 & "',1,'02',to_date('" & DocDate.AddDays(60).ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),to_date('" & DocDate.AddDays(60).ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),1,1," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(17)
                            oCommand.CommandText += ",0,0,NULL," & ExcelDataReader.Item(17) & ",0,0,'HKACTTTEST')"
                            Try
                                oCommand.ExecuteNonQuery()
                            Catch ex As Exception
                                Tran1.Rollback()
                                MsgBox(ex.Message())
                                Exit While
                            End Try
                            'End If                                                                                                          '191205 add by Brady
                        End If
                        ' AP 處理 20190610
                        'If (IsDBNull(Vac_Invoice_No) Or Vac_Invoice_No <> Convert.ToString(ExcelDataReader.Item(7))) Then                 '191205 add by Brady
                        apa01 = String.Empty
                        apa01 = Getapa01()
                        Apa12 = GetApa12()
                        oCommand.CommandText = "INSERT INTO apa_file VALUES ('12','" & apa01 & "',to_date('" & DocDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') ,'D0005','D0005','VN Action','" & apa01 & "',to_date('" & DocDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') ,'12',to_date('" & Apa12.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') ,'USD',1,"
                        oCommand.CommandText += "'P07',0,'3','XX','3',NULL,NULL,NULL,NULL,NULL,0,'HK00001','D0180',NULL,0,'" & ExcelDataReader.Item(7) & "',0,0,0,0,0,0,0,0,0,0,'003','N','N',NULL,NULL,NULL,NULL,NULL,NULL,NULL,'220202','1','0',0,0,NULL,NULL,0,0,0,0,NULL,'0',to_date('" & Apa12.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') ,0,0,NULL,NULL,NULL,NULL,NULL,NULL,1,0,"
                        oCommand.CommandText += "'N','N',NULL,to_date('" & Today.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') , 'N',NULL,NULL,NULL,NULL,NULL,0,'Y','HK00001','D0180','HK00001',to_date('" & Today.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') ,'HKACTTEST',NULL,NULL,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'1','HKACTTTEST',NULL,'HK00001','D0180',NULL,'0')"
                        Try
                            oCommand.ExecuteNonQuery()
                        Catch ex As Exception
                            Tran1.Rollback()
                            MsgBox(ex.Message())
                            Exit While
                        End Try

                        apb02 = 1
                        Dim Tprice As Decimal = ExcelDataReader.Item(14)
                        Dim TQty As Decimal = ExcelDataReader.Item(12)
                        Dim TAmount As Decimal = Tprice * TQty
                        '191205,191206 add by Brady
                        'oCommand.CommandText = "INSERT INTO apb_file VALUES ('" & apa01 & "'," & apb02 & ",NULL,NULL,0,NULL,NULL," & Tprice & "," & Tprice & "," & TQty & "," & TAmount & "," & TAmount & ",NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," & Tprice & "," & TAmount & ",'1405',NULL,'" & ExcelDataReader.Item(9) & "','" & ExcelDataReader.Item(13) & "','1',NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,'" & ExcelDataReader.Item(7) & "',"
                        oCommand.CommandText = "INSERT INTO apb_file VALUES ('" & apa01 & "'," & apb02 & ",NULL,NULL,0,NULL,NULL," & Tprice & "," & Tprice & "," & TQty & "," & TAmount & "," & TAmount & ",NULL,'" & ExcelDataReader.Item(27) & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," & Tprice & "," & TAmount & ",'1405',NULL,'" & ExcelDataReader.Item(9) & "','" & ExcelDataReader.Item(13) & "','1',NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,'" & ExcelDataReader.Item(7) & "',"
                        '191205,191206 add by Brady END
                        oCommand.CommandText += "NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'HKACTTTEST','HKACTTEST')"
                        Try
                            oCommand.ExecuteNonQuery()
                        Catch ex As Exception
                            Tran1.Rollback()
                            MsgBox(ex.Message())
                            Exit While
                        End Try

                        oCommand.CommandText = "INSERT INTO apc_file VALUES ('" & apa01 & "'," & apb02 & ",'09',to_date('" & Apa12.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') ,to_date('" & Apa12.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') ,1,1,0,0,0,0,'" & apa01 & "',0,0,0,0,'HKACTTTEST')"
                        Try
                            oCommand.ExecuteNonQuery()
                        Catch ex As Exception
                            Tran1.Rollback()
                            MsgBox(ex.Message())
                            Exit While
                        End Try
                        WriteBack1(apa01)
                        'End If                                                                                                               '191205 add by Brady
                    Else
                        If l_FOC <> "FOC" Then
                            'If (IsDBNull(Hac_Invoice_No) Or Hac_Invoice_No <> Convert.ToString(ExcelDataReader.Item(6))) Then              '191205 add by Brady
                            oCommand.CommandText = "Select Count(*) from oma_file where oma01 = '" & oma01 & "' "
                            Dim HaveHead As Int16 = oCommand.ExecuteScalar()
                            If HaveHead = 0 Then
                                oCommand.CommandText = "INSERT INTO oma_file VALUES ('14','" & oma01 & "',to_date('" & DocDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),'D0003','Austria Action','D0003',NULL,'Action Composites GmbH','Action Composites GmbH','1',NULL,"
                                '200102 add by Brady
                                'oCommand.CommandText += "'N','2',to_date('" & DocDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),NULL,to_date('" & DocDate.AddDays(120).ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),to_date('" & DocDate.AddDays(120).ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),"
                                oCommand.CommandText += "'N','2',to_date('" & DocDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')," & ExcelDataReader.Item(6) & ",to_date('" & DocDate.AddDays(120).ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),to_date('" & DocDate.AddDays(120).ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),"
                                '200102 add by Brady END
                                oCommand.CommandText += "'002','HK00001','D0180',NULL,0,100,0,'1','XX','3',NULL,NULL,NULL,'112202',NULL,'Y','S02',0,'C','N','" & ExcelDataReader.Item(19) & "',1,'001',NULL,'11',NULL,NULL,NULL,NULL,NULL,NULL,NULL,'Y'," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(17) & ",0,0"
                                oCommand.CommandText += "," & ExcelDataReader.Item(17) & ",0," & ExcelDataReader.Item(17) & ",0," & ExcelDataReader.Item(17) & ",0," & ExcelDataReader.Item(17) & ",0,1," & ExcelDataReader.Item(17) & ",0," & ExcelDataReader.Item(17) & ",1," & ExcelDataReader.Item(17) & ",NULL,NULL,NULL,'N','N',NULL,"
                                oCommand.CommandText += "0,'HK00001','D0180','HK00001',to_date('" & DocDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd'),'0','N','1',0,0,'HKACTTEST',NULL,NULL,NULL,'D0003','Austria Action',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'2','HKACTTTEST',"
                                oCommand.CommandText += "NULL,'HK00001','D0180',NULL,NULL,0,0,1,NULL)"

                                Try
                                    oCommand.ExecuteNonQuery()
                                    omb03 = 0
                                Catch ex As Exception
                                    Tran1.Rollback()
                                    MsgBox(ex.Message())
                                    Exit While
                                End Try
                            End If
                            omb03 += 1
                            '191205 add by Brady
                            'oCommand.CommandText = "INSERT INTO omb_file VALUES ('14','" & oma01 & "'," & omb03 & ",NULL,'" & ExcelDataReader.Item(13) & "','" & ExcelDataReader.Item(7) & "'," & ExcelDataReader.Item(12) & "," & ExcelDataReader.Item(15) & "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(15)
                            'omb01-omb15
                            oCommand.CommandText = "INSERT INTO omb_file VALUES ('14','" & oma01 & "'," & omb03 & ",'" & ExcelDataReader.Item(27) & "','" & ExcelDataReader.Item(13) & "','" & ExcelDataReader.Item(7) & "'," & ExcelDataReader.Item(12) & "," & ExcelDataReader.Item(15) & "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(15)
                            '191205 add by Brady END

                            '200321 add by Brady
                            'oCommand.CommandText += "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(15) & "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(17) & ",NULL,NULL,NULL,0,0,NULL,NULL,NULL,NULL,'603','99','N',NULL,NULL,NULL,'" & ExcelDataReader.Item(10) & "',"
                            'omb16-ombud01
                            oCommand.CommandText += "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(15) & "," & ExcelDataReader.Item(17) & "," & ExcelDataReader.Item(17) & ",NULL,NULL,'600101',0,0,NULL,NULL,NULL,NULL,'603','99','N',NULL,NULL,NULL,'" & ExcelDataReader.Item(10) & "',"
                            '200321 add by Brady END

                            oCommand.CommandText += "NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'HKACTTTEST','HKACTTEST',NULL,NULL,NULL)"
                            Try
                                oCommand.ExecuteNonQuery()
                            Catch ex As Exception
                                Tran1.Rollback()
                                MsgBox(ex.Message())
                                Exit While
                            End Try
                            WriteBack(oma01)
                            'End If                                                                                                          '191205 add by Brady
                        End If
                        'If (IsDBNull(Vac_Invoice_No) Or Vac_Invoice_No <> Convert.ToString(ExcelDataReader.Item(7))) Then                 '191205 add by Brady
                        apb02 += 1
                        Dim Tprice As Decimal = ExcelDataReader.Item(14)
                        Dim TQty As Decimal = ExcelDataReader.Item(12)
                        Dim TAmount As Decimal = Tprice * TQty
                        '191205,191206 add by Brady
                        'oCommand.CommandText = "INSERT INTO apb_file VALUES ('" & apa01 & "'," & apb02 & ",NULL,NULL,0,NULL,NULL," & Tprice & "," & Tprice & "," & TQty & "," & TAmount & "," & TAmount & ",NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," & Tprice & "," & TAmount & ",'1405',NULL,'" & ExcelDataReader.Item(9) & "','" & ExcelDataReader.Item(13) & "','1',NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,'" & ExcelDataReader.Item(7) & "',"
                        oCommand.CommandText = "INSERT INTO apb_file VALUES ('" & apa01 & "'," & apb02 & ",NULL,NULL,0,NULL,NULL," & Tprice & "," & Tprice & "," & TQty & "," & TAmount & "," & TAmount & ",NULL,'" & ExcelDataReader.Item(27) & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," & Tprice & "," & TAmount & ",'1405',NULL,'" & ExcelDataReader.Item(9) & "','" & ExcelDataReader.Item(13) & "','1',NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,'" & ExcelDataReader.Item(7) & "',"
                        '191205,191206 add by Brady END
                        oCommand.CommandText += "NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'HKACTTTEST','HKACTTEST')"
                        Try
                            oCommand.ExecuteNonQuery()
                        Catch ex As Exception
                            Tran1.Rollback()
                            MsgBox(ex.Message())
                            Exit While
                        End Try
                        WriteBack1(apa01)
                        'End If                                                                                                          '191205 add by Brady
                    End If
                End While
                If Not IsDBNull(Tran1.Connection) Then
                    Tran1.Commit()
                End If
            End If
            ExcelDataReader.Close()
            MsgBox("Done")
        End If
    End Sub

    Private Sub Form169_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("hkacttest")
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommander2.Connection = oConnection
                oCommander2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub
    Public Function Getoma01()
        Dim AB As String = String.Empty
        AB = "HR141-" & PaperDate.ToString("yy") & PaperDate.ToString("MM")
        oCommander2.CommandText = "select nvl(MAX(SUBSTR(oma01,11,4)),0) from oma_file where oma01 LIKE '" & AB & "%'"
        Dim MaxInt As Integer = oCommander2.ExecuteScalar()
        MaxInt += 1
        Select Case Strings.Len(MaxInt.ToString())
            Case 1
                AB = AB & "000" & MaxInt
            Case 2
                AB = AB & "00" & MaxInt
            Case 3
                AB = AB & "0" & MaxInt
            Case 4
                AB = AB & MaxInt
        End Select
        Return AB
    End Function
    Public Function Getapa01()
        Dim AB As String = String.Empty
        AB = "HP121-" & PaperDate.ToString("yy") & PaperDate.ToString("MM")
        oCommander2.CommandText = "select nvl(MAX(SUBSTR(apa01,11,4)),0) from apa_file where apa01 LIKE '" & AB & "%'"
        Dim MaxInt As Integer = oCommander2.ExecuteScalar()
        MaxInt += 1
        Select Case Strings.Len(MaxInt.ToString())
            Case 1
                AB = AB & "000" & MaxInt
            Case 2
                AB = AB & "00" & MaxInt
            Case 3
                AB = AB & "0" & MaxInt
            Case 4
                AB = AB & MaxInt
        End Select
        Return AB
    End Function
    Private Sub WriteBack(ByVal oma01 As String)
        oCommander2.CommandText = "SELECT nvl(SUM(omb14),0) from hkacttest.omb_file where omb01 = '" & oma01 & "'"
        Dim TotalV As Decimal = oCommander2.ExecuteScalar()
        oCommander2.CommandText = "UPDATE oma_file SET oma50 = " & TotalV & ",oma50t = " & TotalV & ",oma54 =" & TotalV & ",oma54t =" & TotalV & ",oma56 =" & TotalV & ",oma56t = " & TotalV
        oCommander2.CommandText += ",oma59 =" & TotalV & ",oma59t = " & TotalV & ",oma61 = " & TotalV & " WHERE oma01 = '" & oma01 & "'"
        Try
            oCommander2.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

        oCommander2.CommandText = "UPDATE omc_file SET omc08 = " & TotalV & ",omc09 = " & TotalV & ",omc13 =" & TotalV & " WHERE omc01 = '" & oma01 & "'"
        Try
            oCommander2.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
    End Sub
    Private Function GetApa12()
        Dim l_date As Int16 = PaperDate.Day
        oCommander2.CommandText = "Select apz56 from apz_file "
        Dim m_date As Int16 = oCommander2.ExecuteScalar()
        oCommander2.CommandText = "select pma08 from pma_file where pma01 = '12'"
        Dim DA As Int16 = oCommander2.ExecuteScalar()   ' 要加上的天數
        Dim DecideDate As New Date
        If l_date > m_date Then   ' 已過結帳日, 以下個月底計算
            DecideDate = PaperDate.AddDays((l_date - 1) * Decimal.MinusOne)
            DecideDate = DecideDate.AddMonths(2).AddDays(-1)
            DecideDate = DecideDate.AddDays(DA)
        Else  ' 未過結帳日, 以AP 日計算
            DecideDate = PaperDate.AddDays(DA)
        End If
        If Weekday(DecideDate, FirstDayOfWeek.Sunday) = 1 Then
            DecideDate = DecideDate.AddDays(1)
        End If
        Return DecideDate
    End Function
    Private Sub WriteBack1(ByVal apa01 As String)
        oCommander2.CommandText = "SELECT nvl(SUM(apb24),0) from hkacttest.apb_file where apb01 = '" & apa01 & "'"
        Dim TotalV As Decimal = oCommander2.ExecuteScalar()
        oCommander2.CommandText = "UPDATE apa_file SET apa31f = " & TotalV & ",apa34f = " & TotalV & ",apa31 =" & TotalV & ",apa34 =" & TotalV & ",apa57f =" & TotalV & ",apa57 = " & TotalV
        oCommander2.CommandText += ",apa73 =" & TotalV & " WHERE apa01 = '" & apa01 & "'"
        Try
            oCommander2.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

        oCommander2.CommandText = "UPDATE apc_file SET apc08 = " & TotalV & ",apc09 = " & TotalV & ",apc13 =" & TotalV & " WHERE apc01 = '" & apa01 & "'"
        Try
            oCommander2.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
    End Sub
End Class