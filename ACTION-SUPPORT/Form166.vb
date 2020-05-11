Public Class Form166
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand3 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim sDate1 As Date
    Dim eDate1 As Date
    Dim pYear As Int16 = 0
    Dim pMonth As Int16 = 0
    Private Sub Form166_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
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
        tYear = DateTimePicker1.Value.Year
        tMonth = DateTimePicker1.Value.Month
        sDate1 = Convert.ToDateTime(tYear & "/" & tMonth & "/01")
        eDate1 = sDate1.AddMonths(1).AddDays(-1)
        pMonth = tMonth - 1
        If pMonth = 0 Then
            oCommand.CommandText = "SELECT COUNT(*) FROM ACA_FIFO_RECORD WHERE year1 = " & tYear & " AND month1 = 0"
            Dim ZeroData As Decimal = oCommand.ExecuteScalar()
            If ZeroData = 0 Then
                pMonth = 12
                pYear = tYear - 1
            Else
                pYear = tYear
            End If
        Else
            pYear = tYear
        End If

        oCommand.CommandText = "DELETE ACA_FIFO_RECORD WHERE year1 =" & tYear & " AND month1 =" & tMonth
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception

        End Try

        oCommand.CommandText = "INSERT INTO ACA_FIFO_RECORD Select " & tYear & "," & tMonth & "," & "PN, LOT, SUM(QTY),COSTPRICE FROM ( Select " & tYear & "," & tMonth & ", PN, LOT ,QTY,COSTPRICE from aca_fifo_record where year1 = " & pYear & " and month1 = " & pMonth & " "
        oCommand.CommandText += "UNION ALL "
        oCommand.CommandText += "select " & tYear & ", " & tMonth & ", PN1, D1,SUM(QUANTITY) AS T1, nvl(ccc23,0) FROM ( select (case when s2.dacpn is null then s1.pn else s2.dacpn end) as PN1, to_char(OGA02, 'yyyymmdd') as d1, QUANTITy from Aca_Goods_Received s1 "
        oCommand.CommandText += "left join oga_file on INVOICENO = oga27 left join aca_pn s2 on s1.pn = s2.acapn where Date1 between to_date('"
        oCommand.CommandText += sDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and oga02 is not null " 'oga02 is not null add on 20200421
        oCommand.CommandText += "union all "
        oCommand.CommandText += "Select (case when s2.dacpn is null then s1.pn else s2.dacpn end) as PN1, to_char(date1,'yyyymmdd') , Quantity from ACA_Shipment_Return  s1 left join aca_pn s2 on s1.pn = s2.acapn where date1 between to_date('"
        oCommand.CommandText += sDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') ) AB "
        oCommand.CommandText += "LEFT JOIN ccc_file on PN1 = ccc_file.ccc01 and ccc02 = " & tYear & " and ccc03 = " & tMonth & " GROUP BY PN1,D1, ccc23 ) group by pn,lot,costprice"


        'oCommand.CommandText = "INSERT INTO ACA_FIFO_RECORD Select " & tYear & "," & tMonth & ", PN, LOT ,QTY,COSTPRICE from aca_fifo_record where year1 = " & pYear & " and month1 = " & pMonth & " "
        'oCommand.CommandText += "UNION ALL "
        'oCommand.CommandText += "select " & tYear & ", " & tMonth & ", PN1, D1,SUM(QUANTITY) AS T1, nvl(ccc23,0) FROM ( select (case when s2.dacpn is null then s1.pn else s2.dacpn end) as PN1, to_char(Date1,'yyyymmdd') as d1, QUANTITy from Aca_Goods_Received s1 "
        'oCommand.CommandText += "left join aca_pn s2 on s1.pn = s2.acapn where Date1 between to_date('"
        'oCommand.CommandText += sDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        ''oCommand.CommandText += "union all "
        ''oCommand.CommandText += "Select (case when s2.dacpn is null then s1.pn else s2.dacpn end) as PN1, to_char(date1,'yyyymmdd') , Quantity from ACA_Shipment_Return  s1 left join aca_pn s2 on s1.pn = s2.acapn where date1 between to_date('"
        ''oCommand.CommandText += sDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') 
        'oCommand.CommandText += " ) AB "
        'oCommand.CommandText += "LEFT JOIN ccc_file on PN1 = ccc_file.ccc01 and ccc02 = " & tYear & " and ccc03 = " & tMonth & " GROUP BY PN1,D1, ccc23"

        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
            Return
        End Try


        ' 計算出貨資料
        oCommand.CommandText = "select (case when s2.dacpn is null then s1.pn else s2.dacpn end) as c1, PRICE1, DOCCURR, SUM(QUANTITY) AS t1, Date1 from aca_shipment s1 left join aca_pn s2 on s1.pn = s2.acapn where date1 between to_date('"
        oCommand.CommandText += sDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') group by s2.dacpn,s1.pn,price1,Doccurr, Date1 order by Date1, c1"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Dim SQ As Decimal = oReader.Item("t1")  '出貨數量給值
                oCommand2.CommandText = "select * from aca_fifo_record where year1 = " & tYear & " and month1 = " & tMonth & " and pn = '" & oReader.Item("c1") & "' and qty > 0 order by lot"
                oReader2 = oCommand2.ExecuteReader()
                If oReader2.HasRows() Then
                    While oReader2.Read()
                        ' 找到, 開始處理
                        If SQ <= 0 Then
                            Continue While
                        End If
                        If SQ <= oReader2.Item("Qty") Then   ' 如果SQ 小於等於 Qty , 則讓 Qty -SQ  回寫到 Fifo
                            Dim XX1 As Decimal = oReader2.Item("Qty") - SQ
                            oCommand3.CommandText = "UPDATE aca_fifo_record SET Qty = " & XX1 & " WHERE year1 = " & tYear & " AND month1 =" & tMonth & " and PN ='" & oReader2.Item("PN") & "' and LOT = '" & oReader2.Item("LOT") & "'"
                            Try
                                oCommand3.ExecuteNonQuery()
                            Catch ex As Exception
                                MsgBox(ex.Message())
                            End Try
                            ' 回寫完要寫入 ShipmentCost 表
                            oCommand3.CommandText = "INSERT INTO aca_shipment_cost VALUES ('" & oReader.Item("C1") & "'," & oReader.Item("PRICE1") & ",'" & oReader.Item("DOCCURR") & "'," & SQ & ",to_date('" & oReader.Item("Date1") & "','yyyy/mm/dd') , '" & oReader2.Item("LOT") & "', " & oReader2.Item("CostPrice") & ")"
                            Try
                                oCommand3.ExecuteNonQuery()
                            Catch ex As Exception
                                MsgBox(ex.Message())
                            End Try
                            SQ = 0
                        Else
                            ' SQ > 0 則 FIFO 記錄設為0
                            oCommand3.CommandText = "UPDATE aca_fifo_record SET Qty = 0 WHERE year1 = " & tYear & " AND month1 =" & tMonth & " and PN ='" & oReader2.Item("PN") & "' and LOT = '" & oReader2.Item("LOT") & "'"
                            Try
                                oCommand3.ExecuteNonQuery()
                            Catch ex As Exception
                                MsgBox(ex.Message())
                            End Try
                            ' 回寫完要寫入 ShipmentCost 表
                            oCommand3.CommandText = "INSERT INTO aca_shipment_cost VALUES ('" & oReader.Item("C1") & "'," & oReader.Item("PRICE1") & ",'" & oReader.Item("DOCCURR") & "'," & oReader2.Item("Qty") & ",to_date('" & oReader.Item("Date1") & "','yyyy/mm/dd') , '" & oReader2.Item("LOT") & "', " & oReader2.Item("CostPrice") & ")"
                            Try
                                oCommand3.ExecuteNonQuery()
                            Catch ex As Exception
                                MsgBox(ex.Message())
                            End Try
                            SQ = SQ - oReader2.Item("Qty")
                        End If
                    End While
                    If SQ > 0 Then
                        'MsgBox("庫存有誤")
                        ' MsgBox(oReader.Item("C1") & "庫存為負")
                    End If
                Else
                    'MsgBox("找不到FIFO")
                    ' MsgBox(oReader.Item("C1") & "NO FIFO")

                End If
                oReader2.Close()
            End While
        Else
            MsgBox("無出貨資料")
        End If
        oReader.Close()


        '' 計算出貨資料
        'oCommand.CommandText = "select (case when s2.dacpn is null then s1.pn else s2.dacpn end) as c1, SUM(QUANTITY) AS t1, Date1, s1.price1,s1.doccurr from aca_shipment s1 left join aca_pn s2 on s1.pn = s2.acapn where date1 between to_date('"
        'oCommand.CommandText += sDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & eDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') group by s2.dacpn,s1.pn, Date1, s1.price1,s1.doccurr order by Date1, c1"
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        Dim SQ As Decimal = oReader.Item("t1")  '出貨數量給值
        '        oCommand2.CommandText = "select * from aca_fifo_record where year1 = " & tYear & " and month1 = " & tMonth & " and pn = '" & oReader.Item("c1") & "' and qty > 0 order by lot"
        '        oReader2 = oCommand2.ExecuteReader()
        '        If oReader2.HasRows() Then
        '            While oReader2.Read()
        '                ' 找到, 開始處理
        '                If SQ <= 0 Then
        '                    Continue While
        '                End If
        '                If SQ <= oReader2.Item("Qty") Then   ' 如果SQ 小於等於 Qty , 則讓 Qty -SQ  回寫到 Fifo
        '                    Dim XX1 As Decimal = oReader2.Item("Qty") - SQ
        '                    oCommand3.CommandText = "UPDATE aca_fifo_record SET Qty = " & XX1 & " WHERE year1 = " & tYear & " AND month1 =" & tMonth & " and PN ='" & oReader2.Item("PN") & "' and LOT = '" & oReader2.Item("LOT") & "'"
        '                    Try
        '                        oCommand3.ExecuteNonQuery()
        '                    Catch ex As Exception
        '                        MsgBox(ex.Message())
        '                    End Try
        '                    ' 回寫完要寫入 ShipmentCost 表
        '                    Dim PriceX As Decimal = 0
        '                    If IsDBNull(oReader.Item("price1")) Then
        '                        PriceX = 0
        '                    End If
        '                    If Not IsNumeric(oReader.Item("price1")) Then
        '                        PriceX = 0
        '                    Else
        '                        PriceX = oReader.Item("price1")
        '                    End If
        '                    oCommand3.CommandText = "INSERT INTO aca_shipment_cost VALUES ('" & oReader.Item("C1") & "'," & PriceX & ",'" & oReader.Item("doccurr") & "'," & SQ & ",to_date('" & oReader.Item("Date1") & "','yyyy/mm/dd') , '" & oReader2.Item("LOT") & "', " & oReader2.Item("CostPrice") & ")"
        '                    Try
        '                        oCommand3.ExecuteNonQuery()
        '                    Catch ex As Exception
        '                        MsgBox(ex.Message())
        '                    End Try
        '                    SQ = 0
        '                Else
        '                    ' SQ > 0 則 FIFO 記錄設為0
        '                    oCommand3.CommandText = "UPDATE aca_fifo_record SET Qty = 0 WHERE year1 = " & tYear & " AND month1 =" & tMonth & " and PN ='" & oReader2.Item("PN") & "' and LOT = '" & oReader2.Item("LOT") & "'"
        '                    Try
        '                        oCommand3.ExecuteNonQuery()
        '                    Catch ex As Exception
        '                        MsgBox(ex.Message())
        '                    End Try
        '                    ' 回寫完要寫入 ShipmentCost 表
        '                    Dim PriceX As Decimal = 0
        '                    If IsDBNull(oReader.Item("price1")) Then
        '                        PriceX = 0
        '                    End If
        '                    If Not IsNumeric(oReader.Item("price1")) Then
        '                        PriceX = 0
        '                    Else
        '                        PriceX = oReader.Item("price1")
        '                    End If
        '                    oCommand3.CommandText = "INSERT INTO aca_shipment_cost VALUES ('" & oReader.Item("C1") & "'," & PriceX & ",'" & oReader.Item("doccurr") & "'," & oReader2.Item("Qty") & ",to_date('" & oReader.Item("Date1") & "','yyyy/mm/dd') , '" & oReader2.Item("LOT") & "', " & oReader2.Item("CostPrice") & ")"
        '                    Try
        '                        oCommand3.ExecuteNonQuery()
        '                    Catch ex As Exception
        '                        MsgBox(ex.Message())
        '                    End Try
        '                    SQ = SQ - oReader2.Item("Qty")
        '                End If
        '            End While
        '            If SQ > 0 Then
        '                'MsgBox("庫存有誤")
        '                'MsgBox(oReader.Item("C1") & "庫存為負")
        '            End If
        '        Else
        '            'MsgBox("找不到FIFO")
        '            'MsgBox(oReader.Item("C1") & "NO FIFO")

        '        End If
        '        oReader2.Close()
        '    End While
        'Else
        '    MsgBox("無出貨資料")
        'End If
        'oReader.Close()

        MsgBox("Done")
    End Sub
End Class