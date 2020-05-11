Public Class Form71
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0

    Private Sub Form71_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'OpenFileDialog1.Title = "OPEN EXCEL"
        'Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        'If selectOK = System.Windows.Forms.DialogResult.OK Then
        '    Dim ExcelPath As String = OpenFileDialog1.FileName
        '    Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
        '    Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
        '    Excelconn.Open()
        '    Dim ExcelString = "SELECT oeb04 FROM [sheet1$]"
        '    Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)
        '    Dim DS As Data.DataSet = New DataSet()
        '    Try
        '        ExcelAdapater.Fill(DS, "table1")
        '    Catch ex As Exception
        '        MsgBox(ex.Message())
        '    End Try

        '    For i As Integer = 0 To DS.Tables("table1").Rows.Count - 1 Step 1
        '        oCommand.CommandText = "select t1 from ( "
        '        oCommand.CommandText += "Select Round((oeb13 * oea24 / 6.4), 4) as t1 from oeb_file,oea_file where oea01 = oeb01 and oea01 not like 'D2309%' and oea99 is not null and oeb13 <> 0 and oeaconf  = 'Y' "
        '        oCommand.CommandText += "and oea02 < to_date('2017/03/18','yyyy/mm/dd') and oeb04 = '" & DS.Tables("table1").Rows(i).Item(0).ToString & "'  order by oea02 desc ) where rownum <= 1"
        '        Dim P1 As Decimal = oCommand.ExecuteScalar()
        '        oCommander2.CommandText = "INSERT INTO price_temp (oeb04,price) VALUES ('" & DS.Tables("table1").Rows(i).Item(0).ToString & "'," & P1 & ")"
        '        oCommander2.ExecuteNonQuery()
        '    Next
        '    MsgBox("DONE")
        'End If
        'oCommand.CommandText = "select * from aaa_temp"
        'oReader = oCommand.ExecuteReader
        'If oReader.HasRows() Then
        'While oReader.Read()
        'oCommander2.CommandText = "update bmb_file set bmb03 = '526000020014' where bmb01 = '" & oReader.Item("bmb01") & "' and bmb02 = " & oReader.Item("bmb02") & " and bmb03 = '526000020013'"
        'Try
        'oCommander2.ExecuteNonQuery()
        'Catch ex As Exception
        'MsgBox(ex.Message())
        'End Try
        'End While
        'End If
        'oReader.Close()
        'MsgBox("OK")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        xExcel.DisplayAlerts = False
        mSQLS1.CommandText = "select SUBSTRING(photo_filename, 2, 100)   from z_ms_equipment where equipment_id = 'A0400061'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Dim AA As String = "http://192.168.10.254/IQMES"
                Dim AB As String = AA & mSQLReader.Item(0)
                Ws.Shapes.AddPicture(AB, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 30, 40, 50, 60)

            End While
        End If
        mSQLReader.Close()
        SaveExcel()
    End Sub


    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Test"
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
                'Module1.KillExcelProcess(OldExcel)
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
        Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
        Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
        Dim oCommander3 As New Oracle.ManagedDataAccess.Client.OracleCommand
        Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
        Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommander2.Connection = oConnection
                oCommander2.CommandType = CommandType.Text
                oCommander3.Connection = oConnection
                oCommander3.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        oCommand.CommandText = "select (case when s2.acapn is null then s1.pn else dacpn end) as pn, s1.pn as pn1 ,price1, doccurr, date1 from aca_shipment s1 left join aca_pn s2 on s1.pn = s2.acapn  where price1 is null order by pn"
        oReader = oCommand.ExecuteReader()
        Dim AA = 0
        If oReader.HasRows() Then
            While oReader.Read()
                AA += 1
                Me.Label1.Text = AA
                Me.Label1.Refresh()
                oCommander2.CommandText = "select * from aca_price_list where pn = '" & oReader.Item("pn") & "' and vdate > to_date('" & oReader.Item("date1") & "','yyyy/mm/dd') order by vdate"
                oReader2 = oCommander2.ExecuteReader()
                If oReader2.HasRows() Then
                    oReader2.Read()
                    oCommander3.CommandText = "UPDATE aca_shipment SET price1 = " & oReader2.Item("price1") & ", doccurr ='" & oReader2.Item("currency") & "' WHERE pn = '" & oReader.Item("pn1") & "' AND Date1 = to_date('" & oReader.Item("date1") & "','yyyy/mm/dd') "
                    Try
                        oCommander3.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                Else

                End If
                oReader2.Close()
            End While
        End If
        oReader.Close()
        MsgBox("Done")
    End Sub
End Class