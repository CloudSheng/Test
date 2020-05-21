Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants

Public Class Form301
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oSQLReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oSQLS1 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim TYear As String = String.Empty
    Dim TMonth As String = String.Empty
    Dim LYear As String = String.Empty
    Dim LMonth As String = String.Empty
    Dim Time1 As Date
    Dim Time2 As Date
    Dim D5 As Date
    Dim C1 As String = String.Empty
    Dim C2 As String = String.Empty
    Dim LineZ As Integer = 0
    Dim mAdapter1 As New SqlClient.SqlDataAdapter
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub SumRow()
        Ws.Cells(LineZ, 39) = "=SUM(H" & LineZ & ":AK" & LineZ & ")"
        Ws.Cells(LineZ, 40) = "=SUM(E" & LineZ & ":AK" & LineZ & ")"
        Ws.Cells(LineZ, 41) = "=SUM(N" & LineZ & ":AK" & LineZ & ")"
        Ws.Cells(LineZ, 42) = "=SUM(AE" & LineZ & ":AH" & LineZ & ")" '+SUM(AA" & LineZ & ":AD" & LineZ & ")"
    End Sub
    Private Sub Form301_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If Now.Month < 10 Then
            TextBox1.Text = Now.Year & "0" & Now.Month
        Else
            TextBox1.Text = Now.Year & Now.Month
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oSQLS1.Connection = oConnection
                oSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        TYear = Strings.Left(TextBox1.Text, 4)
        TMonth = Strings.Right(TextBox1.Text, 2)
        Time1 = Convert.ToDateTime(TYear & "/" & TMonth & "/01")
        Time2 = Time1.AddMonths(1).AddDays(-1)
        If TMonth > 1 Then
            LYear = TYear
            LMonth = TMonth - 1
        ElseIf TMonth = 1 Then
            LYear = TYear - 1
            LMonth = 12
        End If
        BackgroundWorker1.RunWorkerAsync()
        'ExportToExcel()
        'SaveExcel()
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "资材仓月度出入仓数量与金额"
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
                Module1.KillExcelProcess(OldExcel)
                MsgBox("Finished")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub ExportToExcel()


        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        Ws.Name = "ALL"
        AdjustExcelFormat()
        oCommand.CommandText = "select imk02,imd02,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8 from ("
        oCommand.CommandText += " select imk02,sum(imk09) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8 from imk_file where imk05 = " & LYear & " and imk06 = " & LMonth & " and imk02 in ('D146101','D146102','D146103','D146104','D146106','D146107','D146108') group by imk02"
        oCommand.CommandText += " union all select tlf902,0,sum(tlf10*tlf12),0,0,0,0,0,0 from tlf_file where tlf06 between to_date('" & Time1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        oCommand.CommandText += " and to_date('" & Time2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tlf907 = 1 and tlf902 in ('D146101','D146102','D146103','D146104','D146106','D146107','D146108') group by tlf902"
        oCommand.CommandText += " union all select tlf902,0,0,sum(tlf10*tlf12),0,0,0,0,0 from tlf_file where tlf06 between to_date('" & Time1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        oCommand.CommandText += " and to_date('" & Time2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tlf907 = -1 and tlf902 in ('D146101','D146102','D146103','D146104','D146106','D146107','D146108') group by tlf902"
        oCommand.CommandText += " union all select imk02,0,0,0,sum(imk09),0,0,0,0 from imk_file where imk05 = " & TYear & " and imk06 = " & TMonth & " and imk02 in ('D146101','D146102','D146103','D146104','D146106','D146107','D146108') group by imk02"
        oCommand.CommandText += " union all select imk02,0,0,0,0,sum(imk09*(stb07+stb08+stb09)),0,0,0 from imk_file,stb_file where imk01=stb01 and imk05=stb02 and imk06=stb03 and imk05 = " & LYear & " and imk06 = " & LMonth & " and imk02 in ('D146101','D146102','D146103','D146104','D146106','D146107','D146108') group by imk02"
        oCommand.CommandText += " union all select tlf902,0,0,0,0,0,sum(tlf10*tlf12*(stb07+stb08+stb09)),0,0 from tlf_file,stb_file where tlf01=stb01 and stb02=" & TYear & " and stb03=" & TMonth & " and tlf06 between to_date('" & Time1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        oCommand.CommandText += " and to_date('" & Time2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tlf907 = 1 and tlf902 in ('D146101','D146102','D146103','D146104','D146106','D146107','D146108') group by tlf902"
        oCommand.CommandText += " union all select tlf902,0,0,0,0,0,0,sum(tlf10*tlf12*(stb07+stb08+stb09)),0 from tlf_file,stb_file where tlf01=stb01 and stb02=" & TYear & " and stb03=" & TMonth & " and tlf06 between to_date('" & Time1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        oCommand.CommandText += " and to_date('" & Time2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tlf907 = -1 and tlf902 in ('D146101','D146102','D146103','D146104','D146106','D146107','D146108') group by tlf902"
        oCommand.CommandText += " union all select imk02,0,0,0,0,0,0,0,sum(imk09*(stb07+stb08+stb09)) from imk_file,stb_file where imk01=stb01 and imk05=stb02 and imk06=stb03 and imk05 = " & TYear & " and imk06 = " & TMonth & " and imk02 in ('D146101','D146102','D146103','D146104','D146106','D146107','D146108') group by imk02"
        oCommand.CommandText += " ),imd_file where imk02=imd01 group by imk02,imd02"


        oSQLReader = oCommand.ExecuteReader()
        If oSQLReader.HasRows() Then
            While oSQLReader.Read()
                Ws.Cells(LineZ, 1) = oSQLReader.Item("imk02")
                Ws.Cells(LineZ, 2) = oSQLReader.Item("imd02")
                Ws.Cells(LineZ, 3) = oSQLReader.Item("t1")
                Ws.Cells(LineZ, 4) = oSQLReader.Item("t2")
                Ws.Cells(LineZ, 5) = oSQLReader.Item("t3")
                Ws.Cells(LineZ, 6) = oSQLReader.Item("t4")
                Ws.Cells(LineZ, 7) = oSQLReader.Item("t5") / 6.4
                Ws.Cells(LineZ, 8) = oSQLReader.Item("t6") / 6.4
                Ws.Cells(LineZ, 9) = oSQLReader.Item("t7") / 6.4
                Ws.Cells(LineZ, 10) = oSQLReader.Item("t8") / 6.4
                LineZ += 1
            End While
        End If
        oSQLReader.Close()

        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        Ws.Name = "FG"
        AdjustExcelFormat1()
        oCommand.CommandText = "select imk01,ima02,ima021,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8 from ("
        oCommand.CommandText += " select imk01,sum(imk09) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8 from imk_file where imk05 = " & LYear & " and imk06 = " & LMonth & " and imk02='D146103' group by imk01"
        oCommand.CommandText += " union all select tlf01,0,sum(tlf10*tlf12),0,0,0,0,0,0 from tlf_file where tlf06 between to_date('" & Time1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        oCommand.CommandText += " and to_date('" & Time2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tlf907 = 1 and tlf902='D146103' group by tlf01"
        oCommand.CommandText += " union all select tlf01,0,0,sum(tlf10*tlf12),0,0,0,0,0 from tlf_file where tlf06 between to_date('" & Time1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        oCommand.CommandText += " and to_date('" & Time2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tlf907 = -1 and tlf902='D146103' group by tlf01"
        oCommand.CommandText += " union all select imk01,0,0,0,sum(imk09),0,0,0,0 from imk_file where imk05 = " & TYear & " and imk06 = " & TMonth & " and imk02='D146103' group by imk01"
        oCommand.CommandText += " union all select imk01,0,0,0,0,sum(imk09*(stb07+stb08+stb09)),0,0,0 from imk_file,stb_file where imk01=stb01 and imk05=stb02 and imk06=stb03 and imk05 = " & LYear & " and imk06 = " & LMonth & " and imk02='D146103' group by imk01"
        oCommand.CommandText += " union all select tlf01,0,0,0,0,0,sum(tlf10*tlf12*(stb07+stb08+stb09)),0,0 from tlf_file,stb_file where tlf01=stb01 and stb02=" & TYear & " and stb03=" & TMonth & " and tlf06 between to_date('" & Time1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        oCommand.CommandText += " and to_date('" & Time2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tlf907 = 1 and tlf902='D146103' group by tlf01"
        oCommand.CommandText += " union all select tlf01,0,0,0,0,0,0,sum(tlf10*tlf12*(stb07+stb08+stb09)),0 from tlf_file,stb_file where tlf01=stb01 and stb02=" & TYear & " and stb03=" & TMonth & " and tlf06 between to_date('" & Time1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        oCommand.CommandText += " and to_date('" & Time2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tlf907 = -1 and tlf902='D146103' group by tlf01"
        oCommand.CommandText += " union all select imk01,0,0,0,0,0,0,0,sum(imk09*(stb07+stb08+stb09)) from imk_file,stb_file where imk01=stb01 and imk05=stb02 and imk06=stb03 and imk05 = " & TYear & " and imk06 = " & TMonth & " and imk02='D146103' group by imk01"
        oCommand.CommandText += " ),ima_file where imk01=ima01 group by imk01,ima02,ima021"
        oSQLReader = oCommand.ExecuteReader()
        If oSQLReader.HasRows() Then
            While oSQLReader.Read()
                Ws.Cells(LineZ, 1) = oSQLReader.Item("imk01")
                Ws.Cells(LineZ, 2) = oSQLReader.Item("ima02")
                Ws.Cells(LineZ, 3) = oSQLReader.Item("ima021")
                Ws.Cells(LineZ, 4) = oSQLReader.Item("t1")
                Ws.Cells(LineZ, 5) = oSQLReader.Item("t2")
                Ws.Cells(LineZ, 6) = oSQLReader.Item("t3")
                Ws.Cells(LineZ, 7) = oSQLReader.Item("t4")
                Ws.Cells(LineZ, 8) = oSQLReader.Item("t5") / 6.4
                Ws.Cells(LineZ, 9) = oSQLReader.Item("t6") / 6.4
                Ws.Cells(LineZ, 10) = oSQLReader.Item("t7") / 6.4
                Ws.Cells(LineZ, 11) = oSQLReader.Item("t8") / 6.4
                Ws.Cells(1, 4) = "=SUM(D3:D" & LineZ & ")"
                Ws.Cells(1, 5) = "=SUM(E3:E" & LineZ & ")"
                Ws.Cells(1, 6) = "=SUM(F3:F" & LineZ & ")"
                Ws.Cells(1, 7) = "=SUM(G3:G" & LineZ & ")"
                Ws.Cells(1, 8) = "=SUM(H3:H" & LineZ & ")"
                Ws.Cells(1, 9) = "=SUM(I3:I" & LineZ & ")"
                Ws.Cells(1, 10) = "=SUM(J3:J" & LineZ & ")"
                Ws.Cells(1, 11) = "=SUM(K3:K" & LineZ & ")"

                LineZ += 1
            End While
        End If
        oSQLReader.Close()

        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
            Return
        End Try


        'WIP   

        Ws = xWorkBook.Sheets(3)
        Ws.Name = "WIP"
        Ws.Activate()
        AdjustExcelFormat3()
        'If oConnection.State <> ConnectionState.Open Then
        '    oConnection.Open()
        '    oSQLS1.Connection = oConnection
        '    oSQLS1.CommandType = CommandType.Text
        'End If
        oSQLS1.CommandText = "DROP TABLE MES_TEMP1"
        Try
            oSQLS1.ExecuteNonQuery()
        Catch ex As Exception
            'MsgBox(ex.Message())
            'Return
        End Try

        oSQLS1.CommandText = "CREATE TABLE mes_temp1 (cf01 nvarchar2(500),station nvarchar2(4),t1 DEC(10,0))"
        Try
            oSQLS1.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
            Return
        End Try

        Dim mConnectionBuilder As New SqlClient.SqlConnectionStringBuilder
        Dim mConnection As New SqlClient.SqlConnection
        Dim mSQLS1 As New SqlClient.SqlCommand
        mConnectionBuilder.DataSource = "192.168.10.254"
        mConnectionBuilder.InitialCatalog = "ERPSUPPORT"
        'mConnectionBuilder.InitialCatalog = "IQMES3"
        mConnectionBuilder.IntegratedSecurity = False
        mConnectionBuilder.UserID = "sa"
        mConnectionBuilder.Password = "p@$$w0rd"
        mConnection.ConnectionString = mConnectionBuilder.ConnectionString

        If mConnection.State <> ConnectionState.Open Then
            mConnection.Open()
            mSQLS1.Connection = mConnection
            mSQLS1.CommandType = CommandType.Text
        End If

        mSQLS1.CommandText = "select sERPPN as cf01,sStation as station,sum(sqty) as t1 from WIPSaveData where sYear = " & TYear & " and sMonth=" & TMonth & " group by sERPPN,sStation"
        Dim mSQLReader As SqlClient.SqlDataReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                oSQLS1.CommandText = "insert into mes_temp1 (cf01,station,t1) VALUES ('"
                oSQLS1.CommandText += mSQLReader.Item("cf01").ToString & "','" & mSQLReader.Item("station").ToString & "'," & mSQLReader.Item("t1") & ")"
                Try
                    oSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                    Return
                End Try
            End While
        End If
        mSQLReader.Close()

        mSQLS1.CommandText = "select smodel,modelname,value,sum(sqty) as Count1,AA from ( SELECT smodel,modelname,value,sqty," _
      & "(case when sStation  in ('0055','0080','0100','0110','0111') then '1裁纱' " _
    & "when sStation  in ('0112','0113') then '2备料' " _
    & "when sStation  in ('0130','0140','0150','0151','0160','0170','0172','0175','0177','0180','0193') then '3预型' " _
    & "when sStation  in ('0165','0173','0174','0175','0190','0195','0200','0210','0215','0220','0223','0225','0230','0231','0240','0250','0255','0260','0280','0300','0315','0316','0320','0321','0325','0326','0330','0331','0333','0390','0395') then '4成型' " _
    & "when sStation  in ('0335','0340','0350','0360','0370','0380','0385','0390','0495','0500','0510','0520','0530') then '5CNC' " _
    & "when sStation  in ('0400','0435','0478','0479','0480','0485','0490','0491','0492','0493','0605','0610','0611','0620','0623','0627') then '6胶合' " _
    & "when sStation  in ('0625','0629','0630','0633','0635','0640','0645') then '9拋光' " _
    & "when sStation  in ('0642','0649','0650','0652','0657','0658','0659','0660','0665','0666','0667','0668','0669','0670','0673','0674') then 'A待包裝' " _
    & "when sStation  in ('0675','0680','0690') then 'B已包裝' " _
    & "when sStation  in ('0642','0649','0650','0652','0657','0658','0659','0660','0665','0666','0667','0668','0669','0670','0673','0674','0675','0680','0690') then 'C包裝' " _
    & "when sStation  in ('0405') then 'D底漆防漆' " _
    & "when sStation  in ('0410','0415','0416','0417','0440','0475') then 'E底漆研磨' " _
    & "when sStation  in ('0418','0420','0430','0441','0445','0450','0455') then 'F底漆涂装' " _
    & "when sStation  in ('0460','0461','0465','0540','0541','0545','0567','0570','0575','0583','0584') then 'G涂装研磨' " _
    & "when sStation  in ('0470','0550','0560','0563','0580','0585','0587','0590','0591','0592','0595') then 'H面漆涂装' " _
    & "when sStation  in ('BLCK') then 'I隔離品' else '~XX' end) as AA " _
    & "FROM ERPSUPPORT.dbo.WIPSaveData left join IQMES3.dbo.model on smodel = model.model left join IQMES3.dbo.model_paravalue on model_paravalue.parameter = 'ERP PN' and sModel = model_paravalue.model "
        mSQLS1.CommandText += "WHERE sYear = " & TYear & " and sMonth = " & TMonth & ") AS B  GROUP BY smodel,modelname,value,AA order by smodel "
        mSQLReader = mSQLS1.ExecuteReader()
        Dim CheckFormat As String = String.Empty
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                If String.IsNullOrEmpty(CheckFormat) Then
                    SumRow1()
                    CheckFormat = mSQLReader.Item("smodel")
                    Ws.Cells(LineZ, 1) = mSQLReader.Item("Value")
                    Ws.Cells(LineZ, 2) = CheckFormat
                    Ws.Cells(LineZ, 3) = mSQLReader.Item("modelname")
                End If
                If Not CheckFormat = mSQLReader.Item("smodel") Then
                    SumRow1()
                    LineZ += 1
                    Ws.Cells(LineZ, 1) = mSQLReader.Item("Value")
                    Ws.Cells(LineZ, 2) = mSQLReader.Item("smodel")
                    Ws.Cells(LineZ, 3) = mSQLReader.Item("modelname")
                    CheckFormat = mSQLReader.Item("smodel")
                End If
                Dim CheckPosition As String = Strings.Left(mSQLReader.Item("AA").ToString, 1)
                Select Case CheckPosition
                    Case "1"
                        Ws.Cells(LineZ, 4) = mSQLReader.Item("Count1")
                    Case "2"
                        Ws.Cells(LineZ, 5) = mSQLReader.Item("Count1")
                    Case "3"
                        Ws.Cells(LineZ, 6) = mSQLReader.Item("Count1")
                    Case "4"
                        Ws.Cells(LineZ, 7) = mSQLReader.Item("Count1")
                    Case "5"
                        Ws.Cells(LineZ, 8) = mSQLReader.Item("Count1")
                    Case "6"
                        Ws.Cells(LineZ, 9) = mSQLReader.Item("Count1")
                        'Case "7"
                        '   Ws.Cells(LineZ, 10) = mSQLReader.Item("Count1")
                        'Case "8"
                        '   Ws.Cells(LineZ, 11) = mSQLReader.Item("Count1")
                    Case "9"
                        Ws.Cells(LineZ, 12) = mSQLReader.Item("Count1")
                    Case "C"
                        Ws.Cells(LineZ, 13) = mSQLReader.Item("Count1")
                    Case "A"
                        Ws.Cells(LineZ, 14) = mSQLReader.Item("Count1")
                    Case "B"
                        Ws.Cells(LineZ, 15) = mSQLReader.Item("Count1")
                    Case "D"
                        Ws.Cells(LineZ, 16) = mSQLReader.Item("Count1")
                    Case "E"
                        Ws.Cells(LineZ, 17) = mSQLReader.Item("Count1")
                    Case "F"
                        Ws.Cells(LineZ, 18) = mSQLReader.Item("Count1")
                    Case "G"
                        Ws.Cells(LineZ, 19) = mSQLReader.Item("Count1")
                    Case "H"
                        Ws.Cells(LineZ, 20) = mSQLReader.Item("Count1")
                    Case "I"
                        Ws.Cells(LineZ, 21) = mSQLReader.Item("Count1")
                End Select
            End While
            SumRow1()
        End If
        mSQLReader.Close()
        mConnection.Close()

        '金額
        oSQLS1.CommandText = "SELECT AA,nvl(SUM(T1A),0) AS T1A,nvl(sum(t2a),0) as t2a FROM ( SELECT " _
              & "(case when station in ('0055','0080','0100','0110','0111')  then '1裁纱' " _
           & "when station in ('0112','0113') then '2备料' " _
           & "when station in ('0130','0140','0150','0151','0160','0170','0172','0175','0177','0180','0193') then '3预型' " _
           & "when station in ('0165','0173','0174','0175','0190','0195','0200','0210','0215','0220','0223','0225','0230','0231','0240','0250','0255','0260','0280','0300','0315','0316','0320','0321','0325','0326','0330','0331','0333','0390','0395') then '4成型' " _
           & "when station in ('0335','0340','0350','0360','0370','0380','0385','0390','0495','0500','0510','0520','0530') then '5CNC' " _
           & "when station in ('0400','0435','0478','0479','0480','0485','0490','0491','0492','0493','0605','0610','0611','0620','0623','0627') then '6胶合' " _
           & "when Station  in ('0625','0629','0630','0633','0635','0640','0645') then '9拋光' " _
            & "when Station  in ('0642','0649','0650','0652','0657','0658','0659','0660','0665','0666','0667','0668','0669','0670','0673','0674') then 'A待包裝' " _
            & "when Station  in ('0675','0680','0690') then 'B已包裝' " _
            & "when Station  in ('0642','0649','0650','0652','0657','0658','0659','0660','0665','0666','0667','0668','0669','0670','0673','0674','0675','0680','0690') then 'C包裝' " _
            & "when station  in ('0405') then 'D底漆防漆' " _
            & "when Station  in ('0410','0415','0417','0440','0475') then 'E底漆研磨' " _
            & "when Station  in ('0418','0420','0430','0441','0445','0450','0455') then 'F底漆涂装' " _
            & "when Station  in ('0460','0461','0465','0540','0541','0545','0567','0570','0575','0583','0584') then 'G涂装研磨' " _
            & "when Station  in ('0470','0550','0560','0563','0580','0585','0587','0590','0591','0592','0595') then 'H面漆涂装' " _
            & "when Station  in ('BLCK') then 'I隔離品' else '~XX' end) as AA ,round(sum(t1*(stb07 + stb08 + stb09)),4) as t1a,sum(t1) as t2a FROM MES_TEMP1 left join stb_file on cf01 = stb01 and stb02 = "
        oSQLS1.CommandText += TYear & " and stb03 = " & TMonth & " group by station ) GROUP BY AA"

        oSQLReader = oSQLS1.ExecuteReader()
        If oSQLReader.HasRows() Then
            While oSQLReader.Read()
                Dim CheckPosition As String = Strings.Left(oSQLReader.Item("AA").ToString, 1)
                Select Case CheckPosition
                    Case "1"
                        Ws.Cells(1, 4) = oSQLReader.Item("t1a") / 6.4
                        Ws.Cells(2, 4) = oSQLReader.Item("t2a")
                    Case "2"
                        Ws.Cells(1, 5) = oSQLReader.Item("t1a") / 6.4
                        Ws.Cells(2, 5) = oSQLReader.Item("t2a")
                    Case "3"
                        Ws.Cells(1, 6) = oSQLReader.Item("t1a") / 6.4
                        Ws.Cells(2, 6) = oSQLReader.Item("t2a")
                    Case "4"
                        Ws.Cells(1, 7) = oSQLReader.Item("t1a") / 6.4
                        Ws.Cells(2, 7) = oSQLReader.Item("t2a")
                    Case "5"
                        Ws.Cells(1, 8) = oSQLReader.Item("t1a") / 6.4
                        Ws.Cells(2, 8) = oSQLReader.Item("t2a")
                    Case "6"
                        Ws.Cells(1, 9) = oSQLReader.Item("t1a") / 6.4
                        Ws.Cells(2, 9) = oSQLReader.Item("t2a")
                        'Case "7"
                        '   Ws.Cells(1, 10) = oSQLReader.Item("t1a") / 6.4
                        '   Ws.Cells(2, 10) = oSQLReader.Item("t2a")
                        'Case "8"
                        '   Ws.Cells(1, 11) = oSQLReader.Item("t1a") / 6.4
                        '  Ws.Cells(2, 11) = oSQLReader.Item("t2a")
                    Case "9"
                        Ws.Cells(1, 12) = oSQLReader.Item("t1a") / 6.4
                        Ws.Cells(2, 12) = oSQLReader.Item("t2a")
                    Case "C"
                        Ws.Cells(1, 13) = oSQLReader.Item("t1a") / 6.4
                        Ws.Cells(2, 13) = oSQLReader.Item("t2a")
                    Case "A"
                        Ws.Cells(1, 14) = oSQLReader.Item("t1a") / 6.4
                        Ws.Cells(2, 14) = oSQLReader.Item("t2a")
                    Case "B"
                        Ws.Cells(1, 15) = oSQLReader.Item("t1a") / 6.4
                        Ws.Cells(2, 15) = oSQLReader.Item("t2a")
                    Case "D"
                        Ws.Cells(1, 16) = oSQLReader.Item("t1a") / 6.4
                        Ws.Cells(2, 16) = oSQLReader.Item("t2a")
                    Case "E"
                        Ws.Cells(1, 17) = oSQLReader.Item("t1a") / 6.4
                        Ws.Cells(2, 17) = oSQLReader.Item("t2a")
                    Case "F"
                        Ws.Cells(1, 18) = oSQLReader.Item("t1a") / 6.4
                        Ws.Cells(2, 18) = oSQLReader.Item("t2a")
                    Case "G"
                        Ws.Cells(1, 19) = oSQLReader.Item("t1a") / 6.4
                        Ws.Cells(2, 19) = oSQLReader.Item("t2a")
                    Case "H"
                        Ws.Cells(1, 20) = oSQLReader.Item("t1a") / 6.4
                        Ws.Cells(2, 20) = oSQLReader.Item("t2a")
                    Case "I"
                        Ws.Cells(1, 21) = oSQLReader.Item("t1a") / 6.4
                        Ws.Cells(2, 21) = oSQLReader.Item("t2a")
                End Select
            End While
            Ws.Cells(1, 13) = "=SUM(N1:O1)"
            Ws.Cells(2, 13) = "=SUM(N2:O2)"
            Ws.Cells(1, 10) = "=SUM(P1:R1)"
            Ws.Cells(2, 10) = "=SUM(P2:R2)"
            Ws.Cells(1, 11) = "=SUM(S1:T1)"
            Ws.Cells(2, 11) = "=SUM(S2:T2)"
        End If
        oSQLReader.Close()

        oSQLS1.CommandText = "DROP TABLE MES_TEMP1"
        Try
            oSQLS1.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
            Return
        End Try
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        'Ws.Name = TMonth
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 20
        Ws.Cells(1, 1) = "仓库编码"
        Ws.Cells(1, 2) = "仓别名称"
        Ws.Cells(1, 3) = "期初总数量"
        Ws.Cells(1, 4) = "本期入仓数量"
        Ws.Cells(1, 5) = "本期出仓数量"
        Ws.Cells(1, 6) = "期末总数量"
        Ws.Cells(1, 7) = "期初总金额"
        Ws.Cells(1, 8) = "本期入仓金额"
        Ws.Cells(1, 9) = "本期出仓金额"
        Ws.Cells(1, 10) = "期末总金额"
        oRng = Ws.Range("C1", "J1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00_ "
        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 20
        Ws.Cells(1, 3) = "Total:"
        Ws.Cells(2, 1) = "ERP料号"
        Ws.Cells(2, 2) = "产品名称"
        Ws.Cells(2, 3) = "产品规格"
        Ws.Cells(2, 4) = "期初总数量"
        Ws.Cells(2, 5) = "本期入仓数量"
        Ws.Cells(2, 6) = "本期出仓数量"
        Ws.Cells(2, 7) = "期末总数量"
        Ws.Cells(2, 8) = "期初总金额"
        Ws.Cells(2, 9) = "本期入仓金额"
        Ws.Cells(2, 10) = "本期出仓金额"
        Ws.Cells(2, 11) = "期末总金额"
        oRng = Ws.Range("D1", "K1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00_ "
        LineZ = 3
    End Sub



    Private Sub AdjustExcelFormat3()
        xExcel.ActiveWindow.Zoom = 75
        oRng = Ws.Range("A1", "C1")
        oRng.EntireColumn.ColumnWidth = 40
        oRng = Ws.Range("D1", "T1")
        oRng.EntireColumn.ColumnWidth = 15
        oRng = Ws.Range("A2", "T2")
        oRng.Interior.Color = System.Drawing.Color.LightBlue
        Ws.Cells(1, 3) = "价值合计"
        Ws.Cells(2, 3) = "数量合计"
        Ws.Cells(3, 1) = "ERP 料号"
        Ws.Cells(3, 2) = "产品名称"
        Ws.Cells(3, 3) = "产品名称"
        Ws.Cells(3, 4) = "标签"
        Ws.Cells(3, 5) = "裁紗"
        Ws.Cells(3, 6) = "预型"
        Ws.Cells(3, 7) = "成型"
        Ws.Cells(3, 8) = "CNC"
        Ws.Cells(3, 9) = "胶合"
        Ws.Cells(3, 10) = "补土"
        Ws.Cells(3, 11) = "涂装"
        Ws.Cells(3, 12) = "抛光"
        Ws.Cells(3, 13) = "包装合计"
        Ws.Cells(3, 14) = "待包装"
        Ws.Cells(3, 15) = "已包装"
        Ws.Cells(3, 16) = "底漆防漆"
        Ws.Cells(3, 17) = "底漆研磨"
        Ws.Cells(3, 18) = "底漆涂装"
        Ws.Cells(3, 19) = "涂装研磨"
        Ws.Cells(3, 20) = "面漆涂装"
        Ws.Cells(3, 21) = "Block(隔离品)"
        LineZ = 4
    End Sub
    Private Sub SumRow1()
        Ws.Cells(LineZ, 13) = "=SUM(N" & LineZ & ":O" & LineZ & ")"
        Ws.Cells(LineZ, 10) = "=SUM(P" & LineZ & ":R" & LineZ & ")"
        Ws.Cells(LineZ, 11) = "=SUM(S" & LineZ & ":T" & LineZ & ")"
    End Sub
End Class