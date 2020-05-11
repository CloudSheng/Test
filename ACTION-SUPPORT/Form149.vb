Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form149
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand3 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim LineS1 As Int16 = 0
    Dim tYear As Int16 = 0
    Dim tDate As Date
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form149_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        TextBox1.Text = Today.Year
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If TextBox1.Text.Length <> 4 Then
            MsgBox("ERROR")
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
        tYear = TextBox1.Text
        tDate = Convert.ToDateTime(tYear & "/01/01")
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "审计倒扎表" & tYear
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
        AdjustExcelFormat1()
        GetCCC("ccc12", "'101','104','106'")  '8
        LineZ += 1
        GetCCC("ccc22", "'101','104','106'") '9
        LineZ += 2
        GetCCC("ccc92", "'101','106','104'") '11
        LineZ += 1
        GetCCC("ccc42+ccc93+ccc72", "'101','104','106'", -1) '12
        LineZ += 2
        GetCDB("cdb05", "1") '14
        LineZ += 1
        GetCDB("cdb05", "2,3") '15
        LineZ += 2
        GetCCGANDCCC("ccg12", "ccc12", "'102'") ' 17
        LineZ += 1
        GetCCGANDCCC("ccg92", "ccc92", "'102'") ' 18
        LineZ += 1
        SGetCCC2("(ccc42 * -1)  +(ccc72 * -1)+(ccc93 * -1)-ccc224", "'102'", "ccc22a2", "'101','102','103','104','106'") '19
        LineZ += 2
        GetCCC("ccc12", "'103'")  '21
        LineZ += 1
        SGetCCC("ccc221+ccc224", "'103'", "ccc222", "'102','103'") '22
        LineZ += 1
        GetCCC("ccc92", "'103'")  '23
        LineZ += 1
        GetCCC("ccc42+ccc93+ccc72", "'103'", -1) '24
        LineZ += 12
        GetAAH()

        ' 處理 copy
        oRng = Ws.Range("E13", "E13")
        oRng.AutoFill(Destination:=Ws.Range("E13", "P13"), Type:=xlFillDefault)
        oRng = Ws.Range("E16", "E16")
        oRng.AutoFill(Destination:=Ws.Range("E16", "P16"), Type:=xlFillDefault)
        oRng = Ws.Range("E20", "E20")
        oRng.AutoFill(Destination:=Ws.Range("E20", "P20"), Type:=xlFillDefault)
        'oRng = Ws.Range("E15", "E15")
        'oRng.AutoFill(Destination:=Ws.Range("E15", "P15"), Type:=xlFillDefault)
        oRng = Ws.Range("E25", "E25")
        oRng.AutoFill(Destination:=Ws.Range("E25", "P25"), Type:=xlFillDefault)
        oRng = Ws.Range("E35", "E35")
        oRng.AutoFill(Destination:=Ws.Range("E35", "P35"), Type:=xlFillDefault)
        oRng = Ws.Range("E37", "E37")
        oRng.AutoFill(Destination:=Ws.Range("E37", "P37"), Type:=xlFillDefault)

        ' 自動大小
        oRng = Ws.Range("E3", "V3")
        oRng.EntireColumn.AutoFit()

    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10
        Ws.Columns.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        Ws.Name = "审计倒扎表"
        Ws.Rows.RowHeight = 18

        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 35
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.ColumnWidth = 22
        oRng = Ws.Range("D1", "D1")
        oRng.EntireColumn.ColumnWidth = 40
        oRng = Ws.Range("E7", "P7")
        oRng.NumberFormatLocal = "mmm-yy"
        oRng = Ws.Range("C4", "C4")
        oRng.EntireColumn.NumberFormat = "@"
        LineZ = 8
        Ws.Cells(2, 2) = "公司名称"
        Ws.Cells(2, 3) = "东莞艾可迅复合材料有限公司"
        Ws.Cells(3, 2) = "报表名称"
        Ws.Cells(3, 3) = "成本倒扎表"
        Ws.Cells(4, 2) = "年度截止日期"
        Ws.Cells(4, 3) = tYear & "12/31"
        Ws.Cells(5, 2) = "币别"
        Ws.Cells(5, 3) = "RMB"
        Ws.Cells(6, 2) = "借/(贷)"
        Ws.Cells(7, 2) = "项目内容"
        Ws.Cells(7, 3) = "计算说明"
        Ws.Cells(7, 4) = "数据来源"
        For i As Int16 = 1 To 12 Step 1
            Ws.Cells(7, 4 + i) = tDate.AddMonths(i - 1)
        Next
        Ws.Cells(7, 17) = "Total"
        Ws.Cells(7, 18) = "审计金额"
        Ws.Cells(7, 19) = "差异"
        Ws.Cells(7, 20) = "数据报表来源"
        Ws.Cells(7, 21) = "数据报表来源"
        Ws.Cells(7, 22) = "差异原因备注"
        Ws.Cells(8, 2) = "原材料年初余额"
        Ws.Cells(8, 3) = "1"
        Ws.Cells(8, 4) = "总账""材料.包装物.自制半成品""账户年初余额"
        Ws.Cells(8, 17) = "=SUM(E8:P8)"
        Ws.Cells(8, 19) = "=Q8-R8"
        Ws.Cells(9, 2) = "加：本年购入材料净额"
        Ws.Cells(9, 3) = "2"
        Ws.Cells(9, 4) = """材料""借方购入额扣退货折让金额"
        Ws.Cells(9, 17) = "=SUM(E9:P9)"
        Ws.Cells(9, 19) = "=Q9-R9"
        Ws.Cells(10, 2) = "加：其他增加额(包装物、自制半成品）"
        Ws.Cells(10, 3) = "3"
        Ws.Cells(10, 4) = """材料""借方其他发生额"
        Ws.Cells(10, 17) = "=SUM(E10:P10)"
        Ws.Cells(10, 19) = "=Q10-R10"
        Ws.Cells(11, 2) = "减：年末材料余额"
        Ws.Cells(11, 3) = "4"
        Ws.Cells(11, 4) = "总账""材料.包装物.自制半成品""账户年末余额"
        Ws.Cells(11, 17) = "=SUM(E11:P11)"
        Ws.Cells(11, 19) = "=Q11-R11"
        Ws.Cells(12, 2) = "减：其他发出额"
        Ws.Cells(12, 3) = "5"
        Ws.Cells(12, 4) = """材料""贷方其他发生额"
        Ws.Cells(12, 17) = "=SUM(E12:P12)"
        Ws.Cells(12, 19) = "=Q12-R12"
        Ws.Cells(13, 2) = "直接材料成本"
        Ws.Cells(13, 3) = "6=1+2+3-4-5"
        Ws.Cells(13, 4) = "生产成本明细账"
        Ws.Cells(13, 5) = "=E8+E9+E10-E11-E12"
        Ws.Cells(13, 17) = "=SUM(E13:P13)"
        Ws.Cells(13, 19) = "=Q13-R13"
        Ws.Cells(14, 2) = "直接人工成本"
        Ws.Cells(14, 3) = "7"
        Ws.Cells(14, 4) = "生产成本明细账"
        Ws.Cells(14, 17) = "=SUM(E14:P14)"
        Ws.Cells(14, 19) = "=Q14-R14"
        Ws.Cells(15, 2) = "制造费用"
        Ws.Cells(15, 3) = "8"
        Ws.Cells(15, 4) = "生产成本明细账"
        Ws.Cells(15, 17) = "=SUM(E15:P15)"
        Ws.Cells(15, 19) = "=Q15-R15"
        Ws.Cells(16, 2) = "产品生产成本"
        Ws.Cells(16, 3) = "9=6+7+8"
        Ws.Cells(16, 4) = """生产成本""借方发生额"
        Ws.Cells(16, 5) = "=E13+E14+E15"
        Ws.Cells(16, 17) = "=SUM(E16:P16)"
        Ws.Cells(16, 19) = "=Q16-R16"
        Ws.Cells(17, 2) = "加：在产品年初余额"
        Ws.Cells(17, 3) = "10"
        Ws.Cells(17, 4) = """生产成本""年初余额"
        Ws.Cells(17, 17) = "=SUM(E17:P17)"
        Ws.Cells(17, 19) = "=Q17-R17"
        Ws.Cells(18, 2) = "减：在产品年末余额"
        Ws.Cells(18, 3) = "11"
        Ws.Cells(18, 4) = """生产成本""年末余额"
        Ws.Cells(18, 17) = "=SUM(E18:P18)"
        Ws.Cells(18, 19) = "=Q18-R18"
        Ws.Cells(19, 2) = "减：其他发出额"
        Ws.Cells(19, 3) = "12"
        Ws.Cells(19, 4) = ""
        Ws.Cells(19, 17) = "=SUM(E19:P19)"
        Ws.Cells(19, 19) = "=Q19-R19"
        Ws.Cells(20, 2) = "产成品成本"
        Ws.Cells(20, 3) = "13=9+10-11-12"
        Ws.Cells(20, 4) = """生产成本""转入""产成品""借方金额"
        Ws.Cells(20, 5) = "=E16+E17-E18-E19"
        Ws.Cells(20, 17) = "=SUM(E20:P20)"
        Ws.Cells(20, 19) = "=Q20-R20"
        Ws.Cells(21, 2) = "加：产成品年初余额"
        Ws.Cells(21, 3) = "14"
        Ws.Cells(21, 4) = """产成品""年初余额"
        Ws.Cells(21, 17) = "=SUM(E21:P21)"
        Ws.Cells(21, 19) = "=Q21-R21"
        Ws.Cells(22, 2) = "加：其他增加额(外购等）"
        Ws.Cells(22, 3) = "15"
        Ws.Cells(22, 4) = """产成品""外购会计记录"
        Ws.Cells(22, 17) = "=SUM(E22:P22)"
        Ws.Cells(22, 19) = "=Q22-R22"
        Ws.Cells(23, 2) = "减：产成品年末余额"
        Ws.Cells(23, 3) = "16"
        Ws.Cells(23, 4) = """产成品""账户年末余额"
        Ws.Cells(23, 17) = "=SUM(E23:P23)"
        Ws.Cells(23, 19) = "=Q23-R23"
        Ws.Cells(24, 2) = "减：其他产成品成本"
        Ws.Cells(24, 3) = "17"
        Ws.Cells(24, 4) = ""
        Ws.Cells(24, 17) = "=SUM(E24:P24)"
        Ws.Cells(24, 19) = "=Q24-R24"
        Ws.Cells(25, 2) = "转入发出商品"
        Ws.Cells(25, 3) = "18=13+14+15-16-17"
        Ws.Cells(25, 4) = ""
        Ws.Cells(25, 5) = "=E20+E21+E22-E23-E24"
        Ws.Cells(25, 17) = "=SUM(E25:P25)"
        Ws.Cells(25, 19) = "=Q25-R25"
        Ws.Cells(26, 2) = "加：发出商品分摊成本更新差异"
        Ws.Cells(26, 3) = "19"
        Ws.Cells(26, 4) = """发出商品""材料差异"
        Ws.Cells(26, 17) = "=SUM(E26:P26)"
        Ws.Cells(26, 19) = "=Q26-R26"
        Ws.Cells(27, 2) = "加：发出商品年初余额"
        Ws.Cells(27, 3) = "20"
        Ws.Cells(27, 4) = """发出商品""年初余额"
        Ws.Cells(27, 17) = "=SUM(E27:P27)"
        Ws.Cells(27, 19) = "=Q27-R27"
        Ws.Cells(28, 2) = "加：出口产品分摊关税及不得退税额"
        Ws.Cells(28, 3) = "21"
        Ws.Cells(28, 4) = ""
        Ws.Cells(28, 17) = "=SUM(E28:P28)"
        Ws.Cells(28, 19) = "=Q28-R28"
        Ws.Cells(29, 2) = "加：产成品直接转销售成本"
        Ws.Cells(29, 3) = "22"
        Ws.Cells(29, 4) = ""
        Ws.Cells(29, 17) = "=SUM(E29:P29)"
        Ws.Cells(29, 19) = "=Q29-R29"
        Ws.Cells(30, 2) = "加：退货收回发出商品成本"
        Ws.Cells(30, 3) = "23"
        Ws.Cells(30, 4) = "用户退货会计记录"
        Ws.Cells(30, 17) = "=SUM(E30:P30)"
        Ws.Cells(30, 19) = "=Q30-R30"
        Ws.Cells(31, 2) = "加：发出商品其他转入"
        Ws.Cells(31, 3) = "24"
        Ws.Cells(31, 4) = """发出商品""其他转入"
        Ws.Cells(31, 17) = "=SUM(E31:P31)"
        Ws.Cells(31, 19) = "=Q31-R31"
        Ws.Cells(32, 2) = "减：发出商品年末余额"
        Ws.Cells(32, 3) = "25"
        Ws.Cells(32, 4) = """发出商品""年末余额"
        Ws.Cells(32, 17) = "=SUM(E32:P32)"
        Ws.Cells(32, 19) = "=Q32-R32"
        Ws.Cells(33, 2) = "加：发出商品盘盈及其他转入"
        Ws.Cells(33, 3) = "26"
        Ws.Cells(33, 4) = "存货（发出商品）盘盈会计记录"
        Ws.Cells(33, 17) = "=SUM(E33:P33)"
        Ws.Cells(33, 19) = "=Q33-R33"
        Ws.Cells(34, 2) = "加：外协电机标准成本差异"
        Ws.Cells(34, 3) = "27"
        Ws.Cells(34, 4) = ""
        Ws.Cells(34, 17) = "=SUM(E34:P34)"
        Ws.Cells(34, 19) = "=Q34-R34"
        Ws.Cells(35, 2) = "产品销售成本"
        Ws.Cells(35, 3) = "28=18+19+20+21+22+23+24-25+26+27"
        Ws.Cells(35, 5) = "=E25+E26+E27+E28+E29+E30+E31-E32+E33+E34"
        Ws.Cells(35, 17) = "=SUM(E35:P35)"
        Ws.Cells(35, 19) = "=S25+S26+S27+S28+S29+S30+S31+-S32+S33+S34"
        Ws.Cells(36, 4) = "主营业务成本账上数"
        Ws.Cells(36, 17) = "=SUM(E36:P36)"
        Ws.Cells(36, 18) = "=Q36"
        Ws.Cells(37, 4) = "差异"
        Ws.Cells(37, 5) = "=E35-E36"
        Ws.Cells(37, 17) = "=SUM(E37:P37)"
        Ws.Cells(37, 18) = "=+R36-R35"

        ' 劃線
        oRng = Ws.Range("B7", "V37")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
    End Sub
    Private Sub GetCCC(ByVal p1 As String, p2 As String)
        oCommand.CommandText = "select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( select "
        oCommand.CommandText += "(case when ccc03 =1 then " & p1 & " else 0 end) as t1,"
        oCommand.CommandText += "(case when ccc03 =2 then " & p1 & " else 0 end) as t2,"
        oCommand.CommandText += "(case when ccc03 =3 then " & p1 & " else 0 end) as t3,"
        oCommand.CommandText += "(case when ccc03 =4 then " & p1 & " else 0 end) as t4,"
        oCommand.CommandText += "(case when ccc03 =5 then " & p1 & " else 0 end) as t5,"
        oCommand.CommandText += "(case when ccc03 =6 then " & p1 & " else 0 end) as t6,"
        oCommand.CommandText += "(case when ccc03 =7 then " & p1 & " else 0 end) as t7,"
        oCommand.CommandText += "(case when ccc03 =8 then " & p1 & " else 0 end) as t8,"
        oCommand.CommandText += "(case when ccc03 =9 then " & p1 & " else 0 end) as t9,"
        oCommand.CommandText += "(case when ccc03 =10 then " & p1 & " else 0 end) as t10,"
        oCommand.CommandText += "(case when ccc03 =11 then " & p1 & " else 0 end) as t11,"
        oCommand.CommandText += "(case when ccc03 =12 then " & p1 & " else 0 end) as t12 "
        oCommand.CommandText += "from ccc_file left join ima_file on ccc01 = ima01 where ccc02 = " & tYear & " and ima06 in (" & p2 & ") )"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Integer = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, 5 + i) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub GetCDB(ByVal p1 As String, p2 As String)
        oCommand.CommandText = "select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( select "
        oCommand.CommandText += "(case when cdb02 =1 then " & p1 & " else 0 end) as t1,"
        oCommand.CommandText += "(case when cdb02 =2 then " & p1 & " else 0 end) as t2,"
        oCommand.CommandText += "(case when cdb02 =3 then " & p1 & " else 0 end) as t3,"
        oCommand.CommandText += "(case when cdb02 =4 then " & p1 & " else 0 end) as t4,"
        oCommand.CommandText += "(case when cdb02 =5 then " & p1 & " else 0 end) as t5,"
        oCommand.CommandText += "(case when cdb02 =6 then " & p1 & " else 0 end) as t6,"
        oCommand.CommandText += "(case when cdb02 =7 then " & p1 & " else 0 end) as t7,"
        oCommand.CommandText += "(case when cdb02 =8 then " & p1 & " else 0 end) as t8,"
        oCommand.CommandText += "(case when cdb02 =9 then " & p1 & " else 0 end) as t9,"
        oCommand.CommandText += "(case when cdb02 =10 then " & p1 & " else 0 end) as t10,"
        oCommand.CommandText += "(case when cdb02 =11 then " & p1 & " else 0 end) as t11,"
        oCommand.CommandText += "(case when cdb02 =12 then " & p1 & " else 0 end) as t12 "
        oCommand.CommandText += "from cdb_file where cdb01 = " & tYear & " and cdb04 in (" & p2 & " ) )"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Integer = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, 5 + i) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub GetCCC(ByVal p1 As String, p2 As String, p3 As Decimal)
        oCommand.CommandText = "select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( select "
        oCommand.CommandText += "(case when ccc03 =1 then " & p1 & " else 0 end) as t1,"
        oCommand.CommandText += "(case when ccc03 =2 then " & p1 & " else 0 end) as t2,"
        oCommand.CommandText += "(case when ccc03 =3 then " & p1 & " else 0 end) as t3,"
        oCommand.CommandText += "(case when ccc03 =4 then " & p1 & " else 0 end) as t4,"
        oCommand.CommandText += "(case when ccc03 =5 then " & p1 & " else 0 end) as t5,"
        oCommand.CommandText += "(case when ccc03 =6 then " & p1 & " else 0 end) as t6,"
        oCommand.CommandText += "(case when ccc03 =7 then " & p1 & " else 0 end) as t7,"
        oCommand.CommandText += "(case when ccc03 =8 then " & p1 & " else 0 end) as t8,"
        oCommand.CommandText += "(case when ccc03 =9 then " & p1 & " else 0 end) as t9,"
        oCommand.CommandText += "(case when ccc03 =10 then " & p1 & " else 0 end) as t10,"
        oCommand.CommandText += "(case when ccc03 =11 then " & p1 & " else 0 end) as t11,"
        oCommand.CommandText += "(case when ccc03 =12 then " & p1 & " else 0 end) as t12 "
        oCommand.CommandText += "from ccc_file left join ima_file on ccc01 = ima01 where ccc02 = " & tYear & " and ima06 in (" & p2 & ") )"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Integer = 0 To oReader.FieldCount - 1 Step 1
                    If p3 < 0 Then
                        Ws.Cells(LineZ, 5 + i) = oReader.Item(i) * Decimal.MinusOne
                    Else
                        Ws.Cells(LineZ, 5 + i) = oReader.Item(i)
                    End If

                Next
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub GetCCG(ByVal p1 As String)
        oCommand.CommandText = "select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( select "
        oCommand.CommandText += "(case when ccg03 =1 then " & p1 & " else 0 end) as t1,"
        oCommand.CommandText += "(case when ccg03 =2 then " & p1 & " else 0 end) as t2,"
        oCommand.CommandText += "(case when ccg03 =3 then " & p1 & " else 0 end) as t3,"
        oCommand.CommandText += "(case when ccg03 =4 then " & p1 & " else 0 end) as t4,"
        oCommand.CommandText += "(case when ccg03 =5 then " & p1 & " else 0 end) as t5,"
        oCommand.CommandText += "(case when ccg03 =6 then " & p1 & " else 0 end) as t6,"
        oCommand.CommandText += "(case when ccg03 =7 then " & p1 & " else 0 end) as t7,"
        oCommand.CommandText += "(case when ccg03 =8 then " & p1 & " else 0 end) as t8,"
        oCommand.CommandText += "(case when ccg03 =9 then " & p1 & " else 0 end) as t9,"
        oCommand.CommandText += "(case when ccg03 =10 then " & p1 & " else 0 end) as t10,"
        oCommand.CommandText += "(case when ccg03 =11 then " & p1 & " else 0 end) as t11,"
        oCommand.CommandText += "(case when ccg03 =12 then " & p1 & " else 0 end) as t12 "
        oCommand.CommandText += "from ccg_file where ccg02 = " & tYear & " )"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Integer = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, 5 + i) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub GetAAH()
        oCommand.CommandText = "select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( select "
        oCommand.CommandText += "(case when aah03 =1 then (aah04-aah05) else 0 end) as t1,"
        oCommand.CommandText += "(case when aah03 =2 then (aah04-aah05) else 0 end) as t2,"
        oCommand.CommandText += "(case when aah03 =3 then (aah04-aah05) else 0 end) as t3,"
        oCommand.CommandText += "(case when aah03 =4 then (aah04-aah05) else 0 end) as t4,"
        oCommand.CommandText += "(case when aah03 =5 then (aah04-aah05) else 0 end) as t5,"
        oCommand.CommandText += "(case when aah03 =6 then (aah04-aah05) else 0 end) as t6,"
        oCommand.CommandText += "(case when aah03 =7 then (aah04-aah05) else 0 end) as t7,"
        oCommand.CommandText += "(case when aah03 =8 then (aah04-aah05) else 0 end) as t8,"
        oCommand.CommandText += "(case when aah03 =9 then (aah04-aah05) else 0 end) as t9,"
        oCommand.CommandText += "(case when aah03 =10 then (aah04-aah05) else 0 end) as t10,"
        oCommand.CommandText += "(case when aah03 =11 then (aah04-aah05) else 0 end) as t11,"
        oCommand.CommandText += "(case when aah03 =12 then (aah04-aah05) else 0 end) as t12 "
        oCommand.CommandText += "from aah_file where aah02 = " & tYear & " and aah01 = '640101' )"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Integer = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, 5 + i) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub SGetCCC(ByVal p1 As String, p2 As String, p3 As String, p4 As String)
        oCommand.CommandText = "select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( select "
        oCommand.CommandText += "(case when ccc03 =1 then " & p1 & " else 0 end) as t1,"
        oCommand.CommandText += "(case when ccc03 =2 then " & p1 & " else 0 end) as t2,"
        oCommand.CommandText += "(case when ccc03 =3 then " & p1 & " else 0 end) as t3,"
        oCommand.CommandText += "(case when ccc03 =4 then " & p1 & " else 0 end) as t4,"
        oCommand.CommandText += "(case when ccc03 =5 then " & p1 & " else 0 end) as t5,"
        oCommand.CommandText += "(case when ccc03 =6 then " & p1 & " else 0 end) as t6,"
        oCommand.CommandText += "(case when ccc03 =7 then " & p1 & " else 0 end) as t7,"
        oCommand.CommandText += "(case when ccc03 =8 then " & p1 & " else 0 end) as t8,"
        oCommand.CommandText += "(case when ccc03 =9 then " & p1 & " else 0 end) as t9,"
        oCommand.CommandText += "(case when ccc03 =10 then " & p1 & " else 0 end) as t10,"
        oCommand.CommandText += "(case when ccc03 =11 then " & p1 & " else 0 end) as t11,"
        oCommand.CommandText += "(case when ccc03 =12 then " & p1 & " else 0 end) as t12 "
        oCommand.CommandText += "from ccc_file left join ima_file on ccc01 = ima01 where ccc02 = " & tYear & " and ima06 in (" & p2 & ") "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select "
        oCommand.CommandText += "(case when ccc03 =1 then " & p3 & " else 0 end) as t1,"
        oCommand.CommandText += "(case when ccc03 =2 then " & p3 & " else 0 end) as t2,"
        oCommand.CommandText += "(case when ccc03 =3 then " & p3 & " else 0 end) as t3,"
        oCommand.CommandText += "(case when ccc03 =4 then " & p3 & " else 0 end) as t4,"
        oCommand.CommandText += "(case when ccc03 =5 then " & p3 & " else 0 end) as t5,"
        oCommand.CommandText += "(case when ccc03 =6 then " & p3 & " else 0 end) as t6,"
        oCommand.CommandText += "(case when ccc03 =7 then " & p3 & " else 0 end) as t7,"
        oCommand.CommandText += "(case when ccc03 =8 then " & p3 & " else 0 end) as t8,"
        oCommand.CommandText += "(case when ccc03 =9 then " & p3 & " else 0 end) as t9,"
        oCommand.CommandText += "(case when ccc03 =10 then " & p3 & " else 0 end) as t10,"
        oCommand.CommandText += "(case when ccc03 =11 then " & p3 & " else 0 end) as t11,"
        oCommand.CommandText += "(case when ccc03 =12 then " & p3 & " else 0 end) as t12 "
        oCommand.CommandText += "from ccc_file left join ima_file on ccc01 = ima01 where ccc02 = " & tYear & " and ima06 in (" & p4 & ") )"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Integer = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, 5 + i) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub GetCCGANDCCC(ByVal p1 As String, ByVal p2 As String, ByVal p3 As String)
        oCommand.CommandText = "select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( select "
        oCommand.CommandText += "(case when ccg03 =1 then " & p1 & " else 0 end) as t1,"
        oCommand.CommandText += "(case when ccg03 =2 then " & p1 & " else 0 end) as t2,"
        oCommand.CommandText += "(case when ccg03 =3 then " & p1 & " else 0 end) as t3,"
        oCommand.CommandText += "(case when ccg03 =4 then " & p1 & " else 0 end) as t4,"
        oCommand.CommandText += "(case when ccg03 =5 then " & p1 & " else 0 end) as t5,"
        oCommand.CommandText += "(case when ccg03 =6 then " & p1 & " else 0 end) as t6,"
        oCommand.CommandText += "(case when ccg03 =7 then " & p1 & " else 0 end) as t7,"
        oCommand.CommandText += "(case when ccg03 =8 then " & p1 & " else 0 end) as t8,"
        oCommand.CommandText += "(case when ccg03 =9 then " & p1 & " else 0 end) as t9,"
        oCommand.CommandText += "(case when ccg03 =10 then " & p1 & " else 0 end) as t10,"
        oCommand.CommandText += "(case when ccg03 =11 then " & p1 & " else 0 end) as t11,"
        oCommand.CommandText += "(case when ccg03 =12 then " & p1 & " else 0 end) as t12 "
        oCommand.CommandText += "from ccg_file where ccg02 = " & tYear
        oCommand.CommandText += " union all select "
        oCommand.CommandText += "(case when ccc03 =1 then " & p2 & " else 0 end) as t1,"
        oCommand.CommandText += "(case when ccc03 =2 then " & p2 & " else 0 end) as t2,"
        oCommand.CommandText += "(case when ccc03 =3 then " & p2 & " else 0 end) as t3,"
        oCommand.CommandText += "(case when ccc03 =4 then " & p2 & " else 0 end) as t4,"
        oCommand.CommandText += "(case when ccc03 =5 then " & p2 & " else 0 end) as t5,"
        oCommand.CommandText += "(case when ccc03 =6 then " & p2 & " else 0 end) as t6,"
        oCommand.CommandText += "(case when ccc03 =7 then " & p2 & " else 0 end) as t7,"
        oCommand.CommandText += "(case when ccc03 =8 then " & p2 & " else 0 end) as t8,"
        oCommand.CommandText += "(case when ccc03 =9 then " & p2 & " else 0 end) as t9,"
        oCommand.CommandText += "(case when ccc03 =10 then " & p2 & " else 0 end) as t10,"
        oCommand.CommandText += "(case when ccc03 =11 then " & p2 & " else 0 end) as t11,"
        oCommand.CommandText += "(case when ccc03 =12 then " & p2 & " else 0 end) as t12 "
        oCommand.CommandText += "from ccc_file left join ima_file on ccc01 = ima01 where ccc02 = " & tYear & " and ima06 in (" & p3 & ") )"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Integer = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, 5 + i) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub SGetCCC2(ByVal p1 As String, p2 As String, p3 As String, p4 As String)
        oCommand.CommandText = "select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( select "
        oCommand.CommandText += "(case when ccc03 =1 then " & p1 & " else 0 end) as t1,"
        oCommand.CommandText += "(case when ccc03 =2 then " & p1 & " else 0 end) as t2,"
        oCommand.CommandText += "(case when ccc03 =3 then " & p1 & " else 0 end) as t3,"
        oCommand.CommandText += "(case when ccc03 =4 then " & p1 & " else 0 end) as t4,"
        oCommand.CommandText += "(case when ccc03 =5 then " & p1 & " else 0 end) as t5,"
        oCommand.CommandText += "(case when ccc03 =6 then " & p1 & " else 0 end) as t6,"
        oCommand.CommandText += "(case when ccc03 =7 then " & p1 & " else 0 end) as t7,"
        oCommand.CommandText += "(case when ccc03 =8 then " & p1 & " else 0 end) as t8,"
        oCommand.CommandText += "(case when ccc03 =9 then " & p1 & " else 0 end) as t9,"
        oCommand.CommandText += "(case when ccc03 =10 then " & p1 & " else 0 end) as t10,"
        oCommand.CommandText += "(case when ccc03 =11 then " & p1 & " else 0 end) as t11,"
        oCommand.CommandText += "(case when ccc03 =12 then " & p1 & " else 0 end) as t12 "
        oCommand.CommandText += "from ccc_file left join ima_file on ccc01 = ima01 where ccc02 = " & tYear & " and ima06 in (" & p2 & ") "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select "
        oCommand.CommandText += "(case when ccc03 =1 then " & p3 & " else 0 end) as t1,"
        oCommand.CommandText += "(case when ccc03 =2 then " & p3 & " else 0 end) as t2,"
        oCommand.CommandText += "(case when ccc03 =3 then " & p3 & " else 0 end) as t3,"
        oCommand.CommandText += "(case when ccc03 =4 then " & p3 & " else 0 end) as t4,"
        oCommand.CommandText += "(case when ccc03 =5 then " & p3 & " else 0 end) as t5,"
        oCommand.CommandText += "(case when ccc03 =6 then " & p3 & " else 0 end) as t6,"
        oCommand.CommandText += "(case when ccc03 =7 then " & p3 & " else 0 end) as t7,"
        oCommand.CommandText += "(case when ccc03 =8 then " & p3 & " else 0 end) as t8,"
        oCommand.CommandText += "(case when ccc03 =9 then " & p3 & " else 0 end) as t9,"
        oCommand.CommandText += "(case when ccc03 =10 then " & p3 & " else 0 end) as t10,"
        oCommand.CommandText += "(case when ccc03 =11 then " & p3 & " else 0 end) as t11,"
        oCommand.CommandText += "(case when ccc03 =12 then " & p3 & " else 0 end) as t12 "
        oCommand.CommandText += "from ccc_file left join ima_file on ccc01 = ima01 where ccc02 = " & tYear & " and ima06 in (" & p4 & ") "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select (case when ccg03 = 1 then cch32d * -1 else 0 end) as t1,"
        oCommand.CommandText += "(case when ccg03 = 2 then cch32d * -1 else 0 end) as t2,"
        oCommand.CommandText += "(case when ccg03 = 3 then cch32d * -1 else 0 end) as t3,"
        oCommand.CommandText += "(case when ccg03 = 4 then cch32d * -1 else 0 end) as t4,"
        oCommand.CommandText += "(case when ccg03 = 5 then cch32d * -1 else 0 end) as t5,"
        oCommand.CommandText += "(case when ccg03 = 6 then cch32d * -1 else 0 end) as t6,"
        oCommand.CommandText += "(case when ccg03 = 7 then cch32d * -1 else 0 end) as t7,"
        oCommand.CommandText += "(case when ccg03 = 8 then cch32d * -1 else 0 end) as t8,"
        oCommand.CommandText += "(case when ccg03 = 9 then cch32d * -1 else 0 end) as t9,"
        oCommand.CommandText += "(case when ccg03 = 10 then cch32d * -1 else 0 end) as t10,"
        oCommand.CommandText += "(case when ccg03 = 11 then cch32d * -1 else 0 end) as t11,"
        oCommand.CommandText += "(case when ccg03 = 12 then cch32d * -1 else 0 end) as t12 from ccg_file left join ima_file on ccg04 = ima01 left join cch_file on ccg01 = cch01 and cch02 =ccg02 and cch03 = ccg03 "
        oCommand.CommandText += "where ccg02 = " & tYear & "  and ima08 = 'S' and cch05  = 'S' )"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Integer = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, 5 + i) = oReader.Item(i)
                Next
            End While
        End If
        oReader.Close()
    End Sub
End Class