Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form148
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

    Private Sub Form148_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        SaveFileDialog1.FileName = "内部倒扎表" & tYear
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
        GetCCC("ccc12", "'101'")  '5
        LineZ += 1
        GetCCC("ccc12", "'106'") '6
        LineZ += 1
        GetCCC("ccc12", "'104'") '7
        LineZ += 2
        GetCCC("ccc221", "'101'") '9
        LineZ += 1
        GetCCC("ccc221", "'106'") '10
        LineZ += 1
        GetCCC("ccc221", "'104'") '11
        LineZ += 2
        GetCCC("ccc224", "'101','104','106'") '13
        LineZ += 1
        GetCCC("ccc222", "'101','104','106'") '14
        LineZ += 2
        GetCCC("ccc92", "'101'") '16
        LineZ += 1
        GetCCC("ccc92", "'106'") '17
        LineZ += 1
        GetCCC("ccc92", "'104'") '18
        LineZ += 2
        GetCCC("ccc42", "'101','104','106'", -1) '20
        LineZ += 1
        GetCCC("ccc93", "'101','104','106'", -1) '21
        LineZ += 1
        GetCCC("ccc72", "'101','104','106'", -1) '22
        LineZ += 3
        GetCDB("cdb05", "1") '25
        LineZ += 1
        GetCDB("cdb05", "2, 3") '26
        LineZ += 3
        GetCCG("ccg12a") ' 29
        LineZ += 1
        GetCCG("ccg12b") '30
        LineZ += 1
        GetCCG("ccg12c+ccg12e+ccg12d") '31
        LineZ += 1
        GetCCC("ccc12", "'102'") '32
        LineZ += 3
        GetCCG("ccg92a") '35
        LineZ += 1
        GetCCG("ccg92b") '36
        LineZ += 1
        GetCCG("ccg92c+ccg92e+ccg92d") '37
        LineZ += 1
        GetCCC("ccc92", "'102'") '38
        LineZ += 1
        SGetCCC2("ccc22a2", "'101','102','103','104','106'") '39
        LineZ += 2
        GetCCC("(ccc42 * -1)-ccc224", "102") '41
        LineZ += 1
        GetCCC("ccc72", "'102'", -1) '42
        LineZ += 1
        GetCCC("ccc93", "'102'", -1) '43
        LineZ += 2
        GetCCC("ccc12", "'103'") '45
        LineZ += 1
        SGetCCC() '46
        LineZ += 1
        GetCCC("ccc224", "'103'")
        LineZ += 1
        GetCCC("ccc92", "'103'")
        LineZ += 1
        GetCCC("ccc42", "'103'", -1)
        LineZ += 1
        GetCCC("ccc72", "'103'", -1)
        LineZ += 1
        GetCCC("ccc93", "'103'", -1)
        LineZ += 2
        GetCCC("ccc62", "'101', '102','103','104','106'", -1)
        LineZ += 2
        GetAAH()
        ' 處理 copy
        oRng = Ws.Range("E4", "E4")
        oRng.AutoFill(Destination:=Ws.Range("E4", "P4"), Type:=xlFillDefault)
        oRng = Ws.Range("E8", "E8")
        oRng.AutoFill(Destination:=Ws.Range("E8", "P8"), Type:=xlFillDefault)
        oRng = Ws.Range("E12", "E12")
        oRng.AutoFill(Destination:=Ws.Range("E12", "P12"), Type:=xlFillDefault)
        oRng = Ws.Range("E15", "E15")
        oRng.AutoFill(Destination:=Ws.Range("E15", "P15"), Type:=xlFillDefault)
        oRng = Ws.Range("E19", "E19")
        oRng.AutoFill(Destination:=Ws.Range("E19", "P19"), Type:=xlFillDefault)
        oRng = Ws.Range("E23", "E24")
        oRng.AutoFill(Destination:=Ws.Range("E23", "P24"), Type:=xlFillDefault)
        oRng = Ws.Range("E27", "E28")
        oRng.AutoFill(Destination:=Ws.Range("E27", "P28"), Type:=xlFillDefault)
        oRng = Ws.Range("E33", "E34")
        oRng.AutoFill(Destination:=Ws.Range("E33", "P34"), Type:=xlFillDefault)
        oRng = Ws.Range("E40", "E40")
        oRng.AutoFill(Destination:=Ws.Range("E40", "P40"), Type:=xlFillDefault)
        oRng = Ws.Range("E44", "E44")
        oRng.AutoFill(Destination:=Ws.Range("E44", "P44"), Type:=xlFillDefault)
        oRng = Ws.Range("E52", "E52")
        oRng.AutoFill(Destination:=Ws.Range("E52", "P52"), Type:=xlFillDefault)
        oRng = Ws.Range("E54", "E54")
        oRng.AutoFill(Destination:=Ws.Range("E54", "P54"), Type:=xlFillDefault)
        oRng = Ws.Range("E56", "E56")
        oRng.AutoFill(Destination:=Ws.Range("E56", "P56"), Type:=xlFillDefault)
        ' 自動大小
        oRng = Ws.Range("E3", "Q3")
        oRng.EntireColumn.AutoFit()

    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.Font.Name = "Arial"
        Ws.Columns.Font.Size = 10
        Ws.Columns.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        Ws.Name = "倒扎表明细"
        Ws.Rows.RowHeight = 18

        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.ColumnWidth = 35
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.ColumnWidth = 22
        oRng = Ws.Range("D1", "D1")
        oRng.EntireColumn.ColumnWidth = 40
        oRng = Ws.Range("E3", "P3")
        oRng.NumberFormatLocal = "mmm-yy"
        oRng = Ws.Range("C4", "C4")
        oRng.EntireColumn.NumberFormat = "@"
        LineZ = 5
        Ws.Cells(2, 2) = "公司名称：东莞艾可迅复合材料有限公司"
        Ws.Cells(2, 3) = "Currency:RMB"
        Ws.Cells(3, 2) = "项目内容"
        Ws.Cells(3, 3) = "计算说明"
        Ws.Cells(3, 4) = "数据来源"
        For i As Int16 = 1 To 12 Step 1
            Ws.Cells(3, 4 + i) = tDate.AddMonths(i - 1)
        Next
        Ws.Cells(3, 17) = "Total"
        Ws.Cells(4, 2) = "加：材料年初余额"
        Ws.Cells(4, 3) = "1=2+3+4"
        Ws.Cells(4, 4) = "总账""原材料.包装物.低耗品""账户年初余额"
        Ws.Cells(4, 5) = "=E5+E6+E7"
        Ws.Cells(4, 17) = "=SUM(E4:P4)"
        Ws.Cells(5, 2) = "原材料年初金额"
        Ws.Cells(5, 3) = "2"
        Ws.Cells(5, 4) = "总账""原材料""账户年初余额"
        Ws.Cells(5, 17) = "=SUM(E5:P5)"
        Ws.Cells(6, 2) = "包材年初金额"
        Ws.Cells(6, 3) = "3"
        Ws.Cells(6, 4) = "总账""包装物""账户年初余额"
        Ws.Cells(6, 17) = "=SUM(E6:P6)"
        Ws.Cells(7, 2) = "低值易耗品年初金额"
        Ws.Cells(7, 3) = "4"
        Ws.Cells(7, 4) = "总账""低耗品""账户年初余额"
        Ws.Cells(7, 17) = "=SUM(E7:P7)"
        Ws.Cells(8, 2) = "加：本年购入材料净额"
        Ws.Cells(8, 3) = "5=6+7+8"
        Ws.Cells(8, 4) = """原材料.包装物.低耗品""借方购入额扣退货折让金额"
        Ws.Cells(8, 5) = "=E9+E10+E11"
        Ws.Cells(8, 17) = "=SUM(E8:P8)"
        Ws.Cells(9, 2) = "原材料本年金额"
        Ws.Cells(9, 3) = "6"
        Ws.Cells(9, 4) = """原材料""借方购入额扣退货折让金额"
        Ws.Cells(9, 17) = "=SUM(E9:P9)"
        Ws.Cells(10, 2) = "包材本年金额"
        Ws.Cells(10, 3) = "7"
        Ws.Cells(10, 4) = """包材""借方购入额扣退货折让金额"
        Ws.Cells(10, 17) = "=SUM(E10:P10)"
        Ws.Cells(11, 2) = "低值易耗品本年金额"
        Ws.Cells(11, 3) = "8"
        Ws.Cells(11, 4) = """低耗品""借方购入额扣退货折让金额"
        Ws.Cells(11, 17) = "=SUM(E11:P11)"
        Ws.Cells(12, 2) = "加：其他增加额(非外购/委外入库）"
        Ws.Cells(12, 3) = "9=10"
        Ws.Cells(12, 4) = """原材料.包材.低耗品""借方其他发生额"
        Ws.Cells(12, 5) = "=E13"
        Ws.Cells(12, 17) = "=SUM(E12:P12)"
        Ws.Cells(13, 2) = "其他增加额(非外购/委外入库）"
        Ws.Cells(13, 3) = "10"
        Ws.Cells(13, 4) = "杂项入库金额"
        Ws.Cells(13, 17) = "=SUM(E13:P13)"
        Ws.Cells(14, 2) = "加：物料委外入库（不含产品）"
        Ws.Cells(14, 3) = "11"
        Ws.Cells(14, 4) = "物料委外入库（不含产品）"
        Ws.Cells(14, 17) = "=SUM(E14:P14)"
        Ws.Cells(15, 2) = "减：年末材料余额"
        Ws.Cells(15, 3) = "12=13+14+15"
        Ws.Cells(15, 4) = "总账""材料.包装物.低耗""账户年末余额"
        Ws.Cells(15, 5) = "=E16+E17+E18"
        Ws.Cells(15, 17) = "=SUM(E15:P15)"
        Ws.Cells(16, 2) = "原材料年末金额"
        Ws.Cells(16, 3) = "13"
        Ws.Cells(16, 4) = "总账""原材料""账户年末余额"
        Ws.Cells(16, 17) = "=SUM(E16:P16)"
        Ws.Cells(17, 2) = "包材年末金额"
        Ws.Cells(17, 3) = "14"
        Ws.Cells(17, 4) = "总账""包装物""账户年末余额"
        Ws.Cells(17, 17) = "=SUM(E17:P17)"
        Ws.Cells(18, 2) = "低值易耗品年末金额"
        Ws.Cells(18, 3) = "15"
        Ws.Cells(18, 4) = "总账""低耗品""账户年末余额"
        Ws.Cells(18, 17) = "=SUM(E18:P18)"
        Ws.Cells(19, 2) = "减：其他发出额(非外购/委外出库）"
        Ws.Cells(19, 3) = "16=17+18+19"
        Ws.Cells(19, 4) = """原材料.包材.低耗品""借方其他发生额"
        Ws.Cells(19, 5) = "=E20+E21+E22"
        Ws.Cells(19, 17) = "=SUM(E19:P19)"
        Ws.Cells(20, 2) = "其他发出额(非外购/委外出库）"
        Ws.Cells(20, 3) = "17"
        Ws.Cells(20, 4) = "杂项领用成本"
        Ws.Cells(20, 17) = "=SUM(E20:P20)"
        Ws.Cells(21, 2) = "其他发出额(非外购/委外出库）"
        Ws.Cells(21, 3) = "18"
        Ws.Cells(21, 4) = "结存调整金额"
        Ws.Cells(21, 17) = "=SUM(E21:P21)"
        Ws.Cells(22, 2) = "其他发出额(非外购/委外出库）"
        Ws.Cells(22, 3) = "19"
        Ws.Cells(22, 4) = "盘盈亏金额"
        Ws.Cells(22, 17) = "=SUM(E22:P22)"
        Ws.Cells(23, 2) = "加：本期投入生产成本"
        Ws.Cells(23, 3) = "20=21+22+23"
        Ws.Cells(23, 4) = "本期投入生产成本：直接材料.直接人工.制造费用"
        Ws.Cells(23, 5) = "=E24+E25+E26"
        Ws.Cells(23, 17) = "=SUM(E23:P23)"
        Ws.Cells(24, 2) = "直接材料成本"
        Ws.Cells(24, 3) = "21=1+5+9+11-12-16"
        Ws.Cells(24, 4) = "生产成本明细账"
        Ws.Cells(24, 5) = "=E4+E8+E12+E14-E15-E19"
        Ws.Cells(24, 17) = "=SUM(E24:P24)"
        Ws.Cells(25, 2) = "直接人工成本"
        Ws.Cells(25, 3) = "22"
        Ws.Cells(25, 4) = "生产成本明细账"
        Ws.Cells(25, 17) = "=SUM(E25:P25)"
        Ws.Cells(26, 2) = "制造费用"
        Ws.Cells(26, 3) = "23"
        Ws.Cells(26, 4) = "生产成本明细账"
        Ws.Cells(26, 17) = "=SUM(E26:P26)"
        Ws.Cells(27, 2) = "加：在产品年初余额"
        Ws.Cells(27, 3) = "24=25+29"
        Ws.Cells(27, 4) = """在制品.半成品""年初余额"
        Ws.Cells(27, 5) = "=E28+E32"
        Ws.Cells(27, 17) = "=SUM(E27:P27)"
        Ws.Cells(28, 2) = "加：在制品年初金额"
        Ws.Cells(28, 3) = "25=26+27+28"
        Ws.Cells(28, 4) = """生产成本""年初余额"
        Ws.Cells(28, 5) = "=E29+E30+E31"
        Ws.Cells(28, 17) = "=SUM(E28:P28)"
        Ws.Cells(29, 2) = "在制品年初金额"
        Ws.Cells(29, 3) = "26"
        Ws.Cells(29, 4) = """生产成本-直接材料""年初余额"
        Ws.Cells(29, 17) = "=SUM(E29:P29)"
        Ws.Cells(30, 2) = "在制品年初金额"
        Ws.Cells(30, 3) = "27"
        Ws.Cells(30, 4) = """生产成本-直接人工""年初余额"
        Ws.Cells(30, 17) = "=SUM(E30:P30)"
        Ws.Cells(31, 2) = "在制品年初金额"
        Ws.Cells(31, 3) = "28"
        Ws.Cells(31, 4) = """生产成本-制造费用及加工费用""年初余额"
        Ws.Cells(31, 17) = "=SUM(E31:P31)"
        Ws.Cells(32, 2) = "半成品年初金额"
        Ws.Cells(32, 3) = "29"
        Ws.Cells(32, 4) = """半成品""年初余额"
        Ws.Cells(32, 17) = "=SUM(E32:P32)"
        Ws.Cells(33, 2) = "减：在产品年末余额"
        Ws.Cells(33, 3) = "30=31+35"
        Ws.Cells(33, 4) = """生产成本.半成品""年末余额"
        Ws.Cells(33, 5) = "=E34+E38"
        Ws.Cells(33, 17) = "=SUM(E33:P33)"
        Ws.Cells(34, 2) = "减：在制品年末金额"
        Ws.Cells(34, 3) = "31=32+33+34"
        Ws.Cells(34, 4) = """生产成本""年末余额"
        Ws.Cells(34, 5) = "=E35+E36+E37"
        Ws.Cells(34, 17) = "=SUM(E34:P34)"
        Ws.Cells(35, 2) = "在制品年末金额"
        Ws.Cells(35, 3) = "32"
        Ws.Cells(35, 4) = """生产成本-直接材料""年末余额"
        Ws.Cells(35, 17) = "=SUM(E35:P35)"
        Ws.Cells(36, 2) = "在制品年末金额"
        Ws.Cells(36, 3) = "33"
        Ws.Cells(36, 4) = """生产成本-直接人工""年末余额"
        Ws.Cells(36, 17) = "=SUM(E36:P36)"
        Ws.Cells(37, 2) = "在制品年末金额"
        Ws.Cells(37, 3) = "34"
        Ws.Cells(37, 4) = """生产成本-制造费用及加工费用""年末余额"
        Ws.Cells(37, 17) = "=SUM(E37:P37)"
        Ws.Cells(38, 2) = "半成品年末金额"
        Ws.Cells(38, 3) = "35"
        Ws.Cells(38, 4) = """半成品""年末余额"
        Ws.Cells(38, 17) = "=SUM(E38:P38)"
        Ws.Cells(39, 2) = "减：委外物料入库金额（扣减外发原材料入库金额）"
        Ws.Cells(39, 3) = "36"
        Ws.Cells(39, 4) = "委外物料入库金额（扣减外发原材料入库金额）"
        Ws.Cells(39, 17) = "=SUM(E39:P39)"
        Ws.Cells(40, 2) = "减：半成品其他减少额(非工单领用）"
        Ws.Cells(40, 3) = "37=38+39+40"
        Ws.Cells(40, 4) = "半成品其他减少额(非工单领用）"
        Ws.Cells(40, 5) = "=E41+E42+E43"
        Ws.Cells(40, 17) = "=SUM(E40:P40)"
        Ws.Cells(41, 2) = "半成品其他减少额(非工单领用）"
        Ws.Cells(41, 3) = "38"
        Ws.Cells(41, 4) = "杂项领用成本-杂项入库金额"
        Ws.Cells(41, 17) = "=SUM(E41:P41)"
        Ws.Cells(42, 2) = "半成品其他减少额(非工单领用）"
        Ws.Cells(42, 3) = "39"
        Ws.Cells(42, 4) = "盘盈亏金额"
        Ws.Cells(42, 17) = "=SUM(E42:P42)"
        Ws.Cells(43, 2) = "半成品其他减少额(非工单领用）"
        Ws.Cells(43, 3) = "40"
        Ws.Cells(43, 4) = "结存调整金额"
        Ws.Cells(43, 17) = "=SUM(E43:P43)"
        Ws.Cells(44, 2) = "加：本期产成品成本"
        Ws.Cells(44, 3) = "41=20+24-30-36-37"
        Ws.Cells(44, 4) = """生产成本""转入""产成品""借方金额"
        Ws.Cells(44, 5) = "=E23+E27-E33-E39-E40"
        Ws.Cells(44, 17) = "=SUM(E44:P44)"
        Ws.Cells(45, 2) = "加：库存商品年初余额"
        Ws.Cells(45, 3) = "42"
        Ws.Cells(45, 4) = """库存商品""年初余额"
        Ws.Cells(45, 17) = "=SUM(E45:P45)"
        Ws.Cells(46, 2) = "加：其他增加额(外购）"
        Ws.Cells(46, 3) = "43"
        Ws.Cells(46, 4) = """库存商品""外购会计记录（含委外）"
        Ws.Cells(46, 17) = "=SUM(E46:P46)"
        Ws.Cells(47, 2) = "加：其他增加额(杂收）"
        Ws.Cells(47, 3) = "44"
        Ws.Cells(47, 4) = """库存商品""杂收会计记录"
        Ws.Cells(47, 17) = "=SUM(E47:P47)"
        Ws.Cells(48, 2) = "减：库存商品年末余额"
        Ws.Cells(48, 3) = "45"
        Ws.Cells(48, 4) = """库存商品""账户年末余额"
        Ws.Cells(48, 17) = "=SUM(E48:P48)"
        Ws.Cells(49, 2) = "减：其他增加额(杂发）"
        Ws.Cells(49, 3) = "46"
        Ws.Cells(49, 4) = """库存商品""杂发会计记录"
        Ws.Cells(49, 17) = "=SUM(E49:P49)"
        Ws.Cells(50, 2) = "减：其他增加额(盘盈亏）"
        Ws.Cells(50, 3) = "47"
        Ws.Cells(50, 4) = """库存商品""盘盈亏会计记录"
        Ws.Cells(50, 17) = "=SUM(E50:P50)"
        Ws.Cells(51, 2) = "减：其他增加额(结存调整金额）"
        Ws.Cells(51, 3) = "48"
        Ws.Cells(51, 4) = """库存商品""结存调整金额会计记录"
        Ws.Cells(51, 17) = "=SUM(E51:P51)"
        Ws.Cells(52, 2) = "产品销售成本"
        Ws.Cells(52, 3) = "49=41+42+43+44-45-46-47-48"
        Ws.Cells(52, 4) = "产品销售成本"
        Ws.Cells(52, 5) = "=E44+E45+E46+E47-E48-E49-E50-E51"
        Ws.Cells(52, 17) = "=SUM(E52:P52)"
        Ws.Cells(53, 2) = "成本模块产品销售成本"
        Ws.Cells(53, 3) = "50"
        Ws.Cells(53, 4) = "成本模块产品销售成本"
        Ws.Cells(53, 17) = "=SUM(E53:P53)"
        Ws.Cells(54, 2) = "差异"
        Ws.Cells(54, 3) = "51=49-50"
        Ws.Cells(54, 4) = "差异"
        Ws.Cells(54, 5) = "=E52-E53"
        Ws.Cells(54, 17) = "=SUM(E54:P54)"
        Ws.Cells(55, 2) = "总账主营业务成本"
        Ws.Cells(55, 3) = "52"
        Ws.Cells(55, 4) = "总账主营业务成本"
        Ws.Cells(55, 17) = "=SUM(E55:P55)"
        Ws.Cells(56, 2) = "差异"
        Ws.Cells(56, 3) = "53=49-52"
        Ws.Cells(56, 4) = "差异"
        Ws.Cells(56, 5) = "=+E55-E52"
        Ws.Cells(56, 17) = "=SUM(E56:P56)"

        Ws.Cells(58, 2) = "备注：存货出库均用负数"" - ""符号表示，在报表中需要注意使用正数表示"
        Ws.Cells(59, 2) = "1.报表编辑逻辑：存货的期初库存金额+存货的本期入库金额-存货的本期出库金额=存货的期末库存金额；产品成本的费用演变过程：由费用—生产成本-产品成本-销货成本"
        Ws.Cells(60, 2) = "2.存货的本期入库金额=存货的本期采购入库金额+存货的本期委外入库金额+存货的本期工单入库+存货的本期返工入库金额+存货的本期杂收入库金额"
        Ws.Cells(61, 2) = "3.存货的本期出库金额=存货的本期工单领用出库金额+存货的本期返工工单领用出库金额+存货的本期杂发领用出库金额+存货的本期销货出库金额+存货的本期盘盈亏金额进出库金额+存货的本期结存调整进出库金额"
        Ws.Cells(62, 2) = "4.本期投入的直接材料成本=原材料/包材/低值易耗品期初金额+原材料/包材/低值易耗品本期入库金额-原材料/包材/低值易耗品期末金额-原材料/包材/低值易耗品（+/-本期杂发领用金额+/-盘盈亏金额+/-结存调整金额）-原材料/包材/低值易耗品本期直接销货成本"
        Ws.Cells(63, 2) = "5.直接人工成本是本期总账制造费用-直接人工累计之和。也即是axct311中对应年度和月份的直接人工一栏位的金额"
        Ws.Cells(64, 2) = "6.制造费用是（本期总账制造费用累计之和-本期总账制造费用-直接人工累计之和）的余额。也即是axct311中对应年度和月份的制造费用一栏位的金额"
        Ws.Cells(65, 2) = "7.本期投入的生产成本=本期直接材料+本期直接人工+本期制造费用"
        Ws.Cells(66, 2) = "8.本期产成品成本=（期初在制品金额-期末在制金额-委外物料本期入库的材料金额）+（期初半成品金额-期末半成品金额+本期半成品杂收金额-本期半成品杂发金额-本期半成品盘盈亏金额-本期结存调整金额）"
        Ws.Cells(67, 2) = "9.本期成品的销货成本=期初成品金额+本期成品外购入库金额（不含委外）+本期成品杂收入库金额-期末成品库存金额-本期成品杂发金额-本期成品盘盈亏金额-本期成品结存调整金额"
        Ws.Cells(68, 2) = "10.所有的工单/返工工单入库领用成本不用调整"
        Ws.Cells(69, 2) = "11.因本期委外物料入库金额包括材料金额和加工费用两部分，其中材料成本已经在材料工单/返工工单领用成本中。同时委外物料的期初在制金额减去委外物料的期末在制金额等于 委外物料入库金额包括材料金额和加工费用两部分。故需要把本期委外物料入库材料金额减去"
        Ws.Cells(70, 2) = "12.需要注意各个栏位累计金额的正负情况"

        ' 劃線
        oRng = Ws.Range("B3", "Q56")
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
        oCommand.CommandText += "from cdb_file where cdb01 = " & tYear & " and cdb04 in (" & p2 & ") )"

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
    Private Sub SGetCCC()
        oCommand.CommandText = "select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( select "
        oCommand.CommandText += "(case when ccc03 =1 then ccc221 else 0 end) as t1,"
        oCommand.CommandText += "(case when ccc03 =2 then ccc221 else 0 end) as t2,"
        oCommand.CommandText += "(case when ccc03 =3 then ccc221 else 0 end) as t3,"
        oCommand.CommandText += "(case when ccc03 =4 then ccc221 else 0 end) as t4,"
        oCommand.CommandText += "(case when ccc03 =5 then ccc221 else 0 end) as t5,"
        oCommand.CommandText += "(case when ccc03 =6 then ccc221 else 0 end) as t6,"
        oCommand.CommandText += "(case when ccc03 =7 then ccc221 else 0 end) as t7,"
        oCommand.CommandText += "(case when ccc03 =8 then ccc221 else 0 end) as t8,"
        oCommand.CommandText += "(case when ccc03 =9 then ccc221 else 0 end) as t9,"
        oCommand.CommandText += "(case when ccc03 =10 then ccc221 else 0 end) as t10,"
        oCommand.CommandText += "(case when ccc03 =11 then ccc221 else 0 end) as t11,"
        oCommand.CommandText += "(case when ccc03 =12 then ccc221 else 0 end) as t12 "
        oCommand.CommandText += "from ccc_file left join ima_file on ccc01 = ima01 where ccc02 = " & tYear & " and ima06 in ('103') "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select (case when ccc03 =1 then ccc222 else 0 end) as t1,"
        oCommand.CommandText += "(case when ccc03 =2 then ccc222 else 0 end) as t2,"
        oCommand.CommandText += "(case when ccc03 =3 then ccc222 else 0 end) as t3,"
        oCommand.CommandText += "(case when ccc03 =4 then ccc222 else 0 end) as t4,"
        oCommand.CommandText += "(case when ccc03 =5 then ccc222 else 0 end) as t5,"
        oCommand.CommandText += "(case when ccc03 =6 then ccc222 else 0 end) as t6,"
        oCommand.CommandText += "(case when ccc03 =7 then ccc222 else 0 end) as t7,"
        oCommand.CommandText += "(case when ccc03 =8 then ccc222 else 0 end) as t8,"
        oCommand.CommandText += "(case when ccc03 =9 then ccc222 else 0 end) as t9,"
        oCommand.CommandText += "(case when ccc03 =10 then ccc222 else 0 end) as t10,"
        oCommand.CommandText += "(case when ccc03 =11 then ccc222 else 0 end) as t11,"
        oCommand.CommandText += "(case when ccc03 =12 then ccc222 else 0 end) as t12 "
        oCommand.CommandText += "from ccc_file left join ima_file on ccc01 = ima01 where ccc02 = " & tYear & " and ima06 in ('102','103') )"

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
    Private Sub SGetCCC2(ByVal p1 As String, p2 As String)
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
        oCommand.CommandText += "from ccc_file left join ima_file on ccc01 = ima01 where ccc02 = " & tYear & " and ima06 in (" & p2 & ")"
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