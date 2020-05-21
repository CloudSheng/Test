Public Class Form1
    Dim CheckAuthorize As Boolean = False
    Dim USN As String = Environment.GetEnvironmentVariable("USERNAME")
    'Private Sub 客ToolStripMenuItem_Click(sender As Object, e As EventArgs)
    '  Form2.Show()
    'Form2.Focus()
    'End Sub

    Private Sub 离开系统ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 离开系统ToolStripMenuItem.Click
        If MsgBox("确定离开系统?", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
            Dim AllFormsCount As Int16 = My.Application.OpenForms.Count
            For i As Int16 = AllFormsCount - 1 To 0 Step -1
                My.Application.OpenForms(i).Close()
            Next
        End If
    End Sub

    Private Sub 成品批次特采ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 成品批次特采ToolStripMenuItem.Click
        Form3.Show()
        Form3.Focus()
    End Sub

    Private Sub 客制工站报表ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Form4.Show()
        Form4.Focus()
    End Sub

    Private Sub 工站呆滞报表ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Form5.Show()
        Form5.Focus()
    End Sub

    Private Sub MES连线测试ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MES连线测试ToolStripMenuItem.Click
        Form6.Show()
        Form6.Focus()
    End Sub

    Private Sub SN批次报废ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SN批次报废ToolStripMenuItem.Click
        Form7.Show()
        Form7.Focus()
    End Sub

    Private Sub 每日盘点设置ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 每日盘点设置ToolStripMenuItem.Click
        Form8.Show()
        Form8.Focus()
    End Sub

    Private Sub 集团资金预估表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 集团资金预估表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form9", USN)
        If CheckAuthorize = True Then
            Form9.Show()
            Form9.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 客制工站完工报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 客制工站完工报表ToolStripMenuItem.Click
        Form10.Show()
        Form10.Focus()
    End Sub

    Private Sub 客制订单报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 客制订单报表ToolStripMenuItem.Click
        Form11.Show()
        Form11.Focus()
    End Sub

    Private Sub ERP连线测试ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ERP连线测试ToolStripMenuItem.Click
        Form12.Show()
        Form12.Focus()
    End Sub

    Private Sub 出货金额及入库金额ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 出货金额及入库金额ToolStripMenuItem.Click
        Form13.Show()
        Form13.Focus()
    End Sub

    Private Sub 客制返工报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 客制返工报表ToolStripMenuItem.Click
        Form14.Show()
        Form14.Focus()
    End Sub

    Private Sub 客制采购单价报表ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Form15.Show()
        Form15.Focus()
    End Sub

    Private Sub 客制采购单价报表非BOMToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Form16.Show()
        Form16.Focus()
    End Sub

    Private Sub 版本资讯ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 版本资讯ToolStripMenuItem.Click
        Form17.Show()
        Form17.Focus()
    End Sub

    Private Sub 客制隔离待判报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 客制隔离待判报表ToolStripMenuItem.Click
        Form18.Show()
        Form18.Focus()
    End Sub

    Private Sub 搬家ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 搬家ToolStripMenuItem.Click
        Form19.Show()
        Form19.Focus()
    End Sub

    Private Sub 客制询价次数报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 客制询价次数报表ToolStripMenuItem.Click
        Form20.Show()
        Form20.Focus()
    End Sub

    Private Sub 模拟展料报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 模拟展料报表ToolStripMenuItem.Click
        Form21.Show()
        Form21.Focus()
    End Sub

    Private Sub 客制报废报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 客制报废报表ToolStripMenuItem.Click
        Form22.Show()
        Form22.Focus()
    End Sub

    Private Sub 客制采购年度报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 客制采购年度报表ToolStripMenuItem.Click
        Form23.Show()
        Form23.Focus()
    End Sub

    Private Sub 客制询价明细报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 客制询价明细报表ToolStripMenuItem.Click
        Form24.Show()
        Form24.Focus()
    End Sub

    Private Sub 产ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 产ToolStripMenuItem.Click
        ' 自動產生返工工單
        Form25.Show()
        Form25.Focus()
    End Sub

    Private Sub TESTToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TESTToolStripMenuItem.Click
        Form71.Show()
        Form71.Focus()
    End Sub

    Private Sub MES型号站别与ERP料号对应表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MES型号站别与ERP料号对应表ToolStripMenuItem.Click
        Form27.Show()
        Form27.Focus()
    End Sub

    'Private Sub ERP自动发料ToolStripMenuItem_Click(sender As Object, e As EventArgs)
    'Form28.Show()
    'Form28.Focus()
    'End Sub

    Private Sub 生管用完工报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 生管用完工报表ToolStripMenuItem.Click
        Form28.Show()
        Form28.Focus()
    End Sub

    Private Sub 付款期报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 付款期报表ToolStripMenuItem.Click
        Form29.Show()
        Form29.Focus()
    End Sub

    Private Sub 专案报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 专案报表ToolStripMenuItem.Click
        Form30.Show()
        Form30.Focus()
    End Sub

    Private Sub 专案人力输入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 专案人力输入ToolStripMenuItem.Click
        Form31.Show()
        Form31.Focus()
    End Sub

    'Private Sub 东莞应付帐龄表ToolStripMenuItem_Click(sender As Object, e As EventArgs)
    '    Form32.Show()
    '    Form32.Focus()
    'End Sub

    Private Sub 东莞物料需求AP模拟ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 东莞物料需求AP模拟ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form33", USN)
        If CheckAuthorize = True Then
            Form33.Show()
            Form33.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 报废报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 报废报表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form34", USN)
        If CheckAuthorize = True Then
            Form34.Show()
            Form34.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 客户别销售WIP库存成本ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 客户别销售WIP库存成本ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form35", USN)
        If CheckAuthorize = True Then
            Form35.Show()
            Form35.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 客制采购月度报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 客制采购月度报表ToolStripMenuItem.Click
        Form38.Show()
        Form38.Focus()
    End Sub

    'Private Sub 经过次数报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 经过次数报表ToolStripMenuItem.Click
    '    Form39.Show()
    '    Form39.Focus()
    'End Sub

    Private Sub 模具AP报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 模具AP报表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form37", USN)
        If CheckAuthorize = True Then
            Form37.Show()
            Form37.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 采购成本指标ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 采购成本指标ToolStripMenuItem.Click
        Form40.Show()
        Form40.Focus()
    End Sub

    Private Sub 模拟展料报表全部料件ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 模拟展料报表全部料件ToolStripMenuItem.Click
        Form41.Show()
        Form41.Focus()
    End Sub

    Private Sub 客制交接报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 客制交接报表ToolStripMenuItem.Click
        Form42.Show()
        Form42.Focus()
    End Sub

    Private Sub 客制盘点报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 客制盘点报表ToolStripMenuItem.Click
        Form129.Show()
        Form129.Focus()
    End Sub

    Private Sub 财务用盘点报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 财务用盘点报表ToolStripMenuItem.Click
        Form43.Show()
        Form43.Focus()
    End Sub

    Private Sub 客制日报废报表ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 客制报废日报表ToolStripMenuItem1.Click
        Form44.Show()
        Form44.Focus()
    End Sub

    Private Sub 抛光段个人绩效报表ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Form45.Show()
        Form45.Focus()
    End Sub

    Private Sub 未结案工单明细ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 未结案工单明细ToolStripMenuItem.Click
        Form46.Show()
        Form46.Focus()
    End Sub

    Private Sub 产品静置时间表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 产品静置时间表ToolStripMenuItem.Click
        Form47.Show()
        Form47.Focus()
    End Sub

    Private Sub 工单用量分析表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 工单用量分析表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form48", USN)
        If CheckAuthorize = True Then
            Form48.Show()
            Form48.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 杂收与库存对照表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 杂收与库存对照表ToolStripMenuItem.Click
        Form49.Show()
        Form49.Focus()
    End Sub

    Private Sub HACToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HACToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form50", USN)
        If CheckAuthorize = True Then
            Form50.Show()
            Form50.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub DAC月销售报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DAC月销售报表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form51", USN)
        If CheckAuthorize = True Then
            Form51.Show()
            Form51.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub HACAR账龄报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HACAR账龄报表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form52", USN)
        If CheckAuthorize = True Then
            Form52.Show()
            Form52.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If

    End Sub

    Private Sub DACAR账龄报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DACAR账龄报表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form53", USN)
        If CheckAuthorize = True Then
            Form53.Show()
            Form53.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub ACAAR账龄报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ACAAR账龄报表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form54", USN)
        If CheckAuthorize = True Then
            Form54.Show()
            Form54.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 抛光段月ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Form55.Show()
        Form55.Focus()
    End Sub

    Private Sub DAC最后单价ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DAC最后单价ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form56", USN)
        If CheckAuthorize = True Then
            Form56.Show()
            Form56.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub HACAR账龄报表ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles HACAR账龄报表ToolStripMenuItem1.Click
        CheckAuthorize = CheckAuthorizeByUser("Form57", USN)
        If CheckAuthorize = True Then
            Form57.Show()
            Form57.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub DACAR账龄报表ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles DACAR账龄报表ToolStripMenuItem1.Click
        CheckAuthorize = CheckAuthorizeByUser("Form58", USN)
        If CheckAuthorize = True Then
            Form58.Show()
            Form58.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub ACA账龄报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ACA账龄报表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form59", USN)
        If CheckAuthorize = True Then
            Form59.Show()
            Form59.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub ERP入库与MES移转查核ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ERP入库与MES移转查核ToolStripMenuItem.Click
        Form60.Show()
        Form60.Focus()
    End Sub

    Private Sub 生管用月报废报表ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Form61.Show()
        Form61.Focus()
    End Sub

    Private Sub 返修记录报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 返修记录报表ToolStripMenuItem.Click
        Form62.Show()
        Form62.Focus()
    End Sub

    Private Sub 年度应付报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 年度应付报表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form63", USN)
        If CheckAuthorize = True Then
            Form63.Show()
            Form63.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub MES与ERP数量比对表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MES与ERP数量比对表ToolStripMenuItem.Click
        Form64.Show()
        Form64.Focus()
    End Sub

    Private Sub 主材料损耗率报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 主材料损耗率报表ToolStripMenuItem.Click
        Form65.Show()
        Form65.Focus()
    End Sub

    Private Sub MES基础资料报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MES基础资料报表ToolStripMenuItem.Click
        Form66.Show()
        Form66.Focus()
    End Sub

    Private Sub 计算机ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 计算机ToolStripMenuItem.Click
        Form67.Show()
        Form67.Focus()
    End Sub


    Private Sub 不良品库存与返工工单匹配表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 不良品库存与返工工单匹配表ToolStripMenuItem.Click
        Form68.Show()
        Form68.Focus()
    End Sub

    Private Sub 请购单汇入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 请购单汇入ToolStripMenuItem.Click
        Form69.Show()
        Form69.Focus()
    End Sub

    Private Sub 专案工时报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 专案工时报表ToolStripMenuItem.Click
        Form70.Show()
        Form70.Focus()
    End Sub

    Private Sub WIP分站存量报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WIP分站存量报表ToolStripMenuItem.Click
        Form72.Show()
        Form72.Focus()
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        CheckAuthorize = CheckAuthorizeByUser("Form73", USN)
        If CheckAuthorize = True Then
            Form73.Show()
            Form73.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If

    End Sub

    Private Sub BVIAR账龄报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BVIAR账龄报表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form74", USN)
        If CheckAuthorize = True Then
            Form74.Show()
            Form74.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub MES样品进度报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MES样品进度报表ToolStripMenuItem.Click
        Form75.Show()
        Form75.Focus()
    End Sub

    Private Sub HAC销售日报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HAC销售日报表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form76", USN)
        If CheckAuthorize = True Then
            Form76.Show()
            Form76.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 涂装WIP报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 涂装WIP报表ToolStripMenuItem.Click
        Form77.Show()
        Form77.Focus()
    End Sub

    Private Sub WIP库龄报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WIP库龄报表ToolStripMenuItem.Click
        Form78.Show()
        Form78.Focus()
    End Sub

    Private Sub 资材ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 资材ToolStripMenuItem.Click
        Form301.Show()
        Form301.Focus()
    End Sub

    Private Sub 批号SN明细表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 批号SN明细表ToolStripMenuItem.Click
        Form79.Show()
        Form79.Focus()
    End Sub

    Private Sub 成本变动表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 成本变动表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form80", USN)
        If CheckAuthorize = True Then
            Form80.Show()
            Form80.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 单阶材料用途表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 单阶材料用途表ToolStripMenuItem.Click
        Form302.Show()
        Form302.Focus()
    End Sub

    Private Sub 材料用途查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 材料用途查询ToolStripMenuItem.Click
        Form81.Show()
        Form81.Focus()
    End Sub

    Private Sub 采购及委外成本变动表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 采购及委外成本变动表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form303", USN)
        If CheckAuthorize = True Then
            Form303.Show()
            Form303.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 杂发单ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 杂发单ToolStripMenuItem.Click
        Form304.Show()
        Form304.Focus()
    End Sub

    Private Sub 销售费用统计表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 销售费用统计表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form82", USN)
        If CheckAuthorize = True Then
            Form82.Show()
            Form82.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 制造费用表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 制造费用表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form83", USN)
        If CheckAuthorize = True Then
            Form83.Show()
            Form83.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 研发费用表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 研发费用表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form84", USN)
        If CheckAuthorize = True Then
            Form84.Show()
            Form84.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 管理费用表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 管理费用表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form85", USN)
        If CheckAuthorize = True Then
            Form85.Show()
            Form85.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 排程汇入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 排程汇入ToolStripMenuItem.Click
        Form86.Show()
        Form86.Focus()
    End Sub

    Private Sub 胶合周排程用量表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 胶合周排程用量表ToolStripMenuItem.Click
        Form87.Show()
        Form87.Focus()
    End Sub

    'Private Sub 毛利分析表ToolStripMenuItem_Click(sender As Object, e As EventArgs)
    '   Form88.Show()
    '  Form88.Focus()
    'End Sub

    Private Sub 材料用途报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 材料用途报表ToolStripMenuItem.Click
        Form89.Show()
        Form89.Focus()
    End Sub

    Private Sub WorkingCapitalToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WorkingCapitalToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form90", USN)
        If CheckAuthorize = True Then
            Form90.Show()
            Form90.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 采购委外料件杂收发明细表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 采购委外料件杂收发明细表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form91", USN)
        If CheckAuthorize = True Then
            Form91.Show()
            Form91.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 工时变更汇入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 工时变更汇入ToolStripMenuItem.Click
        Form305.Show()
        Form305.Focus()
    End Sub

    Private Sub DAC客户价格比价ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DAC客户价格比价ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form92", USN)
        If CheckAuthorize = True Then
            Form92.Show()
            Form92.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 毛利分析表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 毛利分析表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form88", USN)
        If CheckAuthorize = True Then
            Form88.Show()
            Form88.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 项目案预算比较表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 项目案预算比较表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form93", USN)
        If CheckAuthorize = True Then
            Form93.Show()
            Form93.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 客户预ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 客户预ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form94", USN)
        If CheckAuthorize = True Then
            Form94.Show()
            Form94.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub MES报废月报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MES报废月报表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form95", USN)
        If CheckAuthorize = True Then
            Form95.Show()
            Form95.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 银行存款余额调节表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 银行存款余额调节表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form96", USN)
        If CheckAuthorize = True Then
            Form96.Show()
            Form96.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub ERP工单套数表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ERP工单套数表ToolStripMenuItem.Click
        Form97.Show()
        Form97.Focus()
    End Sub

    Private Sub BOM表物料标准成本与实际成本明细表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BOM表物料标准成本与实际成本明细表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form98", USN)
        If CheckAuthorize = True Then
            Form98.Show()
            Form98.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 工单上阶在制成本明细表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 工单上阶在制成本明细表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form99", USN)
        If CheckAuthorize = True Then
            Form99.Show()
            Form99.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub TEST2ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TEST2ToolStripMenuItem.Click
        Form100.Show()
        Form100.Focus()
    End Sub

    Private Sub 标准成本变动表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 标准成本变动表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form101", USN)
        If CheckAuthorize = True Then
            Form101.Show()
            Form101.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 标准成本明细表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 标准成本明细表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form102", USN)
        If CheckAuthorize = True Then
            Form102.Show()
            Form102.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 验退仓退明细表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 验退仓退明细表ToolStripMenuItem.Click
        Form103.Show()
        Form103.Focus()
    End Sub

    Private Sub 胶合WIP报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 胶合WIP报表ToolStripMenuItem.Click
        Form104.Show()
        Form104.Focus()
    End Sub

    Private Sub 接收站库龄表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 接收站库龄表ToolStripMenuItem.Click
        Form105.Show()
        Form105.Focus()
    End Sub

    Private Sub 工单下阶在制成本明细表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 工单下阶在制成本明细表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form106", USN)
        If CheckAuthorize = True Then
            Form106.Show()
            Form106.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 报废品分析表依人力分布ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Form107.Show()
        Form107.Focus()
    End Sub

    Private Sub 成品库龄报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 成品库龄报表ToolStripMenuItem.Click
        Form108.Show()
        Form108.Focus()
    End Sub

    Private Sub 库龄成本分析ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 库龄成本分析ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form109", USN)
        If CheckAuthorize = True Then
            Form109.Show()
            Form109.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 客制报废解锁报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 客制报废解锁报表ToolStripMenuItem.Click
        Form110.Show()
        Form110.Focus()
    End Sub

    Private Sub 生管用过程缺陷报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 生管用过程缺陷报表ToolStripMenuItem.Click
        Form111.Show()
        Form111.Focus()
    End Sub

    Private Sub 周出货计划跟进表ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Form112.Show()
        Form112.Focus()
    End Sub

    'Private Sub ShipmentManagementToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShipmentManagementToolStripMenuItem.Click
    '   Form113.Show()
    '  Form113.Focus()
    'End Sub

    Private Sub HAC销售周报ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HAC销售周报ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form114", USN)
        If CheckAuthorize = True Then
            Form114.Show()
            Form114.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub DACInvoice单价分析报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DACInvoice单价分析报表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form115", USN)
        If CheckAuthorize = True Then
            Form115.Show()
            Form115.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 入库计划汇入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 入库计划汇入ToolStripMenuItem.Click
        Form116.Show()
        Form116.Focus()
    End Sub

    Private Sub ERP安全存量汇入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ERP安全存量汇入ToolStripMenuItem.Click
        Form117.Show()
        Form117.Focus()
    End Sub

    'Private Sub 多交期汇入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 多交期汇入ToolStripMenuItem.Click
    '   Form118.Show()
    '  Form118.Focus()
    'End Sub

    Private Sub 标准采购单价与实际采购单价比较表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 标准采购单价与实际采购单价比较表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form119", USN)
        If CheckAuthorize = True Then
            Form119.Show()
            Form119.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub MESToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles MESToolStripMenuItem1.Click
        Form120.Show()
        Form120.Focus()
    End Sub

    'Private Sub DemandAndScheduleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DemandAndScheduleToolStripMenuItem.Click
    '    Form121.Show()
    '    Form121.Focus()
    'End Sub

    Private Sub HACToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles HACToolStripMenuItem1.Click
        CheckAuthorize = CheckAuthorizeByUser("Form122", USN)
        If CheckAuthorize = True Then
            Form122.Show()
            Form122.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 东莞应付帐龄表ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 东莞应付帐龄表ToolStripMenuItem1.Click
        CheckAuthorize = CheckAuthorizeByUser("Form32", USN)
        If CheckAuthorize = True Then
            Form32.Show()
            Form32.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub DAC利润表ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        CheckAuthorize = CheckAuthorizeByUser("Form123", USN)
        If CheckAuthorize = True Then
            Form123.Show()
            Form123.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub FinicialReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FinicialReportToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form124", USN)
        If CheckAuthorize = True Then
            Form124.Show()
            Form124.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub ISToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ISToolStripMenuItem.Click
        ' Cindy 無法做權限檢查
        If USN = "Cindy.Ye" Then
            Form123.Show()
            Form123.Focus()
        Else
            CheckAuthorize = CheckAuthorizeByUser("Form123", USN)
            If CheckAuthorize = True Then
                Form123.Show()
                Form123.Focus()
            Else
                MsgBox("UnAuthorized")
                Return
            End If
        End If
    End Sub

    Private Sub BSToolStripMenuItem_Click(sender As Object, e As EventArgs)
        CheckAuthorize = CheckAuthorizeByUser("Form125", USN)
        If CheckAuthorize = True Then
            Form125.Show()
            Form125.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 合并财务报表ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        CheckAuthorize = CheckAuthorizeByUser("Form125", USN)
        If CheckAuthorize = True Then
            Form125.Show()
            Form125.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub BSToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles BSToolStripMenuItem.Click
        ' Cindy 無法做權限檢查
        If USN = "Cindy.Ye" Then
            Form125.Show()
            Form125.Focus()
        Else
            CheckAuthorize = CheckAuthorizeByUser("Form125", USN)
            If CheckAuthorize = True Then
                Form125.Show()
                Form125.Focus()
            Else
                MsgBox("UnAuthorized")
                Return
            End If
        End If
    End Sub

    Private Sub AccountingReportsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AccountingReportsToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form126", USN)
        If CheckAuthorize = True Then
            Form126.Show()
            Form126.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 品质报废周报金额统计表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 品质报废周报金额统计表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form350", USN)
        If CheckAuthorize = True Then
            Form350.Show()
            Form350.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub MPL接单记录ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MPL接单记录ToolStripMenuItem.Click
        Form351.Show()
        Form351.Focus()
    End Sub

    Private Sub CashFlowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CashFlowToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form127", USN)
        If CheckAuthorize = True Then
            Form127.Show()
            Form127.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub CashFlowForeCastToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CashFlowForeCastToolStripMenuItem.Click
        'CheckAuthorize = CheckAuthorizeByUser("Form128", USN)
        'If CheckAuthorize = True Then
        'Form128.Show()
        'Form128.Focus()
        'Else
        'MsgBox("UnAuthorized")
        'Return
        'End If
    End Sub

    Private Sub 客诉费用统计表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 客诉费用统计表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form352", USN)
        If CheckAuthorize = True Then
            Form352.Show()
            Form352.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 空运费用统计表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 空运费用统计表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form353", USN)
        If CheckAuthorize = True Then
            Form353.Show()
            Form353.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 运费汇入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 运费汇入ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form130", USN)
        If CheckAuthorize = True Then
            Form130.Show()
            Form130.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 运费明细表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 运费明细表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form131", USN)
        If CheckAuthorize = True Then
            Form131.Show()
            Form131.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 出口货物报关单ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 出口货物报关单ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form354", USN)
        If CheckAuthorize = True Then
            Form354.Show()
            Form354.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub SalesCompareWithBudgetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalesCompareWithBudgetToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form132", USN)
        If CheckAuthorize = True Then
            Form132.Show()
            Form132.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub MachineDefectReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MachineDefectReportToolStripMenuItem.Click
        Form133.Show()
        Form133.Focus()
    End Sub

    Private Sub 部门费用表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 部门费用表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form134", USN)
        If CheckAuthorize = True Then
            Form134.Show()
            Form134.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 部门杂收发表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 部门杂收发表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form135", USN)
        If CheckAuthorize = True Then
            Form135.Show()
            Form135.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 采购入库金额单价采购量统计表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 采购入库金额单价采购量统计表ToolStripMenuItem.Click
        Form136.Show()
        Form136.Focus()
    End Sub

    Private Sub 產品直通率週報ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 產品直通率週報ToolStripMenuItem.Click
        Form137.Show()
        Form137.Focus()
    End Sub

    Private Sub PWC产品分析明细表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PWC产品分析明细表ToolStripMenuItem.Click
        Form138.Show()
        Form138.Focus()
    End Sub

    Private Sub InventoryCountToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InventoryCountToolStripMenuItem.Click
        Form139.Show()
        Form139.Focus()
    End Sub

    Private Sub 总工时重计ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 总工时重计ToolStripMenuItem.Click
        Form140.Show()
        Form140.Focus()
    End Sub

    Private Sub RD纱料用量表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RD纱料用量表ToolStripMenuItem.Click
        Form141.Show()
        Form141.Focus()
    End Sub

    Private Sub 进口手册报关单ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 进口手册报关单ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form355", USN)
        If CheckAuthorize = True Then
            Form355.Show()
            Form355.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 进口贸易报关单ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 进口贸易报关单ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form356", USN)
        If CheckAuthorize = True Then
            Form356.Show()
            Form356.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 外币余额表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 外币余额表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form357", USN)
        If CheckAuthorize = True Then
            Form357.Show()
            Form357.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub RD人工工时月报ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RD人工工时月报ToolStripMenuItem.Click
        Form143.Show()
        Form143.Focus()
    End Sub

    Private Sub RollingForecastToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RollingForecastToolStripMenuItem.Click
        'Form144.Show()
        'Form144.Focus()
    End Sub

    Private Sub 经过次数报表ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 经过次数报表ToolStripMenuItem1.Click
        Form145.Show()
        Form145.Focus()
    End Sub

    Private Sub 應收帳款预测表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 應收帳款预测表ToolStripMenuItem.Click
        'CheckAuthorize = CheckAuthorizeByUser("Form146", USN)
        'If CheckAuthorize = True Then
        'Form146.Show()
        'Form146.Focus()
        'Else
        'MsgBox("UnAuthorized")
        'Return
        'End If
    End Sub

    Private Sub QC称重客制报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QC称重客制报表ToolStripMenuItem.Click
        Form147.Show()
        Form147.Focus()
    End Sub

    Private Sub ForecastToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ForecastToolStripMenuItem.Click
        Form358.Show()
        Form358.Focus()
    End Sub

    Private Sub 成本倒扎表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 成本倒扎表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form148", USN)
        If CheckAuthorize = True Then
            Form148.Show()
            Form148.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 审计倒扎表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 审计倒扎表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form149", USN)
        If CheckAuthorize = True Then
            Form149.Show()
            Form149.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
        Form359.Show()
        Form359.Focus()
    End Sub

    Private Sub 样品报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 样品报表ToolStripMenuItem.Click
        Form150.Show()
        Form150.Focus()
    End Sub

    Private Sub ToolStripMenuItem7_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem7.Click
        Form151.Show()
        Form151.Focus()
    End Sub

    Private Sub ToolStripMenuItem8_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem8.Click
        Form152.Show()
        Form152.Focus()
    End Sub

    Private Sub ToolStripMenuItem9_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem9.Click
        Form153.Show()
        Form153.Focus()
    End Sub

    Private Sub 损益表汇入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 损益表汇入ToolStripMenuItem.Click
        Form154.Show()
        Form154.Focus()
    End Sub

    Private Sub ACAISToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ACAISToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form155", USN)
        If CheckAuthorize = True Then
            Form155.Show()
            Form155.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 资产负债表汇入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 资产负债表汇入ToolStripMenuItem.Click
        Form156.Show()
        Form156.Focus()
    End Sub

    Private Sub ACABSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ACABSToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form157", USN)
        If CheckAuthorize = True Then
            Form157.Show()
            Form157.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub ToolStripMenuItem10_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem10.Click
        CheckAuthorize = CheckAuthorizeByUser("Form158", USN)
        If CheckAuthorize = True Then
            Form158.Show()
            Form158.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 关键零部件报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 关键零部件报表ToolStripMenuItem.Click
        Form159.Show()
        Form159.Focus()
    End Sub

    Private Sub SN库龄表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SN库龄表ToolStripMenuItem.Click
        Form160.Show()
        Form160.Focus()
    End Sub

    Private Sub 部门费用及预算汇总表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 部门费用及预算汇总表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form161", USN)
        If CheckAuthorize = True Then
            Form161.Show()
            Form161.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 功能主管费用表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 功能主管费用表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form162", USN)
        If CheckAuthorize = True Then
            Form162.Show()
            Form162.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub CallOff汇入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CallOff汇入ToolStripMenuItem.Click
        Form163.Show()
        Form163.Focus()
    End Sub

    Private Sub 顾客PN检查报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 顾客PN检查报表ToolStripMenuItem.Click
        Form164.Show()
        Form164.Focus()
    End Sub

    Private Sub DAC销售成本毛利资料报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DAC销售成本毛利资料报表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form165", USN)
        If CheckAuthorize = True Then
            Form165.Show()
            Form165.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub FIFO计算ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FIFO计算ToolStripMenuItem.Click
        Form166.Show()
        Form166.Focus()
    End Sub

    Private Sub 資金計畫表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 資金計畫表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form168", USN)
        If CheckAuthorize = True Then
            Form168.Show()
            Form168.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 传票汇入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 传票汇入ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form167", USN)
        If CheckAuthorize = True Then
            Form167.Show()
            Form167.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 汇入VACARToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 汇入VACARToolStripMenuItem.Click
        ' Cindy 無法做權限檢查
        If USN = "Cindy.Ye" Then
            Form169.Show()
            Form169.Focus()
        Else
            CheckAuthorize = CheckAuthorizeByUser("Form169", USN)
            If CheckAuthorize = True Then
                Form169.Show()
                Form169.Focus()
            Else
                MsgBox("UnAuthorized")
                Return
            End If
        End If
    End Sub

    Private Sub 資金計畫表匯入入庫計劃ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 資金計畫表匯入入庫計劃ToolStripMenuItem.Click
        Form170.Show()
        Form170.Focus()
    End Sub

    Private Sub TrainfreightSIToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TrainfreightSIToolStripMenuItem.Click
        Form361.Show()
        Form361.Focus()
    End Sub

    Private Sub Tooling完工比例报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Tooling完工比例报表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form171", USN)
        If CheckAuthorize = True Then
            Form171.Show()
            Form171.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub SeafreightSIToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SeafreightSIToolStripMenuItem.Click
        Form362.Show()
        Form362.Focus()
    End Sub

    Private Sub 产品EOP时间批次更新ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 产品EOP时间批次更新ToolStripMenuItem.Click
        Form363.Show()
        Form363.Focus()
    End Sub

    Private Sub VDALabel列印资料ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VDALabel列印资料ToolStripMenuItem.Click
        Form364.Show()
        Form364.Focus()
    End Sub

    Private Sub 周生产资源计划报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 周生产资源计划报表ToolStripMenuItem.Click
        Form172.Show()
        Form172.Focus()
    End Sub

    Private Sub 客制移站报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 客制移站报表ToolStripMenuItem.Click
        Form176.Show()
        Form176.Focus()
    End Sub

    Private Sub DAC销售预算报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DAC销售预算报表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form177", USN)
        If CheckAuthorize = True Then
            Form177.Show()
            Form177.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub DACToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DACToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form178", USN)
        If CheckAuthorize = True Then
            Form178.Show()
            Form178.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 主要物料进料需求表购料资金ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 主要物料进料需求表购料资金ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form179", USN)
        If CheckAuthorize = True Then
            Form179.Show()
            Form179.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 每周费用汇总报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 每周费用汇总报表ToolStripMenuItem.Click
        Form180.Show()
        Form180.Focus()
    End Sub

    Private Sub 待结案料号明细表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 待结案料号明细表ToolStripMenuItem.Click
        Form181.Show()
        Form181.Focus()
    End Sub

    Private Sub 原材料进料计划达成汇总报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 原材料进料计划达成汇总报表ToolStripMenuItem.Click
        Form182.Show()
        Form182.Focus()
    End Sub

    Private Sub 成型模具使用次数ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 成型模具使用次数ToolStripMenuItem.Click
        Form183.Show()
        Form183.Focus()
    End Sub

    Private Sub Calloff多交期汇入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Calloff多交期汇入ToolStripMenuItem.Click
        Form365.Show()
        Form365.Focus()
    End Sub

    Private Sub 原料料进料计划汇入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 原料料进料计划汇入ToolStripMenuItem.Click
        Form366.Show()
        Form366.Focus()
    End Sub

    Private Sub 毛利率报表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 毛利率报表ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form186", USN)
        If CheckAuthorize = True Then
            Form186.Show()
            Form186.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub 关务cxmt691资料回写ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 关务cxmt691资料回写ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form367", USN)
        If CheckAuthorize = True Then
            Form367.Show()
            Form367.Focus()
        Else
            'MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub SN扫描状态ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SN扫描状态ToolStripMenuItem.Click
        Form188.Show()
        Form188.Focus()
    End Sub

    Private Sub ToolStripMenuItem11_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem11.Click
        CheckAuthorize = CheckAuthorizeByUser("Form368", USN)
        If CheckAuthorize = True Then
            Form368.Show()
            Form368.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub
    Private Sub ToolStripMenuItem12_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem12.Click
        CheckAuthorize = CheckAuthorizeByUser("Form369", USN)
        If CheckAuthorize = True Then
            Form369.Show()
            Form369.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub
    Private Sub ToolStripMenuItem13_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem13.Click
        CheckAuthorize = CheckAuthorizeByUser("Form191", USN)
        If CheckAuthorize = True Then
            Form191.Show()
            Form191.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub
    Private Sub ToolStripMenuItem14_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem14.Click
        Form192.Show()
        Form192.Focus()
    End Sub
    Private Sub ToolStripMenuItem15_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem15.Click
        Form193.Show()
        Form193.Focus()
    End Sub
    Private Sub ToolStripMenuItem16_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem16.Click
        CheckAuthorize = CheckAuthorizeByUser("Form194", USN)
        If CheckAuthorize = True Then
            Form194.Show()
            Form194.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub
    Private Sub EDI資料產生ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EDI資料產生ToolStripMenuItem.Click
        CheckAuthorize = CheckAuthorizeByUser("Form370", USN)
        If CheckAuthorize = True Then
            Form370.Show()
            Form370.Focus()
        Else
            MsgBox("UnAuthorized")
            Return
        End If
    End Sub

    Private Sub AaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AaToolStripMenuItem.Click
        Form184.Show()
        Form184.Focus()
    End Sub
End Class
