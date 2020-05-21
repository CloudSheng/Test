Public Class Form67

   
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RadioButton1.Checked Then
            If String.IsNullOrEmpty(TextBox1.Text) Then
                MsgBox("请输入买价")
                Return
            End If
            If Not Decimal.TryParse(TextBox1.Text, 1) Then
                MsgBox("买价请输入数字")
                Return
            End If
            If Me.TextBox1.Text <= 0 Then
                MsgBox("买价应大于0")
                Return
            End If
            If String.IsNullOrEmpty(TextBox2.Text) Then
                MsgBox("请输入卖价")
                Return
            End If
            If Not Decimal.TryParse(TextBox2.Text, 1) Then
                MsgBox("卖价请输入数字")
                Return
            End If
            If Me.TextBox2.Text <= 0 Then
                MsgBox("卖价应大于0")
                Return
            End If
            If String.IsNullOrEmpty(TextBox3.Text) Then
                MsgBox("请输入数量")
                Return
            End If
            If Not Decimal.TryParse(TextBox3.Text, 1) Then
                MsgBox("数量请输入数字")
                Return
            End If
            If Me.TextBox3.Text <= 0 Then
                MsgBox("数量应大于0")
                Return
            End If
            Count1()
        End If
        If RadioButton2.Checked Then
            If String.IsNullOrEmpty(TextBox1.Text) Then
                MsgBox("请输入买价")
                Return
            End If
            If Not Decimal.TryParse(TextBox1.Text, 1) Then
                MsgBox("买价请输入数字")
                Return
            End If
            If Me.TextBox1.Text <= 0 Then
                MsgBox("买价应大于0")
                Return
            End If
            If String.IsNullOrEmpty(TextBox4.Text) Then
                MsgBox("请输入预期利润")
                Return
            End If
            If Not Decimal.TryParse(TextBox4.Text, 1) Then
                MsgBox("预期利润请输入数字")
                Return
            End If
            'If Me.TextBox2.Text <= 0 Then
            'MsgBox("卖价应大于0")
            'Return
            'End If
            If String.IsNullOrEmpty(TextBox3.Text) Then
                MsgBox("请输入数量")
                Return
            End If
            If Not Decimal.TryParse(TextBox3.Text, 1) Then
                MsgBox("数量请输入数字")
                Return
            End If
            If Me.TextBox3.Text <= 0 Then
                MsgBox("数量应大于0")
                Return
            End If
            Count2()
        End If
        If RadioButton3.Checked Then
            If String.IsNullOrEmpty(TextBox2.Text) Then
                MsgBox("请输入卖价")
                Return
            End If
            If Not Decimal.TryParse(TextBox2.Text, 1) Then
                MsgBox("卖价请输入数字")
                Return
            End If
            If Me.TextBox2.Text <= 0 Then
                MsgBox("卖价应大于0")
                Return
            End If
            If String.IsNullOrEmpty(TextBox4.Text) Then
                MsgBox("请输入预期利润")
                Return
            End If
            If Not Decimal.TryParse(TextBox4.Text, 1) Then
                MsgBox("预期利润请输入数字")
                Return
            End If
            If String.IsNullOrEmpty(TextBox3.Text) Then
                MsgBox("请输入数量")
                Return
            End If
            If Not Decimal.TryParse(TextBox3.Text, 1) Then
                MsgBox("数量请输入数字")
                Return
            End If
            If Me.TextBox3.Text <= 0 Then
                MsgBox("数量应大于0")
                Return
            End If
            Count3()
        End If
    End Sub
    Private Sub Count1()
        Dim BuyAmount As Decimal = TextBox1.Text * TextBox3.Text
        Label5.Text = BuyAmount
        Dim BuyCharge As Decimal = Decimal.Floor(BuyAmount * 0.0008)
        Label9.Text = BuyCharge
        Dim BuyTotal As Decimal = BuyAmount + BuyCharge
        Label15.Text = BuyTotal

        'Sell
        Dim SellAmount As Decimal = TextBox2.Text * TextBox3.Text
        Label7.Text = SellAmount
        Dim SellCharge As Decimal = Decimal.Floor(SellAmount * 0.0008)
        Label11.Text = SellCharge
        Dim SellTax As Decimal = Decimal.Round(SellAmount * 0.003)
        Label13.Text = SellTax
        Dim SellTotal As Decimal = SellAmount - SellCharge - SellTax
        Label17.Text = SellTotal

        'Profit
        Dim PF As Decimal = SellTotal - BuyTotal
        TextBox4.Text = PF
        Dim PFP As Decimal = PF / BuyAmount
        Label20.Text = Decimal.Round(PFP * 100, 2) & "%"
    End Sub
    Private Sub Count2()
        Dim BuyAmount As Decimal = TextBox1.Text * TextBox3.Text
        Label5.Text = BuyAmount
        Dim BuyCharge As Decimal = Decimal.Floor(BuyAmount * 0.0008)
        Label9.Text = BuyCharge
        Dim BuyTotal As Decimal = BuyAmount + BuyCharge
        Label15.Text = BuyTotal

        Dim ExpectSellTotal As Decimal = BuyTotal + TextBox4.Text
        Dim ExpectSellAmount As Decimal = Decimal.Floor(ExpectSellTotal / 0.9962)
        Label7.Text = ExpectSellAmount
        Dim ExpectSellPrice As Decimal = ExpectSellAmount / Me.TextBox3.Text
        Select Case ExpectSellPrice
            Case Is < 10
                ExpectSellPrice = Decimal.Round(ExpectSellPrice, 2)
            Case 10.01 To 50
                ExpectSellPrice = Decimal.Round(ExpectSellPrice * 2, 1) / 2
            Case 50 To 100
                ExpectSellPrice = Decimal.Round(ExpectSellPrice, 1)
            Case 100 To 500
                ExpectSellPrice = Decimal.Round(ExpectSellPrice * 2, 0) / 2
            Case 500 To 1000
                ExpectSellPrice = Decimal.Round(ExpectSellPrice, 0)
        End Select
        TextBox2.Text = ExpectSellPrice
        Count1()
    End Sub
    Private Sub Count3()
        Dim SellAmount As Decimal = TextBox2.Text * TextBox3.Text
        Label7.Text = SellAmount
        Dim SellCharge As Decimal = Decimal.Floor(SellAmount * 0.0008)
        Label11.Text = SellCharge
        Dim SellTax As Decimal = Decimal.Round(Convert.ToDecimal(SellAmount * 0.003), 0)
        Dim SellTotal As Decimal = SellAmount - SellCharge - SellTax
        Label17.Text = SellTotal

        Dim ExpectBuyTotal As Decimal = SellTotal - TextBox4.Text
        Dim ExpectBuyAmount As Decimal = Decimal.Floor(ExpectBuyTotal / 0.9992)
        Label5.Text = ExpectBuyAmount
        Dim ExpectBuyPrice As Decimal = ExpectBuyAmount / Me.TextBox3.Text
        Select Case ExpectBuyPrice
            Case Is < 10
                ExpectBuyPrice = Decimal.Round(ExpectBuyPrice, 2)
            Case 10.01 To 50
                ExpectBuyPrice = Decimal.Round(ExpectBuyPrice * 2, 1) / 2
            Case 50.01 To 100
                ExpectBuyPrice = Decimal.Round(ExpectBuyPrice, 1)
            Case 100.01 To 500
                ExpectBuyPrice = Decimal.Round(ExpectBuyPrice * 2, 0) / 2
            Case 500 To 1000
                ExpectBuyPrice = Decimal.Round(ExpectBuyPrice, 0)
        End Select
        TextBox1.Text = ExpectBuyPrice
        Count1()
    End Sub
End Class