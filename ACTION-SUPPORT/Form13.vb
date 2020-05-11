Public Class Form13
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    'Dim Depart As String = String.Empty

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        oCommand.CommandText = "select sum(oga50 * oga24) from oga_file where oga09 in ('2','3','4','6') and ogaconf = 'Y' and ogapost = 'Y' "
        oCommand.CommandText += "and oga55 = '1' and oga02 between to_date('" & DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += DateTimePicker2.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        Dim TotalSale As Decimal = oCommand.ExecuteScalar()
        Label4.Text = TotalSale
    End Sub

    Private Sub Form13_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If Now.Month < 10 Then
            TextBox1.Text = Now.Year & "0" & Now.Month
        Else
            TextBox1.Text = Now.Year & Now.Month
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If IsNothing(ComboBox1.SelectedItem) Then
            MsgBox("请选择部门")
            Return
        End If
        If String.IsNullOrEmpty(TextBox1.Text) Then
            MsgBox("请输入年月")
            Return
        End If
        Dim Depart As String = String.Empty
        Select Case ComboBox1.SelectedItem.ToString()
            Case "裁纱"
                Depart = "D3531"
            Case "预型"
                Depart = "D3532"
            Case "成型"
                Depart = "D3535"
            Case "CNC"
                Depart = "D3536"
            Case "补土"
                Depart = "D3561"
            Case "涂装"
                Depart = "D3563"
            Case "胶合"
                Depart = "D3564"
            Case "抛光"
                Depart = "D3565"
            Case "包装"
                Depart = "D3566"
        End Select
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        oCommand.CommandText = "select nvl(sum(sfv09 * ccc23),0) from sfu_file,sfv_file,ccc_file where sfu01 = sfv01 and sfupost = 'Y' "
        oCommand.CommandText += "and sfuconf = 'Y' and sfu02 between to_date('" & DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += DateTimePicker2.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and sfu04 = '"
        oCommand.CommandText += Depart & "' and sfv04 = ccc01 and ccc02 = " & Strings.Left(TextBox1.Text, 4)
        oCommand.CommandText += "and ccc03 = " & Strings.Right(TextBox1.Text, 2)
        Dim TotalIn As Decimal = oCommand.ExecuteScalar()
        Label6.Text = TotalIn
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter
        
    End Sub
End Class