Public Class Form174
    Dim Ds As New DataSet()
    Dim Sda As New SqlClient.SqlDataAdapter
    Dim conn As New SqlClient.SqlConnection()
    Private Sub Form174_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conn.ConnectionString = Module1.OpenConnectionOfERPSUPPORT()
        If conn.State <> ConnectionState.Open Then
            Try
                conn.Open()
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        Sda = New SqlClient.SqlDataAdapter("select * from IES7", conn)
        Sda.Fill(Ds)
        Me.DataGridView1.DataSource = Ds.Tables(0)
        Me.DataGridView1.Columns(0).HeaderText = "周数"
        Me.DataGridView1.Columns(1).HeaderText = "起始时间"
        Me.DataGridView1.Columns(2).HeaderText = "结束时间"
        Me.DataGridView1.Columns(1).DefaultCellStyle.Format = "yyyy-MM-dd hh:mm:ss"
        Me.DataGridView1.Columns(2).DefaultCellStyle.Format = "yyyy-MM-dd hh:mm:ss"
        Me.DataGridView1.Columns(3).HeaderText = "假日名称"
        Me.DataGridView1.Columns(4).HeaderText = "假日数"
        Me.DataGridView1.Columns(5).HeaderText = "是否为休假周"
        Me.DataGridView1.EditMode = DataGridViewEditMode.EditOnEnter
        Me.DataGridView1.AutoResizeColumns()
        Me.DataGridView1.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Ds.HasChanges() Then
            Dim cb As New SqlClient.SqlCommandBuilder(Sda)
            Sda.Update(Ds.Tables(0))
            Ds.Tables(0).AcceptChanges()
            Me.DataGridView1.Update()


        End If
    End Sub
End Class