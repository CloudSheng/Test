<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form8
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意:  以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form8))
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.InventoryCountPrepareBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.IQMES3DataSet = New ACTION_SUPPORT.IQMES3DataSet()
        Me.InventoryCountPrepareTableAdapter = New ACTION_SUPPORT.IQMES3DataSetTableAdapters.InventoryCountPrepareTableAdapter()
        Me.Button1 = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.InventoryCountPrepareBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.IQMES3DataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(12, 3)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 27
        Me.DataGridView1.Size = New System.Drawing.Size(471, 526)
        Me.DataGridView1.TabIndex = 0
        '
        'InventoryCountPrepareBindingSource
        '
        Me.InventoryCountPrepareBindingSource.DataMember = "InventoryCountPrepare"
        Me.InventoryCountPrepareBindingSource.DataSource = Me.IQMES3DataSet
        '
        'IQMES3DataSet
        '
        Me.IQMES3DataSet.DataSetName = "IQMES3DataSet"
        Me.IQMES3DataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'InventoryCountPrepareTableAdapter
        '
        Me.InventoryCountPrepareTableAdapter.ClearBeforeFill = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(515, 79)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "UPDATE"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Form8
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlDark
        Me.ClientSize = New System.Drawing.Size(608, 583)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form8"
        Me.Text = "每日盘点"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.InventoryCountPrepareBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.IQMES3DataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents IQMES3DataSet As ACTION_SUPPORT.IQMES3DataSet
    Friend WithEvents InventoryCountPrepareBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents InventoryCountPrepareTableAdapter As ACTION_SUPPORT.IQMES3DataSetTableAdapters.InventoryCountPrepareTableAdapter
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
