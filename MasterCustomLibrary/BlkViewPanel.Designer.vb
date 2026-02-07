<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BlkViewPanel
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.DGV1 = New System.Windows.Forms.DataGridView()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.ButOk = New System.Windows.Forms.Button()
        Me.RefHandle = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RefName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Insertion = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Rotation = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ReferenceId = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.HasAttributes = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IsDynamic = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.RefCount = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DynamicName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.DGV1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 85.3531!))
        Me.TableLayoutPanel1.Controls.Add(Me.DGV1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.FlowLayoutPanel1, 0, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 91.11111!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 8.888889!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1294, 450)
        Me.TableLayoutPanel1.TabIndex = 3
        '
        'DGV1
        '
        Me.DGV1.AllowUserToAddRows = False
        Me.DGV1.AllowUserToDeleteRows = False
        Me.DGV1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGV1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.RefHandle, Me.RefName, Me.Insertion, Me.Rotation, Me.ReferenceId, Me.HasAttributes, Me.IsDynamic, Me.RefCount, Me.DynamicName})
        Me.DGV1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGV1.Location = New System.Drawing.Point(3, 3)
        Me.DGV1.Name = "DGV1"
        Me.DGV1.ReadOnly = True
        Me.DGV1.Size = New System.Drawing.Size(1288, 404)
        Me.DGV1.TabIndex = 1
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Controls.Add(Me.ButOk)
        Me.FlowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(3, 413)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(1288, 34)
        Me.FlowLayoutPanel1.TabIndex = 5
        '
        'ButOk
        '
        Me.ButOk.Location = New System.Drawing.Point(1210, 3)
        Me.ButOk.Name = "ButOk"
        Me.ButOk.Size = New System.Drawing.Size(75, 23)
        Me.ButOk.TabIndex = 4
        Me.ButOk.Text = "Close"
        Me.ButOk.UseVisualStyleBackColor = True
        '
        'RefHandle
        '
        Me.RefHandle.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.RefHandle.HeaderText = "RefHandle"
        Me.RefHandle.Name = "RefHandle"
        Me.RefHandle.ReadOnly = True
        Me.RefHandle.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.RefHandle.Width = 83
        '
        'RefName
        '
        Me.RefName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.RefName.HeaderText = "RefName"
        Me.RefName.Name = "RefName"
        Me.RefName.ReadOnly = True
        Me.RefName.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.RefName.Width = 77
        '
        'Insertion
        '
        Me.Insertion.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Insertion.HeaderText = "Insertion"
        Me.Insertion.Name = "Insertion"
        Me.Insertion.ReadOnly = True
        Me.Insertion.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Insertion.Width = 72
        '
        'Rotation
        '
        Me.Rotation.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Rotation.HeaderText = "Rotation"
        Me.Rotation.Name = "Rotation"
        Me.Rotation.ReadOnly = True
        Me.Rotation.Width = 72
        '
        'ReferenceId
        '
        Me.ReferenceId.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ReferenceId.HeaderText = "ReferenceId"
        Me.ReferenceId.Name = "ReferenceId"
        Me.ReferenceId.ReadOnly = True
        Me.ReferenceId.Width = 91
        '
        'HasAttributes
        '
        Me.HasAttributes.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.HasAttributes.HeaderText = "HasAttributes"
        Me.HasAttributes.Name = "HasAttributes"
        Me.HasAttributes.ReadOnly = True
        Me.HasAttributes.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.HasAttributes.Width = 95
        '
        'IsDynamic
        '
        Me.IsDynamic.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.IsDynamic.FalseValue = "false"
        Me.IsDynamic.HeaderText = "IsDynamic"
        Me.IsDynamic.Name = "IsDynamic"
        Me.IsDynamic.ReadOnly = True
        Me.IsDynamic.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.IsDynamic.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.IsDynamic.TrueValue = "true"
        Me.IsDynamic.Width = 81
        '
        'RefCount
        '
        Me.RefCount.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.RefCount.HeaderText = "RefCount"
        Me.RefCount.Name = "RefCount"
        Me.RefCount.ReadOnly = True
        Me.RefCount.Width = 77
        '
        'DynamicName
        '
        Me.DynamicName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.DynamicName.HeaderText = "DynamicName"
        Me.DynamicName.Name = "DynamicName"
        Me.DynamicName.ReadOnly = True
        Me.DynamicName.Width = 101
        '
        'BlkViewPanel
        '
        Me.AcceptButton = Me.ButOk
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(1294, 450)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "BlkViewPanel"
        Me.Text = "Block Data"
        Me.TableLayoutPanel1.ResumeLayout(False)
        CType(Me.DGV1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As Windows.Forms.TableLayoutPanel
    Friend WithEvents DGV1 As Windows.Forms.DataGridView
    Friend WithEvents FlowLayoutPanel1 As Windows.Forms.FlowLayoutPanel
    Friend WithEvents ButOk As Windows.Forms.Button
    Friend WithEvents RefHandle As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RefName As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Insertion As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Rotation As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ReferenceId As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents HasAttributes As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IsDynamic As Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents RefCount As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DynamicName As Windows.Forms.DataGridViewTextBoxColumn
End Class
