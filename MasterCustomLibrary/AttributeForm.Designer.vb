<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AttributeForm
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
        Me.DGV = New System.Windows.Forms.DataGridView()
        Me.Attribute = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Value = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ButClose = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.DGV, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.DGV, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.ButClose, 0, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 92.87128!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 7.128713!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(500, 505)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'DGV
        '
        Me.DGV.AllowUserToAddRows = False
        Me.DGV.AllowUserToDeleteRows = False
        Me.DGV.AllowUserToResizeRows = False
        Me.DGV.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGV.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Attribute, Me.Value})
        Me.DGV.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGV.Location = New System.Drawing.Point(3, 3)
        Me.DGV.Name = "DGV"
        Me.DGV.ReadOnly = True
        Me.DGV.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.DGV.ShowCellToolTips = False
        Me.DGV.ShowEditingIcon = False
        Me.DGV.Size = New System.Drawing.Size(494, 462)
        Me.DGV.TabIndex = 0
        '
        'Attribute
        '
        Me.Attribute.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Attribute.HeaderText = "Attribute"
        Me.Attribute.MinimumWidth = 150
        Me.Attribute.Name = "Attribute"
        Me.Attribute.ReadOnly = True
        Me.Attribute.Width = 150
        '
        'Value
        '
        Me.Value.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Value.HeaderText = "Value"
        Me.Value.MinimumWidth = 300
        Me.Value.Name = "Value"
        Me.Value.ReadOnly = True
        Me.Value.Width = 300
        '
        'ButClose
        '
        Me.ButClose.Dock = System.Windows.Forms.DockStyle.Right
        Me.ButClose.Location = New System.Drawing.Point(393, 471)
        Me.ButClose.Name = "ButClose"
        Me.ButClose.Size = New System.Drawing.Size(104, 31)
        Me.ButClose.TabIndex = 1
        Me.ButClose.Text = "Close"
        Me.ButClose.UseVisualStyleBackColor = True
        '
        'AttributeForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(500, 505)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "AttributeForm"
        Me.Text = "Block Attributes"
        Me.TableLayoutPanel1.ResumeLayout(False)
        CType(Me.DGV, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As Windows.Forms.TableLayoutPanel
    Friend WithEvents DGV As Windows.Forms.DataGridView
    Friend WithEvents Attribute As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Value As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ButClose As Windows.Forms.Button
End Class
