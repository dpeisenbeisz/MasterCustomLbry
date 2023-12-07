<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BlockViewPanel
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
        Me.ButCancel = New System.Windows.Forms.Button()
        Me.ButOk = New System.Windows.Forms.Button()
        Me.RefName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RefID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BlkCount = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BlockName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BTRid = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.InsPt = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Rotation = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.HasAttr = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IsDynamic = New System.Windows.Forms.DataGridViewTextBoxColumn()
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
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1349, 450)
        Me.TableLayoutPanel1.TabIndex = 3
        '
        'DGV1
        '
        Me.DGV1.AllowUserToAddRows = False
        Me.DGV1.AllowUserToDeleteRows = False
        Me.DGV1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.RefName, Me.RefID, Me.BlkCount, Me.BlockName, Me.BTRid, Me.InsPt, Me.Rotation, Me.HasAttr, Me.IsDynamic})
        Me.DGV1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGV1.Location = New System.Drawing.Point(3, 3)
        Me.DGV1.Name = "DGV1"
        Me.DGV1.ReadOnly = True
        Me.DGV1.Size = New System.Drawing.Size(1343, 403)
        Me.DGV1.TabIndex = 1
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Controls.Add(Me.ButCancel)
        Me.FlowLayoutPanel1.Controls.Add(Me.ButOk)
        Me.FlowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(3, 412)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(1343, 35)
        Me.FlowLayoutPanel1.TabIndex = 5
        '
        'ButCancel
        '
        Me.ButCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ButCancel.Location = New System.Drawing.Point(1265, 3)
        Me.ButCancel.Name = "ButCancel"
        Me.ButCancel.Size = New System.Drawing.Size(75, 23)
        Me.ButCancel.TabIndex = 3
        Me.ButCancel.Text = "Cancel"
        Me.ButCancel.UseVisualStyleBackColor = True
        '
        'ButOk
        '
        Me.ButOk.Location = New System.Drawing.Point(1184, 3)
        Me.ButOk.Name = "ButOk"
        Me.ButOk.Size = New System.Drawing.Size(75, 23)
        Me.ButOk.TabIndex = 4
        Me.ButOk.Text = "OK"
        Me.ButOk.UseVisualStyleBackColor = True
        '
        'RefName
        '
        Me.RefName.HeaderText = "Blk Ref Name"
        Me.RefName.Name = "RefName"
        Me.RefName.ReadOnly = True
        Me.RefName.Width = 150
        '
        'RefID
        '
        Me.RefID.HeaderText = "Block Reference Id"
        Me.RefID.Name = "RefID"
        Me.RefID.ReadOnly = True
        Me.RefID.Width = 150
        '
        'BlkCount
        '
        Me.BlkCount.HeaderText = "Block Count"
        Me.BlkCount.Name = "BlkCount"
        Me.BlkCount.ReadOnly = True
        '
        'BlockName
        '
        Me.BlockName.HeaderText = "Block Name"
        Me.BlockName.Name = "BlockName"
        Me.BlockName.ReadOnly = True
        Me.BlockName.Width = 150
        '
        'BTRid
        '
        Me.BTRid.HeaderText = "Block Table Rec ID"
        Me.BTRid.Name = "BTRid"
        Me.BTRid.ReadOnly = True
        Me.BTRid.Width = 150
        '
        'InsPt
        '
        Me.InsPt.HeaderText = "Block Insertion Pt"
        Me.InsPt.Name = "InsPt"
        Me.InsPt.ReadOnly = True
        Me.InsPt.Width = 250
        '
        'Rotation
        '
        Me.Rotation.HeaderText = "Rotation"
        Me.Rotation.Name = "Rotation"
        Me.Rotation.ReadOnly = True
        Me.Rotation.Width = 150
        '
        'HasAttr
        '
        Me.HasAttr.HeaderText = "Has Attributes"
        Me.HasAttr.Name = "HasAttr"
        Me.HasAttr.ReadOnly = True
        '
        'IsDynamic
        '
        Me.IsDynamic.HeaderText = "Is Dynamic"
        Me.IsDynamic.Name = "IsDynamic"
        Me.IsDynamic.ReadOnly = True
        '
        'BlockViewPanel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1349, 450)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "BlockViewPanel"
        Me.Text = "Block Data"
        Me.TableLayoutPanel1.ResumeLayout(False)
        CType(Me.DGV1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As Windows.Forms.TableLayoutPanel
    Friend WithEvents DGV1 As Windows.Forms.DataGridView
    Friend WithEvents FlowLayoutPanel1 As Windows.Forms.FlowLayoutPanel
    Friend WithEvents ButCancel As Windows.Forms.Button
    Friend WithEvents ButOk As Windows.Forms.Button
    Friend WithEvents RefName As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RefID As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BlkCount As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BlockName As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BTRid As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents InsPt As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Rotation As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents HasAttr As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IsDynamic As Windows.Forms.DataGridViewTextBoxColumn
End Class
