<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ArrowTypeBox
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.ButType1 = New System.Windows.Forms.RadioButton()
        Me.ButType2 = New System.Windows.Forms.RadioButton()
        Me.ButType3 = New System.Windows.Forms.RadioButton()
        Me.ButType4 = New System.Windows.Forms.RadioButton()
        Me.FlowLayoutPanel2 = New System.Windows.Forms.FlowLayoutPanel()
        Me.ButOK = New System.Windows.Forms.Button()
        Me.ButCancel = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.FlowLayoutPanel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 62.95503!))
        Me.TableLayoutPanel1.Controls.Add(Me.GroupBox1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.FlowLayoutPanel2, 0, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 77.32558!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 22.67442!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(275, 172)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.FlowLayoutPanel1)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(269, 126)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Select Arrow Type"
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Controls.Add(Me.ButType1)
        Me.FlowLayoutPanel1.Controls.Add(Me.ButType2)
        Me.FlowLayoutPanel1.Controls.Add(Me.ButType3)
        Me.FlowLayoutPanel1.Controls.Add(Me.ButType4)
        Me.FlowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(3, 16)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(263, 107)
        Me.FlowLayoutPanel1.TabIndex = 0
        '
        'ButType1
        '
        Me.ButType1.AutoSize = True
        Me.ButType1.Location = New System.Drawing.Point(3, 3)
        Me.ButType1.Name = "ButType1"
        Me.ButType1.Size = New System.Drawing.Size(235, 17)
        Me.ButType1.TabIndex = 0
        Me.ButType1.TabStop = True
        Me.ButType1.Text = "1 Line Horizontal, Vertical, or Diagonal Arrow"
        Me.ButType1.UseVisualStyleBackColor = True
        '
        'ButType2
        '
        Me.ButType2.AutoSize = True
        Me.ButType2.Location = New System.Drawing.Point(3, 26)
        Me.ButType2.Name = "ButType2"
        Me.ButType2.Size = New System.Drawing.Size(184, 17)
        Me.ButType2.TabIndex = 1
        Me.ButType2.TabStop = True
        Me.ButType2.Text = "2 (or More) Lines Horizontal Arrow"
        Me.ButType2.UseVisualStyleBackColor = True
        '
        'ButType3
        '
        Me.ButType3.AutoSize = True
        Me.ButType3.Location = New System.Drawing.Point(3, 49)
        Me.ButType3.Name = "ButType3"
        Me.ButType3.Size = New System.Drawing.Size(228, 17)
        Me.ButType3.TabIndex = 2
        Me.ButType3.TabStop = True
        Me.ButType3.Text = "2 (or more) Lines Vertical or Diagonal Arrow"
        Me.ButType3.UseVisualStyleBackColor = True
        '
        'ButType4
        '
        Me.ButType4.AutoSize = True
        Me.ButType4.Location = New System.Drawing.Point(3, 72)
        Me.ButType4.Name = "ButType4"
        Me.ButType4.Size = New System.Drawing.Size(121, 17)
        Me.ButType4.TabIndex = 3
        Me.ButType4.TabStop = True
        Me.ButType4.Text = "Vertical Down Arrow"
        Me.ButType4.UseVisualStyleBackColor = True
        '
        'FlowLayoutPanel2
        '
        Me.FlowLayoutPanel2.Controls.Add(Me.ButOK)
        Me.FlowLayoutPanel2.Controls.Add(Me.ButCancel)
        Me.FlowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel2.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft
        Me.FlowLayoutPanel2.Location = New System.Drawing.Point(3, 135)
        Me.FlowLayoutPanel2.Name = "FlowLayoutPanel2"
        Me.FlowLayoutPanel2.Size = New System.Drawing.Size(269, 34)
        Me.FlowLayoutPanel2.TabIndex = 1
        '
        'ButOK
        '
        Me.ButOK.Location = New System.Drawing.Point(191, 3)
        Me.ButOK.Name = "ButOK"
        Me.ButOK.Size = New System.Drawing.Size(75, 23)
        Me.ButOK.TabIndex = 0
        Me.ButOK.Text = "OK"
        Me.ButOK.UseVisualStyleBackColor = True
        '
        'ButCancel
        '
        Me.ButCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ButCancel.Location = New System.Drawing.Point(110, 3)
        Me.ButCancel.Name = "ButCancel"
        Me.ButCancel.Size = New System.Drawing.Size(75, 23)
        Me.ButCancel.TabIndex = 1
        Me.ButCancel.Text = "Cancel"
        Me.ButCancel.UseVisualStyleBackColor = True
        '
        'ArrowTypeBox
        '
        Me.AcceptButton = Me.ButOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.ButCancel
        Me.ClientSize = New System.Drawing.Size(275, 172)
        Me.ControlBox = False
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "ArrowTypeBox"
        Me.Text = "Arrow Type"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.FlowLayoutPanel1.PerformLayout()
        Me.FlowLayoutPanel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As Windows.Forms.TableLayoutPanel
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents FlowLayoutPanel1 As Windows.Forms.FlowLayoutPanel
    Friend WithEvents ButType1 As Windows.Forms.RadioButton
    Friend WithEvents ButType2 As Windows.Forms.RadioButton
    Friend WithEvents ButType3 As Windows.Forms.RadioButton
    Friend WithEvents ButType4 As Windows.Forms.RadioButton
    Friend WithEvents FlowLayoutPanel2 As Windows.Forms.FlowLayoutPanel
    Friend WithEvents ButOK As Windows.Forms.Button
    Friend WithEvents ButCancel As Windows.Forms.Button
End Class
