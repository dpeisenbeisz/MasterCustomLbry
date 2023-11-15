<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LayoutPicker
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
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.OkButton = New System.Windows.Forms.Button()
        Me.CButton = New System.Windows.Forms.Button()
        Me.PickerLabel = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(12, 54)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListBox1.Size = New System.Drawing.Size(246, 303)
        Me.ListBox1.TabIndex = 0
        '
        'OkButton
        '
        Me.OkButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.OkButton.Location = New System.Drawing.Point(205, 386)
        Me.OkButton.Name = "OkButton"
        Me.OkButton.Size = New System.Drawing.Size(75, 23)
        Me.OkButton.TabIndex = 1
        Me.OkButton.Text = "OK"
        Me.OkButton.UseVisualStyleBackColor = True
        '
        'CButton
        '
        Me.CButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CButton.Location = New System.Drawing.Point(124, 386)
        Me.CButton.Name = "CButton"
        Me.CButton.Size = New System.Drawing.Size(75, 23)
        Me.CButton.TabIndex = 2
        Me.CButton.Text = "Cancel"
        Me.CButton.UseVisualStyleBackColor = True
        '
        'PickerLabel
        '
        Me.PickerLabel.AutoSize = True
        Me.PickerLabel.Location = New System.Drawing.Point(12, 24)
        Me.PickerLabel.Name = "PickerLabel"
        Me.PickerLabel.Size = New System.Drawing.Size(39, 13)
        Me.PickerLabel.TabIndex = 3
        Me.PickerLabel.Text = "Label1"
        '
        'LayoutPicker
        '
        Me.AcceptButton = Me.OkButton
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.CButton
        Me.ClientSize = New System.Drawing.Size(292, 421)
        Me.Controls.Add(Me.PickerLabel)
        Me.Controls.Add(Me.CButton)
        Me.Controls.Add(Me.OkButton)
        Me.Controls.Add(Me.ListBox1)
        Me.Name = "LayoutPicker"
        Me.Text = "Layout Picker"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ListBox1 As Windows.Forms.ListBox
    Friend WithEvents OkButton As Windows.Forms.Button
    Friend WithEvents CButton As Windows.Forms.Button
    Friend WithEvents PickerLabel As Windows.Forms.Label
End Class
