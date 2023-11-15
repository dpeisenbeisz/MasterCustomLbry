<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FontPicker
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
        Me.ButCancel = New System.Windows.Forms.Button()
        Me.ButOK = New System.Windows.Forms.Button()
        Me.LabPickerType = New System.Windows.Forms.Label()
        Me.BxFonts = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'ButCancel
        '
        Me.ButCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ButCancel.Location = New System.Drawing.Point(282, 98)
        Me.ButCancel.Name = "ButCancel"
        Me.ButCancel.Size = New System.Drawing.Size(75, 23)
        Me.ButCancel.TabIndex = 3
        Me.ButCancel.Text = "Cancel"
        Me.ButCancel.UseVisualStyleBackColor = True
        '
        'ButOK
        '
        Me.ButOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.ButOK.Location = New System.Drawing.Point(185, 98)
        Me.ButOK.Name = "ButOK"
        Me.ButOK.Size = New System.Drawing.Size(75, 23)
        Me.ButOK.TabIndex = 2
        Me.ButOK.Text = "OK"
        Me.ButOK.UseVisualStyleBackColor = True
        '
        'LabPickerType
        '
        Me.LabPickerType.AutoSize = True
        Me.LabPickerType.Location = New System.Drawing.Point(27, 20)
        Me.LabPickerType.Name = "LabPickerType"
        Me.LabPickerType.Size = New System.Drawing.Size(87, 13)
        Me.LabPickerType.TabIndex = 4
        Me.LabPickerType.Text = "PickerTypeLabel"
        Me.LabPickerType.UseMnemonic = False
        '
        'BxFonts
        '
        Me.BxFonts.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.BxFonts.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.BxFonts.FormattingEnabled = True
        Me.BxFonts.Location = New System.Drawing.Point(30, 45)
        Me.BxFonts.Name = "BxFonts"
        Me.BxFonts.Size = New System.Drawing.Size(309, 21)
        Me.BxFonts.TabIndex = 1
        '
        'FontPicker
        '
        Me.AcceptButton = Me.ButOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.CancelButton = Me.ButCancel
        Me.ClientSize = New System.Drawing.Size(379, 133)
        Me.Controls.Add(Me.BxFonts)
        Me.Controls.Add(Me.LabPickerType)
        Me.Controls.Add(Me.ButOK)
        Me.Controls.Add(Me.ButCancel)
        Me.Name = "FontPicker"
        Me.Text = "Form3"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ButCancel As Windows.Forms.Button
    Friend WithEvents ButOK As Windows.Forms.Button
    Friend WithEvents LabPickerType As Windows.Forms.Label
    Friend WithEvents BxFonts As Windows.Forms.ComboBox
End Class
