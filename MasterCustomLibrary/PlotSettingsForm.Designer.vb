<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PlotSettingsForm
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
        Me.BoxPageSetups = New System.Windows.Forms.ComboBox()
        Me.OkButton = New System.Windows.Forms.Button()
        Me.CButton = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'BoxPageSetups
        '
        Me.BoxPageSetups.FormattingEnabled = True
        Me.BoxPageSetups.Location = New System.Drawing.Point(32, 58)
        Me.BoxPageSetups.Name = "BoxPageSetups"
        Me.BoxPageSetups.Size = New System.Drawing.Size(245, 21)
        Me.BoxPageSetups.TabIndex = 0
        '
        'OkButton
        '
        Me.OkButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.OkButton.Location = New System.Drawing.Point(237, 162)
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
        Me.CButton.Location = New System.Drawing.Point(156, 162)
        Me.CButton.Name = "CButton"
        Me.CButton.Size = New System.Drawing.Size(75, 23)
        Me.CButton.TabIndex = 2
        Me.CButton.Text = "Cancel"
        Me.CButton.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(29, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(102, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Select a page setup"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(29, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(177, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "CLick Cancel to create new settings"
        '
        'PlotSettingsForm
        '
        Me.AcceptButton = Me.OkButton
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoSize = True
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.CButton
        Me.ClientSize = New System.Drawing.Size(318, 197)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CButton)
        Me.Controls.Add(Me.OkButton)
        Me.Controls.Add(Me.BoxPageSetups)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "PlotSettingsForm"
        Me.Text = "Plot Settings"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents BoxPageSetups As Windows.Forms.ComboBox
    Friend WithEvents OkButton As Windows.Forms.Button
    Friend WithEvents CButton As Windows.Forms.Button
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
End Class
