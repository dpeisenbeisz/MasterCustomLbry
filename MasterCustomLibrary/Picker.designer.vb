<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Picker
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
        Me.TopLabel = New System.Windows.Forms.Label()
        Me.ButCancel = New System.Windows.Forms.Button()
        Me.ButOK = New System.Windows.Forms.Button()
        Me.BxList = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'TopLabel
        '
        Me.TopLabel.AutoSize = True
        Me.TopLabel.Location = New System.Drawing.Point(13, 13)
        Me.TopLabel.Name = "TopLabel"
        Me.TopLabel.Size = New System.Drawing.Size(103, 13)
        Me.TopLabel.TabIndex = 0
        Me.TopLabel.Text = "&Pick Layer to Delete"
        '
        'ButCancel
        '
        Me.ButCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ButCancel.Location = New System.Drawing.Point(175, 539)
        Me.ButCancel.Name = "ButCancel"
        Me.ButCancel.Size = New System.Drawing.Size(99, 22)
        Me.ButCancel.TabIndex = 15
        Me.ButCancel.Text = "&Cancel"
        Me.ButCancel.UseVisualStyleBackColor = True
        '
        'ButOK
        '
        Me.ButOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.ButOK.Location = New System.Drawing.Point(88, 539)
        Me.ButOK.Name = "ButOK"
        Me.ButOK.Size = New System.Drawing.Size(81, 23)
        Me.ButOK.TabIndex = 14
        Me.ButOK.Text = "&OK"
        Me.ButOK.UseVisualStyleBackColor = True
        '
        'BxList
        '
        Me.BxList.FormattingEnabled = True
        Me.BxList.Location = New System.Drawing.Point(16, 46)
        Me.BxList.Name = "BxList"
        Me.BxList.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.BxList.Size = New System.Drawing.Size(245, 459)
        Me.BxList.TabIndex = 16
        '
        'Picker
        '
        Me.AcceptButton = Me.ButOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.CancelButton = Me.ButCancel
        Me.ClientSize = New System.Drawing.Size(286, 574)
        Me.Controls.Add(Me.BxList)
        Me.Controls.Add(Me.ButCancel)
        Me.Controls.Add(Me.ButOK)
        Me.Controls.Add(Me.TopLabel)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Picker"
        Me.Text = "Select Properties"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TopLabel As Windows.Forms.Label
    Friend WithEvents ButCancel As Windows.Forms.Button
    Friend WithEvents ButOK As Windows.Forms.Button
    Friend WithEvents BxList As Windows.Forms.ListBox
End Class
