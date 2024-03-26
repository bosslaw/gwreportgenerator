<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCustomize
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
        Me.clbFields = New System.Windows.Forms.CheckedListBox()
        Me.butOK = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'clbFields
        '
        Me.clbFields.CheckOnClick = True
        Me.clbFields.FormattingEnabled = True
        Me.clbFields.Location = New System.Drawing.Point(12, 11)
        Me.clbFields.Name = "clbFields"
        Me.clbFields.Size = New System.Drawing.Size(285, 319)
        Me.clbFields.TabIndex = 0
        '
        'butOK
        '
        Me.butOK.Location = New System.Drawing.Point(112, 338)
        Me.butOK.Name = "butOK"
        Me.butOK.Size = New System.Drawing.Size(85, 34)
        Me.butOK.TabIndex = 1
        Me.butOK.Text = "OK"
        Me.butOK.UseVisualStyleBackColor = True
        '
        'frmCustomize
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(309, 383)
        Me.Controls.Add(Me.butOK)
        Me.Controls.Add(Me.clbFields)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.Name = "frmCustomize"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Customize Displayed Fields"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents clbFields As System.Windows.Forms.CheckedListBox
    Friend WithEvents butOK As System.Windows.Forms.Button
End Class
