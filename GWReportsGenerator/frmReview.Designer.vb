<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReview
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmReview))
        Me.dgvRecords = New System.Windows.Forms.DataGridView()
        Me.butConfirm = New System.Windows.Forms.Button()
        Me.ckCheckAll = New System.Windows.Forms.CheckBox()
        CType(Me.dgvRecords, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvRecords
        '
        Me.dgvRecords.AllowUserToAddRows = False
        Me.dgvRecords.AllowUserToDeleteRows = False
        Me.dgvRecords.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvRecords.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvRecords.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvRecords.Location = New System.Drawing.Point(12, 12)
        Me.dgvRecords.Name = "dgvRecords"
        Me.dgvRecords.RowTemplate.Height = 24
        Me.dgvRecords.Size = New System.Drawing.Size(858, 469)
        Me.dgvRecords.TabIndex = 0
        '
        'butConfirm
        '
        Me.butConfirm.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.butConfirm.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.butConfirm.ForeColor = System.Drawing.Color.Black
        Me.butConfirm.Location = New System.Drawing.Point(361, 510)
        Me.butConfirm.Name = "butConfirm"
        Me.butConfirm.Size = New System.Drawing.Size(160, 43)
        Me.butConfirm.TabIndex = 3
        Me.butConfirm.Text = "Confirm"
        Me.butConfirm.UseVisualStyleBackColor = True
        '
        'ckCheckAll
        '
        Me.ckCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ckCheckAll.AutoSize = True
        Me.ckCheckAll.Checked = True
        Me.ckCheckAll.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckCheckAll.Location = New System.Drawing.Point(12, 487)
        Me.ckCheckAll.Name = "ckCheckAll"
        Me.ckCheckAll.Size = New System.Drawing.Size(104, 21)
        Me.ckCheckAll.TabIndex = 1
        Me.ckCheckAll.Text = "Uncheck All"
        Me.ckCheckAll.UseVisualStyleBackColor = True
        '
        'frmReview
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(882, 563)
        Me.Controls.Add(Me.ckCheckAll)
        Me.Controls.Add(Me.butConfirm)
        Me.Controls.Add(Me.dgvRecords)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmReview"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Review Captain's List"
        CType(Me.dgvRecords, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvRecords As System.Windows.Forms.DataGridView
    Friend WithEvents butConfirm As System.Windows.Forms.Button
    Friend WithEvents ckCheckAll As System.Windows.Forms.CheckBox
End Class
