<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.btnDeleteAll = New System.Windows.Forms.Button()
        Me.btnDeleteOne = New System.Windows.Forms.Button()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.lblStatus = New System.Windows.Forms.Label()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.lblStatus)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnImport)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnDeleteAll)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnDeleteOne)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnRefresh)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.DataGridView1)
        Me.SplitContainer1.Size = New System.Drawing.Size(443, 261)
        Me.SplitContainer1.SplitterDistance = 94
        Me.SplitContainer1.TabIndex = 0
        '
        'btnImport
        '
        Me.btnImport.Location = New System.Drawing.Point(13, 102)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(75, 23)
        Me.btnImport.TabIndex = 3
        Me.btnImport.Text = "Import"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'btnDeleteAll
        '
        Me.btnDeleteAll.Location = New System.Drawing.Point(13, 72)
        Me.btnDeleteAll.Name = "btnDeleteAll"
        Me.btnDeleteAll.Size = New System.Drawing.Size(75, 23)
        Me.btnDeleteAll.TabIndex = 2
        Me.btnDeleteAll.Text = "Delete All"
        Me.btnDeleteAll.UseVisualStyleBackColor = True
        '
        'btnDeleteOne
        '
        Me.btnDeleteOne.Location = New System.Drawing.Point(13, 42)
        Me.btnDeleteOne.Name = "btnDeleteOne"
        Me.btnDeleteOne.Size = New System.Drawing.Size(75, 23)
        Me.btnDeleteOne.TabIndex = 1
        Me.btnDeleteOne.Text = "Delete"
        Me.btnDeleteOne.UseVisualStyleBackColor = True
        '
        'btnRefresh
        '
        Me.btnRefresh.Location = New System.Drawing.Point(12, 12)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(75, 23)
        Me.btnRefresh.TabIndex = 0
        Me.btnRefresh.Text = "Refresh"
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(345, 261)
        Me.DataGridView1.TabIndex = 0
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Filter = "XLS|*.xls;*.xlsx|CSV|*.csv"
        '
        'lblStatus
        '
        Me.lblStatus.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(13, 236)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(0, 13)
        Me.lblStatus.TabIndex = 4
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(443, 261)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "Form1"
        Me.Text = "ZenDirect Contacts"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents btnImport As System.Windows.Forms.Button
    Friend WithEvents btnDeleteAll As System.Windows.Forms.Button
    Friend WithEvents btnDeleteOne As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents lblStatus As System.Windows.Forms.Label

End Class
