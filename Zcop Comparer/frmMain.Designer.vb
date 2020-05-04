<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.grpChooseFile = New System.Windows.Forms.GroupBox()
        Me.btnBrowseMappingFile = New System.Windows.Forms.Button()
        Me.txtMappingFile = New System.Windows.Forms.TextBox()
        Me.lblMappingFile = New System.Windows.Forms.Label()
        Me.btnBrowseNewFile = New System.Windows.Forms.Button()
        Me.txtNewFile = New System.Windows.Forms.TextBox()
        Me.lblNewFile = New System.Windows.Forms.Label()
        Me.btnBrowseOldFile = New System.Windows.Forms.Button()
        Me.txtOldFile = New System.Windows.Forms.TextBox()
        Me.lblOldFile = New System.Windows.Forms.Label()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.lstProcess = New System.Windows.Forms.ListBox()
        Me.opnOldFile = New System.Windows.Forms.OpenFileDialog()
        Me.opnNewFile = New System.Windows.Forms.OpenFileDialog()
        Me.lblSubProgress = New System.Windows.Forms.Label()
        Me.opnMappingFile = New System.Windows.Forms.OpenFileDialog()
        Me.grpChooseFile.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.BackColor = System.Drawing.Color.SkyBlue
        Me.Panel3.Location = New System.Drawing.Point(725, 1)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(13, 423)
        Me.Panel3.TabIndex = 14
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.Color.SkyBlue
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(13, 423)
        Me.Panel1.TabIndex = 15
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BackColor = System.Drawing.Color.SkyBlue
        Me.Panel2.Location = New System.Drawing.Point(1, 0)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(737, 14)
        Me.Panel2.TabIndex = 16
        '
        'Panel4
        '
        Me.Panel4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel4.BackColor = System.Drawing.Color.SkyBlue
        Me.Panel4.Location = New System.Drawing.Point(1, 410)
        Me.Panel4.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(737, 14)
        Me.Panel4.TabIndex = 17
        '
        'grpChooseFile
        '
        Me.grpChooseFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpChooseFile.Controls.Add(Me.btnBrowseMappingFile)
        Me.grpChooseFile.Controls.Add(Me.txtMappingFile)
        Me.grpChooseFile.Controls.Add(Me.lblMappingFile)
        Me.grpChooseFile.Controls.Add(Me.btnBrowseNewFile)
        Me.grpChooseFile.Controls.Add(Me.txtNewFile)
        Me.grpChooseFile.Controls.Add(Me.lblNewFile)
        Me.grpChooseFile.Controls.Add(Me.btnBrowseOldFile)
        Me.grpChooseFile.Controls.Add(Me.txtOldFile)
        Me.grpChooseFile.Controls.Add(Me.lblOldFile)
        Me.grpChooseFile.Location = New System.Drawing.Point(21, 22)
        Me.grpChooseFile.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.grpChooseFile.Name = "grpChooseFile"
        Me.grpChooseFile.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.grpChooseFile.Size = New System.Drawing.Size(697, 133)
        Me.grpChooseFile.TabIndex = 18
        Me.grpChooseFile.TabStop = False
        Me.grpChooseFile.Text = "Choose Files"
        '
        'btnBrowseMappingFile
        '
        Me.btnBrowseMappingFile.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.btnBrowseMappingFile.Location = New System.Drawing.Point(619, 96)
        Me.btnBrowseMappingFile.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnBrowseMappingFile.Name = "btnBrowseMappingFile"
        Me.btnBrowseMappingFile.Size = New System.Drawing.Size(75, 28)
        Me.btnBrowseMappingFile.TabIndex = 9
        Me.btnBrowseMappingFile.Text = "Browse"
        Me.btnBrowseMappingFile.UseVisualStyleBackColor = True
        '
        'txtMappingFile
        '
        Me.txtMappingFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMappingFile.Location = New System.Drawing.Point(138, 97)
        Me.txtMappingFile.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtMappingFile.Name = "txtMappingFile"
        Me.txtMappingFile.Size = New System.Drawing.Size(473, 22)
        Me.txtMappingFile.TabIndex = 8
        '
        'lblMappingFile
        '
        Me.lblMappingFile.AutoSize = True
        Me.lblMappingFile.Location = New System.Drawing.Point(7, 101)
        Me.lblMappingFile.Name = "lblMappingFile"
        Me.lblMappingFile.Size = New System.Drawing.Size(125, 17)
        Me.lblMappingFile.TabIndex = 7
        Me.lblMappingFile.Text = "Mapping File Path:"
        '
        'btnBrowseNewFile
        '
        Me.btnBrowseNewFile.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.btnBrowseNewFile.Location = New System.Drawing.Point(619, 59)
        Me.btnBrowseNewFile.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnBrowseNewFile.Name = "btnBrowseNewFile"
        Me.btnBrowseNewFile.Size = New System.Drawing.Size(75, 28)
        Me.btnBrowseNewFile.TabIndex = 6
        Me.btnBrowseNewFile.Text = "Browse"
        Me.btnBrowseNewFile.UseVisualStyleBackColor = True
        '
        'txtNewFile
        '
        Me.txtNewFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNewFile.Location = New System.Drawing.Point(138, 60)
        Me.txtNewFile.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtNewFile.Name = "txtNewFile"
        Me.txtNewFile.Size = New System.Drawing.Size(473, 22)
        Me.txtNewFile.TabIndex = 5
        '
        'lblNewFile
        '
        Me.lblNewFile.AutoSize = True
        Me.lblNewFile.Location = New System.Drawing.Point(7, 64)
        Me.lblNewFile.Name = "lblNewFile"
        Me.lblNewFile.Size = New System.Drawing.Size(98, 17)
        Me.lblNewFile.TabIndex = 4
        Me.lblNewFile.Text = "New File Path:"
        '
        'btnBrowseOldFile
        '
        Me.btnBrowseOldFile.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.btnBrowseOldFile.Location = New System.Drawing.Point(619, 23)
        Me.btnBrowseOldFile.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnBrowseOldFile.Name = "btnBrowseOldFile"
        Me.btnBrowseOldFile.Size = New System.Drawing.Size(75, 28)
        Me.btnBrowseOldFile.TabIndex = 3
        Me.btnBrowseOldFile.Text = "Browse"
        Me.btnBrowseOldFile.UseVisualStyleBackColor = True
        '
        'txtOldFile
        '
        Me.txtOldFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOldFile.Location = New System.Drawing.Point(138, 23)
        Me.txtOldFile.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtOldFile.Name = "txtOldFile"
        Me.txtOldFile.Size = New System.Drawing.Size(473, 22)
        Me.txtOldFile.TabIndex = 1
        '
        'lblOldFile
        '
        Me.lblOldFile.AutoSize = True
        Me.lblOldFile.Location = New System.Drawing.Point(7, 27)
        Me.lblOldFile.Name = "lblOldFile"
        Me.lblOldFile.Size = New System.Drawing.Size(93, 17)
        Me.lblOldFile.TabIndex = 0
        Me.lblOldFile.Text = "Old File Path:"
        '
        'btnStop
        '
        Me.btnStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnStop.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStop.Location = New System.Drawing.Point(619, 162)
        Me.btnStop.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(100, 34)
        Me.btnStop.TabIndex = 19
        Me.btnStop.Text = "Stop"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStart.Location = New System.Drawing.Point(509, 162)
        Me.btnStart.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(100, 34)
        Me.btnStart.TabIndex = 20
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'lstProcess
        '
        Me.lstProcess.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstProcess.ForeColor = System.Drawing.Color.FromArgb(CType(CType(29, Byte), Integer), CType(CType(29, Byte), Integer), CType(CType(29, Byte), Integer))
        Me.lstProcess.FormattingEnabled = True
        Me.lstProcess.ItemHeight = 16
        Me.lstProcess.Location = New System.Drawing.Point(19, 209)
        Me.lstProcess.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.lstProcess.Name = "lstProcess"
        Me.lstProcess.Size = New System.Drawing.Size(699, 196)
        Me.lstProcess.TabIndex = 29
        '
        'opnOldFile
        '
        '
        'opnNewFile
        '
        '
        'lblSubProgress
        '
        Me.lblSubProgress.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSubProgress.Location = New System.Drawing.Point(21, 158)
        Me.lblSubProgress.Name = "lblSubProgress"
        Me.lblSubProgress.Size = New System.Drawing.Size(445, 46)
        Me.lblSubProgress.TabIndex = 30
        Me.lblSubProgress.Text = "Sub Progress"
        '
        'opnMappingFile
        '
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(739, 423)
        Me.Controls.Add(Me.lblSubProgress)
        Me.Controls.Add(Me.lstProcess)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.btnStop)
        Me.Controls.Add(Me.grpChooseFile)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel3)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Zcop Comparer"
        Me.grpChooseFile.ResumeLayout(False)
        Me.grpChooseFile.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel3 As Panel
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel4 As Panel
    Friend WithEvents grpChooseFile As GroupBox
    Friend WithEvents lblOldFile As Label
    Friend WithEvents btnBrowseOldFile As Button
    Friend WithEvents txtOldFile As TextBox
    Friend WithEvents btnBrowseNewFile As Button
    Friend WithEvents txtNewFile As TextBox
    Friend WithEvents lblNewFile As Label
    Friend WithEvents btnStop As Button
    Friend WithEvents btnStart As Button
    Friend WithEvents lstProcess As ListBox
    Friend WithEvents opnOldFile As OpenFileDialog
    Friend WithEvents opnNewFile As OpenFileDialog
    Friend WithEvents lblSubProgress As Label
    Friend WithEvents btnBrowseMappingFile As Button
    Friend WithEvents txtMappingFile As TextBox
    Friend WithEvents lblMappingFile As Label
    Friend WithEvents opnMappingFile As OpenFileDialog
End Class
