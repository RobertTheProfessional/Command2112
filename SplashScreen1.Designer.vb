<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SplashScreen1
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
		Me.PictureBox1 = New System.Windows.Forms.PictureBox
		Me.ApplicationTitle = New System.Windows.Forms.Label
		Me.TextBox1 = New System.Windows.Forms.TextBox
		Me.Label1 = New System.Windows.Forms.Label
		Me.Copyright = New System.Windows.Forms.Label
		Me.Version = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'PictureBox1
		'
		Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Top
		Me.PictureBox1.ErrorImage = Nothing
        Me.PictureBox1.Image = Global.Command2112.My.Resources.Resources.AVR2112CI
        Me.PictureBox1.InitialImage = Nothing
        Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(698, 302)
        Me.PictureBox1.TabIndex = 1
        Me.PictureBox1.TabStop = False
        '
        'ApplicationTitle
        '
        Me.ApplicationTitle.AutoSize = True
        Me.ApplicationTitle.BackColor = System.Drawing.Color.Transparent
        Me.ApplicationTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ApplicationTitle.Location = New System.Drawing.Point(512, 305)
        Me.ApplicationTitle.Name = "ApplicationTitle"
        Me.ApplicationTitle.Size = New System.Drawing.Size(176, 29)
        Me.ApplicationTitle.TabIndex = 2
        Me.ApplicationTitle.Text = "Command2112"
		Me.ApplicationTitle.TextAlign = System.Drawing.ContentAlignment.MiddleRight
		'
		'TextBox1
		'
		Me.TextBox1.Anchor = System.Windows.Forms.AnchorStyles.Left
		Me.TextBox1.BackColor = System.Drawing.SystemColors.Control
		Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.TextBox1.Cursor = System.Windows.Forms.Cursors.Default
		Me.TextBox1.Location = New System.Drawing.Point(12, 373)
		Me.TextBox1.Name = "TextBox1"
		Me.TextBox1.ReadOnly = True
		Me.TextBox1.Size = New System.Drawing.Size(496, 13)
		Me.TextBox1.TabIndex = 2
		Me.TextBox1.TabStop = False
		Me.TextBox1.WordWrap = False
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Location = New System.Drawing.Point(12, 324)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(285, 13)
		Me.Label1.TabIndex = 3
        Me.Label1.Text = "Above Denon AVR-2112CI Photo: Copyright D&&M Holdings, Inc."
		'
		'Copyright
		'
		Me.Copyright.Anchor = System.Windows.Forms.AnchorStyles.None
		Me.Copyright.AutoSize = True
		Me.Copyright.BackColor = System.Drawing.Color.Transparent
		Me.Copyright.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Copyright.Location = New System.Drawing.Point(12, 305)
		Me.Copyright.Name = "Copyright"
		Me.Copyright.Size = New System.Drawing.Size(58, 15)
		Me.Copyright.TabIndex = 2
		Me.Copyright.Text = "Copyright"
		'
		'Version
		'
		Me.Version.Anchor = System.Windows.Forms.AnchorStyles.None
		Me.Version.AutoSize = True
		Me.Version.BackColor = System.Drawing.Color.Transparent
		Me.Version.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Version.Location = New System.Drawing.Point(514, 334)
		Me.Version.Name = "Version"
		Me.Version.Size = New System.Drawing.Size(48, 15)
		Me.Version.TabIndex = 1
		Me.Version.Text = "Version"
		'
		'Label2
		'
		Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Location = New System.Drawing.Point(12, 341)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(418, 32)
		Me.Label2.TabIndex = 4
		Me.Label2.Text = "This program is in no way associated with, produced by, or other otherwise" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "endor" & _
			"sed by D&&M Holdings, Inc."
		'
		'SplashScreen1
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(698, 398)
		Me.ControlBox = False
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.Version)
		Me.Controls.Add(Me.Copyright)
		Me.Controls.Add(Me.ApplicationTitle)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.TextBox1)
		Me.Controls.Add(Me.PictureBox1)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.Name = "SplashScreen1"
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub
	Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
	Friend WithEvents ApplicationTitle As System.Windows.Forms.Label
	Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
	Friend WithEvents Version As System.Windows.Forms.Label
	Friend WithEvents Copyright As System.Windows.Forms.Label
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents Label2 As System.Windows.Forms.Label

End Class
