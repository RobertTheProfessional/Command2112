''''''''' ''''''''' ''''''''' ''''''''' ''''''''' 
''''''''' Copyright 2007-2009 Brian Saville
''''''''' Command3808 is free for non-commercial use.
''''''''' Reuse of the source code is free for non-commercial use
''''''''' provided that the source code used is credited to Command3808.
''''''''' ''''''''' ''''''''' ''''''''' '''''''''
Public NotInheritable Class SplashScreen1

	Public Sub SetStatus(ByVal Status As String)
		Me.Invoke(New dSetStatus(AddressOf iSetStatus), Status)
	End Sub

	Private Delegate Sub dSetStatus(ByVal Status As String)

	Private Sub iSetStatus(ByVal Status As String)
		Me.TextBox1.Text = Status
	End Sub

	Private Sub SplashScreen1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Version.Text = "v" & My.Application.Info.Version.ToString(4)

		Copyright.Text = My.Application.Info.Copyright

	End Sub

	Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
		MyBase.OnPaint(e)
		Dim p As New Pen(System.Drawing.SystemColors.ControlText, 1)
		Dim g As Graphics

		g = Me.CreateGraphics
		g.DrawLine(p, PictureBox1.Left, PictureBox1.Bottom + 1, PictureBox1.Right, PictureBox1.Bottom + 1)
		p.Dispose()
		g.Dispose()
	End Sub

End Class
