''''''''' ''''''''' ''''''''' ''''''''' ''''''''' 
''''''''' Copyright 2007-2009 Brian Saville
''''''''' Command3808 is free for non-commercial use.
''''''''' Reuse of the source code is free for non-commercial use
''''''''' provided that the source code used is credited to Command3808.
''''''''' ''''''''' ''''''''' ''''''''' '''''''''
Public NotInheritable Class AboutBox1

	Private Sub AboutBox1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

		Me.LabelVersion.Text = "v" & My.Application.Info.Version.ToString(4)
		Me.Label1.Text = "DB v" & DB.DBVersion & " (" & DB.DBVersionDate & ")"
		Me.LabelCopyright.Text = My.Application.Info.Copyright
		Me.TextBoxDescription.Text = My.Application.Info.Description

	End Sub

	Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click
		Me.Close()
	End Sub

End Class
