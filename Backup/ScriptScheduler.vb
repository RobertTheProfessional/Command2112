''''''''' ''''''''' ''''''''' ''''''''' ''''''''' 
''''''''' Copyright 2007-2009 Brian Saville
''''''''' Command3808 is free for non-commercial use.
''''''''' Reuse of the source code is free for non-commercial use
''''''''' provided that the source code used is credited to Command3808.
''''''''' ''''''''' ''''''''' ''''''''' '''''''''
Imports System.Windows.Forms

Public Class ScriptScheduler

	Private Sub ScriptScheduler_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Dim m_s As Schedule

		DateTimePicker1.CustomFormat = Helper.CurrentCulture.DateTimeFormat.ShortTimePattern
		DateTimePicker2.CustomFormat = DateTimePicker1.CustomFormat
		DateTimePicker3.CustomFormat = DateTimePicker1.CustomFormat
		DateTimePicker4.CustomFormat = DateTimePicker1.CustomFormat
		DateTimePicker5.CustomFormat = DateTimePicker1.CustomFormat
		DateTimePicker6.CustomFormat = DateTimePicker1.CustomFormat
		DateTimePicker7.CustomFormat = DateTimePicker1.CustomFormat
		DateTimePicker8.CustomFormat = DateTimePicker1.CustomFormat
		DateTimePicker9.CustomFormat = DateTimePicker1.CustomFormat
		DateTimePicker10.CustomFormat = DateTimePicker1.CustomFormat
		DateTimePicker11.CustomFormat = DateTimePicker1.CustomFormat
		DateTimePicker12.CustomFormat = DateTimePicker1.CustomFormat
		DateTimePicker13.CustomFormat = DateTimePicker1.CustomFormat
		DateTimePicker14.CustomFormat = DateTimePicker1.CustomFormat

		GroupBox2.Parent = Me.TabPage1
		GroupBox2.Location = New Point(6, 63)
		GroupBox3.Parent = Me.TabPage1
		GroupBox3.Location = New Point(6, 63)
		GroupBox4.Parent = Me.TabPage1
		GroupBox4.Location = New Point(6, 63)
		GroupBox5.Parent = Me.TabPage1
		GroupBox5.Location = New Point(6, 63)

		' Clear form
		RadioButton1.Checked = False
		RadioButton2.Checked = False
		RadioButton3.Checked = False
		RadioButton4.Checked = False
		TextBox1.Text = ScheduleName
		TextBox2.Text = ScheduleCode
		TextBox3.Text = Nothing
		TabControl1.SelectedIndex = 0
		CheckBox14.Checked = False

		' Clear One Time
		MonthCalendar1.SelectionStart = Now
		DateTimePicker1.Value = MonthCalendar1.SelectionStart

		' Clear  Day / Week
		For intX As Integer = 0 To 6
			CheckedListBox1.SetItemChecked(intX, False)
		Next

		CheckBox1.Checked = False
		DateTimePicker2.Value = MonthCalendar1.SelectionStart
		DateTimePicker3.Value = MonthCalendar1.SelectionStart
		DateTimePicker4.Value = MonthCalendar1.SelectionStart
		DateTimePicker5.Value = MonthCalendar1.SelectionStart
		DateTimePicker6.Value = MonthCalendar1.SelectionStart
		DateTimePicker7.Value = MonthCalendar1.SelectionStart

		' Clear  Month
		CheckBox10.Checked = False
		CheckBox11.Checked = False
		ComboBox1.SelectedIndex = 0
		ComboBox2.SelectedIndex = 0
		ComboBox3.SelectedIndex = 0
		ComboBox4.SelectedIndex = 0
		DateTimePicker8.Value = MonthCalendar1.SelectionStart
		DateTimePicker9.Value = MonthCalendar1.SelectionStart
		DateTimePicker10.Value = MonthCalendar1.SelectionStart
		DateTimePicker11.Value = MonthCalendar1.SelectionStart
		DateTimePicker12.Value = MonthCalendar1.SelectionStart
		DateTimePicker13.Value = MonthCalendar1.SelectionStart

		' Clear Interval
		NumericUpDown1.Value = 0
		NumericUpDown2.Value = 1
		NumericUpDown3.Value = 0
		NumericUpDown4.Value = 0
		MonthCalendar2.SelectionStart = MonthCalendar1.SelectionStart
		DateTimePicker14.Value = MonthCalendar1.SelectionStart

		If Not String.IsNullOrEmpty(ScheduleCode) Then
			m_s = New Schedule(ScheduleCode, ScriptID, ScriptScheduleID)

			If m_s.IsDisabled Then
				CheckBox14.Checked = True
			End If

			Select Case m_s.Type
				Case Schedule.ScheduleType.DailyWeekly
					RadioButton2.Checked = True

					For intX As Integer = 0 To m_s.DayWeekDays.Count - 1
						CheckedListBox1.SetItemChecked(CInt(m_s.DayWeekDays(intX)), True)
					Next

					If m_s.DayWeekTimes.Count > 0 Then
						DateTimePicker2.Value = m_s.DayWeekTimes(0)
					End If

					If m_s.DayWeekTimes.Count > 1 Then
						CheckBox1.Checked = True
						DateTimePicker3.Value = m_s.DayWeekTimes(1)
					End If

					If m_s.DayWeekTimes.Count > 2 Then
						CheckBox2.Checked = True
						DateTimePicker4.Value = m_s.DayWeekTimes(2)
					End If

					If m_s.DayWeekTimes.Count > 3 Then
						CheckBox3.Checked = True
						DateTimePicker5.Value = m_s.DayWeekTimes(3)
					End If

					If m_s.DayWeekTimes.Count > 4 Then
						CheckBox4.Checked = True
						DateTimePicker6.Value = m_s.DayWeekTimes(4)
					End If

					If m_s.DayWeekTimes.Count > 5 Then
						CheckBox5.Checked = True
						DateTimePicker7.Value = m_s.DayWeekTimes(5)
					End If

				Case Schedule.ScheduleType.Interval
					RadioButton4.Checked = True

					NumericUpDown1.Value = m_s.IntervalTimeSpan.Days
					NumericUpDown2.Value = m_s.IntervalTimeSpan.Hours
					NumericUpDown3.Value = m_s.IntervalTimeSpan.Minutes
					NumericUpDown4.Value = m_s.IntervalTimeSpan.Seconds

					MonthCalendar2.SelectionStart = m_s.IntervalStartDateTime.Date

					DateTimePicker14.Value = m_s.IntervalStartDateTime

				Case Schedule.ScheduleType.Monthly
					RadioButton3.Checked = True

					If m_s.MonthDates.Count > 0 Then
						ComboBox1.SelectedItem = m_s.MonthDates(0) & Helper.GetDaySuffix(m_s.MonthDates(0))
					End If

					If m_s.MonthDates.Count > 1 Then
						CheckBox11.Checked = True
						ComboBox2.SelectedItem = m_s.MonthDates(1) & Helper.GetDaySuffix(m_s.MonthDates(1))
					End If

					If m_s.MonthDates.Count > 2 Then
						CheckBox12.Checked = True
						ComboBox3.SelectedItem = m_s.MonthDates(2) & Helper.GetDaySuffix(m_s.MonthDates(2))
					End If

					If m_s.MonthDates.Count > 3 Then
						CheckBox13.Checked = True
						ComboBox4.SelectedItem = m_s.MonthDates(3) & Helper.GetDaySuffix(m_s.MonthDates(3))
					End If

					If m_s.MonthTimes.Count > 0 Then
						DateTimePicker13.Value = m_s.MonthTimes(0)
					End If

					If m_s.MonthTimes.Count > 1 Then
						CheckBox10.Checked = True
						DateTimePicker12.Value = m_s.MonthTimes(1)
					End If

					If m_s.MonthTimes.Count > 2 Then
						CheckBox9.Checked = True
						DateTimePicker11.Value = m_s.MonthTimes(2)
					End If

					If m_s.MonthTimes.Count > 3 Then
						CheckBox8.Checked = True
						DateTimePicker10.Value = m_s.MonthTimes(3)
					End If

					If m_s.MonthTimes.Count > 4 Then
						CheckBox7.Checked = True
						DateTimePicker9.Value = m_s.MonthTimes(4)
					End If

					If m_s.MonthTimes.Count > 5 Then
						CheckBox6.Checked = True
						DateTimePicker8.Value = m_s.MonthTimes(5)
					End If

				Case Schedule.ScheduleType.OneTime
					RadioButton1.Checked = True

					MonthCalendar1.SelectionStart = m_s.OneTimeDateTime.Date
					DateTimePicker1.Value = m_s.OneTimeDateTime

			End Select

		End If

	End Sub

	Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

		Dim s As Schedule
		Dim booContinue As Boolean = True

		s = BuildScheduleObject()

		If s Is Nothing Then
			booContinue = False
		End If

		If booContinue Then
			ScheduleCode = s.GetCode

			If Not String.IsNullOrEmpty(TextBox1.Text) Then
				ScheduleName = TextBox1.Text
			Else
				ScheduleName = ScheduleCode
			End If

		Else
			ScheduleName = Nothing
			ScheduleCode = Nothing
		End If

		If booContinue Then
			Me.DialogResult = Windows.Forms.DialogResult.OK
			Me.Close()
		End If

	End Sub

	Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
		ScheduleCode = Nothing
		ScheduleName = Nothing
		Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.Close()
	End Sub

	Private Function BuildScheduleObject() As Schedule

		Dim sReturn As New Schedule(ScriptID, ScriptScheduleID)
		Dim dtp As DateTimePicker
		Dim cb As ComboBox
		Dim booContinue As Boolean = True
		Dim intZ As Integer
		Dim datX As Date

		If CheckBox14.Checked Then
			sReturn.IsDisabled = True
		End If

		If RadioButton1.Checked Then
			sReturn.Type = Schedule.ScheduleType.OneTime

			sReturn.OneTimeDateTime = New Date(MonthCalendar1.SelectionStart.Year, MonthCalendar1.SelectionStart.Month, MonthCalendar1.SelectionStart.Day, DateTimePicker1.Value.Hour, DateTimePicker1.Value.Minute, 0)

		ElseIf RadioButton2.Checked Then
			sReturn.Type = Schedule.ScheduleType.DailyWeekly

			For intX As Integer = 0 To 6

				If CheckedListBox1.GetItemChecked(intX) Then
					sReturn.DayWeekDays.Add(CType(intX, DayOfWeek))
				End If

			Next

			If sReturn.DayWeekDays.Count = 0 Then
				MsgBox("You must pick at least one day for this schedule to run.")
				booContinue = False
			End If

			If booContinue Then

				dtp = DateTimePicker2
				sReturn.DayWeekTimes.Add(New Date(1753, 1, 1, dtp.Value.Hour, dtp.Value.Minute, 0))

				dtp = DateTimePicker3
				If dtp.Enabled Then
					datX = New Date(1753, 1, 1, dtp.Value.Hour, dtp.Value.Minute, 0)
					If Not sReturn.DayWeekTimes.Contains(datX) Then
						sReturn.DayWeekTimes.Add(datX)
					End If
				End If

				dtp = DateTimePicker4
				If dtp.Enabled Then
					datX = New Date(1753, 1, 1, dtp.Value.Hour, dtp.Value.Minute, 0)
					If Not sReturn.DayWeekTimes.Contains(datX) Then
						sReturn.DayWeekTimes.Add(datX)
					End If
				End If

				dtp = DateTimePicker5
				If dtp.Enabled Then
					datX = New Date(1753, 1, 1, dtp.Value.Hour, dtp.Value.Minute, 0)
					If Not sReturn.DayWeekTimes.Contains(datX) Then
						sReturn.DayWeekTimes.Add(datX)
					End If
				End If

				dtp = DateTimePicker6
				If dtp.Enabled Then
					datX = New Date(1753, 1, 1, dtp.Value.Hour, dtp.Value.Minute, 0)
					If Not sReturn.DayWeekTimes.Contains(datX) Then
						sReturn.DayWeekTimes.Add(datX)
					End If
				End If

				dtp = DateTimePicker7
				If dtp.Enabled Then
					datX = New Date(1753, 1, 1, dtp.Value.Hour, dtp.Value.Minute, 0)
					If Not sReturn.DayWeekTimes.Contains(datX) Then
						sReturn.DayWeekTimes.Add(datX)
					End If
				End If

				sReturn.DayWeekTimes.Sort()

			End If

		ElseIf RadioButton3.Checked Then
			sReturn.Type = Schedule.ScheduleType.Monthly

			cb = ComboBox1
			sReturn.MonthDates.Add(CInt(cb.SelectedItem.ToString.Remove(cb.SelectedItem.ToString.Length - 2, 2)))

			cb = ComboBox2
			If cb.Enabled Then
				intZ = CInt(cb.SelectedItem.ToString.Remove(cb.SelectedItem.ToString.Length - 2, 2))
				If Not sReturn.MonthDates.Contains(intZ) Then
					sReturn.MonthDates.Add(intZ)
				End If
			End If

			cb = ComboBox3
			If cb.Enabled Then
				intZ = CInt(cb.SelectedItem.ToString.Remove(cb.SelectedItem.ToString.Length - 2, 2))
				If Not sReturn.MonthDates.Contains(intZ) Then
					sReturn.MonthDates.Add(intZ)
				End If
			End If

			cb = ComboBox4
			If cb.Enabled Then
				intZ = CInt(cb.SelectedItem.ToString.Remove(cb.SelectedItem.ToString.Length - 2, 2))
				If Not sReturn.MonthDates.Contains(intZ) Then
					sReturn.MonthDates.Add(intZ)
				End If
			End If

			dtp = DateTimePicker13
			sReturn.MonthTimes.Add(New Date(1753, 1, 1, dtp.Value.Hour, dtp.Value.Minute, 0))

			dtp = DateTimePicker12
			If dtp.Enabled Then
				datX = New Date(1753, 1, 1, dtp.Value.Hour, dtp.Value.Minute, 0)
				If Not sReturn.MonthTimes.Contains(datX) Then
					sReturn.MonthTimes.Add(datX)
				End If
			End If

			dtp = DateTimePicker11
			If dtp.Enabled Then
				datX = New Date(1753, 1, 1, dtp.Value.Hour, dtp.Value.Minute, 0)
				If Not sReturn.MonthTimes.Contains(datX) Then
					sReturn.MonthTimes.Add(datX)
				End If
			End If

			dtp = DateTimePicker10
			If dtp.Enabled Then
				datX = New Date(1753, 1, 1, dtp.Value.Hour, dtp.Value.Minute, 0)
				If Not sReturn.MonthTimes.Contains(datX) Then
					sReturn.MonthTimes.Add(datX)
				End If
			End If

			dtp = DateTimePicker9
			If dtp.Enabled Then
				datX = New Date(1753, 1, 1, dtp.Value.Hour, dtp.Value.Minute, 0)
				If Not sReturn.MonthTimes.Contains(datX) Then
					sReturn.MonthTimes.Add(datX)
				End If
			End If

			dtp = DateTimePicker8
			If dtp.Enabled Then
				datX = New Date(1753, 1, 1, dtp.Value.Hour, dtp.Value.Minute, 0)
				If Not sReturn.MonthTimes.Contains(datX) Then
					sReturn.MonthTimes.Add(datX)
				End If
			End If

			sReturn.MonthDates.Sort()
			sReturn.MonthTimes.Sort()

		ElseIf RadioButton4.Checked Then
			sReturn.Type = Schedule.ScheduleType.Interval

			sReturn.IntervalTimeSpan = New TimeSpan(CInt(NumericUpDown1.Value), CInt(NumericUpDown2.Value), CInt(NumericUpDown3.Value), CInt(NumericUpDown4.Value), 0)

			If sReturn.IntervalTimeSpan.TotalSeconds < 5 Then
				MsgBox("Your interval must be at least five seconds.")
				booContinue = False
			End If

			If booContinue Then
				sReturn.IntervalStartDateTime = New Date(MonthCalendar2.SelectionStart.Year, MonthCalendar2.SelectionStart.Month, MonthCalendar2.SelectionStart.Day, DateTimePicker14.Value.Hour, DateTimePicker14.Value.Minute, 0)
			End If

		End If

		If Not booContinue Then
			sReturn = Nothing
		End If

		Return sReturn

	End Function

	Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged

		If RadioButton1.Checked Then
			GroupBox2.Visible = True
		Else
			GroupBox2.Visible = False
		End If
	End Sub

	Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged

		If RadioButton2.Checked Then
			GroupBox3.Visible = True
		Else
			GroupBox3.Visible = False
		End If

	End Sub

	Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged

		If RadioButton3.Checked Then
			GroupBox4.Visible = True
		Else
			GroupBox4.Visible = False
		End If

	End Sub

	Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged

		If RadioButton4.Checked Then
			GroupBox5.Visible = True
		Else
			GroupBox5.Visible = False
		End If

	End Sub

	Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged

		If CheckBox1.Checked Then
			DateTimePicker3.Enabled = True
			CheckBox2.Enabled = True
		Else
			DateTimePicker3.Enabled = False
			CheckBox2.Enabled = False
		End If

		CheckBox2.Checked = False

	End Sub

	Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged

		If CheckBox2.Checked Then
			DateTimePicker4.Enabled = True
			CheckBox3.Enabled = True
		Else
			DateTimePicker4.Enabled = False
			CheckBox3.Enabled = False
		End If

		CheckBox3.Checked = False

	End Sub

	Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged

		If CheckBox3.Checked Then
			DateTimePicker5.Enabled = True
			CheckBox4.Enabled = True
		Else
			DateTimePicker5.Enabled = False
			CheckBox4.Enabled = False
		End If

		CheckBox4.Checked = False

	End Sub

	Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged

		If CheckBox4.Checked Then
			DateTimePicker6.Enabled = True
			CheckBox5.Enabled = True
		Else
			DateTimePicker6.Enabled = False
			CheckBox5.Enabled = False
		End If

		CheckBox5.Checked = False

	End Sub

	Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged

		If CheckBox5.Checked Then
			DateTimePicker7.Enabled = True
		Else
			DateTimePicker7.Enabled = False
		End If

	End Sub

	Private Sub CheckBox10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox10.CheckedChanged

		If CheckBox10.Checked Then
			DateTimePicker12.Enabled = True
			CheckBox9.Enabled = True
		Else
			DateTimePicker12.Enabled = False
			CheckBox9.Enabled = False
		End If

		CheckBox9.Checked = False

	End Sub

	Private Sub CheckBox9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox9.CheckedChanged

		If CheckBox9.Checked Then
			DateTimePicker11.Enabled = True
			CheckBox8.Enabled = True
		Else
			DateTimePicker11.Enabled = False
			CheckBox8.Enabled = False
		End If

		CheckBox8.Checked = False

	End Sub

	Private Sub CheckBox8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox8.CheckedChanged

		If CheckBox8.Checked Then
			DateTimePicker10.Enabled = True
			CheckBox7.Enabled = True
		Else
			DateTimePicker10.Enabled = False
			CheckBox7.Enabled = False
		End If

		CheckBox7.Checked = False

	End Sub

	Private Sub CheckBox7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox7.CheckedChanged

		If CheckBox7.Checked Then
			DateTimePicker9.Enabled = True
			CheckBox6.Enabled = True
		Else
			DateTimePicker9.Enabled = False
			CheckBox6.Enabled = False
		End If

		CheckBox6.Checked = False

	End Sub

	Private Sub CheckBox6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged

		If CheckBox6.Checked Then
			DateTimePicker8.Enabled = True
		Else
			DateTimePicker8.Enabled = False
		End If

	End Sub

	Private Sub CheckBox11_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox11.CheckedChanged

		If CheckBox11.Checked Then
			ComboBox2.Enabled = True
			CheckBox12.Enabled = True
		Else
			ComboBox2.Enabled = False
			CheckBox12.Enabled = False
		End If

		CheckBox12.Checked = False

	End Sub

	Private Sub CheckBox12_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox12.CheckedChanged

		If CheckBox12.Checked Then
			ComboBox3.Enabled = True
			CheckBox13.Enabled = True
		Else
			ComboBox3.Enabled = False
			CheckBox13.Enabled = False
		End If

		CheckBox13.Checked = False

	End Sub

	Private Sub CheckBox13_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox13.CheckedChanged

		If CheckBox13.Checked Then
			ComboBox4.Enabled = True
		Else
			ComboBox4.Enabled = False
		End If

	End Sub

	Private m_ScheduleCode As String
	Private m_ScheduleName As String
	Private m_ScriptID As Integer

	Private m_ScriptScheduleID As Integer

	Public Property ScriptScheduleID() As Integer
		Get
			Return m_ScriptScheduleID
		End Get
		Set(ByVal value As Integer)
			m_ScriptScheduleID = value
		End Set
	End Property

	Public Property ScriptID() As Integer
		Get
			Return m_ScriptID
		End Get
		Set(ByVal value As Integer)
			m_ScriptID = value
		End Set
	End Property

	Public Property ScheduleName() As String
		Get
			Return m_ScheduleName
		End Get
		Set(ByVal value As String)
			m_ScheduleName = value
		End Set
	End Property

	Public Property ScheduleCode() As String
		Get
			Return m_ScheduleCode
		End Get
		Set(ByVal value As String)
			m_ScheduleCode = value
		End Set
	End Property

	Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

		Dim s As Schedule

		s = BuildScheduleObject()

		TextBox2.Clear()
		TextBox3.Clear()

		If s IsNot Nothing Then
			TextBox2.Text = s.GetCode
			TextBox3.Text = s.GetNextRunDate(Now).ToString
		End If

	End Sub
End Class
