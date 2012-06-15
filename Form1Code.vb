''''''''' ''''''''' ''''''''' ''''''''' ''''''''' 
''''''''' Copyright 2007-2009 Brian Saville
''''''''' Command3808 is free for non-commercial use.
''''''''' Reuse of the source code is free for non-commercial use
''''''''' provided that the source code used is credited to Command3808.
''''''''' ''''''''' ''''''''' ''''''''' ''''''''' 
Public Class Form1

	Private lSources As New Generic.List(Of ComboBox)
	Private booLoading, booFormLoadStarted As Boolean
	Private WithEvents lTunerPresets As New System.ComponentModel.BindingList(Of TunerPreset)
	Friend WithEvents lTabInfo As System.ComponentModel.BindingList(Of TabPageInfo)
	Friend WithEvents lSourceInfo As SortedList(Of String, Source)
	Private lControlQueue As New SortedList(Of String, Date)
	Private lGroupBoxes As New List(Of GroupBox)
	Private dsScripts As Data.DataSet
	Friend RecordingScriptID As Integer
	Friend RecordingScriptName As String
	Private lScriptSchedules As New List(Of Schedule)

	Private Delegate Sub dHandleResponseReceived(ByVal l As Generic.List(Of String))

	Private Sub HandleResponseReceived(ByVal l As Generic.List(Of String))

		Me.BeginInvoke(New dHandleResponseReceived(AddressOf iHandleResponseReceived), New Object() {l})

	End Sub

	Private Sub iHandleResponseReceived(ByVal l As Generic.List(Of String))
		Try
			Commands.ParseResponse(l)
		Catch ex As Exception
			OutputError(ex)
		End Try
	End Sub

	Private Delegate Sub dHandleMessageReceived(ByVal kvp As ReadWriteKeyValuePair(Of String, Exception))

	Private Sub HandleMessageReceived(ByVal m As ReadWriteKeyValuePair(Of String, Exception))

		Me.BeginInvoke(New dHandleMessageReceived(AddressOf iHandleMessageReceived), New Object() {m})

	End Sub

	Private Sub iHandleMessageReceived(ByVal kvp As ReadWriteKeyValuePair(Of String, Exception))

		WriteDebug(kvp.Key, True)

		If Not IsNothing(kvp.Value) Then
			OutputError(kvp.Value)
		End If

	End Sub

	Private Delegate Sub dDoRunScriptSchedule(ByVal s As Schedule)

	Public Sub DoRunScriptSchedule(ByVal s As Schedule)
		Me.BeginInvoke(New dDoRunScriptSchedule(AddressOf iDoRunScriptSchedule), New Object() {s})
	End Sub

	Public Sub iDoRunScriptSchedule(ByVal s As Schedule)
		s.Run()
	End Sub

	Function GetNowString() As String
		Return Helper.GetNowString
	End Function

	Sub OutputError(ByVal ex As Exception)

		Me.TextBox4.AppendText(vbCrLf & vbCrLf)
		Me.TextBox4.AppendText("An exception occured at " & GetNowString() & ": ")
		Me.TextBox4.AppendText(vbCrLf)
		Me.TextBox4.AppendText(Strings.StrDup(50, "-"))
		Me.TextBox4.AppendText(vbCrLf & vbCrLf)
		Me.TextBox4.AppendText(ex.ToString)
		Me.TextBox4.AppendText(vbCrLf & vbCrLf)
		Me.TextBox4.AppendText(Strings.StrDup(50, "-"))
		Me.TextBox4.AppendText(vbCrLf & vbCrLf)

		MsgBox("An exception occured and it's details have been written to the debug window.", MsgBoxStyle.Critical)

	End Sub

	Sub WriteDebug(ByVal ValueToWrite As String, ByVal IncludeCRLF As Boolean)

		Me.TextBox4.AppendText(ValueToWrite)

		If IncludeCRLF Then
			Me.TextBox4.AppendText(vbCrLf)
		End If

	End Sub

	Sub AddHandlers()
		AddHandler Network_MT.ResponseReady, AddressOf HandleResponseReceived
		AddHandler Network_MT.MessageReady, AddressOf HandleMessageReceived

		SplashScreen1.SetStatus("Adding Handlers")

		For Each c As Control In GetAllControls(Me)

			If TypeOf c Is RangeChangeButton Then
				AddHandler DirectCast(c, RangeChangeButton).Click, AddressOf HandleRangeChangeButtonClick
			End If

			If TypeOf c Is GroupBox Then
				lGroupBoxes.Add(DirectCast(c, GroupBox))
			End If

		Next

	End Sub

	Function GetAllControls(ByVal ParentControl As Control) As Generic.List(Of Control)

		Dim l As New Generic.List(Of Control)

		l.Add(ParentControl)

		If ParentControl.HasChildren Then

			For Each c As Control In ParentControl.Controls
				l.AddRange(GetAllControls(c))
			Next

		End If

		Return l

	End Function

	Sub LoadSettings()

		Dim strX() As String

		If Not String.IsNullOrEmpty(My.CustomSettings.IPAddress.Instance.Value) Then
			SplashScreen1.SetStatus("Setting IP Address")
			strX = My.CustomSettings.IPAddress.Instance.Value.Split("."c)

			NumericUpDown2.Value = CInt(strX(0))
			NumericUpDown3.Value = CInt(strX(1))
			NumericUpDown4.Value = CInt(strX(2))
			NumericUpDown5.Value = CInt(strX(3))

		End If

		SplashScreen1.SetStatus("Initializing User Settings")

		Me.CheckBox2.Checked = My.CustomSettings.PreviewCommands.Instance.Value
		Me.CheckBox3.Checked = My.CustomSettings.AutoClearDebug.Instance.Value
		Me.CheckBox7.Checked = My.CustomSettings.AutoRefreshNet.Instance.Value
		Me.CheckBox8.Checked = My.CustomSettings.MinimizeToTray.Instance.Value
		Me.CheckBox9.Checked = My.CustomSettings.AutoRefreshIPod.Instance.Value

		Me.NumericUpDown1.Value = My.CustomSettings.AutoClearLines.Instance.Value
		Me.NumericUpDown6.Value = My.CustomSettings.AutoRefreshNetSecs.Instance.Value
		Me.NumericUpDown7.Value = My.CustomSettings.DefaultCommandPauseMS.Instance.Value
		Me.NumericUpDown8.Value = My.CustomSettings.AutoRefreshIPodSecs.Instance.Value

		SplashScreen1.SetStatus("Initializing Saved Commands")
		FillSavedCommands()

		SplashScreen1.SetStatus("Initializing Master & Zone Volume Levels")
		FillMasterAndZoneVolumeLevels()

		SplashScreen1.SetStatus("Initializing Channel Volume Levels")
		FillChannelVolumeLevels()

		SplashScreen1.SetStatus("Initializing Range Drop Down Lists")
		FillRangeControls()

		'' Sources
		SplashScreen1.SetStatus("Initializing Source Drop Down Lists")
		FillSources() ' ComboBox3
		FillVideoSelect() ' ComboBox12
		iFillAllSources(ComboBox33)
		iFillAllSources(ComboBox36)
		iFillAllSources(ComboBox57)
		iFillAllSources(ComboBox59)
		iFillAllSources(ComboBox61)

		lSources.Add(ComboBox3)
		lSources.Add(ComboBox12)
		lSources.Add(ComboBox33)
		lSources.Add(ComboBox36)
		lSources.Add(ComboBox57)
		lSources.Add(ComboBox59)
		lSources.Add(ComboBox61)

		SplashScreen1.SetStatus("Initializing On / Off Buttons")
		ConfigureOnOffRadioButtons()

		SplashScreen1.SetStatus("Initializing Other Drop Down Lists")
		ConfigureFixedListComboBoxes()

		SplashScreen1.SetStatus("Initializing Tuner Presets")
		DataGridView1.DataSource = lTunerPresets
		DataGridView1.ClearSelection()
		DataGridView1.CurrentCell = Nothing
		FillTunerPresetEditControls()

		SplashScreen1.SetStatus("Initializing Tab Details")
		FillTabs()

		SplashScreen1.SetStatus("Loading Scripts")
		FillScripts(Nothing, Nothing, Nothing)

	End Sub

	Sub FillScripts(ByVal SelectedScriptValue As String, ByVal SelectedScheduleValue As String, ByVal SelectedEntryValue As String)

		ListBox1.DataSource = Nothing
		ListBox1.Items.Clear()

		dsScripts = New DataSet

		dsScripts.Tables.Add(DB.Script.ReadAllASCByName)
		dsScripts.Tables.Add(DB.ScriptSchedule.ReadAll)
		dsScripts.Tables.Add(DB.ScriptEntry.ReadAllASCBySequence)
		dsScripts.Relations.Add("Script_ScriptSchedule", dsScripts.Tables(0).Columns("ScriptID"), dsScripts.Tables(1).Columns("ScriptID"))
		dsScripts.Relations.Add("Script_ScriptEntry", dsScripts.Tables(0).Columns("ScriptID"), dsScripts.Tables(2).Columns("ScriptID"))

		ListBox1.DataSource = dsScripts.Tables(0)
		ListBox1.DisplayMember = "name"
		ListBox1.ValueMember = "name"

		If Not String.IsNullOrEmpty(SelectedScriptValue) Then
			ListBox1.SelectedValue = SelectedScriptValue
		ElseIf ListBox1.Items.Count >= 1 Then
			ListBox1.SelectedIndex = 0
		End If

		FillScriptDetail(SelectedScheduleValue, SelectedEntryValue)

	End Sub

	Sub FillScriptDetail(ByVal SelectedScheduleValue As String, ByVal SelectedEntryValue As String)

		Dim dvSchedules As DataView

		ListBox2.DataSource = Nothing
		ListBox3.DataSource = Nothing
		ListBox2.Items.Clear()
		ListBox3.Items.Clear()

		If ListBox1.SelectedItem IsNot Nothing Then
			dvSchedules = DirectCast(ListBox1.SelectedItem, DataRowView).CreateChildView("Script_ScriptSchedule")
			ListBox2.DataSource = dvSchedules
			ListBox2.DisplayMember = "name"
			ListBox2.ValueMember = "name"

			If Not String.IsNullOrEmpty(SelectedScheduleValue) Then
				ListBox2.SelectedValue = SelectedScheduleValue
			ElseIf ListBox2.Items.Count >= 1 Then
				ListBox2.SelectedIndex = 0
			End If

			SyncLock objSL

				lScriptSchedules.Clear()

				For Each drv As DataRowView In dvSchedules
					lScriptSchedules.Add(New Schedule(drv("code").ToString, CInt(drv("ScriptID")), CInt(drv("ScriptScheduleID"))))
				Next

			End SyncLock

			ListBox3.DataSource = DirectCast(ListBox1.SelectedItem, DataRowView).CreateChildView("Script_ScriptEntry")
			ListBox3.DisplayMember = "description"
			ListBox3.ValueMember = "description"

			If Not String.IsNullOrEmpty(SelectedEntryValue) Then
				ListBox3.SelectedValue = SelectedEntryValue
			ElseIf ListBox3.Items.Count >= 1 Then
				ListBox3.SelectedIndex = 0
			End If
		End If

	End Sub

	Sub ConfigureOnOffRadioButtons()

		RadioButton1.Tag = OnOffCommand.PowerOff
        RadioButton2.Tag = OnOffCommand.PowerOn
        RadioButton3.Tag = OnOffCommand.ToneCtrlOn
        RadioButton4.Tag = OnOffCommand.ToneCtrlOff
        RadioButton5.Tag = OnOffCommand.CinemaEQOn
        RadioButton6.Tag = OnOffCommand.CinemaEQOff
        RadioButton7.Tag = OnOffCommand.AFDOn
        RadioButton8.Tag = OnOffCommand.AFDOff
        RadioButton9.Tag = OnOffCommand.SWAttOn
        RadioButton10.Tag = OnOffCommand.SWAttOff
        RadioButton11.Tag = OnOffCommand.SWOn
        RadioButton12.Tag = OnOffCommand.SWOff
        RadioButton13.Tag = OnOffCommand.PanoramaOn
        RadioButton14.Tag = OnOffCommand.PanoramaOff
        RadioButton15.Tag = OnOffCommand.ZoneMainOn
        RadioButton16.Tag = OnOffCommand.ZoneMainOff
        RadioButton17.Tag = OnOffCommand.Zone2On
        RadioButton18.Tag = OnOffCommand.Zone2Off
        RadioButton19.Tag = OnOffCommand.Zone3On
        RadioButton20.Tag = OnOffCommand.Zone3Off
        RadioButton21.Tag = OnOffCommand.DSPEffectOn
        RadioButton22.Tag = OnOffCommand.DSPEffectOff
        RadioButton23.Tag = OnOffCommand.Zone2HPFOn
        RadioButton24.Tag = OnOffCommand.Zone2HPFOff
        RadioButton25.Tag = OnOffCommand.Zone3HPFOn
        RadioButton26.Tag = OnOffCommand.Zone3HPFOff
        RadioButton33.Tag = OnOffCommand.PowerOff
        RadioButton32.Tag = OnOffCommand.PowerOn
        RadioButton34.Tag = OnOffCommand.ZoneMainOn
        RadioButton35.Tag = OnOffCommand.ZoneMainOff
        RadioButton36.Tag = OnOffCommand.Zone2On
        RadioButton37.Tag = OnOffCommand.Zone2Off
        RadioButton38.Tag = OnOffCommand.Zone3On
        RadioButton39.Tag = OnOffCommand.Zone3Off
        RadioButton40.Tag = OnOffCommand.DynamicEQOn
        RadioButton41.Tag = OnOffCommand.DynamicEQOff

    End Sub

    Sub ConfigureFixedListComboBoxes()

        Dim l As KeyValueList(Of String, String)

        ' Audio Input Mode
        l = New KeyValueList(Of String, String)
        l.AddNew("Auto", "AUTO")
        l.AddNew("HDMI", "HDMI")
        l.AddNew("Digital", "DIGITAL")
        l.AddNew("Analog", "ANALOG")
        l.AddNew("EXT. IN", "EXT.IN-1")
        iConfigureFixedListComboBoxes(FixedListComboBox.AudioInputMode, l, "SD")

        ' Digital Audio Mode
        l = New KeyValueList(Of String, String)
        l.AddNew("Auto", "AUTO")
        l.AddNew("PCM", "PCM")
        l.AddNew("DTS", "DTS")
        iConfigureFixedListComboBoxes(FixedListComboBox.DigitalAudioMode, l, "DC")

        ' Zone 2 Channel Setting
        l = New KeyValueList(Of String, String)
        l.AddNew("Stereo", "ST")
        l.AddNew("Mono", "MONO")
        iConfigureFixedListComboBoxes(FixedListComboBox.Zone2ChannelSetting, l, Helper.GetZonePrefix(Zones.Zone2) & "CS")

        ' Zone 3 Channel Setting
        l = New KeyValueList(Of String, String)
        l.AddNew("Stereo", "ST")
        l.AddNew("Mono", "MONO")
        iConfigureFixedListComboBoxes(FixedListComboBox.Zone3ChannelSetting, l, Helper.GetZonePrefix(Zones.Zone3) & "CS")

        ' DRC
        l = New KeyValueList(Of String, String)
        l.AddNew("Off", "OFF")
        l.AddNew("Auto", "AUTO")
        l.AddNew("Low", "LOW")
        l.AddNew("Medium", "MID")
        l.AddNew("High", "HI")
        iConfigureFixedListComboBoxes(FixedListComboBox.DRC, l, "PSDRC ")

        ' D.COMP
        l = New KeyValueList(Of String, String)
        l.AddNew("Off", "OFF")
        l.AddNew("Low", "LOW")
        l.AddNew("Medium", "MID")
        l.AddNew("High", "HI")
        iConfigureFixedListComboBoxes(FixedListComboBox.DCOMP, l, "PSDCO ")

        ' Room Size
        l = New KeyValueList(Of String, String)
        l.AddNew("Small", "S")
        l.AddNew("Medium-Small", "MS")
        l.AddNew("Medium", "M")
        l.AddNew("Medium-Large", "ML")
        l.AddNew("Large", "L")
        iConfigureFixedListComboBoxes(FixedListComboBox.RoomSize, l, "PSRSZ ")

        ' Night Mode
        l = New KeyValueList(Of String, String)
        l.AddNew("Off", "OFF")
        l.AddNew("Low", "LOW")
        l.AddNew("Medium", "MID")
        l.AddNew("High", "HI")
        iConfigureFixedListComboBoxes(FixedListComboBox.Night, l, "PSNIGHT ")

        ' Restorer
        l = New KeyValueList(Of String, String)
        l.AddNew("Off", "OFF")
        l.AddNew("Mode 1", "MODE1")
        l.AddNew("Mode 2", "MODE2")
        l.AddNew("Mode 3", "MODE3")
        iConfigureFixedListComboBoxes(FixedListComboBox.Restorer, l, "PSRSTR ")

        ' Resolution
        l = New KeyValueList(Of String, String)
        l.AddNew("Auto", "AUTO")
        l.AddNew("480p", "48P")
        l.AddNew("720p", "72P")
        l.AddNew("1080i", "10I")
        l.AddNew("1080p", "10P")
        iConfigureFixedListComboBoxes(FixedListComboBox.Resolution, l, "VSSC")

        ' Aspect Ratio
        l = New KeyValueList(Of String, String)
        l.AddNew("Normal", "NRM")
        l.AddNew("Full", "FUL")
        iConfigureFixedListComboBoxes(FixedListComboBox.AspectRatio, l, "VSASP")

        ' Surround Mode
        l = New KeyValueList(Of String, String)
        l.AddNew("Music", "MUSIC")
        l.AddNew("Cinema", "CINEMA")
        l.AddNew("Game", "GAME")
        l.AddNew("Pro Logic", "PRO LOGIC")
        l.AddNew("PLIIz Height", "HEIGHT")
        iConfigureFixedListComboBoxes(FixedListComboBox.SurroundMode, l, "PSMODE:")

        ' Surround Back Mode
        l = New KeyValueList(Of String, String)
        l.AddNew("Matrix On", "MTRX ON")
        l.AddNew("Matrix Off", "NON MTRX")
        l.AddNew("PLIIx Cinema", "PL2X CINEMA")
        l.AddNew("PLIIx Music", "PL2X MUSIC")
        l.AddNew("On", "ON")
        l.AddNew("Off", "OFF")
        iConfigureFixedListComboBoxes(FixedListComboBox.SurroundBackMode, l, "PSSB:")

        ' Surround Setting
        l = New KeyValueList(Of String, String)
        l.AddNew("Standard (Dolby)", "DOLBY DIGITAL")
        l.AddNew("Standard (DTS)", "DTS SURROUND")
        l.AddNew("Stereo", "STEREO")
        l.AddNew("Direct", "DIRECT")
        l.AddNew("Pure Direct", "PURE DIRECT")
        l.AddNew("Multi Channel Stereo", "MCH STEREO")
        l.AddNew("DSP Rock Arena", "ROCK ARENA")
        l.AddNew("DSP Jazz Club", "JAZZ CLUB")
        l.AddNew("DSP Mono Movie", "MONO MOVIE")
        l.AddNew("DSP Matrix", "MATRIX")
        l.AddNew("DSP Video Game", "VIDEO GAME")
        l.AddNew("DSP Virtual", "VIRTUAL")
        iConfigureFixedListComboBoxes(FixedListComboBox.SurroundSetting, l, "MS")

        ' Audyssey
        l = New KeyValueList(Of String, String)
        l.AddNew("Audyssey", "AUDYSSEY")
        l.AddNew("Audyssey Bypass Fronts", "BYP.LR")
        l.AddNew("Audyssey Flat", "FLAT")
        l.AddNew("Custom Curve", "MANUAL")
        l.AddNew("Off", "OFF")
        iConfigureFixedListComboBoxes(FixedListComboBox.AudysseyMode, l, "PSMULTEQ:")

        ' Dynamic Volume
        l = New KeyValueList(Of String, String)
        l.AddNew("Midnight", "NGT")
        l.AddNew("Evening", "EVE")
        l.AddNew("Day", "DAY")
        l.AddNew("Off", "OFF")
        iConfigureFixedListComboBoxes(FixedListComboBox.DynamicVolume, l, "PSDYNVOL ")

        ' Reference Offset
        l = New KeyValueList(Of String, String)
        l.AddNew("0", "0")
        l.AddNew("5", "5")
        l.AddNew("10", "10")
        l.AddNew("15", "15")
        iConfigureFixedListComboBoxes(FixedListComboBox.ReferenceLevel, l, "PSREFLEV ")

    End Sub

    Sub iConfigureFixedListComboBoxes(ByVal m As FixedListComboBox, ByVal l As KeyValueList(Of String, String), ByVal CommandTag As String)

        Dim cb As ComboBox

        cb = GetFixedListComboBox(m)

        cb.Tag = CommandTag
        cb.DisplayMember = "Key"
        cb.ValueMember = "Value"
        cb.DataSource = l
        cb.SelectedIndex = -1

        AddHandler cb.SelectedIndexChanged, AddressOf HandleFixedListComboBoxChange

    End Sub

    Sub StartupQuery()

        For Each tpi As TabPageInfo In lTabInfo

            If tpi.QueryOnLoad Then

                Select Case tpi.TabName
                    Case tabMain.Name
                        Query(QueryType.Main)
                    Case tabSurroundParams.Name
                        Query(QueryType.SurroundParams)
                    Case tabVideoParams.Name
                        Query(QueryType.VideoParams)
                    Case tabOtherParams.Name
                        Query(QueryType.OtherParams)
                    Case tabZone2.Name
                        Query(QueryType.Zone2)
                    Case tabZone3.Name
                        Query(QueryType.Zone3)
                    Case tabTuner.Name
                        Query(QueryType.Tuner)
                    Case tabNet.Name
                        Query(QueryType.NET)
                End Select

            End If

        Next

        Query(QueryType.Sources)
        Query(QueryType.Presets)

    End Sub

    Sub DrawTabs()

        Dim intY As Integer
        Dim tps As New Generic.List(Of TabPage)
        Dim tp As TabPage

        SplashScreen1.SetStatus("Initializing Tab Order")

        For Each tp In TabControl1.TabPages
            tps.Add(tp)
        Next

        TabControl1.TabPages.Clear()

        For Each tpi As TabPageInfo In lTabInfo

            If tpi.IsDefault AndAlso intY = 0 Then
                intY = tpi.TabOrder - 1
            End If

            For Each tp In tps
                If tp.Name = tpi.TabName Then
                    TabControl1.TabPages.Add(tp)
                    Exit For
                End If
            Next

        Next

        ' Always ddd script, debug, and settings tabs at the end
        TabControl1.TabPages.Add(tps(tps.Count - 3))
        TabControl1.TabPages.Add(tps(tps.Count - 2))
        TabControl1.TabPages.Add(tps(tps.Count - 1))

        For intX As Integer = 0 To TabControl1.TabCount - 1
            SplashScreen1.SetStatus("Initializing " & TabControl1.TabPages(intX).Text & " Tab (" & intX + 1 & " of " & TabControl1.TabCount & ")")
            TabControl1.SelectTab(intX)
        Next

        TabControl1.SelectTab(intY)

    End Sub

    Sub Query(ByVal qt As QueryType)

        Select Case qt
            Case QueryType.Full
                QueryMain()
                QuerySurroundParams(False)
                QueryVideoParams()
                QueryOtherParams()
                QueryZone2()
                QueryZone3()
                QuerySources()
                QueryTuner()
                QueryNet()
                QueryIPod()
            Case QueryType.Main
                QueryMain()
            Case QueryType.SurroundParams
                QuerySurroundParams(True)
            Case QueryType.VideoParams
                QueryVideoParams()
            Case QueryType.OtherParams
                QueryOtherParams()
            Case QueryType.Zone2
                QueryZone2()
            Case QueryType.Zone3
                QueryZone3()
            Case QueryType.Sources
                QuerySources()
            Case QueryType.Presets
                QueryPresets()
            Case QueryType.Tuner
                QueryTuner()
            Case QueryType.NET
                QueryNet()
            Case QueryType.iPod
                QueryIPod()
        End Select

    End Sub

    Sub QueryMain()

        Commands.QueryOnOff(OnOffCommand.Power)
        Commands.QueryOnOff(OnOffCommand.ZoneMain)
        Commands.QueryMasterVolume()
        Commands.QueryChannelVolume(Channel.LeftFront) ' this will query all channels
        Commands.QueryOnOff(OnOffCommand.Mute)
        Commands.QuerySource()
        Commands.QuerySourceStatus()
        Commands.QueryVideoSelect()
        Commands.QueryAudioInputMode()
        Commands.QueryDigitalAudioMode()
        Commands.QuerySurroundSettings()  ' Surround setting, Surround mode (surround tab), surround back mode (surround tab)

    End Sub

    Sub QuerySurroundParams(ByVal QuerySurroundSettings As Boolean)

        Commands.QueryRange(RangeControlType.LFE)
        Commands.QueryRange(RangeControlType.Dimension)
        Commands.QueryRange(RangeControlType.CenterWidth)
        Commands.QueryRange(RangeControlType.CenterImage)
        Commands.QueryRange(RangeControlType.DSPSurroundDelay)

        Commands.QueryOnOff(OnOffCommand.CinemaEQ)
        Commands.QueryOnOff(OnOffCommand.AFD)
        Commands.QueryOnOff(OnOffCommand.SWAtt)
        Commands.QueryOnOff(OnOffCommand.SW)
        Commands.QueryOnOff(OnOffCommand.Panorama)
        Commands.QueryOnOff(OnOffCommand.DSPEffect)

        Commands.QueryThis(ComboBox37.Tag.ToString) ' DRC
        Commands.QueryThis(ComboBox38.Tag.ToString) ' D.COMP
        Commands.QueryThis(ComboBox39.Tag.ToString) ' Room Size

        If QuerySurroundSettings Then
            Commands.QuerySurroundSettings() ' Surround mode, surround back mode, surround setting (main tab)
        End If

    End Sub

    Sub QueryVideoParams()

        Commands.QueryRange(RangeControlType.Contrast)
        Commands.QueryRange(RangeControlType.Brightness)
        Commands.QueryRange(RangeControlType.Chroma)
        Commands.QueryRange(RangeControlType.Hue)
        Commands.QueryRange(RangeControlType.AudioDelay)
        Commands.QueryThis(ComboBox49.Tag.ToString) ' Resolution
        Commands.QueryThis(ComboBox50.Tag.ToString) ' Aspect Ratio

    End Sub

    Sub QueryOtherParams()

        Commands.QueryRange(RangeControlType.Bass)
        Commands.QueryRange(RangeControlType.Treble)

        Commands.QueryOnOff(OnOffCommand.ToneCtrl)
        Commands.QueryOnOff(OnOffCommand.DynamicEQ)

        Commands.QueryThis(ComboBox40.Tag.ToString) ' Night Mode
        Commands.QueryThis(ComboBox41.Tag.ToString) ' Restorer Mode

        Commands.QueryAudysseyMode()
        Commands.QueryDynVol()
        Commands.QueryRefLev()

        Commands.QueryRemoteAndPanelLock()

    End Sub

    Sub QueryZone2()

        Commands.QueryOnOff(OnOffCommand.Zone2) ' this also fills zone volume and source
        Commands.QueryOnOff(OnOffCommand.Zone2Mute)
        Commands.QueryOnOff(OnOffCommand.Zone2HPF)

        Commands.QueryRange(RangeControlType.Zone2Bass)
        Commands.QueryRange(RangeControlType.Zone2Treble)

        Commands.QueryChannelVolume(Channel.Zone2Left) ' this will query both channels
        Commands.QueryZoneChannelSetting(Zones.Zone2)

    End Sub

    Sub QueryZone3()

        Commands.QueryOnOff(OnOffCommand.Zone3) ' this also fills zone volume and source
        Commands.QueryOnOff(OnOffCommand.Zone3Mute)
        Commands.QueryOnOff(OnOffCommand.Zone3HPF)

        Commands.QueryRange(RangeControlType.Zone3Bass)
        Commands.QueryRange(RangeControlType.Zone3Treble)

        Commands.QueryChannelVolume(Channel.Zone3Left)  ' this will query both channels
        Commands.QueryZoneChannelSetting(Zones.Zone3)

    End Sub

    Sub QuerySources()

        Commands.QuerySourceNames()
        Commands.QuerySourceStatus()

    End Sub

    Sub QueryPresets()

        Commands.QueryPresets()

    End Sub

    Sub QueryTuner()
        Commands.QueryTuner()
    End Sub

    Sub QueryNet()
        ClearNet()
        Commands.QueryNet()
    End Sub

    Sub QueryIPod()
        ClearNet()
        Commands.QueryIPod()
    End Sub

    Sub FillMasterAndZoneVolumeLevels()

        Dim v As MasterVolume
        Dim l As Generic.List(Of MasterVolume)
        Dim cb As ComboBox

        cb = ComboBox2
        cb.Items.Clear()
        cb.DisplayMember = "DisplayValue"
        l = New Generic.List(Of MasterVolume)

        For decX As Decimal = -80.5D To 18 Step 0.5D
            v = New MasterVolume(decX, -80)
            l.Add(v)
        Next

        cb.Items.AddRange(l.ToArray)

        cb = ComboBox58
        cb.Items.Clear()
        cb.DisplayMember = "DisplayValue"
        cb.Items.AddRange(l.ToArray)

        cb = ComboBox23
        cb.Items.Clear()
        cb.DisplayMember = "DisplayValue"
        l = New Generic.List(Of MasterVolume)

        For decX As Decimal = -71D To 18
            v = New MasterVolume(decX, -70)
            l.Add(v)
        Next

        cb.Items.AddRange(l.ToArray)

        cb = ComboBox60
        cb.Items.Clear()
        cb.DisplayMember = "DisplayValue"
        cb.Items.AddRange(l.ToArray)

        cb = ComboBox24
        cb.Items.Clear()
        cb.DisplayMember = "DisplayValue"
        l = New Generic.List(Of MasterVolume)

        For decX As Decimal = -71D To 18
            v = New MasterVolume(decX, -70)
            l.Add(v)
        Next

        cb.Items.AddRange(l.ToArray)

        cb = ComboBox62
        cb.Items.Clear()
        cb.DisplayMember = "DisplayValue"
        cb.Items.AddRange(l.ToArray)

    End Sub

    Sub FillChannelVolumeLevels()

        iFillChannelVolumeLevels(ComboBox4, Channel.LeftFront)
        iFillChannelVolumeLevels(ComboBox5, Channel.Center)
        iFillChannelVolumeLevels(ComboBox6, Channel.RightFront)
        iFillChannelVolumeLevels(ComboBox7, Channel.LeftSurround)
        iFillChannelVolumeLevels(ComboBox8, Channel.RightSurround)
        iFillChannelVolumeLevels(ComboBox67, Channel.LeftHeight)
        iFillChannelVolumeLevels(ComboBox66, Channel.RightHeight)
        iFillChannelVolumeLevels(ComboBox68, Channel.LeftWide)
        iFillChannelVolumeLevels(ComboBox69, Channel.RightWide)
        iFillChannelVolumeLevels(ComboBox9, Channel.LeftRear)
        iFillChannelVolumeLevels(ComboBox10, Channel.RightRear)
        iFillChannelVolumeLevels(ComboBox11, Channel.Subwoofer)
        iFillChannelVolumeLevels(ComboBox29, Channel.Zone2Right)
        iFillChannelVolumeLevels(ComboBox30, Channel.Zone2Left)
        iFillChannelVolumeLevels(ComboBox31, Channel.Zone3Right)
        iFillChannelVolumeLevels(ComboBox32, Channel.Zone3Left)

    End Sub

    Sub iFillChannelVolumeLevels(ByVal cb As ComboBox, ByVal c As Channel)

        Dim v As ChannelVolume
        Dim decY, decStep As Decimal
        Dim l As New Generic.List(Of ChannelVolume)

        cb.DisplayMember = "DisplayValue"
        cb.Tag = c

        If c <> Channel.Subwoofer Then
            decY = -12D
        Else
            decY = -12.5D
        End If

        Select Case c
            Case Channel.Zone2Left, Channel.Zone2Right, Channel.Zone3Left, Channel.Zone3Right
                decStep = 1
            Case Else
                decStep = 0.5D
        End Select

        Dim sw As Stopwatch

        For decX As Decimal = decY To 12 Step decStep
            sw = Stopwatch.StartNew()
            v = New ChannelVolume(decX)
            l.Add(v)
        Next

        cb.Items.AddRange(l.ToArray)

    End Sub

    Sub FillRangeControls()

        iFillRangeControls(ComboBox16, -6, 6, 50, 1, RangeControlType.Bass, False, "dB")
        iFillRangeControls(ComboBox17, -6, 6, 50, 1, RangeControlType.Treble, False, "dB")
        iFillRangeControls(ComboBox18, -10, 0, 0, 1, RangeControlType.LFE, True, "dB")
        iFillRangeControls(ComboBox19, 0, 6, 0, 1, RangeControlType.Dimension, False, "")
        iFillRangeControls(ComboBox20, 0, 7, 0, 1, RangeControlType.CenterWidth, False, "")
        iFillRangeControls(ComboBox21, 0, 1, 0, 0.1D, RangeControlType.CenterImage, False, "")
        iFillRangeControls(ComboBox26, -10, 10, 50, 1, RangeControlType.Zone2Bass, False, "dB")
        iFillRangeControls(ComboBox25, -10, 10, 50, 1, RangeControlType.Zone2Treble, False, "dB")
        iFillRangeControls(ComboBox28, -10, 10, 50, 1, RangeControlType.Zone3Bass, False, "dB")
        iFillRangeControls(ComboBox27, -10, 10, 50, 1, RangeControlType.Zone3Treble, False, "dB")
        iFillRangeControls(ComboBox43, 0, 300, 0, 10, RangeControlType.DSPSurroundDelay, False, "ms")
        iFillRangeControls(ComboBox44, 0, 200, 0, 1, RangeControlType.AudioDelay, False, "ms")
        iFillRangeControls(ComboBox45, -6, 6, 50, 1, RangeControlType.Contrast, False, "dB")
        iFillRangeControls(ComboBox46, 0, 12, 0, 1, RangeControlType.Brightness, False, "dB")
        iFillRangeControls(ComboBox47, -6, 6, 50, 1, RangeControlType.Chroma, False, "dB")
        iFillRangeControls(ComboBox48, -6, 6, 50, 1, RangeControlType.Hue, False, "dB")
        iFillRangeControls(ComboBox42, 1, 15, 0, 1, RangeControlType.DSPEffectLevel, False, "")

    End Sub

    Sub iFillRangeControls(ByVal cb As ComboBox, ByVal MinVal As Decimal, ByVal MaxVal As Decimal, ByVal DenonValueForZero As Integer, ByVal MinimumIncrement As Decimal, ByVal dt As RangeControlType, ByVal DenonValueDecrements As Boolean, ByVal UnitName As String)

        Dim v As DenonRangeControl
        Dim l As New Generic.List(Of DenonRangeControl)

        cb.DisplayMember = "DisplayValue"
        cb.Tag = dt

        For decX As Decimal = MinVal To MaxVal Step MinimumIncrement
            v = New DenonRangeControl(decX, DenonValueForZero, MinimumIncrement, DenonValueDecrements, UnitName, dt)
            l.Add(v)
        Next

        cb.Items.AddRange(l.ToArray)

        AddHandler cb.SelectedIndexChanged, AddressOf HandleDenonControlValueChange

    End Sub

    Sub FillSavedCommands()

        Dim l As New Generic.List(Of String)

        ComboBox1.Items.Clear()

        If IsNothing(My.CustomSettings.SavedCommands.Instance.Value) Then
            My.CustomSettings.SavedCommands.Instance.Value = New Collections.Specialized.StringCollection
        End If

        If My.CustomSettings.SavedCommands.Instance.Value.Count > 0 Then

            l.Add("-- Select An Item From Below --")

            For Each strX As String In My.CustomSettings.SavedCommands.Instance.Value
                l.Add(strX)
            Next

            ComboBox1.Items.AddRange(l.ToArray)

            ComboBox1.SelectedIndex = 0

        End If

    End Sub

    Sub FillTabs()

        Dim tp As TabPage
        Dim l As System.ComponentModel.BindingList(Of TabPageInfo)
        Dim lInfo, lTabs As New Generic.List(Of String)
        Dim intReservedTabs As Integer = 3 ' scripts, debug, settings - these will always appear, so don't put them in the tab info grid

        If My.CustomSettings.Tabs.Instance.Value.Count = 0 Then

            l = New System.ComponentModel.BindingList(Of TabPageInfo)

            For intX As Integer = 0 To TabControl1.TabPages.Count - (intReservedTabs + 1)
                tp = TabControl1.TabPages(intX)

                l.Add(New TabPageInfo(tp.Name, tp.Text, intX + 1, False, False))

                If intX = 0 Then
                    l(intX).IsDefault = True
                End If

            Next

            My.CustomSettings.Tabs.Instance.Value = l

        Else
            l = My.CustomSettings.Tabs.Instance.Value

            If l.Count <> TabControl1.TabPages.Count - intReservedTabs Then

                For intX As Integer = 0 To TabControl1.TabPages.Count - (intReservedTabs + 1)
                    lTabs.Add(TabControl1.TabPages(intX).Name)
                Next

                For intX As Integer = 0 To l.Count - 1
                    lInfo.Add(l(intX).TabName)
                Next

                If l.Count < TabControl1.TabPages.Count - intReservedTabs Then

                    For Each strX As String In lTabs
                        If Not lInfo.Contains(strX) Then
                            tp = DirectCast(TabControl1.Controls(strX), TabPage)
                            l.Add(New TabPageInfo(tp.Name, tp.Text, l.Count + 1, False, False))
                        End If
                    Next
                Else

                End If

                My.CustomSettings.Tabs.Instance.Value = l

            End If

        End If
        lTabInfo = l
        DataGridView2.DataSource = lTabInfo

    End Sub

    Sub FillSources()

        Dim s As Source
        Dim l As SortedList(Of String, Source)
        Dim l2 As New Generic.List(Of Source)

        ComboBox3.Items.Clear()
        ComboBox3.DisplayMember = "SourceName"

        If My.CustomSettings.Sources.Instance.Value.Count = 0 Then

            l = New SortedList(Of String, Source)

            s = New Source("TUNER", SourceType.TUNER)
            l.Add(s.SourceName, s)
            s = New Source("CD", SourceType.CD)
            l.Add(s.SourceName, s)
            s = New Source("BD", SourceType.BD)
            l.Add(s.SourceName, s)
            s = New Source("DVD", SourceType.DVD)
            l.Add(s.SourceName, s)
            s = New Source("TV", SourceType.TV)
            l.Add(s.SourceName, s)
            s = New Source("SAT/CBL", SourceType.SATCBL)
            l.Add(s.SourceName, s)
            s = New Source("GAME", SourceType.GAME)
            l.Add(s.SourceName, s)
            s = New Source("GAME2", SourceType.GAME2)
            l.Add(s.SourceName, s)
            s = New Source("V.AUX", SourceType.VAUX)
            l.Add(s.SourceName, s)
            s = New Source("DOCK", SourceType.DOCK)
            l.Add(s.SourceName, s)
            s = New Source("NET/USB", SourceType.NETUSB)
            l.Add(s.SourceName, s)
            s = New Source("Source", SourceType.SOURCE)
            l.Add(s.SourceName, s)

            My.CustomSettings.Sources.Instance.Value = l

        Else
            l = My.CustomSettings.Sources.Instance.Value
        End If

        lSourceInfo = l

        For Each kvp As KeyValuePair(Of String, Source) In lSourceInfo

            s = kvp.Value

            If s.IsActive AndAlso s.SourceType <> SourceType.SOURCE Then
                l2.Add(s)
            End If

        Next

        ComboBox3.Items.AddRange(l2.ToArray)

    End Sub

    Sub FillVideoSelect()

        Dim s As Source
        Dim l As New Generic.List(Of Source)

        ComboBox12.Items.Clear()
        ComboBox12.DisplayMember = "SourceName"

        For Each kvp As KeyValuePair(Of String, Source) In lSourceInfo

            s = kvp.Value

            If s.IsActive Then

                Select Case s.SourceType
                    Case SourceType.BD, SourceType.DVD, SourceType.TV, SourceType.SATCBL, SourceType.GAME, SourceType.GAME2, SourceType.VAUX, SourceType.SOURCE
                        l.Add(s)
                End Select

            End If

        Next

        ComboBox12.Items.AddRange(l.ToArray)

    End Sub

    Sub iFillAllSources(ByVal cb As ComboBox)

        Dim s As Source
        Dim l As New Generic.List(Of Source)

        cb.Items.Clear()
        cb.DisplayMember = "SourceName"

        For Each kvp As KeyValuePair(Of String, Source) In lSourceInfo

            s = kvp.Value

            If s.IsActive Then
                l.Add(s)
            End If

        Next

        cb.Items.AddRange(l.ToArray)

    End Sub

    Sub FillTunerPresetEditControls()

        Dim strX As String
        Dim l As New KeyValueList(Of String, String)
        Dim cb As ComboBox

        cb = ComboBox51

        cb.DisplayMember = "Key"
        cb.ValueMember = "Value"

        For intX As Integer = 8750 To 10790 Step 20

            strX = intX.ToString.PadLeft(6, "0"c)

            l.AddNew(Helper.GetTunerFrequencyDisplayValue(strX), strX)
        Next

        For intX As Integer = 52000 To 171000 Step 1000

            strX = intX.ToString.PadLeft(6, "0"c)

            l.AddNew(Helper.GetTunerFrequencyDisplayValue(strX), strX)
        Next

        cb.DataSource = l
        cb.SelectedIndex = 0

        cb = ComboBox53
        cb.BindingContext = New BindingContext
        cb.DisplayMember = "Key"
        cb.ValueMember = "Value"
        cb.DataSource = l
        cb.SelectedIndex = 0

        l = New KeyValueList(Of String, String)
        cb = ComboBox52

        cb.DisplayMember = "Key"
        cb.ValueMember = "Value"

        For intX As Integer = Asc("A") To Asc("G")

            For intY As Integer = 1 To 8

                strX = Chr(intX) & intY.ToString

                l.AddNew(strX, strX)
            Next

        Next

        cb.DataSource = l
        cb.SelectedIndex = -1

    End Sub

    Function CalculateVolumeIndex(ByVal cb As ComboBox, ByVal rcb As RangeChangeButton) As Integer

        Dim intX As Integer = cb.SelectedIndex

        If intX < 0 Then
            intX = 0
        End If

        If intX = 0 AndAlso rcb.IndexChangeAmount > 0 Then
            intX = 1
        End If

        intX += rcb.IndexChangeAmount

        If intX < 0 Then
            intX = 0
        ElseIf intX > cb.Items.Count - 1 Then
            intX = cb.Items.Count - 1
        End If

        Return intX

    End Function

    Function CalculateRangeIndex(ByVal cb As ComboBox, ByVal rcb As RangeChangeButton) As Integer

        Dim intX As Integer = cb.SelectedIndex

        If intX < 0 Then
            intX = 0
        End If

        intX += rcb.IndexChangeAmount

        If intX < 0 Then
            intX = 0
        ElseIf intX > cb.Items.Count - 1 Then
            intX = cb.Items.Count - 1
        End If

        Return intX

    End Function

    Sub SetLastControlInfo(ByVal c As Control)

        Dim gb As GroupBox = Helper.GetParentGroupBoxForControl(c)

        If Not gb Is Nothing Then
            ToolStripStatusLabel1.Text = "Last Control Updated: " & c.Parent.Text & " @ " & GetNowString()
        End If

    End Sub

    Sub SetControlWaiting(ByVal c As Control)

        Dim gb As GroupBox = Helper.GetParentGroupBoxForControl(c)

        If Not My.CustomSettings.PreviewCommands.Instance.Value AndAlso RecordingScriptID = 0 Then

            If Not gb Is Nothing Then
                gb.ForeColor = Color.Red
            End If

            If Not lControlQueue.ContainsKey(gb.Name) Then
                lControlQueue.Add(gb.Name, Now)
            End If

        End If

    End Sub

    Sub ClearControlWaiting(ByVal c As Control)

        Dim gb As GroupBox = Helper.GetParentGroupBoxForControl(c)

        ClearControlWaiting(gb)

    End Sub

    Sub ClearControlWaiting(ByVal gb As GroupBox)

        If Not gb Is Nothing Then
            gb.ForeColor = System.Drawing.SystemColors.ControlText

            If lControlQueue.ContainsKey(gb.Name) Then
                lControlQueue.Remove(gb.Name)
            End If

        End If

    End Sub

    Function GetFixedListComboBox(ByVal m As FixedListComboBox) As ComboBox

        Dim cb As ComboBox

        Select Case m
            Case FixedListComboBox.DCOMP
                cb = ComboBox38
            Case FixedListComboBox.DRC
                cb = ComboBox37
            Case FixedListComboBox.Night
                cb = ComboBox40
            Case FixedListComboBox.Restorer
                cb = ComboBox41
            Case FixedListComboBox.RoomSize
                cb = ComboBox39
            Case FixedListComboBox.Resolution
                cb = ComboBox49
            Case FixedListComboBox.AspectRatio
                cb = ComboBox50
            Case FixedListComboBox.Zone2ChannelSetting
                cb = ComboBox34
            Case FixedListComboBox.Zone3ChannelSetting
                cb = ComboBox35
            Case FixedListComboBox.AudioInputMode
                cb = ComboBox14
            Case FixedListComboBox.DigitalAudioMode
                cb = ComboBox15
            Case FixedListComboBox.TunerFrequency
                cb = ComboBox53
            Case FixedListComboBox.SurroundMode
                cb = ComboBox54
            Case FixedListComboBox.SurroundBackMode
                cb = ComboBox55
            Case FixedListComboBox.SurroundSetting
                cb = ComboBox22
            Case FixedListComboBox.AudysseyMode
                cb = ComboBox56
            Case FixedListComboBox.DynamicVolume
                cb = ComboBox63
            Case FixedListComboBox.ReferenceLevel
                cb = ComboBox64
        End Select

        Return cb

    End Function

    Sub SetNetRefeshTimer()

        If CheckBox7.Checked AndAlso TabControl1.SelectedTab.Name = tabNet.Name Then
            Timer1.Enabled = False
            Timer1.Stop()
            Timer1.Interval = CInt(NumericUpDown6.Value * 1000)
            Timer1.Enabled = True
            Timer1.Start()
        Else
            Timer1.Enabled = False
            Timer1.Stop()
        End If

    End Sub

    Sub SetIPodRefeshTimer()

        If CheckBox9.Checked AndAlso TabControl1.SelectedTab.Name = tabIPod.Name Then
            Timer2.Enabled = False
            Timer2.Stop()
            Timer2.Interval = CInt(NumericUpDown8.Value * 1000)
            Timer2.Enabled = True
            Timer2.Start()
        Else
            Timer2.Enabled = False
            Timer2.Stop()
        End If

    End Sub

    Sub ClearNet()

        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox10.Clear()
        TextBox11.Clear()
        TextBox12.Clear()
        TextBox13.Clear()
        TextBox14.Clear()

        Button63.Enabled = False
        Button64.Enabled = False
        Button65.Enabled = False
        Button66.Enabled = False
        Button67.Enabled = False
        Button68.Enabled = False

        Label6.Visible = False
        Label7.Visible = False
        Label8.Visible = False
        Label9.Visible = False
        Label11.Visible = False
        Label12.Visible = False

        SetNetRefeshTimer()

    End Sub

    Sub ClearIPod()

        TextBox15.Clear()
        TextBox16.Clear()
        TextBox17.Clear()
        TextBox18.Clear()
        TextBox19.Clear()
        TextBox20.Clear()
        TextBox21.Clear()
        TextBox22.Clear()
        TextBox23.Clear()
        TextBox24.Clear()

        Label17.Visible = False
        Label18.Visible = False
        Label19.Visible = False
        Label20.Visible = False
        Label21.Visible = False
        Label22.Visible = False
        Label23.Visible = False

        SetIPodRefeshTimer()

    End Sub

#Region "Set UI"

    Public Sub SetSurroundSetting(ByVal v As String)

        booLoading = True

        Label14.Text = v

        booLoading = False

    End Sub

    Public Sub SetNETInfo(ByVal v As String, ByVal LineNumber As Integer, ByVal Playable As Boolean, ByVal CursorSelect As Boolean)

        booLoading = True

        Select Case LineNumber
            Case 0
                TextBox6.Text = v
            Case 1
                TextBox7.Text = v
                Button63.Enabled = False
                Label6.Visible = False

                If Playable Then
                    Button63.Enabled = True
                End If

                If CursorSelect Then
                    Label6.Visible = True
                End If

            Case 2
                TextBox8.Text = v
                Button64.Enabled = False
                Label7.Visible = False

                If Playable Then
                    Button64.Enabled = True
                End If

                If CursorSelect Then
                    Label7.Visible = True
                End If

            Case 3
                TextBox9.Text = v
                Button65.Enabled = False
                Label8.Visible = False

                If Playable Then
                    Button65.Enabled = True
                End If

                If CursorSelect Then
                    Label8.Visible = True
                End If

            Case 4
                TextBox10.Text = v
                Button66.Enabled = False
                Label9.Visible = False

                If Playable Then
                    Button66.Enabled = True
                End If

                If CursorSelect Then
                    Label9.Visible = True
                End If

            Case 5
                TextBox11.Text = v
                Button67.Enabled = False
                Label11.Visible = False

                If Playable Then
                    Button67.Enabled = True
                End If

                If CursorSelect Then
                    Label11.Visible = True
                End If

            Case 6
                TextBox12.Text = v
                Button68.Enabled = False
                Label12.Visible = False

                If Playable Then
                    Button68.Enabled = True
                End If

                If CursorSelect Then
                    Label12.Visible = True
                End If

            Case 7
                TextBox13.Text = v
            Case 8
                TextBox14.Text = v
        End Select

        booLoading = False

    End Sub

    Public Sub SetIPodInfo(ByVal v As String, ByVal LineNumber As Integer, ByVal Playable As Boolean, ByVal DisplayOnly As Boolean, ByVal CursorSelect As Boolean)

        booLoading = True

        Select Case LineNumber
            Case 0
                TextBox23.Text = v
            Case 1
                TextBox22.Text = v
                Label22.Visible = False

                If CursorSelect Then
                    Label22.Visible = True
                End If

            Case 2
                TextBox21.Text = v
                Label21.Visible = False

                If CursorSelect Then
                    Label21.Visible = True
                End If

            Case 3
                TextBox20.Text = v
                Label20.Visible = False

                If CursorSelect Then
                    Label20.Visible = True
                End If

            Case 4
                TextBox19.Text = v
                Label19.Visible = False

                If CursorSelect Then
                    Label19.Visible = True
                End If

            Case 5
                TextBox18.Text = v
                Label18.Visible = False

                If CursorSelect Then
                    Label18.Visible = True
                End If

            Case 6
                TextBox17.Text = v
                Label17.Visible = False

                If CursorSelect Then
                    Label17.Visible = True
                End If

            Case 7
                TextBox16.Text = v
                Label23.Visible = False

                If CursorSelect Then
                    Label23.Visible = True
                End If

            Case 8
                TextBox15.Text = v
            Case 9
                TextBox24.Text = v
        End Select

        booLoading = False

    End Sub

    Public Sub SetTunerPreset(ByVal v As String)

        booLoading = True

        Dim intX As Integer

        Select Case v
            Case "OFF"
                DataGridView1.ClearSelection()
                DataGridView1.CurrentCell = Nothing
            Case Else
                intX = ComboBox52.FindStringExact(v)

                If intX >= 0 Then
                    DataGridView1.BindingContext(lTunerPresets).Position = intX
                End If

        End Select

        SetLastControlInfo(DataGridView1)
        ClearControlWaiting(DataGridView1)

        booLoading = False

    End Sub

    Public Sub SetFixedListComboBox(ByVal v As String, ByVal m As FixedListComboBox)

        booLoading = True

        Dim cb As ComboBox

        cb = GetFixedListComboBox(m)

        cb.SelectedValue = v

        SetLastControlInfo(cb)
        ClearControlWaiting(cb)

        booLoading = False

    End Sub

    Public Sub SetRange(ByVal dt As RangeControlType, ByVal db As DenonRangeControl)

        booLoading = True

        Dim cb As ComboBox

        Select Case dt
            Case RangeControlType.Bass
                cb = ComboBox16
            Case RangeControlType.Treble
                cb = ComboBox17
            Case RangeControlType.LFE
                cb = ComboBox18
            Case RangeControlType.Dimension
                cb = ComboBox19
            Case RangeControlType.CenterWidth
                cb = ComboBox20
            Case RangeControlType.CenterImage
                cb = ComboBox21
            Case RangeControlType.Zone2Bass
                cb = ComboBox26
            Case RangeControlType.Zone2Treble
                cb = ComboBox25
            Case RangeControlType.Zone3Bass
                cb = ComboBox28
            Case RangeControlType.Zone3Treble
                cb = ComboBox27
            Case RangeControlType.DSPSurroundDelay
                cb = ComboBox43
            Case RangeControlType.AudioDelay
                cb = ComboBox44
            Case RangeControlType.Contrast
                cb = ComboBox45
            Case RangeControlType.Brightness
                cb = ComboBox46
            Case RangeControlType.Chroma
                cb = ComboBox47
            Case RangeControlType.Hue
                cb = ComboBox48
            Case RangeControlType.DSPEffectLevel
                cb = ComboBox42
        End Select

        If Not IsNothing(db) Then
            cb.SelectedItem = db
        End If

        SetLastControlInfo(cb)
        ClearControlWaiting(cb)

        booLoading = False
    End Sub

    Public Sub SetOnOff(ByVal cmd As OnOffCommand)

        Dim rb, rb2 As RadioButton
        Dim cb, cb2 As CheckBox

        booLoading = True

        Select Case cmd
            Case OnOffCommand.DynamicEQOff
                rb = RadioButton41
            Case OnOffCommand.DynamicEQOn
                rb = RadioButton40
            Case OnOffCommand.ToneCtrlOff
                rb = RadioButton4
            Case OnOffCommand.ToneCtrlOn
                rb = RadioButton3
            Case OnOffCommand.PowerOff
                rb = RadioButton1
                rb2 = RadioButton33
            Case OnOffCommand.PowerOn
                rb = RadioButton2
                rb2 = RadioButton32
            Case OnOffCommand.CinemaEQOff
                rb = RadioButton6
            Case OnOffCommand.CinemaEQOn
                rb = RadioButton5
            Case OnOffCommand.AFDOff
                rb = RadioButton8
            Case OnOffCommand.AFDOn
                rb = RadioButton7
            Case OnOffCommand.SWAttOff
                rb = RadioButton10
            Case OnOffCommand.SWAttOn
                rb = RadioButton9
            Case OnOffCommand.SWOff
                rb = RadioButton12
            Case OnOffCommand.SWOn
                rb = RadioButton11
            Case OnOffCommand.PanoramaOff
                rb = RadioButton14
            Case OnOffCommand.PanoramaOn
                rb = RadioButton13
            Case OnOffCommand.ZoneMainOff
                rb = RadioButton16
                rb2 = RadioButton35
            Case OnOffCommand.ZoneMainOn
                rb = RadioButton15
                rb2 = RadioButton34
            Case OnOffCommand.Zone2Off
                rb = RadioButton18
                rb2 = RadioButton37
            Case OnOffCommand.Zone2On
                rb = RadioButton17
                rb2 = RadioButton36
            Case OnOffCommand.Zone3Off
                rb = RadioButton20
                rb2 = RadioButton39
            Case OnOffCommand.Zone3On
                rb = RadioButton19
                rb2 = RadioButton38
            Case OnOffCommand.DSPEffectOff
                rb = RadioButton22
            Case OnOffCommand.DSPEffectOn
                rb = RadioButton21
            Case OnOffCommand.Zone2HPFOff
                rb = RadioButton24
            Case OnOffCommand.Zone2HPFOn
                rb = RadioButton23
            Case OnOffCommand.Zone3HPFOff
                rb = RadioButton26
            Case OnOffCommand.Zone3HPFOn
                rb = RadioButton25
            Case OnOffCommand.Zone2MuteOff, OnOffCommand.Zone2MuteOn
                cb = CheckBox5
                cb2 = CheckBox11
            Case OnOffCommand.Zone3MuteOff, OnOffCommand.Zone3MuteOn
                cb = CheckBox6
                cb2 = CheckBox12
            Case OnOffCommand.MuteOff, OnOffCommand.MuteOn
                cb = CheckBox1
                cb2 = CheckBox10
        End Select

        If Not rb Is Nothing Then
            rb.Checked = True

            If cmd = OnOffCommand.DSPEffectOff Then
                RangeChangeButton29.Enabled = False
                RangeChangeButton30.Enabled = False
                ComboBox42.Enabled = False
            ElseIf cmd = OnOffCommand.DSPEffectOn Then
                RangeChangeButton29.Enabled = True
                RangeChangeButton30.Enabled = True
                ComboBox42.Enabled = True
            End If

            SetLastControlInfo(rb)
            ClearControlWaiting(rb)

            If Not rb2 Is Nothing Then
                rb2.Checked = True
                SetLastControlInfo(rb2)
                ClearControlWaiting(rb2)
            End If

        ElseIf Not cb Is Nothing Then

            Select Case cmd
                Case OnOffCommand.Zone2MuteOff, OnOffCommand.Zone3MuteOff, OnOffCommand.MuteOff
                    cb.Checked = False
                    cb2.Checked = False
                Case OnOffCommand.Zone2MuteOn, OnOffCommand.Zone3MuteOn, OnOffCommand.MuteOn
                    cb.Checked = True
                    cb2.Checked = True
            End Select

            SetLastControlInfo(cb)
            ClearControlWaiting(cb)
            SetLastControlInfo(cb2)
            ClearControlWaiting(cb2)

        End If

        booLoading = False

    End Sub

    Public Sub SetRemoteAndPanelLock(ByVal cmd As RemoteAndPanelLock)

        Dim rb As RadioButton

        booLoading = True

        Select Case cmd
            Case RemoteAndPanelLock.PanelFullLock
                rb = RadioButton29
            Case RemoteAndPanelLock.PanelFullLockExceptMV
                rb = RadioButton31
            Case RemoteAndPanelLock.PanelUnlock
                rb = RadioButton30
            Case RemoteAndPanelLock.RemoteLock
                rb = RadioButton27
            Case RemoteAndPanelLock.RemoteUnlock
                rb = RadioButton28
        End Select

        rb.Checked = True

        SetLastControlInfo(rb)
        ClearControlWaiting(rb)

        booLoading = False

    End Sub

    Public Sub SetMasterVolume(ByVal v As MasterVolume, ByVal z As Zones)

        booLoading = True

        Dim cb, cb2 As ComboBox

        Select Case z
            Case Zones.ZoneMain
                cb = ComboBox2
                cb2 = ComboBox58
            Case Zones.Zone2
                cb = ComboBox23
                cb2 = ComboBox60
            Case Zones.Zone3
                cb = ComboBox24
                cb2 = ComboBox62
        End Select

        If Not IsNothing(v) Then
            cb.SelectedItem = v
            SetLastControlInfo(cb)
            ClearControlWaiting(cb)
            cb2.SelectedItem = v
            SetLastControlInfo(cb2)
            ClearControlWaiting(cb2)
        End If

        booLoading = False

    End Sub

    Public Sub SetChannelVolume(ByVal c As Channel, ByVal v As ChannelVolume)

        Dim cb As ComboBox

        Select Case c
            Case Channel.LeftFront
                cb = ComboBox4
            Case Channel.Center
                cb = ComboBox5
            Case Channel.RightFront
                cb = ComboBox6
            Case Channel.LeftSurround
                cb = ComboBox7
            Case Channel.RightSurround
                cb = ComboBox8
            Case Channel.LeftRear
                cb = ComboBox9
            Case Channel.RightRear
                cb = ComboBox10
            Case Channel.LeftHeight
                cb = ComboBox67
            Case Channel.RightHeight
                cb = ComboBox66
            Case Channel.LeftWide
                cb = ComboBox68
            Case Channel.RightWide
                cb = ComboBox69
            Case Channel.Subwoofer
                cb = ComboBox11
            Case Channel.Zone2Left
                cb = ComboBox30
            Case Channel.Zone2Right
                cb = ComboBox29
            Case Channel.Zone3Left
                cb = ComboBox32
            Case Channel.Zone3Right
                cb = ComboBox31
        End Select

        booLoading = True

        If Not IsNothing(v) Then
            cb.SelectedItem = v
            SetLastControlInfo(cb)
            ClearControlWaiting(cb)
        End If

        booLoading = False

    End Sub

    Public Sub SetSource(ByVal v As Source)

        booLoading = True

        If Not IsNothing(v) Then
            ComboBox3.SelectedItem = v
            SetLastControlInfo(ComboBox3)
            ClearControlWaiting(ComboBox3)
            ComboBox57.SelectedItem = v
            SetLastControlInfo(ComboBox57)
            ClearControlWaiting(ComboBox57)
        End If

        booLoading = False

    End Sub

    Public Sub SetSourceName(ByVal v As Source)

        booLoading = True

        Dim intX As Integer

        If Not IsNothing(v) Then

            For Each cb As ComboBox In lSources

                intX = cb.Items.IndexOf(v)

                If intX >= 0 Then
                    cb.Items(intX) = v
                End If

            Next

        End If

        booLoading = False

    End Sub

    Public Sub SetSourceStatus(ByVal v As Source)

        booLoading = True

        Dim intX As Integer

        If Not IsNothing(v) Then

            For Each cb As ComboBox In lSources

                intX = cb.Items.IndexOf(v)

                If intX >= 0 Then

                    If Not v.IsActive Then
                        cb.Items.Remove(v)
                    End If

                End If

            Next

        End If

        booLoading = False

    End Sub

    Public Sub SetTunerPreset(ByVal v As TunerPreset)

        booLoading = True

        Dim intX As Integer

        intX = lTunerPresets.IndexOf(v)

        If intX >= 0 Then
            lTunerPresets.RemoveAt(intX)
        End If

        lTunerPresets.Add(v)

        booLoading = False

    End Sub

    Public Sub SetZoneSource(ByVal z As Zones, ByVal v As Source)

        booLoading = True

        Dim cb, cb2 As ComboBox

        Select Case z
            Case Zones.Zone2
                cb = ComboBox33
                cb2 = ComboBox59
            Case Zones.Zone3
                cb = ComboBox36
                cb2 = ComboBox61
        End Select

        If Not IsNothing(v) Then
            cb.SelectedItem = v
            SetLastControlInfo(cb)
            ClearControlWaiting(cb)
            cb2.SelectedItem = v
            SetLastControlInfo(cb2)
            ClearControlWaiting(cb2)
        End If

        booLoading = False

    End Sub

    Public Sub SetVideoSelect(ByVal v As Source)

        booLoading = True

        If Not IsNothing(v) Then
            ComboBox12.SelectedItem = v
            SetLastControlInfo(ComboBox12)
            ClearControlWaiting(ComboBox12)
        End If

        booLoading = False

    End Sub

#End Region

#Region "Global Error Handler"

	Public Sub GlobalErrorHandler(ByVal sender As Object, ByVal e As System.Threading.ThreadExceptionEventArgs)
		Dim strX As String = "I am sorry. A serious error has occured." & vbCrLf
		Dim strY As String

		Try
			TabControl1.SelectTab(TabControl1.TabCount - 2)

			strX &= vbCrLf & vbCrLf
			strX &= "For your convenience, the exception information is output below. Please contact technical support. This application will now close."
			strX &= vbCrLf & vbCrLf
			strX &= e.Exception.ToString
			strX &= vbCrLf & vbCrLf
			strX &= "Before the application closes, would you like to copy the exception info to your clipboard?"

			If MsgBox(strX, MsgBoxStyle.Critical Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2) = MsgBoxResult.Yes Then
				My.Computer.Clipboard.SetText(e.Exception.ToString)
			End If

		Catch ex As Exception
			strY = "An unhandled error occured that was caught by the application's global error handler. However, during processing of that unhandled error, the global error handler itself encountered an unhandled exception. Here is what I can tell you. " & vbCrLf & vbCrLf & "strX = " & strX & vbCrLf & vbCrLf & "e.Exception.ToString = " & e.Exception.ToString & vbCrLf & vbCrLf & "ex.ToString = " & ex.ToString
			If MsgBox(strX, MsgBoxStyle.Critical Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2) = MsgBoxResult.Yes Then
				My.Computer.Clipboard.SetText(e.Exception.ToString & vbCrLf & vbCrLf & vbCrLf & ex.ToString)
			End If
		Finally
			Me.Close()
		End Try

	End Sub

	Public Sub New()

		' This call is required by the Windows Form Designer.
		InitializeComponent()

		' Add any initialization after the InitializeComponent() call.
		AddHandler System.Windows.Forms.Application.ThreadException, AddressOf GlobalErrorHandler
		SplashScreen1.SetStatus("Initializing UI")

		Me.Icon = My.Resources.denon
		Me.NotifyIcon1.Icon = My.Resources.denon
		Me.NotifyIcon2.Icon = My.Resources.denon
		AboutBox1.Icon = My.Resources.denon

		Dim di As New IO.DirectoryInfo(Application.UserAppDataPath)

		AppDomain.CurrentDomain.SetData("DataDirectory", di.Parent.FullName)

	End Sub

#End Region

    Private Sub iConfigureFixedListComboBoxes(fixedListComboBox As FixedListComboBox, p2 As Integer, p3 As String)
        Throw New NotImplementedException
    End Sub

End Class
