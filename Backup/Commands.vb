''''''''' ''''''''' ''''''''' ''''''''' ''''''''' 
''''''''' Copyright 2007-2009 Brian Saville
''''''''' Command3808 is free for non-commercial use.
''''''''' Reuse of the source code is free for non-commercial use
''''''''' provided that the source code used is credited to Command3808.
''''''''' ''''''''' ''''''''' ''''''''' '''''''''
Imports System.Text.RegularExpressions
Public Class Commands

	Private Shared Sub EnqueueRequest(ByVal Command As String, ByVal TimeOut As Integer)

		If Form1.RecordingScriptID <> 0 Then
			DB.ScriptEntry.Create(Form1.RecordingScriptID, Command, Command)
			Form1.WriteDebug(Helper.GetNowString & " - RECORDED: " & Command, True)
			Form1.FillScripts(Form1.RecordingScriptName, Nothing, Nothing)
		ElseIf Form1.CheckBox2.Checked Then
			Form1.WriteDebug(Helper.GetNowString & " - PREVIEW: " & Command, True)
		Else
			RaiseEvent RequestReady(New ReadWriteKeyValuePair(Of String, Integer)(Command, TimeOut))
		End If

	End Sub

	Public Shared Event RequestReady(ByVal kvp As ReadWriteKeyValuePair(Of String, Integer))

#Region "Query"

	Public Shared Sub QueryThis(ByVal CommandWithoutQuestionMark As String)
		EnqueueRequest(CommandWithoutQuestionMark & "?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

	Public Shared Sub QueryRange(ByVal db As RangeControlType)

		EnqueueRequest(Helper.GetCommandPrefixForRangeControl(db) & "?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)

	End Sub

	Public Shared Sub QueryRemoteAndPanelLock()
		EnqueueRequest("SYPANEL LOCK ?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
		EnqueueRequest("SYREMOTE LOCK ?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

	Public Shared Sub QueryOnOff(ByVal cmd As OnOffCommand)

		EnqueueRequest(Helper.GetCommandPrefixForOnOff(cmd) & "?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)

	End Sub

	Public Shared Sub QueryMasterVolume()
		EnqueueRequest("MV?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
		EnqueueRequest("MU?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

	Public Shared Sub QueryChannelVolume(ByVal c As Channel)

		Select Case c
			Case Channel.LeftFront
				EnqueueRequest("CVFL?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.Center
				EnqueueRequest("CVC?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.RightFront
				EnqueueRequest("CVFR?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.LeftSurround
				EnqueueRequest("CVSL?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.RightSurround
				EnqueueRequest("CVSR?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.LeftRear
				EnqueueRequest("CVSBL?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.RightRear
				EnqueueRequest("CVSBR?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.Subwoofer
				EnqueueRequest("CVSW?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.Zone2Left, Channel.Zone2Right
				EnqueueRequest(Helper.GetZonePrefix(Zones.Zone2) & "CV?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.Zone3Left, Channel.Zone3Right
				EnqueueRequest(Helper.GetZonePrefix(Zones.Zone3) & "CV?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
		End Select

	End Sub

	Public Shared Sub QuerySource()
		EnqueueRequest("SI?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

	Public Shared Sub QueryVideoSelect()
		EnqueueRequest("SV?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

	Public Shared Sub QueryRecordSelect()
		EnqueueRequest("SR?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

	Public Shared Sub QueryAudioInputMode()
		EnqueueRequest("SD?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

	Public Shared Sub QuerySurroundSettings()
		EnqueueRequest("MS?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value) ' main tab
		EnqueueRequest("PSSB: ?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value) ' surround tab
		EnqueueRequest("PSMODE: ?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)	' surround tab
	End Sub

	Public Shared Sub QueryRoomEQ()
		EnqueueRequest("PSROOM EQ: ?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)	' surround tab
	End Sub

	Public Shared Sub QueryDigitalAudioMode()
		EnqueueRequest("DC?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

	Public Shared Sub QueryZoneChannelSetting(ByVal z As Zones)
		EnqueueRequest(Helper.GetZonePrefix(z) & "CS?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

	Public Shared Sub QuerySourceNames()
		EnqueueRequest("SSFUN ?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

	Public Shared Sub QuerySourceStatus()
		EnqueueRequest("SSSOD ?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

	Public Shared Sub QueryPresets()
		EnqueueRequest("SSTPN ?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
		EnqueueRequest("SSXPN ?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

	Public Shared Sub QueryTuner()
		EnqueueRequest("TFAN?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
		EnqueueRequest("TPAN?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
		EnqueueRequest("TMAN?", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

	Public Shared Sub QueryNet()
		EnqueueRequest("NSE", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

	Public Shared Sub QueryIPod()
		EnqueueRequest("IPE", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

#End Region

#Region "Command"

	Public Shared Sub SendIPodCommand(ByVal ipc As IPodCommand)

		Dim strCmd As String = "IP"

		Select Case ipc
			Case IPodCommand.Down
				strCmd &= "91"
			Case IPodCommand.EnterPlayPause
				strCmd &= "94"
			Case IPodCommand.Left
				strCmd &= "92"
			Case IPodCommand.PageDown
				strCmd &= "9X" ' documentation says 9Y
			Case IPodCommand.PageUp
				strCmd &= "9Y" ' documentation says 9X
			Case IPodCommand.PlayPause
				strCmd &= "9A"
			Case IPodCommand.Right
				strCmd &= "93"
			Case IPodCommand.Stop
				strCmd &= "9C"
			Case IPodCommand.Up
				strCmd &= "90"
		End Select

		EnqueueRequest(strCmd, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)

	End Sub

	Public Shared Sub SendNETCommand(ByVal nc As NETCommand)

		Dim strCmd As String = "NS"

		Select Case nc
			Case NETCommand.Down
				strCmd &= "91"
			Case NETCommand.EnterPlayPause
				strCmd &= "94"
			Case NETCommand.Left
				strCmd &= "92"
			Case NETCommand.PageDown
				strCmd &= "9X" ' documentation says 9Y
			Case NETCommand.PageUp
				strCmd &= "9Y" ' documentation says 9X
			Case NETCommand.Pause
				strCmd &= "9B"
			Case NETCommand.Play
				strCmd &= "9A"
			Case NETCommand.Right
				strCmd &= "93"
			Case NETCommand.Stop
				strCmd &= "9C"
			Case NETCommand.Up
				strCmd &= "90"
		End Select

		EnqueueRequest(strCmd, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)

	End Sub

	Public Shared Sub SendMenuCommand(ByVal mc As MenuCommand)

		Dim strCmd As String = "MN"

		Select Case mc
			Case MenuCommand.Down
				strCmd &= "CDN"
			Case MenuCommand.Enter
				strCmd &= "ENT"
			Case MenuCommand.GUIOff
				strCmd &= "MEN OFF"
			Case MenuCommand.GUIOn
				strCmd &= "MEN ON"
			Case MenuCommand.Left
				strCmd &= "CLT"
			Case MenuCommand.Return
				strCmd &= "RTN"
			Case MenuCommand.Right
				strCmd &= "CRT"
			Case MenuCommand.Up
				strCmd &= "CUP"
		End Select

		EnqueueRequest(strCmd, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)

	End Sub

	Public Shared Sub SendRange(ByVal dt As RangeControlType, ByVal db As DenonRangeControl)

		EnqueueRequest(Helper.GetCommandPrefixForRangeControl(dt) & db.DenonCommandValue, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)

	End Sub

	Public Shared Sub SendOnOff(ByVal cmd As OnOffCommand)

		Select Case CInt(cmd)
			Case 101 To 200
				EnqueueRequest(Helper.GetCommandPrefixForOnOff(cmd) & "ON", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case 201 To 299
				Select Case cmd
					Case OnOffCommand.PowerOff
						EnqueueRequest(Helper.GetCommandPrefixForOnOff(cmd) & "STANDBY", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
					Case Else
						EnqueueRequest(Helper.GetCommandPrefixForOnOff(cmd) & "OFF", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
				End Select

		End Select

	End Sub

	Public Shared Sub SendRemoteAndPanelLock(ByVal cmd As RemoteAndPanelLock)

		Select Case cmd
			Case RemoteAndPanelLock.PanelFullLock
				EnqueueRequest("SYPANEL+V LOCK ON", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case RemoteAndPanelLock.PanelFullLockExceptMV
				EnqueueRequest("SYPANEL LOCK ON", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case RemoteAndPanelLock.PanelUnlock
				EnqueueRequest("SYPANEL LOCK OFF", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case RemoteAndPanelLock.RemoteLock
				EnqueueRequest("SYREMOTE LOCK ON", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case RemoteAndPanelLock.RemoteUnlock
				EnqueueRequest("SYREMOTE LOCK OFF", My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
		End Select

	End Sub

	Public Shared Sub SendThis(ByVal m As String)
		SendCommand(m, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

	Public Shared Sub SendCommand(ByVal m As String, ByVal TimeOut As Integer)
		EnqueueRequest(m, TimeOut)
	End Sub

	Public Shared Sub SendMasterVolume(ByVal v As MasterVolume)
		EnqueueRequest("MV" & v.DenonCommandValue, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

	Public Shared Sub SendZoneVolume(ByVal v As MasterVolume, ByVal z As Zones)
		EnqueueRequest(Helper.GetZonePrefix(z) & v.DenonCommandValue, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

	Public Shared Sub SendChannelVolume(ByVal c As Channel, ByVal v As ChannelVolume)

		Select Case c
			Case Channel.LeftFront
				EnqueueRequest("CVFL " & v.DenonCommandValue, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.Center
				EnqueueRequest("CVC " & v.DenonCommandValue, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.RightFront
				EnqueueRequest("CVFR " & v.DenonCommandValue, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.LeftSurround
				EnqueueRequest("CVSL " & v.DenonCommandValue, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.RightSurround
				EnqueueRequest("CVSR " & v.DenonCommandValue, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.LeftRear
				EnqueueRequest("CVSBL " & v.DenonCommandValue, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.RightRear
				EnqueueRequest("CVSBR " & v.DenonCommandValue, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.Subwoofer
				EnqueueRequest("CVSW " & v.DenonCommandValue, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.Zone2Left
				EnqueueRequest(Helper.GetZonePrefix(Zones.Zone2) & "CVFL " & v.DenonCommandValue, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.Zone2Right
				EnqueueRequest(Helper.GetZonePrefix(Zones.Zone2) & "CVFR " & v.DenonCommandValue, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.Zone3Left
				EnqueueRequest(Helper.GetZonePrefix(Zones.Zone3) & "CVFL " & v.DenonCommandValue, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
			Case Channel.Zone3Right
				EnqueueRequest(Helper.GetZonePrefix(Zones.Zone3) & "CVFR " & v.DenonCommandValue, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
		End Select

	End Sub

	Public Shared Sub SendSource(ByVal v As Source)
		EnqueueRequest("SI" & v.DefaultSourceName, 4000)
	End Sub

	Public Shared Sub SendZoneSource(ByVal z As Zones, ByVal v As Source)
		EnqueueRequest(Helper.GetZonePrefix(z) & v.DefaultSourceName, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

	Public Shared Sub SendVideoSelect(ByVal v As Source)
		EnqueueRequest("SV" & v.DefaultSourceName, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

	Public Shared Sub SendRecordSelect(ByVal v As Source)
		EnqueueRequest("SR" & v.DefaultSourceName, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
	End Sub

#End Region

#Region "Parse"

	Public Shared Sub ParseResponse(ByVal l As Generic.List(Of String))

		Dim kvp As ReadWriteKeyValuePair(Of Channel, ChannelVolume)

		For Each strX As String In l

			If Not String.IsNullOrEmpty(strX) Then

				Select Case strX.Substring(0, 2)
					Case "PW"
						Form1.SetOnOff(ParsePowerResponse(strX))
					Case "MV"
						Form1.SetMasterVolume(ParseMasterVolumeResponse(strX), Zones.ZoneMain)
					Case "MU"
						Form1.SetOnOff(ParseMuteResponse(strX))
					Case "SI"
						Form1.SetSource(ParseSourceResponse(strX))
					Case "CV"
						kvp = ParseChannelVolumeResponse(strX)

						If Not IsNothing(kvp) AndAlso kvp.Key <> Channel.Undefined Then
							Form1.SetChannelVolume(kvp.Key, kvp.Value)
						End If
					Case "SV"
						Form1.SetVideoSelect(ParseVideoSelectResponse(strX))
					Case "SR"
						Form1.SetRecordSelect(ParseRecordSelectResponse(strX))
						Form1.SetZoneSource(Zones.Zone2, ParseRecordSelectResponse(strX))
					Case "SD"
						Form1.SetFixedListComboBox(ParseAudioInputModeResponse(strX), FixedListComboBox.AudioInputMode)
					Case "DC"
						Form1.SetFixedListComboBox(ParseDigitalAudioModeResponse(strX), FixedListComboBox.DigitalAudioMode)
					Case "PS"
						ParsePSResponse(strX)
					Case "MS"
						ParseMSResponse(strX)
					Case "ZM"
						ParseZoneMain(strX)
					Case "Z2"
						ParseZone2(strX)
					Case "Z3"
						ParseZone3(strX)
					Case "SS"
						ParseSS(strX)
					Case "PV"
						ParsePVResponse(strX)
					Case "VS"
						ParseVSResponse(strX)
					Case "TF"
						Form1.SetFixedListComboBox(strX.Substring(4), FixedListComboBox.TunerFrequency)
					Case "TP"
						Form1.SetTunerPreset(strX.Substring(4))
					Case "NS"
						ParseNSResponse(strX)
					Case "IP"
						ParseIPResponse(strX)
					Case "SY"
						ParseSYResponse(strX)
				End Select

			End If

		Next

	End Sub

	Public Shared Sub ParseSYResponse(ByVal m As String)

		Select Case m.Substring(2)
			Case "REMOTE LOCK ON"
				Form1.SetRemoteAndPanelLock(RemoteAndPanelLock.RemoteLock)
			Case "REMOTE LOCK OFF"
				Form1.SetRemoteAndPanelLock(RemoteAndPanelLock.RemoteUnlock)
			Case "PANEL LOCK ON"
				Form1.SetRemoteAndPanelLock(RemoteAndPanelLock.PanelFullLockExceptMV)
			Case "PANEL+V LOCK ON"
				Form1.SetRemoteAndPanelLock(RemoteAndPanelLock.PanelFullLock)
			Case "PANEL LOCK OFF"
				Form1.SetRemoteAndPanelLock(RemoteAndPanelLock.PanelUnlock)
		End Select

	End Sub

	Public Shared Sub ParseNSResponse(ByVal m As String)

		Dim intX, intY As Integer
		Dim rm As Match
		Dim booPlayable, booCursorSelect As Boolean
		Dim l As Generic.List(Of Byte)

		rm = Regex.Match(m, "NSE(?<1>\d)")

		If rm.Success Then

			intX = CInt(rm.Groups(1).Value)

			l = New Generic.List(Of Byte)(System.Text.ASCIIEncoding.ASCII.GetBytes(m.Substring(4)))

			Select Case intX
				Case 0, 7, 8

					intY = l.IndexOf(0)

					If l.IndexOf(0) >= 0 Then
						l.RemoveRange(intY, l.Count - intY)
					End If

				Case Else
					' Bit 1 (decimal 1) = playable
					' Bit 4 (decimal 8) = cursor select

					If (l(0) And 1) = 1 Then
						booPlayable = True
					End If

					If (l(0) And 8) = 8 Then
						booCursorSelect = True
					End If

					l.RemoveAt(0)

					intY = l.IndexOf(0)

					If l.IndexOf(0) >= 0 Then
						l.RemoveRange(intY, l.Count - intY)
					End If

			End Select

			l.RemoveAll(AddressOf Helper.SystemPredicateRemoveNonAsciiBytes)

			Form1.SetNETInfo(System.Text.ASCIIEncoding.ASCII.GetString(l.ToArray).Trim, intX, booPlayable, booCursorSelect)

		End If

	End Sub

	Public Shared Sub ParseIPResponse(ByVal m As String)

		Dim intX, intY As Integer
		Dim rm As Match
		Dim booPlayable, booDisplayOnly, booCursorSelect As Boolean
		Dim l As Generic.List(Of Byte)

		rm = Regex.Match(m, "IPE(?<1>\d)")

		If rm.Success Then

			intX = CInt(rm.Groups(1).Value)

			l = New Generic.List(Of Byte)(System.Text.ASCIIEncoding.ASCII.GetBytes(m.Substring(4)))

			Select Case intX
				Case 0, 8, 9

					intY = l.IndexOf(0)

					If l.IndexOf(0) >= 0 Then
						l.RemoveRange(intY, l.Count - intY)
					End If

				Case Else
					' Bit 1 (decimal 1) = playable
					' Bit 2 (decimal 2) = display only
					' Bit 4 (decimal 8) = cursor select

					If (l(0) And 1) = 1 Then
						booPlayable = True
					End If

					If (l(0) And 2) = 2 Then
						booDisplayOnly = True
					End If

					If (l(0) And 8) = 8 Then
						booCursorSelect = True
					End If

					l.RemoveAt(0)

					intY = l.IndexOf(0)

					If l.IndexOf(0) >= 0 Then
						l.RemoveRange(intY, l.Count - intY)
					End If

			End Select

			l.RemoveAll(AddressOf Helper.SystemPredicateRemoveNonAsciiBytes)

			Form1.SetIPodInfo(System.Text.ASCIIEncoding.ASCII.GetString(l.ToArray).Trim, intX, booPlayable, booDisplayOnly, booCursorSelect)

		End If

	End Sub

	Public Shared Sub ParseVSResponse(ByVal m As String)

		Select Case m.Substring(2, 2)
			Case "AS"
				Form1.SetFixedListComboBox(m.Substring(5), FixedListComboBox.AspectRatio)
			Case "SC"
				Form1.SetFixedListComboBox(m.Substring(4), FixedListComboBox.Resolution)
		End Select

	End Sub

	Public Shared Sub ParsePVResponse(ByVal m As String)

		Select Case m.Substring(2, 2)
			Case "CN"
				Form1.SetRange(RangeControlType.Contrast, DenonRangeControl.CreateFromDenonCommandValue(m.Substring(5), 50, 1, False, "dB", RangeControlType.Contrast))
			Case "BR"
				Form1.SetRange(RangeControlType.Brightness, DenonRangeControl.CreateFromDenonCommandValue(m.Substring(5), 0, 1, False, "dB", RangeControlType.Brightness))
			Case "CM"
				Form1.SetRange(RangeControlType.Chroma, DenonRangeControl.CreateFromDenonCommandValue(m.Substring(5), 50, 1, False, "dB", RangeControlType.Chroma))
			Case "HU"
				Form1.SetRange(RangeControlType.Hue, DenonRangeControl.CreateFromDenonCommandValue(m.Substring(6), 50, 1, False, "dB", RangeControlType.Hue))
		End Select

	End Sub

	Public Shared Sub ParseSS(ByVal m As String)

		Select Case m.Substring(2, 3)
			Case "FUN"
				ParseSourceName(m)
			Case "SOD"
				ParseSourceStatus(m)
			Case "TPN"
				ParseTunerPreset(m)
		End Select

	End Sub

	Public Shared Sub ParseTunerPreset(ByVal m As String)

		Form1.SetTunerPreset(TunerPreset.ParseFromSSTPNString(m.Substring(5)))

	End Sub

	Public Shared Sub ParseSourceName(ByVal m As String)

		Dim rm As Match
		Dim s As Source
		Dim strX As String

		rm = Regex.Match(m, "SSFUN(?<1>.*?)\x20(?<2>.*)")

		If rm.Success Then

			If Form1.lSourceInfo.ContainsKey(rm.Groups(1).Value) Then

				s = DirectCast(Form1.lSourceInfo(rm.Groups(1).Value), Source)

				strX = rm.Groups(2).Value.Trim(" "c)

				If s.DefaultSourceName <> strX Then
					s.CustomSourceName = strX
				Else
					s.CustomSourceName = Nothing
				End If

				Form1.SetSourceName(s)

			End If

		End If

	End Sub

	Public Shared Sub ParseSourceStatus(ByVal m As String)

		Dim rm As Match
		Dim s As Source
		Dim strX As String

		rm = Regex.Match(m, "SSSOD(?<1>.*?)\x20(?<2>.*)")

		If rm.Success Then

			If Form1.lSourceInfo.ContainsKey(rm.Groups(1).Value) Then

				s = DirectCast(Form1.lSourceInfo(rm.Groups(1).Value), Source)

				strX = rm.Groups(2).Value.Trim(" "c)

				If strX = "USE" Then
					s.IsActive = True
				ElseIf strX = "DEL" Then
					s.IsActive = False
				End If

				Form1.SetSourceStatus(s)

			End If

		End If

	End Sub

	Public Shared Sub ParseZoneMain(ByVal m As String)

		Select Case m
			Case Helper.GetCommandPrefixForOnOff(OnOffCommand.ZoneMainOff) & "OFF"
				Form1.SetOnOff(OnOffCommand.ZoneMainOff)
			Case Helper.GetCommandPrefixForOnOff(OnOffCommand.ZoneMainOn) & "ON"
				Form1.SetOnOff(OnOffCommand.ZoneMainOn)
		End Select

	End Sub

	Public Shared Sub ParseZone2(ByVal m As String)

		Dim strX As String = Helper.GetZonePrefix(Zones.Zone2)
		Dim rx As New Regex(strX & ".*?(?:OFF|ON)")
		Dim rx2 As New Regex(strX & "\d\d")
		Dim booProcessed As Boolean

		If rx.IsMatch(m) Then

			Select Case m
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.Zone2Off) & "OFF"
					Form1.SetOnOff(OnOffCommand.Zone2Off)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.Zone2On) & "ON"
					Form1.SetOnOff(OnOffCommand.Zone2On)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.Zone2MuteOff) & "OFF"
					Form1.SetOnOff(OnOffCommand.Zone2MuteOff)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.Zone2MuteOn) & "ON"
					Form1.SetOnOff(OnOffCommand.Zone2MuteOn)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.Zone2HPFOff) & "OFF"
					Form1.SetOnOff(OnOffCommand.Zone2HPFOff)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.Zone2HPFOn) & "ON"
					Form1.SetOnOff(OnOffCommand.Zone2HPFOn)
					booProcessed = True

			End Select
		End If

		If Not booProcessed AndAlso rx2.IsMatch(m) Then
			Form1.SetMasterVolume(MasterVolume.CreateFromDenonCommandValue(m.Substring(2), -70), Zones.Zone2)
			booProcessed = True
		End If

		If Not booProcessed AndAlso m.Length >= 4 Then

			Select Case m.Substring(0, 4)
				Case strX & "CS"
					Form1.SetFixedListComboBox(m.Substring(4), FixedListComboBox.Zone2ChannelSetting)
					booProcessed = True
				Case strX & "CV"

					If m.Substring(4, 2) = "FL" Then
						Form1.SetChannelVolume(Channel.Zone2Left, ChannelVolume.CreateFromDenonCommandValue(m.Substring(7)))
						booProcessed = True
					ElseIf m.Substring(4, 2) = "FR" Then
						Form1.SetChannelVolume(Channel.Zone2Right, ChannelVolume.CreateFromDenonCommandValue(m.Substring(7)))
						booProcessed = True
					End If

				Case strX & "PS"

					If m.Substring(0, 8) = Helper.GetCommandPrefixForRangeControl(RangeControlType.Zone2Bass) Then
						Form1.SetRange(RangeControlType.Zone2Bass, DenonRangeControl.CreateFromDenonCommandValue(m.Substring(8), 50, 1, False, "dB", RangeControlType.Zone2Bass))
						booProcessed = True
					ElseIf m.Substring(0, 8) = Helper.GetCommandPrefixForRangeControl(RangeControlType.Zone2Treble) Then
						Form1.SetRange(RangeControlType.Zone2Treble, DenonRangeControl.CreateFromDenonCommandValue(m.Substring(8), 50, 1, False, "dB", RangeControlType.Zone2Treble))
						booProcessed = True
					End If

			End Select

		End If

		If Not booProcessed Then
			Form1.SetZoneSource(Zones.Zone2, Source.CreateFromDenonCommandValue(m.Substring(2)))
			Form1.SetRecordSelect(Source.CreateFromDenonCommandValue(m.Substring(2)))
		End If

	End Sub

	Public Shared Sub ParseZone3(ByVal m As String)

		Dim strX As String = Helper.GetZonePrefix(Zones.Zone3)
		Dim rx As New Regex(strX & ".*?(?:OFF|ON)")
		Dim rx2 As New Regex(strX & "\d\d")
		Dim booProcessed As Boolean

		If rx.IsMatch(m) Then

			Select Case m
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.Zone3Off) & "OFF"
					Form1.SetOnOff(OnOffCommand.Zone3Off)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.Zone3On) & "ON"
					Form1.SetOnOff(OnOffCommand.Zone3On)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.Zone3MuteOff) & "OFF"
					Form1.SetOnOff(OnOffCommand.Zone3MuteOff)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.Zone3MuteOn) & "ON"
					Form1.SetOnOff(OnOffCommand.Zone3MuteOn)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.Zone3HPFOff) & "OFF"
					Form1.SetOnOff(OnOffCommand.Zone3HPFOff)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.Zone3HPFOn) & "ON"
					Form1.SetOnOff(OnOffCommand.Zone3HPFOn)
					booProcessed = True
			End Select

		End If

		If Not booProcessed AndAlso rx2.IsMatch(m) Then

			Form1.SetMasterVolume(MasterVolume.CreateFromDenonCommandValue(m.Substring(2), -70), Zones.Zone3)
			booProcessed = True
		End If

		If Not booProcessed AndAlso m.Length >= 4 Then

			Select Case m.Substring(0, 4)
				Case strX & "CS"
					Form1.SetFixedListComboBox(m.Substring(4), FixedListComboBox.Zone3ChannelSetting)
					booProcessed = True
				Case strX & "CV"

					If m.Substring(4, 2) = "FL" Then
						Form1.SetChannelVolume(Channel.Zone3Left, ChannelVolume.CreateFromDenonCommandValue(m.Substring(7)))
						booProcessed = True
					ElseIf m.Substring(4, 2) = "FR" Then
						Form1.SetChannelVolume(Channel.Zone3Right, ChannelVolume.CreateFromDenonCommandValue(m.Substring(7)))
						booProcessed = True
					End If

				Case strX & "PS"

					If m.Substring(0, 8) = Helper.GetCommandPrefixForRangeControl(RangeControlType.Zone3Bass) Then
						Form1.SetRange(RangeControlType.Zone3Bass, DenonRangeControl.CreateFromDenonCommandValue(m.Substring(8), 50, 1, False, "dB", RangeControlType.Zone3Bass))
						booProcessed = True
					ElseIf m.Substring(0, 8) = Helper.GetCommandPrefixForRangeControl(RangeControlType.Zone3Treble) Then
						Form1.SetRange(RangeControlType.Zone3Treble, DenonRangeControl.CreateFromDenonCommandValue(m.Substring(8), 50, 1, False, "dB", RangeControlType.Zone3Treble))
						booProcessed = True
					End If

			End Select

		End If

		If Not booProcessed Then
			Form1.SetZoneSource(Zones.Zone3, Source.CreateFromDenonCommandValue(m.Substring(2)))
		End If

	End Sub

	Public Shared Sub ParseMSResponse(ByVal m As String)

		Dim strMode As String
		Dim strSetting As String

		Select Case m

			Case "MSDIRECT"
				strMode = "DIRECT"
				strSetting = strMode
			Case "MSPURE DIRECT"
				strMode = "PURE DIRECT"
				strSetting = strMode
			Case "MSSTEREO"
				strMode = "STEREO"
				strSetting = strMode
			Case "MSMULTI CH IN"
				strMode = "HDMI MULTI CHANNEL IN"
			Case "MSM CH IN+PL2X C"
				strMode = "HDMI MULTI CHANNEL IN PL2X CINEMA"
			Case "MSM CH IN+PL2X M"
				strMode = "HDMI MULTI CHANNEL IN PL2X MUSIC"
			Case "MSMULTI CH DIRECT"
				strMode = "HDMI MULTI CHANNEL IN DIRECT"
				strSetting = "DIRECT"
			Case "MSMULTI CH PURE"
				strMode = "HDMI MULTI CHANNEL IN PURE DIRECT"
				strSetting = "PURE DIRECT"
			Case "MSMULTI CH IN 7.1"
				strMode = "ANALOG 7.1 MULTI CHANNEL IN"
			Case "MSM CH DRCT+PL2X C"
				strMode = "HDMI MULTI CHANNEL IN DIRECT PL2X CINEMA"
				strSetting = "DIRECT"
			Case "MSM CH DRCT+PL2X M"
				strMode = "HDMI MULTI CHANNEL IN DIRECT PL2X MUSIC"
				strSetting = "DIRECT"
			Case "MSM CH PURE D+PL2X C"
				strMode = "HDMI MULTI CHANNEL IN PURE DIRECT PL2X CINEMA"
				strSetting = "PURE DIRECT"
			Case "MSM CH PURE D+PL2X M"
				strMode = "HDMI MULTI CHANNEL IN PURE DIRECT PL2X MUSIC"
				strSetting = "PURE DIRECT"
			Case "MSM DIRECT 7.1"
				strMode = "ANALOG 7.1 MULTI CHANNEL IN DIRECT"
				strSetting = "DIRECT"
			Case "MSM CH PURE D 7.1"
				strMode = "ANALOG 7.1 MULTI CHANNEL IN PURE DIRECT"
				strSetting = "PURE DIRECT"
			Case "MSDOLBY PRO LOGIC"
				strMode = "DOLBY PRO LOGIC (PL)"
				strSetting = "DOLBY DIGITAL"
			Case "MSDOLBY PL2 C"
				strMode = "DOLBY PRO LOGIC 2 (PL2) CINEMA"
				strSetting = "DOLBY DIGITAL"
			Case "MSDOLBY PL2 M"
				strMode = "DOLBY PRO LOGIC 2 (PL2) MUSIC"
				strSetting = "DOLBY DIGITAL"
			Case "MSDOLBY PL2 G"
				strMode = "DOLBY PRO LOGIC 2 (PL2) GAME"
				strSetting = "DOLBY DIGITAL"
			Case "MSDOLBY PL2X C"
				strMode = "DOLBY PRO LOGIC 2X (PL2X) CINEMA"
				strSetting = "DOLBY DIGITAL"
			Case "MSDOLBY PL2X M"
				strMode = "DOLBY PRO LOGIC 2X (PL2X) MUSIC"
				strSetting = "DOLBY DIGITAL"
			Case "MSDOLBY PL2X G"
				strMode = "DOLBY PRO LOGIC 2X (PL2X) GAME"
				strSetting = "DOLBY DIGITAL"
			Case "MSDOLBY DIGITAL"
				strMode = "DOLBY DIGITAL"
				strSetting = strMode
			Case "MSDOLBY D EX"
				strMode = "DOLBY DIGITAL EX"
				strSetting = "DOLBY DIGITAL"
			Case "MSDOLBY D+PL2X C"
				strMode = "DOLBY DIGITAL PL2X CINEMA"
				strSetting = "DOLBY DIGITAL"
			Case "MSDOLBY D+PL2X M"
				strMode = "DOLBY DIGITAL PL2X MUSIC"
				strSetting = "DOLBY DIGITAL"
			Case "MSDTS NEO:6 C"
				strMode = "DTS NEO:6 CINEMA"
				strSetting = "DTS SURROUND"
			Case "MSDTS NEO:6 M"
				strMode = "DTS NEO:6 MUSIC"
				strSetting = "DTS SURROUND"
			Case "MSDTS SURROUND"
				strMode = "DTS SURROUND"
				strSetting = "DTS SURROUND"
			Case "MSDTS ES DSCRT6.1"
				strMode = "DTS ES DISCREET 6.1"
				strSetting = "DTS SURROUND"
			Case "MSDTS ES MTRX6.1"
				strMode = "DTS ES MATRIX 6.1"
				strSetting = "DTS SURROUND"
			Case "MSDTS+PL2X C"
				strMode = "DTS SURROUND PL2X CINEMA"
				strSetting = "DTS SURROUND"
			Case "MSDTS+PL2X M"
				strMode = "DTS SURROUND PL2X MUSIC"
				strSetting = "DTS SURROUND"
			Case "MSWIDE SCREEN"
				strMode = "DSP WIDE SCREEN"
				strSetting = m.Substring(2)
			Case "MS5CH STEREO"
				strMode = "DSP 5 CHANNEL STEREO"
				strSetting = m.Substring(2)
			Case "MS7CH STEREO"
				strMode = "DSP 7 CHANNEL STEREO"
				strSetting = m.Substring(2)
			Case "MSSUPER STADIUM"
				strMode = "DSP SUPER STADIUM"
				strSetting = m.Substring(2)
			Case "MSROCK ARENA"
				strMode = "DSP ROCK ARENA"
				strSetting = m.Substring(2)
			Case "MSJAZZ CLUB"
				strMode = "DSP JAZZ CLUB"
				strSetting = m.Substring(2)
			Case "MSCLASSIC CONCERT"
				strMode = "DSP CLASSICAL MUSIC CONCERT"
				strSetting = m.Substring(2)
			Case "MSMONO MOVIE"
				strMode = "DSP MONO MOVIE"
				strSetting = m.Substring(2)
			Case "MSMATRIX"
				strMode = "DSP MATRIX"
				strSetting = m.Substring(2)
			Case "MSVIDEO GAME"
				strMode = "DSP VIDEO GAME"
				strSetting = m.Substring(2)
			Case "MSVIRTUAL"
				strMode = "DSP VIRTUAL"
				strSetting = m.Substring(2)
			Case "MSDSD DIRECT"
				strMode = "DSD DIRECT"
				strSetting = "DIRECT"
			Case "MSDSD PURE DIRECT"
				strMode = "DSD PURE DIRECT"
				strSetting = "PURE DIRECT"
			Case "MSDSD MULTI DRCT"
				strMode = "DSD MULTI CHANNEL DIRECT"
				strSetting = "DIRECT"
			Case "MSDSD MULTI PURE"
				strMode = "DSD MULTIPLE CHANNEL PURE DIRECT"
				strSetting = "PRUE DIRECT"
			Case "MSNEURAL"
				strMode = "NEURAL SURROUND"
				strSetting = m.Substring(2)
			Case "MSDOLBY D+"
				strMode = "DOLBY DIGITAL PLUS"
				strSetting = "DOLBY DIGITAL"
			Case "MSDOLBY D+ +EX"
				strMode = "DOLBY DIGITAL PLUS EX"
				strSetting = "DOLBY DIGITAL"
			Case "MSDOLBY D+ +PL2X C"
				strMode = "DOLBY DIGITAL PLUS PL2X CINEMA"
				strSetting = "DOLBY DIGITAL"
			Case "MSDOLBY D+ +PL2X M"
				strMode = "DOLBY DIGITAL PLUS PL2X MUSIC"
				strSetting = "DOLBY DIGITAL"
			Case "MSDOLBY HD"
				strMode = "DOLBY TRUE HD"
				strSetting = "DOLBY DIGITAL"
			Case "MSDOLBY HD+EX"
				strMode = "DOLBY TRUE HD EX"
				strSetting = "DOLBY DIGITAL"
			Case "MSDOLBY HD+PL2X C"
				strMode = "DOLBY TRUE HD PL2X CINEMA"
				strSetting = "DOLBY DIGITAL"
			Case "MSDOLBY HD+PL2X M"
				strMode = "DOLBY TRUE HD PL2X MUSIC"
				strSetting = "DOLBY DIGITAL"
			Case "MSDTS HD"
				strMode = "DTS HD"
				strSetting = "DTS SURROUND"
			Case "MSDTS HD MSTR"
				strMode = "DTS HD MASTER"
				strSetting = "DTS SURROUND"
			Case "MSDTS  ES 8CH DSCRT"
				strMode = "DTS ES 8 CHANNEL DISCREEET"
				strSetting = "DTS SURROUND"
			Case "MSMPEG2 AAC"
				strMode = "MPEG2 AAC"
			Case "MSAAC+DOLBY EX"
				strMode = "AAC DOLBY EX"
				strSetting = "DOLBY DIGITAL"
			Case "MSAAC+PL2X C"
				strMode = "AAC DOLBY PL2X CINEMA"
				strSetting = "DOLBY DIGITAL"
			Case "MSAAC+PL2X M"
				strMode = "AAC DOLBY PL2X MUSIC"
				strSetting = "DOLBY DIGITAL"
			Case "MSDTS+NEO:6"
				strMode = "DTS SURROUND NEO:6"
				strSetting = "DTS SURROUND"
			Case "MSDTS96/24"
				strMode = "DTS SURROUND 96KHZ 24BIT"
				strSetting = "DTS SURROUND"
			Case "MSDTS96 ES MTRX"
				strMode = "DTS SURROUND ES MATRIX 96KHZ 24BIT"
				strSetting = "DTS SURROUND"
			Case Else
				strMode = "UNKNOWN SURROUND SETTINGS: " & m
		End Select

		Form1.SetSurroundSetting(strMode)

		If Not String.IsNullOrEmpty(strSetting) Then
			Form1.SetFixedListComboBox(strSetting, FixedListComboBox.SurroundSetting)
		End If


	End Sub

	Public Shared Sub ParsePSResponse(ByVal m As String)

		Dim rx As New Regex("PS.*?(?:OFF|ON)")
		Dim booProcessed As Boolean

		If rx.IsMatch(m) Then

			Select Case m
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.ToneDefeatOff) & "OFF"
					Form1.SetOnOff(OnOffCommand.ToneDefeatOff)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.ToneDefeatOn) & "ON"
					Form1.SetOnOff(OnOffCommand.ToneDefeatOn)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.CinemaEQOff) & "OFF"
					Form1.SetOnOff(OnOffCommand.CinemaEQOff)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.CinemaEQOn) & "ON"
					Form1.SetOnOff(OnOffCommand.CinemaEQOn)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.AFDOff) & "OFF"
					Form1.SetOnOff(OnOffCommand.AFDOff)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.AFDOn) & "ON"
					Form1.SetOnOff(OnOffCommand.AFDOn)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.SWAttOff) & "OFF"
					Form1.SetOnOff(OnOffCommand.SWAttOff)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.SWAttOn) & "ON"
					Form1.SetOnOff(OnOffCommand.SWAttOn)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.SWOff) & "OFF"
					Form1.SetOnOff(OnOffCommand.SWOff)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.SWOn) & "ON"
					Form1.SetOnOff(OnOffCommand.SWOn)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.PanoramaOff) & "OFF"
					Form1.SetOnOff(OnOffCommand.PanoramaOff)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.PanoramaOn) & "ON"
					Form1.SetOnOff(OnOffCommand.PanoramaOn)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.DSPEffectOff) & "OFF"
					Form1.SetOnOff(OnOffCommand.DSPEffectOff)
					booProcessed = True
				Case Helper.GetCommandPrefixForOnOff(OnOffCommand.DSPEffectOn) & "ON"
					Form1.SetOnOff(OnOffCommand.DSPEffectOn)
					booProcessed = True
			End Select

		End If

		If Not booProcessed Then

			If m.Length >= 6 AndAlso m.Substring(0, 6) = Helper.GetCommandPrefixForRangeControl(RangeControlType.Bass) Then
				Form1.SetRange(RangeControlType.Bass, DenonRangeControl.CreateFromDenonCommandValue(m.Substring(6), 50, 1, False, "dB", RangeControlType.Bass))
			ElseIf m.Length >= 6 AndAlso m.Substring(0, 6) = Helper.GetCommandPrefixForRangeControl(RangeControlType.Treble) Then
				Form1.SetRange(RangeControlType.Treble, DenonRangeControl.CreateFromDenonCommandValue(m.Substring(6), 50, 1, False, "dB", RangeControlType.Treble))
			ElseIf m.Length >= 6 AndAlso m.Substring(0, 6) = Helper.GetCommandPrefixForRangeControl(RangeControlType.LFE) Then
				Form1.SetRange(RangeControlType.LFE, DenonRangeControl.CreateFromDenonCommandValue(m.Substring(6), 0, 1, True, "dB", RangeControlType.LFE))
			ElseIf m.Length >= 6 AndAlso m.Substring(0, 6) = Helper.GetCommandPrefixForRangeControl(RangeControlType.Dimension) Then
				Form1.SetRange(RangeControlType.Dimension, DenonRangeControl.CreateFromDenonCommandValue(m.Substring(6), 0, 1, False, "", RangeControlType.Dimension))
			ElseIf m.Length >= 6 AndAlso m.Substring(0, 6) = Helper.GetCommandPrefixForRangeControl(RangeControlType.CenterWidth) Then
				Form1.SetRange(RangeControlType.CenterWidth, DenonRangeControl.CreateFromDenonCommandValue(m.Substring(6), 0, 1, False, "", RangeControlType.CenterWidth))
			ElseIf m.Length >= 6 AndAlso m.Substring(0, 6) = Helper.GetCommandPrefixForRangeControl(RangeControlType.CenterImage) Then
				Form1.SetRange(RangeControlType.CenterImage, DenonRangeControl.CreateFromDenonCommandValue(m.Substring(6), 0, 0.1D, False, "", RangeControlType.CenterImage))
			ElseIf m.Length >= 6 AndAlso m.Substring(0, 6) = Helper.GetCommandPrefixForRangeControl(RangeControlType.DSPSurroundDelay) Then
				Form1.SetRange(RangeControlType.DSPSurroundDelay, DenonRangeControl.CreateFromDenonCommandValue(m.Substring(6), 0, 10, False, "ms", RangeControlType.DSPSurroundDelay))
			ElseIf m.Length >= 6 AndAlso m.Substring(0, 6) = Helper.GetCommandPrefixForRangeControl(RangeControlType.DSPEffectLevel) Then
				Form1.SetRange(RangeControlType.DSPEffectLevel, DenonRangeControl.CreateFromDenonCommandValue(m.Substring(6), 0, 1, False, "", RangeControlType.DSPEffectLevel))
			ElseIf m.Length >= 6 AndAlso m.Substring(0, 6) = "PSDRC " Then
				Form1.SetFixedListComboBox(m.Substring(6), FixedListComboBox.DRC)
			ElseIf m.Length >= 6 AndAlso m.Substring(0, 6) = "PSDCO " Then
				Form1.SetFixedListComboBox(m.Substring(6), FixedListComboBox.DCOMP)
			ElseIf m.Length >= 6 AndAlso m.Substring(0, 6) = "PSRSZ " Then
				Form1.SetFixedListComboBox(m.Substring(6), FixedListComboBox.RoomSize)
			ElseIf m.Length >= 7 AndAlso m.Substring(0, 7) = "PSRSTR " Then
				Form1.SetFixedListComboBox(m.Substring(7), FixedListComboBox.Restorer)
			ElseIf m.Length >= 8 AndAlso m.Substring(0, 8) = "PSNIGHT " Then
				Form1.SetFixedListComboBox(m.Substring(8), FixedListComboBox.Night)
			ElseIf m.Length >= 7 AndAlso m.Substring(0, 7) = "PSMODE:" Then
				Form1.SetFixedListComboBox(m.Substring(7), FixedListComboBox.SurroundMode)
			ElseIf m.Length >= 5 AndAlso m.Substring(0, 5) = "PSSB:" Then
				Form1.SetFixedListComboBox(m.Substring(5), FixedListComboBox.SurroundBackMode)
			ElseIf m.Length >= 10 AndAlso m.Substring(0, 10) = "PSROOM EQ:" Then
				Form1.SetFixedListComboBox(m.Substring(10), FixedListComboBox.RoomEQ)
			ElseIf m.Length >= 8 AndAlso m.Substring(0, 8) = Helper.GetCommandPrefixForRangeControl(RangeControlType.AudioDelay) Then
				Form1.SetRange(RangeControlType.AudioDelay, DenonRangeControl.CreateFromDenonCommandValue(m.Substring(8), 0, 1, False, "ms", RangeControlType.AudioDelay))
			End If

		End If

	End Sub

	Public Shared Function ParseChannelVolumeResponse(ByVal m As String) As ReadWriteKeyValuePair(Of Channel, ChannelVolume)

		Dim r As ChannelVolume
		Dim c As Channel
		Dim rx As New Regex("CV(?<1>[A-Z]{1,3})\x20(?<2>\d{2,3})")
		Dim kvp As ReadWriteKeyValuePair(Of Channel, ChannelVolume)
		Dim ma As Match

		ma = rx.Match(m)

		If Not IsNothing(ma) AndAlso ma.Success Then

			Select Case ma.Groups(1).Value
				Case "FL"
					c = Channel.LeftFront
				Case "C"
					c = Channel.Center
				Case "FR"
					c = Channel.RightFront
				Case "SL"
					c = Channel.LeftSurround
				Case "SR"
					c = Channel.RightSurround
				Case "SBL"
					c = Channel.LeftRear
				Case "SBR"
					c = Channel.RightRear
				Case "SW"
					c = Channel.Subwoofer
			End Select

			r = ChannelVolume.CreateFromDenonCommandValue(ma.Groups(2).Value)

			kvp = New ReadWriteKeyValuePair(Of Channel, ChannelVolume)(c, r)

		End If

		Return kvp

	End Function

	Public Shared Function ParseMasterVolumeResponse(ByVal m As String) As MasterVolume

		Dim r As MasterVolume
		Dim rx As New Regex("MV\d{2,3}")

		If rx.IsMatch(m) Then
			r = MasterVolume.CreateFromDenonCommandValue(m.Substring(2), -80)
		End If

		Return r

	End Function

	Public Shared Function ParsePowerResponse(ByVal m As String) As OnOffCommand

		Dim r As OnOffCommand
		Dim rx As New Regex("PW.*")

		If rx.IsMatch(m) Then
			Select Case m.ToUpper
				Case "PWON"
					r = OnOffCommand.PowerOn
				Case "PWSTANDBY"
					r = OnOffCommand.PowerOff
			End Select
		End If

		Return r

	End Function

	Public Shared Function ParseMuteResponse(ByVal m As String) As OnOffCommand

		Dim r As OnOffCommand
		Dim rx As New Regex("MU.*")

		If rx.IsMatch(m) Then
			Select Case m.ToUpper
				Case "MUON"
					r = OnOffCommand.MuteOn
				Case "MUOFF"
					r = OnOffCommand.MuteOff
			End Select
		End If

		Return r

	End Function

	Public Shared Function ParseSourceResponse(ByVal m As String) As Source

		Dim r As Source
		Dim rx As New Regex("SI.*")

		If rx.IsMatch(m) Then
			r = Source.CreateFromDenonCommandValue(m.Substring(2))
		End If

		Return r

	End Function

	Public Shared Function ParseVideoSelectResponse(ByVal m As String) As Source

		Dim r As Source
		Dim rx As New Regex("SV.*")

		If rx.IsMatch(m) Then
			r = Source.CreateFromDenonCommandValue(m.Substring(2))
		End If

		Return r

	End Function

	Public Shared Function ParseRecordSelectResponse(ByVal m As String) As Source

		Dim r As Source
		Dim rx As New Regex("SR.*")

		If rx.IsMatch(m) Then
			r = Source.CreateFromDenonCommandValue(m.Substring(2))
		End If

		Return r

	End Function

	Public Shared Function ParseAudioInputModeResponse(ByVal m As String) As String

		Dim rx As New Regex("SD.*")
		Dim r As String

		If rx.IsMatch(m) Then
			r = m.Substring(2)
		End If

		Return r

	End Function

	Public Shared Function ParseDigitalAudioModeResponse(ByVal m As String) As String

		Dim rx As New Regex("DC.*")
		Dim r As String

		If rx.IsMatch(m) Then
			r = m.Substring(2)
		End If

		Return r

	End Function

#End Region

End Class
