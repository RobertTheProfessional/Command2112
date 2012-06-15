''''''''' ''''''''' ''''''''' ''''''''' ''''''''' 
''''''''' Copyright 2007-2009 Brian Saville
''''''''' Command2112 is free for non-commercial use.
''''''''' Reuse of the source code is free for non-commercial use
''''''''' provided that the source code used is credited to Command2112.
''''''''' ''''''''' ''''''''' ''''''''' '''''''''
Public Class Helper

	Public Shared Function SystemPredicateRemoveNonAsciiBytes(ByVal obj As Byte) As Boolean

		If obj <= 31 OrElse obj >= 127 Then
			Return True
		End If

	End Function

	Public Shared Function GetTunerFrequencyDisplayValue(ByVal FrequencyValue As String) As String

		Dim decX As Decimal = CDec(FrequencyValue.Insert(4, CultureDecimalSeperator))

		If decX < 500 Then
			Return decX.ToString("#.00") & " MHz (FM)"
		Else
			Return decX.ToString("#") & " kHz (AM)"
		End If

	End Function

	Public Shared Function CultureDecimalSeperator() As String
		Return CurrentCulture.NumberFormat.NumberDecimalSeparator
	End Function

	Public Shared Function CurrentCulture() As System.Globalization.CultureInfo
		Return System.Globalization.CultureInfo.CurrentCulture
	End Function

	Public Shared Function GetNowString() As String
		Return Now.ToString("yyyy-MM-dd hh:mm:ss.fff")
	End Function

	Public Shared Function CheckIP(ByVal SuppressMessageBox As Boolean, ByVal SuppressLogging As Boolean) As Boolean

		If String.IsNullOrEmpty(My.CustomSettings.IPAddress.Instance.Value) OrElse My.CustomSettings.IPAddress.Instance.Value = "10.0.0.0" OrElse Not System.Net.IPAddress.TryParse(My.CustomSettings.IPAddress.Instance.Value, Nothing) Then

			If Not SuppressMessageBox Then
				MsgBox("Please enter a valid IP address")
			End If

			If Not SuppressLogging Then
				Form1.WriteDebug("No valid IP address.", True)
			End If

			Return False
		Else
			Return True
		End If

	End Function

	Public Shared Function GetCommandPrefixForOnOff(ByVal cmd As OnOffCommand) As String

		Dim strX As String

        Select Case cmd
            Case OnOffCommand.DynamicEQ, OnOffCommand.DynamicEQOff, OnOffCommand.DynamicEQOn
                strX = "PSDYNEQ "
            Case OnOffCommand.ToneCtrl, OnOffCommand.ToneCtrlOff, OnOffCommand.ToneCtrlOn
                strX = "PSTONE CTRL "
            Case OnOffCommand.Power, OnOffCommand.PowerOff, OnOffCommand.PowerOn
                strX = "PW"
            Case OnOffCommand.CinemaEQ
                strX = "PSCINEMA EQ. "
            Case OnOffCommand.CinemaEQOff, OnOffCommand.CinemaEQOn
                strX = "PSCINEMA EQ."
            Case OnOffCommand.AFD, OnOffCommand.AFDOff, OnOffCommand.AFDOn
                strX = "PSAFD "
            Case OnOffCommand.SWAtt, OnOffCommand.SWAttOff, OnOffCommand.SWAttOn
                strX = "PSATT "
            Case OnOffCommand.SW, OnOffCommand.SWOff, OnOffCommand.SWOn
                strX = "PSSWR "
            Case OnOffCommand.Panorama, OnOffCommand.PanoramaOff, OnOffCommand.PanoramaOn
                strX = "PSPAN "
            Case OnOffCommand.ZoneMain, OnOffCommand.ZoneMainOff, OnOffCommand.ZoneMainOn
                strX = GetZonePrefix(Zones.ZoneMain)
            Case OnOffCommand.Zone2, OnOffCommand.Zone2Off, OnOffCommand.Zone2On
                strX = GetZonePrefix(Zones.Zone2)
            Case OnOffCommand.Zone3, OnOffCommand.Zone3Off, OnOffCommand.Zone3On
                strX = GetZonePrefix(Zones.Zone3)
            Case OnOffCommand.Zone2Mute, OnOffCommand.Zone2MuteOff, OnOffCommand.Zone2MuteOn
                strX = GetZonePrefix(Zones.Zone2) & "MU"
            Case OnOffCommand.Zone3Mute, OnOffCommand.Zone3MuteOff, OnOffCommand.Zone3MuteOn
                strX = GetZonePrefix(Zones.Zone3) & "MU"
            Case OnOffCommand.Zone2HPF, OnOffCommand.Zone2HPFOff, OnOffCommand.Zone2HPFOn
                strX = GetZonePrefix(Zones.Zone2) & "HPF"
            Case OnOffCommand.Zone3HPF, OnOffCommand.Zone3HPFOff, OnOffCommand.Zone3HPFOn
                strX = GetZonePrefix(Zones.Zone3) & "HPF"
            Case OnOffCommand.Mute, OnOffCommand.MuteOff, OnOffCommand.MuteOn
                strX = "MU"
            Case OnOffCommand.DSPEffect, OnOffCommand.DSPEffectOff, OnOffCommand.DSPEffectOn
                strX = "PSEFF "
        End Select

        Return strX

    End Function

    Public Shared Function GetCommandPrefixForRangeControl(ByVal db As RangeControlType) As String

        Dim strX As String

        Select Case db
            Case RangeControlType.Bass
                strX = "PSBAS "
            Case RangeControlType.Treble
                strX = "PSTRE "
            Case RangeControlType.LFE
                strX = "PSLFE "
            Case RangeControlType.Dimension
                strX = "PSDIM "
            Case RangeControlType.CenterWidth
                strX = "PSCEN "
            Case RangeControlType.CenterImage
                strX = "PSCEI "
            Case RangeControlType.Zone2Bass
                strX = GetZonePrefix(Zones.Zone2) & "PSBAS "
            Case RangeControlType.Zone2Treble
                strX = GetZonePrefix(Zones.Zone2) & "PSTRE "
            Case RangeControlType.Zone3Bass
                strX = GetZonePrefix(Zones.Zone3) & "PSBAS "
            Case RangeControlType.Zone3Treble
                strX = GetZonePrefix(Zones.Zone3) & "PSTRE "
            Case RangeControlType.DSPSurroundDelay
                strX = "PSDEL "
            Case RangeControlType.AudioDelay
                strX = "PSDELAY "
            Case RangeControlType.Contrast
                strX = "PVCN "
            Case RangeControlType.Brightness
                strX = "PVBR "
            Case RangeControlType.Chroma
                strX = "PVCM "
            Case RangeControlType.Hue
                strX = "PVHUE "
            Case RangeControlType.DSPEffectLevel
                strX = "PSEFF "
        End Select

        Return strX

    End Function

    Public Shared Function GetZonePrefix(ByVal z As Zones) As String

        Dim strX As String

        Select Case z
            Case Zones.ZoneMain
                strX = "ZM"
            Case Zones.Zone2
                strX = "Z2"
            Case Zones.Zone3
                strX = "Z3"
        End Select

        Return strX

    End Function

    Public Shared Function GetParentGroupBoxForControl(ByVal c As Control) As GroupBox

        Dim gb As GroupBox

        If Not c.Parent Is Nothing AndAlso TypeOf c.Parent Is GroupBox Then
            gb = DirectCast(c.Parent, GroupBox)
        End If

        Return gb

    End Function

    Public Shared Function LimitedSizeInputBoxInputBox(ByVal Prompt As String, ByVal MaxSize As Integer, Optional ByVal Title As String = "", Optional ByVal DefaultResponse As String = "", Optional ByVal XPos As Integer = -1, Optional ByVal YPos As Integer = -1) As String

        Dim strX As String
        Dim strDefault As String = DefaultResponse

        Do

            Prompt &= vbCrLf & vbCrLf & "Max " & MaxSize & " characters."

            strX = InputBox(Prompt, Title, strDefault, XPos, YPos)

            If strX.Length <= MaxSize Then
                Exit Do
            Else
                strDefault = strX
                MsgBox("Your response of " & strX.Length & " characters is too long." & vbCrLf & vbCrLf & "Please limit your response to " & MaxSize & " characters.", MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
            End If

        Loop

        Return strX

    End Function

    Public Shared Function GetDaySuffix(ByVal Day As Integer) As String

        Select Case Day

            Case 1, 21, 31
                Return "st"
            Case 2, 22
                Return "nd"
            Case 3, 23
                Return "rd"
            Case Else
                Return "th"
        End Select

    End Function

    ''' <summary>
    ''' Converts a string to the SHA256 representation of that string
    ''' </summary>
    Public Shared Function StringToSHA256String(ByVal Value As String) As String

        Dim sham As New System.Security.Cryptography.SHA256Managed
        Dim byt() As Byte

        byt = System.Text.Encoding.Default.GetBytes(Value)
        byt = sham.ComputeHash(byt)

        Return BitConverter.ToString(byt).Replace("-", "").ToLower

    End Function

End Class

Public Class MasterVolume

    Private m_DecibalValue As Decimal
    Private m_MinimumDecibalValueAllowed As Decimal

    Private ReadOnly Property MinimumDenonValueAllowed() As Decimal
        Get
            Return m_MinimumDecibalValueAllowed + 80
        End Get
    End Property

    Public ReadOnly Property DenonCommandValue() As String
        Get
            Dim strX As String
            Dim decX As Decimal = DenonValue

            If decX >= MinimumDenonValueAllowed AndAlso decX <= 98 Then
                If decX Mod 1 = 0 Then
                    strX = decX.ToString("0.0").Replace(Helper.CultureDecimalSeperator & "0", "").PadLeft(2, "0"c)
                Else
                    strX = decX.ToString("0.0").Replace(Helper.CultureDecimalSeperator, "").PadLeft(3, "0"c)
                End If
            ElseIf decX < MinimumDenonValueAllowed OrElse decX = 99 Then
                strX = "99"
            Else
                strX = "98"
            End If

            Return strX
        End Get
    End Property

    Public ReadOnly Property DenonValue() As Decimal
        Get
            Dim decX As Decimal

            If DecibalValue >= m_MinimumDecibalValueAllowed AndAlso DecibalValue <= 18 Then
                decX = DecibalValue + 80
            ElseIf DecibalValue < m_MinimumDecibalValueAllowed Then
                decX = 99
            Else
                decX = 98
            End If

            Return decX
        End Get
    End Property

    Public ReadOnly Property DecibalValue() As Decimal
        Get
            Return m_DecibalValue
        End Get
    End Property

    Public ReadOnly Property DisplayValue() As String
        Get

            Dim strX As String

            If DecibalValue >= m_MinimumDecibalValueAllowed AndAlso DecibalValue <= 18 Then
                strX = Format(DecibalValue, "0.0") & " dB"
            ElseIf DecibalValue < m_MinimumDecibalValueAllowed Then
                strX = "Off"
            Else
                strX = "Max"
            End If

            Return strX

        End Get
    End Property

    Public Shared Function CreateFromDenonCommandValue(ByVal DenonCommandValue As String, ByVal MinimumDecibalValueAllowed As Decimal) As MasterVolume

        If Not Short.TryParse(DenonCommandValue, Nothing) OrElse DenonCommandValue.Length < 2 OrElse DenonCommandValue.Length > 3 Then
            Throw New Exception("DenonCommandValue must either be 2 or 3 numbers.")
        End If

        If DenonCommandValue.Length = 2 Then
            If DenonCommandValue <> "99" Then
                Return New MasterVolume(CDec(DenonCommandValue) - 80, MinimumDecibalValueAllowed)
            Else
                Return New MasterVolume(-80.5D, MinimumDecibalValueAllowed)
            End If
        Else
            Return New MasterVolume(CDec(DenonCommandValue.Insert(2, Helper.CultureDecimalSeperator)) - 80, MinimumDecibalValueAllowed)
        End If

    End Function

    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf (obj) Is MasterVolume Then
            If DirectCast(obj, MasterVolume).DenonCommandValue = Me.DenonCommandValue Then
                Return True
            End If
        End If
    End Function

    Sub New(ByVal DecibalValue As Decimal, ByVal MinimumDecibalValueAllowed As Decimal)

        If DecibalValue Mod 0.5D <> 0 Then
            Throw New Exception("DecibalValue must be a multiple of 0.5")
        End If

        Me.m_DecibalValue = DecibalValue
        Me.m_MinimumDecibalValueAllowed = MinimumDecibalValueAllowed

    End Sub

End Class

Public Class ChannelVolume

    Private m_DecibalValue As Decimal

    Public ReadOnly Property DenonCommandValue() As String
        Get
            Dim strX As String
            Dim decX As Decimal = DenonValue

            If decX >= 38 AndAlso decX <= 62 Then
                If decX Mod 1 = 0 Then
                    strX = decX.ToString("0.0").Replace(Helper.CultureDecimalSeperator & "0", "").PadLeft(2, "0"c)
                Else
                    strX = decX.ToString("0.0").Replace(Helper.CultureDecimalSeperator, "").PadLeft(3, "0"c)
                End If
            ElseIf decX < 38 Then
                strX = "00"
            Else
                strX = "62"
            End If

            Return strX
        End Get
    End Property

    Public ReadOnly Property DenonValue() As Decimal
        Get
            Dim decX As Decimal

            If DecibalValue >= -12 AndAlso DecibalValue <= 12 Then
                decX = DecibalValue + 50
            ElseIf DecibalValue < -12 Then
                decX = 0
            Else
                decX = 62
            End If

            Return decX
        End Get
    End Property

    Public ReadOnly Property DecibalValue() As Decimal
        Get
            Return m_DecibalValue
        End Get
    End Property

    Public ReadOnly Property DisplayValue() As String
        Get

            Dim strX As String

            If DecibalValue >= -12 AndAlso DecibalValue <= 12 Then
                strX = Format(DecibalValue, "0.0") & " dB"
            ElseIf DecibalValue < -12 Then
                strX = "Off"
            Else
                strX = "Max"
            End If

            Return strX

        End Get
    End Property

    Public Shared Function CreateFromDenonCommandValue(ByVal DenonCommandValue As String) As ChannelVolume

        If Not Short.TryParse(DenonCommandValue, Nothing) OrElse DenonCommandValue.Length < 2 OrElse DenonCommandValue.Length > 3 Then
            Throw New Exception("DenonCommandValue must either be 2 or 3 numbers.")
        End If

        If DenonCommandValue.Length = 2 Then
            If DenonCommandValue <> "00" Then
                Return New ChannelVolume(CDec(DenonCommandValue) - 50)
            Else
                Return New ChannelVolume(-12.5D)
            End If
        Else
            Return New ChannelVolume(CDec(DenonCommandValue.Insert(2, Helper.CultureDecimalSeperator)) - 50)
        End If

    End Function

    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf (obj) Is ChannelVolume Then
            If DirectCast(obj, ChannelVolume).DenonCommandValue = Me.DenonCommandValue Then
                Return True
            End If
        End If
    End Function

    Sub New(ByVal DecibalValue As Decimal)

        If DecibalValue Mod 0.5D <> 0 Then
            Throw New Exception("DecibalValue must be a multiple of 0.5")
        End If

        Me.m_DecibalValue = DecibalValue

    End Sub

End Class

Public Class DenonRangeControl

    Private m_RangeValue As Decimal
    Private m_DenonValueForZero As Integer
    Private m_MinimumIncrement As Decimal
    Private m_DenonValueDecrements As Boolean
    Private m_UnitName As String
    Private m_CommandLength As String
    Private m_RCT As RangeControlType

    Public ReadOnly Property DenonCommandValue() As String
        Get
            Dim strX As String

            strX = CalculateDenonCommandValue()

            Return strX
        End Get
    End Property

    Private Function CalculateDenonCommandValue() As String

        Dim strX As String
        Dim decX As Decimal = DenonValue

        Select Case m_RCT
            Case RangeControlType.DSPSurroundDelay, RangeControlType.AudioDelay
                strX = decX.ToString("#.#").Replace(Helper.CultureDecimalSeperator, "").PadLeft(3, "0"c)
            Case RangeControlType.Brightness
                strX = decX.ToString("#.#").Replace(Helper.CultureDecimalSeperator, "").PadLeft(1, "0"c)
            Case Else
                strX = decX.ToString("#.#").Replace(Helper.CultureDecimalSeperator, "").PadLeft(2, "0"c)
        End Select

        Return strX

    End Function

    Public ReadOnly Property DenonValue() As Decimal
        Get
            Dim decX As Decimal

            decX = CalculateDenonValue()

            Return decX
        End Get
    End Property

    Private Function CalculateDenonValue() As Decimal

        Dim decX As Decimal

        Select Case m_RCT
            Case RangeControlType.CenterImage
                decX = RangeValue * 10
            Case Else
                If Not m_DenonValueDecrements Then
                    decX = RangeValue + m_DenonValueForZero
                Else
                    decX = RangeValue * -1 + m_DenonValueForZero
                End If
        End Select

        Return decX

    End Function

    Public ReadOnly Property RangeValue() As Decimal
        Get
            Return m_RangeValue
        End Get
    End Property

    Public ReadOnly Property DisplayValue() As String
        Get

            Dim strX As String

            strX = Format(RangeValue, "0.0") & " " & m_UnitName

            Return strX

        End Get
    End Property

    Private Shared Function CalculateRangeValue(ByVal DenonCommandValue As String, ByVal rct As RangeControlType) As Decimal

        Dim decX As Decimal

        Select Case rct
            Case RangeControlType.CenterImage
                decX = CDec(DenonCommandValue) / 10
            Case Else
                decX = CDec(DenonCommandValue)
        End Select

        Return decX

    End Function

    Public Shared Function CreateFromDenonCommandValue(ByVal DenonCommandValue As String, ByVal DenonValueForZero As Integer, ByVal MinimumIncrement As Decimal, ByVal DenonValueDecrements As Boolean, ByVal UnitName As String, ByVal rct As RangeControlType) As DenonRangeControl

        Dim intX As Integer = 1

        If Not Short.TryParse(DenonCommandValue, Nothing) OrElse DenonCommandValue.Length < 1 OrElse DenonCommandValue.Length > 3 Then
            Throw New Exception("DenonCommandValue must either be 1, 2 or 3 numbers.")
        End If

        If DenonValueDecrements Then
            intX = -1
        End If

        If DenonCommandValue.Length = 2 Then
            Return New DenonRangeControl(CalculateRangeValue(DenonCommandValue, rct) * intX - DenonValueForZero, DenonValueForZero, MinimumIncrement, DenonValueDecrements, UnitName, rct)
        Else
            Select Case rct
                Case RangeControlType.AudioDelay, RangeControlType.DSPSurroundDelay, RangeControlType.Brightness
                    Return New DenonRangeControl(CalculateRangeValue(DenonCommandValue, rct) * intX - DenonValueForZero, DenonValueForZero, MinimumIncrement, DenonValueDecrements, UnitName, rct)
                Case Else
                    Return New DenonRangeControl(CalculateRangeValue(DenonCommandValue.Insert(2, "."), rct) * intX - DenonValueForZero, DenonValueForZero, MinimumIncrement, DenonValueDecrements, UnitName, rct)
            End Select

        End If

    End Function

    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf (obj) Is DenonRangeControl Then
            If DirectCast(obj, DenonRangeControl).DenonCommandValue = Me.DenonCommandValue Then
                Return True
            End If
        End If
    End Function

    Sub New(ByVal RangeValue As Decimal, ByVal DenonValueForZero As Integer, ByVal MinimumIncrement As Decimal, ByVal DenonValueDecrements As Boolean, ByVal UnitName As String, ByVal rct As RangeControlType)

        If RangeValue Mod MinimumIncrement <> 0 Then
            Throw New Exception("DenonRangeControl must be a multiple of " & MinimumIncrement)
        End If

        Me.m_RangeValue = RangeValue
        Me.m_DenonValueForZero = DenonValueForZero
        Me.m_MinimumIncrement = MinimumIncrement
        Me.m_DenonValueDecrements = DenonValueDecrements
        Me.m_UnitName = UnitName
        Me.m_RCT = rct
    End Sub

End Class

Public Class RangeChangeButton : Inherits Button

    Private m_IndexChangeAmount As Integer
    Private m_ButtonType As RangeChangeButtonType

    Public Property IndexChangeAmount() As Integer
        Get
            Return m_IndexChangeAmount
        End Get
        Set(ByVal value As Integer)
            m_IndexChangeAmount = value
        End Set
    End Property

    Public Property ButtonType() As RangeChangeButtonType
        Get
            Return m_ButtonType
        End Get
        Set(ByVal value As RangeChangeButtonType)
            m_ButtonType = value
        End Set
    End Property

End Class

<Serializable()> _
Public Class Source

    Private m_IsActive As Boolean
    Private m_CustomerSourceName As String
    Private m_DefaultSourceName As String
    Private m_SourceType As SourceType

    Public Property IsActive() As Boolean
        Get
            Return m_IsActive
        End Get
        Set(ByVal value As Boolean)
            m_IsActive = value
        End Set
    End Property

    Public ReadOnly Property SourceName() As String
        Get
            If Not String.IsNullOrEmpty(CustomSourceName) Then
                Return CustomSourceName
            Else
                Return DefaultSourceName
            End If
        End Get
    End Property

    Public Property CustomSourceName() As String
        Get
            Return m_CustomerSourceName
        End Get
        Set(ByVal value As String)
            m_CustomerSourceName = value
        End Set
    End Property

    Public ReadOnly Property DefaultSourceName() As String
        Get
            Return m_DefaultSourceName
        End Get
    End Property

    Public Shared Function CreateFromDenonCommandValue(ByVal DenonCommandValue As String) As Source

        Dim s As Source

        If Not IsNothing(Form1.lSourceInfo(DenonCommandValue)) Then
            s = Form1.lSourceInfo(DenonCommandValue)
        End If

        Return s

    End Function

    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf (obj) Is Source Then
            If DirectCast(obj, Source).DefaultSourceName = Me.DefaultSourceName Then
                Return True
            End If
        End If
    End Function

    Public Property SourceType() As SourceType
        Get
            Return m_SourceType
        End Get
        Set(ByVal value As SourceType)
            m_SourceType = value
        End Set
    End Property

    Sub New(ByVal DefaultSourceName As String, ByVal st As SourceType)
        m_DefaultSourceName = DefaultSourceName
        m_IsActive = True
        m_SourceType = st
    End Sub

End Class

Public Class KeyValueList(Of K, V) : Inherits Generic.List(Of ReadWriteKeyValuePair(Of K, V))

    Sub AddNew(ByVal key As K, ByVal value As V)
        MyBase.Add(New ReadWriteKeyValuePair(Of K, V)(key, value))
    End Sub

End Class

Public Class ReadWriteKeyValuePair(Of K, V)

    Public m_Key As K
    Public m_Value As V

    Public Property Key() As K
        Get
            Return m_Key
        End Get
        Set(ByVal value As K)
            m_Key = value
        End Set
    End Property

    Public Property Value() As V
        Get
            Return m_Value
        End Get
        Set(ByVal value As V)
            m_Value = value
        End Set
    End Property

    Sub New()

    End Sub

    Sub New(ByVal key As K, ByVal value As V)
        m_Key = key
        m_Value = value
    End Sub
End Class

Public Class TunerPreset : Implements IComparable

    Private m_Code As String
    Private m_Name As String
    Private m_Frequency As String

    Public ReadOnly Property FrequencyForDisplay() As String
        Get
            Return Helper.GetTunerFrequencyDisplayValue(Frequency)
        End Get
    End Property

    Public ReadOnly Property Frequency() As String
        Get
            Return m_Frequency
        End Get
    End Property

    Public ReadOnly Property Name() As String
        Get
            Return m_Name
        End Get
    End Property

    Public ReadOnly Property Code() As String
        Get
            Return m_Code
        End Get
    End Property

    Sub New(ByVal PresetCode As String, ByVal PresetName As String, ByVal PresetFrequency As String)
        m_Code = PresetCode
        m_Name = PresetName
        m_Frequency = PresetFrequency
    End Sub

    Private Sub New()

    End Sub

    Public Shared Function ParseFromSSTPNString(ByVal SSTPN As String) As TunerPreset
        Return New TunerPreset(SSTPN.Substring(0, 2), SSTPN.Substring(2, 9).Trim, SSTPN.Substring(11))
    End Function

    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf obj Is TunerPreset AndAlso DirectCast(obj, TunerPreset).Code = Code Then
            Return True
        End If
    End Function

    Public Function CompareTo(ByVal obj As Object) As Integer Implements System.IComparable.CompareTo
        If TypeOf obj Is TunerPreset Then
            Return Code.CompareTo(DirectCast(obj, TunerPreset).Code)
        Else
            Throw New ArgumentException("obj is not the same type as this instance.")
        End If
    End Function

End Class

<Serializable()> _
Public Class TabPageInfo

    Private m_TabName As String
    Private m_TabOrder As Integer
    Private m_QueryOnLoad As Boolean
    Private m_QueryOnSelect As Boolean
    Private m_TabText As String
    Private m_IsDefault As Boolean

    Public Property IsDefault() As Boolean
        Get
            Return m_IsDefault
        End Get
        Set(ByVal value As Boolean)
            m_IsDefault = value
        End Set
    End Property

    Public ReadOnly Property TabText() As String
        Get
            Return m_TabText
        End Get
    End Property

    Public Property QueryOnSelect() As Boolean
        Get
            Return m_QueryOnSelect
        End Get
        Set(ByVal value As Boolean)
            m_QueryOnSelect = value
        End Set
    End Property

    Public Property QueryOnLoad() As Boolean
        Get
            Return m_QueryOnLoad
        End Get
        Set(ByVal value As Boolean)
            m_QueryOnLoad = value
        End Set
    End Property

    Public Property TabOrder() As Integer
        Get
            Return m_TabOrder
        End Get
        Set(ByVal value As Integer)
            m_TabOrder = value
        End Set
    End Property

    Public ReadOnly Property TabName() As String
        Get
            Return m_TabName
        End Get
    End Property

    Sub New(ByVal Name As String, ByVal Text As String, ByVal Order As Integer, ByVal QueryOnLoad As Boolean, ByVal QueryOnSelect As Boolean)
        m_TabName = Name
        m_TabText = Text
        m_TabOrder = Order
        m_QueryOnLoad = QueryOnLoad
        m_QueryOnSelect = QueryOnSelect
    End Sub

    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf obj Is TabPageInfo AndAlso DirectCast(obj, TabPageInfo).TabName = TabName Then
            Return True
        End If
    End Function

End Class

Public Class Schedule

    Private m_OneTimeDateTime As Date
    Private m_IntervalStartDateTime As Date
    Private m_DayWeekDays As New List(Of DayOfWeek)
    Private m_DayWeekTimes As New List(Of Date)
    Private m_MonthDates As New List(Of Integer)
    Private m_MonthTimes As New List(Of Date)
    Private m_IntervalTimeSpan As TimeSpan
    Private m_ScriptID As Integer
    Private m_ScriptScheduleID As Integer
    Private m_Type As ScheduleType
    Private m_IsDisabled As Boolean
    Private Shared m_LastRunTimeCache As New Generic.Dictionary(Of Integer, Date)

    Public Shared Function AccuracyPaddingSeconds() As Integer
        Return 3
    End Function

    Public Function GetNextRunDate(ByVal CurrentRunDate As Date) As Date
        Return CalculateNextRunDate(CurrentRunDate)
    End Function

    Private Function CalculateNextRunDate(ByVal RunTime As Date) As Date

        Dim datNow As Date = New Date(RunTime.Year, RunTime.Month, RunTime.Day, RunTime.Hour, RunTime.Minute, RunTime.Second)
        Dim datReturn, datTemp As Date
        Dim dow As Nullable(Of DayOfWeek)
        Dim intDay As Integer
        Dim datNextMonth As Date

        Select Case Type
            Case ScheduleType.OneTime
                datReturn = OneTimeDateTime
            Case ScheduleType.DailyWeekly
                DayWeekDays.Sort()

                For Each d As DayOfWeek In DayWeekDays
                    If d >= datNow.DayOfWeek Then
                        dow = d
                        Exit For
                    End If
                Next

                If Not dow.HasValue Then
                    dow = DayWeekDays(0)
                End If

                DayWeekTimes.Sort()

                If dow = datNow.DayOfWeek Then
                    datTemp = New Date(1753, 1, 1, datNow.Hour, datNow.Minute, datNow.Second)

                    For Each dat As Date In DayWeekTimes
                        If dat >= datTemp Then
                            datReturn = New Date(datNow.Year, datNow.Month, datNow.Day, dat.Hour, dat.Minute, dat.Second)
                            Exit For
                        End If
                    Next

                End If

                If datReturn = Date.MinValue Then

                    If dow.Value = datNow.DayOfWeek Then

                        dow = New Nullable(Of DayOfWeek)

                        For Each d As DayOfWeek In DayWeekDays
                            If d > datNow.DayOfWeek Then
                                dow = d
                                Exit For
                            End If
                        Next

                        If Not dow.HasValue Then
                            dow = DayWeekDays(0)
                        End If

                    End If

                    If dow.Value > datNow.DayOfWeek Then
                        datTemp = datNow.Date.AddDays(dow.Value - datNow.DayOfWeek)
                    Else
                        datTemp = datNow.Date.AddDays(DayOfWeek.Saturday - datNow.DayOfWeek + dow.Value).AddDays(1)
                    End If

                    datReturn = New Date(datTemp.Year, datTemp.Month, datTemp.Day, DayWeekTimes(0).Hour, DayWeekTimes(0).Minute, DayWeekTimes(0).Second)

                End If

            Case ScheduleType.Interval

                If IntervalStartDateTime <= datNow AndAlso m_LastRunTimeCache.ContainsKey(ScriptScheduleID) Then
                    datReturn = m_LastRunTimeCache(ScriptScheduleID).Add(IntervalTimeSpan)

                    If datReturn < datNow Then
                        datReturn = datNow
                    End If

                ElseIf IntervalStartDateTime <= datNow Then
                    datReturn = datNow
                Else
                    datReturn = IntervalStartDateTime
                End If

            Case ScheduleType.Monthly
                MonthDates.Sort()

                For Each d As Integer In MonthDates
                    If d >= datNow.Day Then
                        intDay = d
                        Exit For
                    End If
                Next

                If intDay = 0 Then
                    intDay = MonthDates(0)
                End If

                MonthTimes.Sort()

                If intDay = datNow.Day Then
                    datTemp = New Date(1753, 1, 1, datNow.Hour, datNow.Minute, datNow.Second)

                    For Each dat As Date In MonthTimes
                        If dat >= datTemp Then
                            datReturn = New Date(datNow.Year, datNow.Month, datNow.Day, dat.Hour, dat.Minute, dat.Second)
                            Exit For
                        End If
                    Next

                End If

                If datReturn = Date.MinValue Then

                    If intDay = datNow.Day Then

                        intDay = 0

                        For Each d As Integer In MonthDates
                            If d > datNow.Day Then
                                intDay = d
                                Exit For
                            End If
                        Next

                        If intDay = 0 Then
                            intDay = MonthDates(0)
                        End If

                    End If

                    If intDay > datNow.Day Then
                        If Date.DaysInMonth(datNow.Year, datNow.Month) >= intDay Then
                            datTemp = datNow.Date.AddDays(intDay - datNow.Day)
                        Else
                            datTemp = New Date(datNow.Year, datNow.Month, Date.DaysInMonth(datNow.Year, datNow.Month))
                        End If
                    Else
                        datNextMonth = New Date(datNow.Year, datNow.Month, 1).AddMonths(1)
                        If Date.DaysInMonth(datNextMonth.Year, datNextMonth.Month) >= intDay Then
                            datTemp = New Date(datNextMonth.Year, datNextMonth.Month, intDay)
                        Else
                            datTemp = New Date(datNextMonth.Year, datNextMonth.Month, Date.DaysInMonth(datNextMonth.Year, datNextMonth.Month))
                        End If

                    End If

                    datReturn = New Date(datTemp.Year, datTemp.Month, datTemp.Day, MonthTimes(0).Hour, MonthTimes(0).Minute, MonthTimes(0).Second)

                End If

        End Select

        If datReturn < datNow Then
            IsDisabled = True
        End If

        Return datReturn

    End Function

    Public Sub BeginRun(ByVal RunTime As Date)

        If Type = ScheduleType.Interval AndAlso ScriptScheduleID <> 0 Then
            m_LastRunTimeCache(ScriptScheduleID) = RunTime
        End If

        Form1.DoRunScriptSchedule(Me)

    End Sub

    Public Sub Run()

        Dim dt As DataTable

        dt = DB.ScriptEntry.ReadByScriptIDASCBySequence(ScriptID)

        For Each dr As DataRow In dt.Rows
            Commands.SendCommand(dr("command").ToString, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
        Next

    End Sub

    Public Property IsDisabled() As Boolean
        Get
            Return m_IsDisabled
        End Get
        Set(ByVal value As Boolean)
            m_IsDisabled = value
        End Set
    End Property

    Public Property Type() As ScheduleType
        Get
            Return m_Type
        End Get
        Set(ByVal value As ScheduleType)
            m_Type = value
        End Set
    End Property

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

    Public Property IntervalTimeSpan() As TimeSpan
        Get
            Return m_IntervalTimeSpan
        End Get
        Set(ByVal value As TimeSpan)
            m_IntervalTimeSpan = value
        End Set
    End Property

    Public ReadOnly Property MonthTimes() As List(Of Date)
        Get
            Return m_MonthTimes
        End Get
    End Property

    Public ReadOnly Property MonthDates() As List(Of Integer)
        Get
            Return m_MonthDates
        End Get
    End Property

    Public ReadOnly Property DayWeekTimes() As List(Of Date)
        Get
            Return m_DayWeekTimes
        End Get
    End Property

    Public ReadOnly Property DayWeekDays() As List(Of DayOfWeek)
        Get
            Return m_DayWeekDays
        End Get
    End Property

    Public Property IntervalStartDateTime() As Date
        Get
            Return m_IntervalStartDateTime
        End Get
        Set(ByVal value As Date)
            m_IntervalStartDateTime = value
        End Set
    End Property

    Public Property OneTimeDateTime() As Date
        Get
            Return m_OneTimeDateTime
        End Get
        Set(ByVal value As Date)
            m_OneTimeDateTime = value
        End Set
    End Property

    Public Function GetCode() As String

        Dim strCode As String

        Select Case Type
            Case ScheduleType.DailyWeekly
                ' 0_1_2
                ' W_DaysOfWeek2DigitsPerEntry_TimeOfDay4DigitsPerEntry
                ' W_0001020304050607_0000060512302345
                strCode = "W_"

                For Each d As DayOfWeek In DayWeekDays
                    strCode &= CInt(d).ToString.PadLeft(2, "0"c)
                Next

                strCode &= "_"

                For Each d As Date In DayWeekTimes
                    strCode &= d.Hour.ToString.PadLeft(2, "0"c) & d.Minute.ToString.PadLeft(2, "0"c)
                Next

            Case ScheduleType.Interval
                ' 0_1_2
                ' I_Days3DigitsHours2DigitsMinutes2DigitsSeconds2Digits_Year4DigitsMonth2DigitsDay2DigitsHour2DigitsMinute2Digits
                ' I_005230559_200902280149
                strCode = "I_"

                strCode &= IntervalTimeSpan.Days.ToString.PadLeft(3, "0"c)
                strCode &= IntervalTimeSpan.Hours.ToString.PadLeft(2, "0"c)
                strCode &= IntervalTimeSpan.Minutes.ToString.PadLeft(2, "0"c)
                strCode &= IntervalTimeSpan.Seconds.ToString.PadLeft(2, "0"c)

                strCode &= "_"

                strCode &= IntervalStartDateTime.Year.ToString.PadLeft(4, "0"c) & IntervalStartDateTime.Month.ToString.PadLeft(2, "0"c) & IntervalStartDateTime.Day.ToString.PadLeft(2, "0"c)
                strCode &= IntervalStartDateTime.Hour.ToString.PadLeft(2, "0"c) & IntervalStartDateTime.Minute.ToString.PadLeft(2, "0"c)

            Case ScheduleType.Monthly
                ' 0_1_2
                ' M_DaysOfMonth2DigitsPerEntry_TimeOfDay4DigitsPerEntry
                ' M_0105101531_0000060512302345
                strCode = "M_"

                For Each i As Integer In MonthDates
                    strCode &= CInt(i).ToString.PadLeft(2, "0"c)
                Next

                strCode &= "_"

                For Each d As Date In MonthTimes
                    strCode &= d.Hour.ToString.PadLeft(2, "0"c) & d.Minute.ToString.PadLeft(2, "0"c)
                Next

            Case ScheduleType.OneTime
                ' 0_1_2
                ' O_Year4DigitsMonth2DigitsDay2DigitsHour2DigitsMinute2Digits
                ' O_200902280149
                strCode = "O_"

                strCode &= OneTimeDateTime.Year.ToString.PadLeft(4, "0"c) & OneTimeDateTime.Month.ToString.PadLeft(2, "0"c) & OneTimeDateTime.Day.ToString.PadLeft(2, "0"c)
                strCode &= OneTimeDateTime.Hour.ToString.PadLeft(2, "0"c) & OneTimeDateTime.Minute.ToString.PadLeft(2, "0"c)

        End Select

        If IsDisabled Then
            strCode &= "X"
        End If

        Return strCode

    End Function

    Sub New(ByVal Code As String, ByVal ScriptID As Integer, ByVal ScriptScheduleID As Integer)

        Me.New(ScriptID, ScriptScheduleID)

        Dim strSplit() As String

        If Code.Contains("X") Then
            IsDisabled = True
            Code.Replace("X", "")
        End If

        strSplit = Code.Split("_"c)

        Select Case strSplit(0)
            Case "W"
                Type = ScheduleType.DailyWeekly

                ' Part 1
                For intX As Integer = 0 To strSplit(1).Length - 2 Step 2
                    DayWeekDays.Add(CType(CInt(strSplit(1).Substring(intX, 2)), DayOfWeek))
                Next

                ' Part 2
                For intX As Integer = 0 To strSplit(2).Length - 4 Step 4
                    DayWeekTimes.Add(New Date(1753, 1, 1, CInt(strSplit(2).Substring(intX, 2)), CInt(strSplit(2).Substring(intX + 2, 2)), 0))
                Next

            Case "I"
                Type = ScheduleType.Interval

                ' Part 1
                IntervalTimeSpan = New TimeSpan(CInt(strSplit(1).Substring(0, 3)), CInt(strSplit(1).Substring(3, 2)), CInt(strSplit(1).Substring(5, 2)), CInt(strSplit(1).Substring(7, 2)))

                ' Part 2
                IntervalStartDateTime = New Date(CInt(strSplit(2).Substring(0, 4)), CInt(strSplit(2).Substring(4, 2)), CInt(strSplit(2).Substring(6, 2)), CInt(strSplit(2).Substring(8, 2)), CInt(strSplit(2).Substring(10, 2)), 0)

            Case "M"
                Type = ScheduleType.Monthly

                ' Part 1
                For intX As Integer = 0 To strSplit(1).Length - 2 Step 2
                    MonthDates.Add(CInt(strSplit(1).Substring(intX, 2)))
                Next

                ' Part 2
                For intX As Integer = 0 To strSplit(2).Length - 4 Step 4
                    MonthTimes.Add(New Date(1753, 1, 1, CInt(strSplit(2).Substring(intX, 2)), CInt(strSplit(2).Substring(intX + 2, 2)), 0))
                Next

            Case "O"
                Type = ScheduleType.OneTime

                ' Part 1
                OneTimeDateTime = New Date(CInt(strSplit(1).Substring(0, 4)), CInt(strSplit(1).Substring(4, 2)), CInt(strSplit(1).Substring(6, 2)), CInt(strSplit(1).Substring(8, 2)), CInt(strSplit(1).Substring(10, 2)), 0)
        End Select

    End Sub

    Sub New(ByVal ScriptID As Integer, ByVal ScriptScheduleID As Integer)

        Me.ScriptID = ScriptID
        Me.ScriptScheduleID = ScriptScheduleID

    End Sub

    Public Enum ScheduleType
        Undefined = 0
        OneTime = 1
        DailyWeekly = 2
        Monthly = 3
        Interval = 4
    End Enum

    Public Class NextRunDateComparer : Implements IComparer(Of Schedule)


        Public Function Compare(ByVal x As Schedule, ByVal y As Schedule) As Integer Implements System.Collections.Generic.IComparer(Of Schedule).Compare

            Dim datX, datY As Date

            datX = x.GetNextRunDate(Now)
            datY = y.GetNextRunDate(Now)

            If datX < datY Then
                Return -1
            ElseIf datX > datY Then
                Return 1
            Else
                Return 0
            End If

        End Function
    End Class

End Class

Public Enum Channel
    Undefined = 0
    LeftFront = 1
    Center = 2
    RightFront = 3
    LeftSurround = 4
    RightSurround = 5
    LeftHeight = 13
    RightHeight = 14
    LeftWide = 15
    RightWide = 16
    LeftRear = 6
    RightRear = 7
    Subwoofer = 8
    Zone2Left = 9
    Zone2Right = 10
    Zone3Left = 11
    Zone3Right = 12
End Enum

Public Enum SourceType
    Undefined = 0
    TUNER = 1
    CD = 2
    BD = 3
    DVD = 4
    TV = 5
    SATCBL = 6
    GAME = 7
    GAME2 = 8
    VAUX = 9
    DOCK = 10
    NETUSB = 11
    SOURCE = 13
End Enum

Public Enum OnOffCommand
    Undefined = 0
    ToneCtrl = 1
    ToneCtrlOn = 101
    ToneCtrlOff = 201
    Power = 2
    PowerOn = 102
    PowerOff = 202
    CinemaEQ = 3
    CinemaEQOn = 103
    CinemaEQOff = 203
    AFD = 4
    AFDOn = 104
    AFDOff = 204
    SWAtt = 5
    SWAttOn = 105
    SWAttOff = 205
    SW = 6
    SWOn = 106
    SWOff = 206
    Panorama = 7
    PanoramaOn = 107
    PanoramaOff = 207
    ZoneMain = 8
    ZoneMainOn = 108
    ZoneMainOff = 208
    Zone2 = 9
    Zone2On = 109
    Zone2Off = 209
    Zone3 = 10
    Zone3On = 110
    Zone3Off = 210
    Zone2Mute = 11
    Zone2MuteOn = 111
    Zone2MuteOff = 211
    Zone3Mute = 12
    Zone3MuteOn = 112
    Zone3MuteOff = 212
    Zone2HPF = 13
    Zone2HPFOn = 113
    Zone2HPFOff = 213
    Zone3HPF = 14
    Zone3HPFOn = 114
    Zone3HPFOff = 214
    Mute = 15
    MuteOn = 115
    MuteOff = 215
    DSPEffect = 16
    DSPEffectOn = 116
    DSPEffectOff = 216
    DynamicEQ = 17
    DynamicEQOn = 117
    DynamicEQOff = 217
End Enum

Public Enum RangeControlType
	Undefined = 0
	Bass = 1
	Treble = 2
	LFE = 3
	Dimension = 4
	CenterWidth = 5
	CenterImage = 6
	Zone2Bass = 7
	Zone2Treble = 8
	Zone3Bass = 9
	Zone3Treble = 10
	DSPSurroundDelay = 11
	AudioDelay = 12
	Contrast = 13
	Brightness = 14
	Chroma = 15
	Hue = 16
	DSPEffectLevel = 17
End Enum

Public Enum QueryType
	Undefined = 0
	Full = 1
	Main = 2
	OtherParams = 3
	Zone2 = 4
	Sources = 5
	Zone3 = 6
	SurroundParams = 7
	VideoParams = 8
	Presets = 9
	Tuner = 10
	NET = 11
	iPod = 12
End Enum

Public Enum RangeChangeButtonType
	Undefined = 0
	Normal = 1
	MasterVolume = 2
	ChannelVolume = 3
End Enum

Public Enum Zones
	Undefined = 0
	ZoneMain = 1
	Zone2 = 2
	Zone3 = 3
End Enum

Public Enum MenuCommand
	Undefined = 0
	Up = 1
	Down = 2
	Left = 3
	Right = 4
	Enter = 5
	[Return] = 6
	GUIOn = 7
	GUIOff = 8
End Enum

Public Enum FixedListComboBox
	Undefined = 0
	DRC = 1
	DCOMP = 2
	RoomSize = 3
	Night = 4
	Restorer = 5
	Resolution = 6
	AspectRatio = 7
	Zone2ChannelSetting = 8
	Zone3ChannelSetting = 9
	AudioInputMode = 10
	DigitalAudioMode = 11
	TunerFrequency = 12
	SurroundMode = 13
	SurroundBackMode = 14
	SurroundSetting = 15
    AudysseyMode = 16
    DynamicVolume = 17
    ReferenceLevel = 18
End Enum

Public Enum NETCommand
	Undefined = 0
	Up = 1
	Down = 2
	Left = 3
	Right = 4
	EnterPlayPause = 5
	Play = 6
	Pause = 7
	[Stop] = 8
	PageUp = 9
	PageDown = 10
End Enum

Public Enum IPodCommand
	Undefined = 0
	Up = 1
	Down = 2
	Left = 3
	Right = 4
	EnterPlayPause = 5
	PlayPause = 6
	[Stop] = 8
	PageUp = 9
	PageDown = 10
End Enum

Public Enum RemoteAndPanelLock
	Undefined = 0
	RemoteUnlock = 1
	RemoteLock = 2
	PanelUnlock = 3
	PanelFullLock = 4
	PanelFullLockExceptMV = 5
End Enum