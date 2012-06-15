Imports System.Data.SqlServerCe

Namespace My

	'This class allows you to handle specific events on the settings class:
	' The SettingChanging event is raised before a setting's value is changed.
	' The PropertyChanged event is raised after a setting's value is changed.
	' The SettingsLoaded event is raised after the setting values are loaded.
	' The SettingsSaving event is raised before the setting values are saved.
	Partial Friend NotInheritable Class MySettings



		Private Sub MySettings_SettingsLoaded(ByVal sender As Object, ByVal e As System.Configuration.SettingsLoadedEventArgs) Handles Me.SettingsLoaded

		End Sub
	End Class

	Namespace CustomSettings

		Public MustInherit Class CustomSettingBase(Of T)

			Private m_ValueIsSet As Boolean
			Private m_DefaultValue As T
			Private m_SettingName As String
			Private m_Value As T

			Public Sub CreateOrReset()

				If DB.Setting.Exists(m_SettingName) Then
					Value = m_DefaultValue
				Else
					Create(m_DefaultValue)
				End If

			End Sub

			Public Property Value() As T
				Get

					If Not m_ValueIsSet Then
						m_Value = Read()
						m_ValueIsSet = True
					End If

					Return m_Value
				End Get
				Set(ByVal value As T)

					m_Value = value
					Update(value)
					m_ValueIsSet = True
				End Set
			End Property

			' Private Members

			Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs)
				Value = m_Value
			End Sub

			Private Sub Update(ByVal Value As T)
				DB.Setting.Update(m_SettingName, Serialize(Value))
			End Sub

			Private Sub Create(ByVal Value As T)
				m_Value = Value
				DB.Setting.Create(m_SettingName, Serialize(Value))
				m_ValueIsSet = True
			End Sub

			Private Function Read() As T
				Return Deserialize(DB.Setting.Read(m_SettingName))
			End Function

			Private Function Serialize(ByVal o As T) As String

				Dim bf As New Runtime.Serialization.Formatters.Binary.BinaryFormatter
				Dim ms As New IO.MemoryStream
				Dim byt() As Byte
				Dim strX As String

				bf.Serialize(ms, o)

				byt = ms.GetBuffer

				Array.Resize(byt, CInt(ms.Length))

				strX = Convert.ToBase64String(byt)

				ms.Close()
				ms.Dispose()

				Return strX

			End Function

			Private Function Deserialize(ByVal data As String) As T

				Dim o As T
				Dim bf As New Runtime.Serialization.Formatters.Binary.BinaryFormatter
				Dim ms As New IO.MemoryStream
				Dim byt() As Byte

				byt = Convert.FromBase64String(data)

				ms.Write(byt, 0, byt.Length)

				ms.Position = 0

				Try
					o = DirectCast(bf.Deserialize(ms), T)
				Catch ex As Exception
					Throw ex
				End Try

				ms.Close()
				ms.Dispose()

				Return o

			End Function

			Protected Sub New(ByVal SettingName As String, ByVal DefaultValue As T)
				m_SettingName = SettingName
				m_DefaultValue = DefaultValue
				AddHandler Form1.FormClosing, AddressOf Form1_FormClosing
			End Sub

		End Class

		Public Class AutoClearDebug : Inherits CustomSettingBase(Of Boolean)

			Protected Shared m_Instance As AutoClearDebug

			Public Shared Function Instance() As AutoClearDebug

				If m_Instance Is Nothing Then
					m_Instance = DirectCast(System.Activator.CreateInstance(Reflection.MethodBase.GetCurrentMethod.DeclaringType, True), AutoClearDebug)
				End If

				Return m_Instance

			End Function

			Private Sub New()
				MyBase.New(Reflection.MethodBase.GetCurrentMethod.DeclaringType.Name, True)
			End Sub

		End Class

		Public Class AutoClearLines : Inherits CustomSettingBase(Of Integer)

			Protected Shared m_Instance As AutoClearLines

			Public Shared Function Instance() As AutoClearLines

				If m_Instance Is Nothing Then
					m_Instance = DirectCast(System.Activator.CreateInstance(Reflection.MethodBase.GetCurrentMethod.DeclaringType, True), AutoClearLines)
				End If

				Return m_Instance

			End Function

			Private Sub New()
				MyBase.New(Reflection.MethodBase.GetCurrentMethod.DeclaringType.Name, 100)
			End Sub

		End Class

		Public Class AutoRefreshIPod : Inherits CustomSettingBase(Of Boolean)

			Protected Shared m_Instance As AutoRefreshIPod

			Public Shared Function Instance() As AutoRefreshIPod

				If m_Instance Is Nothing Then
					m_Instance = DirectCast(System.Activator.CreateInstance(Reflection.MethodBase.GetCurrentMethod.DeclaringType, True), AutoRefreshIPod)
				End If

				Return m_Instance

			End Function

			Private Sub New()
				MyBase.New(Reflection.MethodBase.GetCurrentMethod.DeclaringType.Name, False)
			End Sub

		End Class

		Public Class AutoRefreshIPodSecs : Inherits CustomSettingBase(Of Integer)

			Protected Shared m_Instance As AutoRefreshIPodSecs

			Public Shared Function Instance() As AutoRefreshIPodSecs

				If m_Instance Is Nothing Then
					m_Instance = DirectCast(System.Activator.CreateInstance(Reflection.MethodBase.GetCurrentMethod.DeclaringType, True), AutoRefreshIPodSecs)
				End If

				Return m_Instance

			End Function

			Private Sub New()
				MyBase.New(Reflection.MethodBase.GetCurrentMethod.DeclaringType.Name, 5)
			End Sub

		End Class

		Public Class AutoRefreshNet : Inherits CustomSettingBase(Of Boolean)

			Protected Shared m_Instance As AutoRefreshNet

			Public Shared Function Instance() As AutoRefreshNet

				If m_Instance Is Nothing Then
					m_Instance = DirectCast(System.Activator.CreateInstance(Reflection.MethodBase.GetCurrentMethod.DeclaringType, True), AutoRefreshNet)
				End If

				Return m_Instance

			End Function

			Private Sub New()
				MyBase.New(Reflection.MethodBase.GetCurrentMethod.DeclaringType.Name, False)
			End Sub

		End Class

		Public Class AutoRefreshNetSecs : Inherits CustomSettingBase(Of Integer)

			Protected Shared m_Instance As AutoRefreshNetSecs

			Public Shared Function Instance() As AutoRefreshNetSecs

				If m_Instance Is Nothing Then
					m_Instance = DirectCast(System.Activator.CreateInstance(Reflection.MethodBase.GetCurrentMethod.DeclaringType, True), AutoRefreshNetSecs)
				End If

				Return m_Instance

			End Function

			Private Sub New()
				MyBase.New(Reflection.MethodBase.GetCurrentMethod.DeclaringType.Name, 5)
			End Sub

		End Class

		Public Class DefaultCommandPauseMS : Inherits CustomSettingBase(Of Integer)

			Protected Shared m_Instance As DefaultCommandPauseMS

			Public Shared Function Instance() As DefaultCommandPauseMS

				If m_Instance Is Nothing Then
					m_Instance = DirectCast(System.Activator.CreateInstance(Reflection.MethodBase.GetCurrentMethod.DeclaringType, True), DefaultCommandPauseMS)
				End If

				Return m_Instance

			End Function

			Private Sub New()
				MyBase.New(Reflection.MethodBase.GetCurrentMethod.DeclaringType.Name, 200)
			End Sub

		End Class

		Public Class IPAddress : Inherits CustomSettingBase(Of String)

			Protected Shared m_Instance As IPAddress

			Public Shared Function Instance() As IPAddress

				If m_Instance Is Nothing Then
					m_Instance = DirectCast(System.Activator.CreateInstance(Reflection.MethodBase.GetCurrentMethod.DeclaringType, True), IPAddress)
				End If

				Return m_Instance

			End Function

			Private Sub New()
				MyBase.New(Reflection.MethodBase.GetCurrentMethod.DeclaringType.Name, "10.0.0.0")
			End Sub

		End Class

		Public Class MinimizeToTray : Inherits CustomSettingBase(Of Boolean)

			Protected Shared m_Instance As MinimizeToTray

			Public Shared Function Instance() As MinimizeToTray

				If m_Instance Is Nothing Then
					m_Instance = DirectCast(System.Activator.CreateInstance(Reflection.MethodBase.GetCurrentMethod.DeclaringType, True), MinimizeToTray)
				End If

				Return m_Instance

			End Function

			Private Sub New()
				MyBase.New(Reflection.MethodBase.GetCurrentMethod.DeclaringType.Name, False)
			End Sub

		End Class

		Public Class PreviewCommands : Inherits CustomSettingBase(Of Boolean)

			Protected Shared m_Instance As PreviewCommands

			Public Shared Function Instance() As PreviewCommands

				If m_Instance Is Nothing Then
					m_Instance = DirectCast(System.Activator.CreateInstance(Reflection.MethodBase.GetCurrentMethod.DeclaringType, True), PreviewCommands)
				End If

				Return m_Instance

			End Function

			Private Sub New()
				MyBase.New(Reflection.MethodBase.GetCurrentMethod.DeclaringType.Name, False)
			End Sub

		End Class

		Public Class SavedCommands : Inherits CustomSettingBase(Of Specialized.StringCollection)

			Protected Shared m_Instance As SavedCommands

			Public Shared Function Instance() As SavedCommands

				If m_Instance Is Nothing Then
					m_Instance = DirectCast(System.Activator.CreateInstance(Reflection.MethodBase.GetCurrentMethod.DeclaringType, True), SavedCommands)
				End If

				Return m_Instance

			End Function

			Private Sub New()
				MyBase.New(Reflection.MethodBase.GetCurrentMethod.DeclaringType.Name, New Specialized.StringCollection)
			End Sub

		End Class

		Public Class Sources : Inherits CustomSettingBase(Of SortedList(Of String, Source))

			Protected Shared m_Instance As Sources

			Public Shared Function Instance() As Sources

				If m_Instance Is Nothing Then
					m_Instance = DirectCast(System.Activator.CreateInstance(Reflection.MethodBase.GetCurrentMethod.DeclaringType, True), Sources)
				End If

				Return m_Instance

			End Function

			Private Sub New()
				MyBase.New(Reflection.MethodBase.GetCurrentMethod.DeclaringType.Name, New SortedList(Of String, Source))
			End Sub

		End Class

		Public Class Tabs : Inherits CustomSettingBase(Of System.ComponentModel.BindingList(Of TabPageInfo))

			Protected Shared m_Instance As Tabs

			Public Shared Function Instance() As Tabs

				If m_Instance Is Nothing Then
					m_Instance = DirectCast(System.Activator.CreateInstance(Reflection.MethodBase.GetCurrentMethod.DeclaringType, True), Tabs)
				End If

				Return m_Instance

			End Function

			Private Sub New()
				MyBase.New(Reflection.MethodBase.GetCurrentMethod.DeclaringType.Name, New System.ComponentModel.BindingList(Of TabPageInfo))
			End Sub

		End Class

	End Namespace

End Namespace
