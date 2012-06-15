''''''''' ''''''''' ''''''''' ''''''''' ''''''''' 
''''''''' Copyright 2007-2009 Brian Saville
''''''''' Command3808 is free for non-commercial use.
''''''''' Reuse of the source code is free for non-commercial use
''''''''' provided that the source code used is credited to Command3808.
''''''''' ''''''''' ''''''''' ''''''''' '''''''''
Imports System.Data.SqlServerCe
Imports System.Data
Public Class DB


	Private Const CodeVersion As Integer = 3
	Private Shared m_DBConn As SqlCeConnection
	Private Shared m_DBVersion As Integer
	Private Shared m_DBVersionDate As String

	Public Class Setting

		Public Shared Sub Create(ByVal Name As String, ByVal Value As String)

			Dim sqlCmd As New SqlCeCommand

			sqlCmd.Parameters.AddWithValue("@name", Name)
			sqlCmd.Parameters.AddWithValue("@value", Value)

			Using Tran As New TransactionScope
				DB.ExecuteQuery(My.Resources.DBSetting_Create, sqlCmd)
				Tran.Complete()
			End Using

		End Sub

		Public Shared Function Exists(ByVal Name As String) As Boolean
			Dim sqlCmd As New SqlCeCommand
			sqlCmd.Parameters.AddWithValue("@name", Name)

			If CInt(DB.GetDataTable(My.Resources.DBSetting_Exists, sqlCmd).Rows(0)(0)) >= 1 Then
				Return True
			End If

		End Function

		Public Shared Function Read(ByVal Name As String) As String

			Dim sqlCmd As New SqlCeCommand
			sqlCmd.Parameters.AddWithValue("@name", Name)


			Return DB.GetDataTable(My.Resources.DBSetting_Read, sqlCmd).Rows(0)(0).ToString
		End Function

		Public Shared Sub Update(ByVal Name As String, ByVal Value As String)
			Dim sqlCmd As New SqlCeCommand

			sqlCmd.Parameters.AddWithValue("@name", Name)
			sqlCmd.Parameters.AddWithValue("@value", Value)

			Using Tran As New TransactionScope
				DB.ExecuteQuery(My.Resources.DBSetting_Update, sqlCmd)
				Tran.Complete()
			End Using

		End Sub

	End Class

	Public Class SystemInfo

		Public Shared Sub Create(ByVal Name As String, ByVal Value As String)

			Dim sqlCmd As New SqlCeCommand

			sqlCmd.Parameters.AddWithValue("@name", Name)
			sqlCmd.Parameters.AddWithValue("@value", Value)

			Using Tran As New TransactionScope
				DB.ExecuteQuery(My.Resources.DBSystemInfo_Create, sqlCmd)
				Tran.Complete()
			End Using

		End Sub

		Public Shared Function Exists(ByVal Name As String) As Boolean
			Dim sqlCmd As New SqlCeCommand
			sqlCmd.Parameters.AddWithValue("@name", Name)

			If CInt(DB.GetDataTable(My.Resources.DBSystemInfo_Exists, sqlCmd).Rows(0)(0)) >= 1 Then
				Return True
			End If

		End Function

		Public Shared Function Read(ByVal Name As String) As String

			Dim sqlCmd As New SqlCeCommand
			sqlCmd.Parameters.AddWithValue("@name", Name)


			Return DB.GetDataTable(My.Resources.DBSystemInfo_Read, sqlCmd).Rows(0)(0).ToString
		End Function

		Public Shared Sub Update(ByVal Name As String, ByVal Value As String)
			Dim sqlCmd As New SqlCeCommand

			sqlCmd.Parameters.AddWithValue("@name", Name)
			sqlCmd.Parameters.AddWithValue("@value", Value)

			Using Tran As New TransactionScope
				DB.ExecuteQuery(My.Resources.DBSystemInfo_Update, sqlCmd)
				Tran.Complete()
			End Using

		End Sub

	End Class

	Public Class Script

		Public Shared Sub CreateFromDR(ByVal dr As DataRow)

			Dim dt As DataTable

			Using Tran As New TransactionScope

				Create(dr("name").ToString)
				dt = GetMaxScriptID()

				For Each dr2 As DataRow In dr.GetChildRows("Script_ScriptSchedule")
					ScriptSchedule.Create(CInt(dt.Rows(0)(0)), dr2("code").ToString, dr2("name").ToString)
				Next

				For Each dr2 As DataRow In dr.GetChildRows("Script_ScriptEntry")
					ScriptEntry.Create(CInt(dt.Rows(0)(0)), dr2("command").ToString, dr2("description").ToString)
				Next

			End Using

		End Sub

		Public Shared Sub Create(ByVal Name As String)

			Dim sqlCmd As New SqlCeCommand

			sqlCmd.Parameters.AddWithValue("@name", Name)

			Using Tran As New TransactionScope
				DB.ExecuteQuery(My.Resources.DBScript_Create, sqlCmd)
				Tran.Complete()
			End Using

		End Sub

		Public Shared Function ReadAllASCByName() As DataTable

			Dim sqlCmd As New SqlCeCommand

			Return DB.GetDataTable(My.Resources.DBScript_ReadAllASCByName, sqlCmd)

		End Function

		Public Shared Function GetMaxScriptID() As DataTable

			Dim sqlCmd As New SqlCeCommand

			Return DB.GetDataTable(My.Resources.DBScript_GetMaxScriptID, sqlCmd)

		End Function

		Public Shared Function ReadScriptName(ByVal ScriptID As Integer) As String

			Dim sqlCmd As New SqlCeCommand

			sqlCmd.Parameters.AddWithValue("@ScriptID", ScriptID)

			Return DB.GetDataTable(My.Resources.DBScript_ReadScriptName, sqlCmd).Rows(0)(0).ToString

		End Function

		Public Shared Sub UpdateName(ByVal ScriptID As Integer, ByVal Name As String)
			Dim sqlCmd As New SqlCeCommand

			sqlCmd.Parameters.AddWithValue("@ScriptID", ScriptID)
			sqlCmd.Parameters.AddWithValue("@name", Name)

			Using Tran As New TransactionScope
				DB.ExecuteQuery(My.Resources.DBScript_UpdateName, sqlCmd)
				Tran.Complete()
			End Using

		End Sub

		Public Shared Sub CascadeDelete(ByVal ScriptID As Integer)
			ScriptEntry.DeleteByScriptID(ScriptID)
			ScriptSchedule.DeleteByScriptID(ScriptID)

			Dim sqlCmd As New SqlCeCommand

			sqlCmd.Parameters.AddWithValue("@ScriptID", ScriptID)

			Using Tran As New TransactionScope
				DB.ExecuteQuery(My.Resources.DBScript_Delete, sqlCmd)
				Tran.Complete()
			End Using

		End Sub

	End Class

	Public Class ScriptEntry

		Public Shared Sub Create(ByVal ScriptID As Integer, ByVal Command As String, ByVal Description As String)

			Dim sqlCmd As New SqlCeCommand
			Dim intX As Integer = 1
			Dim dt As DataTable

			sqlCmd.Parameters.AddWithValue("@ScriptID", ScriptID)
			sqlCmd.Parameters.AddWithValue("@command", Command)
			sqlCmd.Parameters.AddWithValue("@description", Description)

			Using Tran As New TransactionScope
				dt = DB.GetDataTable(My.Resources.DBScriptEntry_GetMaxSequenceNumberByScriptID, sqlCmd)

				If dt.Rows.Count > 0 AndAlso dt.Rows(0)(0) IsNot DBNull.Value Then
					intX = CInt(dt.Rows(0)(0))
				End If

				sqlCmd.Parameters.AddWithValue("@sequence", intX)

				DB.ExecuteQuery(My.Resources.DBScriptEntry_Create, sqlCmd)
				Tran.Complete()
			End Using

		End Sub

		Public Shared Function ReadAllASCBySequence() As DataTable

			Dim sqlCmd As New SqlCeCommand

			Return DB.GetDataTable(My.Resources.DBScriptEntry_ReadAllASCBySequence, sqlCmd)

		End Function

		Public Shared Function ReadByScriptIDASCBySequence(ByVal ScriptID As Integer) As DataTable

			Dim sqlCmd As New SqlCeCommand

			sqlCmd.Parameters.AddWithValue("@ScriptID", ScriptID)

			Return DB.GetDataTable(My.Resources.DBScriptEntry_ReadByScriptIDASCBySequence, sqlCmd)

		End Function

		Public Shared Sub UpdateDescription(ByVal ScriptEntryID As Integer, ByVal Description As String)
			Dim sqlCmd As New SqlCeCommand

			sqlCmd.Parameters.AddWithValue("@ScriptEntryID", ScriptEntryID)
			sqlCmd.Parameters.AddWithValue("@description", Description)

			Using Tran As New TransactionScope
				DB.ExecuteQuery(My.Resources.DBScriptEntry_UpdateDescription, sqlCmd)
				Tran.Complete()
			End Using

		End Sub

		Public Shared Sub UpdateSequence(ByVal SourceScriptEntryID As Integer, ByVal DestinationScriptEntryID As Integer, ByVal SourceSequence As Integer, ByVal DestinationSequence As Integer)
			Dim sqlCmd As New SqlCeCommand

			sqlCmd.Parameters.AddWithValue("@ScriptEntryID", SourceScriptEntryID)
			sqlCmd.Parameters.AddWithValue("@sequence", DestinationSequence)

			Using Tran As New TransactionScope
				DB.ExecuteQuery(My.Resources.DBScriptEntry_UpdateSequence, sqlCmd)

				sqlCmd.Parameters("@ScriptEntryID").Value = DestinationScriptEntryID
				sqlCmd.Parameters("@sequence").Value = SourceSequence

				DB.ExecuteQuery(My.Resources.DBScriptEntry_UpdateSequence, sqlCmd)

				Tran.Complete()
			End Using

		End Sub

		Public Shared Sub DeleteByScriptID(ByVal ScriptID As Integer)
			Dim sqlCmd As New SqlCeCommand

			sqlCmd.Parameters.AddWithValue("@ScriptID", ScriptID)

			Using Tran As New TransactionScope
				DB.ExecuteQuery(My.Resources.DBScriptEntry_DeleteByScriptID, sqlCmd)
				Tran.Complete()
			End Using

		End Sub

		Public Shared Sub Delete(ByVal ScriptEntryID As Integer)
			Dim sqlCmd As New SqlCeCommand

			sqlCmd.Parameters.AddWithValue("@ScriptEntryID", ScriptEntryID)

			Using Tran As New TransactionScope
				DB.ExecuteQuery(My.Resources.DBScriptEntry_Delete, sqlCmd)
				Tran.Complete()
			End Using

		End Sub

	End Class

	Public Class ScriptSchedule

		Public Shared Sub Create(ByVal ScriptID As Integer, ByVal Code As String, ByVal Name As String)
			Dim sqlCmd As New SqlCeCommand

			sqlCmd.Parameters.AddWithValue("@ScriptID", ScriptID)
			sqlCmd.Parameters.AddWithValue("@code", Code)
			sqlCmd.Parameters.AddWithValue("@name", Name)

			Using Tran As New TransactionScope
				DB.ExecuteQuery(My.Resources.DBScriptSchedule_Create, sqlCmd)
				Tran.Complete()
			End Using

		End Sub

		Public Shared Function ReadAll() As DataTable

			Dim sqlCmd As New SqlCeCommand

			Return DB.GetDataTable(My.Resources.DBScriptSchedule_ReadAll, sqlCmd)

		End Function

		Public Shared Sub DeleteByScriptID(ByVal ScriptID As Integer)
			Dim sqlCmd As New SqlCeCommand

			sqlCmd.Parameters.AddWithValue("@ScriptID", ScriptID)

			Using Tran As New TransactionScope
				DB.ExecuteQuery(My.Resources.DBScriptSchedule_DeleteByScriptID, sqlCmd)
				Tran.Complete()
			End Using

		End Sub

		Public Shared Sub Update(ByVal ScriptScheduleID As Integer, ByVal Code As String, ByVal Name As String)

			Dim sqlCmd As New SqlCeCommand

			sqlCmd.Parameters.AddWithValue("@ScriptScheduleID", ScriptScheduleID)
			sqlCmd.Parameters.AddWithValue("@code", Code)
			sqlCmd.Parameters.AddWithValue("@name", Name)

			Using Tran As New TransactionScope
				DB.ExecuteQuery(My.Resources.DBScriptSchedule_Update, sqlCmd)
				Tran.Complete()
			End Using

		End Sub

		Public Shared Sub Delete(ByVal ScriptScheduleID As Integer)
			Dim sqlCmd As New SqlCeCommand

			sqlCmd.Parameters.AddWithValue("@ScriptScheduleID", ScriptScheduleID)

			Using Tran As New TransactionScope
				DB.ExecuteQuery(My.Resources.DBScriptSchedule_Delete, sqlCmd)
				Tran.Complete()
			End Using

		End Sub

	End Class

	Private Shared Function ExecuteQuery(ByVal CommandText As String, ByVal sqlCmd As SqlCeCommand) As Integer

		Dim intResult As Integer

		Using Tran As New TransactionScope

			sqlCmd.CommandText = CommandText
			sqlCmd.Connection = DBConn

			intResult = sqlCmd.ExecuteNonQuery()

			Tran.Complete()

		End Using

		Return intResult

	End Function

	Private Shared Function GetDataTable(ByVal CommandText As String, ByVal sqlCmd As SqlCeCommand) As DataTable

		Dim dt As New DataTable
		Dim sqlAdap As New SqlCeDataAdapter

		sqlCmd.CommandText = CommandText
		sqlCmd.Connection = DBConn

		sqlAdap.SelectCommand = sqlCmd

		sqlAdap.Fill(dt)

		Return dt

	End Function

	Public Shared ReadOnly Property DBVersion() As Integer
		Get
			Return m_DBVersion
		End Get
	End Property

	Public Shared ReadOnly Property DBVersionDate() As String
		Get
			Return m_DBVersionDate
		End Get
	End Property

	Private Shared ReadOnly Property DBConn() As SqlCeConnection
		Get
			If IsNothing(m_DBConn) Then
				m_DBConn = New SqlCeConnection(My.Settings.Command3808ConnectionString)
			End If

			Return m_DBConn
		End Get
	End Property

	Private Shared Sub UpdateDBIfRequired()

		If Not System.IO.File.Exists(DBConn.Database) Then
			Dim sqlCeEngine As New SqlCeEngine(DBConn.ConnectionString)
			sqlCeEngine.CreateDatabase()
			DBConn.Open()
		Else
			DBConn.Open()
			' pre-upgrade version number
			m_DBVersion = CInt(DB.SystemInfo.Read("DBVersion"))
		End If

		iUpdateDBIfRequired()

	End Sub

	Private Shared Sub UpdateDBVersion(ByVal NewVersion As Integer)

		If DB.SystemInfo.Exists("DBVersionDate") Then
			DB.SystemInfo.Update("DBVersionDate", Helper.GetNowString)
		Else
			DB.SystemInfo.Create("DBVersionDate", Helper.GetNowString)
		End If

		If DB.SystemInfo.Exists("DBVersion") Then
			DB.SystemInfo.Update("DBVersion", NewVersion.ToString)
		Else
			DB.SystemInfo.Create("DBVersion", NewVersion.ToString)
		End If

	End Sub

	Private Shared Sub iUpdateDBIfRequired()

		If CodeVersion > DBVersion Then

			For intX As Integer = DBVersion + 1 To CodeVersion

				If SplashScreen1.Visible Then
					SplashScreen1.SetStatus("Upgrading Internal Database To v" & intX)
				End If

				' when adding a new version don't forget to change CodeVersion constant above
				Select Case intX
					Case 1
						RunScript(My.Resources.DBv1)
						My.CustomSettings.AutoClearDebug.Instance.CreateOrReset()
						My.CustomSettings.AutoClearLines.Instance.CreateOrReset()
						My.CustomSettings.AutoRefreshIPod.Instance.CreateOrReset()
						My.CustomSettings.AutoRefreshIPodSecs.Instance.CreateOrReset()
						My.CustomSettings.AutoRefreshNet.Instance.CreateOrReset()
						My.CustomSettings.AutoRefreshNetSecs.Instance.CreateOrReset()
						My.CustomSettings.DefaultCommandPauseMS.Instance.CreateOrReset()
						My.CustomSettings.IPAddress.Instance.CreateOrReset()
						My.CustomSettings.MinimizeToTray.Instance.CreateOrReset()
						My.CustomSettings.PreviewCommands.Instance.CreateOrReset()
						My.CustomSettings.SavedCommands.Instance.CreateOrReset()
						My.CustomSettings.Sources.Instance.CreateOrReset()
						My.CustomSettings.Tabs.Instance.CreateOrReset()
					Case 2
						' Do nothing
					Case 3
						RunScript(My.Resources.DBv3)
				End Select

				UpdateDBVersion(intX)

			Next

		End If

		m_DBVersion = CInt(DB.SystemInfo.Read("DBVersion"))
		m_DBVersionDate = DB.SystemInfo.Read("DBVersionDate")

	End Sub

	Private Shared Sub RunScript(ByVal Resource As String)

		Dim strCommands() As String = Resource.Split(New String() {";"}, StringSplitOptions.RemoveEmptyEntries)

		Dim sqlCmd As New SqlCeCommand
		sqlCmd.Connection = DBConn

		For Each strCommand As String In strCommands
			strCommand = strCommand.Trim
			If Not String.IsNullOrEmpty(strCommand) AndAlso strCommand.Substring(0, 2) <> "--" Then
				sqlCmd.CommandText = strCommand
				sqlCmd.ExecuteNonQuery()
			End If
		Next

	End Sub

	Public Shared Sub Reset()

		Dim strPath As String = DBConn.Database

		If DBConn IsNot Nothing Then
			DBConn.Close()
			DBConn.Dispose()
			m_DBConn = Nothing
		End If

		If System.IO.File.Exists(strPath) Then
			IO.File.Move(strPath, strPath.Replace(IO.Path.GetExtension(strPath), "_old_v" & DBVersion & "_" & Now.Ticks & ".sdf"))
		End If

		m_DBVersion = 0

		UpdateDBIfRequired()

	End Sub

	Shared Sub New()

		UpdateDBIfRequired()

	End Sub

	Protected Overrides Sub Finalize()
		MyBase.Finalize()
		DBConn.Close()
		DBConn.Dispose()
	End Sub



End Class
