﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.17379
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System

Namespace My.Resources
    
    'This class was auto-generated by the StronglyTypedResourceBuilder
    'class via a tool like ResGen or Visual Studio.
    'To add or remove a member, edit your .ResX file then rerun ResGen
    'with the /str option, or rebuild your VS project.
    '''<summary>
    '''  A strongly-typed resource class, for looking up localized strings, etc.
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0"),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.Microsoft.VisualBasic.HideModuleNameAttribute()>  _
    Friend Module Resources
        
        Private resourceMan As Global.System.Resources.ResourceManager
        
        Private resourceCulture As Global.System.Globalization.CultureInfo
        
        '''<summary>
        '''  Returns the cached ResourceManager instance used by this class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend ReadOnly Property ResourceManager() As Global.System.Resources.ResourceManager
            Get
                If Object.ReferenceEquals(resourceMan, Nothing) Then
                    Dim temp As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager("Command2112.Resources", GetType(Resources).Assembly)
                    resourceMan = temp
                End If
                Return resourceMan
            End Get
        End Property
        
        '''<summary>
        '''  Overrides the current thread's CurrentUICulture property for all
        '''  resource lookups using this strongly typed resource class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend Property Culture() As Global.System.Globalization.CultureInfo
            Get
                Return resourceCulture
            End Get
            Set
                resourceCulture = value
            End Set
        End Property
        
        Friend ReadOnly Property AVR2112CI() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("AVR2112CI", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        Friend ReadOnly Property AVR2112CIsm() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("AVR2112CIsm", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to INSERT INTO Script (name) VALUES (@name).
        '''</summary>
        Friend ReadOnly Property DBScript_Create() As String
            Get
                Return ResourceManager.GetString("DBScript_Create", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to DELETE FROM Script WHERE ScriptID = @ScriptID.
        '''</summary>
        Friend ReadOnly Property DBScript_Delete() As String
            Get
                Return ResourceManager.GetString("DBScript_Delete", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to SELECT MAX(ScriptID) FROM Script.
        '''</summary>
        Friend ReadOnly Property DBScript_GetMaxScriptID() As String
            Get
                Return ResourceManager.GetString("DBScript_GetMaxScriptID", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to SELECT ScriptID, name FROM Script ORDER BY name.
        '''</summary>
        Friend ReadOnly Property DBScript_ReadAllASCByName() As String
            Get
                Return ResourceManager.GetString("DBScript_ReadAllASCByName", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to SELECT name FROM Script WHERE ScriptID = @ScriptID.
        '''</summary>
        Friend ReadOnly Property DBScript_ReadScriptName() As String
            Get
                Return ResourceManager.GetString("DBScript_ReadScriptName", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to UPDATE Script SET name = @name WHERE ScriptID = @ScriptID.
        '''</summary>
        Friend ReadOnly Property DBScript_UpdateName() As String
            Get
                Return ResourceManager.GetString("DBScript_UpdateName", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to INSERT INTO ScriptEntry (ScriptID, sequence, command, description) VALUES (@ScriptID, @sequence, @command, @description).
        '''</summary>
        Friend ReadOnly Property DBScriptEntry_Create() As String
            Get
                Return ResourceManager.GetString("DBScriptEntry_Create", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to DELETE FROM ScriptEntry WHERE ScriptEntryID = @ScriptEntryID.
        '''</summary>
        Friend ReadOnly Property DBScriptEntry_Delete() As String
            Get
                Return ResourceManager.GetString("DBScriptEntry_Delete", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to DELETE FROM ScriptEntry WHERE ScriptID = @ScriptID.
        '''</summary>
        Friend ReadOnly Property DBScriptEntry_DeleteByScriptID() As String
            Get
                Return ResourceManager.GetString("DBScriptEntry_DeleteByScriptID", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to SELECT MAX (sequence) + 1 FROM ScriptEntry WHERE ScriptID = @ScriptID.
        '''</summary>
        Friend ReadOnly Property DBScriptEntry_GetMaxSequenceNumberByScriptID() As String
            Get
                Return ResourceManager.GetString("DBScriptEntry_GetMaxSequenceNumberByScriptID", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to SELECT ScriptEntryID, ScriptID, sequence, command, description FROM ScriptEntry ORDER BY ScriptID, sequence.
        '''</summary>
        Friend ReadOnly Property DBScriptEntry_ReadAllASCBySequence() As String
            Get
                Return ResourceManager.GetString("DBScriptEntry_ReadAllASCBySequence", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to SELECT ScriptEntryID, ScriptID, sequence, command, description FROM ScriptEntry WHERE ScriptID = @ScriptID ORDER BY sequence.
        '''</summary>
        Friend ReadOnly Property DBScriptEntry_ReadByScriptIDASCBySequence() As String
            Get
                Return ResourceManager.GetString("DBScriptEntry_ReadByScriptIDASCBySequence", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to UPDATE ScriptEntry SET description = @description WHERE ScriptEntryID = @ScriptEntryID.
        '''</summary>
        Friend ReadOnly Property DBScriptEntry_UpdateDescription() As String
            Get
                Return ResourceManager.GetString("DBScriptEntry_UpdateDescription", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to UPDATE ScriptEntry SET sequence = @sequence WHERE ScriptEntryID = @ScriptEntryID.
        '''</summary>
        Friend ReadOnly Property DBScriptEntry_UpdateSequence() As String
            Get
                Return ResourceManager.GetString("DBScriptEntry_UpdateSequence", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to INSERT INTO ScriptSchedule (ScriptID, code, name) VALUES (@ScriptID, @code, @name).
        '''</summary>
        Friend ReadOnly Property DBScriptSchedule_Create() As String
            Get
                Return ResourceManager.GetString("DBScriptSchedule_Create", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to DELETE FROM ScriptSchedule WHERE ScriptScheduleID = @ScriptScheduleID.
        '''</summary>
        Friend ReadOnly Property DBScriptSchedule_Delete() As String
            Get
                Return ResourceManager.GetString("DBScriptSchedule_Delete", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to DELETE FROM ScriptSchedule WHERE ScriptID = @ScriptID.
        '''</summary>
        Friend ReadOnly Property DBScriptSchedule_DeleteByScriptID() As String
            Get
                Return ResourceManager.GetString("DBScriptSchedule_DeleteByScriptID", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to SELECT ScriptScheduleID, ScriptID, code, name FROM ScriptSchedule ORDER BY ScriptID, name.
        '''</summary>
        Friend ReadOnly Property DBScriptSchedule_ReadAll() As String
            Get
                Return ResourceManager.GetString("DBScriptSchedule_ReadAll", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to UPDATE ScriptSchedule SET name = @name, code = @code WHERE ScriptScheduleID = @ScriptScheduleID.
        '''</summary>
        Friend ReadOnly Property DBScriptSchedule_Update() As String
            Get
                Return ResourceManager.GetString("DBScriptSchedule_Update", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to INSERT INTO Setting (name, value) VALUES (@name, @value).
        '''</summary>
        Friend ReadOnly Property DBSetting_Create() As String
            Get
                Return ResourceManager.GetString("DBSetting_Create", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to SELECT COUNT(*) FROM Setting WHERE name = @name.
        '''</summary>
        Friend ReadOnly Property DBSetting_Exists() As String
            Get
                Return ResourceManager.GetString("DBSetting_Exists", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to SELECT value FROM Setting WHERE name = @name.
        '''</summary>
        Friend ReadOnly Property DBSetting_Read() As String
            Get
                Return ResourceManager.GetString("DBSetting_Read", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to UPDATE Setting SET value = @value WHERE name = @name.
        '''</summary>
        Friend ReadOnly Property DBSetting_Update() As String
            Get
                Return ResourceManager.GetString("DBSetting_Update", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to INSERT INTO SystemInfo (name, value) VALUES (@name, @value).
        '''</summary>
        Friend ReadOnly Property DBSystemInfo_Create() As String
            Get
                Return ResourceManager.GetString("DBSystemInfo_Create", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to SELECT COUNT(*) FROM SystemInfo WHERE name = @name.
        '''</summary>
        Friend ReadOnly Property DBSystemInfo_Exists() As String
            Get
                Return ResourceManager.GetString("DBSystemInfo_Exists", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to SELECT value FROM SystemInfo WHERE name = @name.
        '''</summary>
        Friend ReadOnly Property DBSystemInfo_Read() As String
            Get
                Return ResourceManager.GetString("DBSystemInfo_Read", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to UPDATE SystemInfo SET value = @value WHERE name = @name.
        '''</summary>
        Friend ReadOnly Property DBSystemInfo_Update() As String
            Get
                Return ResourceManager.GetString("DBSystemInfo_Update", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to -- v1;
        '''-- put a semi-colon after each command including comments;
        '''-- only put comments at start of line. and only use the double hyphen format for comments;
        '''-- do not include any blank lines;
        '''CREATE TABLE Setting
        '''(
        '''	name nvarchar (100) NOT NULL CONSTRAINT Setting_PK PRIMARY KEY
        '''	, value ntext NOT NULL
        '''
        ''')
        ''';
        '''CREATE TABLE SystemInfo
        '''(
        '''	name nvarchar (100) NOT NULL CONSTRAINT SystemInfo_PK PRIMARY KEY
        '''	, value ntext NOT NULL
        '''
        ''')
        ''';.
        '''</summary>
        Friend ReadOnly Property DBv1() As String
            Get
                Return ResourceManager.GetString("DBv1", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to -- v3;
        '''-- put a semi-colon after each command including comments;
        '''-- only put comments at start of line. and only use the double hyphen format for comments;
        '''-- do not include any blank lines;
        '''CREATE TABLE Script
        '''(
        '''	ScriptID int IDENTITY NOT NULL CONSTRAINT Script_PK PRIMARY KEY
        '''	, name nvarchar (100) NOT NULL
        ''')
        ''';
        '''CREATE TABLE ScriptEntry
        '''(
        '''	ScriptEntryID int IDENTITY NOT NULL CONSTRAINT ScriptEntry_PK PRIMARY KEY
        '''	, ScriptID int NOT NULL REFERENCES Script (ScriptID)
        '''	, sequence smallint NOT NU [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property DBv3() As String
            Get
                Return ResourceManager.GetString("DBv3", resourceCulture)
            End Get
        End Property
        
        Friend ReadOnly Property denon() As System.Drawing.Icon
            Get
                Dim obj As Object = ResourceManager.GetObject("denon", resourceCulture)
                Return CType(obj,System.Drawing.Icon)
            End Get
        End Property
        
        Friend ReadOnly Property PauseRecorderHS() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("PauseRecorderHS", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        Friend ReadOnly Property Play() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("Play", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        Friend ReadOnly Property PlayHS() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("PlayHS", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        Friend ReadOnly Property RecordHS() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("RecordHS", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        Friend ReadOnly Property StopHS() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("StopHS", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
    End Module
End Namespace
