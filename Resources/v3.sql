-- v3;
-- put a semi-colon after each command including comments;
-- only put comments at start of line. and only use the double hyphen format for comments;
-- do not include any blank lines;
CREATE TABLE Script
(
	ScriptID int IDENTITY NOT NULL CONSTRAINT Script_PK PRIMARY KEY
	, name nvarchar (100) NOT NULL
)
;
CREATE TABLE ScriptEntry
(
	ScriptEntryID int IDENTITY NOT NULL CONSTRAINT ScriptEntry_PK PRIMARY KEY
	, ScriptID int NOT NULL REFERENCES Script (ScriptID)
	, sequence smallint NOT NULL
	, command nvarchar (250) NOT NULL
	, description nvarchar (250) NOT NULL
)
;
CREATE TABLE ScriptSchedule
(
	ScriptScheduleID int IDENTITY NOT NULL CONSTRAINT ScriptSchedule_PK PRIMARY KEY
	, ScriptID int NOT NULL REFERENCES Script (ScriptID)
	, code nvarchar (250) NOT NULL
	, name nvarchar (100) NOT NULL
)
;