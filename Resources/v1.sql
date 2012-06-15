-- v1;
-- put a semi-colon after each command including comments;
-- only put comments at start of line. and only use the double hyphen format for comments;
-- do not include any blank lines;
CREATE TABLE Setting
(
	name nvarchar (100) NOT NULL CONSTRAINT Setting_PK PRIMARY KEY
	, value ntext NOT NULL

)
;
CREATE TABLE SystemInfo
(
	name nvarchar (100) NOT NULL CONSTRAINT SystemInfo_PK PRIMARY KEY
	, value ntext NOT NULL

)
;