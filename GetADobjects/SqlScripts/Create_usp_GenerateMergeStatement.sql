USE [AD_DW]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		Snorri Kristjansson
-- Create date: 05.09.2015
-- Description:	Generate MERGE statement from table definition.
-- =============================================
CREATE PROCEDURE [dbo].[usp_GenerateMergeStatement]
	@TableName nvarchar(64),
	@tempTableName nvarchar(64),
	@MergeSQL nvarchar(max) OUTPUT
AS
BEGIN
	SET NOCOUNT ON;
	SET @MergeSQL = '';

	-- Get table columns as CSV
	DECLARE @CSV nvarchar(4000);
	SET @CSV = (SELECT LEFT(ColNames, LEN(ColNames) - 1)
	FROM (
	select '[' + c.name + '],' from sys.columns c 
	join sys.tables t on c.object_id = t.object_id
	where t.name = @TableName
	order by c.column_id
	FOR XML PATH('')
	) c (ColNames)) + CHAR(13);

	-- Get table columns as CSV formatted for MERGE INSERT clause
	DECLARE @CSV2 nvarchar(4000);
	SET @CSV2 = (SELECT LEFT(ColNames, LEN(ColNames) - 1)
	FROM (
	select 'T.[' + c.name + '] = S.[' + c.name + '],' from sys.columns c 
	join sys.tables t on c.object_id = t.object_id
	where t.name = @TableName AND c.name != 'ObjectGUID'
	order by c.column_id
	FOR XML PATH('')
	) c (ColNames)) + CHAR(13);

	-- Get table columns as CSV formatted for MERGE UPDATE clause - VALUES
	DECLARE @CSV3 nvarchar(4000);
	SET @CSV3 = (SELECT LEFT(ColNames, LEN(ColNames) - 1)
	FROM (
	select 'S.[' + c.name + '],' from sys.columns c 
	join sys.tables t on c.object_id = t.object_id
	where t.name = @TableName
	order by c.column_id
	FOR XML PATH('')
	) c (ColNames)) + CHAR(13);

	DECLARE @M1 nvarchar(1000) =
	  'MERGE dbo.' + @TableName + ' WITH (HOLDLOCK) AS T' + CHAR(13)
	+ 'USING ' + @tempTableName + ' AS S ' + CHAR(13)
	+ 'ON (T.ObjectGUID = S.ObjectGUID) ' + CHAR(13)
	+ 'WHEN MATCHED THEN ' + CHAR(13)
	+ 'UPDATE SET ' + CHAR(13);
	DECLARE @M2 nvarchar(1000) =
	  'WHEN NOT MATCHED BY TARGET THEN ' + CHAR(13)
	+ 'INSERT (' + CHAR(13);
	DECLARE @M3 nvarchar(1000) =
	  ') ' + CHAR(13)
	+ 'VALUES (' + CHAR(13);
	DECLARE @M4 nvarchar(1000) =
	+ ') ' + CHAR(13)
	+ 'WHEN NOT MATCHED BY SOURCE THEN ' + CHAR(13)
	+ 'DELETE;'
	SET @MergeSQL =
	  CAST(@M1 AS nvarchar(max))	-- Need to cast to nvarchar(max) because otherwise output will be truncated to 4000 chars.
	+ CAST(@CSV2 AS nvarchar(max))
	+ CAST(@M2 AS nvarchar(max))
	+ CAST(@CSV AS nvarchar(max))
	+ CAST(@M3 AS nvarchar(max))
	+ CAST(@CSV3 AS nvarchar(max))
	+ CAST(@M4 AS nvarchar(max));
END

GO


