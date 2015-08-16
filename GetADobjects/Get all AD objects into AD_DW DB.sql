USE [AD_DW]
GO

-- Set Active Directory domain to use.
DECLARE @ADpath nvarchar(64) = 'LDAP://DC=veca,DC=is';

-------------------
-- Users
-------------------

-- Create ##ADgroups (global) temp table dynamically. Note is global cuz of scope issue.
DECLARE @table_name sysname = 'dbo.ADusers';
DECLARE @tempTableName nvarchar(50) = '##ADusers';
DECLARE @SQL nvarchar(max) = '';
EXECUTE [dbo].[usp_GenerateTempTableScript] @table_name, @tempTableName, @SQL OUTPUT;
EXEC (@SQL);

-- Get all groups from AD into temp table
DECLARE @ADfilter nvarchar(64) = '(&(objectCategory=person)(objectClass=user))';
DECLARE @Members XML;
INSERT INTO ##ADusers EXEC dbo.clr_GetADobjects @ADpath, @ADfilter, @Members OUTPUT;

-- Generate MERGE statement dynamically from table definition.
DECLARE @TableName nvarchar(64) = 'ADusers';
SET @SQL = '';
EXECUTE [dbo].[usp_GenerateMergeStatement] @TableName, @tempTableName, @SQL OUTPUT

-- Execute MERGE statement.
EXECUTE (@SQL)

-- DROP temp table.
SET @SQL = 'DROP TABLE ' + @tempTableName + ';';
EXECUTE (@SQL)


-------------------
-- Contacts
-------------------

-- Create (global) temp table dynamically. Note is global cuz of scope issue.
SET @table_name = 'dbo.ADcontacts';
SET @tempTableName = '##ADcontacts';
SET @SQL = '';
EXECUTE [dbo].[usp_GenerateTempTableScript] @table_name, @tempTableName, @SQL OUTPUT;
EXEC (@SQL);

-- Get all groups from AD into temp table
SET @ADfilter = '(&(objectCategory=person)(objectClass=contact))';
INSERT INTO ##ADcontacts EXEC dbo.clr_GetADobjects @ADpath, @ADfilter, @Members OUTPUT;

-- Generate MERGE statement dynamically from table definition.
SET @TableName = 'ADcontacts';
SET @SQL = '';
EXECUTE [dbo].[usp_GenerateMergeStatement] @TableName, @tempTableName, @SQL OUTPUT

-- Execute MERGE statement.
EXECUTE (@SQL)

-- DROP temp table.
SET @SQL = 'DROP TABLE ' + @tempTableName + ';';
EXECUTE (@SQL)


-------------------
-- Computers
-------------------

-- Create (global) temp table dynamically. Note is global cuz of scope issue.
SET @table_name = 'dbo.ADcomputers';
SET @tempTableName = '##ADcomputers';
SET @SQL = '';
EXECUTE [dbo].[usp_GenerateTempTableScript] @table_name, @tempTableName, @SQL OUTPUT;
EXEC (@SQL);

-- Get all groups from AD into temp table
SET @ADfilter = '(objectCategory=computer)';
INSERT INTO ##ADcomputers EXEC dbo.clr_GetADobjects @ADpath, @ADfilter, @Members OUTPUT;

-- Generate MERGE statement dynamically from table definition.
SET @TableName = 'ADcomputers';
SET @SQL = '';
EXECUTE [dbo].[usp_GenerateMergeStatement] @TableName, @tempTableName, @SQL OUTPUT

-- Execute MERGE statement.
EXECUTE (@SQL)

-- DROP temp table.
SET @SQL = 'DROP TABLE ' + @tempTableName + ';';
EXECUTE (@SQL)


-------------------
-- Well known SIDs
-------------------

-- Create (global) temp table dynamically. Note is global cuz of scope issue.
SET @table_name = 'dbo.ADwell_known_sids';
SET @tempTableName = '##ADwell_known_sids';
SET @SQL = '';
EXECUTE [dbo].[usp_GenerateTempTableScript] @table_name, @tempTableName, @SQL OUTPUT;
EXEC (@SQL);

-- Get all groups from AD into temp table
DECLARE @ADfilterX nvarchar(4000) = '(|(objectSID=S-1-0)(objectSID=S-1-0-0)(objectSID=S-1-1)(objectSID=S-1-1-0)(objectSID=S-1-2)(objectSID=S-1-2-0)(objectSID=S-1-2-1)(objectSID=S-1-3)(objectSID=S-1-3-0)(objectSID=S-1-3-1)(objectSID=S-1-3-2)(objectSID=S-1-3-3)(objectSID=S-1-3-4)(objectSID=S-1-5-80-0)(objectSID=S-1-4)(objectSID=S-1-5)(objectSID=S-1-5-1)(objectSID=S-1-5-2)(objectSID=S-1-5-3)(objectSID=S-1-5-4)(objectSID=S-1-5-6)(objectSID=S-1-5-7)(objectSID=S-1-5-8)(objectSID=S-1-5-9)(objectSID=S-1-5-10)(objectSID=S-1-5-11)(objectSID=S-1-5-12)(objectSID=S-1-5-13)(objectSID=S-1-5-14)(objectSID=S-1-5-15)(objectSID=S-1-5-17)(objectSID=S-1-5-18)(objectSID=S-1-5-19)(objectSID=S-1-5-20)(objectSID=S-1-5-64-10)(objectSID=S-1-5-64-14)(objectSID=S-1-5-64-21)(objectSID=S-1-5-80)(objectSID=S-1-5-80-0)(objectSID=S-1-5-83-0)(objectSID=S-1-16-0)(objectSID=S-1-16-4096)(objectSID=S-1-16-8192)(objectSID=S-1-16-8448)(objectSID=S-1-16-12288)(objectSID=S-1-16-16384)(objectSID=S-1-16-20480)(objectSID=S-1-16-28672))';
INSERT INTO ##ADwell_known_sids EXEC dbo.clr_GetADobjects @ADpath, @ADfilterX, @Members OUTPUT;

-- Get DisplayName and Description from lookup table
UPDATE ##ADwell_known_sids
SET ##ADwell_known_sids.[DisplayName] = L.[DisplayName],
	##ADwell_known_sids.[Description] = L.[Description]
FROM ##ADwell_known_sids W
INNER JOIN dbo.ADwell_known_sids_lookup L ON W.[SID] = L.[SID] COLLATE Latin1_General_CI_AS

-- Generate MERGE statement dynamically from table definition.
SET @TableName = 'ADwell_known_sids';
SET @SQL = '';
EXECUTE [dbo].[usp_GenerateMergeStatement] @TableName, @tempTableName, @SQL OUTPUT

-- Execute MERGE statement.
EXECUTE (@SQL)

-- DROP temp table.
SET @SQL = 'DROP TABLE ' + @tempTableName + ';';
EXECUTE (@SQL)


-------------------
-- Groups
-------------------

-- Create (global) temp table dynamically. Note is global cuz of scope issue.
SET @table_name = 'dbo.ADgroups';
SET @tempTableName = '##ADgroups';
SET @SQL = '';
EXECUTE [dbo].[usp_GenerateTempTableScript] @table_name, @tempTableName, @SQL OUTPUT;
EXEC (@SQL);

-- Get all groups from AD into temp table
SET @ADfilter = '(objectCategory=group)';
INSERT INTO ##ADgroups EXEC dbo.clr_GetADobjects @ADpath, @ADfilter, @Members OUTPUT;

-- Generate MERGE statement dynamically from table definition.
SET @TableName = 'ADgroups';
SET @SQL = '';
EXECUTE [dbo].[usp_GenerateMergeStatement] @TableName, @tempTableName, @SQL OUTPUT

-- Execute MERGE statement.
EXECUTE (@SQL)

-- DROP temp table.
SET @SQL = 'DROP TABLE ' + @tempTableName + ';';
EXECUTE (@SQL)

-- Create temp table for group members
IF OBJECT_ID('tempdb..#ADgroup_members') IS NOT NULL DROP TABLE #ADgroup_members;
CREATE TABLE #ADgroup_members(
[GroupGUID] [uniqueidentifier] NOT NULL,
[MemberGUID] [uniqueidentifier] NOT NULL,
[MemberType] [nvarchar](64) NOT NULL);

-- INSERT group members into temp table.
WITH MemberList AS ( -- Process group members XML data.
select Tg.Cg.value('@GrpDS', 'nvarchar(256)') as GroupDS,
		Tm.Cm.value('@MemberDS', 'nvarchar(256)') as MemberDS
from @Members.nodes('/body') as Tb(Cb)
  outer apply Tb.Cb.nodes('Group') as Tg(Cg)
  outer apply Tg.Cg.nodes('Member') AS Tm(Cm)
)
INSERT INTO #ADgroup_members 
  SELECT 
	G.objectGUID AS GroupGUID,
	COALESCE(U.objectGUID, GM.objectGUID, C.objectGUID, CN.objectGUID, W.ObjectGUID) AS MemberGUID,
	COALESCE(U.ObjectClass, GM.ObjectClass, C.ObjectClass, CN.ObjectClass, W.ObjectClass) AS MemberType
FROM MemberList M
LEFT JOIN dbo.ADgroups G ON M.GroupDS = G.DistinguishedName
LEFT JOIN dbo.ADgroups GM ON M.MemberDS = GM.DistinguishedName
LEFT JOIN dbo.ADcomputers C ON M.MemberDS = C.DistinguishedName
LEFT JOIN dbo.ADusers U ON M.MemberDS = U.DistinguishedName
LEFT JOIN dbo.ADcontacts CN ON M.MemberDS = CN.DistinguishedName
LEFT JOIN dbo.ADwell_known_sids W ON M.MemberDS = W.DistinguishedName;

MERGE dbo.ADgroup_members WITH (HOLDLOCK) AS T
USING #ADgroup_members AS S 
ON (T.GroupGUID = S.GroupGUID AND T.MemberGUID = S.MemberGUID) 
WHEN MATCHED THEN 
UPDATE SET 
T.MemberType = S.MemberType
WHEN NOT MATCHED BY TARGET THEN 
INSERT (
GroupGUID, MemberGUID, MemberType
) 
VALUES (
S.GroupGUID, S.MemberGUID, S.MemberType
) 
WHEN NOT MATCHED BY SOURCE THEN 
DELETE;

DROP TABLE #ADgroup_members;
