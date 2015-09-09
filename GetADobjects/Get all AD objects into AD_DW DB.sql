USE [AD_DW]
GO

-- Set Active Directory domain to use.
DECLARE @ADpath nvarchar(64) = 'LDAP://DC=veca,DC=is';

-- Declare vars used.
DECLARE @table_name sysname;
DECLARE @tempTableName nvarchar(50), @TableName nvarchar(64),  @ADfilter nvarchar(4000), @CurrentOp nvarchar(50);
DECLARE @SQL nvarchar(max);
DECLARE @Members XML;

-------------------
-- Users
-------------------
SET @CurrentOp = 'Get Users';
SET @TableName = 'ADusers';
PRINT '------------------------------------------------------------------------';
PRINT @CurrentOp + ' from AD INTO ' + @TableName + ' table.';

-- Create (global) temp table dynamically. Note is global cuz of scope issue.
SET @table_name = 'dbo.' + @TableName;
SET @tempTableName = '##' + @TableName;
EXECUTE [dbo].[usp_GenerateTempTableScript] @table_name, @tempTableName, @SQL OUTPUT;
EXEC (@SQL);

-- Get all users from AD into temp table
SET @ADfilter = '(&(objectCategory=person)(objectClass=user))';
INSERT INTO ##ADusers EXEC dbo.clr_GetADobjects @ADpath, @ADfilter, @Members OUTPUT;

-- Insert into DeletedADusers table all rows from ADusers table that don't exist in the temp table.
-- (i.e. the rows that will be deleted by the MERGE statement.)
INSERT INTO [dbo].[DeletedADusers]
	SELECT a.*, GETDATE() AS [DeletedDate]
	FROM [AD_DW].[dbo].[ADusers] a
	WHERE ObjectGUID NOT IN (SELECT ObjectGUID FROM ##ADusers);

-- Generate MERGE statement dynamically from table definition.
EXECUTE [dbo].[usp_GenerateMergeStatement] @TableName, @tempTableName, @SQL OUTPUT;

-- Execute MERGE statement.
PRINT 'MERGE ' + @TableName;
EXECUTE (@SQL)

-- DROP temp table.
SET @SQL = 'DROP TABLE ' + @tempTableName + ';';
EXECUTE (@SQL)

-- UPDATE [ManagerGUID]
UPDATE [dbo].[ADusers] 
SET [ManagerGUID] = (SELECT [ObjectGUID] FROM [dbo].[ADusers] WHERE a.[Manager] = DistinguishedName)
FROM [dbo].[ADusers] a WHERE Manager IS NOT NULL

--Example query
--SELECT [DisplayName],[Title],[Department],[Company]
--	  ,CASE WHEN [ManagerGUID] IS NULL THEN '' 
--			WHEN [ManagerGUID] IS NOT NULL THEN (SELECT DisplayName FROM [dbo].[ADusers] WHERE a.[ManagerGUID] = [ObjectGUID])
--	  END AS ManagerName
--FROM [dbo].[ADusers] a WHERE ManagerGUID IS NOT NULL

-- UPDATE UserMustChangePasswordAtNextLogon
UPDATE [dbo].[ADusers] 
SET UserMustChangePasswordAtNextLogon = (CASE WHEN [PasswordLastSet] IS NULL AND [PasswordNeverExpires] = 0 AND [PasswordNotRequired] = 0 THEN 1 ELSE 0 END);

-------------------
-- Contacts
-------------------
SET @CurrentOp = 'Get Contacts';
SET @TableName = 'ADcontacts';
PRINT '------------------------------------------------------------------------';
PRINT @CurrentOp + ' from AD INTO ' + @TableName + ' table.';

-- Create (global) temp table dynamically. Note is global cuz of scope issue.
SET @table_name = 'dbo.' + @TableName;
SET @tempTableName = '##' + @TableName;
EXECUTE [dbo].[usp_GenerateTempTableScript] @table_name, @tempTableName, @SQL OUTPUT;
EXEC (@SQL);

-- Get all Contacts from AD into temp table
SET @ADfilter = '(&(objectCategory=person)(objectClass=contact))';
INSERT INTO ##ADcontacts EXEC dbo.clr_GetADobjects @ADpath, @ADfilter, @Members OUTPUT;

-- Generate MERGE statement dynamically from table definition.
EXECUTE [dbo].[usp_GenerateMergeStatement] @TableName, @tempTableName, @SQL OUTPUT

-- Execute MERGE statement.
PRINT 'MERGE ' + @TableName;
EXECUTE (@SQL)

-- DROP temp table.
SET @SQL = 'DROP TABLE ' + @tempTableName + ';';
EXECUTE (@SQL)


-------------------
-- Computers
-------------------
SET @CurrentOp = 'Get Computers';
SET @TableName = 'ADcomputers';
PRINT '------------------------------------------------------------------------';
PRINT @CurrentOp + ' from AD INTO ' + @TableName + ' table.';

-- Create (global) temp table dynamically. Note is global cuz of scope issue.
SET @table_name = 'dbo.' + @TableName;
SET @tempTableName = '##' + @TableName;
EXECUTE [dbo].[usp_GenerateTempTableScript] @table_name, @tempTableName, @SQL OUTPUT;
EXEC (@SQL);

-- Get all computers from AD into temp table
SET @ADfilter = '(objectCategory=computer)';
INSERT INTO ##ADcomputers EXEC dbo.clr_GetADobjects @ADpath, @ADfilter, @Members OUTPUT;

-- Generate MERGE statement dynamically from table definition.
EXECUTE [dbo].[usp_GenerateMergeStatement] @TableName, @tempTableName, @SQL OUTPUT

-- Execute MERGE statement.
PRINT 'MERGE ' + @TableName;
EXECUTE (@SQL)

-- DROP temp table.
SET @SQL = 'DROP TABLE ' + @tempTableName + ';';
EXECUTE (@SQL)


-------------------
-- Well known SIDs
-------------------
SET @CurrentOp = 'Get Well known SIDs';
SET @TableName = 'ADwell_known_sids';
PRINT '------------------------------------------------------------------------';
PRINT @CurrentOp + ' from AD INTO ' + @TableName + ' table.';

-- Create (global) temp table dynamically. Note is global cuz of scope issue.
SET @table_name = 'dbo.' + @TableName;
SET @tempTableName = '##' + @TableName;
EXECUTE [dbo].[usp_GenerateTempTableScript] @table_name, @tempTableName, @SQL OUTPUT;
EXEC (@SQL);

-- Get all Well known SIDs from AD into temp table
SET @ADfilter = '(|(objectSID=S-1-0)(objectSID=S-1-0-0)(objectSID=S-1-1)(objectSID=S-1-1-0)(objectSID=S-1-2)(objectSID=S-1-2-0)(objectSID=S-1-2-1)(objectSID=S-1-3)(objectSID=S-1-3-0)(objectSID=S-1-3-1)(objectSID=S-1-3-2)(objectSID=S-1-3-3)(objectSID=S-1-3-4)(objectSID=S-1-5-80-0)(objectSID=S-1-4)(objectSID=S-1-5)(objectSID=S-1-5-1)(objectSID=S-1-5-2)(objectSID=S-1-5-3)(objectSID=S-1-5-4)(objectSID=S-1-5-6)(objectSID=S-1-5-7)(objectSID=S-1-5-8)(objectSID=S-1-5-9)(objectSID=S-1-5-10)(objectSID=S-1-5-11)(objectSID=S-1-5-12)(objectSID=S-1-5-13)(objectSID=S-1-5-14)(objectSID=S-1-5-15)(objectSID=S-1-5-17)(objectSID=S-1-5-18)(objectSID=S-1-5-19)(objectSID=S-1-5-20)(objectSID=S-1-5-64-10)(objectSID=S-1-5-64-14)(objectSID=S-1-5-64-21)(objectSID=S-1-5-80)(objectSID=S-1-5-80-0)(objectSID=S-1-5-83-0)(objectSID=S-1-16-0)(objectSID=S-1-16-4096)(objectSID=S-1-16-8192)(objectSID=S-1-16-8448)(objectSID=S-1-16-12288)(objectSID=S-1-16-16384)(objectSID=S-1-16-20480)(objectSID=S-1-16-28672))';
INSERT INTO ##ADwell_known_sids EXEC dbo.clr_GetADobjects @ADpath, @ADfilter, @Members OUTPUT;

-- Get DisplayName and Description from lookup table
UPDATE ##ADwell_known_sids
SET ##ADwell_known_sids.[DisplayName] = L.[DisplayName],
	##ADwell_known_sids.[Description] = L.[Description]
FROM ##ADwell_known_sids W
INNER JOIN dbo.ADwell_known_sids_lookup L ON W.[SID] = L.[SID] COLLATE Latin1_General_CI_AS

-- Generate MERGE statement dynamically from table definition.
EXECUTE [dbo].[usp_GenerateMergeStatement] @TableName, @tempTableName, @SQL OUTPUT

-- Execute MERGE statement.
PRINT 'MERGE ' + @TableName;
EXECUTE (@SQL)

-- DROP temp table.
SET @SQL = 'DROP TABLE ' + @tempTableName + ';';
EXECUTE (@SQL)


-------------------
-- Groups
-------------------
SET @CurrentOp = 'Get Groups';
SET @TableName = 'ADgroups';
PRINT '------------------------------------------------------------------------';
PRINT @CurrentOp + ' from AD INTO ' + @TableName + ' table.';

-- Create (global) temp table dynamically. Note is global cuz of scope issue.
SET @table_name = 'dbo.' + @TableName;
SET @tempTableName = '##' + @TableName;
EXECUTE [dbo].[usp_GenerateTempTableScript] @table_name, @tempTableName, @SQL OUTPUT;
EXEC (@SQL);

-- Get all groups from AD into temp table - Note group members returned in @Members as XML.
SET @ADfilter = '(objectCategory=group)';
INSERT INTO ##ADgroups EXEC dbo.clr_GetADobjects @ADpath, @ADfilter, @Members OUTPUT;

-- Generate MERGE statement dynamically from table definition.
EXECUTE [dbo].[usp_GenerateMergeStatement] @TableName, @tempTableName, @SQL OUTPUT

-- Execute MERGE statement.
PRINT 'MERGE ' + @TableName;
EXECUTE (@SQL)

-- DROP temp table.
SET @SQL = 'DROP TABLE ' + @tempTableName + ';';
EXECUTE (@SQL)


PRINT 'Get Group members from XML data.';

-- Create (global) temp table dynamically. Note is global cuz of scope issue.
SET @table_name = 'dbo.ADgroup_members';
SET @tempTableName = '##ADgroup_members';
EXECUTE [dbo].[usp_GenerateTempTableScript] @table_name, @tempTableName, @SQL OUTPUT;
EXEC (@SQL);

-- INSERT group members into temp table.
WITH MemberList AS ( -- Process group members XML data.
select Tg.Cg.value('@GrpDS', 'nvarchar(256)') as GroupDS,
		Tm.Cm.value('@MemberDS', 'nvarchar(256)') as MemberDS
from @Members.nodes('/body') as Tb(Cb)
  outer apply Tb.Cb.nodes('Group') as Tg(Cg)
  outer apply Tg.Cg.nodes('Member') AS Tm(Cm)
)
INSERT INTO ##ADgroup_members 
  SELECT 
	G.objectGUID AS GroupGUID,
	COALESCE(U.objectGUID, GM.objectGUID, C.objectGUID, CN.objectGUID, W.ObjectGUID) AS MemberGUID,
	COALESCE(U.ObjectClass, GM.ObjectClass, C.ObjectClass, CN.ObjectClass, W.ObjectClass) AS MemberType,
	M.GroupDS AS GroupDistinguishedName,
	M.MemberDS AS [MemberDistinguishedName]
FROM MemberList M
LEFT JOIN dbo.ADgroups G ON M.GroupDS = G.DistinguishedName
LEFT JOIN dbo.ADgroups GM ON M.MemberDS = GM.DistinguishedName
LEFT JOIN dbo.ADcomputers C ON M.MemberDS = C.DistinguishedName
LEFT JOIN dbo.ADusers U ON M.MemberDS = U.DistinguishedName
LEFT JOIN dbo.ADcontacts CN ON M.MemberDS = CN.DistinguishedName
LEFT JOIN dbo.ADwell_known_sids W ON M.MemberDS = W.DistinguishedName;

PRINT 'MERGE Group members';
MERGE dbo.ADgroup_members WITH (HOLDLOCK) AS T
USING ##ADgroup_members AS S 
ON (T.GroupGUID = S.GroupGUID AND T.MemberGUID = S.MemberGUID) 
WHEN MATCHED THEN 
UPDATE SET 
T.MemberType = S.MemberType,
T.[GroupDistinguishedName] = S.[GroupDistinguishedName],
T.[MemberDistinguishedName] = S.[MemberDistinguishedName]
WHEN NOT MATCHED BY TARGET THEN 
INSERT (GroupGUID, MemberGUID, MemberType, [GroupDistinguishedName], [MemberDistinguishedName]) 
VALUES (S.GroupGUID, S.MemberGUID, S.MemberType, S.[GroupDistinguishedName], S.[MemberDistinguishedName]) 
WHEN NOT MATCHED BY SOURCE THEN 
DELETE;

DROP TABLE ##ADgroup_members;

-------------------
-- Users photos
-------------------
SET @CurrentOp = 'Get Users photos';
SET @TableName = 'ADusersPhotos';
PRINT '------------------------------------------------------------------------';
PRINT @CurrentOp + ' from AD INTO ' + @TableName + ' table.';

-- Create (global) temp table dynamically. Note is global cuz of scope issue.
SET @table_name = 'dbo.' + @TableName;
SET @tempTableName = '##' + @TableName;
EXECUTE [dbo].[usp_GenerateTempTableScript] @table_name, @tempTableName, @SQL OUTPUT;
EXEC (@SQL);

-- Get all photos from AD into temp table
SET @ADfilter = '(&(objectCategory=person)(objectClass=user))';
INSERT INTO ##ADusersPhotos EXEC clr_GetADusersPhotos @ADpath, @ADfilter;

-- Generate MERGE statement dynamically from table definition.
EXECUTE [dbo].[usp_GenerateMergeStatement] @TableName, @tempTableName, @SQL OUTPUT

-- Execute MERGE statement.
PRINT 'MERGE ' + @TableName;
EXECUTE (@SQL)

-- DROP temp table.
SET @SQL = 'DROP TABLE ' + @tempTableName + ';';
EXECUTE (@SQL)


