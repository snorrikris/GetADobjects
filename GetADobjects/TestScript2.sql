CREATE TABLE #ADgroups(
[Name] [nvarchar](256) NULL,
[GroupCategory] [nvarchar](32) NULL,
[GroupScope] [nvarchar](32) NULL,
[Description] [nvarchar](256) NULL,
[EmailAddress] [nvarchar](256) NULL,
[ManagedBy] [nvarchar](256) NULL,
[DistinguishedName] [nvarchar](512) NULL,
[DisplayName] [nvarchar](128) NULL,
[Created] [datetime] NULL,
[Modified] [datetime] NULL,
[SamAccountName] [nvarchar](128) NULL,
[ObjectCategory] [nvarchar](128) NULL,
[ObjectClass] [nvarchar](64) NULL,
[SID] [nvarchar](128) NULL,
[ObjectGUID] [uniqueidentifier] NOT NULL);

DECLARE @ADpath nvarchar(64) = 'LDAP://DC=veca,DC=is';
DECLARE @ADfilter nvarchar(64) = '(objectCategory=group)';
DECLARE @Members XML;
INSERT INTO #ADgroups EXEC dbo.clr_GetADobjects @ADpath, @ADfilter, @Members OUTPUT;

SELECT * FROM #ADgroups;

