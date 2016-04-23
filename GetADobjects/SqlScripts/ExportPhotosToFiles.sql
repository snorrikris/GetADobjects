-- Note - Ole Automation Procedures must be enabled for this script to work.
-- See: https://msdn.microsoft.com/en-us/library/ms191188.aspx

-- Set Domain path.
DECLARE @ADpath nvarchar(64) = 'LDAP://DC=contoso,DC=com';

-- Set file path. (note the SQL service user must have (at least) write permission to this folder).
DECLARE	@FilePath varchar(128) = 'C:\Photos\';

CREATE TABLE #ADusers(
	[GivenName] [nvarchar](128) NULL,
	[Initials] [nvarchar](32) NULL,
	[Surname] [nvarchar](256) NULL,
	[DisplayName] [nvarchar](128) NULL,
	[Description] [nvarchar](256) NULL,
	[Office] [nvarchar](128) NULL,
	[OfficePhone] [nvarchar](32) NULL,
	[EmailAddress] [nvarchar](256) NULL,
	[HomePage] [nvarchar](128) NULL,
	[StreetAddress] [nvarchar](160) NULL,
	[POBox] [nvarchar](32) NULL,
	[City] [nvarchar](128) NULL,
	[State] [nvarchar](64) NULL,
	[PostalCode] [nvarchar](32) NULL,
	[Country] [nvarchar](128) NULL,
	[HomePhone] [nvarchar](32) NULL,
	[Pager] [nvarchar](32) NULL,
	[MobilePhone] [nvarchar](32) NULL,
	[Fax] [nvarchar](128) NULL,
	[Title] [nvarchar](128) NULL,
	[Department] [nvarchar](128) NULL,
	[Company] [nvarchar](128) NULL,
	[Manager] [nvarchar](256) NULL,
	[EmployeeID] [nvarchar](64) NULL,
	[EmployeeNumber] [nvarchar](64) NULL,
	[Division] [nvarchar](128) NULL,
	[Enabled] [bit] NULL,
	[LockedOut] [bit] NULL,
	[MNSLogonAccount] [bit] NULL,
	[CannotChangePassword] [bit] NULL,
	[PasswordExpired] [bit] NULL,
	[PasswordNeverExpires] [bit] NULL,
	[PasswordNotRequired] [bit] NULL,
	[SmartcardLogonRequired] [bit] NULL,
	[DoesNotRequirePreAuth] [bit] NULL,
	[AllowReversiblePasswordEncryption] [bit] NULL,
	[AccountNotDelegated] [bit] NULL,
	[TrustedForDelegation] [bit] NULL,
	[TrustedToAuthForDelegation] [bit] NULL,
	[UseDESKeyOnly] [bit] NULL,
	[UserMustChangePasswordAtNextLogon] [bit] NULL,
	[HomedirRequired] [bit] NULL,
	[LastBadPasswordAttempt] [datetime] NULL,
	[BadLogonCount] [int] NULL,
	[LastLogonDate] [datetime] NULL,
	[LogonCount] [int] NULL,
	[PasswordLastSet] [datetime] NULL,
	[PasswordExpiryTime] [datetime] NULL,
	[AccountLockoutTime] [datetime] NULL,
	[AccountExpirationDate] [datetime] NULL,
	[LogonWorkstations] [nvarchar](128) NULL,
	[HomeDirectory] [nvarchar](128) NULL,
	[HomeDrive] [nvarchar](64) NULL,
	[ProfilePath] [nvarchar](128) NULL,
	[ScriptPath] [nvarchar](256) NULL,
	[userAccountControl] [int] NULL,
	[PrimaryGroupID] [int] NULL,
	[Name] [nvarchar](256) NULL,
	[CN] [nvarchar](256) NULL,
	[UserPrincipalName] [nvarchar](256) NULL,
	[SamAccountName] [nvarchar](128) NULL,
	[DistinguishedName] [nvarchar](512) NULL,
	[Created] [datetime] NULL,
	[Modified] [datetime] NULL,
	[ObjectCategory] [nvarchar](128) NULL,
	[ObjectClass] [nvarchar](64) NULL,
	[SID] [nvarchar](128) NOT NULL,
	[ObjectGUID] [uniqueidentifier] NOT NULL,
	[ManagerGUID] [uniqueidentifier] NULL,
);

CREATE TABLE #ADusersPhotos (
	[ObjectGUID] [uniqueidentifier] NOT NULL,
	[Width] [int] NULL,
	[Height] [int] NULL,
	[Format] [nvarchar](6),
	[Photo] [varbinary](max) NULL
);

-- Insert all users from AD into temp table.
DECLARE @Members XML;
DECLARE @ADfilter nvarchar(4000) = '(&(objectCategory=person)(objectClass=user))';
INSERT INTO #ADusers EXEC clr_GetADobjects @ADpath, @ADfilter, @Members OUTPUT;

-- Insert photos into global temp table (using same AD filter).
INSERT INTO #ADusersPhotos EXEC clr_GetADusersPhotos @ADpath, @ADfilter;

DECLARE	 @Command       VARCHAR(4000), 
         @Format		nvarchar(6),
		 @ObjectGUID	uniqueidentifier,
         @ImageFileName VARCHAR(128),
		 @ImageData		varbinary(max),
		 @ObjectToken	INT,
 		 @SamAccountName nvarchar(128),
		 @DisplayName	nvarchar(128);

DECLARE curPhotoImage CURSOR FOR             -- Cursor for each image in table
	SELECT p.[Format],
		   p.[ObjectGUID],
		   p.[Photo],
		   u.SamAccountName,
		   u.DisplayName
	FROM   #ADusersPhotos p
	JOIN   #ADusers u ON p.ObjectGUID = u.ObjectGUID
	WHERE  p.[Photo] IS NOT NULL;

OPEN curPhotoImage; 

FETCH NEXT FROM curPhotoImage 
	INTO @Format, @ObjectGUID, @ImageData, @SamAccountName, @DisplayName;

WHILE (@@FETCH_STATUS = 0) -- Cursor loop  
  BEGIN 
	SET @ImageFileName = @FilePath + @DisplayName + ' (' + @SamAccountName + ').' + @Format;

	-- Source: http://stackoverflow.com/questions/4056050/script-to-save-varbinary-data-to-disk
	EXEC sp_OACreate 'ADODB.Stream', @ObjectToken OUTPUT;
	EXEC sp_OASetProperty @ObjectToken, 'Type', 1;
	EXEC sp_OAMethod @ObjectToken, 'Open';
	EXEC sp_OAMethod @ObjectToken, 'Write', NULL, @ImageData;
	EXEC sp_OAMethod @ObjectToken, 'SaveToFile', NULL, @ImageFileName, 2;
	EXEC sp_OAMethod @ObjectToken, 'Close';
	EXEC sp_OADestroy @ObjectToken;
     
    FETCH NEXT FROM curPhotoImage 
		INTO @Format, @ObjectGUID, @ImageData, @SamAccountName, @DisplayName;
  END  -- cursor loop 
 
CLOSE curPhotoImage;
DEALLOCATE curPhotoImage;

DROP TABLE #ADusers;
DROP TABLE #ADusersPhotos;
