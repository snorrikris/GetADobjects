-- Set Domain path.
DECLARE @ADpath nvarchar(64) = 'LDAP://DC=veca,DC=is';

PRINT 'Create temp tables.'
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
	[ManagerGUID] [uniqueidentifier] NULL
);


PRINT 'Get all users from AD into temp table'
DECLARE @Members XML;
DECLARE @ADfilter nvarchar(64) = '(&(objectCategory=person)(objectClass=user))';
--INSERT INTO #ADusers 
EXEC clr_GetADobjects @ADpath, @ADfilter, @Members OUTPUT;
SELECT * FROM #ADusers;

PRINT 'DROP temp tables.'
DROP TABLE #ADusers;
