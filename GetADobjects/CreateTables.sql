USE [AD_DW]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ADusers](
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
CONSTRAINT [PK_UserGUID] PRIMARY KEY CLUSTERED 
(
	[ObjectGUID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ADcontacts](
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
[DistinguishedName] [nvarchar](512) NULL,
[Name] [nvarchar](256) NULL,
[CN] [nvarchar](256) NULL,
[Created] [datetime] NULL,
[Modified] [datetime] NULL,
[ObjectCategory] [nvarchar](128) NULL,
[ObjectClass] [nvarchar](64) NULL,
[ObjectGUID] [uniqueidentifier] NOT NULL,
CONSTRAINT [PK_ContactGUID] PRIMARY KEY CLUSTERED 
(
	[ObjectGUID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ADcomputers](
[Name] [nvarchar](256) NULL,
[DNSHostName] [nvarchar](128) NULL,
[Description] [nvarchar](256) NULL,
[Location] [nvarchar](128) NULL,
[OperatingSystem] [nvarchar](64) NULL,
[OperatingSystemVersion] [nvarchar](16) NULL,
[OperatingSystemServicePack] [nvarchar](32) NULL,
[ManagedBy] [nvarchar](256) NULL,
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
[LastBadPasswordAttempt] [datetime] NULL,
[BadLogonCount] [int] NULL,
[LastLogonDate] [datetime] NULL,
[logonCount] [int] NULL,
[PasswordLastSet] [datetime] NULL,
[AccountLockoutTime] [datetime] NULL,
[AccountExpirationDate] [datetime] NULL,
[Created] [datetime] NULL,
[Modified] [datetime] NULL,
[CN] [nvarchar](256) NULL,
[DisplayName] [nvarchar](128) NULL,
[DistinguishedName] [nvarchar](512) NULL,
[PrimaryGroupID] [int] NULL,
[SamAccountName] [nvarchar](128) NULL,
[userAccountControl] [int] NULL,
[ObjectCategory] [nvarchar](128) NULL,
[ObjectClass] [nvarchar](64) NULL,
[SID] [nvarchar](128) NOT NULL,
[ObjectGUID] [uniqueidentifier] NOT NULL,
CONSTRAINT [PK_ComputerGUID] PRIMARY KEY CLUSTERED 
(
	[ObjectGUID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ADgroups](
[Name] [nvarchar](256) NULL,
[GroupCategory] [nvarchar](32) NULL,
[GroupScope] [nvarchar](32) NULL,
[Description] [nvarchar](512) NULL,
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
[ObjectGUID] [uniqueidentifier] NOT NULL,
CONSTRAINT [PK_GroupGUID] PRIMARY KEY CLUSTERED 
(
	[ObjectGUID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ADgroup_members](
[GroupGUID] [uniqueidentifier] NOT NULL,
[MemberGUID] [uniqueidentifier] NOT NULL,
[MemberType] [nvarchar](16) NOT NULL,
CONSTRAINT [PK_GroupMemberGUIDs] PRIMARY KEY CLUSTERED 
(
	[GroupGUID] ASC, [MemberType] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[ADgroup_members] WITH CHECK ADD FOREIGN KEY([GroupGUID])
REFERENCES [dbo].[ADgroups] ([ObjectGUID])
GO

CREATE TABLE [dbo].[ADgroup_user_members](
	[GroupGUID] [uniqueidentifier] NOT NULL,
	[UserGUID] [uniqueidentifier] NOT NULL,
 CONSTRAINT [PK_GroupUserMemberGUIDs] PRIMARY KEY CLUSTERED 
(
	[GroupGUID] ASC,
	[UserGUID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

ALTER TABLE [dbo].[ADgroup_user_members]  WITH CHECK ADD FOREIGN KEY([GroupGUID])
REFERENCES [dbo].[ADgroups] ([ObjectGUID])
GO

ALTER TABLE [dbo].[ADgroup_user_members]  WITH CHECK ADD FOREIGN KEY([UserGUID])
REFERENCES [dbo].[ADusers] ([ObjectGUID])
GO
