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
[MemberType] [nvarchar](64) NOT NULL,
[GroupDistinguishedName] [nvarchar](512) NULL,
[MemberDistinguishedName] [nvarchar](512) NULL,
CONSTRAINT [PK_GroupMemberGUIDs] PRIMARY KEY CLUSTERED 
(
	[GroupGUID] ASC, [MemberGUID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

--ALTER TABLE [dbo].[ADgroup_members] WITH CHECK ADD FOREIGN KEY([GroupGUID])
--REFERENCES [dbo].[ADgroups] ([ObjectGUID])
--GO

--CREATE TABLE [dbo].[ADgroup_user_members](
--	[GroupGUID] [uniqueidentifier] NOT NULL,
--	[UserGUID] [uniqueidentifier] NOT NULL,
-- CONSTRAINT [PK_GroupUserMemberGUIDs] PRIMARY KEY CLUSTERED 
--(
--	[GroupGUID] ASC,
--	[UserGUID] ASC
--)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
--) ON [PRIMARY]

--GO

--ALTER TABLE [dbo].[ADgroup_user_members]  WITH CHECK ADD FOREIGN KEY([GroupGUID])
--REFERENCES [dbo].[ADgroups] ([ObjectGUID])
--GO

--ALTER TABLE [dbo].[ADgroup_user_members]  WITH CHECK ADD FOREIGN KEY([UserGUID])
--REFERENCES [dbo].[ADusers] ([ObjectGUID])
--GO

CREATE TABLE [dbo].[ADwell_known_sids](
[Name] [nvarchar](256) NULL,
[Description] [nvarchar](512) NULL,
[DistinguishedName] [nvarchar](512) NULL,
[DisplayName] [nvarchar](128) NULL,
[ObjectCategory] [nvarchar](128) NULL,
[ObjectClass] [nvarchar](64) NULL,
[SID] [nvarchar](128) NULL,
[ObjectGUID] [uniqueidentifier] NOT NULL,
CONSTRAINT [PK_WellKnownSIDsGUID] PRIMARY KEY CLUSTERED 
(
	[ObjectGUID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ADwell_known_sids_lookup](
[SID] [nvarchar](128) NOT NULL,
[DisplayName] [nvarchar](128) NULL,
[Description] [nvarchar](512) NULL,
CONSTRAINT [PK_WellKnownSIDsLookup] PRIMARY KEY CLUSTERED 
(
	[SID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

INSERT [dbo].[ADwell_known_sids_lookup] ([SID], [DisplayName], [Description]) 
VALUES 
(N'S-1-0', N'Null Authority', N'An identifier authority.'), 
(N'S-1-0-0', N'Nobody', N'No security principal.'), 
(N'S-1-1', N'World Authority', N'An identifier authority.'), 
(N'S-1-1-0', N'Everyone', N'A group that includes all users, even anonymous users and guests. Membership is controlled by the operating system.'), 
(N'S-1-2', N'Local Authority', N'An identifier authority.'), 
(N'S-1-2-0', N'Local', N'A group that includes all users who have logged on locally. '), 
(N'S-1-2-1', N'Console Logon', N'A group that includes users who are logged on to the physical console. '), 
(N'S-1-3', N'Creator Authority', N'An identifier authority.'), 
(N'S-1-3-0', N'Creator Owner', N'A placeholder in an inheritable access control entry (ACE). When the ACE is inherited, the system replaces this SID with the SID for the object''s creator.'), 
(N'S-1-3-1', N'Creator Group', N'A placeholder in an inheritable ACE. When the ACE is inherited, the system replaces this SID with the SID for the primary group of the object''s creator. The primary group is used only by the POSIX subsystem.'), 
(N'S-1-3-2', N'Creator Owner Server', N'This SID is not used in Windows 2000.'), 
(N'S-1-3-3', N'Creator Group Server', N'This SID is not used in Windows 2000.'), 
(N'S-1-3-4', N'Owner Rights', N'A group that represents the current owner of the object. When an ACE that carries this SID is applied to an object, the system ignores the implicit READ_CONTROL and WRITE_DAC permissions for the object owner.'), 
--(N'S-1-5-80-0', N'All Services ', N'A group that includes all service processes configured on the system. Membership is controlled by the operating system. '), 
(N'S-1-4', N'Non-unique Authority', N'An identifier authority.'), 
(N'S-1-5', N'NT Authority', N'An identifier authority.'), 
(N'S-1-5-1', N'Dialup', N'A group that includes all users who have logged on through a dial-up connection. Membership is controlled by the operating system.'), 
(N'S-1-5-2', N'Network', N'A group that includes all users that have logged on through a network connection. Membership is controlled by the operating system.'), 
(N'S-1-5-3', N'Batch', N'A group that includes all users that have logged on through a batch queue facility. Membership is controlled by the operating system.'), 
(N'S-1-5-4', N'Interactive', N'A group that includes all users that have logged on interactively. Membership is controlled by the operating system.'), 
(N'S-1-5-5-X-Y', N'Logon Session', N'A logon session. The X and Y values for these SIDs are different for each session.'), 
(N'S-1-5-6', N'Service', N'A group that includes all security principals that have logged on as a service. Membership is controlled by the operating system.'), 
(N'S-1-5-7', N'Anonymous', N'A group that includes all users that have logged on anonymously. Membership is controlled by the operating system.'), 
(N'S-1-5-8', N'Proxy', N'This SID is not used in Windows 2000.'), 
(N'S-1-5-9', N'Enterprise Domain Controllers', N'A group that includes all domain controllers in a forest that uses an Active Directory directory service. Membership is controlled by the operating system.'), 
(N'S-1-5-10', N'Principal Self', N'A placeholder in an inheritable ACE on an account object or group object in Active Directory. When the ACE is inherited, the system replaces this SID with the SID for the security principal who holds the account.'), 
(N'S-1-5-11', N'Authenticated Users', N'A group that includes all users whose identities were authenticated when they logged on. Membership is controlled by the operating system.'), 
(N'S-1-5-12', N'Restricted Code', N'This SID is reserved for future use.'), 
(N'S-1-5-13', N'Terminal Server Users', N'A group that includes all users that have logged on to a Terminal Services server. Membership is controlled by the operating system.'), 
(N'S-1-5-14', N'Remote Interactive Logon', N'A group that includes all users who have logged on through a terminal services logon. '), 
(N'S-1-5-15', N'This Organization', N'A group that includes all users from the same organization. Only included with AD accounts and only added by a Windows Server 2003 or later domain controller. '), 
(N'S-1-5-17', N'This Organization', N'An account that is used by the default Internet Information Services (IIS) user. '), 
(N'S-1-5-18', N'Local System', N'A service account that is used by the operating system.'), 
(N'S-1-5-19', N'NT Authority', N'Local Service'), 
(N'S-1-5-20', N'NT Authority', N'Network Service'), 
(N'S-1-5-21-domain-500', N'Administrator', N'A user account for the system administrator. By default, it is the only user account that is given full control over the system.'), 
(N'S-1-5-21-domain-501', N'Guest', N'A user account for people who do not have individual accounts. This user account does not require a password. By default, the Guest account is disabled.'), 
(N'S-1-5-21-domain-502', N'KRBTGT', N'A service account that is used by the Key Distribution Center (KDC) service.'), 
(N'S-1-5-21-domain-512', N'Domain Admins', N'A global group whose members are authorized to administer the domain. By default, the Domain Admins group is a member of the Administrators group on all computers that have joined a domain, including the domain controllers. Domain Admins is the default owner of any object that is created by any member of the group.'), 
(N'S-1-5-21-domain-513', N'Domain Users', N'A global group that, by default, includes all user accounts in a domain. When you create a user account in a domain, it is added to this group by default.'), 
(N'S-1-5-21-domain-514', N'Domain Guests', N'A global group that, by default, has only one member, the domain''s built-in Guest account.'), 
(N'S-1-5-21-domain-515', N'Domain Computers', N'A global group that includes all clients and servers that have joined the domain.'), 
(N'S-1-5-21-domain-516', N'Domain Controllers', N'A global group that includes all domain controllers in the domain. New domain controllers are added to this group by default.'), 
(N'S-1-5-21-domain-517', N'Cert Publishers', N'A global group that includes all computers that are running an enterprise certification authority. Cert Publishers are authorized to publish certificates for User objects in Active Directory.'), 
(N'S-1-5-21-root domain-518', N'Schema Admins', N'A universal group in a native-mode domain; a global group in a mixed-mode domain. The group is authorized to make schema changes in Active Directory. By default, the only member of the group is the Administrator account for the forest root domain.'), 
(N'S-1-5-21-root domain-519', N'Enterprise Admins', N'A universal group in a native-mode domain; a global group in a mixed-mode domain. The group is authorized to make forest-wide changes in Active Directory, such as adding child domains. By default, the only member of the group is the Administrator account for the forest root domain.'), 
(N'S-1-5-21-domain-520', N'Group Policy Creator Owners', N'A global group that is authorized to create new Group Policy objects in Active Directory. By default, the only member of the group is Administrator.'), 
(N'S-1-5-21-domain-553', N'RAS and IAS Servers', N'A domain local group. By default, this group has no members. Servers in this group have Read Account Restrictions and Read Logon Information access to User objects in the Active Directory domain local group. '), 
(N'S-1-5-32-544', N'Administrators', N'A built-in group. After the initial installation of the operating system, the only member of the group is the Administrator account. When a computer joins a domain, the Domain Admins group is added to the Administrators group. When a server becomes a domain controller, the Enterprise Admins group also is added to the Administrators group.'), 
(N'S-1-5-32-545', N'Users', N'A built-in group. After the initial installation of the operating system, the only member is the Authenticated Users group. When a computer joins a domain, the Domain Users group is added to the Users group on the computer.'), 
(N'S-1-5-32-546', N'Guests', N'A built-in group. By default, the only member is the Guest account. The Guests group allows occasional or one-time users to log on with limited privileges to a computer''s built-in Guest account.'), 
(N'S-1-5-32-547', N'Power Users', N'A built-in group. By default, the group has no members. Power users can create local users and groups; modify and delete accounts that they have created; remove users from the Power Users, Users, and Guests groups. Power users also can install programs; create, manage and delete local printers; and create and delete file shares.'), 
(N'S-1-5-32-548', N'Account Operators', N'A built-in group that exists only on domain controllers. By default, the group has no members. By default, Account Operators have permission to create, modify, and delete accounts for users, groups, and computers in all containers and organizational units of Active Directory except the Builtin container and the Domain Controllers OU. Account Operators do not have permission to modify the Administrators and Domain Admins groups, nor modify the accounts for members of those groups.'), 
(N'S-1-5-32-549', N'Server Operators', N'A built-in group that exists only on domain controllers. By default, the group has no members. Server Operators can log on to a server interactively; create and delete network shares; start and stop services; back up and restore files; format the hard disk of the computer; and shut down the computer.'), 
(N'S-1-5-32-550', N'Print Operators', N'A built-in group that exists only on domain controllers. By default, the only member is the Domain Users group. Print Operators can manage printers and document queues.'), 
(N'S-1-5-32-551', N'Backup Operators', N'A built-in group. By default, the group has no members. Backup Operators can back up and restore all files on a computer, regardless of the permissions that protect those files. Backup Operators also can log on to the computer and shut it down.'), 
(N'S-1-5-32-552', N'Replicators', N'A built-in group that is used by the File Replication service on domain controllers. By default, the group has no members. Do not add users to this group.'), 
(N'S-1-5-64-10', N'NTLM Authentication ', N'A SID that is used when the NTLM authentication package authenticated the client '), 
(N'S-1-5-64-14', N'SChannel Authentication ', N'A SID that is used when the SChannel authentication package authenticated the client. '), 
(N'S-1-5-64-21', N'Digest Authentication ', N'A SID that is used when the Digest authentication package authenticated the client. '), 
(N'S-1-5-80', N'NT Service ', N'An NT Service account prefix'), 
(N'S-1-5-80-0', N'NT SERVICES\ALL SERVICES', N'A group that includes all service processes that are configured on the system. Membership is controlled by the operating system.'), 
(N'S-1-5-83-0', N'NT VIRTUAL MACHINE\Virtual Machines', N'A built-in group. The group is created when the Hyper-V role is installed. Membership in the group is maintained by the Hyper-V Management Service (VMMS). This group requires the "Create Symbolic Links" right (SeCreateSymbolicLinkPrivilege), and also the "Log on as a Service" right (SeServiceLogonRight). '), 
(N'S-1-16-0', N'Untrusted Mandatory Level ', N'An untrusted integrity level. Note Added in Windows Vista and Windows Server 2008 '), 
(N'S-1-16-4096', N'Low Mandatory Level ', N'A low integrity level. '), 
(N'S-1-16-8192', N'Medium Mandatory Level ', N'A medium integrity level. '), 
(N'S-1-16-8448', N'Medium Plus Mandatory Level ', N'A medium plus integrity level. '), 
(N'S-1-16-12288', N'High Mandatory Level ', N'A high integrity level. '), 
(N'S-1-16-16384', N'System Mandatory Level ', N'A system integrity level. '), 
(N'S-1-16-20480', N'Protected Process Mandatory Level ', N'A protected-process integrity level. '), 
(N'S-1-16-28672', N'Secure Process Mandatory Level ', N'A secure process integrity level. '), 
(N'S-1-5-32-554', N'BUILTIN\Pre-Windows 2000 Compatible Access', N'An alias added by Windows 2000. A backward compatibility group which allows read access on all users and groups in the domain. '), 
(N'S-1-5-32-555', N'BUILTIN\Remote Desktop Users', N'An alias. Members in this group are granted the right to logon remotely. '), 
(N'S-1-5-32-556', N'BUILTIN\Network Configuration Operators', N'An alias. Members in this group can have some administrative privileges to manage configuration of networking features. '), 
(N'S-1-5-32-557', N'BUILTIN\Incoming Forest Trust Builders', N'An alias. Members of this group can create incoming, one-way trusts to this forest. '), 
(N'S-1-5-32-558', N'BUILTIN\Performance Monitor Users', N'An alias. Members of this group have remote access to monitor this computer. '), 
(N'S-1-5-32-559', N'BUILTIN\Performance Log Users', N'An alias. Members of this group have remote access to schedule logging of performance counters on this computer. '), 
(N'S-1-5-32-560', N'BUILTIN\Windows Authorization Access Group', N'An alias. Members of this group have access to the computed tokenGroupsGlobalAndUniversal attribute on User objects. '), 
(N'S-1-5-32-561', N'BUILTIN\Terminal Server License Servers', N'An alias. A group for Terminal Server License Servers. When Windows Server 2003 Service Pack 1 is installed, a new local group is created.'), 
(N'S-1-5-32-562', N'BUILTIN\Distributed COM Users', N'An alias. A group for COM to provide computerwide access controls that govern access to all call, activation, or launch requests on the computer.'), 
(N'S-1-5-21-domain-498', N'Enterprise Read-only Domain Controllers ', N'A Universal group. Members of this group are Read-Only Domain Controllers in the enterprise '), 
(N'S-1-5-21-domain-521', N'Read-only Domain Controllers', N'A Global group. Members of this group are Read-Only Domain Controllers in the domain'), 
(N'S-1-5-32-569', N'BUILTIN\Cryptographic Operators', N'A Builtin Local group. Members are authorized to perform cryptographic operations.'), 
(N'S-1-5-21-domain-571', N'Allowed RODC Password Replication Group ', N'A Domain Local group. Members in this group can have their passwords replicated to all read-only domain controllers in the domain. '), 
(N'S-1-5-21-domain-572', N'Denied RODC Password Replication Group ', N'A Domain Local group. Members in this group cannot have their passwords replicated to any read-only domain controllers in the domain '), 
(N'S-1-5-32-573', N'BUILTIN\Event Log Readers ', N'A Builtin Local group. Members of this group can read event logs from local machine. '), 
(N'S-1-5-32-574', N'BUILTIN\Certificate Service DCOM Access ', N'A Builtin Local group. Members of this group are allowed to connect to Certification Authorities in the enterprise. '), 
(N'S-1-5-21-domain-522', N'Cloneable Domain Controllers', N'A Global group. Members of this group that are domain controllers may be cloned.'), 
(N'S-1-5-32-575', N'BUILTIN\RDS Remote Access Servers', N'A Builtin Local group. Servers in this group enable users of RemoteApp programs and personal virtual desktops access to these resources. In Internet-facing deployments, these servers are typically deployed in an edge network. This group needs to be populated on servers running RD Connection Broker. RD Gateway servers and RD Web Access servers used in the deployment need to be in this group.'), 
(N'S-1-5-32-576', N'BUILTIN\RDS Endpoint Servers', N'A Builtin Local group. Servers in this group run virtual machines and host sessions where users RemoteApp programs and personal virtual desktops run. This group needs to be populated on servers running RD Connection Broker. RD Session Host servers and RD Virtualization Host servers used in the deployment need to be in this group.'), 
(N'S-1-5-32-577', N'BUILTIN\RDS Management Servers', N'A Builtin Local group. Servers in this group can perform routine administrative actions on servers running Remote Desktop Services. This group needs to be populated on all servers in a Remote Desktop Services deployment. The servers running the RDS Central Management service must be included in this group.'), 
(N'S-1-5-32-578', N'BUILTIN\Hyper-V Administrators', N'A Builtin Local group. Members of this group have complete and unrestricted access to all features of Hyper-V.'), 
(N'S-1-5-32-579', N'BUILTIN\Access Control Assistance Operators', N'A Builtin Local group. Members of this group can remotely query authorization attributes and permissions for resources on this computer.'), 
(N'S-1-5-32-580', N'BUILTIN\Remote Management Users', N'A Builtin Local group. Members of this group can access WMI resources over management protocols (such as WS-Management via the Windows Remote Management service). This applies only to WMI namespaces that grant access to the user.');
GO

CREATE TABLE [dbo].[ADusersPhotos](
[ObjectGUID] [uniqueidentifier] NOT NULL,
[Width] [int] NULL,
[Height] [int] NULL,
[Photo] [varbinary](max) NULL,
CONSTRAINT [PK_ADusersPhotosGUIDs] PRIMARY KEY CLUSTERED 
(
	[ObjectGUID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

