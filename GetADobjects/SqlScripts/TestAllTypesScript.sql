-- Set Domain path.
DECLARE @ADpath nvarchar(64) = 'LDAP://DC=veca,DC=is';

PRINT 'Get all users from AD'
DECLARE @Members XML;
DECLARE @ADfilter nvarchar(4000) = '(&(objectCategory=person)(objectClass=user))';
EXEC clr_GetADobjects @ADpath, @ADfilter, @Members OUTPUT;

PRINT 'Get all Contacts from AD'
SET @ADfilter = '(&(objectCategory=person)(objectClass=contact))';
EXEC clr_GetADobjects @ADpath, @ADfilter, @Members OUTPUT;

PRINT 'Get all computers from AD'
SET @ADfilter = '(objectCategory=computer)';
EXEC clr_GetADobjects @ADpath, @ADfilter, @Members OUTPUT;

PRINT 'Get all Well known SIDs from AD'
SET @ADfilter = '(|(objectSID=S-1-0)(objectSID=S-1-0-0)(objectSID=S-1-1)(objectSID=S-1-1-0)(objectSID=S-1-2)(objectSID=S-1-2-0)(objectSID=S-1-2-1)(objectSID=S-1-3)(objectSID=S-1-3-0)(objectSID=S-1-3-1)(objectSID=S-1-3-2)(objectSID=S-1-3-3)(objectSID=S-1-3-4)(objectSID=S-1-5-80-0)(objectSID=S-1-4)(objectSID=S-1-5)(objectSID=S-1-5-1)(objectSID=S-1-5-2)(objectSID=S-1-5-3)(objectSID=S-1-5-4)(objectSID=S-1-5-6)(objectSID=S-1-5-7)(objectSID=S-1-5-8)(objectSID=S-1-5-9)(objectSID=S-1-5-10)(objectSID=S-1-5-11)(objectSID=S-1-5-12)(objectSID=S-1-5-13)(objectSID=S-1-5-14)(objectSID=S-1-5-15)(objectSID=S-1-5-17)(objectSID=S-1-5-18)(objectSID=S-1-5-19)(objectSID=S-1-5-20)(objectSID=S-1-5-64-10)(objectSID=S-1-5-64-14)(objectSID=S-1-5-64-21)(objectSID=S-1-5-80)(objectSID=S-1-5-80-0)(objectSID=S-1-5-83-0)(objectSID=S-1-16-0)(objectSID=S-1-16-4096)(objectSID=S-1-16-8192)(objectSID=S-1-16-8448)(objectSID=S-1-16-12288)(objectSID=S-1-16-16384)(objectSID=S-1-16-20480)(objectSID=S-1-16-28672))';
EXEC clr_GetADobjects @ADpath, @ADfilter, @Members OUTPUT;

PRINT 'Get all groups from AD - Note group members returned in @Members as XML.'
SET @ADfilter = '(objectCategory=group)';
EXEC clr_GetADobjects @ADpath, @ADfilter, @Members OUTPUT;

PRINT 'All group members.';
WITH MemberList AS ( -- Process group members XML data.
select Tg.Cg.value('@GrpDS', 'nvarchar(256)') as GroupDS,
		Tm.Cm.value('@MemberDS', 'nvarchar(256)') as MemberDS
from @Members.nodes('/body') as Tb(Cb)
  outer apply Tb.Cb.nodes('Group') as Tg(Cg)
  outer apply Tg.Cg.nodes('Member') AS Tm(Cm)
)
SELECT * FROM MemberList;

PRINT 'Get user photos.';
SET @ADfilter = '(&(objectCategory=person)(objectClass=user))';
EXEC clr_GetADusersPhotos @ADpath, @ADfilter;
