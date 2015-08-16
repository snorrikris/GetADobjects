DECLARE @ADpath nvarchar(64) = 'LDAP://DC=veca,DC=is';
--DECLARE @ADfilter nvarchar(64) = '(objectCategory=builtinDomain)';
--DECLARE @ADfilter nvarchar(512) = '(|(objectSID=S-1-5-20)(objectSID=S-1-5-11)(objectSID=S-1-5-6)(objectSID=S-1-5-18)(objectSID=S-1-5-17)(objectSID=S-1-5-4)(objectSID=S-1-5-9))';

--DECLARE @ADfilter nvarchar(1024) = '(|(objectSID=S-1-0)(objectSID=S-1-0-0)(objectSID=S-1-1)(objectSID=S-1-1-0)(objectSID=S-1-2)(objectSID=S-1-2-0)(objectSID=S-1-2-1)(objectSID=S-1-3)(objectSID=S-1-3-0)(objectSID=S-1-3-1)(objectSID=S-1-3-2)(objectSID=S-1-3-3)(objectSID=S-1-3-4)(objectSID=S-1-5-80-0)(objectSID=S-1-4))';
--DECLARE @ADfilter nvarchar(1024) = '(|(objectSID=S-1-5)(objectSID=S-1-5-1)(objectSID=S-1-5-2)(objectSID=S-1-5-3)(objectSID=S-1-5-4)(objectSID=S-1-5-6)(objectSID=S-1-5-7)(objectSID=S-1-5-8)(objectSID=S-1-5-9)(objectSID=S-1-5-10)(objectSID=S-1-5-11)(objectSID=S-1-5-12)(objectSID=S-1-5-13)(objectSID=S-1-5-14)(objectSID=S-1-5-15)(objectSID=S-1-5-17)(objectSID=S-1-5-18)(objectSID=S-1-5-19)(objectSID=S-1-5-20))';
DECLARE @ADfilter nvarchar(1024) = '(|(objectSID=S-1-5-32-544)(objectSID=S-1-5-32-545)(objectSID=S-1-5-32-546)(objectSID=S-1-5-32-547)(objectSID=S-1-5-32-548)(objectSID=S-1-5-32-549)(objectSID=S-1-5-32-550)(objectSID=S-1-5-32-551)(objectSID=S-1-5-32-552))';
--DECLARE @ADfilter nvarchar(1024) = '(|(objectSID=S-1-5-64-10)(objectSID=S-1-5-64-14)(objectSID=S-1-5-64-21)(objectSID=S-1-5-80)(objectSID=S-1-5-80-0)(objectSID=S-1-5-83-0)(objectSID=S-1-16-0)(objectSID=S-1-16-4096)(objectSID=S-1-16-8192)(objectSID=S-1-16-8448)(objectSID=S-1-16-12288)(objectSID=S-1-16-16384)(objectSID=S-1-16-20480)(objectSID=S-1-16-28672))';
--DECLARE @ADfilter nvarchar(1024) = '(|(objectSID=S-1-5-32-554)(objectSID=S-1-5-32-555)(objectSID=S-1-5-32-556)(objectSID=S-1-5-32-557)(objectSID=S-1-5-32-558)(objectSID=S-1-5-32-559)(objectSID=S-1-5-32-560)(objectSID=S-1-5-32-561)(objectSID=S-1-5-32-562)(objectSID=S-1-5-32-569)(objectSID=S-1-5-32-573)(objectSID=S-1-5-32-574)(objectSID=S-1-5-32-575)(objectSID=S-1-5-32-576)(objectSID=S-1-5-32-577)(objectSID=S-1-5-32-578)(objectSID=S-1-5-32-579)(objectSID=S-1-5-32-580))';
EXEC clr_GetADpropertiesToFile @ADpath, @ADfilter;

