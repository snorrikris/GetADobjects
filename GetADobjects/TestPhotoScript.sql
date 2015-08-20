DECLARE @ADpath nvarchar(64) = 'LDAP://DC=veca,DC=is';
--DECLARE @ADfilter nvarchar(128) = '(&(objectCategory=person)(objectClass=user)(SamAccountName=snorri))';
DECLARE @ADfilter nvarchar(128) = '(&(objectCategory=person)(objectClass=user))';
EXEC clr_GetADusersPhotos @ADpath, @ADfilter;
