DECLARE @ADpath nvarchar(64) = 'LDAP://DC=veca,DC=is';
DECLARE @ADfilter nvarchar(128) = '(&(objectCategory=person)(objectClass=user))';
INSERT INTO dbo.[ADusersPhotos] EXEC clr_GetADusersPhotos @ADpath, @ADfilter;
