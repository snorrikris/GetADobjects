-- Note - xp_cmdshell must be enabled for this script to work.
-- See: https://msdn.microsoft.com/en-us/library/ms190693.aspx

-- Set Domain path.
DECLARE @ADpath nvarchar(64) = 'LDAP://DC=veca,DC=is';


CREATE TABLE ##ADusersPhotos (
	[ObjectGUID] [uniqueidentifier] NOT NULL,
	[Width] [int] NULL,
	[Height] [int] NULL,
	[Format] [nvarchar](6),
	[Photo] [varbinary](max) NULL
);

PRINT 'INSERT user photos into temp table.';
--DECLARE @ADfilter nvarchar(128) = '(&(objectCategory=person)(objectClass=user)(SamAccountName=snorri))';
DECLARE @ADfilter nvarchar(4000) = '(&(objectCategory=person)(objectClass=user))';
INSERT INTO ##ADusersPhotos EXEC clr_GetADusersPhotos @ADpath, @ADfilter;
--SELECT u.DisplayName, p.Height, p.Width, p.[Format]
--FROM #ADusersPhotos p
--JOIN #ADusers u ON p.ObjectGUID = u.ObjectGUID;

-- Source: http://www.sqlusa.com/bestpractices/imageimportexport/
------------
-- T-SQL Export all images in table to file system folder
-- Source table: Production.ProductPhoto  - Destination: C:\Temp\
------------
 

DECLARE	@FilePath varchar(128) = 'C:\Photos\', 
		 @Command       VARCHAR(4000), 
         @Format		nvarchar(6),
		 @ObjectGUID	uniqueidentifier,
         @ImageFileName VARCHAR(128) 
 
DECLARE curPhotoImage CURSOR FOR             -- Cursor for each image in table
SELECT [Format],
	   [ObjectGUID]
FROM   ##ADusersPhotos 
WHERE  [Photo] IS NOT NULL
 
OPEN curPhotoImage 
 
FETCH NEXT FROM curPhotoImage 
INTO @Format, 
     @ObjectGUID 

 
WHILE (@@FETCH_STATUS = 0) -- Cursor loop  
  BEGIN 
	SET @ImageFileName = @FilePath + CAST(@ObjectGUID AS VARCHAR(128)) + '.' + @Format;

-- Keep the bcp command on ONE LINE - SINGLE LINE!!!  
    --SET @Command = 'bcp "SELECT [Photo] FROM ##ADusersPhotos WHERE [ObjectGUID] = ''' + CAST(@ObjectGUID AS VARCHAR(128)) + '''" queryout "' + @ImageFileName + '" -T -n -S SNORRIDEV\SQLEXPRESS';
    SET @Command = 'bcp "SELECT [Photo] FROM ##ADusersPhotos WHERE [ObjectGUID] = ''' + CAST(@ObjectGUID AS VARCHAR(128)) + '''" queryout "' + @ImageFileName + '" -T -N -S SNORRIDEV\SQLEXPRESS';
     
    PRINT @Command -- debugging  
/* bcp "SELECT LargePhoto FROM AdventureWorks2008.Production.ProductPhoto 
WHERE ProductPhotoID = 69" queryout 
"K:\data\images\productphoto\racer02_black_f_large.gif" -T -n -SHPESTAR
*/
     
    EXEC xp_cmdshell @Command     -- Carry out image export to file from db table
     
    FETCH NEXT FROM curPhotoImage 
    INTO @Format, 
         @ObjectGUID 
  END  -- cursor loop 
 
CLOSE curPhotoImage 
 
DEALLOCATE curPhotoImage 

DROP TABLE ##ADusersPhotos;