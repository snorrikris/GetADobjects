-- Get table columns as CSV
DECLARE @CSV nvarchar(4000);
SET @CSV = (SELECT LEFT(ColNames, LEN(ColNames) - 1)
FROM (
select '[' + c.name + '],' from sys.columns c 
join sys.tables t on c.object_id = t.object_id
where t.name = 'ADgroups'
order by c.column_id
FOR XML PATH('')
) c (ColNames));
SELECT @CSV;

-- Get table columns as CSV formatted for MERGE INSERT clause
DECLARE @CSV2 nvarchar(4000);
SET @CSV2 = (SELECT LEFT(ColNames, LEN(ColNames) - 1)
FROM (
select 'T.[' + c.name + '] = S.[' + c.name + '],' from sys.columns c 
join sys.tables t on c.object_id = t.object_id
where t.name = 'ADgroups'
order by c.column_id
FOR XML PATH('')
) c (ColNames));
SELECT @CSV2;

-- Get table columns as CSV formatted for MERGE UPDATE clause - VALUES
DECLARE @CSV3 nvarchar(4000);
SET @CSV3 = (SELECT LEFT(ColNames, LEN(ColNames) - 1)
FROM (
select 'S.[' + c.name + '],' from sys.columns c 
join sys.tables t on c.object_id = t.object_id
where t.name = 'ADgroups'
order by c.column_id
FOR XML PATH('')
) c (ColNames));
SELECT @CSV3;

