--GroupGUID = '95411D9F-B530-4071-B6B4-01BD4CEFDD7C'
-- Source: http://stackoverflow.com/questions/14599152/recursive-cte-to-get-members-of-a-group
-- Source: http://data.stackexchange.com/stackoverflow/query/94066/recursive-cte-to-get-members-of-a-group

WITH cte AS
(
SELECT [GroupGUID]
      ,[MemberGUID]
      ,[MemberType]
      ,[GroupDistinguishedName]
      ,[MemberDistinguishedName]
	  , 0 AS Level
FROM [dbo].[ADgroup_members]
WHERE GroupGUID = 'D60E6669-FB5D-4502-AF1D-084184E8EB1E' --'95411D9F-B530-4071-B6B4-01BD4CEFDD7C'
--WHERE MemberType = 'user'
UNION ALL
SELECT gm.[GroupGUID]
      ,gm.[MemberGUID]
      ,gm.[MemberType]
      ,gm.[GroupDistinguishedName]
      ,gm.[MemberDistinguishedName]
	  ,c.Level + 1
	  FROM cte c
	  INNER JOIN [dbo].[ADgroup_members] gm ON c.MemberGUID = gm.GroupGUID
)
SELECT [GroupGUID]
      ,[MemberGUID]
      ,[MemberType]
      ,[GroupDistinguishedName]
      ,[MemberDistinguishedName]
	  ,Level
FROM cte
--WHERE MemberType = 'user'
--ORDER BY GroupGUID
--ORDER BY [MemberDistinguishedName]
