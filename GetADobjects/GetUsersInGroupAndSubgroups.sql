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
SELECT DISTINCT [MemberGUID] AS [UserGUID]
FROM cte
WHERE MemberType = 'user'
--ORDER BY GroupGUID
--ORDER BY [MemberDistinguishedName]
