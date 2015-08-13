--EXEC clr_GetADusers

--EXEC clr_GetADcontacts

--EXEC clr_GetADcomputers

--EXEC clr_GetADgroups

--DECLARE @Members XML;
--EXEC clr_GetADgroupsEx @Members OUTPUT;
--WITH MemberList AS (
--select Tg.Cg.value('@GrpDS', 'nvarchar(256)') as GroupDS,
--		Tm.Cm.value('@MemberDS', 'nvarchar(256)') as MemberDS
--from @Members.nodes('/body') as Tb(Cb)
--  outer apply Tb.Cb.nodes('Group') as Tg(Cg)
--  outer apply Tg.Cg.nodes('Member') AS Tm(Cm)
--)
--SELECT M.*, 
--	G.objectGUID AS GroupGUID,
--	COALESCE(U.objectGUID, GM.objectGUID, C.objectGUID) AS MemberGUID,
--	COALESCE(U.ObjectClass, GM.ObjectClass, C.ObjectClass) AS MemberType
--FROM MemberList M
--LEFT JOIN dbo.Groups G ON M.GroupDS = G.DistinguishedName
--LEFT JOIN dbo.Groups GM ON M.MemberDS = GM.DistinguishedName
--LEFT JOIN dbo.Computers C ON M.MemberDS = C.DistinguishedName
--LEFT JOIN dbo.Users U ON M.MemberDS = U.DistinguishedName;

--INSERT INTO xmltmp VALUES ( @Members );
--SELECT @Members

EXEC clr_GetADcomputersEx

--EXEC clr_GetADcontactsEx

--EXEC clr_GetADusersEx

