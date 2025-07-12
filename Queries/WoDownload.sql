INSERT INTO [@P4S_EXPORT] (Code, [Name], U_DateCreated, U_DateModified, U_DocId, U_ObjType,U_ExportFlag,U_Lock,U_Backorder)
SELECT concat(ObjType, DocEntry), concat(ObjType, DocEntry), getdate(), getdate(), DocEntry, ObjType,0,0,0, *
FROM owor
where Status = 'R'