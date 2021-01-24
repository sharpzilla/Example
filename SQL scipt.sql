DECLARE @duration INT = 1; --Задается кол-во дней после которого считается задержка файла.
DECLARE @variableID INT = (SELECT VariableID FROM dbo.Variable WHERE VariableName = 'п_Разраб');--ID=47
DECLARE @currentDate DATETIME = GETDATE();
DECLARE @folderID INT = (SELECT MAX(ProjectTreeRec.Level) FROM dbo.ProjectTreeRec WHERE ChildProjectID = (SELECT projectID FROM dbo.Projects WHERE Name LIKE N'Проекты' AND deleted = 0)) + 1

SELECT DISTINCT d.Filename AS 'FileName',
CONVERT(CHAR(12), h.Date, 104) 'Date',
CASE
WHEN (CAST(DATEDIFF(DAY, h.Date, @currentDate) - @duration AS INT) = 0 OR CAST(DATEDIFF(DAY, h.Date, @currentDate) - @duration AS INT) = -1) THEN 1
ELSE CAST(DATEDIFF(DAY, h.Date, @currentDate) AS INT)
END 'Delay',
d.DocumentID 'DocID',
dbo.acFindAnyVariableValue2(@variableID, d.DocumentID) 'Designer',
dip.ProjectID 'FolderID',
ptr.ParentProjectID 'ProjID',
p.Name 'ProjName'

FROM dbo.Documents d
INNER JOIN dbo.TransitionHIstory h
ON h.DocumentID = d.DocumentID
INNER JOIN dbo.DocumentsInProjects dip
ON d.DocumentID = dip.DocumentID
INNER JOIN dbo.ProjectTreeRec ptr
ON dip.ProjectID = ptr.ChildProjectID
INNER JOIN dbo.Projects p
ON p.ProjectID = ptr.ParentProjectID

WHERE d.CurrentStatusID IN (SELECT s.StatusID FROM dbo.Status s WHERE s.Name LIKE N'%Корректировка%' AND s.Enabled = 1)
AND d.DocumentID = h.DocumentID
AND h.TransitionNr = (SELECT MAX(TransitionNr) FROM dbo.TransitionHistory WHERE h.DocumentID = DocumentID)
AND d.Filename NOT LIKE N'%^ИГТ.%'
AND d.Filename NOT LIKE N'Элемент%'
AND d.Filename NOT LIKE N'Sheet%'
AND d.Filename NOT LIKE N'%^Копия%'
AND DATEDIFF(DAY, h.Date, @currentDate) > @duration
AND d.Deleted = 0
AND ptr.Level = @folderID
ORDER BY 'Delay' ASC, 'Designer' ASC
