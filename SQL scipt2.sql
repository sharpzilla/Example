DECLARE @DrawFileName NVARCHAR(255);
DECLARE @pattern INT = LEN(N'ИГТ.123456.789.01.123');
DECLARE @suffixes NVARCHAR(15) = N'СБ%.slddrw';
DECLARE @currentDate DATETIME = GETDATE(); 
DECLARE @filenametemplate NVARCHAR(255) = N'ИГТ.[0-9][0-9][0-9][0-9][0-9][0-9].[0-9][0-9][0-9].[0-9][0-9].[0-9][0-9][0-9]И1%.docx';
DECLARE @variableID INT = (SELECT VariableID FROM dbo.Variable WHERE VariableName = 'п_Разраб'); 
DECLARE @folderID INT = (SELECT MAX(ProjectTreeRec.Level) FROM dbo.ProjectTreeRec WHERE ChildProjectID = (SELECT projectID FROM dbo.Projects WHERE Name LIKE N'Проекты' AND deleted = 0)) + 1;
DECLARE @temptable TABLE (
DocumentID INT NOT NULL,
ProjectID INT NOT NULL,
Filename NVARCHAR(255),
ProjName NVARCHAR(255),
FirstFileDate DATETIME,
Designer NVARCHAR(100),
DrawFileName NVARCHAR(255),
DrawFileDate DATETIME);

INSERT INTO @temptable (DocumentID, ProjectID, Filename, ProjName, FirstFileDate, Designer, DrawFileName)
SELECT DISTINCT d.DocumentID, dip.ProjectID FolderID, d.Filename FileName, p.Name, ur.FirstFileDate,
dbo.acFindAnyVariableValue2(@variableID, d.DocumentID),
CONCAT(SUBSTRING(d.Filename, 1, @pattern), @suffixes) DrawFileName

FROM dbo.Documents AS d
LEFT OUTER JOIN
(SELECT DocumentID, MAX(Date) AS FirstFileDate FROM dbo.UserRevs GROUP BY DocumentID) ur
ON d.DocumentID = ur.DocumentID
INNER JOIN dbo.DocumentsInProjects dip
ON d.DocumentID = dip.DocumentID
INNER JOIN dbo.ProjectTreeRec ptr
ON dip.ProjectID = ptr.ChildProjectID
INNER JOIN dbo.Projects p
ON p.ProjectID = ptr.ParentProjectID
INNER JOIN dbo.Revisions r
ON r.DocumentID = d.DocumentID
INNER JOIN dbo.Users u
ON u.UserID = r.UserID

WHERE d.Deleted = 0
AND d.Filename NOT LIKE N'%^%'
AND d.Filename LIKE @filenametemplate
AND ptr.Level = @folderID
AND r.RevNr = 1

DECLARE cur
CURSOR FOR
SELECT DrawFileName
FROM @temptable
OPEN cur
FETCH NEXT
FROM cur
INTO @DrawFileName WHILE @@FETCH_STATUS = 0
BEGIN
UPDATE @temptable
SET DrawFileDate = (
SELECT MAX(FirstFileDate1)
FROM dbo.Documents d
LEFT OUTER JOIN (SELECT DocumentID, MAX(Date) FirstFileDate1 FROM dbo.UserRevs GROUP BY DocumentID) ur
ON d.DocumentID = ur.DocumentID
WHERE d.Deleted = 0
AND d.Filename LIKE @DrawFileName)
WHERE DrawFileName = @DrawFileName
FETCH NEXT
FROM cur
INTO @DrawFileName
END CLOSE cur
DEALLOCATE cur

SELECT DocumentID, ProjectID, Filename, ProjName, Designer, DATEDIFF(DAY, DrawFileDate, @currentDate) 'Delay'
FROM @temptable
WHERE DrawFileDate IS NOT NULL
AND (DATEDIFF(DAY, DrawFileDate, FirstFileDate) > 0 OR DATEDIFF(DAY, DrawFileDate, FirstFileDate) IS NULL)
ORDER BY Delay