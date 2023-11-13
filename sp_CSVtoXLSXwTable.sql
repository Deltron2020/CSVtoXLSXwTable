USE Database
GO

if object_id('dbo.sp_CSVtoXLSXwTable', 'P') is not null
	drop procedure dbo.sp_CSVtoXLSXwTable
GO
/* ===============
Created By Tyler T on 3/15/23

This stored procedure converts a csv file to xlsx,
then it uses the passed in rowCount & colCharacter (From Excel)
to convert the data in the xlsx file into a Table.
This makes viewing the data clearer and allows dynamic linking.

Added this portion on 3/28/23
--SELECT @path, dbo.doesFileExist (@path) AS [IsExist] 
 =============== */


CREATE PROCEDURE [dbo].[sp_CSVtoXLSXwTable]( @fullCsvPath VARCHAR(512) , @fullXlsxPath VARCHAR(512) , @rowCount INT , @colCharacter VARCHAR(4) )
AS
BEGIN
SET NOCOUNT ON

DECLARE @sqlCommand NVARCHAR(1000);

/* ================  CSV to Excel Conversion =============== */
DECLARE @isExists BIT = (SELECT dbo.doesFileExist (@fullCsvPath)) ;

IF @isExists = 1
BEGIN;

	SET @sqlCommand = 'powershell.exe -file "C:\PS Scripts\CSVtoXL_Conversion.ps1" "' + @fullCsvPath + '" "' + @fullXlsxPath + '"';
	PRINT @sqlCommand
	EXEC master..xp_cmdshell @sqlCommand

END;

ELSE
BEGIN;

	PRINT 'The file @ '+@fullCsvPath+' does not exist.'

END;


/* ================  Excel Table Creation =============== */
SET @isExists = (SELECT dbo.doesFileExist (@fullXlsxPath)) ;

IF @isExists = 1
BEGIN;

	SET @sqlCommand = 'powershell.exe -file "C:\PS Scripts\CreateXLTable.ps1" "' + @fullXlsxPath + '" "' + @fullXlsxPath + '" "' + CAST(@rowcount AS VARCHAR) + '" "' + @colCharacter + '"';
	PRINT @sqlCommand
	EXEC master..xp_cmdshell @sqlCommand

END;
ELSE
BEGIN;

	PRINT 'The file @ '+@fullXlsxPath+' does not exist.'

END;


END