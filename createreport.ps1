Param(
    [Parameter(Mandatory=$true)][string]$instance,
    [Parameter(Mandatory=$false)][string]$targetOrMaster
 )
 if(!$targetOrMaster) {
     $targetOrMaster = 'master'
     Write-Host $targetOrMaster
 }

### dit script schaamteloos gejat van MSSQLTIPS.COM :
### https://www.mssqltips.com/sqlservertip/5025/stored-procedure-to-generate-html-tables-for-sql-server-query-output/ 
$proc_CreateScript = "
CREATE OR ALTER PROC [dbo].[usp_Query2html] (@SQLQuery NVARCHAR(3000))
AS
BEGIN
   DECLARE @columnslist NVARCHAR (1000) = ''
   DECLARE @restOfQuery NVARCHAR (2000) = ''
   DECLARE @DynTSQL NVARCHAR (3000)
   DECLARE @FROMPOS INT

   SET NOCOUNT ON

   SELECT @columnslist += 'ISNULL (' + NAME + ',' + '''' + ' ' + '''' + ')' + ','
   FROM sys.dm_exec_describe_first_result_set(@SQLQuery, NULL, 0)

   SET @columnslist = left (@columnslist, Len (@columnslist) - 1)
   SET @FROMPOS = CHARINDEX ('FROM', @SQLQuery, 1)
   SET @restOfQuery = SUBSTRING(@SQLQuery, @FROMPOS, LEN(@SQLQuery) - @FROMPOS + 1)
   SET @columnslist = Replace (@columnslist, '),', ') as TD,')
   SET @columnslist += ' as TD'
   SET @DynTSQL = CONCAT (
         'SELECT (SELECT '
         , @columnslist
         ,' '
         , @restOfQuery
         ,' FOR XML RAW (''TR''), ELEMENTS, TYPE) AS ''TBODY'''
         ,' FOR XML PATH (''''), ROOT (''TABLE'')'
         )

   EXEC (@DynTSQL)
   SET NOCOUNT OFF
END
GO
";

$hulpfunctie = "
create or alter function XofNiet(@nudag nvarchar(10), @id int, @freq_type int )
returns nvarchar(10) 
WITH EXECUTE AS CALLER
as
begin
	declare @result nvarchar(10)

    ----/ @freq8intervals heb ik schaamteloos gejat van sqlches https://sqlcodesnippets.com/tag/freq_interval/ zodoende /----

	declare @freq8intervals table (id int, intverval_dag varchar( 70 ))
	insert into @freq8intervals select 1, 'zondag'
	insert into @freq8intervals select 2, 'maandag'
	insert into @freq8intervals select 3, 'zondag, maandag'
	insert into @freq8intervals select 4, 'dinsdag'
	insert into @freq8intervals select 5, 'zondag, dinsdag'
	insert into @freq8intervals select 6, 'maandag, dinsdag'
	insert into @freq8intervals select 7, 'zondag, maandag, dinsdag'
	insert into @freq8intervals select 8, 'woensdag'
	insert into @freq8intervals select 9, 'zondag, woensdag'
	insert into @freq8intervals select 10, 'maandag, woensdag'
	insert into @freq8intervals select 11, 'zondag, maandag, woensdag'
	insert into @freq8intervals select 12, 'dinsdag, woensdag'
	insert into @freq8intervals select 13, 'zondag, dinsdag, woensdag'
	insert into @freq8intervals select 14, 'maandag, dinsdag, woensdag'
	insert into @freq8intervals select 15, 'zondag, maandag, dinsdag, woensdag'
	insert into @freq8intervals select 16, 'donderdag'
	insert into @freq8intervals select 17, 'zondag, donderdag'
	insert into @freq8intervals select 18, 'maandag, donderdag'
	insert into @freq8intervals select 19, 'zondag, maandag, donderdag'
	insert into @freq8intervals select 20, 'dinsdag, donderdag'
	insert into @freq8intervals select 21, 'zondag, dinsdag, donderdag'
	insert into @freq8intervals select 22, 'maandag, dinsdag, donderdag'
	insert into @freq8intervals select 23, 'zondag, maandag, dinsdag, donderdag'
	insert into @freq8intervals select 24, 'woensdag, donderdag'
	insert into @freq8intervals select 25, 'zondag, woensdag, donderdag'
	insert into @freq8intervals select 26, 'maandag, woensdag, donderdag'
	insert into @freq8intervals select 27, 'zondag, maandag, woensdag, donderdag'
	insert into @freq8intervals select 28, 'dinsdag, woensdag, donderdag'
	insert into @freq8intervals select 29, 'zondag, dinsdag, woensdag, donderdag'
	insert into @freq8intervals select 30, 'maandag, dinsdag, woensdag, donderdag'
	insert into @freq8intervals select 31, 'zondag, maandag, dinsdag, woensdag, donderdag'
	insert into @freq8intervals select 32, 'vrijdag'
	insert into @freq8intervals select 33, 'zondag, vrijdag'
	insert into @freq8intervals select 34, 'maandag, vrijdag'
	insert into @freq8intervals select 35, 'zondag, maandag, vrijdag'
	insert into @freq8intervals select 36, 'dinsdag, vrijdag'
	insert into @freq8intervals select 37, 'zondag, dinsdag, vrijdag'
	insert into @freq8intervals select 38, 'maandag, dinsdag, vrijdag'
	insert into @freq8intervals select 39, 'zondag, maandag, dinsdag, vrijdag'
	insert into @freq8intervals select 40, 'woensdag, vrijdag'
	insert into @freq8intervals select 41, 'zondag, woensdag, vrijdag'
	insert into @freq8intervals select 42, 'maandag, woensdag, vrijdag'
	insert into @freq8intervals select 43, 'zondag, maandag, woensdag, vrijdag'
	insert into @freq8intervals select 44, 'dinsdag, woensdag, vrijdag'
	insert into @freq8intervals select 45, 'zondag, dinsdag, woensdag, vrijdag'
	insert into @freq8intervals select 46, 'maandag, dinsdag, woensdag, vrijdag'
	insert into @freq8intervals select 47, 'zondag, maandag, dinsdag, woensdag, vrijdag'
	insert into @freq8intervals select 48, 'donderdag, vrijdag'
	insert into @freq8intervals select 49, 'zondag, donderdag, vrijdag'
	insert into @freq8intervals select 50, 'maandag, donderdag, vrijdag'
	insert into @freq8intervals select 51, 'zondag, maandag, donderdag, vrijdag'
	insert into @freq8intervals select 52, 'dinsdag, donderdag, vrijdag'
	insert into @freq8intervals select 53, 'zondag, dinsdag, donderdag, vrijdag'
	insert into @freq8intervals select 54, 'maandag, dinsdag, donderdag, vrijdag'
	insert into @freq8intervals select 55, 'zondag, maandag, dinsdag, donderdag, vrijdag'
	insert into @freq8intervals select 56, 'woensdag, donderdag, vrijdag'
	insert into @freq8intervals select 57, 'zondag, woensdag, donderdag, vrijdag'
	insert into @freq8intervals select 58, 'maandag, woensdag, donderdag, vrijdag'
	insert into @freq8intervals select 59, 'zondag, maandag, woensdag, donderdag, vrijdag'
	insert into @freq8intervals select 60, 'dinsdag, woensdag, donderdag, vrijdag'
	insert into @freq8intervals select 61, 'zondag, dinsdag, woensdag, donderdag, vrijdag'
	insert into @freq8intervals select 62, 'maandag, dinsdag, woensdag, donderdag, vrijdag'
	insert into @freq8intervals select 63, 'zondag, maandag, dinsdag, woensdag, donderdag, vrijdag'
	insert into @freq8intervals select 64, 'zaterdag'
	insert into @freq8intervals select 65, 'zondag, zaterdag'
	insert into @freq8intervals select 66, 'maandag, zaterdag'
	insert into @freq8intervals select 67, 'zondag, maandag, zaterdag'
	insert into @freq8intervals select 68, 'dinsdag, zaterdag'
	insert into @freq8intervals select 69, 'zondag, dinsdag, zaterdag'
	insert into @freq8intervals select 70, 'maandag, dinsdag, zaterdag'
	insert into @freq8intervals select 71, 'zondag, maandag, dinsdag, zaterdag'
	insert into @freq8intervals select 72, 'woensdag, zaterdag'
	insert into @freq8intervals select 73, 'zondag, woensdag, zaterdag'
	insert into @freq8intervals select 74, 'maandag, woensdag, zaterdag'
	insert into @freq8intervals select 75, 'zondag, maandag, woensdag, zaterdag'
	insert into @freq8intervals select 76, 'dinsdag, woensdag, zaterdag'
	insert into @freq8intervals select 77, 'zondag, dinsdag, woensdag, zaterdag'
	insert into @freq8intervals select 78, 'maandag, dinsdag, woensdag, zaterdag'
	insert into @freq8intervals select 79, 'zondag, maandag, dinsdag, woensdag, zaterdag'
	insert into @freq8intervals select 80, 'donderdag, zaterdag'
	insert into @freq8intervals select 81, 'zondag, donderdag, zaterdag'
	insert into @freq8intervals select 82, 'maandag, donderdag, zaterdag'
	insert into @freq8intervals select 83, 'zondag, maandag, donderdag, zaterdag'
	insert into @freq8intervals select 84, 'dinsdag, donderdag, zaterdag'
	insert into @freq8intervals select 85, 'zondag, dinsdag, donderdag, zaterdag'
	insert into @freq8intervals select 86, 'maandag, dinsdag, donderdag, zaterdag'
	insert into @freq8intervals select 87, 'zondag, maandag, dinsdag, donderdag, zaterdag'
	insert into @freq8intervals select 88, 'woensdag, donderdag, zaterdag'
	insert into @freq8intervals select 89, 'zondag, woensdag, donderdag, zaterdag'
	insert into @freq8intervals select 90, 'maandag, woensdag, donderdag, zaterdag'
	insert into @freq8intervals select 91, 'zondag, maandag, woensdag, donderdag, zaterdag'
	insert into @freq8intervals select 92, 'dinsdag, woensdag, donderdag, zaterdag'
	insert into @freq8intervals select 93, 'zondag, dinsdag, woensdag, donderdag, zaterdag'
	insert into @freq8intervals select 94, 'maandag, dinsdag, woensdag, donderdag, zaterdag'
	insert into @freq8intervals select 95, 'zondag, maandag, dinsdag, woensdag, donderdag, zaterdag'
	insert into @freq8intervals select 96, 'vrijdag, zaterdag'
	insert into @freq8intervals select 97, 'zondag, vrijdag, zaterdag'
	insert into @freq8intervals select 98, 'maandag, vrijdag, zaterdag'
	insert into @freq8intervals select 99, 'zondag, maandag, vrijdag, zaterdag'
	insert into @freq8intervals select 100, 'dinsdag, vrijdag, zaterdag'
	insert into @freq8intervals select 101, 'zondag, dinsdag, vrijdag, zaterdag'
	insert into @freq8intervals select 102, 'maandag, dinsdag, vrijdag, zaterdag'
	insert into @freq8intervals select 103, 'zondag, maandag, dinsdag, vrijdag, zaterdag'
	insert into @freq8intervals select 104, 'woensdag, vrijdag, zaterdag'
	insert into @freq8intervals select 105, 'zondag, woensdag, vrijdag, zaterdag'
	insert into @freq8intervals select 106, 'maandag, woensdag, vrijdag, zaterdag'
	insert into @freq8intervals select 107, 'zondag, maandag, woensdag, vrijdag, zaterdag'
	insert into @freq8intervals select 108, 'dinsdag, woensdag, vrijdag, zaterdag'
	insert into @freq8intervals select 109, 'zondag, dinsdag, woensdag, vrijdag, zaterdag'
	insert into @freq8intervals select 110, 'maandag, dinsdag, woensdag, vrijdag, zaterdag'
	insert into @freq8intervals select 111, 'zondag, maandag, dinsdag, woensdag, vrijdag, zaterdag'
	insert into @freq8intervals select 112, 'donderdag, vrijdag, zaterdag'
	insert into @freq8intervals select 113, 'zondag, donderdag, vrijdag, zaterdag'
	insert into @freq8intervals select 114, 'maandag, donderdag, vrijdag, zaterdag'
	insert into @freq8intervals select 115, 'zondag, maandag, donderdag, vrijdag, zaterdag'
	insert into @freq8intervals select 116, 'dinsdag, donderdag, vrijdag, zaterdag'
	insert into @freq8intervals select 117, 'zondag, dinsdag, donderdag, vrijdag, zaterdag'
	insert into @freq8intervals select 118, 'maandag, dinsdag, donderdag, vrijdag, zaterdag'
	insert into @freq8intervals select 119, 'zondag, maandag, dinsdag, donderdag, vrijdag, zaterdag'
	insert into @freq8intervals select 120, 'woensdag, donderdag, vrijdag, zaterdag'
	insert into @freq8intervals select 121, 'zondag, woensdag, donderdag, vrijdag, zaterdag'
	insert into @freq8intervals select 122, 'maandag, woensdag, donderdag, vrijdag, zaterdag'
	insert into @freq8intervals select 123, 'zondag, maandag, woensdag, donderdag, vrijdag, zaterdag'
	insert into @freq8intervals select 124, 'dinsdag, woensdag, donderdag, vrijdag, zaterdag'
	insert into @freq8intervals select 125, 'zondag, dinsdag, woensdag, donderdag, vrijdag, zaterdag'
	insert into @freq8intervals select 126, 'maandag, dinsdag, woensdag, donderdag, vrijdag, zaterdag'
	insert into @freq8intervals select 127, 'zondag, maandag, dinsdag, woensdag, donderdag, vrijdag, zaterdag'

	declare @freq8dagstring nvarchar(70) = (select intverval_dag from @freq8intervals where id = @id)

	declare @freq32intervals table (id int, intervaldag varchar(10))
	insert into @freq32intervals values (1, 'zondag')
	insert into @freq32intervals values (2, 'maandag')
	insert into @freq32intervals values (3, 'dinsdag')
	insert into @freq32intervals values (4, 'woensdag')
	insert into @freq32intervals values (5, 'donderdag')
	insert into @freq32intervals values (6, 'vrijdag')
	insert into @freq32intervals values (7, 'zaterdag')

	declare @freq32dagstring varchar(10) = (select intervaldag from @freq32intervals where id = @id)

	set @result = 
		case 
			when @freq_type = 4 then 'X'
			when (@freq_type = 8) and CHARINDEX(@nudag, @freq8dagstring) > 0 then 'X' 
			when @freq_type = 32 and charindex(@nudag, @freq32dagstring) > 0 then 'X'
			else '' 
		end;

	return @result
end;"

$viewQuery = "
create or alter view [vw_SqlJob] as
select '' as jobnaam,
	'zondag' as zondag,
	'maandag' as maandag ,
	'dinsdag' as dinsdag,
	'woensdag' as woensdag,
	'donderdag' as donderdag,
	'vrijdag' as vrijdag,
	'zaterdag' as zaterdag 
union all
-- -----dagdelen is nog todo :)
	
--select '',
--	'ochtend','middag','nacht','ochtend','middag','nacht','ochtend','middag','nacht','ochtend','middag','nacht','ochtend','middag','nacht','ochtend','middag','nacht','ochtend','middag','nacht'
--union all
select 
j.name
,dbo.XofNiet('zondag',s.freq_interval ,s.freq_type)
,dbo.XofNiet('maandag',s.freq_interval ,s.freq_type)
,dbo.XofNiet('dinsdag',s.freq_interval ,s.freq_type)
,dbo.XofNiet('woensdag',s.freq_interval ,s.freq_type)
,dbo.XofNiet('donderdag',s.freq_interval,s.freq_type)
,dbo.XofNiet('vrijdag',s.freq_interval,s.freq_type)
,dbo.XofNiet('zaterdag',s.freq_interval,s.freq_type)
from msdb.dbo.sysjobs j
inner join msdb.dbo.sysjobschedules js on js.job_id = j.job_id
inner join msdb.dbo.sysschedules s on s.schedule_id = js.schedule_id;
"

Invoke-DbaQuery -SqlInstance $instance -Database $targetOrMaster -CommandType Text -Query $proc_CreateScript -EnableException
Invoke-DbaQuery -SqlInstance $instance -Database $targetOrMaster -CommandType Text -Query $hulpfunctie -EnableException
Invoke-DbaQuery -SqlInstance $instance -Database $targetOrMaster -CommandType Text -Query $viewQuery -EnableException
$bcpQuery = "exec $targetOrMaster.dbo.usp_Query2html 'select * from vw_SqlJob'"

bcp "$bcpQuery" queryout report.html -c -T -S $instance
#Copy /Y style.txt + report.html s_report.html
exit 1
