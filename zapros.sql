use test_task;
declare @way_to_file varchar(100), @year varchar(10)
set @way_to_file = 'C:\Nikita test\Targets - Sample.xlsx'
set @year = '2017'
EXEC ImportTargets @way_to_file,@year
