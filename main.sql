use test_task; --Enter your bd name
CREATE TABLE PracticeGroup
(
	ID int PRIMARY KEY IDENTITY NOT NULL,
	GroupName varchar(50)
);
go
CREATE TABLE Partner
(
	ID int PRIMARY KEY IDENTITY NOT NULL,
	Practice_Group_ID Int,
	FirstName varchar(50),
	LastName varchar(50),
	FOREIGN KEY (Practice_Group_ID) REFERENCES PracticeGroup(ID)
);
go
CREATE TABLE PartnerTarget
(
	ID int IDENTITY(1,1) NOT NULL,
	partner_id int NOT NULL,
	use_Date datetime NOT NULL,
	Utilization real NULL,
	GrossProfit real NULL,
	Realization real NULL,
	NetBillRate real NULL,
	Revenue real NULL,
	ProductionHours real NULL,
	[AR90+Days] real NULL,
	ProductionAmount real NULL,
	FOREIGN KEY (partner_id) REFERENCES Partner(ID)
);
go
CREATE VIEW ViewTargets as
SELECT dbo.PartnerTarget.ID, dbo.PracticeGroup.ID AS Expr1, dbo.Partner.ID AS Expr2, dbo.PracticeGroup.GroupName, dbo.Partner.FirstName + ' ' + dbo.Partner.LastName AS FullName, dbo.PartnerTarget.use_Date, dbo.PartnerTarget.GrossProfit, dbo.PartnerTarget.Realization, dbo.PartnerTarget.NetBillRate, dbo.PartnerTarget.Revenue, dbo.PartnerTarget.ProductionHours, dbo.PartnerTarget.[AR90+Days],  dbo.PartnerTarget.ProductionAmount
FROM dbo.Partner INNER JOIN
dbo.PartnerTarget ON dbo.Partner.ID = dbo.PartnerTarget.partner_id INNER JOIN
dbo.PracticeGroup ON dbo.Partner.Practice_Group_ID = dbo.PracticeGroup.ID
go
CREATE Procedure ImportTargets
@way varchar(1000),
@year varchar(10)
AS
Begin
	Declare @way_to_file varchar(1000), @setyear varchar (50);
	set @way_to_file = @way;
	set @setyear = @year;
	if ( select use_Date from PartnerTarget where YEAR(use_Date) = @setyear) IS NULL
		Print 'No rows in the table for this year';
	else
		delete PartnerTarget where YEAR(use_Date) = @setyear
	--Adding unique Group Names
	create table groupname
	(
		id_gp int IDENTITY (1,1),
		gp varchar(20)
	);
	DBCC CHECKIDENT (PartnerTarget, RESEED, 1)
	DBCC CHECKIDENT(PracticeGroup, RESEED, 1)
	DBCC CHECKIDENT ([Partner], RESEED, 1)
	declare @way_pg varchar(500)
	set @way_to_file = 'Excel 12.0; Database=' + @way_to_file; -- main path
		--
	set @way_pg = 'Insert Into groupname(gp) 
	Select [Practice Group] From OPENROWSET(''Microsoft.ACE.OLEDB.12.0'',
	''' +  @way_to_file + ''', ''SELECT * FROM [Invoice$B2:B57]'')';
	exec (@way_pg)
	declare @gp_counter int
	set @gp_counter = 1
	while @gp_counter<=55
		if (select GroupName from PracticeGroup where GroupName = (select gp from groupname where id_gp=@gp_counter)) is NULL
			begin
				insert into PracticeGroup(GroupName) select gp from groupname where id_gp = @gp_counter
				set @gp_counter = @gp_counter + 1
			end;
		else
			begin
				set @gp_counter = @gp_counter + 1
			end;
	drop table groupname;
	--Adding Names
	declare @way_partner varchar(1000)
	create table Names
	(
	FullName varchar(50),
	LastName varchar(50),
	FirstName varchar(50)
	)
	create table gp_id
	(
	GroupNam varchar(20),
	id_gp int,
	id int IDENTITY(1,1)
	);
	set @way_partner = 
	'Insert Into Names
	Select Partner,
		Substring(Partner,1, CharIndex('','', Partner) - 1) As FirstName,
		Substring(Partner, CharIndex('','', Partner) + 1, len(Partner)) AS LastName
	From OPENROWSET(''Microsoft.ACE.OLEDB.12.0'', 
		''' + @way_to_file + ''', ''SELECT * FROM [Invoice$A2:A57]'')';
	exec (@way_partner)
	declare @way_partner_gp varchar(300)
	set @way_partner_gp = 'Insert Into gp_id (GroupNam)
	Select * From OPENROWSET(''Microsoft.ACE.OLEDB.12.0'', 
		''' + @way_to_file + ''', ''SELECT * FROM [Invoice$B2:B57]'')';
	exec (@way_partner_gp)
	Insert into Partner(FirstName,LastName)
	select FirstName,LastName
	from Names;
	declare @counter_groupname int , @gp varchar(20), @id_g int
	set @counter_groupname = 1
	while @counter_groupname <=55
		Begin
			set @gp = (select GroupNam from gp_id where id=@counter_groupname);
			if (Select GroupName from PracticeGroup where GroupName = @gp) IS NOT NULL
			begin
				set @id_g = (Select ID from PracticeGroup where GroupName = @gp)
				Update [Partner] set Practice_Group_ID = @id_g where ID = @counter_groupname;
			end;
			set @counter_groupname = @counter_groupname + 1;
		end;

	--drop table Names;
	--Adding date and all columns of table
	--Unpivot the table.  
	Create TABLE test_month (id int IDENTITY (1,1), Months varchar(10), u_date date);
	declare @way_month varchar(1000)
	set @way_month = 
	'Insert Into test_month( Months)
	SELECT Months  
	FROM   
	   (SELECT F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12 
	   FROM OPENROWSET(''Microsoft.ACE.OLEDB.12.0'', 
	''' + @way_to_file + ''', ''SELECT * FROM [Invoice$C1:N2]'')) p  
	UNPIVOT  
	   (Months FOR F IN   
		  (F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12 )  
	)AS unpvt';  
	exec (@way_month)
	
	--While and insert other columns in table

	create table auxiliary
	(
		id int IDENTITY(1,1),
		use_date date,
		grosspro real,
		realization real,
		netbillrate real,
		revenue real,
		prodho real,
		ar90 real,
		prodamo real
	);
	create table revenue(id_rev int IDENTITY(1,1),revenue_import real);
	create table grossprofit(id_gross int IDENTITY(1,1),Grossprofit_import real);
	create table realization(id_real int IDENTITY(1,1),realization_import real);
	create table Netbillrate(id_net int IDENTITY(1,1),netbillrate_import real);
	create table productionHours(id_prod_hours int IDENTITY(1,1),productionhours_import real);
	create table AR90(id_ar int IDENTITY(1,1),ar90_import real);
	create table ProductionAmount(id_prodam int IDENTITY(1,1),prodam_import real);
	declare @counter int,  @count int;
	set @count = 67;
	set @counter = 1;
	WHILE @counter <=12
		BEGIN
			--Update and insert date in another form
			declare @newDate varchar(50), @list varchar(100), @temporary varchar(500)
			set @newDate = (Select Months from test_month where id = @counter ) + '1,' + @setyear
			update test_month set u_date=@newDate where id = @counter;
			--Create auxiliary table
			-----1
				set @list = 'SELECT * FROM [Invoice$' + NCHAR(@count) + '2:' + NCHAR(@count) + '57]'
				set @temporary = 'Insert Into revenue(revenue_import)
				Select * From OPENROWSET(''Microsoft.ACE.OLEDB.12.0'', 
				''' + @way_to_file + ''', ''' + @list +''')';
				exec (@temporary);
				set @temporary = NULL
			-----2
				set @list = 'SELECT * FROM [GP$' + NCHAR(@count) + '2:' + NCHAR(@count) + '57]'
				set @temporary = 'Insert Into grossprofit(Grossprofit_import)
				Select * From OPENROWSET(''Microsoft.ACE.OLEDB.12.0'', 
				''' + @way_to_file + ''', ''' + @list +''')';
				exec (@temporary);
				set @temporary = NULL
			-----3
			set @list = 'SELECT * FROM [WAR Realization$' + NCHAR(@count) + '2:' + NCHAR(@count) + '57]'
				set @temporary = 'Insert Into realization(realization_import)
				Select * From OPENROWSET(''Microsoft.ACE.OLEDB.12.0'', 
				''' + @way_to_file + ''', ''' + @list +''')';
				exec (@temporary);
				set @temporary = NULL
			---4
			set @list = 'SELECT * FROM [Net Bill Rates$' + NCHAR(@count) + '2:' + NCHAR(@count) + '57]'
				set @temporary = 'Insert Into Netbillrate(netbillrate_import)
				Select * From OPENROWSET(''Microsoft.ACE.OLEDB.12.0'', 
				''' + @way_to_file + ''', ''' + @list +''')';
				exec (@temporary);
				set @temporary = NULL
			----5
				set @list = 'SELECT * FROM [Production Hours$' + NCHAR(@count) + '2:' + NCHAR(@count) + '57]'
				set @temporary = 'Insert Into productionHours(productionhours_import)
				Select * From OPENROWSET(''Microsoft.ACE.OLEDB.12.0'', 
				''' + @way_to_file + ''', ''' + @list +''')';
				exec (@temporary);
				set @temporary = NULL
			----6
				set @list = 'SELECT * FROM [AR Aging$' + NCHAR(@count) + '2:' + NCHAR(@count) + '57]'
				set @temporary = 'Insert Into AR90(ar90_import)
				Select * From OPENROWSET(''Microsoft.ACE.OLEDB.12.0'', 
				''' + @way_to_file + ''', ''' + @list +''')';
				exec (@temporary);
				set @temporary = NULL
			-----7
				set @list = 'SELECT * FROM [Prod Amts$' + NCHAR(@count) + '2:' + NCHAR(@count) + '57]'
				set @temporary = 'Insert Into ProductionAmount(prodam_import)
				Select * From OPENROWSET(''Microsoft.ACE.OLEDB.12.0'', 
				''' + @way_to_file + ''', ''' + @list +''')';
				exec (@temporary);
				set @temporary = NULL
			--Insert from lists
			Insert into auxiliary(revenue) select revenue_import from revenue
			declare @Partner_counter int;
			set @Partner_counter = 1;
			WHILE @Partner_counter <=55
			BEGIN		
				update auxiliary set use_date = @newDate;
				update auxiliary set grosspro = Grossprofit_import  from grossprofit where id = id_gross;
				update auxiliary set realization = realization_import  from realization where id = id_real;
				update auxiliary set netbillrate = netbillrate_import  from Netbillrate where id = id_net;
				update auxiliary set prodho = productionhours_import  from productionHours where id = id_prod_hours;
				update auxiliary set ar90 = ar90_import  from AR90 where id = id_ar;
				update auxiliary set prodamo = prodam_import  from ProductionAmount where id = id_prodam;
				set @Partner_counter = @Partner_counter + 1	
			end;
			-- insert all cells into row
			insert into PartnerTarget(partner_id, use_Date,GrossProfit,Realization,NetBillRate,Revenue,ProductionHours,[AR90+Days],ProductionAmount) 
			select id, use_date, grosspro, realization, netbillrate, revenue, prodho, ar90, prodamo  from auxiliary;
			delete auxiliary;
			DBCC CHECKIDENT (auxiliary, RESEED, 0)
			delete revenue;
			DBCC CHECKIDENT (revenue, RESEED, 0)
			delete grossprofit;
			DBCC CHECKIDENT (grossprofit, RESEED, 0)
			delete realization;
			DBCC CHECKIDENT (realization, RESEED, 0)
			delete Netbillrate;
			DBCC CHECKIDENT (Netbillrate, RESEED, 0)
			delete productionHours;
			DBCC CHECKIDENT (productionHours, RESEED, 0)
			delete AR90;
			DBCC CHECKIDENT (AR90, RESEED, 0)
			delete ProductionAmount;
			DBCC CHECKIDENT (ProductionAmount, RESEED, 0)
			set @counter = @counter + 1
			set @count = @count + 1
		end;
	
	drop table gp_id,Names,auxiliary,test_month,revenue,grossprofit,realization,Netbillrate,productionHours,AR90,ProductionAmount;
end;

