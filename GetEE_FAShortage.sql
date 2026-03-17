----------------!! Can only Run after OPS_OPO_new generated !!
---------------Need to run twice to get qi eMsg
/*
drop table #tmp
drop table #MD
drop table #MatMD
drop table #result
drop table #sub
drop table #BL
drop table #PL
drop table #rep
drop table #dd
drop table #dp
drop table #dd_mat
drop table #dd_rep
drop table #dd_result
delete from OPS_BL where ReportDate=convert(char(10),getdate(),111) and Customer='HP'
drop table #BLD
*/
update BLMaterial$ set Remark=substring(Remark,charindex(' )',Remark)+3,1000) where Remark like '%OPO -->%'

select *,AbleToTransfer='N' into #tmp from
(
select MType,Material,description,Buyer,demand_part,ShortageQty,ServiceStock,NBStock=0,ManufactureStock,ManufactureExcessStock,TrialStock,SHStock,Intransit  from 
(
select MType,Material,description,buyer,demand_part,ShortageQty,ServiceStock,ManufactureStock,ManufactureExcessStock,TrialStock,SHStock,Intransit from (
select MType='NB',Material=left(part_no,12),description,buyer,demand_part,ManufactureStock=stock,ServiceStock=fstock,ManufactureExcessStock=jstock,TrialStock=pstock,SHStock=bstock,Intransit=instransit,ShortageQty=total_qty from icc_r_shortage_data 
where len(rtrim(part_no))>12 and (where_use like '11%' or where_use like '13%')
union
select MType='DT',part_no=left(part_no,12),description,buyer,demand_part,stock,fstock,jstock,pstock,bstock,instransit,total_qty from icc_dt_r_shortage_data 
where len(rtrim(part_no))>12 and (where_use like '11%' or where_use like '13%')
) as a
) as a left join
(select * from BuyerCode) as b on a.buyer=b.BuyerCode
) as a 


update #tmp set NBStock=b.Qty from #tmp a,Ivan_CurrentINV b where a.Material=b.MatNo and b.Customer='HP' and INVDate=convert(char(10),getdate(),111)

update #tmp set AbleToTransfer='Y' where ManufactureExcessStock>=ShortageQty

----(2017/01/05) Only MType 'DT' need this logic
update #tmp set AbleToTransfer='N' where NBStock>=ShortageQty and MType='DT'

create table #MD(iid int identity(1,1),MType varchar(20),Material varchar(20),demand_part varchar(8000))

create table #MatMD(MType varchar(20),Material varchar(20),MD varchar(20))

insert #MD
    select distinct MType,Material,demand_part from #tmp
    

declare @i int
declare @j int
select @i=min(iid) from #MD
select @j=max(iid) from #MD

while @i<=@j
begin
     while (select charindex(';',demand_part) from #MD where iid=@i)>0
     begin
         insert #MatMD
            select MType,Material,left(demand_part,charindex(';',demand_part)-1) from #MD where iid=@i
         update #MD set demand_part=substring(demand_part,charindex(';',demand_part)+1,1000) where iid=@i
     end
     select @i=@i+1 
end
 
select distinct a.MType,a.Material,description,PIC='                    ',Buyer,
MinPO=isnull(MinPO,''),MinNeed=isnull(MinNeed,''),MaxPO=isnull(MaxPO,''),MaxNeed=isnull(MaxNeed,''),ShortageQty,ServiceStock,NBStock,
ManufactureStock,ManufactureExcessStock,TrialStock,SHStock,Intransit,AbleToTransfer,VmiStock=0,AbleFromVmi='N',New='Y',a.demand_part into #result from
(select * from #tmp where not Material like '6071%') as a left join
 (
select distinct a.MType,Material,MinPO=min(MinPO),MinNeed=min(MinNeed),MaxPO=max(MaxPO),MaxNeed=max(MaxNeed) from
(select distinct * from #MatMD) as a left join
(
select IECPN,MinPO=min(POReceiveDate),MinNeed=min(NeedShipDate),MaxPO=max(POReceiveDate),MaxNeed=max(NeedShipDate) from (
select * from OPS_OPO_new where Customer='HP' and ReportDate=convert(char(10),getdate(),111)
and IECPN in (select distinct MD from #MatMD) and Shortage='Y') as a group by IECPN
) as b on a.MD=b.IECPN group by a.MType,Material
) as b on a.MType=b.MType and a.Material=b.Material
order by AbleToTransfer desc,description



update #result set PIC=b.PIC from #result a,
(select Material,b.PIC from t_download_matmas_CP60DW a,MType b where a.Material_group=b.Material_group and Material in (select distinct Material from #result)) as b
where a.Material=b.Material 


update #result set New='' from #result a,BLMaterial$ b where a.MType=b.MType and a.Material=b.Material

-----(2017/11/29) Add Vmi Stock
update #result set VmiStock=b.VmiStock from #result a,
(select part_no,VmiStock=sum(available_qty) from t_download_vmi_852 where plant='CP60' group by part_no) as b 
where a.Material=b.part_no

update #result set AbleFromVmi='Y' where VmiStock>=ShortageQty


--drop table #sub
select distinct a.PF,a.Location,a.Component,INV=0 into #sub from SPSLocation a,
(select distinct PF,Location from SPSLocation where Component in (select distinct Material from #result)) as b
where a.PF=b.PF and a.Location=b.Location


delete from #sub from #sub a,(
select * from (
select PF,Location,qty=count(*) from #sub group by PF,Location
) as a where qty=1
) as b where a.PF=b.PF and a.Location=b.Location

---Check data
--select PF,Location,qty=count(*) from #sub group by PF,Location
--select *  from #sub where PF='PFBH02HMB132' and Location='U4701' 
---Update Inventory
update #sub set INV=b.Qty from #sub a,Ivan_CurrentINV b where a.Component=b.MatNo and b.Customer='HP' and INVDate=convert(char(10),getdate(),111)




--drop table #PL
--drop table #rep

create table #PL(pid int identity(1,1),PF varchar(20),Location varchar(20),Component varchar(500))
create table #rep(iid int identity(1,1),Component varchar(20))

insert #PL
select distinct PF,Location,'' from #sub

/*
declare @i int
declare @j int
*/
declare @x int
declare @y int
declare @PF varchar(20)
declare @Location varchar(20)
declare @Component varchar(500)

select @PF=''
select @Location=''
select @i=min(pid) from #PL
--select @j=10 from #PL
select @j=max(pid) from #PL
select @x=0,@y=0

while @i<=@j
begin
    select @PF=PF,@Location=Location from #PL where pid=@i
    insert #rep
        select distinct Component+'*'+convert(varchar(20),INV) from #sub where PF=@PF and Location=@Location order by Component+'*'+convert(varchar(20),INV)
    --select * from #rep
    select @x=min(iid) from #rep
    select @y=max(iid) from #rep
    select @Component=''
    while @x<=@y
    begin
        select @Component=rtrim(@Component)+rtrim(Component)+'~' from #rep where iid=@x    
        select @x=@x+1
    end
    update #PL set Component=@Component where pid=@i    
    select @Component=''
    select @x=0,@y=0    
    select @i=@i+1
    truncate table #rep
end

--select *  from #PL where PF='PFBH02HMB132' and Location='U4701' 

--select * from (
select distinct a.*,[Material ETA]=isnull(b.[Material ETA],'_'),Remark=isnull(b.Remark,'_') into #BL from #result as a left join (select distinct * from BLMaterial$) as b on a.MType=b.MType and a.Material=b.Material order by AbleToTransfer desc,description
--) as a where Material='6011B0114901'


--drop table #dp
--drop table #dd
create table #dp(iid int identity(1,1),Material varchar(20),demand_part varchar(8000))
create table #dd(Material varchar(20),demand_part varchar(8000))

insert #dp
select distinct Material,demand_part from #BL a,(select distinct PF,Component from #PL) as b 
where (charindex(a.Material,b.Component)>0 and charindex(b.PF,a.demand_part)>0)


--select * from #dp

/*
declare @i int
declare @j int
*/
declare @demand_part varchar(5000)
declare @Material varchar(20)
select @demand_part=''
select @Material=''
select @i=min(iid) from #dp
select @j=max(iid) from #dp
while @i<=@j
begin
    select @Material=Material,@demand_part=demand_part from #dp where iid=@i
    while (charindex(';',@demand_part)>0 and len(rtrim(@demand_part))>1)
    begin
        insert #dd
           select @Material,left(@demand_part,charindex(';',@demand_part)-1)
           select @demand_part=substring(@demand_part,charindex(';',@demand_part)+1,5000)
    end   
    select @Material='',@demand_part=''        
    select @i=@i+1
end


update #dd set demand_part=b.NewPF from #dd a,
(select distinct Material,PF,NewPF=PF+' ('+rtrim(Component)+')' from
(
select a.Material,a.demand_part,b.PF,b.Component from
(select * from #result) as a left join
(select distinct PF,Component from #PL) as b on 
(charindex(a.Material,b.Component)>0 and charindex(b.PF,a.demand_part)>0)
) as a where not PF is null
) as b where a.Material=b.Material and a.demand_part=b.PF

--select * from #dd

---------------------
-----------------------
/*
drop table #dd_mat
drop table #dd_rep
drop table #dd_result
*/
create table #dd_mat(pid int identity(1,1),Material varchar(20))
create table #dd_rep(iid int identity(1,1),Material varchar(20),demand_part varchar(1000))
create table #dd_result(Material varchar(20),demand_part varchar(5000))

insert #dd_mat
    select distinct Material from #dd

/*
declare @i int
declare @j int
declare @x int
declare @y int
declare @demand_part varchar(5000)
*/
declare @mat varchar(20)
select @mat=''
select @i=min(pid) from #dd_mat
--select @j=10 from #PL
select @j=max(pid) from #dd_mat
select @x=0,@y=0

while @i<=@j
begin
    select @mat=Material from #dd_mat where pid=@i
    insert #dd_rep
        select * from #dd where Material=@mat order by demand_part
    --select * from #rep
    select @x=min(iid) from #dd_rep
    select @y=max(iid) from #dd_rep
    select @demand_part=''
    while @x<=@y
    begin
        select @demand_part=rtrim(@demand_part)+rtrim(demand_part)+';' from #dd_rep where iid=@x    
        select @x=@x+1
    end
    insert #dd_result values(@mat,@demand_part)  
    select  @demand_part=''
    select @mat=''
    select @x=0,@y=0    
    select @i=@i+1
    truncate table #dd_rep
end

--select * from #dd_result

-----------------------
------------------------
update #BL set demand_part=b.demand_part from #BL a,#dd_result b where a.Material= b.Material

----------------------------------------------
----------------------------------------------
---(2019/02/14) Remove SW27
---(2017/06/06) Add MW27 & MW12 .
---(2017/03/10) Per LG Request ,add Raw OPO 
---(2017/07/11) Combine all to NB .
update #BL set Remark='( '+b.St+' ) '+rtrim(Remark) from #BL a,
(select Material,St='OPO --> '+convert(varchar(20),sum(Open_Qty)) from OPS_RawOPO 
where ReportDate=convert(char(10),getdate(),111) and Customer='HP' and WH in ('SW04','MW12'/*,'SW27'*/,'MW27') group by Material) b 
where a.Material=b.Material and MType='NB'

/*
update #BL set Remark='( '+b.St+' ) '+rtrim(Remark) from #BL a,
(select Material,St='OPO --> '+convert(varchar(20),sum(Open_Qty)) from OPS_RawOPO 
where ReportDate=convert(char(10),getdate(),111) and Customer='HP' and WH in ('SW27','MW27') group by Material) b 
where a.Material=b.Material and MType='DT'
*/
----------------------------------------------
----------------------------------------------
delete from OPS_BL where Customer='HP' and (ReportDate=convert(char(10),getdate(),111) or ReportDate<=convert(char(10),dateadd(dd,-90,getdate()),111))

insert OPS_BL 
   select convert(char(10),getdate(),111),'HP',* from #BL 

--drop table #BLD

select a.MD,ENSDate='0000/00/00',OpenQty=0,a.Material,MType='---------------------',b.description,b.[Material ETA],b.Remark into #BLD from
(select distinct MD,Material from #MatMD) as a inner join
(select Material,description,[Material ETA],Remark from OPS_BL where Customer='HP' and ReportDate=convert(char(10),getdate(),111)) as b
on a.Material=b.Material order by MD,Material

update #BLD set MType='CPU' where Material like '6025%' and MType='---------------------'
update #BLD set MType='Jack/Socket' where Material like '6026%' and MType='---------------------'
update #BLD set MType='Battery' where Material like '6027%' and MType='---------------------'
update #BLD set MType='DRAM' where Material like '6019%'  and [description] like '%RAM%' and MType='---------------------'
update #BLD set MType='Graphic' where Material like '6019%'  and [description] like '%GRAPHIC%' and MType='---------------------'
update #BLD set MType='IC' where Material like '6019%'  and MType='---------------------'
update #BLD set MType='PCB' where Material like '6050%'  and MType='---------------------'
update #BLD set MType='Connector' where Material like '6012%'  and MType='---------------------'
update #BLD set MType='Passive' where MType='---------------------'

update #BLD set ENSDate=b.NeedShipDate,OpenQty=b.OpenQty from #BLD a,
(select IECPN,NeedShipDate=min(NeedShipDate),OpenQty=sum(OpenQty) from OPS_OPO_new where Customer='HP' and ReportDate=convert(char(10),getdate(),111)
group by IECPN ) as b where a.MD=b.IECPN

select MD,ENSDate,OpenQty,MType,PQty=count(*) from #BLD group by MD,ENSDate,OpenQty,MType order by OpenQty desc,MType

select * from #BLD order by OpenQty desc,MType,Material

select ReportDate=convert(char(10),getdate(),111),MType,Qty_ETA=sum(Qty_ETA),Qty_woETA=sum(Qty_woETA) from
(
select MType,Qty_ETA=count(*),Qty_woETA=0 from 
(select distinct Material,MType,[Material ETA]  from #BLD) as a where not  [Material ETA]='_'
group by MType
union
select MType,0,count(*) from 
(select distinct Material,MType,[Material ETA]  from #BLD) as a where  [Material ETA]='_'
group by MType
) as a group by MType order by sum(Qty_ETA)+sum(Qty_woETA) desc


-----(2024/10/25) Add OMS Stock
-----(2024/10/25) Add SW01 005 Stock
-----(2024/10/23) Add SW01 015 Stock


	select distinct a.*,OMSStock=isnull(b.Qty,0) from
	(

		select a.*,SW01005Stock=isnull(b.Qty,0) from
		(

			select a.*,SW01015Stock=isnull(b.Qty,0) from
			(

			select a.*,Priority=isnull(b.Priority,''),eMsg=isnull(b.eMsg,'') from 
			(select * from OPS_BL where Customer='HP' and ReportDate=convert(char(10),getdate(),111) ) as a left join 
			(select * from Ivan_qi where ReportDate=convert(char(10),getdate(),111) and Customer='BL') as b on  a.Material=b.ShortagePN --order by PIC,Priority,eMsg

			) as a left join

			 (select MatNo,Qty from WIP_WHALL_TD where SLoc='WA1' and SType='015' )  as b on a.Material=b.MatNo --order by PIC,Priority,eMsg

	 		) as a left join

		 (select MatNo,Qty from WIP_WHALL_TD where SLoc='WA1' and SType='005' )  as b on a.Material=b.MatNo --order by PIC,Priority,eMsg

	) as a left join
	(
	select Material,Qty from
	(select * from OMS$ where StockDate=(select max(StockDate) from OMS$)) as a, t_download_matmas_CP60DW as b where a.HPPN=Old_Material
	)  as b on a.Material=b.Material order by PIC,Priority,eMsg


/*
-------------------------------------------------------
-------------------------------------------------------
-------Get PF ETA...
-------------------------------------------------------
-------------------------------------------------------
drop table #tmp_DP
drop table #DP
drop table #tmp_BLETA
drop table #BLETA
drop table #BLResult
create table #tmp_DP(iid int identity(1,1),Material varchar(20),demand_part varchar(8000))
create table #DP(Material varchar(20),demand_part varchar(8000))
insert #tmp_DP
    select distinct Material,demand_part from #tmp where demand_part like '%;%'

---Add single item to #SH_ETA
insert #DP
    select distinct Material,demand_part from #tmp where not demand_part like '%;%' 

declare @a int
declare @b int
select @a=min(iid) from #tmp_DP
select @b=max(iid) from #tmp_DP
while @a<=@b
begin
     while (select charindex(';',demand_part) from #tmp_DP where iid=@a)>0
     begin
         insert #DP
            select distinct Material,left(demand_part,charindex(';',demand_part)-1) from #tmp_DP where iid=@a            
         update #tmp_DP set demand_part=substring(demand_part,charindex(';',demand_part)+1,5000) where iid=@a
         
         insert #DP
            select distinct Material,rtrim(demand_part) from #tmp_DP where iid=@a and not demand_part like '%;%'
     end
     select @a=@a+1 
end    

delete #DP from #DP where not demand_part in
(select distinct IECPN from OPS_OPO_new where Customer='HP' and ReportDate=convert(char(10),getdate(),111) and OpenQty>0 and MP='N')

--select distinct * from #DP where Material='6019B1415501'


create table #tmp_BLETA(iid int identity(1,1),Material varchar(20),MaterialETA varchar(1000))
create table #BLETA(Material varchar(20),MaterialETA varchar(1000),ETA char(10),Qty int)

--Only Multiple ETA need to count
insert #tmp_BLETA
    select distinct Material,[Material ETA] from #BL where [Material ETA] like '%,%' 

---Add single item to #SH_ETA
insert #BLETA
    select distinct Material,[Material ETA],'',0 from #BL where not [Material ETA] like '%,%' 
    
declare @c int
declare @d int
select @c=min(iid) from #tmp_BLETA
select @d=max(iid) from #tmp_BLETA
select @d=1

while @c<=@d
begin
     while (select charindex(',',MaterialETA) from #tmp_BLETA where iid=@c)>0
     begin
         insert #BLETA
            select distinct Material,left(MaterialETA,charindex(',',MaterialETA)-1),'','' from #tmp_BLETA where iid=@c
         update #tmp_BLETA set MaterialETA=substring(MaterialETA,charindex(',',MaterialETA)+1,1000) where iid=@c
         
         insert #BLETA
            select distinct Material,rtrim(MaterialETA),'','' from #tmp_BLETA where iid=@c and not MaterialETA like '%,%'
     end
     select @c=@c+1 
end


update #BLETA set MaterialETA='2999/01/01*0' where (MaterialETA='_' or MaterialETA='')

--update ETA /Qty
update #BLETA set ETA=left(MaterialETA,charindex('*',MaterialETA)-1),Qty=convert(int,rtrim(substring(MaterialETA,charindex('*',MaterialETA)+1,1000)))

---Format Date
update #BLETA set ETA=convert(char(10),convert(datetime,rtrim(ETA)+' 00:00'),111) where not (MaterialETA='_' or MaterialETA='')

--select * from #BLETA
--select distinct * from #DP
--drop table #BLResult
select distinct b.demand_part,Model='                    ',OPOQty=0,a.Material,description,MG='                   ',ShortageQty,MaterialETA,ETA,Qty into #BLResult from #BLETA a,#DP b,#BL c 
where a.Material=b.Material and b.Material=c.Material order by b.demand_part,a.Material

update #BLResult set Model=b.Material_group from #BLResult a,t_download_matmas_CP60DW b where a.demand_part=b.Material
update #BLResult set MG=b.Material_group from #BLResult a,t_download_matmas_CP60DW b where a.Material=b.Material

update #BLResult set OPOQty=b.OpenQty from #BLResult a,
(select IECPN,OpenQty=sum(OpenQty) from OPS_OPO_new where Customer='HP' and ReportDate=convert(char(10),getdate(),111) 
and OpenQty>0 and MP='N' group by IECPN) as b where a.demand_part=b.IECPN

-----No ETA item ...
select maxETA,count(*) from(
select demand_part,maxETA=max(ETA) from #BLResult group by demand_part) as a group by maxETA

-----Shortage Material Type by Model

select a.Model,a.PT,b.OPOQty,MG,qty from
(
select Model,MG,PT,qty=count(*) from
(
select distinct Model,PT=substring(demand_part,8,2),MG,Material from #BLResult) as a group by Model,MG,PT --order by Model,count(*) desc
) as a inner join
(select Model,PT,OPOQty=sum(OPOQty) from 
(select distinct Model,PT=substring(demand_part,8,2),demand_part,OPOQty from #BLResult) as a  group by Model,PT) as b on a.Model=b.Model
and a.PT=b.PT order by Model,qty desc

select top 10 * from #BLResult --where Material='6019B1415501' demand_part='PFCU31AMB232'


select a.*,b.ServiceStock from (
select a.Model,a.PT,b.OPOQty,MG,Material,description,ShortageQty,MaterialETA from
(
select Model,PT,MG,Material,description,ShortageQty,MaterialETA,qty=count(*) from
(
select distinct Model,PT=substring(demand_part,8,2),MG,Material,description,ShortageQty,MaterialETA from #BLResult
where not MG='R-LABEL'
) as a group by Model,PT,MG,Material,description,ShortageQty,MaterialETA --order by Model,count(*) desc
) as a inner join
(select Model,PT,OPOQty=sum(OPOQty) from 
(select distinct Model,PT=substring(demand_part,8,2),demand_part,OPOQty from #BLResult) as a  group by Model,PT) as b on a.Model=b.Model
and a.PT=b.PT) as a,(select distinct Material,ServiceStock from #BL) as b where a.Material=b.Material order by Model,PT,MaterialETA



select a.*,ServiceStock,ManufactureStock,ManufactureExcessStock,VmiStock from
(
select a.*,b.ShortageQty from 
(
select Material,description,Item=count(*) from (
select distinct Material,description,demand_part from #BLResult
) as a group by Material,description 
) as a left join
(select distinct Material,ShortageQty from #BLResult) as b on a.Material=b.Material 
) as a left join #BL as b on a.Material=b.Material order by Item desc

*/