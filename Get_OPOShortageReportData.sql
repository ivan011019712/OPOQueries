drop table #DN
go
drop table #TodayDN
go
drop table #OPO
go
drop table #INV
go
drop table #OS
go
drop table #SL
go
drop table #SL1
go
drop table #SLResult
go
drop table #SLResult1
go
drop table #MT
go
drop table #ShipData
go
drop table #ww005
go
drop table #ww015
go
drop table #wwOMS
go
--drop table #MTaddSPStock
--go



--drop table #OS2	
--go
--drop table #OS3
--go 

declare @Customer varchar(20)
declare @Cust varchar(20)
select @Customer='DYNABOOK'
select @Cust=@Customer

---Get Today's DN
select Site,PO,SO,Item,IECPN,PO_Date,Qty856,IES_DN,IES_DNPGI into #TodayDN from Service_APD 
where Site in (Select distinct ZS92Site from SiteMapping where Customer=@Customer and not ZS92Site='') 
and left(SO,2) in ('11','12','20') and IES_DNPGI=case
when datepart(weekday,dateadd(dd,-1,getdate()))='1' then convert(char(10),dateadd(dd,-3,getdate()),111)
when datepart(weekday,dateadd(dd,-1,getdate()))='7' then convert(char(10),dateadd(dd,-2,getdate()),111)
else convert(char(10),dateadd(dd,-1,getdate()),111) end


if @Customer='FJ'
begin
insert #TodayDN
select Site,PO,SO,Item,IECPN,PO_Date=Date850,Qty856,IES_DN,IES_DNPGI from Service_APD where Site like '7054%'
and  left(SO,2) in ('11','12','20') and IES_DNPGI=case
when datepart(weekday,dateadd(dd,-1,getdate()))='1' then convert(char(10),dateadd(dd,-3,getdate()),111)
when datepart(weekday,dateadd(dd,-1,getdate()))='7' then convert(char(10),dateadd(dd,-2,getdate()),111)
else convert(char(10),dateadd(dd,-1,getdate()),111) end
end

if @Customer='ASUS'
begin
insert #TodayDN
select Site,PO,SO,Item,IECPN,PO_Date=Date850,Qty856,IES_DN,IES_DNPGI from Service_APD where PO in (
select distinct PO from (
select distinct PO from ZM57 where Site ='ASUS-CSC'
union
select distinct PO from Service_APD  where 
Site in (select distinct Site from ASUSSite where Loc='ICC')
--('32748','15883','24625','32745','15797','135406','1032','170338','246683','209016','233465','257447','248967',
--'245326','226139','17923','215119','165483','11149','169255','281887','234950','167200')
) as a 
) 
and  left(SO,2) in ('11','12','20') and IES_DNPGI=case
when datepart(weekday,dateadd(dd,-1,getdate()))='1' then convert(char(10),dateadd(dd,-3,getdate()),111)
when datepart(weekday,dateadd(dd,-1,getdate()))='7' then convert(char(10),dateadd(dd,-2,getdate()),111)
else convert(char(10),dateadd(dd,-1,getdate()),111) end
end


if @Customer='ASUS_ITH'
begin
insert #TodayDN
select Site,PO,SO,Item,IECPN,PO_Date=Date850,Qty856,IES_DN,IES_DNPGI from Service_APD where PO in (
select distinct PO from (
select distinct PO from ZM57 where Site ='ASUSTH-CSC'
union
select distinct PO from Service_APD  where 
Site in (select distinct Site from ASUSSite where Loc='ITH')
--('32748','15883','24625','32745','15797','135406','1032','170338','246683','209016','233465','257447','248967',
--'245326','226139','17923','215119','165483','11149','169255','281887','234950','167200')
) as a 
) 
and  left(SO,2) in ('11','12','20') and IES_DNPGI=case
when datepart(weekday,dateadd(dd,-1,getdate()))='1' then convert(char(10),dateadd(dd,-3,getdate()),111)
when datepart(weekday,dateadd(dd,-1,getdate()))='7' then convert(char(10),dateadd(dd,-2,getdate()),111)
else convert(char(10),dateadd(dd,-1,getdate()),111) end
end


----(2016/05/25) Get Original Inventory .
create table #INV(POVendor varchar(20),MP varchar(20),MatNo varchar(20),Qty int)

if @Customer='TOSHIBA'
begin
   select @Cust='TSB'
end

insert #INV
    select POVendor,MP,MatNo,Qty from Ivan_CurrentINV where Customer=@Cust and INVDate=convert(char(10),getdate(),111) 
    
    
------------Get Open PO (ALL)

/*
select iid,Site,PIC,PO,SO,IECPO,POItem,CPQNo,IECPN,ProductFamily,POVendor,POReceiveDate,PO_Type,
Model_Status,MP,FCST_Status,PType,NeedShipDate,POQty,DNQty,DockQty,ShipQty,
OpenQty,RMA_FG,CurrentDone,FutureDone,FutureFirstSupportDate,Shortage,SameMonthFCST,OPOR,Escalation,EscalationDate,Remark from OPS_OPO 
where Customer=@Customer and ReportDate=convert(char(10),getdate(),111) 
*/
---Get DN
select Site,PO,SO,Item,IECPN,PO_Date=Date850,Qty856,IES_DN,IES_DNPGI into #DN from Service_APD 
where /*Site in (Select distinct ZS92Site from SiteMapping where Customer=@Customer and not ZS92Site='') 
and */ PndGIDate='0000/00/00' and not IES_DNPGI='0000/00/00' and left(SO,2) in ('11','12','20') 

/*
if @Customer='FJ'
begin
insert #DN
   select Site,PO,SO,Item,IECPN,PO_Date,Qty856,IES_DN,IES_DNPGI from Service_APD where Site like '7054%'
and SO like '11%' and PndGIDate='0000/00/00' and not IES_DNPGI='0000/00/00' and SO like '11%'
end

if @Customer='ASUS'
begin
insert #DN
   select Site,PO,SO,Item,IECPN,PO_Date,Qty856,IES_DN,IES_DNPGI from Service_APD where 
   Site in (select distinct Site from ASUSSite)
--('32748','15883','24625','32745','15797','135406','1032','170338','246683','209016','233465','257447','248967',
--'245326','226139','17923','215119','165483','11149','169255','281887','234950','167200')
and SO like '11%' and PndGIDate='0000/00/00' and not IES_DNPGI='0000/00/00' and SO like '11%'
end
*/

update #DN set Item=b.IPCSOItem from #DN a,ZM57$ b where a.SO=b.IECSO and a.Item=b.IECSPItem 
update #DN set Item=b.IPCSOItem from #DN a,ZM57 b where a.SO=b.IECSO and a.Item=b.IECSPItem 




--(2021/11/11) Modify queries from left join to update to prevent duplicate
---(2016/03/18) ---Add DN information
select iid,Site,PIC,PO,SO,IECPO,POItem,CPQNo,IECPN,ProductFamily,POVendor,POReceiveDate,PO_Type,
Model_Status,MP,FCST_Status,PType,NeedShipDate,POQty,DNQty,DockQty,ShipQty,
OpenQty,RMA_FG,CurrentDone,FutureDone,FutureFirstSupportDate,Shortage,SameMonthFCST,OPOR,Escalation,EscalationDate,Remark,
IES_DNPGI='          ',IES_DN='                '  into #OPO
from OPS_OPO 
where Customer=@Customer and ReportDate=convert(char(10),getdate(),111) order by PIC,Site,iid

update #OPO set IES_DNPGI=isnull(b.IES_DNPGI,'')  from #OPO a,
(select SO,Item,IECPN,PO_Date,IES_DNPGI=max(IES_DNPGI) from #DN group by SO,Item,IECPN,PO_Date) as b where  
a.SO=b.SO and a.IECPN=b.IECPN and a.POItem=b.Item and a.POReceiveDate=b.PO_Date 

update #OPO set IES_DN=isnull(b.IES_DN,'')  from #OPO a,#DN b where a.SO=b.SO and a.IECPN=b.IECPN and a.POItem=b.Item and a.POReceiveDate=b.PO_Date and a.IES_DNPGI=b.IES_DNPGI

--select * from  #OPO where ProductFamily like '%EOSL%'
--update #OPO set ProductFamily=replace(ProductFamily,'_EOSL','')
update #OPO set ProductFamily=rtrim(ProductFamily)+'_EOSL' from #OPO a,ModelID b where a.Site='ASUS-CSC' and a.ProductFamily=b.SAPPlatform and EOSL_Dt<=getdate()

/*
select * from
(select distinct ProductFamily from #OPO) as a left join
(select * from ModelID) as b on a.ProductFamily=b.SAPPlatform
*/

select Site,PIC,OPOR,NPI_ItemQty,Normal_ItemQty from OPS_OPOR1 where Customer=@Customer and ReportDate=convert(char(10),getdate(),111) order by Customer,Site 

select OPOR,NPI_ItemQty,Normal_ItemQty from OPS_OPOR2 where Customer=@Customer and ReportDate=convert(char(10),getdate(),111) order by Customer,OPOR

--select * from OPS_OPOR2 where ReportDate=(select max(ReportDate) from OPS_OPOR2 where ReportDate<>convert(char(10),getdate(),111))order by Customer,OPOR
--select * from OPS_OPOSummary where ReportDate=convert(char(10),getdate(),111) order by Customer,PType,failqty desc

---For Old OSD 
select PType,Site,PIC,allItem,newItem,allqty,UpsideRate,nearfailItem,nearfailqty,failItem,failqty,failpercentage 
from OPS_OPOSummary where Customer=@Customer and ReportDate=convert(char(10),getdate(),111) order by Customer,PType,failqty desc


---For Old OSD (NPI)
select PType,Site,PIC,NPIallItem,NPInewItem,NPIallqty,0,NPInearfailItem,NPInearfailqty,NPIfailItem,NPIfailqty,NPIfailpercentage 
from OPS_OPOSummary where Customer=@Customer and ReportDate=convert(char(10),getdate(),111) order by Customer,PType,failqty desc


--select * from OPS_OPOSummary2 where Customer=@Customer and PType in ('FRU','Raw Material') and ReportDate=convert(char(10),getdate(),111) order by ReportDate desc,Customer,PType

---For Old OSD
--select Customer,ReportDate,PType,AllQty,FailQty,FailPercentage from OPS_OPOSummary2 where PType in ('FRU','Raw Material') order by ReportDate desc,Customer,PType

-----OPO & Shortage
select iid,Site,IECPN,NeedShipDate,a.POVendor,OpenQty,ShortagePN,Material_descript,NeedQty,OrginalStock=isnull(b.Qty,0),CurrentStock,
FutureStock,PIC,MaterialETA,Remark into #OS from
(
select iid,Site,IECPN,NeedShipDate,POVendor,OpenQty,ShortagePN,Material_descript,NeedQty,CurrentStock,
FutureStock,PIC,MaterialETA,Remark from OPS_OM where Customer=@Customer and ReportDate=convert(char(10),getdate(),111)
) as a left join
(
--select * from #INV where MP='FG'
select POVendor,MS=case when MS is null then MatNo else MS end,Qty=sum(Qty) from(

select distinct a.POVendor,MS,MatNo,Qty from (
select * from #INV where MP='FG') as a left join 
(select distinct MS,ShortagePN from OPS_Alt where Customer=@Customer) as b 
on a.MatNo=b.ShortagePN 

) as a 
group by POVendor,case when MS is null then MatNo else MS end

) as b on a.POVendor=b.POVendor and a.ShortagePN=b.MS order by  iid




----Material---------------------------------------------------------------------------------------------------
select a.POVendor,ShortagePN,PIC,Material_descript,GattingExisted,GattingNeed,
SAPRawMaterialOPO,POBalance,WholeBuyExisted,SODate,NeedShipDate,LatestSODate,LatestNeedShipDate,Site,Item,
RealOpenQty,OrginalStock=isnull(b.Qty,0),RemainedStock,MTStock,MStock,Buyer,Vendor,EstETADate,NewShortage,
MaterialETA,Remark,MD,a.MP,NPI into #MT from
(
select POVendor,ShortagePN,PIC,Material_descript,GattingExisted,GattingNeed,
SAPRawMaterialOPO,POBalance,WholeBuyExisted,SODate,NeedShipDate,LatestSODate,LatestNeedShipDate,Site,Item,
RealOpenQty,RemainedStock,MTStock,MStock,Buyer,Vendor,EstETADate,NewShortage,
MaterialETA,Remark,MD,MP,NPI from OPS_Material where Customer=@Customer and ReportDate=convert(char(10),getdate(),111) 
-----(2020/05/27) remove the items which RemainedStock<RealOpenQty
--and  RemainedStock<RealOpenQty
) as a left join
(
--select * from #INV where MP='FG'
select POVendor,MS=case when MS is null then MatNo else MS end,Qty=sum(Qty) from(

select distinct a.POVendor,MS,MatNo,Qty from (
select * from #INV where MP='FG') as a left join 
(select distinct MS,ShortagePN from OPS_Alt where Customer=@Customer) as b 
on a.MatNo=b.ShortagePN 

) as a 
group by POVendor,case when MS is null then MatNo else MS end

) as b on a.POVendor=b.POVendor and a.ShortagePN=b.MS order by PIC,RealOpenQty desc

update #OS set Material_descript='S'+a.Material_descript from #OS a,#MT b where a.ShortagePN=b.ShortagePN

-----(2024/10/28) Add SW01 015 Stock to MTStock--------------------------------------------------------------------------------------------------
/*
drop table #ww005
drop table #ww015
drop table #wwOMS
*/
-----Get Alternative 
select MS,MatNo,Qty into #ww005 from
(select * from WIP_WHALL_TD where SLoc='WA1' and SType='005') a,
(select MS,ShortagePN from  OPS_Alt where Customer=@Customer and ReportDate=convert(char(10),getdate(),111)) as b where a.MatNo=b.ShortagePN

-----Get Alternative 
select MS,MatNo,Qty into #ww015 from
(select * from WIP_WHALL_TD where SLoc='WA1' and SType='015') a,
(select MS,ShortagePN from  OPS_Alt where Customer=@Customer and ReportDate=convert(char(10),getdate(),111)) as b where a.MatNo=b.ShortagePN

-----Get Alternative --OMS
select MS,MatNo,Qty into #wwOMS from
(
select MatNo=Material,Qty from
(select * from OMS$ where StockDate=(select max(StockDate) from OMS$)) as a, t_download_matmas_CP60DW as b where a.HPPN=Old_Material
) a,
(select MS,ShortagePN from  OPS_Alt where Customer=@Customer and ReportDate=convert(char(10),getdate(),111)) as b where a.MatNo=b.ShortagePN


-----(2024/10/28) Add SW01 015 Stock to MTStock--------------------------------------------------------------------------------------------------end

/*
update #MT set Remark=replace(Remark,'<<No MP>>','')

update #MT set Remark='<<No MP>>'+rtrim(Remark) where ShortagePN in 
(
select distinct ShortagePN from 
(
select * from 
(select distinct IECPN,ShortagePN from OPS_OM where Customer=@Customer and ReportDate=convert(char(10),getdate(),111)) as a left join
(select * from Maggie_FRUnoMP) as b on a.IECPN=b.FRUPN
) as a where not FRUPN is null
)
*/

create table #SL(tid int identity(1,1),iid int,Mt varchar(1000),Rk varchar(5000))
create table #SLResult(iid int,Remark varchar(5000),Rk varchar(5000))
create table #SL1(tid int identity(1,1),iid int,Mt varchar(1000),Rk varchar(5000))
create table #SLResult1(iid int,Remark varchar(5000),Rk varchar(5000))

if @Customer='ASUS'
begin
	select a.*,ASUSPN=rtrim(b.Old_Material)/*+' : '+rtrim(b.Material_descript)*/ into #OS2 from #OS a,t_download_matmas_CP69DW b where left(a.ShortagePN,12)=b.Material

	update #OS2 set MaterialETA='' 
	where not MaterialETA like '%,%'  and not MaterialETA=''  
	and convert(char(10),convert(datetime,left( MaterialETA,charindex('*',MaterialETA)-1)+' 00:00:00'),111)<=convert(char(10),getdate(),111)
	
	

	select sid=identity(int,1,1),* into #OS3 from #OS2 
	where MaterialETA like '%,%'
	---select * from #OS3 where iid='172'
	---test----update #OS3 set MaterialETA=MaterialETA+',2020/06/08*100,2020/07/01*10' where sid=1

	declare @w int
	declare @z int
	declare @ME varchar(200)
	declare @MET varchar(20)
	declare @ME2 varchar(200)
	select @w=min(sid) from #OS3
	select @z=max(sid) from #OS3
	--select @z=5
	while @w<=@z
	begin
		select @ME=MaterialETA from #OS3 where sid=@w		
		select @ME2=''	 
		while not @ME=''
		begin		
			--select @ME
			if charindex(',',@ME)>0
			begin
				select @MET=left(@ME,charindex(',',@ME)-1)
				if convert(char(10),convert(datetime,left( @MET,charindex('*',@MET)-1)+' 00:00:00'),111)<=convert(char(10),getdate(),111)
				begin					
					select @ME=replace(@ME,@MET+',','')
				end
				else
				begin
					select @ME2=rtrim(@ME2)+@MET+','
					select @ME=rtrim(replace(@ME,@MET+',',''))
				end	
			end
			else
			begin
				select @MET=@ME
				if convert(char(10),convert(datetime,left( @MET,charindex('*',@MET)-1)+' 00:00:00'),111)<=convert(char(10),getdate(),111)
				begin					
					select @ME=replace(@ME,@MET,'')
				end
				else
				begin
					select @ME2=rtrim(@ME2)+@MET
					select @ME=rtrim(replace(@ME,@MET,''))
				end	
			end	
		end						
		update #OS3 set MaterialETA=@ME2 where sid=@w
		select @w=@w+1
		end
		--select * from #OS2 a,#OS3 b where a.iid=b.iid and a.IECPN=b.IECPN and a.ShortagePN=b.ShortagePN
		update #OS2 set MaterialETA=b.MaterialETA from #OS2 a,#OS3 b where a.iid=b.iid and a.IECPN=b.IECPN and a.ShortagePN=b.ShortagePN

		insert #SL
		select iid,Mt=
		isnull(ASUSPN+' ('+
		case when MaterialETA='' then 'No ETA' else MaterialETA end
		+')' ,''),Remark
		from #OS2 where Material_descript like 'S(%' 
		-----(2020/05/28) remove the items which RemainedStock<RealOpenQty
		--and not ShortagePN in (select ShortagePN from OPS_Material where Customer=@Customer and ReportDate=convert(char(10),getdate(),111) and  RemainedStock>=RealOpenQty)		
		order by iid ,Mt

	insert #SLResult
		select distinct iid,'','' from #SL order by iid

	declare @xx int
	declare @yy int
	select @xx=min(tid) from #SL
	select @yy=max(tid) from #SL
	while (@xx<=@yy)
		begin
			update #SLResult set Remark=rtrim(Remark)+' ; '+ b.Mt,Rk=rtrim(a.Rk) +' || '+ b.Rk from #SLResult a,(select iid,Mt,Rk from #SL where tid=@xx) as b where a.iid=b.iid
			select @xx=@xx+1
		end

-------(2022/12/28)  Add Priority 9 (¯S«æ) to OPO
-------(2021/11/05)  Add Priority to OPO
	update #OPO set PO=PO+'**' from #OPO a, ASUSPri$ b where a.PO=b.DN and b.Priority='9'
	update #OPO set PO=PO+'*' from #OPO a, ASUSPri$ b where a.PO=b.DN and b.Priority='1'

	select a.*,MRemark=isnull(b.Remark,'') ,Rk=isnull(b.Rk,'') from
	(select * from #OPO) as a left join 
	(select * from #SLResult) as b on a.iid=b.iid

	select * from #OS2  order by iid
    --select a.*,b.Old_Material from #MT a,t_download_matmas_CP69DW b where left(a.ShortagePN,12)=b.Material 	order by PIC,RealOpenQty desc
	select a.*,Priority=isnull(b.Priority,''),eMsg=isnull(b.eMsg,'') from 
	(select a.*,b.Old_Material from #MT a,t_download_matmas_CP69DW b where left(a.ShortagePN,12)=b.Material ) as a left join 
	(select * from Ivan_qi where ReportDate=convert(char(10),getdate(),111) and Customer=@Customer ) as b on a.ShortagePN=b.ShortagePN order by ShortagePN,Priority,eMsg

	drop table #OS2

end
----------------------------------------ASUSITH------------------------------------------------------
if @Customer='ASUS_ITH'
begin
	select a.*,ASUSPN=rtrim(b.Old_Material)/*+' : '+rtrim(b.Material_descript)*/ into #OS21 from #OS a,t_download_matmas_CP69DW b where left(a.ShortagePN,12)=b.Material

	update #OS21 set MaterialETA='' 
	where not MaterialETA like '%,%'  and not MaterialETA=''  
	and convert(char(10),convert(datetime,left( MaterialETA,charindex('*',MaterialETA)-1)+' 00:00:00'),111)<=convert(char(10),getdate(),111)
	
	

	select sid=identity(int,1,1),* into #OS31 from #OS21 
	where MaterialETA like '%,%'
	---select * from #OS3 where iid='172'
	---test----update #OS3 set MaterialETA=MaterialETA+',2020/06/08*100,2020/07/01*10' where sid=1

	declare @w1 int
	declare @z1 int
	declare @ME1 varchar(200)
	declare @MET1 varchar(20)
	declare @ME21 varchar(200)
	select @w1=min(sid) from #OS31
	select @z1=max(sid) from #OS31
	--select @z=5
	while @w<=@z
	begin
		select @ME1=MaterialETA from #OS31 where sid=@w1		
		select @ME21=''	 
		while not @ME1=''
		begin		
			--select @ME
			if charindex(',',@ME1)>0
			begin
				select @MET1=left(@ME1,charindex(',',@ME1)-1)
				if convert(char(10),convert(datetime,left( @MET1,charindex('*',@MET1)-1)+' 00:00:00'),111)<=convert(char(10),getdate(),111)
				begin					
					select @ME=replace(@ME1,@MET1+',','')
				end
				else
				begin
					select @ME21=rtrim(@ME21)+@MET1+','
					select @ME1=rtrim(replace(@ME1,@MET1+',',''))
				end	
			end
			else
			begin
				select @MET1=@ME1
				if convert(char(10),convert(datetime,left( @MET1,charindex('*',@MET1)-1)+' 00:00:00'),111)<=convert(char(10),getdate(),111)
				begin					
					select @ME1=replace(@ME1,@MET1,'')
				end
				else
				begin
					select @ME21=rtrim(@ME21)+@MET1
					select @ME1=rtrim(replace(@ME1,@MET1,''))
				end	
			end	
		end						
		update #OS31 set MaterialETA=@ME21 where sid=@w1
		select @w1=@w1+1
		end
		
		update #OS21 set MaterialETA=b.MaterialETA from #OS21 a,#OS31 b where a.iid=b.iid and a.IECPN=b.IECPN and a.ShortagePN=b.ShortagePN

		insert #SL1
		select iid,Mt=
		isnull(ASUSPN+' ('+
		case when MaterialETA='' then 'No ETA' else MaterialETA end
		+')' ,''),Remark
		from #OS21 where Material_descript like 'S(%' 
		-----(2020/05/28) remove the items which RemainedStock<RealOpenQty
		--and not ShortagePN in (select ShortagePN from OPS_Material where Customer=@Customer and ReportDate=convert(char(10),getdate(),111) and  RemainedStock>=RealOpenQty)		
		order by iid ,Mt

	insert #SLResult1
		select distinct iid,'','' from #SL1 order by iid

	declare @xx1 int
	declare @yy1 int
	select @xx1=min(tid) from #SL1
	select @yy1=max(tid) from #SL1
	while (@xx1<=@yy1)
		begin
			update #SLResult1 set Remark=rtrim(Remark)+' ; '+ b.Mt,Rk=rtrim(a.Rk) +' || '+ b.Rk from #SLResult1 a,(select iid,Mt,Rk from #SL1 where tid=@xx1) as b where a.iid=b.iid
			select @xx1=@xx1+1
		end

-------(2022/12/28)  Add Priority 9 (¯S«æ) to OPO
-------(2021/11/05)  Add Priority to OPO
	update #OPO set PO=PO+'**' from #OPO a, ASUSPri$ b where a.PO=b.DN and b.Priority='9'
	update #OPO set PO=PO+'*' from #OPO a, ASUSPri$ b where a.PO=b.DN and b.Priority='1'

	select a.*,MRemark=isnull(b.Remark,'') ,Rk=isnull(b.Rk,'') from
	(select * from #OPO) as a left join 
	(select * from #SLResult) as b on a.iid=b.iid

	select * from #OS21  order by iid
    --select a.*,b.Old_Material from #MT a,t_download_matmas_CP69DW b where left(a.ShortagePN,12)=b.Material 	order by PIC,RealOpenQty desc
	select a.*,Priority=isnull(b.Priority,''),eMsg=isnull(b.eMsg,'') from 
	(select a.*,b.Old_Material from #MT a,t_download_matmas_CP69DW b where left(a.ShortagePN,12)=b.Material ) as a left join 
	(select * from Ivan_qi where ReportDate=convert(char(10),getdate(),111) and Customer=@Customer ) as b on a.ShortagePN=b.ShortagePN order by ShortagePN,Priority,eMsg

	drop table #OS21
end
----------------------------------------End ASUSITH----------------------------------------------
--else
--begin	
if @Customer in ('HP','H2','DYNABOOK','HP_ITH','HP_IMP')
begin
	if @Customer='HP'
	begin
		---(2022/04/29) add OFilm Info in OSSPPN
		update #OS set IECPN=rtrim(a.IECPN)+'*' from #OS a,OFilmPN$ b where a.IECPN=b.IECPN
	end
	select * from #OPO
	select * from #OS  order by iid
	--select * from #MT order by PIC,RealOpenQty desc


	select a.*,Priority=isnull(b.Priority,''),eMsg=isnull(b.eMsg,''),ww005=0,ww015=0,wwOMS=0 into #MTaddSPStock from 
	(select * from #MT ) as a left join 
	(select * from Ivan_qi where ReportDate=convert(char(10),getdate(),111) and Customer=@Customer) as b on  a.ShortagePN=b.ShortagePN order by PIC,Priority,eMsg

	update #MTaddSPStock set ww005=b.Qty from #MTaddSPStock a,#ww005 b where a.ShortagePN=b.MatNo
	update #MTaddSPStock set ww015=b.Qty from #MTaddSPStock a,#ww015 b where a.ShortagePN=b.MatNo
	update #MTaddSPStock set wwOMS=b.Qty from #MTaddSPStock a,#wwOMS b where a.ShortagePN=b.MatNo

	select * from #MTaddSPStock
	drop table #MTaddSPStock	
end
--end



------------------------------------------------------------------------------------------------------------------

select PIC,Item,Qty,NoETAItem,NoETAQty from OPS_MSummaryPIC where Customer=@Customer and ReportDate=convert(char(10),getdate(),111) order by Customer


select MType,PIC,Item,Qty,NoETAItem,NoETAQty from  OPS_MSummaryType where Customer=@Customer and ReportDate=convert(char(10),getdate(),111) order by Customer,Qty desc


select MS,ShortagePN,Priority,Usage,Item,RefPN from  OPS_Alt where Customer=@Customer and ReportDate=convert(char(10),getdate(),111) order by Customer,MS


select POVendor,MaterialProperty,MatNo,Material_descript,RemainedQty,StockExistedDate,ConsumerStatus from  OPS_Inventory 
where Customer=@Customer and ReportDate=convert(char(10),getdate(),111) order by Customer,MatNo


select PO_Number,Item,Plnt,WH,Material,Description,Vendor,Buyer,Create_Dt,PO_Qty,Open_Qty,MaterialETA,Remark from OPS_RawOPO where Customer=@Customer and ReportDate=convert(char(10),getdate(),111) order by Customer


----DN HitRate
if (select count(*) from #TodayDN)=0
begin
select 'No DN'
end
else
begin
select IES_DNPGI,OpenDN=sum(OpenDN),Total=sum(Total),HitRate=convert(decimal(8,2),convert(float,sum(Total)-sum(OpenDN))/convert(float,sum(Total))*100) from 
(
select IES_DNPGI,OpenDN=count(*),Total=0 from #OPO where /*IES_DNPGI=convert(char(10),dateadd(dd,-1,getdate()),111)*/
IES_DNPGI=case
when datepart(weekday,dateadd(dd,-1,getdate()))='1' then convert(char(10),dateadd(dd,-3,getdate()),111)
when datepart(weekday,dateadd(dd,-1,getdate()))='7' then convert(char(10),dateadd(dd,-2,getdate()),111)
else convert(char(10),dateadd(dd,-1,getdate()),111) end

 and DNQty>0 group by IES_DNPGI
union
select IES_DNPGI,0,count(*) from #TodayDN group by IES_DNPGI
) as a group by IES_DNPGI
end

create table #ShipData(Site varchar(20),ProductFamily varchar(20),CPQNo varchar(20),IECPN varchar(20),
PO varchar(20),PO_Type varchar(20),FCST_Status varchar(20),Model_Status varchar(20),
SO varchar(20),SOItem varchar(20),POReceiveDate char(10),ShipQty int,NeedShipDate char(10),
IES_DN varchar(20),DNDate char(10),ShipDate char(10),RSD char(10))


--(2020/04/07) Add RSD per HP Annie's request.
--select * from Service_APD where SO='1107743144'
------Shipped data

if @Customer='FJ' 
begin
insert #ShipData
select *  from
(
select distinct Site,ProductFamily,CPQPN,IECPN,PO,PO_Type,FCST_Status,Model_Status,SO,Item,Date850,Qty856,SO_First_Date,IES_DN,IES_DNPGI,PndGIDate,
RSD=left(SO_ReqDate,4)+'/'+substring(SO_ReqDate,5,2)+'/'+substring(SO_ReqDate,7,2) from Service_APD where 
Site in (Select distinct ZS92Site from SiteMapping where Customer=@Customer and not ZS92Site='') 
and PndGIDate>=convert(char(10),dateadd(dd,-40,getdate()),111)
and not PndGIDate='0000/00/00'
union
select distinct Site,ProductFamily,CPQPN,IECPN,PO,PO_Type,FCST_Status,Model_Status,SO,Item,Date850,Qty856,SO_First_Date,IES_DN,IES_DNPGI,PndGIDate,
RSD=left(SO_ReqDate,4)+'/'+substring(SO_ReqDate,5,2)+'/'+substring(SO_ReqDate,7,2) from Service_APD where 
Site like '70%' and PndGIDate>=convert(char(10),dateadd(dd,-40,getdate()),111)
and not PndGIDate='0000/00/00'
) as a order by PndGIDate
end
else if @Customer='TOSHIBA' 
begin
insert #ShipData
select distinct Site,ProductFamily,CPQPN,IECPN,PO,Order_Reason,FCST_Status,Model_Status,SO,Item,Date850,Qty856,SO_First_Date,IES_DN,IES_DNPGI,PndGIDate,
RSD=left(SO_ReqDate,4)+'/'+substring(SO_ReqDate,5,2)+'/'+substring(SO_ReqDate,7,2) from Service_APD where 
Site in (Select distinct ZS92Site from SiteMapping where Customer='TOSHIBA' and not ZS92Site='') 
and PndGIDate>=convert(char(10),dateadd(dd,-40,getdate()),111) and left(IECPN,1) in ('T','6')
and not PndGIDate='0000/00/00' order by PndGIDate
end
else if @Customer='ASUS' 
begin
insert #ShipData
select distinct Site,ProductFamily,CPQPN,IECPN,PO,Order_Reason,FCST_Status,Model_Status,SO,Item,Date850,Qty856,SO_First_Date,IES_DN,IES_DNPGI,PndGIDate,
RSD=left(SO_ReqDate,4)+'/'+substring(SO_ReqDate,5,2)+'/'+substring(SO_ReqDate,7,2) from Service_APD where  PO in (
select distinct PO from (
select distinct PO from ZM57 where Site ='ASUS-CSC'
union
select distinct PO from Service_APD  where 
Site in (select distinct Site from ASUSSite where Loc='ICC')
--('32748','15883','24625','32745','15797','135406','1032','170338','246683','209016','233465','257447','248967',
--'245326','226139','17923','215119','165483','11149','169255','281887','234950','167200')
) as a 
) 
and PndGIDate>=convert(char(10),dateadd(dd,-40,getdate()),111) and left(IECPN,1) in ('L','6','R')
and not PndGIDate='0000/00/00' order by PndGIDate
end
else if @Customer='DYNABOOK' 
begin
insert #ShipData
select distinct Site,ProductFamily,CPQPN,IECPN,PO,Order_Reason,FCST_Status,Model_Status,SO,Item,Date850,Qty856,SO_First_Date,IES_DN,IES_DNPGI,PndGIDate,
RSD=left(SO_ReqDate,4)+'/'+substring(SO_ReqDate,5,2)+'/'+substring(SO_ReqDate,7,2) from Service_APD where 
Site in (Select distinct ZS92Site from SiteMapping where Customer='DYNABOOK' and not ZS92Site='') 
and PndGIDate>=convert(char(10),dateadd(dd,-40,getdate()),111) and left(IECPN,1) in ('L','6')
and not PndGIDate='0000/00/00' order by PndGIDate
end
else if @Customer='ASUS_ITH' 
begin
insert #ShipData
select distinct Site,ProductFamily,CPQPN,IECPN,PO,Order_Reason,FCST_Status,Model_Status,SO,Item,Date850,Qty856,SO_First_Date,IES_DN,IES_DNPGI,PndGIDate,
RSD=left(SO_ReqDate,4)+'/'+substring(SO_ReqDate,5,2)+'/'+substring(SO_ReqDate,7,2) from Service_APD where  PO in (
select distinct PO from (
select distinct PO from ZM57 where Site ='ASUSTH-CSC'
union
select distinct PO from Service_APD  where 
Site in (select distinct Site from ASUSSite where Loc='ITH')
--('32748','15883','24625','32745','15797','135406','1032','170338','246683','209016','233465','257447','248967',
--'245326','226139','17923','215119','165483','11149','169255','281887','234950','167200')
) as a 
) 
and PndGIDate>=convert(char(10),dateadd(dd,-40,getdate()),111) and left(IECPN,1) in ('L','6','R')
and not PndGIDate='0000/00/00' order by PndGIDate
end
else if @Customer='DYNABOOK' 
begin
insert #ShipData
select distinct Site,ProductFamily,CPQPN,IECPN,PO,Order_Reason,FCST_Status,Model_Status,SO,Item,Date850,Qty856,SO_First_Date,IES_DN,IES_DNPGI,PndGIDate,
RSD=left(SO_ReqDate,4)+'/'+substring(SO_ReqDate,5,2)+'/'+substring(SO_ReqDate,7,2) from Service_APD where 
Site in (Select distinct ZS92Site from SiteMapping where Customer='DYNABOOK' and not ZS92Site='') 
and PndGIDate>=convert(char(10),dateadd(dd,-40,getdate()),111) and left(IECPN,1) in ('L','6')
and not PndGIDate='0000/00/00' order by PndGIDate
end
else if @Customer='ACER' 
begin
insert #ShipData
select distinct Site,ProductFamily,CPQPN,IECPN,PO,Order_Reason,FCST_Status,Model_Status,SO,Item,Date850,Qty856,SO_First_Date,IES_DN,IES_DNPGI,PndGIDate,
RSD=left(SO_ReqDate,4)+'/'+substring(SO_ReqDate,5,2)+'/'+substring(SO_ReqDate,7,2) from Service_APD where PO in (
select distinct PO from (
select distinct PO from (
select * from ZM57 where Site in (select distinct SoldToParty from SiteMapping where Customer='ACER')
union
select * from ZM57 where Site in (select distinct ShipToParty from SiteMapping where Customer='ACER')
) as a
union
select distinct PO from Service_APD a,SiteMapping b where 
a.Site=b.ZS92Site and b.Customer='ACER' and not ZS92Site=''
) as a
)
and PndGIDate>=convert(char(10),dateadd(dd,-40,getdate()),111) and left(IECPN,1) in ('L','6')
and not PndGIDate='0000/00/00' order by PndGIDate
end
else
begin
insert #ShipData
select distinct Site,ProductFamily,CPQPN,IECPN,
PO=case when PO in ('14100777','14100780','14100781','14100782','14100783',
'14100784','14100785','14100786','14100789','14110504','14110604','14129933','14135339') then rtrim(PO)+' (Tariff)'
when PO in ('14149391','14149393') then rtrim(PO)+' (IND BTB)'
when PO in ('14152531','14152542','14152549','14152813','14152815','14152822','14152823','14152976','14154791','14155764','14155766','14162279','14162285') then rtrim(PO)+' (IND BTB 2)'
when PO in ('14158402','14158411','14158413','14165677','14165633','14165649','14165659','14165677','14181963') then rtrim(PO)+' (IND BTB 3)'
when PO in ('114162175','14162180','14165705','14165724','14165738','14170303','14170305','14165705','14165724','14162175','14177162','14177163') then rtrim(PO)+' (IND BTB 4)'
when PO in ('14158415','14158433','14162808','14162811') then rtrim(PO)+' (Tariff 2)'
when PO in ('14170309','14170312','14170313','14170315','14175564','14175578','14175584','14175586','14175588','14175592','14175606','14175609','14175611') then rtrim(PO)+' (IND BTB 5)'
when PO in ('14182384','14182389','14182883','14182888','14182891','14182893','14185211','14185218','14185224','14185211','14185218','14185224','14186499','14186501','14188677','14197041','14198790','14200347','14202091') then rtrim(PO)+' (IND BTB 6)'
else PO end
,PO_Type,FCST_Status,Model_Status,SO,Item,Date850,Qty856,SO_First_Date,IES_DN,IES_DNPGI,PndGIDate,
RSD=left(SO_ReqDate,4)+'/'+substring(SO_ReqDate,5,2)+'/'+substring(SO_ReqDate,7,2) from Service_APD where 
Site in (Select distinct ZS92Site from SiteMapping where Customer=@Customer and not ZS92Site='') 
and PndGIDate>=convert(char(10),dateadd(dd,-40,getdate()),111)
and not PndGIDate='0000/00/00' order by PndGIDate
end

--(2020/10/16) Add "*" as long as the PO ever set to "P1"
update #ShipData set PO=a.PO+'*' from #ShipData a,SMSCM b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='P1' 
update #ShipData set PO=a.PO+'*' from #ShipData a,SMSCM_History b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='P1' and not a.PO like '%*'


---(2019/03/06) add SMS SMSCQ Containment 
update #ShipData set PO=rtrim(a.PO)+'_CS' from #ShipData a,SMSCM b where replace(a.PO,'*','')=b.PO and a.CPQNo=b.OSSPPN and b.PT='CS'
---(2019/03/25) add SMS SMSCQ Containment 2 
update #ShipData set PO=rtrim(a.PO)+'_CS2' from #ShipData a,SMSCM b where replace(a.PO,'*','')=b.PO and a.CPQNo=b.OSSPPN and b.PT='CS2'
---(2019/04/23) add SMS SMSCQ Containment 3 
update #ShipData set PO=rtrim(a.PO)+'_CS3' from #ShipData a,SMSCM b where replace(a.PO,'*','')=b.PO and a.CPQNo=b.OSSPPN and b.PT='CS3'
---(2019/05/24) add SMS SMSCQ Containment 4 
update #ShipData set PO=rtrim(a.PO)+'_CS4' from #ShipData a,SMSCM b where replace(a.PO,'*','')=b.PO and a.CPQNo=b.OSSPPN and b.PT='CS4'
---(2019/06/20) add SMS SMSCQ Containment 5 
update #ShipData set PO=rtrim(a.PO)+'_CS5' from #ShipData a,SMSCM b where replace(a.PO,'*','')=b.PO and a.CPQNo=b.OSSPPN and b.PT='CS5'
---(2019/07/29) add SMS SMSCQ Containment 6 
update #ShipData set PO=rtrim(a.PO)+'_CS6' from #ShipData a,SMSCM b where replace(a.PO,'*','')=b.PO and a.CPQNo=b.OSSPPN and b.PT='CS6'

---(2019/08/27) add SMS SMSCQ Priority PO 
update #ShipData set PO=rtrim(a.PO)+'_P' from #ShipData a,SMSCM b where replace(a.PO,'*','')=b.PO and a.CPQNo=b.OSSPPN and b.PT='P'

---(2023/05/15) add SMS BP (BigBuy POs) 
update #ShipData set PO=rtrim(a.PO)+'_BP' from #ShipData a,SMSBP b where replace(a.PO,'*','')=b.PO and a.CPQNo=b.OSSPPN and b.PT='BP'

---(2023/06/04) add BR_BO
update #ShipData set PO=rtrim(a.PO)+'_BR_BO' from #ShipData a,SMSBP b where replace(a.PO,'*','')=b.PO and a.CPQNo=b.OSSPPN and b.PT='BR_BO'

---(2025/10/30) add EMEAKBQ
update #ShipData set PO=rtrim(a.PO)+'_EMEAKBQ' from #ShipData a,SMSBP b where replace(a.PO,'*','')=b.PO and a.CPQNo=b.OSSPPN and b.PT='EMEAKBQ'

---(2025/04/20) add AMSKBQ
update #ShipData set PO=rtrim(a.PO)+'_AMSKBQ' from #ShipData a,SMSBP b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='AMSKBQ'

---(2025/05/04) add VDS
update #ShipData set PO=rtrim(a.PO)+'_VDS' from #ShipData a,SMSBP b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='VDS'

---(2023/06/29) add SMS TR (Tariff POs) 
update #ShipData set PO=rtrim(a.PO)+'_TR' from #ShipData a,SMSTR b where replace(a.PO,'*','')=b.PO and a.CPQNo=b.OSSPPN and b.PT='TR'

---(2025/01/07) add SMS TR2 (Tariff POs) 
update #ShipData set PO=rtrim(a.PO)+'_TR2' from #ShipData a,SMSTR b where replace(a.PO,'*','')=b.PO and a.CPQNo=b.OSSPPN and b.PT='TR2'

---(2025/01/07) add SMS TR12 (Tariff POs) 
update #ShipData set PO=rtrim(a.PO)+'_TR12' from #ShipData a,SMSTR b where replace(a.PO,'*','')=b.PO and a.CPQNo=b.OSSPPN and b.PT='TR12'




---(2024/01/16) add SMS APJSLA
update #ShipData set PO=rtrim(a.PO)+'_APJSLA' from #ShipData a,SMSTR b where replace(a.PO,'*','')=b.PO and a.CPQNo=b.OSSPPN and b.PT='APJSLA'


---(2023/07/24) add SMS APIRR
update #ShipData set PO=rtrim(a.PO)+'_APIRR' from #ShipData a,SMSAPIRR b where replace(a.PO,'*','')=b.PO and a.CPQNo=b.OSSPPN and b.PT='APIRR'

---(2023/08/04) add SMS Top 300
update #ShipData set PO=rtrim(a.PO)+'_APJ300' from #ShipData a,SMSTop300 b where replace(a.PO,'*','')=b.PO and a.CPQNo=b.OSSPPN and b.PT='APJ300'

---(2024/12/26) Add charindex to ensure all "_" items can be added
---(2020/02/19) add SMS SMSCQ Priority PO 
update #ShipData set PO=rtrim(a.PO)+'_P1' from #ShipData a,SMSCM b where left(a.PO,case when charindex('_',a.PO)>0 then charindex('_',a.PO)-1 else 20 end)=b.PO and a.CPQNo=b.OSSPPN and b.PT='P1'
update #ShipData set PO=rtrim(a.PO)+'_P2' from #ShipData a,SMSCM b where left(a.PO,case when charindex('_',a.PO)>0 then charindex('_',a.PO)-1 else 20 end)=b.PO and a.CPQNo=b.OSSPPN and b.PT='P2'
update #ShipData set PO=rtrim(a.PO)+'_P3' from #ShipData a,SMSCM b where left(a.PO,case when charindex('_',a.PO)>0 then charindex('_',a.PO)-1 else 20 end)=b.PO and a.CPQNo=b.OSSPPN and b.PT='P3'
update #ShipData set PO=rtrim(a.PO)+'_P4' from #ShipData a,SMSCM b where left(a.PO,case when charindex('_',a.PO)>0 then charindex('_',a.PO)-1 else 20 end)=b.PO and a.CPQNo=b.OSSPPN and b.PT='P4'

----(2020/12/23) from HP Critical Items
update #ShipData set PO=rtrim(a.PO)+'_P0' from #ShipData a,SMSCM b where left(a.PO,case when charindex('_',a.PO)>0 then charindex('_',a.PO)-1 else 20 end)=b.PO and a.CPQNo=b.OSSPPN and b.PT='P0'
update #ShipData set PO=rtrim(a.PO)+'_P5' from #ShipData a,SMSCM b where left(a.PO,case when charindex('_',a.PO)>0 then charindex('_',a.PO)-1 else 20 end)=b.PO and a.CPQNo=b.OSSPPN and b.PT='P5'

---(2024/12/26) Add charindex for "*"
update #ShipData set PO=rtrim(a.PO)+'_P1' from #ShipData a,SMSCM b where left(a.PO,case when charindex('*',a.PO)>0 then charindex('*',a.PO)-1 else 20 end)=b.PO and a.CPQNo=b.OSSPPN and b.PT='P1'
update #ShipData set PO=rtrim(a.PO)+'_P2' from #ShipData a,SMSCM b where left(a.PO,case when charindex('*',a.PO)>0 then charindex('*',a.PO)-1 else 20 end)=b.PO and a.CPQNo=b.OSSPPN and b.PT='P2'
update #ShipData set PO=rtrim(a.PO)+'_P3' from #ShipData a,SMSCM b where left(a.PO,case when charindex('*',a.PO)>0 then charindex('*',a.PO)-1 else 20 end)=b.PO and a.CPQNo=b.OSSPPN and b.PT='P3'
update #ShipData set PO=rtrim(a.PO)+'_P4' from #ShipData a,SMSCM b where left(a.PO,case when charindex('*',a.PO)>0 then charindex('*',a.PO)-1 else 20 end)=b.PO and a.CPQNo=b.OSSPPN and b.PT='P4'
update #ShipData set PO=rtrim(a.PO)+'_P0' from #ShipData a,SMSCM b where left(a.PO,case when charindex('*',a.PO)>0 then charindex('*',a.PO)-1 else 20 end)=b.PO and a.CPQNo=b.OSSPPN and b.PT='P0'
update #ShipData set PO=rtrim(a.PO)+'_P5' from #ShipData a,SMSCM b where left(a.PO,case when charindex('*',a.PO)>0 then charindex('*',a.PO)-1 else 20 end)=b.PO and a.CPQNo=b.OSSPPN and b.PT='P5'



----(2024/06/11) Add EXPRESS
update #ShipData set PO=rtrim(a.PO)+'_EXPRESS' from #ShipData a,SMSCM b where replace(a.PO,'*','')=b.PO and a.CPQNo=b.OSSPPN and b.PT='EXPRESS'

---(2019/08/29) add SMS SMSCQ BTB 
update #ShipData set PO=rtrim(a.PO)+'_BTB' from #ShipData a,SMSCM b where replace(a.PO,'*','')=b.PO and a.CPQNo=b.OSSPPN and b.PT='BTB'

---(2025/11/13) Find HP ITH and IMP
update #ShipData set Site='SMSUS_IMP' from #ShipData a,Service_APD b where a.Site=b.Site and a.SO=b.SO and b.Site='HP-FSMSUS1' and Plant='UM60'
update #ShipData set Site='SMSUS_ITH' from #ShipData a,Service_APD b where a.Site=b.Site and a.SO=b.SO and b.Site='HP-FSMSUS1' and Plant='TP01'

select * from #ShipData

--------(2020/08/07) Addd Checking where have X01 BOM
declare @PP varchar(500)
declare @PT varchar(20)
--declare @Customer varchar(20)
--select @Customer='FJ'

select @PT=case 
when @Customer='HP' then 'CP60DW'
when @Customer='HP_ITH' then 'TH02DW'
when @Customer='HP_IMP' then 'UM60DW'
when @Customer='TOSHIBA' then 'CP62DW'
when @Customer='FJ' then 'CP62DW'
when @Customer='DYNABOOK' then 'CP62DW'
when @Customer='ACER' then 'CP69DW'
when @Customer='ASUS' then 'CP69DW'
when @Customer='ASUS_ITH' then 'TH03DW'
when @Customer='TINY' then 'CP65DW' end

select @PP='select'''+@Customer+''' ,IECPN,ProductFamily from OPS_OPO where Customer='''+@Customer+'''  and ReportDate=convert(char(10),getdate(),111)
and IECPN in (select distinct Material from t_download_org_bom_'+@PT+ ' where substring(Material,2,1)=''F'' and Bom_status=''05'')'

execute (@PP)


--------------------------------------------------------
-------------------------------------------------------
--------------------------------------------------------
-------------------------------------------------------
----All OTD Failed Summary 
--------------------------------------------------------
---------------------------------------------------------------------------------------------------------------
-------------------------------------------------------


--select * from  OPS_OTDFailD where ReportDate=convert(char(10),getdate(),111) order by Customer,OPOType,OpenQty desc
--select * from  OPS_OTDFailD where ReportDate=(select max(ReportDate) from OPS_OTDFailD where ReportDate<>convert(char(10),getdate(),111)) order by Customer,OPOType,OpenQty desc

----Summary
select Customer,MType,PType,ItemQty,PTypeQty from  OPS_OTDFailS where ReportDate=convert(char(10),getdate(),111) and OPOType='Normal OTD' order by Customer,OPOType,PTypeQty desc
select Customer,MType,PType,ItemQty,PTypeQty from  OPS_OTDFailS where ReportDate=convert(char(10),getdate(),111) and OPOType='NPI OTD' order by Customer,OPOType,PTypeQty desc
select Customer,MType,PType,ItemQty,PTypeQty from  OPS_OTDFailS where ReportDate=convert(char(10),getdate(),111) and OPOType='Normal Potential' order by Customer,OPOType,PTypeQty desc
select Customer,MType,PType,ItemQty,PTypeQty from  OPS_OTDFailS where ReportDate=convert(char(10),getdate(),111) and OPOType='NPI Potential' order by Customer,OPOType,PTypeQty desc

/*
----Summary last day
select Customer,MType,PType,ItemQty,PTypeQty from  OPS_OTDFailS where ReportDate=(select max(ReportDate) from OPS_OTDFailS where ReportDate<>convert(char(10),getdate(),111)) and OPOType='Normal OTD' order by Customer,OPOType,PTypeQty desc
select Customer,MType,PType,ItemQty,PTypeQty from  OPS_OTDFailS where ReportDate=(select max(ReportDate) from OPS_OTDFailS where ReportDate<>convert(char(10),getdate(),111)) and OPOType='NPI OTD' order by Customer,OPOType,PTypeQty desc
select Customer,MType,PType,ItemQty,PTypeQty from  OPS_OTDFailS where ReportDate=(select max(ReportDate) from OPS_OTDFailS where ReportDate<>convert(char(10),getdate(),111)) and OPOType='Normal Potential' order by Customer,OPOType,PTypeQty desc
select Customer,MType,PType,ItemQty,PTypeQty from  OPS_OTDFailS where ReportDate=(select max(ReportDate) from OPS_OTDFailS where ReportDate<>convert(char(10),getdate(),111)) and OPOType='NPI Potential' order by Customer,OPOType,PTypeQty desc
*/


--select Customer,OPOType,MType,PType,ItemQty,PTypeQty from  OPS_OTDFailS where ReportDate=(select max(ReportDate) from OPS_OTDFailS where ReportDate<>convert(char(10),getdate(),111)) order by Customer,OPOType,PTypeQty desc


----Detail
select Customer,IECPN,ProductFamily,NeedShipDate,OpenQty,ShortagePN,Material_descript,MaterialETA,Remark from  OPS_OTDFailD 
where ReportDate=convert(char(10),getdate(),111) and OPOType='Normal OTD' order by Customer,OPOType,OpenQty desc

select Customer,IECPN,ProductFamily,NeedShipDate,OpenQty,ShortagePN,Material_descript,MaterialETA,Remark from  OPS_OTDFailD 
where ReportDate=convert(char(10),getdate(),111) and OPOType='NPI OTD' order by Customer,OPOType,OpenQty desc

select Customer,IECPN,ProductFamily,NeedShipDate,OpenQty,ShortagePN,Material_descript,MaterialETA,Remark from  OPS_OTDFailD 
where ReportDate=convert(char(10),getdate(),111) and OPOType='Normal Potential' order by Customer,OPOType,OpenQty desc

select Customer,IECPN,ProductFamily,NeedShipDate,OpenQty,ShortagePN,Material_descript,MaterialETA,Remark from  OPS_OTDFailD 
where ReportDate=convert(char(10),getdate(),111) and OPOType='NPI Potential' order by Customer,OPOType,OpenQty desc

/*
----Detail Last Day
select Customer,IECPN,ProductFamily,NeedShipDate,OpenQty,ShortagePN,Material_descript,MaterialETA,Remark from  OPS_OTDFailD 
where ReportDate=(select max(ReportDate) from OPS_OTDFailS where ReportDate<>convert(char(10),getdate(),111)) and OPOType='Normal OTD' order by Customer,OPOType,OpenQty desc

select Customer,IECPN,ProductFamily,NeedShipDate,OpenQty,ShortagePN,Material_descript,MaterialETA,Remark from  OPS_OTDFailD 
where ReportDate=(select max(ReportDate) from OPS_OTDFailS where ReportDate<>convert(char(10),getdate(),111)) and OPOType='NPI OTD' order by Customer,OPOType,OpenQty desc

select Customer,IECPN,ProductFamily,NeedShipDate,OpenQty,ShortagePN,Material_descript,MaterialETA,Remark from  OPS_OTDFailD 
where ReportDate=(select max(ReportDate) from OPS_OTDFailS where ReportDate<>convert(char(10),getdate(),111)) and OPOType='Normal Potential' order by Customer,OPOType,OpenQty desc

select Customer,IECPN,ProductFamily,NeedShipDate,OpenQty,ShortagePN,Material_descript,MaterialETA,Remark from  OPS_OTDFailD 
where ReportDate=(select max(ReportDate) from OPS_OTDFailS where ReportDate<>convert(char(10),getdate(),111)) and OPOType='NPI Potential' order by Customer,OPOType,OpenQty desc




select * from  OPS_OTDFailD where ReportDate=(select max(ReportDate) from OPS_OTDFailD where ReportDate<>convert(char(10),getdate(),111)) order by Customer,OPOType,OpenQty desc
*/

-----Get Mail content
select * from OPS_OPOSummary2 where PType in ('FRU','Raw Material') and ReportDate=convert(char(10),getdate(),111) order by ReportDate desc,Customer,PType
