----Check ZM57
--select * from ZM57 where IPCSO=''
--delete from ZM57 where IPCSO=''
/*
drop table #INV
drop table #init
drop table #init2
drop table #Mat_PF
drop table #Mat_SA
drop table #MPinit
drop table #MPmid
drop table #MPresult
--drop table #Dock
drop table #init_tmp_OPO
drop table #FSD
drop table #TT
drop table #ETA
--(cancel)drop table #BOMinit
--(cancel)drop table #mid
--(cancel)drop table #final
drop table #init131
drop table #mid131
drop table #final131
drop table #finalPF
drop table #tmp2 
drop table #tmp3
drop table #initAssy
drop table #midAssy
drop table #finalAssy
drop table #GetSub
drop table #AA
drop table #AAmid
--drop table #b2b
drop table #FindQty
drop table #midFindQty
----Need to keep
drop table #tmp_OPO
drop table #result
drop table #AAResult
drop table #Alternative


--drop table #oriINV
drop table #tmpSeq
drop table #tmpWH
drop table #midWH
drop table #altWH
drop table #WHResult
Drop table #ETAfinal
drop table #WHFinal
drop table #tmpETA
drop table #WHmid
drop table #WHPN
drop table #midFinal
drop table #WHcount
drop table #ETAQty  
drop table #OPOSummary
drop table #DD
drop table #MV
drop table #VD
drop table #OPODetail
drop table #ShortageDetail
*/

------------Get Open PO (ALL)
create table #INV(POVendor varchar(20),MaterialProperty varchar(20),MatNo varchar(20),Qty int,INVDate char(10),Remark varchar(2000))

insert #INV
    select POVendor,MP,MatNo,Qty,INVDate,'' from Ivan_CurrentINV  where Customer='DYNABOOK' and INVDate=convert(char(10),getdate(),111) 
 
create table #init_tmp_OPO(Site varchar(20),PO varchar(20),SO varchar(20),Item varchar(20),IECPO varchar(20),POItem varchar(20),
CPQNo varchar(20),IECPN varchar(20),ProductFamily varchar(20),POVendor varchar(20),
POReceiveDate char(10),SORequestDate char(10),CustETADate char(10),PO_Type varchar(20),Model_Status varchar(20),FCST_Status varchar(20),PType varchar(20),TCDate char(10),NeedShipDate char(10),
POQty int,DNQty int,DockQty int,ShipQty int,OpenQty int,Shortage varchar(10),PIC varchar(20),MP varchar(10),Escalation varchar(20),EscalationDate varchar(20),Remark varchar(1000))

create table #tmp_OPO(iid int identity(1,1),Site varchar(20),PO varchar(20),SO varchar(20),Item varchar(20),IECPO varchar(20),POItem varchar(20),
CPQNo varchar(20),IECPN varchar(20),ProductFamily varchar(20),POVendor varchar(20),
POReceiveDate char(10),SORequestDate char(10),CustETADate char(10),PO_Type varchar(20),Model_Status varchar(20),FCST_Status varchar(20),PType varchar(20),TCDate char(10),NeedShipDate char(10),
POQty int,DNQty int,DockQty int,ShipQty int,OpenQty int,Shortage varchar(10),PIC varchar(20),MP varchar(10),Escalation varchar(20),EscalationDate varchar(20),Remark varchar(1000))



insert #init_tmp_OPO
select distinct Site,PO,SO,Item,'','',rtrim(CPQPN),IECPN,ProductFamily,'',Date850,SO_First_Date,'',
isnull(Order_Reason,''),isnull(Model_Status,''),isnull(FCST_Status,''),'',TCDate='',NeedShipDate='',
POQty=convert(int,Qty850),0,0,Ship=0,OpenQty=0,'','','','','','' from Service_APD a,SiteMapping b where 
a.Site=b.ZS92Site and b.Customer='DYNABOOK' and not ZS92Site=''/*IES_DNPGI='0000/00/00' and*/ and (SO like '11%' or SO like '2%' or SO like '12%') --and Plant in ('CP81','CP60','TP01')
and not IECPN like 'TF%'


--update DN Qty
update #init_tmp_OPO set DNQty=b.DNQty from #init_tmp_OPO a,
(select Site,IECPN,SO,Item,DNQty=convert(int,sum(Qty856)) from Service_APD where Qty856>0  and PndGIDate='0000/00/00' and not IES_DNPGI='0000/00/00'
group by Site,IECPN,SO,Item) as b where a.Site=b.Site and a.SO=b.SO and a.Item=b.Item and a.IECPN=b.IECPN


----Update Ship Qty
update #init_tmp_OPO set ShipQty=b.ShipQty from #init_tmp_OPO a,
(select Site,IECPN,SO,Item,ShipQty=convert(int,sum(Qty856)) from Service_APD where Qty856>0  and not PndGIDate='0000/00/00'--and not IES_DNPGI='0000/00/00'
group by Site,IECPN,SO,Item) as b where a.Site=b.Site and a.SO=b.SO and a.Item=b.Item and a.IECPN=b.IECPN



---Update POVendor ,PO,Type ....
update #init_tmp_OPO set POVendor=b.POVendor from #init_tmp_OPO a,ZSD65 b where a.SO=b.SO and a.Item=b.SOItem
update #init_tmp_OPO set POVendor='IES' where POVendor='IES-CP07'
update #init_tmp_OPO set POVendor='ICC' where POVendor='ICC-CP62'

---(2020/08/07) Force to Add POVendor.
update #init_tmp_OPO set POVendor='ICC' where POVendor=''
----SO initial is 2 ,set POVendor='IES'
--update #init_tmp_OPO set POVendor='IES' where POVendor=''


update #init_tmp_OPO set IECPO=b.IPCSO,POItem=b.IPCSOItem,PType=b.POType from #init_tmp_OPO a,ZM57$ b where a.SO=b.IECSO and a.Item=b.IECSPItem and a.IECPO=''


----(2017/11/06) need to sort out the HP_AIODOA SOs) ...
update #init_tmp_OPO set IECPO=convert(varchar(20),convert(decimal(20,0),b.IPCSO)),POItem=b.IPCSOItem,PType=b.POType from 
#init_tmp_OPO a,
(
select * from (
select IECSO,IECSPItem,IPCSO,IPCSOItem,POType from #init_tmp_OPO a,ZM57 b where a.SO=b.IECSO and a.Item=b.IECSPItem and a.IECPO='' and not a.SO in (
select distinct SO from Service_APD where SO like '12%' and IECPN like 'LF%')
) as a where not IPCSO=''
) as b where a.SO=b.IECSO and a.Item=b.IECSPItem and a.IECPO='' 




------Define MP/EOL ..
update #init_tmp_OPO set CustETADate=b.EOL_Dt from #init_tmp_OPO a,
(select CPQNo,EOL_Dt=convert(char(10),max(EOL_Dt),111) from PNModel a,ModelID b where a.FamilyNo=b.FamilyNo and CPQNo in (select distinct CPQNo from #init_tmp_OPO)
group by CPQNo) as b where a.CPQNo=b.CPQNo

update #init_tmp_OPO set Model_Status='MP' where CustETADate>=POReceiveDate
update #init_tmp_OPO set Model_Status='EOL' where CustETADate<POReceiveDate

--Check whether there is any PN without EOL information .
--select * from #init_tmp_OPO where SO='1106379479'
/*
--drop table #TT
--drop table #ETA
------Find ETA date of 2 ,4,5,9 weeks .
create table #TT(iid int identity(1,1),IssueDate varchar(10))
create table #ETA(IssueDate varchar(10),ETAWeek int,ETADate varchar(10))

insert #TT
select distinct POReceiveDate from #init_tmp_OPO

declare @x int
declare @y int
declare @dd varchar(10)
select @x=min(iid) from #TT
select @y=max(iid) from #TT
while @x<=@y
begin
-------Change Date to CalendarDate
    select @dd=IssueDate from #TT where iid=@x
/*
    insert #ETA
       select @dd,'2',convert(char(10),max(Dt),111) from(
       select Top 14/*10*/ DID,Dt from Calendar3 where Dt>convert(datetime,@dd+' 00:00') /*and WorkingDay='1'*/) as a 
    insert #ETA
       select @dd,'4',convert(char(10),max(Dt),111) from(
       select Top 28/*20*/ DID,Dt from Calendar3 where Dt>convert(datetime,@dd+' 00:00') /*and WorkingDay='1'*/) as a 
*/
    insert #ETA
       select @dd,'2',convert(char(10),max(Dt),111) from(
       select Top 14/*25*/ DID,Dt from Calendar3 where Dt>convert(datetime,@dd+' 00:00') /*and WorkingDay='1'*/) as a 
    insert #ETA
       select @dd,'4',convert(char(10),max(Dt),111) from(
       select Top 30/*45*/ DID,Dt from Calendar3 where Dt>convert(datetime,@dd+' 00:00') /*and WorkingDay='1'*/) as a 
   
    select @x=@x+1
end

----Update IECETA
update #init_tmp_OPO set TCDate=b.ETADate from #init_tmp_OPO a,#ETA b where a.POReceiveDate=b.IssueDate  and a.Model_Status='MP' and b.ETAWeek='2'
update #init_tmp_OPO set TCDate=b.ETADate from #init_tmp_OPO a,#ETA b where a.POReceiveDate=b.IssueDate  and a.Model_Status='EOL' and b.ETAWeek='4'
*/
----NPI set 2 weeks .
--update #init_tmp_OPO set TCDate=b.ETADate from #init_tmp_OPO a,#ETA b where a.POReceiveDate=b.IssueDate and a.PO_Type='NPI' and b.ETAWeek='2'


update #init_tmp_OPO set OpenQty=POQty-ShipQty
delete from #init_tmp_OPO where OpenQty=0

------(2015/04/09) update POVendor for IHC PO
update #init_tmp_OPO set POVendor=b.[PO2 Delivery Plant] from #init_tmp_OPO a,zpu38 b where a.IECPO=[PO1 Number]  and a.POItem=b.[PO1 Item]

update #init_tmp_OPO set POVendor='IES' where POVendor='CP07'
update #init_tmp_OPO set POVendor='ICC' where POVendor='CP62'


---Update CustETADate (-7 days is the Cust NeedShipDate)
---(2014/09/05) remove deduction
--Update #init_tmp_OPO set CustETADate=(select convert(char(10),max(Dt),111) from Calendar3 where Dt<=dateadd(dd,-7,convert(datetime,SORequestDate+' 00:00')) and WorkingDay=1)


---Update TCDate='' to '2099/01/01'  
update #init_tmp_OPO set TCDate='2099/01/01' where TCDate=''
update #init_tmp_OPO set TCDate=SORequestDate


----Compare NeedShipDate & SORequestDate
update #init_tmp_OPO set NeedShipDate=case when TCDate>=SORequestDate then TCDate else SORequestDate end
--update #init_tmp_OPO set NeedShipDate=TCDate where PO_Type='NPI'


-----Update Site to MSite
update #init_tmp_OPO set Site=b.MSite from #init_tmp_OPO a,(select distinct MSite,ZS92Site from SiteMapping where not ZS92Site='') as b where a.Site=b.ZS92Site

---Update PIC
update #init_tmp_OPO set PIC=b.PIC from #init_tmp_OPO a,OSSPPIC b where a.Site=b.Site


----Define MP
update #init_tmp_OPO set MP='Y' where IECPN like '6%' and MP='' --Raw Material


update #init_tmp_OPO set MP=b.MP from #init_tmp_OPO a,
(select CPQNo,StopProduceDate=dateadd(mm,-3,max(EOL_Dt)),MP=case when dateadd(mm,-3,max(EOL_Dt))>=getdate() then 'Y' else 'N' end from (
select a.CPQNo,EOL_Dt from
(select distinct CPQNo from #init_tmp_OPO) as a inner join
(select CPQNo,EOL_Dt from PNModel a,ModelID b where a.FamilyNo=b.FamilyNo ) as b on a.CPQNo=b.CPQNo
) as a group by CPQNo
) as b where a.CPQNo=b.CPQNo and a.MP=''



---Fill not find data MP to 'Y'
update #init_tmp_OPO set MP='Y' where MP=''

-----(2015/07/08) User request to add all open MO------Start
/*
-----(2015/07/08) Add MO
update #init_tmp_OPO set Remark=b.MO from #init_tmp_OPO a,
(select a.[SO#],a.[IEC#],
MO='<< PC Plan : '+convert(char(10),convert(datetime,replace(MO#,right(MO#,4),'')),111)+'*'
+convert(varchar(10),b.Qty)/*+'  ('+right(MO#,4)+', '+rtrim(isnull(convert(char(10),b.Shipment_Date,111),''))+'),  '+rtrim(b.Note)*/+' >>',
b.Qty,
MO#=replace(MO#,right(MO#,4),''),Plant=right(MO#,4),b.Note,
Shipment_Date=isnull(convert(char(10),b.Shipment_Date,111),'') from [10.99.252.108].[Plan].dbo.SO_Detial a,[10.99.252.108].[Plan].dbo.MO b where a.ID=b.ID
and convert(datetime,replace(MO#,right(MO#,4),''))>=getdate()) as b where
a.SO COLLATE Chinese_Taiwan_Stroke_BIN=b.[SO#] and a.IECPN COLLATE Chinese_Taiwan_Stroke_BIN=b.[IEC#] 
*/
--drop table #TSB_MO
--drop table #TSBMO
create table #TSB_MO(SO varchar(20),IECPN varchar(20),MO varchar(500))
create table #TSBMO(iid int identity(1,1),SO varchar(20),IECPN varchar(20),MO varchar(50))

/*
insert #TSB_MO
select a.[SO#],a.[IEC#],
MO='<< PC Plan : '+convert(char(10),convert(datetime,replace(MO#,right(MO#,4),'')),111)+'*'+convert(varchar(10),b.Qty)+' >>'
/*,b.Qty,MO#=replace(MO#,right(MO#,4),''),Plant=right(MO#,4),b.Note*/
from [10.96.3.127].[Plan].dbo.SO_Detial a,[10.96.3.127].[Plan].dbo.MO b where a.ID=b.ID and Shipment_Date is null
*/

insert #TSBMO 
select a.* from #TSB_MO a,
(
select SO,IECPN from (
select SO,IECPN,qty=count(*) from #TSB_MO group by SO,IECPN) as a where qty>1
) as b where a.SO=b.SO and a.IECPN=b.IECPN order by SO,IECPN,MO

delete from #TSB_MO from #TSB_MO a,
(
select SO,IECPN from (
select SO,IECPN,qty=count(*) from #TSB_MO group by SO,IECPN) as a where qty>1
) as b where a.SO=b.SO and a.IECPN=b.IECPN

declare @momin int
declare @momax int
declare @mSO varchar(20)
declare @mIECPN varchar(20)
declare @mMO varchar(50)
select @momin=min(iid) from #TSBMO
select @momax=max(iid) from #TSBMO

while (@momin<=@momax)
begin
      select @mSO=SO,@mIECPN=IECPN,@mMO=MO from #TSBMO where iid=@momin
      
      if exists(select * from #TSB_MO where SO=@mSO and IECPN=@mIECPN)
      begin
         update #TSB_MO set MO=rtrim(MO)+rtrim(@mMO) where SO=@mSO and IECPN=@mIECPN
      end
      else
      begin
         insert #TSB_MO values(@mSO,@mIECPN,@mMO)
      end
      select @momin=@momin+1
end
-----Add End signal
update #TSB_MO set MO=rtrim(MO)+'|'

update #init_tmp_OPO set Remark=b.MO from #init_tmp_OPO a,#TSB_MO b where
a.SO COLLATE Chinese_Taiwan_Stroke_BIN=b.SO and a.IECPN COLLATE Chinese_Taiwan_Stroke_BIN=b.IECPN 

-----(2015/07/08) User request to add all open MO------End

--update TSB_OPO$ set Remark=replace(Remark,left(Remark,charindex('>>',Remark)+1),'') where Remark like '%<< PC Plan%>>%'
update TSB_OPO$ set Remark=replace(Remark,left(Remark,charindex('>>|',Remark)+2),'') where Remark like '%<< PC Plan%>>|%'

--Update Escalation
update #init_tmp_OPO set /*Escalation=b.Escalation,EscalationDate=b.[Escalation Date],*/Remark=rtrim(a.Remark)+b.Remark from #init_tmp_OPO a,TSB_OPO$ b where a.Site=b.Site and a.SO=b.SO and a.IECPN=b.IECPN


/*
-----Change NPI order's NeedShipDate to FSD Date if the NeedShipDate earlier than FSD date .
create table #FSD(IECPN varchar(20),FSD varchar(20),FRUShipDate varchar(20))

insert #FSD
select distinct IECPN,'','' from #init_tmp_OPO where PO_Type='NPI'

update #FSD set FSD=b.FSD from #FSD a,
(
select ODMPartNumber,FSD=min(FSD) from (
select ODMPartNumber,FSD from
(
select * from
(select distinct SpareKitPN,ODMPartNumber from SPB where ODMPartNumber in (select distinct IECPN from #FSD)) as a left join
(select PN,FSD from [IEC1-CUSTOMER].EService.dbo.PNProject) as b on a.SpareKitPN=b.PN
) as a where not FSD is null
) as a group by ODMPartNumber
) as b where a.IECPN=b.ODMPartNumber

select * from #FSD where FSD=''
update #FSD set FSD='1900/01/01' where FSD=''

update #FSD set FRUShipDate=convert(char(10),dateadd(mm,-1,convert(datetime,FSD+' 00:00')),111)
update #init_tmp_OPO set NeedShipDate=b.FRUShipDate from #init_tmp_OPO a,#FSD b where a.PO_Type='NPI' and a.IECPN=b.IECPN and FRUShipDate>NeedShipDate
*/
-----(2014/01/14)Insert the old OPO which not existed in Service_APD back to #init_tmp_OPO
/*
insert #init_tmp_OPO
select 
Site,PO,SO,Item,IECPO,POItem,CPQNo,IECPN,ProductFamily,POVendor,POReceiveDate,SORequestDate,CustETADate,PO_Type,Model_Status,FCST_Status,PType ,TCDate ,NeedShipDate ,
POQty,DNQty,DockQty,ShipQty,OpenQty,Shortage,PIC,MP,Escalation,EscalationDate,Remark
 from #tmp_OPO where iid in(
select a.iid from(
select a.iid,b.SO from 
(select * from #tmp_OPO) as a left join
(select * from #init_tmp_OPO) as b on a.Site=b.Site and a.PO=b.PO and a.SO=b.SO and a.Item=b.Item and a.IECPN=b.IECPN
) as a where SO is null
)
*/

--update Site to 'FJ EDI' for the 705 PO ..
--update #init_tmp_OPO set Site='FJ EDI' where Site like '705%'

--Remove strange data
delete from #init_tmp_OPO where Site=''

--Update Special Orders' NeedShipDate
update #init_tmp_OPO set NeedShipDate=convert(char(10),dateadd(dd,30,convert(datetime,POReceiveDate+' 00:00:00')),111) where Site in ('KOYO','IHC') and PType in ('NB','ZB')
update #init_tmp_OPO set NeedShipDate=convert(char(10),dateadd(dd,14,convert(datetime,POReceiveDate+' 00:00:00')),111) where Site in ('KOYO','IHC') and PType in ('ZN','ZC')

--Update Model
--update #init_tmp_OPO set ProductFamily=b.Material_group from #init_tmp_OPO a,t_download_matmas_CP07DW b where a.IECPN=b.Material and a.ProductFamily=''
update #init_tmp_OPO set ProductFamily=b.Material_group from #init_tmp_OPO a,t_download_matmas_CP62DW b where a.IECPN=b.Material and a.ProductFamily=''


insert #tmp_OPO
   select distinct * from #init_tmp_OPO order by NeedShipDate          

-------------------------(2019/12/13)
---select * from #tmp_OPO 

/*
drop table #MPinit
drop table #MPmid
drop table #MPresult
*/
--create table #MPIECPN(IECPN varchar(20),ASPN varchar(20))
create table #MPinit(POVendor varchar(20),IECPN varchar(20),ASPN varchar(20),SAPN varchar(20),AG varchar(20),Priority varchar(20),Usage varchar(20),Item varchar(20),Qty float)
create table #MPmid(POVendor varchar(20),IECPN varchar(20),ASPN varchar(20),SAPN varchar(20),AG varchar(20),Priority varchar(20),Usage varchar(20),Item varchar(20),Qty float)
create table #MPresult(POVendor varchar(20),IECPN varchar(20),ASPN varchar(20),SAPN varchar(20),AG varchar(20),Priority varchar(20),Usage varchar(20),Item varchar(20),Qty float)

insert #MPinit
   select distinct POVendor,IECPN,IECPN,Component,Alternative_item_group,Priority,Usage_probability,Item_number,Quantity from (
   select distinct POVendor,IECPN from #tmp_OPO where /*MP='Y' and*/ not IECPN like '6%') as a,t_download_org_bom_CP62DW b where a.IECPN=b.Material and left(b.Component,1) in ('1','2')

--update #MPinit set POVendor=case when ShipSite='SH' then 'IES' else 'ICC' end from #MPinit a,PNSite b where a.IECPN=b.IECPN and a.POVendor=''

while (select count(*) from #MPinit)>0
begin
insert #MPresult
    select distinct * from #MPinit where SAPN like '13%'  

delete from #MPinit  where SAPN like '13%'  

insert #MPmid
   select distinct * from #MPinit

delete from #MPinit

insert #MPinit
   select distinct a.POVendor,a.IECPN,a.SAPN,b.Component,b.Alternative_item_group,b.Priority,b.Usage_probability,b.Item_number,Quantity from
   #MPmid a,t_download_org_bom_CP62DW b where a.SAPN=b.Material and left(b.Component,1) in ('1','2')

delete #MPmid
end

-------(2020/03/12) Change Dock Qty allocation method by referring the items with DN only
/*
drop table #tmp_OPO_DN0
drop table #tmp_OPO_DN
drop table #DN
*/

select  *,IES_DN='--------------------',IES_DNPGI='0000/00/00'  into #tmp_OPO_DN0 from #tmp_OPO

---Get DN
select Site,PO,SO,Item,IECPN,PO_Date,Qty856,IES_DN,IES_DNPGI into #DN from Service_APD 
where Site in (Select distinct ZS92Site from SiteMapping where Customer='DYNABOOK' and not ZS92Site='') 
and PndGIDate='0000/00/00' and not IES_DNPGI='0000/00/00' and SO like '11%'

update #tmp_OPO_DN0 set IES_DNPGI=b.IES_DNPGI from #tmp_OPO_DN0 a,
(
select a.SO,SOItem=a.Item,Item=convert(int,IPCSOItem),a.IECPN,PO_Date,IES_DNPGI=max(IES_DNPGI) from #DN a,ZM57 b where 
a.SO=b.IECSO and a.Item=b.IECSPItem --and a.IECPN='JFBL01BMB002' 
group by a.SO,a.Item,convert(int,IPCSOItem),a.IECPN,a.PO_Date
) as b where a.SO=b.SO and a.IECPN=b.IECPN and a.POReceiveDate=b.PO_Date and a.POItem=b.Item

update #tmp_OPO_DN0 set IES_DN=b.IES_DN from #tmp_OPO_DN0 a,
(select distinct a.*,newItem=convert(int,b.IPCSOItem) from #DN a,ZM57 b where a.SO=b.IECSO and a.Item=b.IECSPItem) as b
where a.SO=b.SO and a.IECPN=b.IECPN and a.POReceiveDate=b.PO_Date and a.IES_DNPGI=b.IES_DNPGI and convert(int,a.POItem)=b.newItem


create table #tmp_OPO_DN(pid int identity(1,1),iid int ,Site varchar(20),PO varchar(20),SO varchar(20),Item varchar(20),IECPO varchar(20),POItem varchar(20),
CPQNo varchar(20),IECPN varchar(20),ProductFamily varchar(20),POVendor varchar(20),
POReceiveDate char(10),SORequestDate char(10),CustETADate char(10),PO_Type varchar(20),Model_Status varchar(20),FCST_Status varchar(20),PType varchar(20),TCDate char(10),NeedShipDate char(10),
POQty int,DNQty int,DockQty int,ShipQty int,OpenQty int,Shortage varchar(10),PIC varchar(20),MP varchar(10),Escalation varchar(100),EscalationDate varchar(20),Remark varchar(1000),
IES_DN varchar(20),IES_DNPGI  varchar(20))

insert #tmp_OPO_DN
	select *  from #tmp_OPO_DN0 where not DNQty=0 order by IES_DNPGI

--select * from #tmp_OPO_DN


declare @DX int
declare @DY int
declare @POVendor varchar(20)
declare @IECPN varchar(20)
declare @OpenQty int
declare @DNQty int
declare @DockQty int

select @DX=min(pid) from #tmp_OPO_DN
select @DY=max(pid) from #tmp_OPO_DN

while @DX<=@DY
begin
    select @POVendor=POVendor,@IECPN=IECPN,@OpenQty=OpenQty,@DNQty=DNQty from #tmp_OPO_DN where pid=@DX
    if @DNQty>0
    begin
		if exists(select * from #INV where POVendor=@POVendor and MatNo=@IECPN and MaterialProperty='SHIP' and Qty>0)
		begin
			select @DockQty=Qty from #INV where POVendor=@POVendor and MaterialProperty='SHIP' and MatNo=@IECPN
			if @DNQty>=@DockQty
			begin
			update #tmp_OPO_DN set DockQty=@DockQty where pid=@DX
			update #INV set Qty=0,Remark=rtrim(Remark)+';('+convert(varchar(5),@DX)+')-->'+convert(varchar(5),@DockQty) where POVendor=@POVendor and MatNo=@IECPN and MaterialProperty='SHIP'
			end
			else
			begin
			update #tmp_OPO_DN set DockQty=@DNQty where pid=@DX
			update #INV set Qty=@DockQty-@DNQty,Remark=rtrim(Remark)+';('+convert(varchar(5),@DX)+')-->'+convert(varchar(5),@DNQty) where POVendor=@POVendor and MatNo=@IECPN and MaterialProperty='SHIP'
			end         
		end
    end		
    select @DX=@DX+1
end

----If DockQty exist ,then deduct DNQty
update #tmp_OPO_DN set DNQty=DNQty-DockQty

update #tmp_OPO set DNQty=b.DNQty,DockQty=b.DockQty  from #tmp_OPO a,#tmp_OPO_DN b where a.iid=b.iid
--select * from #tmp_OPO
------------------------
------------------------
------------------------
------------------------
------------------------
-----Check initial #tmp_OPO ,especially for DockQty
--select count(*) from #tmp_OPO
--select * from #tmp_OPO where DockQty>0
--select * from #INV where not Remark=''



--select * from #init
------Get Raw Material & PF Relation from SA shortage report
create table #init(iid int identity(1,1),POVendor varchar(20),Material varchar(20),WU varchar(8000))
create table #Mat_PF(POVendor varchar(20),Pno varchar(20),Material varchar(20))



insert #init
select distinct POVendor,procument_part,demand_part from
(
select distinct POVendor='ICC',procument_part=left(part_no,12),demand_part from DYNABOOK_r_shortage_data where len(rtrim(part_no))>12  and demand_part like '%F%'
) as a 

declare @i int
declare @j int
select @i=min(iid) from #init
select @j=max(iid) from #init

while @i<=@j
begin
     while (select charindex(';',WU) from #init where iid=@i)>0
     begin
         insert #Mat_PF
            select POVendor,left(WU,charindex(';',WU)-1),Material from #init where iid=@i
         update #init set WU=substring(WU,charindex(';',WU)+1,1000) where iid=@i
     end
     select @i=@i+1 
end

----Keep LF only
delete from #Mat_PF where not (Pno like 'LF%' or Pno like '6%')
---select * from #Mat_PF

------Get Raw Material & Whereuse Relation from SA shortage report
-----(2015/04/27) for the ICC r_shortage_data "where_use" column which changed to material_group
-----(2015/04/28) modify more ,arrange current where_use back to the dara we need .
/*
drop table #pre_init2
drop table #mid_init2
drop table #MatCount
drop table #WhereUse
drop table #init2

drop table #init2
drop table #Mat_SA
*/
create table #pre_init2(iid int identity(1,1),POVendor varchar(20),Material varchar(20),where_use varchar(20))
create table #mid_init2(iid int,POVendor varchar(20),Material varchar(20),where_use varchar(20))
create table #MatCount(iid int identity(1,1),POVendor varchar(20),Material varchar(20),where_use varchar(5000))


insert #pre_init2
	select distinct POVendor='ICC',procument_part=left(part_no,12),where_use from DYNABOOK_r_shortage_data where len(rtrim(part_no))=12 and demand_part like '%F%'order by procument_part


insert #MatCount
select distinct POVendor,Material,'' from #pre_init2

declare @s int
declare @t int
declare @u int
declare @v int
select @s=min(iid) from #MatCount
select @t=max(iid) from #MatCount

while @s<=@t
begin
     insert #mid_init2
          select a.* from #pre_init2 a,(select * from #MatCount where iid=@s) as b where a.POVendor=b.POVendor and a.Material=b.Material order by a.iid
     select @u=min(iid) from #mid_init2
     select @v=max(iid) from #mid_init2
     while @u<=@v
     begin
     update #MatCount set where_use=rtrim(a.where_use)+b.where_use+';' from #MatCount a,
     (select * from #mid_init2 where iid=@u) as b where a.POVendor=b.POVendor and a.Material=b.Material and a.iid=@s     
     select @u=@u+1
     end
     truncate table #mid_init2
     select @s=@s+1
end



create table #init2(iid int identity(1,1),POVendor varchar(20),Material varchar(20),Buyer varchar(50),SA varchar(5000))
create table #Mat_SA(POVendor varchar(20),Pno varchar(20),Material varchar(20),Buyer Varchar(50))

insert #init2
select distinct POVendor,procument_part,buyer,where_use from
(
select distinct POVendor='ICC',procument_part=left(part_no,12),buyer,b.where_use from DYNABOOK_r_shortage_data a,
(select * from #MatCount where POVendor='ICC') as b where left(a.part_no,12)=b.Material and len(rtrim(part_no))>12 and demand_part like '%F%'
) as a 


declare @a int
declare @b int
select @a=min(iid) from #init2
select @b=max(iid) from #init2

while @a<=@b
begin
     while (select charindex(';',SA) from #init2 where iid=@a)>0
     begin
         insert #Mat_SA
            select POVendor,left(SA,charindex(';',SA)-1),Material,Buyer from #init2 where iid=@a
         update #init2 set SA=substring(SA,charindex(';',SA)+1,1000) where iid=@a
     end
     select @a=@a+1 
end



---Get tsb_icc_r_shortage_data
update #Mat_SA set Buyer=b.Buyer from #Mat_SA a,BuyerCode b where a.Buyer=b.BuyerCode



---Add MP M/B shortage .....
delete #Mat_PF from #Mat_PF a,#MPresult b where a.POVendor=b.POVendor and a.Pno=b.IECPN and a.Material=b.SAPN
insert #Mat_PF 
    select distinct POVendor,IECPN,SAPN from #MPresult
  
delete #Mat_SA from #Mat_SA a,#MPresult b where a.POVendor=b.POVendor and a.Pno=b.IECPN and a.Material=b.SAPN
insert #Mat_SA
    select distinct POVendor,IECPN,SAPN,'' from #MPresult

select distinct a.POVendor,a.Pno,a.Material,b.Buyer into #tmp_Mat_SA from #Mat_PF a,#Mat_SA b where a.Material=b.Material and a.Material like '15%'

insert #tmp_Mat_SA
    select distinct a.POVendor,a.Pno,a.Material,b.Buyer from #Mat_PF a,#Mat_SA b where a.Material=b.Material and a.Material like '6%' and b.Pno like '15%'

delete #Mat_SA from #Mat_SA a,#tmp_Mat_SA b where a.POVendor=b.POVendor and a.Pno=b.Pno and a.Material=b.Material and a.Buyer=b.Buyer

insert #Mat_SA
    select * from #tmp_Mat_SA

------------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------------------
-------(Modify on 2012/12/27) Find Top 13%
create table #init131(Material varchar(20),mid varchar(20),nextlevel varchar(20))
create table #mid131(Material varchar(20),mid varchar(20),nextlevel varchar(20))
create table #final131(Material varchar(20),mid varchar(20),nextlevel varchar(20))
create table #finalPF(Material varchar(20),mid varchar(20),nextlevel varchar(20))

insert #init131
   select distinct Pno,Pno,Pno from #Mat_SA where left(Pno,2) in ('11','13','19')

declare @count int
select @count=0

while (select count(*) from #init131)>0
begin
if @count>10
begin
    select 'Too Many loops'
    return
end

insert #mid131
select Material,mid,Mat from (
--select a.Material,a.mid,Mat=b.Material from #init131 a,t_download_org_bom_CP07DW b where a.mid=b.Component
--union
select a.Material,a.mid,Mat=b.Material from #init131 a,t_download_org_bom_CP62DW b where a.mid=b.Component
)as a 

insert #final131 
select distinct * from #mid131 where not nextlevel like '13%'

delete from #mid131 where not nextlevel like '13%'

delete from #init131
insert #init131
    select Material,nextlevel,nextlevel from #mid131    

delete from #mid131
select @count=@count+1
end

--select * from #final131

insert #finalPF
   select * from #final131 where left(nextlevel,2)='LF'
   
delete from #final131 where left(nextlevel,2)='LF'

delete from #init131
delete from #mid131

-----Find PF

insert #init131
 select distinct nextlevel,nextlevel,nextlevel from #final131

--declare @count int
select @count=0

while (select count(*) from #init131)>0
begin
if @count>10
begin
    select 'Too Many loops'
    return
end

insert #mid131
select Material,mid,Mat from (
--select a.Material,a.mid,Mat=b.Material from #init131 a,t_download_org_bom_CP07DW b where a.mid=b.Component
--union
select a.Material,a.mid,Mat=b.Material from #init131 a,t_download_org_bom_CP62DW b where a.mid=b.Component
)as a

insert #finalPF 
select distinct * from #mid131 where left(nextlevel,2)='LF'

delete from #mid131 where left(nextlevel,2)='LF'

delete from #init131
insert #init131
    select Material,nextlevel,nextlevel from #mid131    

delete from #mid131
select @count=@count+1
end



select distinct POVendor,PFnextlevel,PFmid,mid into #tmp2 from (
select * from
(
------
select distinct * from
(
select distinct Material,mid,PFmid,PFnextlevel from
(
select * from
(select * from #final131) as a left join
(select PFMaterial=Material,PFmid=mid,PFnextlevel=nextlevel from #finalPF) as b on a.nextlevel=b.PFMaterial
) as a where not PFMaterial is null
union
---(2013/01/04)Up level should be PF
select distinct Material,mid,nextlevel,nextlevel from #finalPF where left(nextlevel,2)='LF' and (Material like '13%' or Material like '11%')
) as a 
-------
) as a left join
(select distinct POVendor,PFPno=IECPN from #tmp_OPO
) as b on a.PFnextlevel=b.PFPno 
) as a where not PFPno is null  

--select distinct * from #tmp2
update #tmp2 set mid=PFmid where mid like '192%'

create table #result (POVendor varchar(20),PF varchar(20),RefPN varchar(20),ShortagePN varchar(20))

insert #result
    select distinct * from #tmp2

----(2013/02/04) Remove the 13% which RefPN exists PF & 13% both
delete from #result where ShortagePN in (
select distinct ShortagePN from #result where ShortagePN in (
select distinct ShortagePN from #result where RefPN=ShortagePN and ShortagePN like '13%') and left(RefPN,2)='LF') and RefPN=ShortagePN


----Insert the Material which parent is PF
insert #result
     select distinct POVendor,Pno,Pno,Material from #Mat_SA where (Material like '6%' or Material like '1%') and left(Pno,2)='LF'



----Find the PF for 1xx 
create table #initAssy(POVendor varchar(20),Assy varchar(20),PF varchar(20))
create table #midAssy(POVendor varchar(20),Assy varchar(20),PF varchar(20))
create table #finalAssy(POVendor varchar(20),Assy varchar(20),PF varchar(20))

--declare @count int
select @count=0

insert #initAssy
 select distinct POVendor,Pno,Pno from #Mat_SA where Material like '6%' and left (Pno,1)='1' and not left(Pno,2) in ('11','13','19','PF','SF','JF') 

while (select count(*) from #initAssy)>0
begin
if @count>10
begin
    select 'Too Many loops'
    return
end

insert #midAssy
select POVendor,Assy,Material from (
--select POVendor,a.Assy,b.Material from #initAssy a,t_download_org_bom_CP07DW b where a.PF=b.Component and not (b.Material like 'LC%' or b.Material like 'JC%')
--union
select POVendor,a.Assy,b.Material from #initAssy a,t_download_org_bom_CP62DW b where a.PF=b.Component and not (b.Material like 'LC%' or b.Material like 'JC%')
) as a

insert #finalAssy 
select distinct * from #midAssy where left(PF,2)='LF'

--delete from #midAssy from #midAssy a,#finalAssy b where a.Assy=b.Assy
delete from #midAssy where left(PF,2)='LF'

delete from #initAssy
insert #initAssy
    select * from #midAssy    

delete from #midAssy
select @count=@count+1
end
delete from  #midAssy   
insert #midAssy    
    select distinct * from #finalAssy
    
delete from #initAssy 
insert #initAssy  
     select * from #midAssy   

--select * from #initAssy 

select POVendor,Assy,PF into #tmp3 from (
select * from
(select * from #finalAssy) as a left join
(select distinct PV=POVendor,PFPno=IECPN from #tmp_OPO
--select distinct PV=POVendor,PFPno=Pno from #Mat_PF
) as b on a.PF=b.PFPno and a.POVendor=b.PV
) as a where not PFPno is null      
 
delete from #finalAssy
insert #finalAssy
    select * from #tmp3  
 
insert #result
select distinct a.POVendor,a.PF,b.Pno,b.Material from #finalAssy a,
(select distinct * from #Mat_SA where Material like '6%' and left (Pno,1)='1' and not left(Pno,2) in ('11','13','19','LF','JF')) as b
where a.POVendor=b.POVendor and a.Assy=b.Pno

---(2013/06/18)----Get Substitution (even 0%)
select distinct a.POVendor,PF,RefPN,b.Component into #GetSub from #result a,
(
/*
select distinct a.Material,a.Component from t_download_org_bom_CP07DW a,(
select b.Material,b.Alternative_item_group,b.Item_number from #result a,t_download_org_bom_CP07DW b where a.RefPN=b.Material and a.ShortagePN=b.Component) as b 
where a.Material=b.Material and a.Alternative_item_group=b.Alternative_item_group and a.Item_number=b.Item_number
union
*/
select distinct a.Material,a.Component from t_download_org_bom_CP62DW a,(
select b.Material,b.Alternative_item_group,b.Item_number from #result a,t_download_org_bom_CP62DW b where a.RefPN=b.Material and a.ShortagePN=b.Component) as b 
where a.Material=b.Material and a.Alternative_item_group=b.Alternative_item_group and a.Item_number=b.Item_number
) as b where a.RefPN=b.Material

delete from #result from #result a,#GetSub b where a.POVendor=b.POVendor and a.PF=b.PF and a.RefPN=b.RefPN and a.ShortagePN=b.Component

insert #result
    select * from #GetSub

-----------Strange Data ,manual modify    
--select * from #result where RefPN='1510B1248105' and ShortagePN='6037B0065512'
--update #result set ShortagePN='6037B0080212' where RefPN='1510B1379429' and ShortagePN='6037B0081012'
--update #result set ShortagePN='6037B0065612' where RefPN='1510B1248105' and ShortagePN='6037B0065512'

--(2013/01/04) -----Add Agency Label back .
--Remove Carton
delete from #result where ShortagePN like '606%' and not ShortagePN in (select distinct Material from t_download_matmas_CP62DW where Material_descript like 'LABEL,AGENCY%')

--------------Get Alternative-----(2012/11/09)There are some item which usage Qty is less than 1 ...
create table #AA(POVendor varchar(20),PF varchar(20),Component varchar(20),RefPN varchar(20),ShortagePN varchar(20),AG varchar(20),Priority varchar(20),Usage varchar(20),Item varchar(20),Qty float)
create table #AAmid(POVendor varchar(20),PF varchar(20),Component varchar(20),RefPN varchar(20),ShortagePN varchar(20),AG varchar(20),Priority varchar(20),Usage varchar(20),Item varchar(20),Qty float)
create table #AAResult(POVendor varchar(20),PF varchar(20),RefPN varchar(20),ShortagePN varchar(20),AG varchar(20),Priority varchar(20),Usage varchar(20),Item varchar(20),Qty float)


---Get Component
insert #AA
   select distinct POVendor,PF,PF,RefPN,ShortagePN,'','','','','' from #result --where PF in (select distinct IECPN from #MPresult)

---Find Single Shortage
update #AA set AG='NA' from #AA a,
(select distinct PF,RefPN from (
select PF,RefPN,qty=count(*) from #result group by PF,RefPN) as a where qty=1) as b where a.PF=b.PF and a.RefPN=b.RefPN

while (select count(*) from #AA)>0
begin   
--Find Component
update #AA set AG=b.Alternative_item_group,Priority=b.Priority,Usage=b.Usage_probability,Item=b.Item_number,Qty=/*convert(int,*/convert(float,replace(replace(b.Quantity,' ',''),',',''))/*)*/ from #AA a,
(
select distinct Material,Component,Alternative_item_group,Priority,Usage_probability,Item_number,Quantity from (
--select Material,Component,Alternative_item_group,Priority,Usage_probability,Item_number,Quantity from t_download_org_bom_CP07DW
--union
select Material,Component,Alternative_item_group,Priority,Usage_probability,Item_number,Quantity from t_download_org_bom_CP62DW
) as a 
) as b where a.Component=b.Material and a.ShortagePN=b.Component
and a.AG=''

-- If Find Component the AG will be Null (No substitution) or number  (Has Substitution) ,update Null the "NA"
update #AA set AG='NA' where AG is null

/*
select distinct a.POVendor,a.PF,a.Component,a.RefPN,ShortagePN=b.Component,AG=b.Alternative_item_group,
Priority=b.Priority,Usage=b.Usage_probability,Item=b.Item_number,Qty=convert(float,replace(replace(b.Quantity,' ',''),',','')) into #AATmp from 
(select * from #AA where not AG in ('','NA')) as a,t_download_org_bom_CP07DW b where a.Component=b.Material and a.Item=b.Item_number 
/*and not b.Usage_probability=' 0'*/ order by a.PF,a.Component,b.Item_number
*/

select distinct POVendor,PF,Component,RefPN,ShortagePN,AG,Priority,Usage,Item,Qty into #AATmp from (
/*
select distinct a.POVendor,a.PF,a.Component,a.RefPN,ShortagePN=b.Component,AG=b.Alternative_item_group,
Priority=b.Priority,Usage=b.Usage_probability,Item=b.Item_number,Qty=convert(float,replace(replace(b.Quantity,' ',''),',','')) from 
(select * from #AA where not AG in ('','NA')) as a,t_download_org_bom_CP07DW b where a.Component=b.Material and a.Item=b.Item_number 
union
*/
select distinct a.POVendor,a.PF,a.Component,a.RefPN,ShortagePN=b.Component,AG=b.Alternative_item_group,
Priority=b.Priority,Usage=b.Usage_probability,Item=b.Item_number,Qty=convert(float,replace(replace(b.Quantity,' ',''),',','')) from 
(select * from #AA where not AG in ('','NA')) as a,t_download_org_bom_CP62DW b where a.Component=b.Material and a.Item=b.Item_number 
) as a order by PF,Component,Item


delete from #AA where not AG in ('','NA') 

insert #AA
    select * from #AATmp

drop table #AATmp


--Update AG to Yes if find same item have multiple material
update #AA set AG='Y' from #AA a,(
select distinct Component,RefPN,Item from(
select Component,RefPN,Item,qty=count(*) from #AA where not AG in ('','NA') group by Component,RefPN,Item 
) as a where qty>1
) as b where a.Component=b.Component and a.RefPN=b.RefPN and a.Item=b.Item and not a.AG in ('','NA')

update #AA set AG='NA' where not AG in ('','NA','Y')


insert #AAResult
   select distinct POVendor,PF,RefPN,ShortagePN,AG,Priority,Usage,Item,Qty from #AA where AG in ('NA','Y') 

delete from #AA where AG in ('NA','Y') 


insert #AAmid
select distinct POVendor,PF,Component,RefPN,ShortagePN,'','','','',Qty from (

/*
select distinct POVendor,PF,b.Component,a.RefPN,a.ShortagePN,a.Qty--,b.Alternative_item_group,b.Priority,b.Usage_probability,b.Item_number
from #AA a,t_download_org_bom_CP07DW b where a.Component=b.Material and not b.Component like '6%'
and Alternative_item_group is null 
union
*/
select distinct POVendor,PF,b.Component,a.RefPN,a.ShortagePN,a.Qty--,b.Alternative_item_group,b.Priority,b.Usage_probability,b.Item_number
from #AA a,t_download_org_bom_CP62DW b where a.Component=b.Material and not b.Component like '6%'
and Alternative_item_group is null 
) as a order by PF,ShortagePN,RefPN
--and (Alternative_item_group is null or (not Alternative_item_group is null and not Usage_probability=' 0')) order by PF,ShortagePN,RefPN


delete from #AA

insert #AA
    select distinct * from #AAmid

delete from #AAmid


delete from #AA from #AA a,#AAResult b where a.POVendor=b.POVendor and a.PF=b.PF and a.ShortagePN=b.ShortagePN

end

----Strange data force to change 
--select * from #AAResult where ShortagePN='605406220203'
--update #AAResult set ShortagePN='605406220203' where ShortagePN='6037B0081012'
--update #AAResult set ShortagePN='6037B0065612' where ShortagePN='6037B0065512'
update #AAResult set Qty='1' where ShortagePN='605406220203'

---(2013/01/28) Insert Raw Material to #AAResult
insert #AAResult
    select distinct POVendor,IECPN,'0',IECPN,'NA','','','','1' from #tmp_OPO where IECPN like '6%'


----Find the Qty
create table #FindQty(PF varchar(20),SA varchar(20),RefPN varchar(20),ShortagePN varchar(20),Qty float)
create table #midFindQty(PF varchar(20),SA varchar(20),RefPN varchar(20),ShortagePN varchar(20),Qty float)
insert #FindQty
   select distinct PF,PF,RefPN,ShortagePN,0 from #AAResult where Qty=0
   

while (select count(*) from #FindQty)>0
begin   
--update #FindQty set Qty=convert(int,convert(float,replace(replace(b.Quantity,' ',''),',',''))) from #FindQty a,t_download_org_bom_CP07DW b where a.SA=b.Material and a.ShortagePN=b.Component
--and a.Qty=0 

update #FindQty set Qty=convert(int,convert(float,replace(replace(b.Quantity,' ',''),',',''))) from #FindQty a,t_download_org_bom_CP62DW b where a.SA=b.Material and a.ShortagePN=b.Component
and a.Qty=0 

--Update found Qty to #AAResult
update #AAResult set Qty=b.Qty from #AAResult a,(select distinct PF,RefPN,ShortagePN,Qty from #FindQty where Qty>0) as b where a.PF=b.PF 
and a.RefPN=b.RefPN and a.ShortagePN=b.ShortagePN and a.Qty=0

--Delete the Qty>0
delete from #FindQty where Qty>0

--Delete the extra data
delete from #FindQty from #FindQty a,#AAResult b where a.PF=b.PF and a.RefPN=b.RefPN and a.ShortagePN=b.ShortagePN and b.Qty>0


insert #midFindQty
     select distinct PF,Component,RefPN,ShortagePN,0 from (
     --select a.PF,b.Component,a.RefPN,a.ShortagePN from #FindQty a,t_download_org_bom_CP07DW b where a.SA=b.Material and not Component like '6%'
     --union
     select a.PF,b.Component,a.RefPN,a.ShortagePN from #FindQty a,t_download_org_bom_CP62DW b where a.SA=b.Material and not Component like '6%'
     ) as a


delete from #FindQty

insert #FindQty 
     select * from #midFindQty


delete from #midFindQty
end

-----------------------------------------------------------------------------------------------
/*
select * from #tmp_OPO where IECPN='6054B1138701'
update #tmp_OPO set Shortage='Y' from #tmp_OPO a,#result b where a.POVendor=b.POVendor and a.IECPN=b.PF
update #tmp_OPO set Shortage='N' where Shortage=''
*/

update #AAResult set Usage=replace(Usage,' ','')

---(2013/06/18)-----Get Alternative
--drop table #Alternative
create table #Alternative(POVendor varchar(20),PF varchar(20),MS varchar(20),RefPN varchar(20),ShortagePN varchar(20),
AG varchar(20),Priority varchar(20),Usage varchar(20),Item varchar(20),Qty float,INV int)

--drop table #tmpAlt
create table #tmpAlt(iid int identity(1,1),POVendor varchar(20),PF varchar(20),RefPN varchar(20),Item varchar(20))

--drop table #tmpAltResult
create table #tmpAltResult(POVendor varchar(20),PF varchar(20),RefPN varchar(20),MS varchar(20),Item varchar(20))

insert #tmpAlt
     select POVendor,PF,RefPN,Item from(
        select POVendor,PF,RefPN,Item,qty=count(*) from #AAResult where AG='Y'  group by POVendor,PF,RefPN,Item) as a where qty>1  
     
declare @Ai int
declare @Aj int
declare @APOVendor varchar(20)
declare @APF varchar(20)
declare @ARefPN varchar(20)
declare @AItem varchar(20)
select @Ai=min(iid) from #tmpAlt
select @Aj=max(iid) from #tmpAlt


while @Ai<=@Aj
begin
    select @APOVendor=POVendor,@APF=PF,@ARefPN=RefPN,@AItem=Item from #tmpAlt where iid=@Ai
           
    insert #tmpAltResult
         select top 1 POVendor,PF,RefPN,ShortagePN+'*',Item from #AAResult where 
              POVendor=@APOVendor and PF=@APF and RefPN=@ARefPN and Item=@AItem order by Usage desc,ShortagePN desc
    
    select @Ai=@Ai+1
end


insert #Alternative
     select distinct a.POVendor,a.PF,b.MS,a.RefPN,a.ShortagePN,a.AG,a.Priority,a.Usage,a.Item,a.Qty,0 from #AAResult a,#tmpAltResult b where 
              a.POVendor=b.POVendor and a.PF=b.PF and a.RefPN=b.RefPN and a.Item=b.Item

delete from #AAResult where AG='Y'

insert #AAResult
   select distinct POVendor,PF,RefPN,MS,'','','','',Qty from #Alternative


--(2013/01/18 --> Strange ,remark first ,forget why add this )
/*
----Remove multiple ShortagePN but exist in 17%
delete from #AAResult from #AAResult where PF in (
select PF from (
select PF,qty=count(*) from(
select a.* from #AAResult a,
(select PF,ShortagePN from #AAResult where RefPN like '17%') as b where a.PF=b.PF and a.ShortagePN=b.ShortagePN
) as a group by PF) as a where qty>1
) and RefPN like '17%'
*/

--drop table #tmpSeq
create table #tmpSeq(iid int identity(1,1),POVendor varchar(20),PF varchar(20),RefPN varchar(20))

insert #tmpSeq
select distinct a.POVendor,a.PF,a.RefPN from #AAResult a,(
select distinct PF from
(select PF,ShortagePN,qty=count(*) from #AAResult group by PF,ShortagePN) as a where qty>1 ) as b where a.PF=b.PF order by a.PF,RefPN

update #AAResult set RefPN=convert(varchar(10),b.iid) from #AAResult a,#tmpSeq b where a.POVendor=b.POVendor and a.PF=b.PF and a.RefPN=b.RefPN
update #AAResult set RefPN='0' where len(RefPN)='12' 
--select * from #AAResult where Qty<1
update #AAResult set Qty=1 where Qty<1

/*
delete #INV where MaterialProperty='FG'
insert #INV
  select POVendor,MP,MatNo,Qty,INVDate,'' from Ivan_CurrentINV  where INVDate>=convert(char(10),getdate(),111) and MP='FG'
drop table #tmpWH
drop table #midWH
drop table #altWH
drop table #WHResult
*/
---Debug
--update #INV set Qty='481' where MatNo='6054B0196601'
--insert #INV values('IES','FG','6054B0196601','1000','2013/01/31','')


create table #tmpWH(OpenQty int,RefPN int,ShortagePN varchar(20),Qty float,INV float)
create table #midWH(iid int identity(1,1),OpenQty float,ShortagePN varchar(20),Qty float,INV float,fillQty float,NeedQty float)
create table #altWH(iid int identity(1,1),NeedQty float,ShortagePN varchar(20))
create table #WHResult(Wid int identity(1,1),iid int,PN varchar(20),SubPN varchar(20),Usage float,Qty float,INVDate char(10))

declare @WX int
declare @WY int
declare @iX int
declare @iY int
declare @usage float
declare @mwX int
declare @mwY int
declare @WPOVendor varchar(20)
declare @WIECPN varchar(20)
declare @WOpenQty float

declare @minQty float 
declare @mX int
declare @mY int                         
declare @mQty float                      
declare @mat varchar(20)

declare @altX int
declare @altY int
declare @altQty float
declare @altINV int
declare @mainPN varchar(20)
declare @DetailQty float
declare @INVDate char(10)
declare @altMatNo varchar(20)

select @WX=min(iid) from #tmp_OPO
select @WY=max(iid) from #tmp_OPO

--select @WX=11
--select @WY=11

while @WX<=@WY  ----1st loop begin
begin --111
----Find OPO item (2012/12/21 --> needs to remove PType='ZN')
       ----(2013/01/18 add ,find PF to deduct )
       ----(2015/01/29 add ,Don't need to check if no shortage)
       ----(2015/02/11 clear the variant)
       select @WPOVendor='',@WIECPN='',@WOpenQty=0
       select @WPOVendor=isnull(a.POVendor,''),@WIECPN=isnull(IECPN,''),@WOpenQty=isnull(OpenQty,0)-isnull(DockQty,0) from #tmp_OPO a, #AAResult b
       where a.IECPN=b.PF and a.POVendor=b.POVendor and iid=@WX and PType in ('NB','ZB')
       --select @WPOVendor=isnull(POVendor,''),@WIECPN=isnull(IECPN,''),@WOpenQty=isnull(OpenQty,0)-isnull(DockQty,0) from #tmp_OPO where iid=@WX and PType in ('NB','ZB')
--select @WPOVendor,@WIECPN,@WOpenQty
---Only OpenQty>0 need to check detail ,else ,go next .
if (@WOpenQty>0 /*and left(@WIECPN,2) in ('PF','SF','JF')*/)
begin --1121          
---------------------------------------------------------------------------------------------------------------------------------------------
-----(2014/02/13)deduct PF & 60 first in Inventory exists
---------------------------------------------------------------------------------------------------------------------------------------------
	 if exists(select * from #INV where POVendor=@WPOVendor and MatNo=@WIECPN and MaterialProperty='FG' /*and left(MatNo,2) in ('PF','SF','JF')*/ and Qty>0)
	 begin --10
	 --while @WOpenQty>0
	 --		begin --12			    
				select top 1 @DetailQty=Qty,@INVDate=INVDate from #INV where POVendor=@WPOVendor and MatNo=@WIECPN and MaterialProperty='FG' /*and left(MatNo,2) in ('PF','SF','JF')*/ and Qty>0 order by INVDate
				--select @DetailQty,@INVDate,@WOpenQty
				if @WOpenQty>=@DetailQty
				begin --13
					 ----INV 變 0
					update #INV set Qty=0,Remark=rtrim(Remark)+';('+convert(varchar(10),@WX)+')-->'+convert(varchar(10),@DetailQty)
					 where MatNo=@WIECPN and INVDate=@INVDate and MaterialProperty='FG' and POVendor=@WPOVendor
					/*'+convert(varchar(10),@WX)+'*/
					insert #WHResult values(@WX,@WIECPN,@WIECPN,1,@DetailQty,@INVDate)
					select @WOpenQty=@WOpenQty-@DetailQty					
				 end  --13
				 else
				 begin  --14
					  ---INV 減少
					update #INV set Qty=@DetailQty-@WOpenQty,Remark=rtrim(Remark)+';('+convert(varchar(10),@WX)+')-->'+convert(varchar(10),@WOpenQty)
					where MatNo=@WIECPN and INVDate=@INVDate and MaterialProperty='FG' and POVendor=@WPOVendor
					/*'+convert(varchar(10),@WX)+'*/
					insert #WHResult values(@WX,@WIECPN,@WIECPN,1,@WOpenQty,@INVDate)
					select @WOpenQty=0  																	                               
				end    --14  
		--	end  --12 
	------------------------------------------------------------------------------ 
    end  --10    
    select @DetailQty=0 
end --1121
---------------------------------------------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------------------------------------------
---Only OpenQty>0 need to check detail ,else ,go next .
if @WOpenQty>0 
begin --112     
       insert #tmpWH
            select distinct @WOpenQty,RefPN,ShortagePN,Qty,0 from #AAResult where POVendor=@WPOVendor and PF=@WIECPN order by RefPN
            
   
            select @iX=min(RefPN) from #tmpWH
            select @iY=max(RefPN) from #tmpWH
            
      
            while @iX<=@iY ----2nd loop begin
            begin --1
                     insert #midWH
                          select distinct OpenQty,ShortagePN,Qty,INV,0,0 from #tmpWH where RefPN=@iX
                         
                     select @mwX=min(iid) from #midWH
                     select @mwY=max(iid) from #midWH
---Debug 
--select * from #midWH 
                 
              while @mwX<=@mwY
              begin --301             
						if exists(select * from #midWH where ShortagePN like '%*' and iid=@mwX)
						begin --2
						 update #midWH set INV=b.Qty from #midWH a,
						 (	
						    select MS,Qty=sum(Qty) from
						    (					 
						    select MS,Qty=sum(a.Qty) from #INV a,
						   (select distinct MS,ShortagePN from #Alternative where MS in
						    (select distinct ShortagePN from #midWH where ShortagePN like '%*' and iid=@mwX) and POVendor=@WPOVendor and PF=@WIECPN) 
						    as b where a.MatNo=b.ShortagePN and a.MaterialProperty='FG'
						    group by MS
						    union
						    select MS=MatNo,Qty=sum(Qty) from #INV where MatNo=(select distinct ShortagePN from #midWH where iid=@mwX) group by MatNo					    
						    ) as a group by MS

						  ) as b where a.ShortagePN=b.MS and a.iid=@mwX						  

/*-----test
						    select MS,Qty=sum(Qty) from
						    (					 
						    select MS,Qty=sum(a.Qty) from #INV a,
						   (select distinct MS,ShortagePN from #Alternative where MS in
						    (select distinct ShortagePN from #midWH where ShortagePN like '%*' and iid=@mwX) and POVendor=@WPOVendor and PF=@WIECPN) 
						    as b where a.MatNo=b.ShortagePN and a.MaterialProperty='FG'
						    group by MS
						    union
						    select MS=MatNo,Qty=sum(Qty) from #INV where MatNo=(select distinct ShortagePN from #midWH where iid=@mwX) group by MatNo					    
						    ) as a group by MS
-----*/						    


						end   --2
						else						 
						begin --3
							update #midWH set INV=b.Qty from #midWH a,(select POVendor,MatNo,Qty=sum(Qty)from #INV where MaterialProperty='FG' group by POVendor,MatNo)
							b where a.ShortagePN=b.MatNo and a.iid=@mwX --and b.MaterialProperty='FG'
 							                
						end  --3
			   select @mwX=@mwX+1			
               end --301  
                                    
                 ---Get fillQty
                     update #midWH set fillQty=convert(int,INV/Qty)            
                 
---Debug 
--select * from #midWH   
if exists(select * from #midWH where fillQty=0)
begin --20131121_1 
   ----Some Material shortage ,ignore all WHResult for this ID
   --select 'YES'
   truncate table #midWH
   delete from #tmpWH
   delete from #WHResult where iid=@WX
end --20131121_1 
else
begin  --20131121_2       
                      
                      
                        select @minQty=min(fillQty) from #midWH --找成套量 (fillQty 最小值) 
                        
                       ---Get NeedQty
                        update #midWH set NeedQty=case when @minQty*Qty>OpenQty*Qty then OpenQty*Qty else @minQty*Qty end   
                                        
---Debug 
--select * from #midWH
 
                        select @mX=min(iid) from #midWH
                        select @mY=max(iid) from #midWH
                        while @mX<=@mY
                        begin --6
                             select @mat=ShortagePN,@mQty=NeedQty,@usage=Qty from #midWH where iid=@mX
--debug
--select @mat,@minQty,@mQty
--select * from #midWH 
                             
                             if(@mat like '%*') ----If Alternative existed 
                             begin --7
                             ---Original PN need to add
                             ---(2012/12/21) need to get original PN first
                                 insert #altWH
                                 select @mQty,ShortagePN from #Alternative where POVendor=@WPOVendor and PF=@WIECPN and MS=@mat   
                             
                                 insert #altWH
                                 select @mQty,@mat from #midWH where iid=@mX

                                 
--select @mat 
--select * from #midWH                                
--select * from #altWH                              
                                 select @altX=min(iid) from #altWH
                                 select @altY=max(iid) from #altWH 
                          
                                 while @altX<=@altY ----3rd Loop begin                                 
                                 begin --9
										select @altQty=NeedQty,@altMatNo=ShortagePN from #altWH where iid=@altX
--select '@altQty'=@altQty,'@altMatNo'=@altMatNo,'@mat'=@mat
										--if exists(select * from #INV where POVendor=@WPOVendor and MatNo=@altMatNo and MaterialProperty='FG' and Qty>0)
										--begin --10
										-- (2012/11/05)select top 1  @altINV=Qty from #INV where POVendor=@WPOVendor and MatNo=@altMatNo and MaterialProperty='FG' order by INVDate
                                       ----------------------------------------------------------------------------
											while exists(select * from #INV where POVendor=@WPOVendor and MatNo=@altMatNo and MaterialProperty='FG' and Qty>0) and @altQty>0
											begin --12
												select top 1 @DetailQty=Qty,@INVDate=INVDate from #INV where POVendor=@WPOVendor and MatNo=@altMatNo and MaterialProperty='FG' and Qty>0 order by INVDate
--debug
--select '@DetailQty'=@DetailQty,@altQty
--select top 1 * from #INV where POVendor=@WPOVendor and MatNo=@altMatNo and MaterialProperty='FG' and Qty>0 order by INVDate	
												if @altQty>=@DetailQty
												 begin --13
													 ----INV 變 0
													update #INV set Qty=0,Remark=rtrim(Remark)+';('+convert(varchar(10),@WX)+')-->'+convert(varchar(10),@DetailQty)
													where MatNo=@altMatNo and INVDate=@INVDate and MaterialProperty='FG' and POVendor=@WPOVendor													
													/*'+convert(varchar(10),@WX)+'*/
													insert #WHResult values(@WX,@mat,@altMatNo,@usage,@DetailQty,@INVDate)
													 set @altQty=@altQty-@DetailQty	
													 --set @DetailQty=0									 
													 --select @altQty,@DetailQty																							 
													 update #altWH set NeedQty=@altQty --where iid=@altX
													 --select * from #altWH
													 update #midWH set NeedQty=@altQty where ShortagePN=@mat 
											     end  --13
												 else
												 begin  --14
													---INV 減少
													update #INV set Qty=@DetailQty-@altQty,Remark=rtrim(Remark)+';('+convert(varchar(10),@WX)+')-->'+convert(varchar(10),@altQty)
													where MatNo=@altMatNo and INVDate=@INVDate and MaterialProperty='FG' and POVendor=@WPOVendor
													/*'+convert(varchar(10),@WX)+'*/
													insert #WHResult values(@WX,@mat,@altMatNo,@usage,@altQty,@INVDate)																											
													set @altQty=0  
													--set @DetailQty=@DetailQty-@altQty
													--select @altQty,@DetailQty
													update #altWH set NeedQty=0--@altQty --where iid=@altX	
													update #midWH set NeedQty=0 where ShortagePN=@mat											                               
												 end    --14 					 											 
										     end  --12 
										------------------------------------------------------------------------------ 
                                        --end  --10    
                                        select @altX=@altX+1 
                                        --set @altQty=0
                                        --set @altMatNo=''
                                 end  --9 ----3rd Loop end 
                                 --update #midWH set NeedQty=(select sum(@altQty) from #altWH) where iid=@mX  
--select * from #midWH                                
--select * from #altWH                                     
                                 
                                 truncate table #altWH	
                             end --7
                             if ((@mat like '13%' and len(rtrim(@mat))=12) or len(rtrim(@mat))=12) ----(2013/02/05) ,add "and len(rtrim(@mat))=12" . If Alternative deosn't existed 
                             begin --8
                                  --select * from #midWH
                                  if exists(select * from #INV where POVendor=@WPOVendor and MatNo=@mat and MaterialProperty='FG' and Qty>0)
								  begin --15
								  while @mQty>0 --19							  
                                  begin
--Debug
--select '@mQty'=@mQty	                                 
								      select top 1 @DetailQty=Qty,@INVDate=INVDate from #INV where POVendor=@WPOVendor and MatNo=@mat and MaterialProperty='FG' and Qty>0 order by INVDate
								      --select '@DetailQty'=@DetailQty
								      ----Debug
								      --select @mQty,@mat
								      --select * from #INV where MatNo='6070B0471301'
								      if @mQty>=@DetailQty
									  begin  --17
											----INV 變 0
											update #INV set Qty=0,Remark=rtrim(Remark)+';('+convert(varchar(10),@WX)+')-->'+convert(varchar(10),@DetailQty)
											where MatNo=@mat and INVDate=@INVDate and MaterialProperty='FG'	 and POVendor=@WPOVendor										
											insert #WHResult values(@WX,@mat,@mat,@usage,@DetailQty,@INVDate)
											select @mQty=@mQty-@DetailQty
											update #midWH set NeedQty=@mQty where ShortagePN=@mat
                                               
									  end     --17
									  else
									  begin   --18
											---INV 減少
											update #INV set Qty=@DetailQty-@mQty,Remark=rtrim(Remark)+';('+convert(varchar(10),@WX)+')-->'+convert(varchar(10),@mQty)
											where MatNo=@mat and INVDate=@INVDate and MaterialProperty='FG' and POVendor=@WPOVendor
											/*'+convert(varchar(10),@WX)+'*/
											insert #WHResult values(@WX,@mat,@mat,@usage,@mQty,@INVDate)
											select @mQty=0   
											update #midWH set NeedQty=@mQty where ShortagePN=@mat                               
									  end     --18
								   end --19								      
								  end  --15	
							 end --8  
						select @mX=@mX+1                        					
                        end --6                          

                 select @iX=@iX+1
                 --update #midWH set NeedQty=@mQty where iid=@mX 
                 --select * from #midWH
                 --delete from #midWH           
            
            --delete from #tmpWH 
            truncate table #midWH
            --select * from #midWH
            --select * from #tmpWH  
end  --20131121_2  
end  --1----2nd loop end        
     select @WX=@WX+1  
     delete from #tmpWH 
--------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------
end   --112
else
begin  --113
    select @WX=@WX+1
end    --113
end   --111 ----1st loop end

---------Get ETA
/*
drop table #tmpETA
drop table #WHmid
drop table #WHPN
drop table #midFinal
drop table #WHFinal
drop table #WHcount
drop table #ETAQty
*/
create table #tmpETA(iid int,PN varchar(20),Qty float,INVDate char(10))
create table #WHmid(PN varchar(20),Qty float,INVDate char(10))
create table #midFinal(PN varchar(20),Qty float,INVDate char(10))
create table #WHPN(pid int identity(1,1),PN varchar(20))
create table #WHFinal(iid int,Qty float,INVDate char(10))
create table #WHcount(Wid int identity,iid int,DateQty varchar(50))


insert #WHFinal
select distinct iid,qty=Qty/Usage,ETA=max(INVDate) from #WHResult where not iid in (
select iid from (
select iid,PN,qty=count(*) from #WHResult group by iid,PN) as a where qty>1) group by iid,Qty/Usage order by iid

insert #tmpETA
select distinct iid,PN,qty=Qty/Usage,INVDate from #WHResult where not iid in (select distinct iid from #WHFinal)

declare @mi int
declare @mj int
declare @PN varchar(20)
declare @Qty float
declare @pnINVDate char(10)
declare @pni int
declare @pnj int
declare @mPN varchar(20)
declare @midPN varchar(20)
declare @midQty float
declare @midINVDate char(10)

select @mi=min(iid) from #tmpETA
select @mj=max(iid) from #tmpETA

while @mi<=@mj
begin --111
     insert #WHmid
        select PN,Qty,INVDate from #tmpETA where iid=@mi order by Qty
     while exists(select * from #WHmid)
     begin --112     
		select top 1 @PN=PN,@Qty=Qty,@pnINVDate=INVDate from #WHmid where not Qty=0  order by Qty,INVDate  
		--select @PN,@Qty,@INVDate      
		insert #midFinal values(@PN,@Qty,@pnINVDate)       
        
		delete from #WHmid where PN=@PN and Qty=@Qty and INVDate=@pnINVDate
        
		insert #WHPN
			select distinct PN from #WHmid where not PN=@PN 
        
		select @pni=min(pid) from #WHPN
		select @pnj=max(pid) from #WHPN
     
		while @pni<=@pnj
		begin --113
			select @mPN=PN from #WHPN where pid=@pni
         

			select top 1 @midPN=PN,@midQty=Qty,@midINVDate=INVDate from #WHmid where PN=@mPN order by Qty
       		
       			            
			update #WHmid set Qty=@midQty-@Qty where PN=@midPN and Qty=@midQty and INVDate=@midINVDate
			insert #midFinal values(@midPN,@Qty,@midINVDate) 
            --select * from #WHmid
			delete from #WHmid where Qty=0            
			select @pni=@pni+1
		end --113		            
      truncate table #WHPN  
      
      insert #WHFinal 
          select @mi,Qty,Max(INVDate) from #midFinal group by Qty      
      delete #midFinal
      end --112      
      delete from #WHmid
      select @mi=@mi+1
 end--111  

--select * from #WHFinal where Qty>0
insert #WHcount
   select iid,INVDate+'*'+convert(varchar(10),Qty) from #WHFinal where INVDate=convert(char(10),getdate(),111) and Qty>=1--(?? need to check) Qty>0
 
--Future Order add 4 more days .
insert #WHcount
   select iid,
   convert(char(10),dateadd(dd,4,convert(datetime,INVDate+' 00:00')),111)
   +'*'+convert(varchar(10),Qty) from #WHFinal where not INVDate=convert(char(10),getdate(),111)
 

--Drop table #ETAQty
create table #ETAQty(iid int,OPOR varchar(200))

insert #ETAQty
    select distinct iid,'' from #WHcount order by iid
    
declare @la int
declare @lb int
select @la=min(Wid) from #WHcount
select @lb=max(Wid) from #WHcount

while @la<=@lb
begin
    update #ETAQty set OPOR=rtrim(a.OPOR)+rtrim(b.DateQty)+' , ' from #ETAQty a,#WHcount b where a.iid=b.iid and b.Wid=@la
    select @la=@la+1
end

update #ETAQty set OPOR=left(OPOR,len(OPOR)-1)


------(2013/07/24) Not Shortage Flag Logic
----Find Current INV existed .
select a.iid,a.OpenQty,b.WHQty,RR=a.OpenQty-b.WHQty into #SHResult from #tmp_OPO a,
(select iid,WHQty=sum(Qty) from #WHFinal where INVDate<=convert(char(10),getdate(),111) group by iid) as b where a.iid=b.iid

update #tmp_OPO set Shortage='N' from #tmp_OPO a,#SHResult b where a.iid=b.iid and RR=0
update #tmp_OPO set Shortage='Y' from #tmp_OPO a,#SHResult b where a.iid=b.iid and RR>0
update #tmp_OPO set Shortage='Y' from  #tmp_OPO a,#AAResult b where a.POVendor=b.POVendor and a.IECPN=b.PF and a.Shortage=''

update #tmp_OPO set Shortage='N' where Shortage=''

----2013/08/22 ---Correct OpenQty calculation rule
update #tmp_OPO set OpenQty=POQty-DockQty-ShipQty where PType='ZN'
update #tmp_OPO set Shortage='N' where PType='ZN' and OpenQty='0'


/*
Drop table #ETAfinal
*/

create table #ETAfinal(iid int,OpenQty int,CurrentDone int,FutureDone int,FirstSupportDate char(10),LastSupportDate char(10))
insert #ETAfinal
select a.iid,OpenQty,CurrentDone/*=case when FD is null then CurrentDone else OpenQty-FD end*/
,FutureDone=isnull(FD,'') ,FristSupportDate=isnull(MD,''),LastSupportDate=isnull(LD,'') from
(select a.iid,OpenQty=OpenQty-DockQty,
CurrentDone=case when Shortage='N' then OpenQty-DockQty else isnull(CD,'') end from
(select * from #tmp_OPO /*where PType in ('NB','ZB')*/) as a left join
(select iid,CD=sum(Qty) from #WHFinal where INVDate<=convert(char(10),getdate(),111) group by iid) as b on a.iid=b.iid 
) as a left join
(select iid,MD=max(INVDate),LD=max(INVDate),FD=sum(Qty) from #WHFinal where INVDate>convert(char(10),getdate(),111) group by iid) as b on a.iid=b.iid order by a.iid

--insert #ETAfinal
--    select iid,OpenQty,CurrentDone=0,FutureDone=0,FristSupportDate='',LastSupportDate='' from #tmp_OPO where PType='ZN'



---------------------------------------------------New Report -----------------------------------------------------
---------------------------------------------------New Report -----------------------------------------------------
---------------------------------------------------New Report -----------------------------------------------------
---------------------------------------------------New Report -----------------------------------------------------
---------------------------------------------------New Report -----------------------------------------------------
---------------------------------------------------New Report -----------------------------------------------------
---------------------------------------------------New Report -----------------------------------------------------
---------------------------------------------------New Report -----------------------------------------------------
---------------------------------------------------New Report -----------------------------------------------------
---------------------------------------------------New Report -----------------------------------------------------
--select * from #tmp_OPO where SO='1105000723'
--select * from #ETAfinal

delete from OPS_OPO where Customer='DYNABOOK' and (ReportDate=convert(char(10),getdate(),111) or ReportDate<=convert(char(10),dateadd(dd,-90,getdate()),111))
delete from OPS_OPOR1 where Customer='DYNABOOK' and (ReportDate=convert(char(10),getdate(),111) or ReportDate<=convert(char(10),dateadd(dd,-90,getdate()),111))
delete from OPS_OPOR2 where Customer='DYNABOOK' and (ReportDate=convert(char(10),getdate(),111) or ReportDate<=convert(char(10),dateadd(dd,-90,getdate()),111))
delete from OPS_OPOSummary where Customer='DYNABOOK' and (ReportDate=convert(char(10),getdate(),111) or ReportDate<=convert(char(10),dateadd(dd,-90,getdate()),111))
delete from OPS_OPOSummary2 where Customer='DYNABOOK' and ReportDate=convert(char(10),getdate(),111)---Do not Clear historical data .
delete from OPS_OM where Customer='DYNABOOK' and (ReportDate=convert(char(10),getdate(),111) or ReportDate<=convert(char(10),dateadd(dd,-90,getdate()),111))
delete from OPS_Material where Customer='DYNABOOK' and (ReportDate=convert(char(10),getdate(),111) or ReportDate<=convert(char(10),dateadd(dd,-90,getdate()),111))
delete from OPS_MSummaryPIC where Customer='DYNABOOK' and (ReportDate=convert(char(10),getdate(),111) or ReportDate<=convert(char(10),dateadd(dd,-90,getdate()),111))
delete from OPS_MSummaryType where Customer='DYNABOOK' and (ReportDate=convert(char(10),getdate(),111) or ReportDate<=convert(char(10),dateadd(dd,-90,getdate()),111))
delete from OPS_Alt where Customer='DYNABOOK' and (ReportDate=convert(char(10),getdate(),111) or ReportDate<=convert(char(10),dateadd(dd,-90,getdate()),111))
delete from OPS_Inventory where Customer='DYNABOOK' and (ReportDate=convert(char(10),getdate(),111) or ReportDate<=convert(char(10),dateadd(dd,-90,getdate()),111))
delete from OPS_RawOPO where Customer='DYNABOOK' and (ReportDate=convert(char(10),getdate(),111) or ReportDate<=convert(char(10),dateadd(dd,-90,getdate()),111))
delete from OPS_OTDFailD where Customer='DYNABOOK' and (ReportDate=convert(char(10),getdate(),111) or ReportDate<=convert(char(10),dateadd(dd,-90,getdate()),111))
delete from OPS_OTDFailS where Customer='DYNABOOK' and (ReportDate=convert(char(10),getdate(),111) or ReportDate<=convert(char(10),dateadd(dd,-90,getdate()),111))
/*
select * from OPS_OPO where Customer='DYNABOOK'
select * from OPS_OPOR1 where Customer='DYNABOOK'
select * from OPS_OPOR2 where Customer='DYNABOOK'
select * from OPS_OPOSummary where Customer='DYNABOOK'
select * from OPS_OPOSummary2 where Customer='DYNABOOK'---Do not Clear data .
select * from OPS_OM where Customer='DYNABOOK'
select * from OPS_Material where Customer='DYNABOOK'
select * from  OPS_MSummaryPIC where Customer='DYNABOOK'
select * from  OPS_MSummaryType where Customer='DYNABOOK'
select * from  OPS_Alt where Customer='DYNABOOK'
select * from  OPS_Inventory where Customer='DYNABOOK'
select * from  OPS_RawOPO where Customer='DYNABOOK'
select * from  OPS_OTDFailD where Customer='DYNABOOK'
select * from  OPS_OTDFailS where Customer='DYNABOOK'
*/

---------------------------------Report -------------------------------
-----Get OPO Report
----OPO SDetail
--drop table #OPODetail
select a.*,OPOR=isnull(OPOR,''),RMA_FG=0 into #OPODetail from
(
select a.*,SameMonthFCST=isnull(b.M1,0) from
(
select a.iid,a.Site,a.PIC,a.PO,a.SO,a.IECPO,a.POItem,a.CPQNo,a.IECPN,ProductFamily,
POVendor,POReceiveDate,PO_Type,Model_Status,MP,FCST_Status,PType,NeedShipDate,POQty,DNQty,
DockQty,ShipQty,b.OpenQty,b.CurrentDone,b.FutureDone,FutureFirstSupportDate=b.FirstSupportDate,FutureLastSupportDate=b.LastSupportDate,
Shortage=case when POQty-DockQty-ShipQty-CurrentDone>0 then 'Y' else 'N' end,Escalation,EscalationDate,a.Remark  from #tmp_OPO a,#ETAfinal b where a.iid=b.iid
) as a left join
(select POVendor=case when POVendor='CP81' then 'IES' when POVendor='CP60' then 'ICC' end,IECPN,M1=sum(M1) from FCST 
where FCSTDate=convert(char(7),dateadd(mm,-1,getdate()),111)+'/01' group by POVendor,IECPN) as b
on a.POVendor=b.POVendor and a.IECPN=b.IECPN 
) as a left join
(select * from #ETAQty) as b on a.iid=b.iid order by iid

----Remove OPOR if PType='ZN'
update #OPODetail set OPOR='' where PType='ZN'
 
update #OPODetail set OPOR='' where Shortage='N' and OpenQty-CurrentDone=0 
--update #OPODetail set OPOR=NeedShipDate+'*'+convert(varchar(10),CurrentDone) where OpenQty=CurrentDone --and OPOR='' 
update #OPODetail set OPOR='Wait for Pick up' where OpenQty=0 and DockQty>0 and OPOR=''
update #OPODetail set OPOR='Checking' where Shortage='Y' and OPOR=''
update #OPODetail set OPOR='DN Created' where Shortage='N' and OpenQty=DNQty and OPOR=''
update #OPODetail set OPOR='Need to Create DN ASAP' where Shortage='N' and OpenQty=CurrentDone 
and datediff(dd,getdate(),convert(datetime,NeedShipDate+' 00:00'))<=7 
and datediff(dd,getdate(),convert(datetime,NeedShipDate+' 00:00'))>=0 and OPOR=''

update #OPODetail set OPOR='Please arrange Shipment ASAP' where CurrentDone>0 and NeedShipDate<=convert(char(10),getdate(),111) and OPOR=''

update #OPODetail set OPOR='(Potential OTD Fail!!) '+rtrim(OPOR) where FutureFirstSupportDate>NeedShipDate and NeedShipDate>=convert(char(10),getdate(),111)

update #OPODetail set OPOR='(OTD Fail!!) '+rtrim(OPOR) where OpenQty>0 and NeedShipDate<=convert(char(10),getdate(),111) and not OPOR like '%OTD Fail!!%'

update #OPODetail set OPOR='(Potential OTD Fail!!) '+rtrim(OPOR) where OpenQty>0
and NeedShipDate>convert(char(10),getdate(),111) and NeedShipDate<=convert(char(10),dateadd(mm,1,getdate()),111) 
and not OPOR like '%OTD Fail!!%' and OPOR like 'Checking%'

update #OPODetail set OPOR='(Potential OTD Fail!!) '+rtrim(OPOR) where OpenQty>0
and NeedShipDate>convert(char(10),getdate(),111) and convert(char(10),dateadd(dd,4,convert(datetime,FutureFirstSupportDate)),111)>=NeedShipDate 
and not OPOR like '%OTD Fail!!%' and Shortage='Y'

update #OPODetail set OPOR='(Potential OTD Fail!!) '+rtrim(OPOR) where OpenQty>0
and NeedShipDate>convert(char(10),getdate(),111) and NeedShipDate<=convert(char(10),dateadd(mm,1,getdate()),111) 
and not OPOR like '%OTD Fail!!%' and Shortage='Y'


--(2013/01/28) Correct Shortage
update #OPODetail set Shortage='Y' where FutureDone>0



---RMA FG Stock
update #OPODetail set RMA_FG=b.Qty from #OPODetail a,#INV b where a.POVendor=b.POVendor and a.IECPN=b.MatNo and b.MaterialProperty='RMA_FG'


--------(2013/02/01) Find PF Substitution------------------------------
--drop table #Sub
--drop table #midSub
create table #Sub(iid int identity(1,1),CPQNo varchar(20),SubPN varchar(200))
create table #midSub(iid int identity(1,1),SubPN varchar(20))

insert #Sub 
     select distinct IECPN,'' from #OPODetail a,SubFRU b where a.IECPN=b.CPQNo and a.Shortage='Y' and not ECRStatus='CANCEL'
   
declare @x int
declare @y int
--declare @i int
--declare @j int
declare @pn varchar(20)
select @x=MIN(iid)from #Sub
select @y=Max(iid)from #Sub


while @x<=@y
  begin
      insert #midSub
          select distinct SubPN from SubFRU where CPQNo=(select rtrim(CPQNo) from #Sub where iid=@x) and not ECRStatus='CANCEL'
      
      select @i=MIN(iid)from #midSub
      select @j=Max(iid)from #midSub
      while @i<=@j
      begin
           select @pn=rtrim(SubPN) from #midSub where iid=@i
           update #Sub set SubPN=RTRIM(SubPN)+@pn+';' where iid=@x
           select @i=@i+1
      end
      truncate table #midSub
      select @x=@x+1 
  end

update #Sub set SubPN=left(SubPN,len(rtrim(SubPN))-1) where len(SubPN)>0

update #OPODetail set OPOR=rtrim(OPOR)+' ,***Substitution : '+rtrim(b.SubPN)+'***' from #OPODetail a,#Sub b where a.IECPN=b.CPQNo
-------------------------------
---OPO Detail
insert OPS_OPO
select Customer='DYNABOOK',ReportDate=convert(char(10),getdate(),111),iid,Site,PIC,PO,SO,IECPO,POItem,CPQNo,IECPN,ProductFamily,
POVendor,POReceiveDate,PO_Type,Model_Status,MP,FCST_Status,PType,NeedShipDate,POQty,DNQty,
DockQty,ShipQty,OpenQty,RMA_FG,CurrentDone,FutureDone,FutureFirstSupportDate,--FutureLastSupportDate,
Shortage,SameMonthFCST,OPOR,Escalation,EscalationDate,Remark from #OPODetail order by PIC,Site,NeedShipDate

------OPOR Status--By Site
insert OPS_OPOR1
select Customer='DYNABOOK',ReportDate=convert(char(10),getdate(),111),Site,PIC,OPOR,NPI_ItemQty=sum(NPI_ItemQty),Normal_ItemQty=sum(Normal_ItemQty) from 
(
select Site,PIC,OPOR,NPI_ItemQty=sum(qty),Normal_ItemQty=0 from (
select Site,PIC,
OPOR=case 
when (OPOR like '%Potential%' and OPOR like '%*%') then 'OPOR applied (Potential OTD Failed)' 
when (OPOR like '%Potential%' and OPOR like '%Checking%') then 'Potential OTD Failed' 
when OPOR like '%Fail%' then 'OPOR applied (OTD Failed)'
when OPOR like '%*%' then 'OPOR Applied' 
when OPOR ='' then 'Under Control' 
else OPOR end,qty=count(*) from #OPODetail where PO_Type='NPI' group by Site,PIC,OPOR
) as a group by Site,PIC,OPOR 
union
select Site,PIC,OPOR,0,sum(qty) from (
select Site,PIC,
OPOR=case 
when (OPOR like '%Potential%' and OPOR like '%*%')  then 'OPOR applied (Potential OTD Failed)' 
when (OPOR like '%Potential%' and OPOR like '%Checking%') then 'Potential OTD Failed' 
when OPOR like '%Fail%' then 'OPOR applied (OTD Failed)'
when OPOR like '%*%' then 'OPOR Applied' 
when OPOR ='' then 'Under Control' 
else OPOR end,qty=count(*) from #OPODetail where not PO_Type='NPI' group by Site,PIC,OPOR
) as a  group by Site,PIC,OPOR 
) as a group by Site,PIC,OPOR order by PIC,Site,OPOR


------OPOR Status --By OPOR
insert OPS_OPOR2
select Customer='DYNABOOK',ReportDate=convert(char(10),getdate(),111),OPOR,NPI_ItemQty=sum(NPI_ItemQty),Normal_ItemQty=sum(Normal_ItemQty) from 
(
select Site,PIC,OPOR,NPI_ItemQty=sum(qty),Normal_ItemQty=0 from (
select Site,PIC,
OPOR=case 
when (OPOR like '%Potential%' and OPOR like '%*%') then 'OPOR applied (Potential OTD Failed)' 
when (OPOR like '%Potential%' and OPOR like '%Checking%') then 'Potential OTD Failed' 
when OPOR like '%Fail%' then 'OPOR applied (OTD Failed)'
when OPOR like '%*%' then 'OPOR Applied' 
when OPOR ='' then 'Under Control' 
else OPOR end,qty=count(*) from #OPODetail where PO_Type='NPI' group by Site,PIC,OPOR
) as a group by Site,PIC,OPOR 
union
select Site,PIC,OPOR,0,sum(qty) from (
select Site,PIC,
OPOR=case 
when (OPOR like '%Potential%' and OPOR like '%*%') then 'OPOR applied (Potential OTD Failed)' 
when (OPOR like '%Potential%' and OPOR like '%Checking%') then 'Potential OTD Failed' 
when OPOR like '%Fail%' then 'OPOR applied (OTD Failed)'
when OPOR like '%*%' then 'OPOR Applied' 
when OPOR ='' then 'Under Control' 
else OPOR end,qty=count(*) from #OPODetail where not PO_Type='NPI' group by Site,PIC,OPOR
) as a  group by Site,PIC,OPOR 
) as a group by OPOR order by OPOR


----OPO Summary
--Drop table #OPOSummary
--OTD fail
select Dt=convert(char(10),getdate(),111),PType,Site,PIC,allItem=sum(allItem),allqty=sum(allqty),newItem=sum(newItem),newQty=sum(newQty),/*,InFCST=sum(InFCST),Upside=sum(Upside)*/
UpsideRate=convert(decimal(8,2),convert(float,sum(Upside))/convert(float,sum(allqty))*100),nearfailItem=sum(nearfailItem),
nearfailqty=sum(nearfailqty),failItem=sum(failItem),failqty=sum(failqty),failpercentage=convert(decimal(8,2),convert(float,sum(failqty))/
case when convert(float,sum(allqty))=0 then 1 else convert(float,sum(allqty)) end *100),
NPIallItem=sum(NPIallItem),NPIallqty=sum(NPIallqty),
NPInewItem=sum(NPInewItem),NPInewQty=sum(NPInewQty),
NPInearfailItem=sum(NPInearfailItem),NPInearfailqty=sum(NPInearfailqty),
NPIfailItem=sum(NPIfailItem),NPIfailqty=sum(NPIfailqty),
NPIfailpercentage=convert(decimal(8,2),convert(float,sum(NPIfailqty))/
case when convert(float,sum(NPIallqty))=0 then 1 else convert(float,sum(NPIallqty)) end *100) into #OPOSummary from
(
select Site,PIC,PType=left(IECPN,1),allItem=count(*),newItem=0,newQty=0,allqty=0,InFCST=0,Upside=0,nearfailItem=0,nearfailqty=0,failItem=0,failqty=0,
NPIallItem=0,NPIallqty=0,NPInewItem=0,NPInewQty=0,NPInearfailItem=0,NPInearfailqty=0,NPIfailItem=0,NPIfailqty=0
from #OPODetail where not PO_Type in ('NPI','NPI2') and not OpenQty=0 group by Site,PIC,left(IECPN,1)
union
select Site,PIC,PType=left(IECPN,1),0,count(*),0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 from #OPODetail where not PO_Type in ('NPI','NPI2') and POReceiveDate=convert(char(10),dateadd(dd,-1,getdate()),111) and not OpenQty=0 group by Site,PIC,left(IECPN,1) 
union
select Site,PIC,PType=left(IECPN,1),0,0,sum(OpenQty),0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 from #OPODetail where not PO_Type in ('NPI','NPI2') and POReceiveDate=convert(char(10),dateadd(dd,-1,getdate()),111) and not OpenQty=0 group by Site,PIC,left(IECPN,1) 
union
select Site,PIC,PType=left(IECPN,1),0,0,0,sum(OpenQty),0,0,0,0,0,0,0,0,0,0,0,0,0,0 from #OPODetail where not PO_Type in ('NPI','NPI2') and not OpenQty=0 group by Site,PIC,left(IECPN,1)
union
select Site,PIC,PType=left(IECPN,1),0,0,0,0,sum(OpenQty),0,0,0,0,0,0,0,0,0,0,0,0,0 from #OPODetail where not PO_Type in ('NPI','NPI2') and FCST_Status='IN'  and not OpenQty=0 group by Site,PIC,left(IECPN,1)
union
select Site,PIC,PType=left(IECPN,1),0,0,0,0,0,sum(OpenQty),0,0,0,0,0,0,0,0,0,0,0,0 from #OPODetail where not PO_Type in ('NPI','NPI2') and FCST_Status in ('UPSID','NO') and not OpenQty=0 group by Site,PIC,left(IECPN,1)
union
select Site,PIC,PType=left(IECPN,1),0,0,0,0,0,0,count(*),0,0,0,0,0,0,0,0,0,0,0 from #OPODetail where not PO_Type in ('NPI','NPI2') and  Shortage='Y' and NeedShipDate>=convert(char(10),getdate(),111) 
and  NeedShipDate<convert(char(10),dateadd(dd,7,getdate()),111) and not OpenQty=0 group by Site,PIC,left(IECPN,1)
union
select Site,PIC,PType=left(IECPN,1),0,0,0,0,0,0,0,sum(OpenQty),0,0,0,0,0,0,0,0,0,0 from #OPODetail where  not PO_Type in ('NPI','NPI2') and Shortage='Y' and NeedShipDate>=convert(char(10),getdate(),111) 
and  NeedShipDate<convert(char(10),dateadd(dd,7,getdate()),111) and not OpenQty=0 group by Site,PIC,left(IECPN,1)
union
select Site,PIC,PType=left(IECPN,1),0,0,0,0,0,0,0,0,count(*),0,0,0,0,0,0,0,0,0 from #OPODetail where  not PO_Type in ('NPI','NPI2') and NeedShipDate<convert(char(10),getdate(),111) and not OpenQty=0 group by Site,PIC,left(IECPN,1)
union
select Site,PIC,PType=left(IECPN,1),0,0,0,0,0,0,0,0,0,sum(OpenQty),0,0,0,0,0,0,0,0 from #OPODetail where  not PO_Type in ('NPI','NPI2') and NeedShipDate<convert(char(10),getdate(),111) and not OpenQty=0 group by Site,PIC,left(IECPN,1)

union
select Site,PIC,PType=left(IECPN,1),0,0,0,0,0,0,0,0,0,0,count(*),0,0,0,0,0,0,0 from #OPODetail where PO_Type in ('NPI','NPI2') and not OpenQty=0 group by Site,PIC,left(IECPN,1)
union
select Site,PIC,PType=left(IECPN,1),0,0,0,0,0,0,0,0,0,0,0,sum(OpenQty),0,0,0,0,0,0 from #OPODetail where  PO_Type in ('NPI','NPI2') and not OpenQty=0 group by Site,PIC,left(IECPN,1)
union
select Site,PIC,PType=left(IECPN,1),0,0,0,0,0,0,0,0,0,0,0,0,count(*),0,0,0,0,0 from #OPODetail where PO_Type in ('NPI','NPI2') and POReceiveDate=convert(char(10),dateadd(dd,-1,getdate()),111) and not OpenQty=0 group by Site,PIC,left(IECPN,1) 
union
select Site,PIC,PType=left(IECPN,1),0,0,0,0,0,0,0,0,0,0,0,0,0,sum(OpenQty),0,0,0,0 from #OPODetail where PO_Type in ('NPI','NPI2') and POReceiveDate=convert(char(10),dateadd(dd,-1,getdate()),111) and not OpenQty=0 group by Site,PIC,left(IECPN,1) 
union
select Site,PIC,PType=left(IECPN,1),0,0,0,0,0,0,0,0,0,0,0,0,0,0,count(*),0,0,0 from #OPODetail where PO_Type in ('NPI','NPI2') and  Shortage='Y' and NeedShipDate>=convert(char(10),getdate(),111) 
and  NeedShipDate<convert(char(10),dateadd(dd,7,getdate()),111) and not OpenQty=0 group by Site,PIC,left(IECPN,1)
union
select Site,PIC,PType=left(IECPN,1),0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,sum(OpenQty),0,0 from #OPODetail where PO_Type in ('NPI','NPI2') and Shortage='Y' and NeedShipDate>=convert(char(10),getdate(),111) 
and  NeedShipDate<convert(char(10),dateadd(dd,7,getdate()),111) and not OpenQty=0 group by Site,PIC,left(IECPN,1)
union
select Site,PIC,PType=left(IECPN,1),0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,count(*),0 from #OPODetail where PO_Type in ('NPI','NPI2') and NeedShipDate<convert(char(10),getdate(),111) and not OpenQty=0 group by Site,PIC,left(IECPN,1)
union
select Site,PIC,PType=left(IECPN,1),0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,sum(OpenQty) from #OPODetail where PO_Type in ('NPI','NPI2') and NeedShipDate<convert(char(10),getdate(),111) and not OpenQty=0 group by Site,PIC,left(IECPN,1)
) as a --where not Site='ARVATO'
group by Site,PIC,PType order by PType,failqty desc


insert OPS_OPOSummary
select Customer='DYNABOOK',* from #OPOSummary

--select from OPS_OPOSummary2

insert OPS_OPOSummary2
select * from (
select Customer,ReportDate,PType=case when PType='6' then 'Raw Material' else 'FRU' end,AllItem=sum(allItem),AllQty=sum(allqty)
,NewItem=sum(newItem),NewQty=sum(newQty),NearFailItem=sum(nearfailItem),NearFailqty=sum(nearfailqty),
FailItem=sum(failItem),Failqty=sum(failqty),FailPercentage=convert(decimal(8,2),convert(float,sum(failqty))/
case when convert(float,sum(allqty))=0 then 1 else convert(float,sum(allqty)) end *100),
NPIAllItem=sum(NPIallItem),NPIAllQty=sum(NPIallqty),
NPINewItem=sum(NPInewItem),NPINewQty=sum(NPInewQty),NPINearFailItem=sum(NPInearfailItem),NPINearFailqty=sum(NPInearfailqty),
NPIFailItem=sum(NPIfailItem),NPIFailqty=sum(NPIfailqty),NPIFailPercentage=convert(decimal(8,2),convert(float,sum(NPIfailqty))/
case when convert(float,sum(NPIallqty))=0 then 1 else convert(float,sum(NPIallqty)) end *100)
from OPS_OPOSummary where Customer='DYNABOOK' and ReportDate=convert(char(10),getdate(),111) and PType='6' group by Customer,ReportDate,PType 
union
select Customer,ReportDate,PType,AllItem=sum(AllItem),AllQty=sum(allqty),
NewItem=sum(NewItem),NewQty=sum(NewQty),NearFailItem=sum(nearfailItem),sum(nearfailqty),
FailItem=sum(failItem),sum(failqty),Failpercentage=convert(decimal(8,2),convert(float,sum(failqty))/
case when convert(float,sum(allqty))=0 then 1 else convert(float,sum(allqty)) end *100),
NPIAllItem=sum(NPIAllItem),NPIAllQty=sum(NPIAllQty),
NPINewItem=sum(NPINewItem),NPINewQty=sum(NPINewQty),NPINearFailItem=sum(NPINearFailItem),NPINearFailqty=sum(NPINearFailqty),
NPIFailItem=sum(NPIFailItem),NPIFailqty=sum(NPIFailqty),NPIFailPercentage=convert(decimal(8,2),convert(float,sum(NPIFailqty))/
case when convert(float,sum(NPIAllQty))=0 then 1 else convert(float,sum(NPIAllQty)) end *100)
from (
select Customer,ReportDate,PType=case when PType='6' then 'Raw Material' else 'FRU' end,AllItem=sum(allItem),allqty=sum(allqty),NewItem=sum(newItem),NewQty=sum(newQty),nearfailItem=sum(nearfailItem),nearfailqty=sum(nearfailqty),failItem=sum(failItem),failqty=sum(failqty),
NPIAllItem=sum(NPIallItem),NPIAllQty=sum(NPIallqty),
NPINewItem=sum(NPInewItem),NPINewQty=sum(NPInewQty),NPINearFailItem=sum(NPInearfailItem),NPINearFailqty=sum(NPInearfailqty),
NPIFailItem=sum(NPIfailItem),NPIFailqty=sum(NPIfailqty)
from OPS_OPOSummary where Customer='DYNABOOK' and ReportDate=convert(char(10),getdate(),111) and not PType='6' group by Customer,ReportDate,PType 
) as a group by Customer,ReportDate,PType 
) as a 

--drop table #OM
--OPO & Shortage Material
select distinct a.*,MaterialETA=isnull(b.MaterialETA,''),Remark=isnull(b.Remark,'') into #OM from
(
select a.iid,a.Site,a.IECPN,a.NeedShipDate,a.POVendor,a.OpenQty,a.ShortagePN,Material_descript,a.NeedQty
,CurrentStock=isnull(b.CurrentQty,''),FutureStock=isnull(b.FutureQty,''),PIC from
(
select a.*,b.ShortagePN,NeedQty=OpenQty*b.Qty,Material_descript='( '+isnull(b.Material_group,'')+' )    '+isnull(b.Material_descript,''),PIC from
(select a.iid,b.Site,b.IECPN,b.POVendor,b.NeedShipDate,OpenQty=a.OpenQty-a.CurrentDone from #ETAfinal a,#tmp_OPO b where a.iid=b.iid and a.OpenQty-a.CurrentDone>0 and PType in ('NB','ZB')) as a left join
(
select a.*,PIC=case when a.ShortagePN like '13%' then 'Joan' when a.ShortagePN like '146%' then 'Joan' when a.ShortagePN like '15%' then 'Iris' else b.PIC end from 
(select a.*,b.Material_descript,b.Material_group from #AAResult a,t_download_matmas_CP62DW b where left(a.ShortagePN,12)=b.Material) as a left join
(select * from MType) as b on a.Material_group=b.Material_group

) as b on a.POVendor=b.POVendor and a.IECPN=b.PF
) as a left join
(

-----2
select POVendor,MatNo,CurrentQty=sum(CurrentQty),FutureQty=sum(FutureQty) from(
----1
select POVendor,MatNo=MatNo,CurrentQty=sum(Qty),FutureQty=0 from #INV where INVDate=convert(char(10),getdate(),111) group by POVendor,MatNo
union
select POVendor,MatNo=MatNo,0,sum(Qty) from #INV where INVDate<>convert(char(10),getdate(),111) group by POVendor,MatNo
union
-----Need to include Alternative Stock
select POVendor,MS,sum(Qty),0 from(
select a.*,Qty=isnull(Qty,0) from
(select distinct POVendor,MS,ShortagePN from #Alternative) as a left join
(select * from #INV where  INVDate=convert(char(10),getdate(),111)) as b on a.ShortagePN=b.MatNo and a.POVendor=b.POVendor
) as a group by POVendor,MS
union
select POVendor,MS,0,sum(Qty) from(
select a.*,Qty=isnull(Qty,0) from
(select distinct POVendor,MS,ShortagePN from #Alternative) as a left join
(select * from #INV where  INVDate<>convert(char(10),getdate(),111)) as b on a.ShortagePN=b.MatNo and a.POVendor=b.POVendor
) as a group by POVendor,MS
-----1
) as a group by POVendor,MatNo 
-----2

) as b on a.ShortagePN=b.MatNo and a.POVendor=b.POVendor 
) as a left join
(select * from Ivan_ShortageMaterialUpdate where convert(char(10),Cdt,111)=convert(char(10),getdate(),111) and Customer='TSB') as b on a.POVendor=b.POVendor and a.ShortagePN=b.Material
order by iid

insert OPS_OM
select Customer='DYNABOOK',ReportDate=convert(char(10),getdate(),111),* from #OM


----Check Date------------By Shortage PN
/*
drop table #Gating
drop table #KBGating
drop table #DD
drop table #MV
drop table #VD
drop table #ShortageDetail
drop table #MDCount
drop table #SMCount
drop table #MDCountTmp
drop table #ShortageDetail_total
drop table #WBinit
drop table #WBresult
drop table #WBmid
*/
------Get Gatting Report
--drop table #Gating
--drop table #KBGating
/*
select distinct POVendor,Material,Buyer,OpenPO,FCSTDemand,FCSTBalance,POBalance  into #Gating from Ivan_Gatting where DocWeek=(select max(DocWeek) from Ivan_Gatting) 

----Get KB Material and transfer to major material
--drop table #KBGating
select POVendor,b.MS,Buyer,OpenPO=sum(OpenPO),FCSTDemand=sum(FCSTDemand),FCSTBalance=sum(FCSTBalance),POBalance=sum(POBalance) into #KBGating from #Gating a,
(select distinct MS,Material from KBAlternative) as  b where a.Material=b.Material group by POVendor,b.MS,Buyer

----update Major KB's OPO
update #KBGating set OpenPO=b.Open_Qty from #KBGating a,
(
select POVendor='IES',b.MS,Open_Qty=sum(convert(float,Open_Qty)) from ipc_t_download_nb_po a,KBAlternative b 
where a.Material=b.Material and S_loc='SW03' group by b.MS
union
select POVendor='ICC',b.MS,Open_Qty=sum(convert(float,Open_Qty)) from icc_t_download_nb_po a,KBAlternative b 
where a.Material=b.Material and S_loc='SW04' group by b.MS
) as b where a.POVendor=b.POVendor and a.MS=b.MS

update #KBGating set POBalance=OpenPO+FCSTBalance

---Remove KB Material
delete from #Gating from #Gating a,KBAlternative b where a.Material=b.Material


----update All OPO
update #Gating set OpenPO=b.Open_Qty from #Gating a,
(
select POVendor='IES',Material,Open_Qty=convert(float,Open_Qty) from ipc_t_download_nb_po a where S_loc='SW03'
union
select POVendor='ICC',Material,Open_Qty=convert(float,Open_Qty) from icc_t_download_nb_po a where S_loc='SW04'
) as b where a.POVendor=b.POVendor and a.Material=b.Material

update #Gating set POBalance=OpenPO+FCSTBalance

---Combine KB to #Gating
update #KBGating set MS=rtrim(MS)+'*'

insert #Gating
    select * from #KBGating
*/
create table #DD(iid int identity(1,1),Dt char(10))

insert #DD
select convert(char(10),Dt,111) from Calendar3 where Dt<=(select dateadd(dd,50,convert(datetime,max(NeedShipDate)+' 00:00')) from #tmp_OPO where not NeedShipDate='2099/01/01' ) and 
Dt>=(select dateadd(dd,-50,convert(datetime,min(NeedShipDate)+' 00:00')) from #tmp_OPO) and WorkingDay=1 order by convert(char(10),Dt,111)


create table #MV(Material varchar(20),PNQty int,PN varchar(1000),VendorQty int,Vendor varchar(1000))
create table #VD(iid int identity(1,1),Material varchar(20),Priority varchar(20),VendorName varchar(50))
insert #MV  
   select distinct Material,0,'',0,'' from(
   select distinct Material=replace(ShortagePN,'*','') from #AAResult   
   ) as a 

insert #VD
select distinct a.Material,a.Priority,b.Vendor_Name2 from t_download_sourcelist_CP62DW a,dbo.t_download_supplier_master b where a.Vendor1=b.Vendor_Code and Material in (
select distinct Material from #MV
) order by Material,Priority

update #MV set VendorQty=b.Qty from #MV a,(select Material,Qty=count(*) from #VD group by Material) as b where a.Material=b.Material

--declare @i int 
--declare @j int 
select @i=min(iid) from #VD
select @j=max(iid) from #VD

while @i<=@j
  begin
      update #MV set Vendor=rtrim(Vendor)+rtrim(b.VendorName)+' /' from #MV a,(select Material,VendorName from #VD where iid=@i) as b where a.Material=b.Material
      select @i=@i+1
  end

--drop table #ShortageDetail
create table #ShortageDetail(POVendor varchar(20),ShortagePN varchar(20),PIC varchar(20),Material_descript varchar(100),SODate varchar(20),
NeedShipDate varchar(20),LatestSODate varchar(20),LatestNeedShipDate varchar(20),Site varchar(20),Item varchar(20),RealOpenQty int,RestNeedQty int,RemainedStock int,Buyer varchar(30),Vendor varchar(500),
EstETADate varchar(20),NewShortage varchar(20),MaterialETA varchar(500),Remark varchar(1000))

insert #ShortageDetail 
select distinct a.POVendor,a.ShortagePN,PIC,a.Material_descript,SODate,NeedShipDate,LatestSODate,LatestNeedShipDate,Site,Item,RealOpenQty,0,0,Buyer=isnull(b.Buyer,''),Vendor,'','','','' from
(

select distinct a.*,b.VendorQty,Vendor from
(select distinct POVendor,ShortagePN,PIC,Material_descript,SODate=min(POReceiveDate),NeedShipDate=min(NeedShipDate),LatestSODate=max(POReceiveDate),LatestNeedShipDate=max(NeedShipDate),
Site=count(distinct Site),Item=count(Item),RealOpenQty=sum(RealOpenQty*Qty) from
 ( 
 
select distinct a.*,ShortagePN=isnull(b.ShortagePN,''),Material_descript='( '+isnull(b.Material_group,'')+' )    '+isnull(b.Material_descript,''),Qty,
PIC=case when ShortagePN like '13%' then 'Joan' when ShortagePN like '146%' then 'Joan' when ShortagePN like '15%' then 'Iris' else b.PIC end from 
(---Get Real OpenQty ---2013/05/16 remove ZN order .
select a.iid,a.Site,a.PO,a.SO,a.Item,a.POVendor,a.POReceiveDate,a.IECPN,NeedShipDate,POQty,DockQty,ShipQty,RealOpenQty=b.OpenQty-b.CurrentDone,b.OpenQty,b.CurrentDone from #tmp_OPO a,#ETAfinal b where PType in ('NB','ZB') and  a.iid=b.iid
and b.OpenQty-b.CurrentDone>0
) as a left join
(
select a.*,b.PIC from 
(select a.*,b.Material_descript,b.Material_group from #AAResult a,t_download_matmas_CP62DW b where left(a.ShortagePN,12)=b.Material) as a left join
(select * from MType) as b on a.Material_group=b.Material_group
) as b on a.POVendor=b.POVendor and a.IECPN=b.PF

) as a  where Not ShortagePN='' group by POVendor,ShortagePN,PIC,Material_descript 
) as a left join
(select distinct * from #MV ) as b on replace(a.ShortagePN,'*','')=b.Material

) as a left join
(select distinct POVendor,Material,Buyer from #Mat_SA) as b on a.POVendor=b.POVendor and replace(a.ShortagePN,'*','')=b.Material--a.ShortagePN=b.Material--


---Update EstETADate
update #ShortageDetail set EstETADate=isnull((select Dt from #DD where iid=(select iid from #DD where Dt=NeedShipDate)-10),'')

---Get Previous MaterialETA & Remark
update #ShortageDetail set NewShortage=RealOpenQty-ShortageQty,MaterialETA=isnull(b.MaterialETA,''),Remark=isnull(b.Remark,'') from #ShortageDetail a,
(
select distinct * from Ivan_ShortageMaterialUpdate where convert(char(10),Cdt,111)=convert(char(10),getdate(),111) and Customer='TSB'
) as b where a.POVendor=b.POVendor and a.ShortagePN=b.Material 


update #ShortageDetail set NewShortage='New' where NewShortage=''

update #ShortageDetail set RemainedStock=isnull(b.Qty,0) from #ShortageDetail a,
(
select POVendor,MS=case when MS is null then MatNo else MS end,Qty=sum(Qty) from(

select distinct a.POVendor,MS,MatNo,Qty from (
select * from #INV where INVDate=convert(char(10),getdate(),111)) as a left join 
(select distinct POVendor,MS,ShortagePN from #Alternative) as b 
on a.MatNo=b.ShortagePN 

) as a 
group by POVendor,case when MS is null then MatNo else MS end

) as b where a.POVendor=b.POVendor and a.ShortagePN=b.MS


---(2013/01/23) solve the Bonnie problem....
update #ShortageDetail set RestNeedQty=b.RestNeedQty from #ShortageDetail a,
(
select ShortagePN,RestNeedQty=sum(RestNeedQty) from (
select distinct iid,POVendor,ShortagePN,RestNeedQty=RestNeedQty*Qty from
(select iid,IECPN,RestNeedQty=OpenQty-CurrentDone-FutureDone from #OPODetail where OpenQty-CurrentDone-FutureDone>0) as a inner join
(select * from #AAResult) as b on a.IECPN=b.PF
) as a group by ShortagePN
) as b where a.ShortagePN=b.ShortagePN


--select * from #ShortageDetail where RemainedStock>RestNeedQty
delete from #ShortageDetail where RemainedStock>RestNeedQty


----------Get Shortage Material Model---------------
/*
drop table #MDCount
drop table #SMCount
drop table #MDCountTmp
drop table #MP_POType
*/
create table #MDCount(PN varchar(20),Model varchar(20))
----(2016/04/13) add MP & NPI .
create table #SMCount(iid int identity(1,1),PN varchar(20),MD varchar(1000),MP varchar(20),NPI varchar(20))
create table #MDCountTmp(mid int identity(1,1),MD varchar(20))

insert #MDCount
select distinct ShortagePN,Material_group from
(

--select distinct a.ShortagePN,b.Material_group from #AAResult a,t_download_matmas_CP07DW b where a.PF=b.Material and a.ShortagePN in (select distinct ShortagePN from #ShortageDetail)
--union

select distinct a.ShortagePN,b.Material_group from #AAResult a,t_download_matmas_CP62DW b where a.PF=b.Material and a.ShortagePN in (select distinct ShortagePN from #ShortageDetail)
) as a 

insert #SMCount
      select distinct PN,'','','' from #MDCount

declare @MDCmin int
declare @MDCmax int
declare @MDtmpmin int
declare @MDtmpmax int

select @MDCmin=min(iid) from #SMCount
select @MDCmax=max(iid) from #SMCount

while @MDCmin<=@MDCmax
begin
   insert #MDCountTmp
        select Model from #MDCount where PN=(select PN from #SMCount where iid=@MDCmin)
   
   select @MDtmpmin=min(mid) from #MDCountTmp
   select @MDtmpmax=max(mid) from #MDCountTmp
   
   while @MDtmpmin<=@MDtmpmax
   begin
       update #SMCount set MD=rtrim(MD)+(select MD from #MDCountTmp where mid=@MDtmpmin)+'/ ' where iid=@MDCmin
       select @MDtmpmin=@MDtmpmin+1
   end
   select  @MDCmin=@MDCmin+1
   truncate table #MDCountTmp
end

---select * from #SMCount

select distinct PN,MP,PO_Type into #MP_POType from
(select a.PN,b.iid from #SMCount a,OPS_OM b where  Customer='DYNABOOK' and ReportDate=convert(char(10),getdate(),111) and a.PN=b.ShortagePN ) as a left join
(select iid,MP,PO_Type from OPS_OPO where Customer='DYNABOOK' and ReportDate=convert(char(10),getdate(),111)) as b on a.iid=b.iid

update #SMCount set NPI='Y' from #SMCount a,#MP_POType b where a.PN=b.PN and b.PO_Type='NPI'
update #SMCount set MP='Y' from #SMCount a,#MP_POType b where a.PN=b.PN and b.MP='Y'
--------------------------------------------------
--drop table #ShortageDetail_total
select POVendor,ShortagePN,PIC,Material_descript,GattingExisted=' ',GattingNeed=0,SAPRawMaterialOPO=0,POBalance=0,WholeBuyExisted='                                                                                                                                                                    '
,SODate,NeedShipDate,LatestSODate,LatestNeedShipDate,Site,Item,RealOpenQty,RemainedStock,MTStock=0,MStock=0,Buyer,Vendor,
EstETADate,NewShortage,MaterialETA,Remark,b.MD,b.MP,b.NPI into #ShortageDetail_total from
(select POVendor,ShortagePN,PIC,Material_descript,SODate,NeedShipDate,LatestSODate,LatestNeedShipDate,Site,Item,RealOpenQty,RemainedStock,Buyer,Vendor,
EstETADate,NewShortage,MaterialETA,Remark from #ShortageDetail) as a left join
(select * from #SMCount) as b
on a.ShortagePN=b.PN order by PIC,NeedShipDate,RealOpenQty desc

/*
update #ShortageDetail_total set GattingExisted='Y',GattingNeed=abs(convert(int,b.FCSTBalance)) from #ShortageDetail_total a,#Gating b 
where replace(a.ShortagePN,' ','')=replace(b.Material,' ','') and a.POVendor=b.POVendor and b.FCSTBalance<0 and a.ShortagePN like '6037%*'

update #ShortageDetail_total set GattingExisted='Y',GattingNeed=abs(convert(int,b.FCSTBalance)) from #ShortageDetail_total a,#Gating b 
where left(a.ShortagePN,12)=left(b.Material,12) and a.POVendor=b.POVendor and b.FCSTBalance<0 and not a.ShortagePN like '6037%*'


insert #ShortageDetail_total 
      select distinct POVendor,Material,'','','Y',GattingNeed,'','','','2099/01/01','2099/01/01','0','0','0','0','0','0',Buyer,'','2099/01/01','','','','' from (
      select * from (select POVendor,Material,Buyer,GattingNeed=abs(FCSTBalance) from #Gating where FCSTBalance<0) as a left join
      (select distinct PO=POVendor,ShortagePN from #ShortageDetail_total) as b on left(a.Material,12)=left(b.ShortagePN,12) and a.POVendor=b.PO
      ) as a where ShortagePN is null 
      
--select * from #ShortageDetail_total where ShortagePN like '6070B0438501%'
*/
update #ShortageDetail_total set Buyer=b.Buyer from #ShortageDetail_total a,BuyerCode b where a.Buyer=b.BuyerCode
update #ShortageDetail_total set Vendor=b.Vendor from #ShortageDetail_total a,#MV b where a.ShortagePN=b.Material and a.Vendor=''
update #ShortageDetail_total set Material_descript='( '+isnull(b.Material_group,'')+' )    '+isnull(b.Material_descript,'') from #ShortageDetail_total a,t_download_matmas_CP62DW b 
where replace(a.ShortagePN,'*','')=b.Material and a.Material_descript=''


----Add Inventory for Gating Items .
update #ShortageDetail_total set RemainedStock=b.Qty from #ShortageDetail_total a,#INV b where a.POVendor=b.POVendor 
and a.ShortagePN=b.MatNo and b.MaterialProperty='FG' and a.SODate='2099/01/01'



update #ShortageDetail_total set PIC=c.PIC from #ShortageDetail_total a,t_download_matmas_CP62DW b,MType c where replace(a.ShortagePN,'*','')=b.Material and b.Material_group=c.Material_group
and a.PIC=''

update #ShortageDetail_total set PIC='Joan' where ShortagePN like '13%' and PIC=''
update #ShortageDetail_total set PIC='Iris' where ShortagePN like '15%' and PIC=''



-------------------=====================================================================================
----Remove the item which Inventory is enough .
--select * from #ShortageDetail_total where SODate='2099/01/01' and RemainedStock>GattingNeed
delete from #ShortageDetail_total where SODate='2099/01/01' and RemainedStock>GattingNeed

----(2013/06/03) get PO Balance
/*
update #ShortageDetail_total set POBalance=b.POBalance from #ShortageDetail_total a, 
(select POVendor,Material,POBalance from #Gating where FCSTBalance<0) as b 
where a.POVendor=b.POVendor and a.ShortagePN=b.Material


----(2013/07/18) get PO Balance by 
--drop table #SAPPO
select Plnt=case when Plnt='CP81' then 'IES' when Plnt='CP60' then 'ICC' else '' end ,Material,Open_Qty=convert(int,sum(convert(float,replace(Open_Qty,' ','')))) into #SAPPO from (
select PO_Number,Item,Plnt,Material,Description,Vendor,Buyer,Create_Dt=convert(char(10),convert(datetime,Create_Dt+' 00:00'),111),PO_Qty,Open_Qty from ipc_t_download_nb_po where S_loc='SW03' 
union
select PO_Number,Item,Plnt,Material,Description,Vendor,Buyer,Create_Dt=convert(char(10),convert(datetime,Create_Dt+' 00:00'),111),PO_Qty,Open_Qty from icc_t_download_nb_po where S_loc='SW04'
) as a group by Plnt,Material


--Get K/B PO Balance
insert #SAPPO
select Plnt,MS=rtrim(MS)+'*',Open_Qty=sum(Open_Qty) from (
select distinct a.Plnt,b.MS,b.Material,a.Open_Qty from
(select * from #SAPPO where Material like '6037%') as a left join
(select * from KBAlternative) as b on a.Material=b.Material) as a where not MS is null group by Plnt,MS

delete from #SAPPO where Material like '6037%' and not Material like '%*'

update #ShortageDetail_total set SAPRawMaterialOPO=b.Open_Qty from #ShortageDetail_total a,#SAPPO b where a.POVendor=b.Plnt and a.ShortagePN=b.Material

--select * from #SAPPO where Material='6070B0372402'
--select top 10 * from #ShortageDetail_total where ShortagePN='6070B0372402'
update #ShortageDetail_total set POBalance=b.Open_Qty-GattingNeed from #ShortageDetail_total a,#SAPPO b where a.POVendor=b.Plnt and a.ShortagePN=b.Material

--select * from #ShortageDetail_total a,#SAPPO b where a.POVendor=b.Plnt and a.ShortagePN=b.Material
*/

----Remove Packing Material
delete from #ShortageDetail_total where ShortagePN like '606%' and SODate='2099/01/01'
delete from #ShortageDetail_total where ShortagePN like '6Z%' and SODate='2099/01/01'
delete from #ShortageDetail_total where ShortagePN like '606%' and SODate='2099/01/01'

----Remove samll EE Material
delete from #ShortageDetail_total where left(ShortagePN,4) in ('6010','6011','6013','6014','6015','6016','6018','6071') and SODate='2099/01/01'


----------------------Get Whole Buy
/*
drop table #WBinit
drop table #WBresult
drop table #WBmid


create table #WBinit(Material varchar(20),SCD varchar(50))
create table #WBresult(iid int identity(1,1),Material varchar(20),SCD varchar(1000))
create table #WBmid(mid int identity(1,1),Material varchar(20),SCD varchar(50))

insert #WBinit
select left(Material,12),
SCD=convert(char(8),ServiceConfirmDate,112)+'-'+
case when PurchaseType='Whole buy' then 'W' 
when PurchaseType='MOQ' then 'M'
when PurchaseType='LTB' then 'L'
else '?' end+'-'+
convert(varchar(10),ServiceDemandQty)+'-'+
convert(varchar(10),BalanceQty)+'-'+
+case when ItemStatus='Open' then 'O'
when ItemStatus='Closed' then 'C'
when ItemStatus='Pending' then 'P'
else '?' end from WholeBuy order by SCD



insert #WBresult
    select distinct Material,'' from #WBinit

declare @WBi int
declare @WBj int
declare @WBa int
declare @WBb int

select @WBi=min(iid) from #WBresult
select @WBj=max(iid) from #WBresult

while @WBi<=@WBj
begin
     insert #WBmid
         select Material,SCD from #WBinit where Material=(select Material from #WBresult where iid=@WBi) order by SCD
         select @WBa=min(mid) from #WBmid
         select @WBb=max(mid) from #WBmid 
         while @WBa<=@WBb 
         begin
           update #WBresult set SCD=rtrim(a.SCD)+b.SCD+' /' from #WBresult a,(select * from #WBmid where mid=@WBa) as b where a.Material=b.Material 
           select @WBa=@WBa+1
         end
     truncate table #WBmid
     select @WBi=@WBi+1
end


-----------------------------------
update #ShortageDetail_total set WholeBuyExisted=b.SCD from #ShortageDetail_total a,#WBresult b where a.ShortagePN=b.Material
*/
--select distinct FF=left(Material_descript,charindex(')',Material_descript)) from #ShortageDetail_total order by FF
--select * from #ShortageDetail_total order by PIC,NeedShipDate,GattingNeed desc where Material_descript like '%LCM%' not ShortagePN like '601%'

-----Updated Material_ETA for Gating items .
update #ShortageDetail_total set NewShortage=case when isnull(b.ShortageQty,0)=0 then 'New' else convert(varchar(10),RealOpenQty-isnull(b.ShortageQty,0)) end 
,MaterialETA=isnull(b.MaterialETA,''),Remark=isnull(b.Remark,'') from #ShortageDetail_total a,
(select distinct * from Ivan_ShortageMaterialUpdate where convert(char(10),Cdt,111)=convert(char(10),getdate(),111) and Customer='TSB') as b
where a.POVendor=b.POVendor and a.ShortagePN=b.Material and a.MaterialETA=''

update #ShortageDetail_total set NewShortage='New' where NewShortage=''


update #ShortageDetail_total set Remark='(ETA Delay : '+convert(varchar(5),b.ETADiff)+' )'+rtrim(Remark) from #ShortageDetail_total a,Ivan_CurrentINV b 
where a.POVendor=b.POVendor and a.ShortagePN=b.MatNo and ETADiff<0 and Customer='TSB'



update #ShortageDetail_total set Remark='(ETA Improve : '+convert(varchar(5),b.ETADiff)+' )'+rtrim(Remark) from #ShortageDetail_total a,Ivan_CurrentINV b 
where a.POVendor=b.POVendor and a.ShortagePN=b.MatNo and ETADiff>0 and not ETADiff=9999 and Customer='TSB'


---(2013/10/07 Get Manufacture Inventory)
/*
update #ShortageDetail_total set MTStock=b.Qty from #ShortageDetail_total a,Ivan_Current_M_INV b where a.POVendor=b.POVendor 
and a.ShortagePN=b.MatNo and b.InventoryType='MT-WH' and b.MP='M_FG' and Customer='TSB'

update #ShortageDetail_total set MStock=b.Qty from #ShortageDetail_total a,Ivan_Current_M_INV b where a.POVendor=b.POVendor 
and a.ShortagePN=b.MatNo and b.InventoryType='M-WH' and b.MP='M_FG' and Customer='TSB'
*/
---(2015/05/08 Add Alternative check )
update #ShortageDetail_total set MTStock=b.Qty from #ShortageDetail_total a,(
select POVendor,ShortagePN,Qty=sum(Qty) from (
select a.*,b.Qty from 
(
select a.POVendor,ShortagePN,SP=case when SP is null then ShortagePN else SP end from (
select a.POVendor,a.ShortagePN,SP=b.ShortagePN from 
(select distinct POVendor,ShortagePN from #ShortageDetail_total) as a left join
(select distinct MS,ShortagePN from #Alternative) as b on a.ShortagePN=b.MS
) as a
) as a left join
(select * from Ivan_Current_M_INV where MP='MT-WH' and Customer='TSB') as b on a.SP=b.MatNo and a.POVendor=b.POVendor
) as a where not Qty is null group by POVendor,ShortagePN
) as b where a.POVendor=b.POVendor and a.ShortagePN=b.ShortagePN


update #ShortageDetail_total set MStock=b.Qty from #ShortageDetail_total a,(
select POVendor,ShortagePN,Qty=sum(Qty) from (
select a.*,b.Qty from 
(
select a.POVendor,ShortagePN,SP=case when SP is null then ShortagePN else SP end from (
select a.POVendor,a.ShortagePN,SP=b.ShortagePN from 
(select distinct POVendor,ShortagePN from #ShortageDetail_total) as a left join
(select distinct MS,ShortagePN from #Alternative) as b on a.ShortagePN=b.MS
) as a
) as a left join
(select * from Ivan_Current_M_INV where MP='M_FG' and Customer='TSB') as b on a.SP=b.MatNo and a.POVendor=b.POVendor
) as a where not Qty is null group by POVendor,ShortagePN
) as b where a.POVendor=b.POVendor and a.ShortagePN=b.ShortagePN

insert OPS_Material
select Customer='DYNABOOK',ReportDate=convert(char(10),getdate(),111),*  from #ShortageDetail_total order by PIC,NeedShipDate,RealOpenQty,GattingNeed desc


-----------Summary (By PIC)-----------------------------------------------------------------
insert OPS_MSummaryPIC
select Customer='DYNABOOK',ReportDate=convert(char(10),getdate(),111),PIC=isnull(PIC,''),Item=sum(Item),Qty=sum(Qty),NoETAItem=sum(NoETAItem),NoETAQty=sum(NoETAQty) from
(
select PIC,Item=count(*),Qty=0,NoETAItem=0,NoETAQty=0 from #ShortageDetail_total group by PIC
union
select PIC,0,sum(RealOpenQty),0,0 from #ShortageDetail_total group by PIC
union
select PIC,0,0,count(*),0 from #ShortageDetail_total where MaterialETA='' group by PIC
union
select PIC,0,0,0,sum(RealOpenQty) from #ShortageDetail_total where MaterialETA='' group by PIC
) as a group by PIC order by Item desc

--Summary (By Type)
insert OPS_MSummaryType
select Customer='DYNABOOK',ReportDate=convert(char(10),getdate(),111),MType,PIC=isnull(PIC,''),Item=sum(Item),Qty=sum(Qty),NoETAItem=sum(NoETAItem),NoETAQty=sum(NoETAQty) from
(
select MType,PIC,Item=sum(Item),Qty=0,NoETAItem=0,NoETAQty=0 from(
select MType=case when replace(replace(left(Material_descript,charindex(')',Material_descript)-1),'(',''),' ','') like 'F%' then 'SA' else
replace(replace(left(Material_descript,charindex(')',Material_descript)-1),'(',''),' ','') 
end,PIC
 ,Item=count(*) from #ShortageDetail_total 
group by replace(replace(left(Material_descript,charindex(')',Material_descript)-1),'(',''),' ',''),PIC
) as a group by MType,PIC
union
select MType,PIC,0,Qty=sum(Qty),0,0 from(
select MType=case when replace(replace(left(Material_descript,charindex(')',Material_descript)-1),'(',''),' ','') like 'F%' then 'SA' else
replace(replace(left(Material_descript,charindex(')',Material_descript)-1),'(',''),' ','') 
end,PIC
 ,Qty=sum(RealOpenQty) from #ShortageDetail_total 
group by replace(replace(left(Material_descript,charindex(')',Material_descript)-1),'(',''),' ',''),PIC
) as a group by MType,PIC
union
select MType,PIC,0,0,Item=sum(Item),0 from(
select MType=case when replace(replace(left(Material_descript,charindex(')',Material_descript)-1),'(',''),' ','') like 'F%' then 'SA' else
replace(replace(left(Material_descript,charindex(')',Material_descript)-1),'(',''),' ','') 
end,PIC
 ,Item=count(*) from #ShortageDetail_total where MaterialETA='' 
group by replace(replace(left(Material_descript,charindex(')',Material_descript)-1),'(',''),' ',''),PIC
) as a group by MType,PIC
union
select MType,PIC,0,0,0,NoETAQty=sum(NoETAQty) from(
select MType=case when replace(replace(left(Material_descript,charindex(')',Material_descript)-1),'(',''),' ','') like 'F%' then 'SA' else
replace(replace(left(Material_descript,charindex(')',Material_descript)-1),'(',''),' ','') 
end,PIC
 ,NoETAQty=sum(RealOpenQty) from #ShortageDetail_total where MaterialETA='' 
group by replace(replace(left(Material_descript,charindex(')',Material_descript)-1),'(',''),' ',''),PIC
) as a group by MType,PIC
) as a group by MType,PIC order by Item desc


---Alternative
insert OPS_Alt
select distinct Customer='DYNABOOK',ReportDate=convert(char(10),getdate(),111),MS,ShortagePN,Priority,Usage,Item,RefPN from #Alternative order by MS,RefPN,Item

---Shortage material Inventory consumer status
insert OPS_Inventory
select distinct Customer='DYNABOOK',ReportDate=convert(char(10),getdate(),111),a.POVendor,a.MaterialProperty,a.MatNo,Material_descript='( '+isnull(b.Material_group,'')+' )    '+isnull(b.Material_descript,''),
RemainedQty=Qty,StockExistedDate=INVDate,ConsumerStatus=a.Remark from #INV a,t_download_matmas_CP62DW b where left(a.MatNo,12)=b.Material 



---(2013/07/01) Raw Material OPOs 
---(2013/07/01) Raw Material OPOs 
insert OPS_RawOPO
select Customer='DYNABOOK',ReportDate=convert(char(10),getdate(),111),a.*,MaterialETA=isnull(MaterialETA,''),Remark=isnull(Remark,'') from
(
select * from (
select distinct PO_Number,Item,Plnt,S_loc,Material,Description,Vendor,Buyer,Create_Dt=convert(char(10),convert(datetime,Create_Dt+' 00:00'),111),
PO_Qty=convert(float,replace(PO_Qty,' ','')),
Open_Qty=convert(float,replace(Open_Qty,' ','')) 
from cp07_t_download_nb_po where Plnt='CP07' and S_loc='SWF3' 
union
select distinct PO_Number,Item,Plnt,S_loc,Material,Description,Vendor,Buyer,Create_Dt=convert(char(10),convert(datetime,Create_Dt+' 00:00'),111),
PO_Qty=convert(float,replace(PO_Qty,' ','')),
Open_Qty=convert(float,replace(Open_Qty,' ','')) 
from cp62_t_download_nb_po where Plnt='CP62' and S_loc='SWF3'
) as a 
) as a left join
(select distinct POVendor=case when POVendor='IES' then 'CP07' when POVendor='ICC' then 'CP62' else '' end,Material,MaterialETA,Remark 
from dbo.Ivan_ShortageMaterialUpdate where convert(char(10),Cdt,111)=convert(char(10),getdate(),111) and Customer='DYNABOOK'
) as b on a.Plnt=b.POVendor and a.Material=b.Material order by Plnt,convert(datetime,Create_Dt+' 00:00')


---------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------
-------New------Summary OPO Status
------Normal OTD
insert OPS_OTDFailD
select Customer='DYNABOOK',ReportDate=convert(char(10),getdate(),111),OPOType='Normal OTD',* from (
select a.*,b.ShortagePN,b.Material_descript,b.MaterialETA,b.Remark from
(select IECPN,ProductFamily,NeedShipDate=min(NeedShipDate),OpenQty=sum(OpenQty) from #OPODetail where OPOR like '(OTD Fail%' and not PO_Type='NPI' group by IECPN,ProductFamily) as a left join
(select distinct IECPN,ShortagePN,Material_descript,MaterialETA,Remark from #OM where NeedQty>CurrentStock) as b on a.IECPN=b.IECPN
) as a where not ShortagePN is null order by OpenQty desc,IECPN

---NPI OTD 
insert OPS_OTDFailD
select Customer='DYNABOOK',ReportDate=convert(char(10),getdate(),111),OPOType='NPI OTD',* from (
select a.*,b.ShortagePN,b.Material_descript,b.MaterialETA,b.Remark from
(select IECPN,ProductFamily,NeedShipDate=min(NeedShipDate),OpenQty=sum(OpenQty) from #OPODetail where OPOR like '(OTD Fail%' and PO_Type='NPI' group by IECPN,ProductFamily) as a left join
(select distinct IECPN,ShortagePN,Material_descript,MaterialETA,Remark from #OM where NeedQty>CurrentStock) as b on a.IECPN=b.IECPN
) as a where not ShortagePN is null order by OpenQty desc,IECPN

---Normal Potential
insert OPS_OTDFailD
select Customer='DYNABOOK',ReportDate=convert(char(10),getdate(),111),OPOType='Normal Potential',* from (
select a.*,b.ShortagePN,b.Material_descript,b.MaterialETA,b.Remark from
(select IECPN,ProductFamily,NeedShipDate=min(NeedShipDate),OpenQty=sum(OpenQty) from #OPODetail where OPOR like '%Potential%' and not PO_Type='NPI' group by IECPN,ProductFamily) as a left join
(select distinct IECPN,ShortagePN,Material_descript,MaterialETA,Remark from #OM where NeedQty>CurrentStock) as b on a.IECPN=b.IECPN
) as a where not ShortagePN is null order by OpenQty desc,IECPN

----NPI Potential
insert OPS_OTDFailD
select Customer='DYNABOOK',ReportDate=convert(char(10),getdate(),111),OPOType='NPI Potential',* from (
select a.*,b.ShortagePN,b.Material_descript,b.MaterialETA,b.Remark from
(select IECPN,ProductFamily,NeedShipDate=min(NeedShipDate),OpenQty=sum(OpenQty) from #OPODetail where OPOR like '%Potential%' and PO_Type='NPI' group by IECPN,ProductFamily) as a left join
(select distinct IECPN,ShortagePN,Material_descript,MaterialETA,Remark from #OM where NeedQty>CurrentStock) as b on a.IECPN=b.IECPN
) as a where not ShortagePN is null order by OpenQty desc,IECPN



---------------------------Summary
---Summary Normal OTD 
insert OPS_OTDFailS
select Customer='DYNABOOK',ReportDate=convert(char(10),getdate(),111),OPOType='Normal OTD',MType=left(Material_descript,charindex(')',Material_descript)),PType=count(distinct ShortagePN),ItemQty=count(POItem),PTypeQty=sum(OpenQty) from (
select * from (
select a.*,b.ShortagePN,b.Material_descript,b.MaterialETA,b.Remark from
(select IECPN,ProductFamily,PO,POItem,NeedShipDate=min(NeedShipDate),OpenQty=sum(OpenQty) from #OPODetail where OPOR like '(OTD Fail%' 
and not PO_Type='NPI' group by IECPN,ProductFamily,PO,POItem) as a left join
(select distinct IECPN,ShortagePN,Material_descript,MaterialETA,Remark from #OM where NeedQty>CurrentStock) as b on a.IECPN=b.IECPN
) as a where not ShortagePN is null
) as a where not Material_descript is null group by left(Material_descript,charindex(')',Material_descript)) order by ItemQty desc

---Summary NPI OTD
insert OPS_OTDFailS
select Customer='DYNABOOK',ReportDate=convert(char(10),getdate(),111),OPOType='NPI OTD',MType=left(Material_descript,charindex(')',Material_descript)),PType=count(distinct ShortagePN),ItemQty=count(POItem),PTypeQty=sum(OpenQty) from (
select * from (
select a.*,b.ShortagePN,b.Material_descript,b.MaterialETA,b.Remark from
(select IECPN,ProductFamily,PO,POItem,NeedShipDate=min(NeedShipDate),OpenQty=sum(OpenQty) from #OPODetail where 
OPOR like '(OTD Fail%' and PO_Type='NPI' group by IECPN,ProductFamily,PO,POItem) as a left join
(select distinct IECPN,ShortagePN,Material_descript,MaterialETA,Remark from #OM where NeedQty>CurrentStock) as b on a.IECPN=b.IECPN
) as a where not ShortagePN is null
) as a where not Material_descript is null group by left(Material_descript,charindex(')',Material_descript)) order by count(*) desc

---Summary
insert OPS_OTDFailS
select Customer='DYNABOOK',ReportDate=convert(char(10),getdate(),111),OPOType='Normal Potential',MType=left(Material_descript,charindex(')',Material_descript)),PType=count(distinct ShortagePN),ItemQty=count(POItem),PTypeQty=sum(OpenQty) from (
select * from (
select a.*,b.ShortagePN,b.Material_descript,b.MaterialETA,b.Remark from
(select IECPN,ProductFamily,PO,POItem,NeedShipDate=min(NeedShipDate),OpenQty=sum(OpenQty) from #OPODetail 
where OPOR like '%Potential%' and not PO_Type='NPI' group by IECPN,ProductFamily,PO,POItem) as a left join
(select distinct IECPN,ShortagePN,Material_descript,MaterialETA,Remark from #OM where NeedQty>CurrentStock) as b on a.IECPN=b.IECPN
) as a where not ShortagePN is null
) as a where not Material_descript is null group by left(Material_descript,charindex(')',Material_descript)) order by count(*) desc


---Summary NPI Potential OTD
insert OPS_OTDFailS
select Customer='DYNABOOK',ReportDate=convert(char(10),getdate(),111),OPOType='NPI Potential',MType=left(Material_descript,charindex(')',Material_descript)),PType=count(distinct ShortagePN),ItemQty=count(POItem),PTypeQty=sum(OpenQty) from (
select * from (
select a.*,b.ShortagePN,b.Material_descript,b.MaterialETA,b.Remark from
(select IECPN,ProductFamily,PO,POItem,NeedShipDate=min(NeedShipDate),OpenQty=sum(OpenQty) from #OPODetail 
where OPOR like '%Potential%' and PO_Type='NPI' group by IECPN,ProductFamily,PO,POItem) as a left join
(select distinct IECPN,ShortagePN,Material_descript,MaterialETA,Remark from #OM where NeedQty>CurrentStock) as b on a.IECPN=b.IECPN
) as a where not ShortagePN is null
) as a where not Material_descript is null group by left(Material_descript,charindex(')',Material_descript)) order by count(*) desc

select 'OK!!'
return


