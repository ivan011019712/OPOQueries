------------------------------------------------
---------------------------------------------------
/*
drop table #tmp
drop table #DN
drop table #result
drop table #OPOresult
drop table #OM
--drop table #Compare
drop table #OMFinal
*/

delete from OldOPO where ReportDate=convert(char(10),getdate(),111) and Customer='H2'
delete from OPS_OPO_new where ReportDate=convert(char(10),getdate(),111) and Customer='H2' 

select *,EST_Date='0000/00/00',IES_DN='--------------------',IES_DNPGI='0000/00/00',MType='--------------------',Category='-----------',LateThanNeedShipDate='-',OTDType='----------',OTDStatus='----------', 
Result='----------------------------------------' into #tmp from OPS_OPO where ReportDate=convert(char(10),getdate(),111) and Customer='H2'

update #tmp set IES_DN='',IES_DNPGI=''

---Get DN
select Site,PO,SO,Item,IECPN,PO_Date,Qty856,IES_DN,IES_DNPGI into #DN from Service_APD 
where Site in (Select distinct ZS92Site from SiteMapping where Customer='HP' and not ZS92Site='') 
and PndGIDate='0000/00/00' and not IES_DNPGI='0000/00/00' and SO like '11%'
 
----(2017/07/12) Modify Item number to PO Item number (join ZM57) to solve strange data .
---Update DN & DN date
/*
update #tmp set IES_DNPGI=b.IES_DNPGI from #tmp a,
(select SO,Item=convert(int,Item),IECPN,PO_Date,IES_DNPGI=max(IES_DNPGI) from #DN group by SO,Item,IECPN,PO_Date) as b where a.SO=b.SO and a.IECPN=b.IECPN and a.POReceiveDate=b.PO_Date and a.POItem=b.Item
*/

update #tmp set IES_DNPGI=b.IES_DNPGI from #tmp a,
(
select a.SO,SOItem=a.Item,Item=convert(int,IPCSOItem),a.IECPN,PO_Date,IES_DNPGI=max(IES_DNPGI) from #DN a,ZM57 b where 
a.SO=b.IECSO and a.Item=b.IECSPItem --and a.IECPN='JFBL01BMB002' 
group by a.SO,a.Item,convert(int,IPCSOItem),a.IECPN,a.PO_Date
) as b where a.SO=b.SO and a.IECPN=b.IECPN and a.POReceiveDate=b.PO_Date and a.POItem=b.Item


update #tmp set IES_DN=b.IES_DN from #tmp a,
(select distinct a.*,newItem=convert(int,b.IPCSOItem) from #DN a,ZM57 b where a.SO=b.IECSO and a.Item=b.IECSPItem) as b
where a.SO=b.SO and a.IECPN=b.IECPN and a.POReceiveDate=b.PO_Date and a.IES_DNPGI=b.IES_DNPGI and convert(int,a.POItem)=b.newItem


------update MType (NB ,DT or Portable...) 
update #tmp set MType='Raw Material' where IECPN like '6%'

--select * from #tmp where MType='--------------------'
/*
update #tmp set MType='HP DT' where IECPN in ('JFCU91DCK002','JFCU51ASP002','JFCU91ASP002','JFCU91AAN002') and  MType='--------------------'
update #tmp set MType='HP PORTABLE' where IECPN='1510B1464601' and  MType='--------------------'
update #tmp set MType='HP PORTABLE' where MType='--------------------' and ProductFamily='F-SAGE10'
update #tmp set MType='HP PORTABLE' where MType='--------------------' and ProductFamily='F-VANIL10'
update #tmp set MType='HP PORTABLE' where MType='--------------------' and IECPN='PFDS62ALB002'
update #tmp set MType='HP DT' where MType='--------------------' and IECPN like 'JF1933%'
update #tmp set MType='HP DT' where MType='--------------------' and IECPN like 'JF%'
update #tmp set MType='HP DT' where MType='--------------------' and IECPN='1650B0148301'
*/
----(2016/12/29 update Schumi to HP DT)
update #tmp set MType='HP DT' where MType='--------------------' and IECPN in (select distinct ODMPartNumber from SPB where Model like 'SCHUMI%' and OSSPOrderable='Yes')
update #tmp set MType='HP DT' where MType='--------------------' and IECPN in (select distinct ODMPartNumber from SPB where Model in (select distinct HPPlatform from ModelID where Customer='HP AIO'))
update #tmp set MType='HP PORTABLE' where MType='--------------------' and IECPN='1310A2597101'
update #tmp set MType='HP PORTABLE' where MType='--------------------' and IECPN='PFDT01BHS002'
update #tmp set MType='HP PORTABLE' where MType='--------------------' and IECPN like 'PF%'
update #tmp set MType='HP DT' where MType='--------------------' and IECPN='JFCU81KBD021'
update #tmp set MType='HP DT' where MType='--------------------' and IECPN='JF1972ALM002'
update #tmp set MType='HP DT' where MType='--------------------' and IECPN like 'JF%'
update #tmp set MType='HP DT' where MType='--------------------' and IECPN='JC1933AAL01Y'
update #tmp set MType='HP PORTABLE' where MType='--------------------' and PO='RMA # 102649'
update #tmp set MType='HP DT' where MType='--------------------' and IECPN='1650B0148301'
update #tmp set MType='HP DT' where MType='--------------------' and PO='D76264502' and CPQNo='Z5M06AA#ABA'

update #tmp set MType=c.Customer from #tmp a,PNModel b,ModelID c where MType='--------------------' and a.IECPN=b.CPQNo and b.FamilyNo=c.FamilyNo and substring(a.IECPN,2,1)='F'

update #tmp set MType='HP DT' where MType='HP AIO'
--select distinct MType from #tmp 
--select * from #tmp where MType='HP AIO'

-----Get Category
update #tmp set Category=substring(IECPN,8,2) where not IECPN like '6%'
update #tmp set Category='Raw' where IECPN like '6%'
--select distinct Category from #tmp 

-----Define OTDType
update #tmp set OTDType='Past Due' where NeedShipDate<convert(char(10),getdate(),111)
update #tmp set OTDType='Potential' where NeedShipDate>=convert(char(10),getdate(),111)


--select Site,SO,IECPN,POReceiveDate,NeedShipDate,POQty,DNQty,DockQty,ShipQty,OpenQty,CurrentDone,FutureDone,FutureFirstSupportDate,Shortage,OPOR,IES_DNPGI from #tmp
---Add EST_Date
---1. (DN exists) If exist DN ,then EST_Date folows DN Date .
update #tmp set EST_Date=IES_DNPGI where not IES_DNPGI='' and EST_Date='0000/00/00'

----(2021/11/05) Change MB from 14 days to 7 days per Sandy's requests.
---2. (FutureSupportDate exists) If MB ,EST_Date=FutureSupportDate+14 ,others  EST_Date=FutureSupportDate+7
--select * from #tmp where not FutureFirstSupportDate='' and EST_Date='0000/00/00' and not substring(IECPN,8,2)='MB'
update #tmp set EST_Date=convert(char(10),dateadd(dd,7,convert(datetime,FutureFirstSupportDate)),111) where not FutureFirstSupportDate='' and EST_Date='0000/00/00' and substring(IECPN,8,2)='MB'
update #tmp set EST_Date=convert(char(10),dateadd(dd,7,convert(datetime,FutureFirstSupportDate)),111) where not FutureFirstSupportDate='' and EST_Date='0000/00/00' and not substring(IECPN,8,2)='MB'

-------(2020/07/22) Add 14 days for the Panel which need to Glue or Clean Room
update #tmp set EST_Date=convert(char(10),dateadd(dd,14,convert(datetime,FutureFirstSupportDate)),111) where not FutureFirstSupportDate='' --and EST_Date='0000/00/00' 
and IECPN in (select distinct IECPN from PanelAssyType$ where (Glue in ('G1','G2','G3') or CR='C1')  ) and IES_DNPGI=''


--3 "ZN" set EST_Date=NeedShipDate
update #tmp set EST_Date=case when NeedShipDate<=convert(char(10),dateadd(dd,4,getdate()),111) 
then convert(char(10),dateadd(dd,4,getdate()),111) 
else NeedShipDate end
where PType='ZN' and EST_Date='0000/00/00'

---3 No Shortage ,set EST_Date=NeedShipDate
update #tmp set EST_Date=case when NeedShipDate<=convert(char(10),dateadd(dd,14,getdate()),111) 
then convert(char(10),dateadd(dd,14,getdate()),111) 
else NeedShipDate end
where Shortage='N' and EST_Date='0000/00/00' and substring(IECPN,8,2) in (/*'MB',*/'LM','BD')

update #tmp set EST_Date=case when NeedShipDate<=convert(char(10),dateadd(dd,7,getdate()),111) 
then convert(char(10),dateadd(dd,7,getdate()),111) 
else NeedShipDate end
where Shortage='N' and EST_Date='0000/00/00' and not substring(IECPN,8,2) in (/*'MB',*/'LM','BD')

---4. Set EST Date equals to NeedShipDate if EST<=NeedShipDate 
update #tmp set EST_Date=NeedShipDate,OTDStatus='CanShip' where EST_Date<NeedShipDate and IES_DNPGI='' 
and not EST_Date='0000/00/00' and NeedShipDate>convert(char(10),getdate(),111)

---5. Set EST Date equals to NeedShipDate if EST Date is blank and NeedShipDate>dateadd(mm,1,Getdate()) 
update #tmp set EST_Date=NeedShipDate,OTDStatus='CanShip' where EST_Date='0000/00/00' and MP='Y' and not Remark like '%No MP%' and not PO_Type in ('NPI','NPI2') and NeedShipDate>convert(char(10),dateadd(mm,1,getdate()),111)
update #tmp set EST_Date=NeedShipDate,OTDStatus='CanShip' where EST_Date='0000/00/00' and MP='N' and not Remark like '%No MP%' and not PO_Type in ('NPI','NPI2') and NeedShipDate>convert(char(10),dateadd(dd,45,getdate()),111)
and not IECPN in ('PFAS02CVG002','PFAS01CHS002','PFCM01ALM002','PFBE01BWL002','PFBE08AWL002','JF1575AMB002',
'PFCJ01CWL002','PFBP31ALM002','PFBP31CLM002','PFBP31DLM002','PFAW21BWL002','PFAY01CWL002','PFBS01CWL002','PFBS01EWL002')
--select * from #tmp where EST_Date='0000/00/00' and MP='Y' and not Remark like '%No MP%' and not PO_Type in ('NPI','NPI2') and NeedShipDate>convert(char(10),dateadd(mm,1,getdate()),111)
--select * from #tmp where EST_Date='0000/00/00' and MP='N' and not Remark like '%No MP%' and not PO_Type in ('NPI','NPI2') and NeedShipDate>convert(char(10),dateadd(dd,45,getdate()),111) and not IECPN in ('PFAS02CVG002','PFAS01CHS002','PFCM01ALM002','PFBE01BWL002','PFBE08AWL002','JF1575AMB002','PFCJ01CWL002','PFBP31ALM002','PFBP31CLM002','PFBP31DLM002','PFAW21BWL002','PFAY01CWL002','PFBS01CWL002','PFBS01EWL002')





--select * from #tmp where EST_Date='0000/00/00'
update #tmp set EST_Date='' where EST_Date='0000/00/00'

----Change from IES_DNPGI to EST_Date
update #tmp set LateThanNeedShipDate='Y' where EST_Date>NeedShipDate and not EST_Date='' and LateThanNeedShipDate='-'
update #tmp set LateThanNeedShipDate='N' where EST_Date<=NeedShipDate and not EST_Date='' and LateThanNeedShipDate='-'
 
--update #tmp set LateThanNeedShipDate='Y' where convert(char(10),dateadd(dd,5,convert(datetime,FutureFirstSupportDate)),111)>NeedShipDate and not FutureFirstSupportDate='' and LateThanNeedShipDate='-'
--update #tmp set LateThanNeedShipDate='N' where convert(char(10),dateadd(dd,5,convert(datetime,FutureFirstSupportDate)),111)<=NeedShipDate and not FutureFirstSupportDate='' and LateThanNeedShipDate='-'

update #tmp set LateThanNeedShipDate='N' where Shortage='N' and OTDType='Potential' and LateThanNeedShipDate='-'
update #tmp set LateThanNeedShipDate='Y' where Shortage='N' and OTDType='Past Due' and LateThanNeedShipDate='-'



------update OTDType
update #tmp set OTDStatus='Dock' where OPOR like '%Wait for Pick up%' and OTDStatus='----------' 
update #tmp set OTDStatus='NotOK' where LateThanNeedShipDate='-' and OTDStatus='----------'

--update #tmp set OTDStatus='OK' where LateThanNeedShipDate='N' and OTDType='Potential' and OTDStatus='----------'

update #tmp set OTDStatus='LateShip' where LateThanNeedShipDate='Y' and OTDType='Potential' and OTDStatus='----------'
update #tmp set OTDStatus='CanShip' where LateThanNeedShipDate='N' and OTDType='Potential' and OTDStatus='----------'

update #tmp set OTDStatus='LateShip' where LateThanNeedShipDate='Y' and OTDType='Past Due' and OTDStatus='----------'
update #tmp set OTDStatus='CanShip' where LateThanNeedShipDate='N' and OTDType='Past Due' and OTDStatus='----------'

update #tmp set OTDStatus='CanShip' where PType='ZN' and not OTDStatus='Dock'




-----Update Result
update #tmp set Result='Past Due with ETA ,in OTD' where OTDType='Past Due' and OTDStatus='CanShip'
update #tmp set Result='Past Due with ETA ,Out of OTD' where OTDType='Past Due' and OTDStatus='LateShip'
update #tmp set Result='Past Due w/o ETA' where OTDType='Past Due' and OTDStatus='NotOK' 

update #tmp set Result='Potential with ETA ,in OTD' where OTDType='Potential' and OTDStatus='CanShip'
update #tmp set Result='Potential with ETA ,Out of OTD' where OTDType='Potential' and OTDStatus='LateShip'
update #tmp set Result='Potential w/o ETA' where OTDType='Potential' and OTDStatus='NotOK' 
update #tmp set Result='Dock' where OTDStatus='Dock' 

-----(2016/05/16) Manual update-----Cannot find the rule
update #tmp set Result='Past Due with ETA ,Out of OTD' where Result like 'Past Due w/o ETA%' and not EST_Date=''


----Remove old Escalation .
update #tmp set Escalation='Y' where Escalation like 'Y 2%'
update #tmp set Escalation='' where not Escalation like 'Y 2%'--Escalation like '%CSO%'


---(2024/09/03) Modify EMEACSO  to Hot List
delete from EMEACSO$ where CPQPN is null
-----(2017/12/07) Update CSO to Escalation
update #tmp set Escalation=/*rtrim(a.Escalation)+' '+*/b.CSO from #tmp a,
(
select Site,CPQPN,CSO=convert(char(10),[CSODate],111)+'_HotList ('+convert(varchar(10),Qty)+')' from EMEACSO$ 
) as b where a.CPQNo=b.CPQPN and  a.Site=b.Site

----(2023/09/03) Mark to run Hot List--------------------------------------------------------------------------------------------
/*
----(2018/01/15) modify CSO structure
/*
drop table #tmpCSO
drop table #CSO
drop table #CSOResult
*/
create table #CSO(iid int identity(1,1),Site varchar(20),CPQPN varchar(20),Qty int,CSODate varchar(20))
create table #CSOResult(Site varchar(20),CPQPN varchar(20),Qty int,CSODate varchar(100))

select Site,CPQPN,Qty=count(*) into #tmpCSO from EMEACSO$ where Qty>0 group by Site,CPQPN

insert #CSO
    select Site,CPQPN,Qty,CSODate from EMEACSO$ where CPQPN in (select CPQPN from #tmpCSO where Qty>1) and Qty>0 order by CPQPN
 
insert #CSOResult
    select Site,CPQPN,Qty,CSODate+'*'+convert(varchar(10),Qty) from EMEACSO$ where CPQPN in (select CPQPN from #tmpCSO where Qty=1) and Qty>0 order by CPQPN
    
insert #CSOResult    
   select Site,CPQPN,sum(Qty),'' from #CSO where CPQPN in (select CPQPN from #tmpCSO where Qty>1) and Qty>0 group by Site,CPQPN

declare @a int
declare @b int

select @a=min(iid) from #CSO
select @b=max(iid) from #CSO

while @a<=@b
begin
    update #CSOResult set CSODate=rtrim(a.CSODate)+b.CSODate+'*'+convert(varchar(10),b.Qty) from #CSOResult a,#CSO b where a.Site=b.Site and a.CPQPN=b.CPQPN and b.iid=@a
    select @a=@a+1
end

--select * from #CSOResult

-----(2017/12/12) Update No SO CSO to Escalation
update #tmp set Escalation=rtrim(a.Escalation)+' CSO_'+b.CSO from #tmp a,
(
select Site,CPQPN,CSO=CSODate from #CSOResult
) as b where a.Site=b.Site and a.CPQNo=b.CPQPN and not Escalation like '%CSO%' --and Site='HP-FCH2'

*/
---(2024/09/03) Modify EMEACSO  to Hot List------------------------------------------------------------------

delete from OldOPO where ReportDate<=convert(char(10),dateadd(dd,-90,getdate()),111)
--select * from #tmp where iid in ('75','53')


-----(2019/01/11) Add BTB_CSO
update #tmp set Escalation=rtrim(a.Escalation)+' BTB_CSO_'+b.CSO from #tmp a,
(
select distinct Site,PO,CPQPN,Qty,CSO=rtrim(CSODate)+'*'+convert(varchar(10),Qty) from BTB$
) as b where a.PO=b.PO and a.Site=b.Site and a.CPQNo=b.CPQPN and not Escalation like '%BTB_CSO%' --and Site='HP-FCH2'

insert OldOPO
   select *  from #tmp 


--select MType,LateThanNeedShipDate,OTDType,OTDStatus,count(*) from #tmp group by MType,LateThanNeedShipDate,OTDType,OTDStatus order by MType,OTDType,OTDStatus
-----Summary 
select a.*,b.Total,Rate=convert(decimal(8,2),convert(float,Items)/convert(float,Total)*100) from
(select MType,Result,Items=count(*) from #tmp group by MType,Result) as a left join
(select MType,Total=count(*)from #tmp group by MType) as b on a.MType=b.MType order by a.MType,a.Result




-----Summary 
select * from
( 
select MType='Dock',Items=count(*),Qty=sum(DockQty) from #tmp where Result like 'Dock%'
union
select MType='Past Due with ETA',count(*),sum(OpenQty) from #tmp where Result like 'Past Due with ETA%'
union
select MType='Past Due w/o ETA',count(*),sum(OpenQty) from #tmp where Result like 'Past Due w/o ETA%'
union
select MType='Potential with ETA before OTD',count(*),sum(OpenQty) from #tmp where Result like 'Potential with ETA%' and LateThanNeedShipDate='N'
union
select MType='Potential with ETA After OTD',count(*),sum(OpenQty) from #tmp where Result like 'Potential with ETA%' and LateThanNeedShipDate='Y'
union
select MType='Potential w/o ETA',count(*),sum(OpenQty) from #tmp where Result like 'Potential w/o ETA%'
) as a




select *,Old_ESTDate='          ',Diff='          ' into #result from #tmp where ReportDate=convert(char(10),getdate(),111)

update #result set Old_ESTDate=isnull(b.EST_Date,'') from #result a,OldOPO b 
where a.SO=b.SO and a.IECPN=b.IECPN and a.POItem=b.POItem 
and b.ReportDate=(select max(ReportDate) from OldOPO where ReportDate<convert(char(10),getdate(),111))



update #result set Diff='New' where (Old_ESTDate='' and EST_Date<>'') 
update #result set Diff='Need Check' where (Old_ESTDate<>'' and EST_Date='') 
update #result set Diff='' where (Old_ESTDate='' and EST_Date='') 
update #result set Diff=datediff(dd,Old_ESTDate,EST_Date) where (Old_ESTDate<>'' and EST_Date<>'') 


---Detail
delete from OPS_OPO_new where ReportDate<=convert(char(10),dateadd(dd,-90,getdate()),111) and Customer='H2'
delete from OPS_OPO_new where ReportDate=convert(char(10),getdate(),111) and Customer='H2'

insert OPS_OPO_new
  select * from #result

--select * from OPS_OPO_new where ReportDate=(select max(ReportDate) from OPS_OPO_new) and PType='ZN' order by PIC,Site

create table #OM(pid int identity(1,1),iid int,Remark varchar(500))
create table #OMFinal(iid int,Remark varchar(5000))

insert #OM
select iid,Remar=rtrim(ShortagePN)+'( '+MaterialETA+' ) : '+Remark from OPS_OM where iid in(
select iid from OPS_OPO_new where ReportDate=(select max(ReportDate) from OPS_OPO_new) and Shortage='Y') and Customer='H2' and ReportDate=(select max(ReportDate) from OPS_OPO_new)

insert #OMFinal
  select distinct iid,'' from #OM


declare @i int
declare @j int
select @i=min(pid) from #OM
select @j=max(pid) from #OM

while @i<=@j
begin
   update #OMFinal set Remark=rtrim(a.Remark)+'//'+b.Remark from #OMFinal a,#OM b where a.iid=b.iid and b.pid=@i
   select @i=@i+1
end

select a.*,OPOPIC=isnull(b.PIC,''),o_PO=a.PO into #OPOresult from 
(
select a.*,RSD=left(b.SO_ReqDate,4)+'/'+substring(b.SO_ReqDate,5,2)+'/'+substring(b.SO_ReqDate,7,2) from
(
select a.*,MaterialRemark=isnull(b.Remark,'') from
(select * from OPS_OPO_new where ReportDate=(select max(ReportDate) from OPS_OPO_new) and Customer='H2') as a left join
(select * from #OMFinal) as b on a.iid=b.iid --order by PIC,Site
) as a left join
(select distinct SO,IECPN,SO_ReqDate from Service_APD) as b
on a.SO=b.SO and a.IECPN=b.IECPN 
) as a left join
(select * from IECPN82PIC where not IECPN82 like '6%') as b on substring(a.IECPN,8,2)=b.IECPN82 and a.Model_Status=b.MPEOL
and substring(a.IECPN,2,1)='F'
--order by PIC,Site


update #OPOresult set OPOPIC=b.PIC from #OPOresult a,IECPN82PIC b where left(a.IECPN,4)=b.IECPN82 and a.Model_Status=b.MPEOL 
and left(IECPN,1)='6' and OPOPIC=''


update #OPOresult set OPOPIC=b.PIC from #OPOresult a,MType b where a.ProductFamily=b.Material_group
and left(IECPN,1)='6' and OPOPIC=''

---(2018/11/02) Add mark for NA Tariff POs
update #OPOresult set PO=rtrim(PO)+'_(Tariff)' where PO in ('14100777','14100780','14100781','14100782','14100783',
'14100784','14100785','14100786','14100789','14110504','14110604','14129933','14135339','15617974','15618362','15623512')

---(2018/12/25) Add mark for India BTB
update #OPOresult set PO=rtrim(PO)+'_(IND BTB)' where PO in ('14149391','14149393')

---(2018/12/25) Add mark for India BTB 2nd
update #OPOresult set PO=rtrim(PO)+'_(IND BTB 2)' where PO in ('14152531','14152542','14152549','14152813','14152815','14152822','14152823','14152976','14154791','14155764','14155766','14162279','14162285')

---(2018/12/26) Add mark for India BTB 3nd
update #OPOresult set PO=rtrim(PO)+'_(IND BTB 3)' where PO in ('14158402','14158411','14158413','14165677','14165633','14165649','14165659','14165677','14181963')

---(2019/01/03) Add mark for NA Tariff 2
update #OPOresult set PO=rtrim(PO)+'_(Tariff 2)' where PO in ('14158415','14158433','14162808','14162811')

---(2019/01/03) Add mark for India BTB 4nd
update #OPOresult set PO=rtrim(PO)+'_(IND BTB 4)' where PO in ('114162175','14162180','14165705','14165724','14165738','14170303','14170305','14165705','14165724','14162175','14177162','14177163')

---(2019/01/15) Add mark for India BTB 5
update #OPOresult set PO=rtrim(PO)+'_(IND BTB 5)' where PO in ('14170309','14170312','14170313','14170315','14175564','14175578','14175584','14175586','14175588','14175592','14175606','14175609','14175611')

---(2019/01/25) Add mark for India BTB 5
update #OPOresult set PO=rtrim(PO)+'_(IND BTB 6)' where PO in ('14182384','14182389','14182883','14182888','14182891','14182893','14185211','14185218','14185224','14185211','14185218','14185224','14186499','14186501','14188677','14197041','14198790','14200347','14202091')

---(2019/03/06) add SMS SMSCQ Containment 
update #OPOresult set PO=rtrim(a.PO)+'_CS' from #OPOresult a,SMSCM b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='CS'
---(2019/03/25) add SMS SMSCQ Containment 2
update #OPOresult set PO=rtrim(a.PO)+'_CS2' from #OPOresult a,SMSCM b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='CS2'
---(2019/04/23) add SMS SMSCQ Containment 3
update #OPOresult set PO=rtrim(a.PO)+'_CS3' from #OPOresult a,SMSCM b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='CS3'
---(2019/05/24) add SMS SMSCQ Containment 4
update #OPOresult set PO=rtrim(a.PO)+'_CS4' from #OPOresult a,SMSCM b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='CS4'
---(2019/06/20) add SMS SMSCQ Containment 5
update #OPOresult set PO=rtrim(a.PO)+'_CS5' from #OPOresult a,SMSCM b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='CS5'
---(2019/07/26) add SMS SMSCQ Containment 6
update #OPOresult set PO=rtrim(a.PO)+'_CS6' from #OPOresult a,SMSCM b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='CS6'

---(2019/08/27) add SMS SMSCQ Priority PO 
update #OPOresult set PO=rtrim(a.PO)+'_P' from #OPOresult a,SMSCM b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='P'


---(2023/05/15) add SMS BP (BigBuy POs) 
update #OPOresult set PO=rtrim(a.PO)+'_BP' from #OPOresult a,SMSBP b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='BP'


---(2024/06/04) add EMEAKBQ
update #OPOresult set PO=rtrim(a.PO)+'_BR_BO' from #OPOresult a,SMSBP b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='BR_BO'


---(2025/10/30) add BR_BO
update #OPOresult set PO=rtrim(a.PO)+'_EMEAKBQ' from #OPOresult a,SMSBP b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='EMEAKBQ'


---(2023/06/29) add SMS TR (Tariff POs) 
update #OPOresult set PO=rtrim(a.PO)+'_TR' from #OPOresult a,SMSTR b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='TR'

---(2025/01/07) add SMS TR2 (Tariff POs) 
update #OPOresult set PO=rtrim(a.PO)+'_TR2' from #OPOresult a,SMSTR b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='TR2'

---(2025/01/07) add SMS TR12 (Tariff POs) 
update #OPOresult set PO=rtrim(a.PO)+'_TR12' from #OPOresult a,SMSTR b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='TR12'


---(2024/01/16) add SMS APJSLA
update #OPOresult set PO=rtrim(a.PO)+'_APJSLA' from #OPOresult a,SMSTR b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='APJSLA'


---(2023/07/24) add SMS APIRR
update #OPOresult set PO=rtrim(a.PO)+'_APIRR' from #OPOresult a,SMSAPIRR b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='APIRR'


---(2023/08/04) add SMS Top 300
update #OPOresult set PO=rtrim(a.PO)+'_APJ300' from #OPOresult a,SMSTop300 b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='APJ300'

---(2024/12/26) Add charindex to ensure all "_" items can be added
---(2020/02/19) add SMS SMSCQ Priority PO 
update #OPOresult set PO=rtrim(a.PO)+'_P1' from #OPOresult a,SMSCM b where left(a.PO,case when charindex('_',a.PO)>0 then charindex('_',a.PO)-1 else 20 end)=b.PO and a.CPQNo=b.OSSPPN and b.PT='P1'
update #OPOresult set PO=rtrim(a.PO)+'_P2' from #OPOresult a,SMSCM b where left(a.PO,case when charindex('_',a.PO)>0 then charindex('_',a.PO)-1 else 20 end)=b.PO and a.CPQNo=b.OSSPPN and b.PT='P2'
update #OPOresult set PO=rtrim(a.PO)+'_P3' from #OPOresult a,SMSCM b where left(a.PO,case when charindex('_',a.PO)>0 then charindex('_',a.PO)-1 else 20 end)=b.PO and a.CPQNo=b.OSSPPN and b.PT='P3'
update #OPOresult set PO=rtrim(a.PO)+'_P4' from #OPOresult a,SMSCM b where left(a.PO,case when charindex('_',a.PO)>0 then charindex('_',a.PO)-1 else 20 end)=b.PO and a.CPQNo=b.OSSPPN and b.PT='P4'

----(2020/12/23) from HP Critical Items
update #OPOresult set PO=rtrim(a.PO)+'_P0' from #OPOresult a,SMSCM b where left(a.PO,case when charindex('_',a.PO)>0 then charindex('_',a.PO)-1 else 20 end)=b.PO and a.CPQNo=b.OSSPPN and b.PT='P0'
update #OPOresult set PO=rtrim(a.PO)+'_P5' from #OPOresult a,SMSCM b where left(a.PO,case when charindex('_',a.PO)>0 then charindex('_',a.PO)-1 else 20 end)=b.PO and a.CPQNo=b.OSSPPN and b.PT='P5'

---(2024/06/11) Add EXPRESS
update #OPOresult set PO=rtrim(a.PO)+'_EXPRESS' from #OPOresult a,SMSCM b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='EXPRESS'

---(2019/08/29) add SMS SMSCQ BTB
update #OPOresult set PO=rtrim(a.PO)+'_BTB' from #OPOresult a,SMSCM b where a.PO=b.PO and a.CPQNo=b.OSSPPN and b.PT='BTB'


---(2020/07/22) add Glue & CR info to Model
update #OPOresult set ProductFamily=rtrim(a.ProductFamily)+'_'+b.Glue+b.CR from #OPOresult a,PanelAssyType$ b where a.IECPN=b.IECPN

---(2022/04/29) add OFilm Info in OSSPPN
update #OPOresult set CPQNo=rtrim(a.CPQNo)+'_*' from #OPOresult a,OFilmPN$ b where a.CPQNo=b.OSSPPN

---(2022/10/04) add POCancel floaf Info in PO
update #OPOresult set PO=rtrim(a.PO)+'_'+rtrim(b.Disposition) from #OPOresult a,HPPOCancelPP b where a.o_PO=b.PO and replace(a.CPQNo,'_*','')=b.HPPN 




select * from #OPOresult

-------(2019/01/25) Get HP DN Date & Qty
select PT,DNDate=IES_DNPGI,Qty=sum(DNQty) from (
select PT=case 
when substring(IECPN,8,2)='MB' then 'MB'
when substring(IECPN,8,2)='LM' then 'LCM'
when substring(IECPN,8,2)='KB' then 'K/B'
when substring(IECPN,8,2)='BD' then 'Whole Unit'
when substring(IECPN,8,2)='CK' then 'Cable'
when substring(IECPN,8,2) in ('TP','LB','BT','LT') then 'Cases'
when left(IECPN,1)='6' then 'Raw Material'
else 'Others' end,DNQty,IES_DNPGI from OPS_OPO_new where ReportDate=convert(char(10),getdate(),111)  and Customer='H2' and not IES_DN='' and DNQty>0
) as a group by PT,IES_DNPGI


-----------------------------------------------------------
----------------------------------------------------------
----------------------------------------------------------
-----------------------------------------------------------
----------------------------------------------------------
----------------------------------------------------------
-----------------------------------------------------------
----------------------------------------------------------
---------------------------------------------------------------------------------------------------------------------
----------------------------------------------------------
----------------------------------------------------------
-------OPOR Report Need to Run every Friday---------------
-----(2018/05/04 Change to all instead of separate NB/DT/Mobile-----------------------------------------------------
------All --------------------------------------------------
select * from (
select MType='Risk Qty',Qty=sum(OpenQty) from #tmp where not PO_Type like 'NPI%'  
and Result in ('Potential w/o ETA','Potential with ETA ,Out of OTD')
and NeedShipDate<=convert(char(10),dateadd(day,14,getdate()),111)
union
select MType='Risk Item',count(*) from #tmp where not PO_Type like 'NPI%'  
and Result in ('Potential w/o ETA','Potential with ETA ,Out of OTD')
and NeedShipDate<=convert(char(10),dateadd(day,14,getdate()),111)
union
select MType='Past Due Qty',sum(OpenQty) from #tmp where not PO_Type like 'NPI%' and Result like 'Past Due%'
union
select MType='Past Due Item',count(*) from #tmp where not PO_Type like 'NPI%' and Result like 'Past Due%'
union
select MType='Open Qty',sum(OpenQty) from #tmp where not PO_Type like 'NPI%' 
union
select MType='Open Item',count(*) from #tmp where not PO_Type like 'NPI%'
) as a
/*
---NB
---(2016/06/03) Risk only need to get the Need Ship Date in 2 weeks .
select * from (
select MType='Risk Qty',Qty=sum(OpenQty) from #tmp where not PO_Type like 'NPI%' and MType in ('HP PORTABLE','Raw Material') 
and Result in ('Potential w/o ETA','Potential with ETA ,Out of OTD')
and NeedShipDate<=convert(char(10),dateadd(day,14,getdate()),111)
union
select MType='Risk Item',count(*) from #tmp where not PO_Type like 'NPI%' and MType in ('HP PORTABLE','Raw Material') 
and Result in ('Potential w/o ETA','Potential with ETA ,Out of OTD')
and NeedShipDate<=convert(char(10),dateadd(day,14,getdate()),111)
union
select MType='Past Due Qty',sum(OpenQty) from #tmp where not PO_Type like 'NPI%' and MType in ('HP PORTABLE','Raw Material') and Result like 'Past Due%'
union
select MType='Past Due Item',count(*) from #tmp where not PO_Type like 'NPI%' and MType in ('HP PORTABLE','Raw Material') and Result like 'Past Due%'
union
select MType='Open Qty',sum(OpenQty) from #tmp where not PO_Type like 'NPI%' and MType in ('HP PORTABLE','Raw Material') 
union
select MType='Open Item',count(*) from #tmp where not PO_Type like 'NPI%' and MType in ('HP PORTABLE','Raw Material')
) as a

---MOBILE
select * from (
select MType='Risk Qty',Qty=sum(OpenQty) from #tmp where not PO_Type like 'NPI%' and MType='HP MOBILE' 
and Result in ('Potential w/o ETA','Potential with ETA ,Out of OTD') 
and NeedShipDate<=convert(char(10),dateadd(day,14,getdate()),111)
union
select MType='Risk Item',count(*) from #tmp where not PO_Type like 'NPI%' and MType='HP MOBILE' 
and Result in ('Potential w/o ETA','Potential with ETA ,Out of OTD')
and NeedShipDate<=convert(char(10),dateadd(day,14,getdate()),111)
union
select MType='Past Due Qty',sum(OpenQty) from #tmp where not PO_Type like 'NPI%' and MType='HP MOBILE' and Result like 'Past Due%'
union
select MType='Past Due Item',count(*) from #tmp where not PO_Type like 'NPI%' and MType='HP MOBILE' and Result like 'Past Due%'
union
select MType='Open Qty',sum(OpenQty) from #tmp where not PO_Type like 'NPI%' and MType='HP MOBILE'
union
select MType='Open Item',count(*) from #tmp where MType='HP MOBILE'
) as a


---DT----(2016/11/14) ---Add AIO .
select * from (
select MType='Risk Qty',Qty=sum(OpenQty) from #tmp where not PO_Type like 'NPI%' and MType in ('HP DT','HP AIO')
and Result in ('Potential w/o ETA','Potential with ETA ,Out of OTD')
and NeedShipDate<=convert(char(10),dateadd(day,14,getdate()),111)
union
select MType='Risk Item',count(*) from #tmp where not PO_Type like 'NPI%' and MType in ('HP DT','HP AIO')
 and Result in ('Potential w/o ETA','Potential with ETA ,Out of OTD')
and NeedShipDate<=convert(char(10),dateadd(day,14,getdate()),111)
union
select MType='Past Due Qty',sum(OpenQty) from #tmp where not PO_Type like 'NPI%' and MType in ('HP DT','HP AIO') and Result like 'Past Due%'
union
select MType='Past Due Item',count(*) from #tmp where not PO_Type like 'NPI%' and MType in ('HP DT','HP AIO') and Result like 'Past Due%'
union
select MType='Open Qty',sum(OpenQty) from #tmp where not PO_Type like 'NPI%' and MType in ('HP DT','HP AIO')
union
select MType='Open Item',count(*) from #tmp where not PO_Type like 'NPI%' and MType in ('HP DT','HP AIO')
) as a
----
*/
----------------------------------------------------------
----------------------------------------------------------
-------OPOR Report Need to Run every Friday- End--------------
----------------------------------------------------------
--------------------------------------------------------

