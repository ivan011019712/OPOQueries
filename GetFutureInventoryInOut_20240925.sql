/*
drop table #DN
drop table #OPO
drop table #Ship
drop table #tmp_SH_ETA
drop table #SH_ETA
*/

declare @USTW int
declare @RMBTW int
Declare @EndDate char(10)
Declare @Dt char(10)

select @Dt=convert(char(10),getdate(),111)
select @EndDate=convert(char(10),dateadd(dd,7,getdate()),111)
select @USTW='28'
select @RMBTW='4'

select @Dt,@EndDate

select Customer='          ',MSite='                    ',
Site,PO,SO,Item,IECPN,PO_Date,Qty856,IES_DN,IES_DNPGI into #DN from Service_APD 
where  PndGIDate='0000/00/00' and not IES_DNPGI='0000/00/00' and not SO like '19%'


--Site in (Select distinct MSite,Customer,ZS92Site from SiteMapping where not ZS92Site='') and
update #DN set Customer=b.Customer,MSite=b.Site from #DN a,
(select * from OPS_OPO where  ReportDate=convert(char(10),getdate(),111)) as b where a.PO=b.PO and a.IECPN=b.IECPN
--(Select distinct MSite,Customer,ZS92Site from SiteMapping where not ZS92Site='') as b where a.Site=b.ZS92Site

update #DN set Customer='HP',MSite='DCE' where Site='HP SCHWEIZ GMBH' and PO='HPI283059'
update #DN set Customer='HP',MSite='DCE' where Site='HP INC UK LIMIT' and PO='HPI283665'
update #DN set Customer='HP',MSite='DCE' where Site='IN CARE OF HP' and PO='PR8180297'
update #DN set Customer='HP',MSite='DCE' where Site='HP PPS AUSTRALI' and PO='HPI289895'
update #DN set Customer='HP',MSite='FLEX EMEA' where Site='HAIO-FRUFG' and PO='200183409'
update #DN set Customer='HP',MSite='DCE' where Site='HEWLETT PACKARD' and PO='HPI297395'
update #DN set Customer='HP',MSite='DCE' where Site='IN CARE OF HP' and Customer='' and MSite=''
update #DN set Customer='HP',MSite='DCE' where Site='IN CARE OF HP.' and Customer='' and MSite=''
update #DN set Customer='HP',MSite='DCE' where Site='HP DEUTSCHLAND'  and PO='HPI305462'
update #DN set Customer='HP',MSite='DCE' where Site='HP JAPAN INC.'  and PO='HPI303536'
update #DN set Customer='HP',MSite='DCE' where Site='HP INC'  and PO='HPI575586'
update #DN set Customer='HP',MSite='CZ' where Site='L6-CZ1'  and PO='20171117-CZ'
update #DN set Customer='HP',MSite='DCE' where PO='HPI471045' and IECPN='PFEF01CMB232'
update #DN set Customer='HP',MSite='DCE' where Site='HP INC'  and PO='No. PR8821408'

delete #DN where Site='IOT-IHS' and SO='2000497362'
delete #DN where Site='D-FCTS' and SO='1108130567'
delete #DN where Site='D-FWWT2' and SO='1108427721'
delete #DN where Site='D-FWWT2' and SO='1108836588' and IECPN='6060B0330401'
delete #DN where Site='IOT-ISV' and SO='2000497362'
delete #DN where Site='IOT-ISV' and SO='2000498279'
delete #DN where Site='IOT-ISV' and SO='2000498546'
delete #DN where Site='24625' and SO='1110150764' and IECPN='6037B0153803'
delete #DN where Site='12607' and SO='1110385230' and IECPN='LF2617ALB002'
delete #DN where Site='' and PO in ('F4009896000','F4009900000','F4009898000','F4009899000','F4009897000')
delete #DN where Site='24625' and  Customer=''
delete #DN where Site='135406' and  Customer='' and PO_Date in ('2020/03/02','2020/05/08','2020/05/11')
delete #DN where Site='15883' and  Customer='' and PO_Date in ('2020/06/01','2020/03/06','2020/03/09','2020/03/11','2020/03/12')
delete #DN where Site='32745' and  Customer='' and PO_Date in ('2020/03/31')
delete #DN where Site='170338' and  Customer='' and PO_Date in ('2020/07/15')
delete #DN where Site='209016' and  Customer='' and PO_Date in ('2020/08/20')
delete #DN where Site='248967' and  Customer='' 
delete #DN where Site='245326' and  Customer=''
delete #DN where Site in ('30019','11149','213758','212827','12398','226139','17923','215119','165483','259941','264376','169255')and  Customer=''
delete from #DN where SO='1200043490'
delete from #DN where SO in ('1200044184','1200043926','1200043917','1200043849','1200043807','1200043794','1200043792','1200043641','2000506359','1200043596','1200043593','2000505646','2000505647','1111686374','1200043578','1200043784',
'1200043784')
delete #DN where Customer='' and SO in ('2000525617','1200044622','1200044613','1200044585','2000523551','1200044571','1200044552','1200044541','2000521924','2000521134','1200044428','1200044422','2000518806','1200044393','1200044385','1200044358','1200044330','1200044324','1200044296','1200044314','1200044297','1200044311','1200044295','1200044282','1200044283','2000514986','1200044251','1200044276','1200044243','1200044252','1200044240','1200044238','2000514049','1200044221','1200044228','1200044208','1200044165','2000512659','2000512622','2000512073','2000511999','1200044128','1200044117','1200043942','1110854177','1110869927','1110926321','1200044032','1200043981','1200043980','1200044030')

--select * from #DN where Customer=''
-----Update Site for FJ
update #DN set Customer='FJ',MSite='FJ EDI' where Customer='' and Site like '7054%'
--select * from #DN where Customer=''
if (select count(*) from #DN where Customer='')>0
begin
    select 'Find Customer is blank ,Pleas check'
    return
end 

update #DN set Item=b.IPCSOItem from #DN a,ZM57$ b where a.SO=b.IECSO and a.Item=b.IECSPItem 
update #DN set Item=b.IPCSOItem from #DN a,ZM57 b where a.SO=b.IECSO and a.Item=b.IECSPItem 


---(2016/03/18) ---Add DN information
select a.*,IES_DN=isnull(b.IES_DN,'') into #OPO from 
(
select a.*,IES_DNPGI=isnull(IES_DNPGI,'') from
(
select Customer,iid,Site,PIC,PO,SO,IECPO,POItem,CPQNo,IECPN,ProductFamily,POVendor,POReceiveDate,PO_Type,
Model_Status,MP,FCST_Status,PType,NeedShipDate,POQty,DNQty,DockQty,ShipQty,
OpenQty,RMA_FG,CurrentDone,FutureDone,FutureFirstSupportDate,Shortage,SameMonthFCST,OPOR,Escalation,EscalationDate,Remark from OPS_OPO 
where ReportDate=convert(char(10),getdate(),111) 
) as a left join
(select SO,Item,IECPN,PO_Date,IES_DNPGI=max(IES_DNPGI) from #DN group by SO,Item,IECPN,PO_Date) as b on 
a.SO=b.SO and a.IECPN=b.IECPN and a.POItem=b.Item and a.POReceiveDate=b.PO_Date
) as a left join
#DN as b on a.SO=b.SO and a.IECPN=b.IECPN and a.POItem=b.Item and a.POReceiveDate=b.PO_Date and a.IES_DNPGI=b.IES_DNPGI
order by PIC,Site,iid

---------Output (FRU Ship)
-------
--drop table #Ship
create table #Ship(Customer varchar(20),PType varchar(20),IECPN varchar(20),Qty float,Price float,Amount float)

insert #Ship
select Customer,PType,IECPN,Qty=sum(Qty),0,0 from (
select Customer,PType='Dock',IECPN,Qty=sum(DockQty) from #OPO where OPOR='Wait for Pick up' group by Customer,IECPN
union
select Customer,PType='DN',IECPN,sum(DNQty) from #OPO where IES_DNPGI<=@EndDate and not IES_DNPGI='' and not OPOR='Wait for Pick up' group by Customer,IECPN
union
select Customer,PType='IN',MatNo,sum(Qty) from Ivan_CurrentINV where INVDate>@Dt and INVDate<=@EndDate group by Customer,MatNo
) as a group by Customer,PType,IECPN


--select Customer,PType='IN',MatNo,sum(Qty),0,0 from Ivan_CurrentINV where INVDate>'2017/03/28' and INVDate<='2018/04/01' group by Customer,MatNo

---------(2017/03/21) Get HP BL material ,too.
------------------------------------------------------------------
create table #tmp_SH_ETA(iid int identity(1,1),Material varchar(20),MaterialETA varchar(200))
create table #SH_ETA(Material varchar(20),MaterialETA varchar(100),ETA char(20),Qty int)


--Only Multiple ETA need to count
insert #tmp_SH_ETA
    select  Material,MaterialETA=replace(upper([Material ETA]),'K','00') from BLMaterial$ 
    where not [Material ETA] in ('_','') and [Material ETA] like '%,%'
    
---Add single item to #SH_ETA
insert #SH_ETA
    select  Material,MaterialETA=replace(upper([Material ETA]),'K','00'),'','' from BLMaterial$ 
    where not [Material ETA] in ('_','') and not [Material ETA] like '%,%'
    
declare @a int
declare @b int
select @a=min(iid) from #tmp_SH_ETA
select @b=max(iid) from #tmp_SH_ETA

while @a<=@b
begin
     while (select charindex(',',MaterialETA) from #tmp_SH_ETA where iid=@a)>0
     begin
         insert #SH_ETA
            select distinct Material,left(MaterialETA,charindex(',',MaterialETA)-1),'','' from #tmp_SH_ETA where iid=@a
         update #tmp_SH_ETA set MaterialETA=substring(MaterialETA,charindex(',',MaterialETA)+1,1000) where iid=@a
         
         insert #SH_ETA
            select distinct Material,rtrim(MaterialETA),'','' from #tmp_SH_ETA where iid=@a and not MaterialETA like '%,%'
     end
     select @a=@a+1 
end

--update ETA /Qty
update #SH_ETA set ETA=left(MaterialETA,charindex('*',MaterialETA)-1),Qty=convert(int,rtrim(substring(MaterialETA,charindex('*',MaterialETA)+1,1000)))

---Format Date
update #SH_ETA set ETA=convert(char(10),convert(datetime,rtrim(ETA)+' 00:00'),111)


/*
select Material,ETA=left(MaterialETA,charindex('*',MaterialETA)-1),
Qty=convert(int,rtrim(substring(MaterialETA,charindex('*',MaterialETA)+1,1000))) from #SH_ETA

select * from BLMaterial$ order by [Material ETA] 
select * from BLMaterial$  where [Material ETA] like '%2024/1/19%'
update BLMaterial$ set [Material ETA]='2024/1/19*922 ' where [Material ETA] like  '%2024/1/19%'
*/
------------------------------------------------------------------------------
-----------------------------------------------------------------------------
insert #Ship
select Customer='HP','BLIN',Material,sum(Qty),0,0 from #SH_ETA where ETA>@Dt and ETA<=@EndDate group by Material


--Get PR00 Price
update #Ship set Price=b.Price/Per*@USTW,Amount=Qty*b.Price/Per*@USTW from #Ship a,PR00 b where a.IECPN=b.IECPN and PType in ('Dock','DN')--and not a.IECPN like '6%'

---Get Raw Material Price
update #Ship set Price=b.StdPrice/b.PriceUnit*@RMBTW,Amount=a.Qty*b.StdPrice/b.PriceUnit*@RMBTW from #Ship a,t_download_matmas_CP60DW b 
where left(a.IECPN,12)=b.Material and (a.IECPN like '6%' or a.IECPN like '1%') and Amount='0' and Customer='HP'

update #Ship set Price=b.StdPrice/b.PriceUnit*@RMBTW,Amount=a.Qty*b.StdPrice/b.PriceUnit*@RMBTW from #Ship a,t_download_matmas_CP62DW b 
where left(a.IECPN,12)=b.Material and (a.IECPN like '6%' or a.IECPN like '1%') and Amount='0' and Customer in ('TOSHIBA','FJ','DYNABOOK')

update #Ship set Price=b.StdPrice/b.PriceUnit*@RMBTW,Amount=a.Qty*b.StdPrice/b.PriceUnit*@RMBTW from #Ship a,t_download_matmas_CP65DW b 
where left(a.IECPN,12)=b.Material and (a.IECPN like '6%' or a.IECPN like '1%') and Amount='0' and Customer='TINY'

update #Ship set Price=b.StdPrice/b.PriceUnit*@RMBTW,Amount=a.Qty*b.StdPrice/b.PriceUnit*@RMBTW from #Ship a,t_download_matmas_CP69DW b 
where left(a.IECPN,12)=b.Material and (a.IECPN like '6%' or a.IECPN like '1%') and Amount='0' and Customer in ('ACER','ASUS')




----Check No Price .
--select * from #Ship where not Qty=0  and Amount=0

select Customer,DN=convert(int,sum(DN)),Dock=convert(int,sum(Dock)),
[OUT]=convert(int,sum(DN))+convert(int,sum(Dock)),[PIN]=convert(int,sum([IN])),BLIN=convert(int,sum(BLIN)),
[IN]=convert(int,sum([IN]))+convert(int,sum(BLIN)),
Summary=convert(int,sum([IN]))+convert(int,sum(BLIN))-convert(int,sum(DN))-convert(int,sum(Dock))
 from (
select Customer,DN=sum(Amount),Dock=0,[IN]=0,BLIN=0 from #Ship where PType='DN' group by Customer --order by Customer
union
select Customer,0,Dock=sum(Amount),0,0 from #Ship where PType='Dock' group by Customer --order by Customer
union
select Customer,0,0,Dock=sum(Amount),0 from #Ship where PType='IN' group by Customer --order by Customer
union
select Customer,0,0,0,Dock=sum(Amount) from #Ship where PType='BLIN' group by Customer --order by Customer
) as a group by Customer


select * from #Ship order by PType,Amount desc



