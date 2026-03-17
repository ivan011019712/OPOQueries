/*
select * from Material$  where [Material ETA] like '%10/22*960%'
select * from Material$ where Material='6037B0146601*'
6037B0152601*	20	2020/6/24*40¡@
1310A2635703
1310A2635704
1310A2727301
1310A2783205*

update Material$ set [Material ETA]='2025/8/25*365,2025/10/22*960,2025/11/5*460' where   [Material ETA] like '%10/22*960%'
update Material$ set [Material ETA]='2012/11/9*300,2012/11/17*300' where Material='6034B0008102' and [Material ETA]='2012/11/9*300,201211/17*300'


select * from TSB_Material$ where Material='6034B0009001'
update TSB_Material$ set [Material ETA]='2015/4/17*800,2015/4/21*2400,2015/4/30*400' where Material='6034B0009001'
*/
---(2016/12/28) Special ....Alorica DT doesn't appear ,manual add....
--select * from Service_APD where SO='1107906215' 
--update Service_APD set Site='HP-FDALO1' where SO='1107906215' 
------Check Inventory date
select Cdt,WIP=count(*) from IEC1_EISDW.INVENTORY.dbo.WIP group by Cdt
select Cdt,WMALL=count(*) from IEC1_EISDW.INVENTORY.dbo.WMALL group by Cdt


select max(Cdt) from WIP_WHALL_TD
select count(*) from WIP_WHALL_TD
select 'HP',count(*) from icc_r_shortage_data
select 'HP-FP',count(*) from [IEC1-FP-01].ICCFRUD.dbo.r_shortage_data

select 'H2',count(*) from icc_r_shortage_data_npi
select 'H2-FP',count(*) from [IEC1-FP-01].ICCFRUD.dbo.r_shortage_data_npi

select 'HP_ITH',count(*) from ith_r_shortage_data
select 'HP_ITH-FP',count(*) from [IEC1-FP-01].TH02FRUD.dbo.r_shortage_data

select 'ASUS',count(*) from ASUS_r_shortage_data
select 'ASUS-FP',count(*) from [IEC1-FP-02].ASUSCQNBD.dbo.r_shortage_data

select 'FJ',count(*) from FJ_r_shortage_data
select 'FJ',count(*) from [IEC1-FP-02].FJCQFRU.dbo.r_shortage_data


select 'DYNA',count(*) from DYNABOOK_r_shortage_data
select 'DYNA',count(*) from [IEC1-FP-02].DYNABOOKCQD.dbo.r_shortage_data




----(2021/11/19) Check if Invetory (WA4) & Intransit exist same PN & Qty
select * from 
(select MatNo,Qty from WIP_WHALL_TD where Plant='CP60' and SLoc='WA4' and SType='001') as a inner join
(select MatNo,Qty from WIP_WHALL_TD where Plant='CP60' and SLoc='PX11' and SType='InTransit') as b on a.MatNo=b.MatNo and a.Qty=b.Qty

select * from 
(select MatNo,Qty from WIP_WHALL_TD where Plant='CP60' and SLoc='WA4' and SType='001') as a inner join
(select MatNo,Qty from WIP_WHALL_TD where Plant='CP60' and SLoc='PX11' and SType='Unstricted ') as b on a.MatNo=b.MatNo and a.Qty=b.Qty


select * from 
(select MatNo,Qty from WIP_WHALL_TD where Plant='CP60' and SLoc='WAB' and SType='001') as a inner join
(select MatNo,Qty from WIP_WHALL_TD where Plant='CP60' and SLoc='PXN1' and SType='Unstricted ') as b on a.MatNo=b.MatNo and a.Qty=b.Qty

select * from 
(select MatNo,Qty from WIP_WHALL_TD where Plant='TH02' and SLoc='WKE' and SType='001') as a inner join
(select MatNo,Qty from WIP_WHALL_TD where Plant='TH02' and SLoc='PX21' and SType='Unstricted ') as b on a.MatNo=b.MatNo and a.Qty=b.Qty

/*
delete from WIP_WHALL_TD where Plant='CP60' and SLoc='WA4' and SType='001' and MatNo in (
     select distinct a.MatNo from 
(select MatNo,Qty from WIP_WHALL_TD where Plant='CP60' and SLoc='WA4' and SType='001') as a inner join
(select MatNo,Qty from WIP_WHALL_TD where Plant='CP60' and SLoc='PX11' and SType='InTransit') as b on a.MatNo=b.MatNo and a.Qty=b.Qty
)

delete from WIP_WHALL_TD where Plant='CP60' and SLoc='WA4' and SType='001' and MatNo in (
     select distinct a.MatNo from 
(select MatNo,Qty from WIP_WHALL_TD where Plant='CP60' and SLoc='WA4' and SType='001') as a inner join
(select MatNo,Qty from WIP_WHALL_TD where Plant='CP60' and SLoc='PX11' and SType='Unstricted') as b on a.MatNo=b.MatNo and a.Qty=b.Qty
)

delete from WIP_WHALL_TD where Plant='CP60' and SLoc='WAB' and SType='001' and MatNo in (
     select distinct a.MatNo from 
(select MatNo,Qty from WIP_WHALL_TD where Plant='CP60' and SLoc='WAB' and SType='001') as a inner join
(select MatNo,Qty from WIP_WHALL_TD where Plant='CP60' and SLoc='PXN1' and SType='Unstricted') as b on a.MatNo=b.MatNo and a.Qty=b.Qty
)

*/

/*-----------------Special item which has been obsolete in PDM, beofore fix, need to manual modify OPO
select * from Service_APD where IECPN='6051B1433701'
delete from #init_tmp_OPO where IECPN='6051B1433701'
delete from #ShortageDetail_total where ShortagePN='6051B1433701'
*/
update  Service_APD set CPQPN=replace(CPQPN,' ','') where CPQPN like ' %'

update Service_APD set CPQPN='90NR0L90-R00010' where IECPN='LFO525AMB001' and CPQPN=''
update Service_APD set CPQPN='90NR0LB0-R02010' where IECPN='LFO525BMB001' and CPQPN=''
update Service_APD set CPQPN='90NR0LB0-R01010' where IECPN='LFO525EMB001' and CPQPN=''


----(2017/06/01) Special OPO ,manual modify
update Service_APD set Site='DCE',SoldToPartner='CPQ-DCE' where SO in ('1108259298','1108251552') 

----Check Whether the Service_APD updated 
select max(Date850) from Service_APD
select count(*) from Service_APD

---Check ZSD65
select * from ZSD65 where SO in (
select distinct SO from Service_APD where Date850=(select max(Date850) from Service_APD))

select POVendor,count(*) from ZSD65 group by POVendor order by POVendor

----Check ZM57
select * from ZM57 where IECSO in (
select distinct SO from Service_APD where Date850=(select max(Date850) from Service_APD))



select distinct Site from ZM57 where IPCSO='' order by Site --and Site='DTC-FASL'
--select * from ZM57 where IPCSO='' and Site='DTC-FASL'
delete from ZM57 where IPCSO='' and Site in ('ASUS-CSC','A-FASC','DTC-FASL','HP-FSMSHUB','TSB-FRUTSP','TSB-FRURTS','TSB-FRUBIZ','D-FAPCC','FJFRU','TSB-FTISH','TSB-FTIPL')




/*
----20240229 ---fix strange data
--select * from ZM57 where IPCSO='' and  Site='ASUS-CSC' order by PODate
--select SO,IPCSO,IPCSOItem from ZSD65 where CustPO in (select PO from ZM57 where IPCSO='' and  Site='ASUS-CSC' )
update ZM57 set IPCSO=b.IPCSO,IPCSOItem=b.IPCSOItem 
 from ZM57 a,
(select CustPO,IPCSO,IPCSOItem='000'+convert(char(2),IPCSOItem) from ZSD65 where CustPO in (select PO from ZM57 where IPCSO='' and  Site='ASUS-CSC' )) as b
where a.PO=b.CustPO
*/




/*


---If the data is wrong need to clear and re-run
drop table #AP
drop table #tmp_SH_ETA
drop table #SH_ETA
delete from Ivan_ShortageMaterialUpdate where convert(char(10),Cdt,111)=convert(char(10),getdate(),111)
*/
----HP
delete from Material$ where Material is null
if exists(select * from Material$)
begin
    update Material$ set [Material ETA]='' where not [Material ETA] like '%*%'
    update Material$ set [Material ETA]=isnull([Material ETA],''),Remark=isnull(Remark,'')
    update Material$ set [Material ETA]=replace([Material ETA],';',',')
    update Material$ set [Material ETA]=replace([Material ETA],' ','')
    update Material$ set [Material ETA]=case when right([Material ETA],1)=',' then left([Material ETA],len([Material ETA])-1) else [Material ETA] end where not [Material ETA]=''
    
    delete from Ivan_ShortageMaterialUpdate where Cdt<dateadd(dd,-90,getdate()) and Customer='HP'
    delete from Ivan_ShortageMaterialUpdate where convert(char(10),Cdt,111)=convert(char(10),getdate(),111) and Customer='HP'
    insert Ivan_ShortageMaterialUpdate
         select distinct POVendor,Material,[Shortage Qty],[Material ETA],Remark,getdate(),'HP' from Material$
end

----H2
delete from H2_Material$ where Material is null
if exists(select * from H2_Material$)
begin
    update H2_Material$ set [Material ETA]='' where not [Material ETA] like '%*%'
    update H2_Material$ set [Material ETA]=isnull([Material ETA],''),Remark=isnull(Remark,'')
    update H2_Material$ set [Material ETA]=replace([Material ETA],';',',')
    update H2_Material$ set [Material ETA]=replace([Material ETA],' ','')
    update H2_Material$ set [Material ETA]=case when right([Material ETA],1)=',' then left([Material ETA],len([Material ETA])-1) else [Material ETA] end where not [Material ETA]=''
    
    delete from Ivan_ShortageMaterialUpdate where Cdt<dateadd(dd,-90,getdate()) and Customer='H2'
    delete from Ivan_ShortageMaterialUpdate where convert(char(10),Cdt,111)=convert(char(10),getdate(),111) and Customer='H2'
    insert Ivan_ShortageMaterialUpdate
         select distinct POVendor,Material,[Shortage Qty],[Material ETA],Remark,getdate(),'H2' from H2_Material$
end


----HPITH
delete from HPITH_Material$ where Material is null
if exists(select * from HPITH_Material$)
begin
    update HPITH_Material$ set [Material ETA]='' where not [Material ETA] like '%*%'
    update HPITH_Material$ set [Material ETA]=isnull([Material ETA],''),Remark=isnull(Remark,'')
    update HPITH_Material$ set [Material ETA]=replace([Material ETA],';',',')
    update HPITH_Material$ set [Material ETA]=replace([Material ETA],' ','')
    update HPITH_Material$ set [Material ETA]=case when right([Material ETA],1)=',' then left([Material ETA],len([Material ETA])-1) else [Material ETA] end where not [Material ETA]=''
    
    delete from Ivan_ShortageMaterialUpdate where Cdt<dateadd(dd,-90,getdate()) and Customer='HP_ITH'
    delete from Ivan_ShortageMaterialUpdate where convert(char(10),Cdt,111)=convert(char(10),getdate(),111) and Customer='HP_ITH'
    insert Ivan_ShortageMaterialUpdate
         select distinct POVendor,Material,[Shortage Qty],[Material ETA],Remark,getdate(),'HP_ITH' from HPITH_Material$
end


----HPIMP
delete from HPIMP_Material$ where Material is null
if exists(select * from HPIMP_Material$)
begin
    update HPIMP_Material$ set [Material ETA]='' where not [Material ETA] like '%*%'
    update HPIMP_Material$ set [Material ETA]=isnull([Material ETA],''),Remark=isnull(Remark,'')
    update HPIMP_Material$ set [Material ETA]=replace([Material ETA],';',',')
    update HPIMP_Material$ set [Material ETA]=replace([Material ETA],' ','')
    update HPIMP_Material$ set [Material ETA]=case when right([Material ETA],1)=',' then left([Material ETA],len([Material ETA])-1) else [Material ETA] end where not [Material ETA]=''
    
    delete from Ivan_ShortageMaterialUpdate where Cdt<dateadd(dd,-90,getdate()) and Customer='HP_IMP'
    delete from Ivan_ShortageMaterialUpdate where convert(char(10),Cdt,111)=convert(char(10),getdate(),111) and Customer='HP_IMP'
    insert Ivan_ShortageMaterialUpdate
         select distinct POVendor,Material,[Shortage Qty],[Material ETA],Remark,getdate(),'HP_IMP' from HPIMP_Material$
end
/*
----TSB
delete from TSB_Material$ where Material is null
if exists(select * from TSB_Material$)
begin
    update TSB_Material$ set [Material ETA]='' where not [Material ETA] like '%*%'
    update TSB_Material$ set [Material ETA]=isnull([Material ETA],''),Remark=isnull(Remark,'')
    update TSB_Material$ set [Material ETA]=replace([Material ETA],';',',')
    update TSB_Material$ set [Material ETA]=replace([Material ETA],' ','')
    update TSB_Material$ set [Material ETA]=case when right([Material ETA],1)=',' then left([Material ETA],len([Material ETA])-1) else [Material ETA] end where not [Material ETA]=''
    
    delete from Ivan_ShortageMaterialUpdate where Cdt<dateadd(dd,-90,getdate()) and Customer='TSB'
    delete from Ivan_ShortageMaterialUpdate where convert(char(10),Cdt,111)=convert(char(10),getdate(),111) and Customer='TSB'
    insert Ivan_ShortageMaterialUpdate
         select distinct POVendor,Material,[Shortage Qty],[Material ETA],Remark,getdate(),'TSB' from TSB_Material$
end
*/

----FJ
delete from FJ_Material$ where Material is null
if exists(select * from FJ_Material$)
begin
    update FJ_Material$ set [Material ETA]='' where not [Material ETA] like '%*%'
    update FJ_Material$ set [Material ETA]=isnull([Material ETA],''),Remark=isnull(Remark,'')
    update FJ_Material$ set [Material ETA]=replace([Material ETA],';',',')
    update FJ_Material$ set [Material ETA]=replace([Material ETA],' ','')
    update FJ_Material$ set [Material ETA]=case when right([Material ETA],1)=',' then left([Material ETA],len([Material ETA])-1) else [Material ETA] end where not [Material ETA]=''
    
    delete from Ivan_ShortageMaterialUpdate where Cdt<dateadd(dd,-90,getdate()) and Customer='FJ'
    delete from Ivan_ShortageMaterialUpdate where convert(char(10),Cdt,111)=convert(char(10),getdate(),111) and Customer='FJ'
    insert Ivan_ShortageMaterialUpdate
         select distinct POVendor,Material,[Shortage Qty],[Material ETA],Remark,getdate(),'FJ' from FJ_Material$
end

----TINY
delete from TINY_Material$ where Material is null
if exists(select * from TINY_Material$)
begin
    update TINY_Material$ set [Material ETA]='' where not [Material ETA] like '%*%'
    update TINY_Material$ set [Material ETA]=isnull([Material ETA],''),Remark=isnull(Remark,'')
    update TINY_Material$ set [Material ETA]=replace([Material ETA],';',',')
    update TINY_Material$ set [Material ETA]=replace([Material ETA],' ','')
    update TINY_Material$ set [Material ETA]=case when right([Material ETA],1)=',' then left([Material ETA],len([Material ETA])-1) else [Material ETA] end where not [Material ETA]=''
    
    delete from Ivan_ShortageMaterialUpdate where Cdt<dateadd(dd,-90,getdate()) and Customer='TINY'
    delete from Ivan_ShortageMaterialUpdate where convert(char(10),Cdt,111)=convert(char(10),getdate(),111) and Customer='TINY'
    insert Ivan_ShortageMaterialUpdate
         select distinct POVendor,Material,[Shortage Qty],[Material ETA],Remark,getdate(),'TINY' from TINY_Material$
end
/*
----MEDION
delete from MED_Material$ where Material is null
if exists(select * from MED_Material$)
begin
    update MED_Material$ set [Material ETA]='' where not [Material ETA] like '%*%'
    update MED_Material$ set [Material ETA]=isnull([Material ETA],''),Remark=isnull(Remark,'')
    update MED_Material$ set [Material ETA]=replace([Material ETA],';',',')
    update MED_Material$ set [Material ETA]=replace([Material ETA],' ','')
    update MED_Material$ set [Material ETA]=case when right([Material ETA],1)=',' then left([Material ETA],len([Material ETA])-1) else [Material ETA] end where not [Material ETA]=''
    
    delete from Ivan_ShortageMaterialUpdate where Cdt<dateadd(dd,-90,getdate()) and Customer='MEDION'
    delete from Ivan_ShortageMaterialUpdate where convert(char(10),Cdt,111)=convert(char(10),getdate(),111) and Customer='MEDION'
    insert Ivan_ShortageMaterialUpdate
         select distinct POVendor,Material,[Shortage Qty],[Material ETA],Remark,getdate(),'MEDION' from MED_Material$
end
*/
---------(2021/06/23) Add CPU/GPU PN into Remark
------------------------------------------------------------------------
--drop table #AP
update ASUS_Material$ set Remark=left(Remark,charindex('@',Remark)-1)  where Remark like '%@%'

select SA,CPU,ACPU=CPU,GPU,AGPU=GPU into #AP from ASUSSABS$

update #AP set ACPU=replace(ACPU,left(ACPU,12),b.Old_Material) from #AP a, t_download_matmas_CP69DW b where 
left(ACPU,12)=b.Material

update #AP set ACPU=replace(ACPU,substring(ACPU,charindex('/',ACPU)+1,12),b.Old_Material) from #AP a, t_download_matmas_CP69DW b where 
substring(ACPU,charindex('/',ACPU)+1,12)=b.Material

update #AP set AGPU=replace(AGPU,left(AGPU,12),b.Old_Material) from #AP a, t_download_matmas_CP69DW b where 
left(AGPU,12)=b.Material

update #AP set AGPU=replace(AGPU,substring(AGPU,charindex('/',AGPU)+1,12),b.Old_Material) from #AP a, t_download_matmas_CP69DW b where 
substring(AGPU,charindex('/',AGPU)+1,12)=b.Material

update ASUS_Material$ set Remark=rtrim(Remark)+'@'+ACPU+','+AGPU from ASUS_Material$ a,#AP b
where a.Material=b.SA and  (Remark like '%CPU%' or Remark like '%GPU%')

-----------------------------------------------------------------
-----------------------------------------------------------
----ASUS
delete from ASUS_Material$ where Material is null
if exists(select * from ASUS_Material$)
begin
    update ASUS_Material$ set [Material ETA]='' where not [Material ETA] like '%*%'
    update ASUS_Material$ set [Material ETA]=isnull([Material ETA],''),Remark=isnull(Remark,'')
    update ASUS_Material$ set [Material ETA]=replace([Material ETA],';',',')
    update ASUS_Material$ set [Material ETA]=replace([Material ETA],' ','')
    update ASUS_Material$ set [Material ETA]=case when right([Material ETA],1)=',' then left([Material ETA],len([Material ETA])-1) else [Material ETA] end where not [Material ETA]=''
    
    delete from Ivan_ShortageMaterialUpdate where Cdt<dateadd(dd,-90,getdate()) and Customer='ASUS'
    delete from Ivan_ShortageMaterialUpdate where convert(char(10),Cdt,111)=convert(char(10),getdate(),111) and Customer='ASUS'
    insert Ivan_ShortageMaterialUpdate
         select distinct POVendor,Material,[Shortage Qty],[Material ETA],Remark,getdate(),'ASUS' from ASUS_Material$
end

-----------------------------------------------------------
----ASUSITH
delete from ASUSITH_Material$ where Material is null
if exists(select * from ASUSITH_Material$)
begin
    update ASUSITH_Material$ set [Material ETA]='' where not [Material ETA] like '%*%'
    update ASUSITH_Material$ set [Material ETA]=isnull([Material ETA],''),Remark=isnull(Remark,'')
    update ASUSITH_Material$ set [Material ETA]=replace([Material ETA],';',',')
    update ASUSITH_Material$ set [Material ETA]=replace([Material ETA],' ','')
    update ASUSITH_Material$ set [Material ETA]=case when right([Material ETA],1)=',' then left([Material ETA],len([Material ETA])-1) else [Material ETA] end where not [Material ETA]=''
    
    delete from Ivan_ShortageMaterialUpdate where Cdt<dateadd(dd,-90,getdate()) and Customer='ASUS_ITH'
    delete from Ivan_ShortageMaterialUpdate where convert(char(10),Cdt,111)=convert(char(10),getdate(),111) and Customer='ASUS_ITH'
    insert Ivan_ShortageMaterialUpdate
         select distinct POVendor,Material,[Shortage Qty],[Material ETA],Remark,getdate(),'ASUS_ITH' from ASUSITH_Material$
end

/*
----ACER
delete from ACER_Material$ where Material is null
if exists(select * from ACER_Material$)
begin
    update ACER_Material$ set [Material ETA]='' where not [Material ETA] like '%*%'
    update ACER_Material$ set [Material ETA]=isnull([Material ETA],''),Remark=isnull(Remark,'')
    update ACER_Material$ set [Material ETA]=replace([Material ETA],';',',')
    update ACER_Material$ set [Material ETA]=replace([Material ETA],' ','')
    update ACER_Material$ set [Material ETA]=case when right([Material ETA],1)=',' then left([Material ETA],len([Material ETA])-1) else [Material ETA] end where not [Material ETA]=''
    
    delete from Ivan_ShortageMaterialUpdate where Cdt<dateadd(dd,-90,getdate()) and Customer='ACER'
    delete from Ivan_ShortageMaterialUpdate where convert(char(10),Cdt,111)=convert(char(10),getdate(),111) and Customer='ACER'
    insert Ivan_ShortageMaterialUpdate
         select distinct POVendor,Material,[Shortage Qty],[Material ETA],Remark,getdate(),'ACER' from ACER_Material$
end

----LENOVO
delete from LEN_Material$ where Material is null
if exists(select * from LEN_Material$)
begin
    update LEN_Material$ set [Material ETA]='' where not [Material ETA] like '%*%'
    update LEN_Material$ set [Material ETA]=isnull([Material ETA],''),Remark=isnull(Remark,'')
    update LEN_Material$ set [Material ETA]=replace([Material ETA],';',',')
    update LEN_Material$ set [Material ETA]=replace([Material ETA],' ','')
    update LEN_Material$ set [Material ETA]=case when right([Material ETA],1)=',' then left([Material ETA],len([Material ETA])-1) else [Material ETA] end where not [Material ETA]=''
    
    delete from Ivan_ShortageMaterialUpdate where Cdt<dateadd(dd,-90,getdate()) and Customer='LENOVO'
    delete from Ivan_ShortageMaterialUpdate where convert(char(10),Cdt,111)=convert(char(10),getdate(),111) and Customer='ACER'
    insert Ivan_ShortageMaterialUpdate
         select distinct POVendor,Material,[Shortage Qty],[Material ETA],Remark,getdate(),'LENOVO' from LEN_Material$
end
*/
--select Cdt,count(*) from Ivan_ShortageMaterialUpdate group by Cdt order by Cdt
--select * from Ivan_ShortageMaterialUpdate where convert(char(10),Cdt,111)=convert(char(10),getdate(),111)
create table #tmp_SH_ETA(iid int identity(1,1),POVendor varchar(20),Material varchar(20),MaterialETA varchar(200),Customer varchar(10))
create table #SH_ETA(POVendor varchar(20),Material varchar(20),MaterialETA varchar(100),ETA char(10),Qty int,Customer varchar(10))

--Only Multiple ETA need to count
insert #tmp_SH_ETA
    select distinct POVendor,Material,MaterialETA,Customer from Ivan_ShortageMaterialUpdate where MaterialETA like '%,%' and convert(char(10),Cdt,111)=convert(char(10),getdate(),111)

---Add single item to #SH_ETA
insert #SH_ETA
    select distinct POVendor,Material,MaterialETA,'','',Customer from Ivan_ShortageMaterialUpdate where not MaterialETA like '%,%' and not MaterialETA='' and convert(char(10),Cdt,111)=convert(char(10),getdate(),111)
    
declare @a int
declare @b int
select @a=min(iid) from #tmp_SH_ETA
select @b=max(iid) from #tmp_SH_ETA

while @a<=@b
begin
     while (select charindex(',',MaterialETA) from #tmp_SH_ETA where iid=@a)>0
     begin
         insert #SH_ETA
            select distinct POVendor,Material,left(MaterialETA,charindex(',',MaterialETA)-1),'','',Customer from #tmp_SH_ETA where iid=@a
         update #tmp_SH_ETA set MaterialETA=substring(MaterialETA,charindex(',',MaterialETA)+1,1000) where iid=@a
         
         insert #SH_ETA
            select distinct POVendor,Material,rtrim(MaterialETA),'','',Customer from #tmp_SH_ETA where iid=@a and not MaterialETA like '%,%'
     end
     select @a=@a+1 
end

--update ETA /Qty
update #SH_ETA set ETA=left(MaterialETA,charindex('*',MaterialETA)-1),Qty=convert(int,rtrim(substring(MaterialETA,charindex('*',MaterialETA)+1,1000)))

---Format Date
update #SH_ETA set ETA=convert(char(10),convert(datetime,rtrim(ETA)+' 00:00'),111)

--select * from ASUS_Material$ where Remark like '%CPU%'
/*
6070B0232801	11/7*914
6070B0232801	11/25*2500
6034B0008102	201211/17*300
6054B1207401	2012/11/09/920
6037B0053801	2012/11/30*200 
select * from #SH_ETA order by ETA where MaterialETA like '%2000%'
select  ETA=left(MaterialETA,charindex('*',MaterialETA)-1),
Qty=convert(int,rtrim(substring(MaterialETA,charindex('*',MaterialETA)+1,1000))) 
from #SH_ETA

select * from Ivan_ShortageMaterialUpdate where Material='6037B0053801' and MaterialETA like '2012/11/30*200%'

update Material$ set [Material ETA]='2012/11/30*200' where Material='6037B0053801' and [Material ETA]='2012/11/30*200 '
update Material$ set [Material ETA]='2012/11/30*200' where Material='6037B0053801' and [Material ETA] like '2012/11/30*200%'
update Material$ set [Material ETA]='2012/11/7*914' where Material='6070B0232801' and [Material ETA]='11/7*914'
update Material$ set [Material ETA]='2012/11/25*2500' where Material='6070B0232801' and [Material ETA]='11/25*2500'


select * from Material$

update #SH_ETA set MaterialETA='2012/11/7*914' where Material='6070B0232801' and MaterialETA='11/7*914'
update #SH_ETA set MaterialETA='2012/11/25*2500' where Material='6070B0232801' and MaterialETA='11/25*2500'

update #SH_ETA set MaterialETA=rtrim(MaterialETA)
select * from #SH_ETA where MaterialETA like '%200%'
*/
--select count(*) from Ivan_OPOMaterialETD

---Get Old ETA
select * into #ICI from Ivan_OPOMaterialETD 

if (select count(*) from #SH_ETA)>0
begin 
     truncate table Ivan_OPOMaterialETD 
     insert dbo.Ivan_OPOMaterialETD
          select distinct POVendor,Material,Qty,ETA,Customer from #SH_ETA
end

---Compare with old ETA
--drop table #ETADiff
select *,ETADiff=0 into #ETADiff  from (
select POVendor,Material,NETA=min(ETD),Customer from Ivan_OPOMaterialETD group by POVendor,Material,Customer) as a left join
(select oPOVendor=POVendor,oMaterial=Material,OETA=min(ETD),oCustomer=Customer from #ICI group by POVendor,Material,Customer) as b 
on a.POVendor=b.oPOVendor and a.Material=b.oMaterial and a.Customer=b.oCustomer

update #ETADiff set ETADiff=9999 where OETA is null
update #ETADiff set ETADiff=datediff(dd,convert(datetime,NETA),convert(datetime,OETA)) where not OETA is null

--select * from #ETADiff



/*
---Get Old ETA
select * into #ICI from Ivan_CurrentINV where INVDate>
case 
when datepart(weekday,getdate())=2 then dateadd(dd,-3,getdate())
when datepart(weekday,getdate())=1 then dateadd(dd,-2,getdate())
else dateadd(dd,-1,getdate()) end 
*/
-----(2020/06/13) Keep old INV to compare Inventory
-----If same date  ,reflash.
if (select min(INVDate) from Ivan_CurrentINV)<>convert(char(10),getdate(),111)
begin
	delete from Ivan_CurrentINV_old
	insert Ivan_CurrentINV_old
		select * from Ivan_CurrentINV
end

---Remove old data 
truncate table Ivan_CurrentINV
truncate table Ivan_Current_M_INV
--drop table #tmpHPINV
---Insert latest Inventory to Ivan_CurrentINV
--select max(Cdt) from WIP_WHALL_TD
--(2016/01/18) Manual convert 146 to 131
--insert Ivan_CurrentINV
--drop table #tmpHPINV
create table #tmpHPINV(POVendor varchar(20),MaterialProperty varchar(20),MatNo varchar(20),Qty int,DD varchar(20),ETADiff varchar(20),Customer varchar(20))

insert #tmpHPINV
select distinct POVendor=case 
when a.Plant='CP81' then 'IES'
when a.Plant='CP60' then 'ICC' 
end
,a.MaterialProperty,a.MatNo,Qty=sum(a.Qty),DD,ETADiff=0,Customer='HP' from (
------(Add c.Customer to prevent distinct)
------(2018/08/07) Add SLoc to prevent overlooking the duplicate Qty but in different SLoc..
select distinct a.SLoc,c.Customer,a.Plant,b.MaterialProperty,a.MatNo,a.Qty,DD=convert(char(10),a.Cdt,111)from WIP_WHALL_TD a,PType b,WType c where a.Plant=b.Plant and a.SLoc=b.SLoc and a.SType=b.SType and 
b.Plant=c.Plant and b.SLoc=c.SLoc and 
b.MaterialProperty in ('SHIP','FG') and b.Plant in ('CP81','CP60') and c.Customer in ('HP','HPTC')  and Cdt=(select max(Cdt) from WIP_WHALL_TD) and not (a.SLoc in ('PXRG','PXRI') and left(a.MatNo,2) in ('PF','SF','JF'))
-----(2018/01/23) add PXRG U stock 60% material into Invan_CurrentINV
union
select distinct a.SLoc,c.Customer,a.Plant,MaterialProperty='FG',a.MatNo,a.Qty,DD=convert(char(10),a.Cdt,111)from WIP_WHALL_TD a,PType b,WType c 
where a.Plant=b.Plant and a.SLoc=b.SLoc and a.SType=b.SType and 
b.Plant=c.Plant and b.SLoc=c.SLoc and a.SLoc='PXRG' and
b.Plant='CP60' and a.SType='Unstricted' and Cdt=(select max(Cdt) from WIP_WHALL_TD) and MatNo like '60%' 

) as a 

group by a.Plant,a.MaterialProperty,a.MatNo,DD order by DD


update #tmpHPINV set MatNo=b.OldMaterial from #tmpHPINV a,AssyMapping b where a.MatNo=b.NewMaterial

insert Ivan_CurrentINV
select POVendor,MaterialProperty,MatNo,Qty=sum(Qty),DD,ETADiff,Customer from #tmpHPINV group by POVendor,MaterialProperty,MatNo,DD,ETADiff,Customer order by DD

delete from #tmpHPINV

-------(2014/04/28) Non-HP Inventory .
--insert Ivan_CurrentINV
insert #tmpHPINV
select distinct POVendor=case 
when a.Plant='CP60' then 'ICC'
when a.Plant='CP62' then 'ICC'
when a.Plant='CP63' then 'ICC'
when a.Plant='CP65' then 'ICC'
when a.Plant='CP69' then 'ICC'
when a.Plant='TH03' then 'ITH'
when a.Plant='TH02' then 'ITH'
when a.Plant='TH05' then 'ITH'
when a.Plant='TH60' then 'ITH'
when a.Plant='UM60' then 'IMP'
else 'IES' end
,a.MaterialProperty,a.MatNo,Qty=sum(Qty),DD,0,a.Customer from (

select distinct a.Plant,a.SLoc,b.MaterialProperty,a.MatNo,a.Qty,DD=convert(char(10),a.Cdt,111),c.Customer from WIP_WHALL_TD a,PType b,WType c where a.Plant=b.Plant and a.SLoc=b.SLoc and a.SType=b.SType and 
b.Plant=c.Plant and b.SLoc=c.SLoc and
b.MaterialProperty in ('SHIP','FG') and not c.Customer in ('HP','HPTC') and Cdt=(select max(Cdt) from WIP_WHALL_TD) 
)as a 

group by a.Plant,a.MaterialProperty,a.MatNo,DD,a.Customer order by DD

update #tmpHPINV set MatNo=b.OldMaterial from #tmpHPINV a,AssyMapping b where a.MatNo=b.NewMaterial

insert Ivan_CurrentINV
select POVendor,MaterialProperty,MatNo,Qty=sum(Qty),DD,ETADiff,Customer from #tmpHPINV group by POVendor,MaterialProperty,MatNo,DD,ETADiff,Customer order by DD


-----(2013/10/07) add Manufacture INV
insert Ivan_Current_M_INV
select distinct POVendor=case 
when a.Plant in ('CP81','CP07') then 'IES'
when a.Plant in ('CP60','CP62','CP65','CP69') then 'ICC' 
when a.Plant in ('TH02','TH03','TH05','TH60') then 'ITH' 
when a.Plant in ('UM60') then 'IMP' 
else 'IES' end
,a.InventoryType
,a.MaterialProperty,a.MatNo,Qty=sum(a.Qty),DD,Customer from (

---(2018/03/20) add M_OBS in MP Inventory 
select distinct a.Plant,a.InventoryType,b.MaterialProperty,a.MatNo,a.Qty,DD=convert(char(10),a.Cdt,111),
Customer=case when c.Customer='HPTC' then 'HP' else Customer end from WIP_WHALL_TD a,PType b,WType c where a.Plant=b.Plant and a.SLoc=b.SLoc and a.SType=b.SType and 
b.Plant=c.Plant and b.SLoc=c.SLoc and
b.MaterialProperty in ('M_FG','M_OBS') and Cdt=(select max(Cdt) from WIP_WHALL_TD) 
)as a 

group by a.Plant,a.InventoryType,a.MaterialProperty,a.MatNo,DD,a.Customer order by DD 
----HP RMA FG (xF)
insert Ivan_CurrentINV
select distinct POVendor=case 
when Plant='CP81' then 'IES'
when Plant='CP60' then 'ICC'
when Plant='TH03' then 'ITH'
when Plant='TH02' then 'ITH'
when Plant='TH05' then 'ITH'
when Plant='TH60' then 'ITH'
when Plant='UM60' then 'IMP'
end
,'RMA_FG',MatNo,Qty=sum(Qty),DD,0,'HP' from (

select distinct Plant,MatNo,Qty,DD=convert(char(10),Cdt,111) from WIP_WHALL_TD where Plant in ('CP81','CP60') and SLoc='PXRC' and substring(MatNo,2,1)='F'
and SType='Unstricted' and Cdt=(select max(Cdt) from WIP_WHALL_TD) 
) as a

group by Plant,MatNo,DD order by DD


----TSB RMA FG (xF)
insert Ivan_CurrentINV
select distinct POVendor=case 
when Plant='CP07' then 'IES'
when Plant='CP62' then 'ICC' end
,'RMA_FG',MatNo,Qty=sum(Qty),DD,0,'TSB' from (

select distinct Plant,MatNo,Qty,DD=convert(char(10),Cdt,111) from WIP_WHALL_TD where Plant in ('CP07','CP62') and SLoc='PXRB' and substring(MatNo,2,1)='F'
and SType='Unstricted' and Cdt=(select max(Cdt) from WIP_WHALL_TD) 
) as a

group by Plant,MatNo,DD order by DD



--Insert future ETA 
insert Ivan_CurrentINV
select distinct POVendor,'FG',Material,Qty,ETD,0,Customer from Ivan_OPOMaterialETD where ETD>convert(char(10),getdate(),111) order by ETD


update Ivan_CurrentINV set ETADiff=b.ETADiff from Ivan_CurrentINV a,#ETADiff b where a.POVendor=b.POVendor and a.MatNo=b.Material and a.INVDate=b.NETA and a.Customer=b.Customer

select min(INVDate) from Ivan_CurrentINV_old
select count(*) from Ivan_CurrentINV
select top 10 * from Ivan_CurrentINV
select top 10 * from Ivan_Current_M_INV


select * from Ivan_CurrentINV where ETADiff<>0
select * from #ETADiff where Material='6043B0116803'
/*
select * from
(select *  from Es$) aS a LEFT join
(select * from OPO$) as b on a.PO=b.SO and a.[IEC PN]=b.IECPN
*/

----HP
--delete from OPO$ where Escalation is null
update OPO$ set Remark='' where Remark is null
update OPO$ set Escalation='' where Escalation is null
update OPO$ set [Escalation Date]='' where [Escalation Date] is null
update OPO$ set Remark='' where Remark='NULL'
update OPO$ set Escalation='' where Escalation='NULL'
update OPO$ set [Escalation Date]='' where [Escalation Date]='NULL'
delete from OPO$ where Escalation='' and [Escalation Date]='' and Remark=''


select * from OPO$

----H2
--delete from H2_OPO$ where Escalation is null
update H2_OPO$ set Remark='' where Remark is null
update H2_OPO$ set Escalation='' where Escalation is null
update H2_OPO$ set [Escalation Date]='' where [Escalation Date] is null
update H2_OPO$ set Remark='' where Remark='NULL'
update H2_OPO$ set Escalation='' where Escalation='NULL'
update H2_OPO$ set [Escalation Date]='' where [Escalation Date]='NULL'
delete from H2_OPO$ where Escalation='' and [Escalation Date]='' and Remark=''


select * from H2_OPO$

----HP_ITH
--delete from HPITH_OPO$ where Escalation is null
update HPITH_OPO$ set Remark='' where Remark is null
update HPITH_OPO$ set Escalation='' where Escalation is null
update HPITH_OPO$ set [Escalation Date]='' where [Escalation Date] is null
update HPITH_OPO$ set Remark='' where Remark='NULL'
update HPITH_OPO$ set Escalation='' where Escalation='NULL'
update HPITH_OPO$ set [Escalation Date]='' where [Escalation Date]='NULL'
delete from HPITH_OPO$ where Escalation='' and [Escalation Date]='' and Remark=''


----HP_IMP
--delete from HPITH_OPO$ where Escalation is null
update HPIMP_OPO$ set Remark='' where Remark is null
update HPIMP_OPO$ set Escalation='' where Escalation is null
update HPIMP_OPO$ set [Escalation Date]='' where [Escalation Date] is null
update HPIMP_OPO$ set Remark='' where Remark='NULL'
update HPIMP_OPO$ set Escalation='' where Escalation='NULL'
update HPIMP_OPO$ set [Escalation Date]='' where [Escalation Date]='NULL'
delete from HPIMP_OPO$ where Escalation='' and [Escalation Date]='' and Remark=''


select * from HPIMP_OPO$
/*
----TSB
--delete from TSB_OPO$ where Escalation is null
update TSB_OPO$ set Remark='' where Remark is null
update TSB_OPO$ set SO='' where SO is null
update TSB_OPO$ set Escalation='' where Escalation is null
update TSB_OPO$ set [Escalation Date]='' where [Escalation Date] is null
update TSB_OPO$ set Remark='' where Remark='NULL'
update TSB_OPO$ set Escalation='' where Escalation='NULL'
update TSB_OPO$ set [Escalation Date]='' where [Escalation Date]='NULL'
delete from TSB_OPO$ where Escalation='' and [Escalation Date]='' and Remark=''


select * from TSB_OPO$
*/

----FJ
--delete from FJ_OPO$ where Escalation is null
update FJ_OPO$ set Remark='' where Remark is null
update FJ_OPO$ set Escalation='' where Escalation is null
update FJ_OPO$ set [Escalation Date]='' where [Escalation Date] is null
update FJ_OPO$ set Remark='' where Remark='NULL'
update FJ_OPO$ set Escalation='' where Escalation='NULL'
update FJ_OPO$ set [Escalation Date]='' where [Escalation Date]='NULL'
delete from FJ_OPO$ where Escalation='' and [Escalation Date]='' and Remark=''


select * from FJ_OPO$

----TINY
--delete from TINY_OPO$ where Escalation is null
update TINY_OPO$ set Remark='' where Remark is null
update TINY_OPO$ set Escalation='' where Escalation is null
update TINY_OPO$ set [Escalation Date]='' where [Escalation Date] is null
update TINY_OPO$ set Remark='' where Remark='NULL'
update TINY_OPO$ set Escalation='' where Escalation='NULL'
update TINY_OPO$ set [Escalation Date]='' where [Escalation Date]='NULL'
delete from TINY_OPO$ where Escalation='' and [Escalation Date]='' and Remark=''


select * from TINY_OPO$
/*
----MEDION
--delete from MED_OPO$ where Escalation is null
update MED_OPO$ set Remark='' where Remark is null
update MED_OPO$ set Escalation='' where Escalation is null
update MED_OPO$ set [Escalation Date]='' where [Escalation Date] is null
update MED_OPO$ set Remark='' where Remark='NULL'
update MED_OPO$ set Escalation='' where Escalation='NULL'
update MED_OPO$ set [Escalation Date]='' where [Escalation Date]='NULL'
delete from MED_OPO$ where Escalation='' and [Escalation Date]='' and Remark=''

select * from MED_OPO$
*/
----ASUS
--select * from ASUS_OPO$ a,(select distinct a.SO from Service_APD a,ASUSDN$ b where a.PO=b.ASUSPO)as b where a.SO=b.SO
--update ASUS_OPO$ set Remark='DN0118,'+a.Remark from ASUS_OPO$ a,(select distinct a.SO from Service_APD a,ASUSDN$ b where a.PO=b.ASUSPO)as b where a.SO=b.SO

--delete from ASUS_OPO$ where Escalation is null
update ASUS_OPO$ set Remark='' where Remark is null
update ASUS_OPO$ set Escalation='' where Escalation is null
update ASUS_OPO$ set [Escalation Date]='' where [Escalation Date] is null
update ASUS_OPO$ set [Escalation Date]='' where [Escalation Date]='0000/00/00'
update ASUS_OPO$ set Remark='' where Remark='NULL'
update ASUS_OPO$ set Escalation='' where Escalation='NULL'
update ASUS_OPO$ set [Escalation Date]='' where [Escalation Date]='NULL'
delete from ASUS_OPO$ where Escalation='' and [Escalation Date]='' and Remark=''


select * from ASUS_OPO$

----ASUSITH

--delete from ASUSITH_OPO$ where Escalation is null
update ASUSITH_OPO$ set Remark='' where Remark is null
update ASUSITH_OPO$ set Escalation='' where Escalation is null
update ASUSITH_OPO$ set [Escalation Date]='' where [Escalation Date] is null
update ASUSITH_OPO$ set [Escalation Date]='' where [Escalation Date]='0000/00/00'
update ASUSITH_OPO$ set Remark='' where Remark='NULL'
update ASUSITH_OPO$ set Escalation='' where Escalation='NULL'
update ASUSITH_OPO$ set [Escalation Date]='' where [Escalation Date]='NULL'
delete from ASUSITH_OPO$ where Escalation='' and [Escalation Date]='' and Remark=''


select * from ASUSITH_OPO$


/*
----ACER
--delete from ACER_OPO$ where Escalation is null
update ACER_OPO$ set Remark='' where Remark is null
update ACER_OPO$ set Escalation='' where Escalation is null
update ACER_OPO$ set [Escalation Date]='' where [Escalation Date] is null
update ACER_OPO$ set [Escalation Date]='' where [Escalation Date]='0000/00/00'
update ACER_OPO$ set Remark='' where Remark='NULL'
update ACER_OPO$ set Escalation='' where Escalation='NULL'
update ACER_OPO$ set [Escalation Date]='' where [Escalation Date]='NULL'
delete from ACER_OPO$ where Escalation='' and [Escalation Date]='' and Remark=''


select * from ACER_OPO$


----LENOVO
--delete from LEN_OPO$ where Escalation is null
update LEN_OPO$ set Remark='' where Remark is null
update LEN_OPO$ set Escalation='' where Escalation is null
update LEN_OPO$ set [Escalation Date]='' where [Escalation Date] is null
update LEN_OPO$ set Remark='' where Remark='NULL'
update LEN_OPO$ set Escalation='' where Escalation='NULL'
update LEN_OPO$ set [Escalation Date]='' where [Escalation Date]='NULL'
delete from LEN_OPO$ where Escalation='' and [Escalation Date]='' and Remark=''

select * from LEN_OPO$
*/

------¬O§_¤J±b
select Customer,Material,INVCHG=isnull(NewQty,0)-isnull(OldQty,0),ETA,Qty from 
(
select * from(
select 
iCustomer=case when a.Customer is null then b.Customer else a.Customer end ,
MatNo=case when a.MatNo is null then b.MatNo else a.MatNo end ,
MP=case when a.MP is null then b.MP else a.MP end ,
NewQty=a.Qty,OldQty=b.Qty from 
(select * from Ivan_CurrentINV where INVDate=(select min(INVDate) from Ivan_CurrentINV)) as a full join
(select * from Ivan_CurrentINV_old where INVDate=(select min(INVDate) from Ivan_CurrentINV_old))  as b
on a.MatNo=b.MatNo and a.Customer=b.Customer and a.MP=b.MP
) as a  right join
(select Material,Customer,ETA,Qty from #SH_ETA where ETA=convert(char(10),getdate(),111)) as b on a.MatNo=Material and a.iCustomer=b.Customer
) as a
order by Customer,Qty desc


-----(2021/10/22) Get 93 Status
--select PoNo,Model,Qty,ShipDate,[Status],Editor from [10.96.183.62,998].IMES_Rep.dbo.[Delivery] where left(Model,2) in ('LF','RF') and ShipDate>convert(char(10),dateadd(dd,-5,getdate()),111)  and Status='93'


