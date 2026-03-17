/*
Follow up		#fwu	　		依ETA追
Every Day	#evd	　		每天追
By Week		#www		ww=01~07	每星期幾追, e.g. : #w05 --> 每星期五追
By Month		#mdd		dd=0~31	每月幾號追, e.g. : #m20 --> 每月20號追
By Date		#mdd		m=1~c, dd=0~31	特定日期追, e.g. : #c13 --> 12月13號追
#neg
#quo
*/

/*
drop table #tmp
delete from Ivan_qi where ReportDate=convert(char(10),getdate(),111) 
*/


declare @dt char(10)
declare @Predt char(10)
select @dt=convert(char(10),dateadd(dd,0,getdate()),111)
select @Predt=convert(char(10),dateadd(dd,-1,getdate()),111)


--select * from OPS_Material where ReportDate='2022/09/13'  and RealOpenQty>0 and Customer='HP'

select ReportDate,Customer,ShortagePN,PIC,RealOpenQty,SODate,LatestSODate,NeedShipDate,MaterialETA,
qi=case when charindex('#',Remark)>0 then substring(Remark,charindex('#',Remark),4) else '' end,Priority='10',eType='-',eMsg='                                                                                                                             ',
Age=0,LD=0,ND=0,
Remark into #tmp from OPS_Material where ReportDate=@dt  and RealOpenQty>0 and Customer in ('HP','ASUS','H2','ASUS_ITH','HP_ITH')
----Remove LatestSODate is 2099/01/01
delete from #tmp where LatestSODate='2099/01/01'

insert #tmp
select ReportDate,'BL',Material,PIC,ShortageQty,MinPO,MaxPO,MinNeed,Material=case when [Material ETA] in ('-','_')  then '' else [Material ETA] end,
qi=case when charindex('#',Remark)>0 then substring(Remark,charindex('#',Remark),4) else '' end,Priority='10',eType='-',eMsg='                                                                                                                             ',
Age=0,LD=0,ND=0,Remark from OPS_BL  where ReportDate=@dt and Customer='HP'

insert #tmp
select ReportDate,'ITH_BL',Material,PIC,ShortageQty,MinPO,MaxPO,MinNeed,Material=case when [Material ETA] in ('-','_')  then '' else [Material ETA] end,
qi=case when charindex('#',Remark)>0 then substring(Remark,charindex('#',Remark),4) else '' end,Priority='10',eType='-',eMsg='                                                                                                                             ',
Age=0,LD=0,ND=0,Remark from OPS_BL  where ReportDate=@dt and Customer='HP_ITH'




---select * from #tmp where SODate='0000/00/00'
---select * from #tmp where MaterialETA='' and SODate='0000/00/00'
---select * from #tmp where SODate='' and LatestSODate=''
update #tmp set SODate=LatestSODate where SODate='0000/00/00'
--delete from #tmp where SODate='' and LatestSODate=''
--delete from #tmp where SODate='0000/00/00' and LatestSODate='0000/00/00'

update #tmp set Age=datediff(dd,SODate,@dt) --where MaterialETA=''
update #tmp set LD=datediff(dd,SODate,left(MaterialETA,charindex('*',MaterialETA)-1)) where not MaterialETA=''
update #tmp set ND=datediff(dd,NeedShipDate,left(MaterialETA,charindex('*',MaterialETA)-1)) where not MaterialETA=''



/*
declare @dt char(10)
declare @Predt char(10)
select @dt=convert(char(10),dateadd(dd,0,getdate()),111)
select @Predt=convert(char(10),dateadd(dd,-1,getdate()),111)
*/

---Find New (前一天的 SO)
update #tmp set eType='N',eMsg='New' where rtrim(LatestSODate)=@Predt and SODate=LatestSODate  and eType='-'


--Find Error ---Incorrect flag
update #tmp set eType='E',eMsg=rtrim(eMsg)+'Error_incorrect_flag' where (not substring(qi,2,1) in ('f','e','m','w','1','2','3','4','5','6','7','8','9','a','b','c','n','q')  
or not  substring(qi,3,1) in ('v','w','s','1','2','3','4','5','6','7','8','9','0') 
or not  substring(qi,4,1) in ('u','d','w','1','2','3','4','5','6','7','8','9') 
)
and not qi=''  
/*
select * from #tmp where (not substring(qi,2,1) in ('f','e','m','w','1','2','3','4','5','6','7','8','9','a','b','c','n','q')  
or not  substring(qi,3,1) in ('v','w','1','2','3','4','5','6','7','8','9','0') 
or not  substring(qi,4,1) in ('u','d','w','1','2','3','4','5','6','7','8','9') 
)
and not qi=''  
*/

--Find Error ---Pre Day but qi is not blank
update #tmp set eType='E',eMsg=rtrim(eMsg)+'Error_qi_not_blank;' where rtrim(LatestSODate)=@Predt and SODate=LatestSODate and qi<>'' and eType='-'
--select * from #tmp where rtrim(LatestSODate)=@Predt and SODate=LatestSODate and qi<>''

--Find Error --Follow but no ETA
update #tmp set eType='E',eMsg=rtrim(eMsg)+'Error_#fwu_but_no_ETA;' where qi='#fwu' and MaterialETA='' and eType='-'
--select * from #tmp where qi='#fwu' and MaterialETA=''

--Find Error --Not fill action
update #tmp set eType='E',eMsg=rtrim(eMsg)+'Error_not_fill_qi;' where  qi='' and MaterialETA='' and not (rtrim(LatestSODate)=@Predt and SODate=LatestSODate)  and eType='-'
--select * from #tmp where  qi='' and MaterialETA='' and not (rtrim(LatestSODate)=@Predt and SODate=LatestSODate) 


-----waring LD-Age > 30 days but still follow up
update #tmp set eType='W',eMsg=rtrim(eMsg)+'Warning_#fwu_but_>30days;' where LD-Age>30 and qi='#fwu' and eType='-' and ND>=-7
--select * from #tmp where LD-Age>30 and qi='#fwu' and ND>=-7


-----waring SO Age > 30 days no ETA but still check every day
update #tmp set eType='W',eMsg=rtrim(eMsg)+'Warning_Age>30days_No_ETA_Still_evd;' where Age>30 and qi='#evd' and LD=0 and eType='-'
--select * from #tmp where Age>30 and qi='#evd' and LD=0

-----waring  Age > 30 days w/ETA but still check every day
update #tmp set eType='W',eMsg=rtrim(eMsg)+'Warning_with_ETA_Still_evd;' where qi='#evd' and LD>0 and eType='-' 
--select * from #tmp where qi='#evd' and LD>0 and ND>=0



-----waring Age > 14 days but no ETA should be escalate
update #tmp set eType='S',eMsg=rtrim(eMsg)+'Escalate_Age>14days_No_ETA;' where LD=0 and Age>14 and eType='-'
--select * from #tmp where qi='#evd' and LD>0


------Set Priority-------------------------------
update #tmp set Priority='01' where eMsg='New'
--select * from #tmp where eMsg='New'

update #tmp set Priority='02' where qi='#evd' 
--select * from #tmp where qi='#evd' order by eMsg

update #tmp set Priority='03' where qi like '#w%' and substring(qi,4,1)=datepart(weekday,getdate())-1
--select * from #tmp where qi like '#w%' and substring(qi,4,1)=datepart(weekday,getdate())-1 order by eMsg

update #tmp set Priority='04'  where qi like '#m%' and substring(qi,3,2)=substring(convert(char(10),getdate(),111),9,2) 
--select  * from #tmp where qi like '#m%' and substring(qi,3,2)=substring(convert(char(10),getdate(),111),9,2) order by eMsg

--update #tmp set Priority='05'  where substring(qi,2,1) in ('1','2','3','4','5','6','7','8','9','a','b','c') and substring(qi,3,2)=substring(convert(char(10),getdate(),111),9,2) 
--select  * from #tmp where substring(qi,2,1) in ('1','2','3','4','5','6','7','8','9','A','B','C') and substring(qi,3,2)=substring(convert(char(10),getdate(),111),9,2) order by eMsg

update #tmp set Priority='05' where substring(qi,2,1) in ('1','2','3','4','5','6','7','8','9','a','b','c') 
and case when substring(qi,2,1)='a' then '10' when substring(qi,2,1)='b' then '11' when substring(qi,2,1)='c' then '12' else  substring(qi,2,1) end =convert(varchar(2),datepart(mm,getdate()))
and substring(qi,3,2)=substring(convert(char(10),getdate(),111),9,2) 

update #tmp set Priority='06' where qi like '#w%' and substring(qi,4,1)<>datepart(weekday,getdate())-1
--select * from #tmp where qi like '#w%' and substring(qi,4,1)<>datepart(weekday,getdate())-1 order by eMsg


update #tmp set Priority='07'  where qi='#fwu' 
--select  * from #tmp where qi='#fwd' order by eMsg

update #tmp set Priority='08'  where qi like '#es%' 
--select  * from #tmp where qi like '#es%' order by eMsg



/*
select * from #qq where substring(qi,2,1) in ('1','2','3','4','5','6','7','8','9','A','B','C') 
and case when substring(qi,2,1)='A' then '10' when substring(qi,2,1)='B' then '11' when substring(qi,2,1)='C' then '12' else  substring(qi,2,1) end =datepart(mm,getdate())
and substring(qi,3,2)=substring(convert(char(10),getdate(),111),9,2) 
*/

--update #tmp set PIC='Var' where Customer='ASUS'

-------------------------Analyze data
--select * from #tmp  where qi='#es1'
--select eType,count(*) from #tmp group by eType order by eType
--select Customer,eType,count(*) from #tmp group by Customer,eType order by Customer,eType
--select PIC,eType,count(*) from #tmp group by PIC,eType order by PIC,eType


--select qi,count(*) from #tmp where not eMsg='New' group by qi order by qi
--select Customer,qi,count(*) from #tmp group by Customer,qi order by Customer,qi
--select PIC,qi,count(*) from #tmp group by PIC,qi order by PIC,qi

---Detail
delete from Ivan_qi where ReportDate<=convert(char(10),dateadd(dd,-90,getdate()),111) 
delete from Ivan_qi where ReportDate=convert(char(10),getdate(),111) 

insert Ivan_qi
	select ReportDate,Customer,ShortagePN,PIC,RealOpenQty,MaterialETA,qi,Priority,eMsg from #tmp order by Customer,PIC,Priority,eMsg


select * from Ivan_qi where ReportDate=convert(char(10),getdate(),111) 
