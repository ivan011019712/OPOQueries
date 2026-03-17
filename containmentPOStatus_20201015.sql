--drop table #tmp

select distinct Site,a.PO,SO,Item,CPQPN,IECPN,ProductFamily,PODate=Date850,NeedShipDate=SO_First_Date,
RSD=left(SO_ReqDate,4)+'/'+substring(SO_ReqDate,5,2)+'/'+substring(SO_ReqDate,7,2),POQty=Qty850,PO_Type,FCST_Status,ShipQty_old=0,
OpenQty_old=0,ShipQty_0304=0,OpenQty_new=0,PI_NS=0,PI_RSD=0,PT
into #tmp from Service_APD a,
SMSCM b
where a.PO=b.PO and a.CPQPN=b.OSSPPN


update #tmp set ShipQty_old=b.ShipQty from #tmp a,
(
select a.SO,a.Item,a.IECPN,ShipQty=sum(Qty856) from Service_APD a,#tmp b where a.SO=b.SO and a.Item=b.Item and a.IECPN=b.IECPN
and not a.PndGIDate='0000/00/00' and a.PndGIDate<'2019/03/04'
group by a.SO,a.Item,a.IECPN
) as b where a.SO=b.SO and a.Item=b.Item and a.IECPN=b.IECPN

update #tmp set ShipQty_0304=b.ShipQty from #tmp a,
(
select a.SO,a.Item,a.IECPN,ShipQty=sum(Qty856) from Service_APD a,#tmp b where a.SO=b.SO and a.Item=b.Item and a.IECPN=b.IECPN
and not a.PndGIDate='0000/00/00' and a.PndGIDate>='2019/03/04'
group by a.SO,a.Item,a.IECPN
) as b where a.SO=b.SO and a.Item=b.Item and a.IECPN=b.IECPN

-----(2019/10/15) Add Pullin NeedShip Date
update #tmp set PI_NS=b.ShipQty from #tmp a,
(
select a.SO,a.Item,a.IECPN,ShipQty=sum(Qty856) from Service_APD a,#tmp b where a.SO=b.SO and a.Item=b.Item and a.IECPN=b.IECPN
and not a.PndGIDate='0000/00/00' and a.PndGIDate>a.SO_First_Date
group by a.SO,a.Item,a.IECPN
) as b where a.SO=b.SO and a.Item=b.Item and a.IECPN=b.IECPN


-----(2019/10/15) Add Pullin RSD
update #tmp set PI_RSD=b.ShipQty from #tmp a,
(
select a.SO,a.Item,a.IECPN,ShipQty=sum(Qty856) from Service_APD a,#tmp b where a.SO=b.SO and a.Item=b.Item and a.IECPN=b.IECPN
and not a.PndGIDate='0000/00/00' and a.PndGIDate>left(SO_ReqDate,4)+'/'+substring(SO_ReqDate,5,2)+'/'+substring(SO_ReqDate,7,2)
group by a.SO,a.Item,a.IECPN
) as b where a.SO=b.SO and a.Item=b.Item and a.IECPN=b.IECPN






update #tmp set OpenQty_old=POQty-ShipQty_old

update #tmp set OpenQty_new=OpenQty_old-ShipQty_0304

--(2020/10/15) Add "*" as long as the PO ever set to "P1"
update #tmp set PO=a.PO+'*' from #tmp a,SMSCM b where a.PO=b.PO and a.CPQPN=b.OSSPPN and b.PT='P1'
update #tmp set PO=a.PO+'*' from #tmp a,SMSCM_History b where a.PO=b.PO and a.CPQPN=b.OSSPPN and b.PT='P1' and not a.PO like '%*'

select * from #tmp --where CPQPN='831837-001'          


select CPQPN,PT,POQty=sum(POQty),ShipQty_old=sum(ShipQty_old),
OpenQty_old=sum(OpenQty_old),ShipQty_0304=sum(ShipQty_0304),OpenQty_new=sum(OpenQty_new) 
from #tmp group by CPQPN,PT order by sum(OpenQty_old) desc

--select CPQPN,sum(OpenQty_old) from #tmp group by CPQPN order by sum(OpenQty_old) desc

--select * from Service_APD where PO='14144079' and CPQPN='736463-001'