select c2.campo2, cli.NOMBRE 
 from clientes cli, (
select t.ID_TARIFA tar,t.NOMBRE as campo2, count(*)
from clientes c, tarifas t
where c.TARIFA_ID = t.ID_TARIFA
and c.TARIFA_ID <> 0
group by t.ID_TARIFA,t.NOMBRE
having count(*) > 1) as c2
where cli.TARIFA_ID = c2.tar


     