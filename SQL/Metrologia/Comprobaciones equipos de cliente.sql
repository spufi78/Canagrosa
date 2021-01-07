-- Equipos de Cliente por cliente
select c.NOMBRE,count(*) from equipos a
left join clientes c on a.CLIENTE_ID = c.ID_CLIENTE 
 where a.NUMERO_EQUIPO_CLIENTE <> '' 
 group by 1;

select a.ID_EQUIPO,a.NOMBRE,a.ESTADO_ID,a.NUMERO_EQUIPO_CLIENTE 
from equipos a
 where a.NUMERO_EQUIPO_CLIENTE <> '' and a.CLIENTE_ID = 0
 group by 1;

-- Equipos Airbus con mantenimiento
select distinct a.ID_EQUIPO,a.NOMBRE,a.ESTADO_ID,a.NUMERO_EQUIPO_CLIENTE 
from equipos a
inner join clientes c on a.CLIENTE_ID = c.ID_CLIENTE 
inner join eq_mantenimiento_equipos b on a.ID_EQUIPO = b.EQUIPO_ID 
where a.NUMERO_EQUIPO_CLIENTE <> '' and c.AIRBUS = 1
and b.ESTADO = 0;

-- Equipos Airbus que nunca se han calibrado, verificado
select distinct a.ID_EQUIPO,a.NOMBRE,a.ALTA_BAJA,a.ESTADO_ID,a.NUMERO_EQUIPO_CLIENTE, MAX(b.FECHA_ACTUAL) as VERIFICACION,MAX(d.FECHA_ACTUAL) as CALIBRACION
from equipos a
inner join clientes c on a.CLIENTE_ID = c.ID_CLIENTE 
left join eq_verificacion_equipos b on a.ID_EQUIPO = b.EQUIPO_ID 
left join eq_calibracion_equipos d on a.ID_EQUIPO = d.EQUIPO_ID 
where a.NUMERO_EQUIPO_CLIENTE <> '' and c.AIRBUS = 1
and isnull(b.ID_VERIFICACION)
and isnull(d.ID_CALIBRACION)
group by 1;

select distinct a.ID_EQUIPO,a.NOMBRE,a.ALTA_BAJA,a.ESTADO_ID,a.NUMERO_EQUIPO_CLIENTE
 ,count(c1.ID_CALIBRACION) AS CAL_PENDIENTES,min(c1.FECHA_PROXIMA) AS CAL_FECHA,count(c2.ID_CALIBRACION) AS CAL_REALIZADAS,max(c2.FECHA_PROXIMA) AS CAL_FECHA
 ,count(v1.ID_VERIFICACION) AS VER_PENDIENTES,min(v1.FECHA_PROXIMA) AS CAL_FECHA,count(v2.ID_VERIFICACION) AS VER_REALIZADAS,max(v2.FECHA_PROXIMA) AS CAL_FECHA
from equipos a
inner join clientes c on a.CLIENTE_ID = c.ID_CLIENTE 
left join eq_calibracion_equipos c1 on a.ID_EQUIPO = c1.EQUIPO_ID and c1.ESTADO = 0
left join eq_calibracion_equipos c2 on a.ID_EQUIPO = c2.EQUIPO_ID and c2.ESTADO <> 0
left join eq_verificacion_equipos v1 on a.ID_EQUIPO = v1.EQUIPO_ID and v1.ESTADO = 0
left join eq_verificacion_equipos v2 on a.ID_EQUIPO = v2.EQUIPO_ID and v2.ESTADO <> 0
where a.NUMERO_EQUIPO_CLIENTE <> '' and c.AIRBUS = 1
group by 1;


select distinct a.ID_EQUIPO,a.NOMBRE,a.ALTA_BAJA,a.ESTADO_ID,a.NUMERO_EQUIPO_CLIENTE
 ,count(c1.ID_CALIBRACION) AS CAL_PENDIENTES,min(c1.FECHA_ACTUAL) AS CAL_FECHA,count(c2.ID_CALIBRACION) AS CAL_REALIZADAS,max(c2.FECHA_ACTUAL) AS CAL_FECHA
 ,count(v1.ID_VERIFICACION) AS VER_PENDIENTES,min(v1.FECHA_ACTUAL) AS VER_FECHA,count(v2.ID_VERIFICACION) AS VER_REALIZADAS,max(v2.FECHA_ACTUAL) AS VER_FECHA
from equipos a
inner join clientes c on a.CLIENTE_ID = c.ID_CLIENTE 
left join eq_calibracion_equipos c1 on a.ID_EQUIPO = c1.EQUIPO_ID and c1.ESTADO = 0
left join eq_calibracion_equipos c2 on a.ID_EQUIPO = c2.EQUIPO_ID and c2.ESTADO <> 0
left join eq_verificacion_equipos v1 on a.ID_EQUIPO = v1.EQUIPO_ID and v1.ESTADO = 0
left join eq_verificacion_equipos v2 on a.ID_EQUIPO = v2.EQUIPO_ID and v2.ESTADO <> 0
where a.NUMERO_EQUIPO_CLIENTE <> '' and c.AIRBUS = 1
group by 1;


select * from eq_calibracion_equipos where equipo_id = 2533;
12117,
12378,
12379,
12830,
13022,
13611);


select * from equipos order by id_equipo desc limit 10; 