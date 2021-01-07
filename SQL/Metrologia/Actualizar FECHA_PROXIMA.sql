update equipos e set e.FECHA_PROX_CALIBRACION = null, e.FECHA_PROX_VERIFICACION = null, e.FECHA_PROX_MANTENIMIENTO = null;

update equipos aa,(
select equipo_id as c1,max(fecha_actual) as c2
from eq_calibracion_equipos a
where a.estado <> 0
group by equipo_id) bb
set aa.FECHA_ULT_CALIBRACION = bb.c2
where aa.ID_EQUIPO = bb.c1;

update equipos aa,(
select equipo_id as c1,max(fecha_actual) as c2
from eq_verificacion_equipos a
where a.estado <> 0
group by equipo_id) bb
set aa.FECHA_ULT_VERIFICACION = bb.c2
where aa.ID_EQUIPO = bb.c1;


update equipos aa,(
select equipo_id as c1,max(fecha_actual) as c2, IF(b.MESES <> 0, date_add(max(fecha_actual), interval b.MESES MONTH), date_add(max(fecha_actual), interval b.DIAS DAY)) as c3
from eq_calibracion_equipos a, equipos e, eq_periodicidad b
where a.estado <> 0
and a.EQUIPO_ID = e.ID_EQUIPO 
and e.PERIODICIDAD_CALIBRACION_ID = b.ID_PERIODICIDAD
group by equipo_id) bb
set aa.FECHA_PROX_CALIBRACION = bb.c3
where aa.ID_EQUIPO = bb.c1;

update equipos aa,(
select equipo_id as c1,max(fecha_actual) as c2, IF(b.MESES <> 0, date_add(max(fecha_actual), interval b.MESES MONTH), date_add(max(fecha_actual), interval b.DIAS DAY)) as c3
from eq_verificacion_equipos a, eq_periodicidad b
where a.estado <> 0
and a.PERIODICIDAD_ID = b.ID_PERIODICIDAD
group by equipo_id) bb
set aa.FECHA_PROX_VERIFICACION = bb.c3
where aa.ID_EQUIPO = bb.c1;

update equipos aa,(
select equipo_id as c1,max(fecha_actual) as c2, IF(b.MESES <> 0, date_add(max(fecha_actual), interval b.MESES MONTH), date_add(max(fecha_actual), interval b.DIAS DAY)) as c3
from eq_mantenimiento_equipos a, equipos e, eq_periodicidad b
where a.estado <> 0
and e.ID_EQUIPO = a.equipo_id
and e.PERIODICIDAD_MANTENIMIENTO_ID = b.ID_PERIODICIDAD
group by equipo_id) bb
set aa.FECHA_PROX_MANTENIMIENTO = bb.c3
where aa.ID_EQUIPO = bb.c1;

select * from equipos where id_equipo = 13456;
select * from eq_mantenimiento_equipos where equipo_id = 13456;

13456

