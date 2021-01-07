select distinct e.ID_EQUIPO, e.NOMBRE, c.FECHA_ACTUAL, d.descripcion,d.resultado
from muestras m, ce_recepcion_equipos b, eq_verificacion_equipos c, eq_verificacion_parametros_resultados d, equipos e
where tipo_muestra_id = 58 and anno = 2013 and anulada = 0
and m.id_muestra = b.MUESTRA_ID 
and b.VERIFICACION_ID <> 0
and c.ID_VERIFICACION = b.VERIFICACION_ID 
and c.ID_VERIFICACION = d.verificacion_id 
and d.equipo_id = c.EQUIPO_ID 
and d.equipo_id = e.ID_EQUIPO 
and d.REALIZADO = 1
order by e.NOMBRE, d.descripcion,c.FECHA_ACTUAL