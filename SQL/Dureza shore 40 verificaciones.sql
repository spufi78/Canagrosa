select a.ID_VERIFICACION, d.DESCRIPCION , b.descripcion, a.FECHA_ACTUAL, c.NOMBRE, b.resultado 
from eq_verificacion_equipos a, eq_verificacion_parametros_resultados b, usuarios c, eq_periodicidad d
where a.equipo_id = 832
and a.ID_VERIFICACION = b.verificacion_id 
and a.VERIFICADOR_INTERNO_ID = c.ID_EMPLEADO 
and a.PERIODICIDAD_ID = d.ID_PERIODICIDAD 
and a.FECHA_ACTUAL >= '2013-01-01' and a.FECHA_ACTUAL <= '2013-12-31'
-- and b.descripcion = '40'