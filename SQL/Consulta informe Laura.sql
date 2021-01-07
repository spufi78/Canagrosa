select d.NOMBRE,count(*)
from muestras a,  determinaciones c, tipos_determinacion d
where a.ID_MUESTRA = c.MUESTRA_ID
  and d.ID_TIPO_DETERMINACION = c.TIPO_DETERMINACION_ID 
  and a.TIPO_MUESTRA_ID = 80
  and a.FECHA_RECEPCION >= '2011-08-01' and a.FECHA_RECEPCION <= '2012-08-31'
  group by d.NOMBRE 