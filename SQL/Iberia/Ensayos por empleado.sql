select  m.ID_GENERAL, m.ANNO,concat(tm.CODIGO,'-',m.ID_PARTICULAR), m.FECHA_RECEPCION, tm.NOMBRE,m.REFERENCIA_CLIENTE,u.NOMBRE 
from muestras m, usuarios u, tipos_muestra tm
where m.ANULADA = 0 and m.CERRADA = 1 and m.CENTRO_ID = 3
and m.CERRADA_USUARIO = u.ID_EMPLEADO 
and m.TIPO_MUESTRA_ID = tm.ID_TIPO_MUESTRA 
order by u.NOMBRE, tm.NOMBRE, m.ID_MUESTRA 