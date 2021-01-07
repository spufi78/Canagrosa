select b.ID_GENERAL,b.FECHA_RECEPCION,b.HORA_RECEPCION,tm.NOMBRE,ta.NOMBRE,b.REFERENCIA_CLIENTE,td.NOMBRE 
from docs_pago_muestras a, muestras b, determinaciones c, tipos_muestra tm, tipos_analisis ta, tipos_determinacion td
where doc_id = 20317
  and a.MUESTRA_ID = b.ID_MUESTRA 
  and a.MUESTRA_ID <> 0  
  and b.ID_MUESTRA = c.MUESTRA_ID 
  and b.TIPO_MUESTRA_ID = tm.ID_TIPO_MUESTRA 
  and b.TIPO_ANALISIS_ID = ta.ID_TIPO_ANALISIS 
  and c.TIPO_DETERMINACION_ID  = td.ID_TIPO_DETERMINACION 
order by 1