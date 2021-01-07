select m.ID_GENERAL, m.FECHA_RECEPCION, tm.NOMBRE, ce.*
from muestras m, tipos_muestra tm, ce_resultados ce
where m.ANNO = 2013
  and m.TIPO_MUESTRA_ID = tm.ID_TIPO_MUESTRA 
  and tm.NOMBRE like '%ADHERENCIA%'
  and m.ID_MUESTRA = ce.MUESTRA_ID 
order by ce.MUESTRA_ID,ce.PROBETA,ce.AREA 