select m.ID_GENERAL,m.FECHA_RECEPCION, tm.NOMBRE, ta.NOMBRE, 
SUBSTRING_INDEX( ce.resultado , '-', 1 ) as medidas,
SUBSTRING_INDEX(SUBSTRING_INDEX( ce.resultado , '-', 2 ),'-',-1) as resultado,
SUBSTRING_INDEX(SUBSTRING_INDEX( ce.resultado , '-', 3 ),'-',-1) as desviacion
from muestras m, tipos_muestra tm, tipos_analisis ta, ce_resultados ce
where m.TIPO_MUESTRA_ID = tm.ID_TIPO_MUESTRA 
and m.TIPO_ANALISIS_ID  = ta.ID_TIPO_ANALISIS 
and m.ID_MUESTRA = ce.MUESTRA_ID 
and tm.ID_TIPO_MUESTRA = 58
and m.ANNO >= 2013 and m.ANULADA = 0 