select a.ID_GENERAL,a.ANNO,a.FECHA_RECEPCION,b.NOMBRE as CLIENTE, c.NOMBRE AS TIPO_MUESTRA,d.NOMBRE AS TIPO_ANALISIS,a.REFERENCIA_CLIENTE,e.IDENTIFICACION_CLIENTE
, SUBSTR(e.IDENTIFICACION_CLIENTE,14,1) as IDEN
, CASE WHEN SUBSTR(e.IDENTIFICACION_CLIENTE,14,1) <> 'E' THEN 1 ELSE 0 END AS 'PINTADA'
, CASE WHEN SUBSTR(e.IDENTIFICACION_CLIENTE,14,1) = 'E' THEN 1 ELSE 0 END AS 'NO PINTADA'
from muestras a
inner join clientes b on a.CLIENTE_ID = b.ID_CLIENTE
inner join tipos_muestra c on a.TIPO_MUESTRA_ID = c.ID_TIPO_MUESTRA
inner join tipos_analisis d on a.TIPO_ANALISIS_ID = d.ID_TIPO_ANALISIS
left join ce_resultados e on e.MUESTRA_ID = a.ID_MUESTRA 
where b.nombre like '%henkel%'
  and a.ANULADA = 0
  and a.TIPO_MUESTRA_ID = 294