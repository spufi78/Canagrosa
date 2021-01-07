select a.ANNO,c.codigo AS CODIGO, c.nombre AS TIPO_MUESTRA, td.NOMBRE as DETERMINACION, count(*) as CANTIDAD
 from muestras a 
inner join clientes b on a.CLIENTE_ID = b.ID_CLIENTE 
inner join tipos_muestra c on a.TIPO_MUESTRA_ID = c.ID_TIPO_MUESTRA 
inner join tipos_analisis ta on a.TIPO_ANALISIS_ID  = ta.ID_TIPO_ANALISIS 
inner join determinaciones det on a.ID_MUESTRA = det.MUESTRA_ID 
inner join tipos_determinacion td on det.TIPO_DETERMINACION_ID = td.ID_TIPO_DETERMINACION 
where b.IBERIA = 1
and a.anno in (2020)
and a.ANULADA = 0
group by 1,2,3,4