select a.ANNO,c.codigo AS CODIGO, c.nombre AS TIPO_MUESTRA, td.NOMBRE as DETERMINACION, tr.NOMBRE, count(*) as CANTIDAD
 from muestras a 
inner join clientes b on a.CLIENTE_ID = b.ID_CLIENTE 
inner join tipos_muestra c on a.TIPO_MUESTRA_ID = c.ID_TIPO_MUESTRA 
inner join tipos_analisis ta on a.TIPO_ANALISIS_ID  = ta.ID_TIPO_ANALISIS 
inner join determinaciones det on a.ID_MUESTRA = det.MUESTRA_ID 
inner join tipos_determinacion td on det.TIPO_DETERMINACION_ID = td.ID_TIPO_DETERMINACION
inner join determinaciones_reactivos dr on det.ID_DETERMINACION = dr.DETERMINACION_ID and dr.TIPO = 'E'
inner join botes_ex bo on dr.BOTE_EX_ID = bo.ID_BOTE_EX
inner join tipos_bote_ex tb on bo.TIPO_BOTE_EX_ID = tb.ID_TIPO_BOTE_EX 
inner join tipos_reactivo_ex tr on tb.TIPO_REACTIVO_EX_ID = tr.ID_TIPO_REACTIVO_EX 
where b.IBERIA = 1
and a.anno in (2020)
and a.ANULADA = 0
and tr.ID_TIPO_REACTIVO_EX in (499,3193)
group by 1,2,3,4,5