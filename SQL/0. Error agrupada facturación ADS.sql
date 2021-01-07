SELECT MONTHNAME(a.fecha_factura),concat(lpad(a.numero,4,'0000'),'/',YEAR(a.FECHA_FACTURA)) AS DELIVERY
      ,d0.descripcion AS PLANTA,d1.descripcion AS ENSAYO,d2.descripcion AS PROGRAMA,d3.descripcion AS SECCION,d4.descripcion AS FLUIDO
 		,IF(isnull(d5.DESCRIPCION),c.REFERENCIA_CLIENTE,d5.DESCRIPCION) AS FACILITY
      ,REPLACE(IF(b.CODIGO <> '',b.TIPO_ANALISIS,tdet.NOMBRE),'*','') AS ANALYSIS
		,IF(b.CODIGO <> '',b.CODIGO,IF(isnull(det.ID_DETERMINACION),b.codigo,dpm2.CODIGO)) AS TARIFA
      ,IF(b.CODIGO <> '',b.PRECIO,IF(isnull(det.ID_DETERMINACION),b.PRECIO,dpm2.PRECIO)) AS COST
      ,COUNT(DISTINCT b.DOC_ID,b.MUESTRA_ID,b.ORDEN) AS SAMPLES
      ,IF(b.CODIGO <> '',b.PRECIO,IF(isnull(det.ID_DETERMINACION),b.PRECIO,dpm2.PRECIO)) * COUNT(DISTINCT b.DOC_ID,b.MUESTRA_ID,b.ORDEN) AS IMPORTE
		,IF (c.OP_REPETICION = 1,'REPETICION',IF(c.OP_NORUTINARIA = 0,'SI','NO')) AS PLANED,group_concat(distinct b.MUESTRA_ID)
from docs_pago a
inner join docs_pago_muestras b on a.id_doc = b.doc_id and b.DETERMINACION_ID = 0
inner join muestras c on c.id_muestra = b.muestra_id
inner join muestras_airbus d ON c.id_muestra = d.muestra_id
inner join clientes e on c.cliente_id = e.id_cliente
inner join tipos_muestra f on c.tipo_muestra_id = f.id_tipo_muestra
inner join tipos_analisis g on c.tipo_analisis_id = g.id_tipo_analisis
left join docs_pago_muestras dpm2 on a.id_doc = dpm2.doc_id AND dpm2.MUESTRA_ID = c.ID_MUESTRA and dpm2.DETERMINACION_ID > 0
left join determinaciones det ON det.ID_DETERMINACION = dpm2.DETERMINACION_ID 
left join tipos_determinacion tdet ON det.tipo_determinacion_id = tdet.ID_TIPO_DETERMINACION
left join decodificadora d0 on d0.codigo = 600 and d0.valor = e.plant_id
left join decodificadora d1 on d1.codigo = 601 and d1.valor = d.ensayo_id
left join decodificadora d2 on d2.codigo = 602 and d2.valor = d.programa_id
left join decodificadora d3 on d3.codigo = 603 and d3.valor = d.section_id
left join decodificadora d4 on d4.codigo = 604 and d4.valor = d.fluid_id
left join decodificadora d5 on d5.codigo = 605 and d5.valor = d.facility_id
 WHERE a.ID_DOC = 35197
  and dpm2.codigo = 'WP1-EF-00001'
  and b.muestra_id = 348045
  group by 1,2,3,4,5,6,7,8,9,10,11,14
  order by group_concat(distinct b.MUESTRA_ID);
  
  select * from determinaciones where muestra_id = 348045;
--  SELECT * FROM MUESTRAS WHERE ID_MUESTRA = 348045;
  select * from docs_pago_muestras where doc_id = 35197 and codigo = 'WP1-EF-00001' order by muestra_id;
