select l.NOMBRE, b.NOMBRE, m.ID_GENERAL, tm.NOMBRE,ta.NOMBRE,b.NOMBRE,c.NOMBRE,m.FECHA_RECEPCION,m.FECHA_CIERRE,td.NOMBRE,d.RESULTADO, wmr.muestra_id, b.LINEA_ID  
  From muestras m 
  inner join banos b on m.BANO_ID = b.ID_BANO
  inner join lineas l on b.LINEA_ID = l.ID_LINEA
  inner join clientes c on m.cliente_id = c.id_cliente
  inner join determinaciones d on m.ID_MUESTRA = d.MUESTRA_ID 
  inner join tipos_determinacion td on d.TIPO_DETERMINACION_ID = td.ID_TIPO_DETERMINACION 
  inner join tipos_muestra tm on m.TIPO_MUESTRA_ID = tm.ID_TIPO_MUESTRA 
  inner join tipos_analisis ta on m.TIPO_ANALISIS_ID = ta.ID_TIPO_ANALISIS
  left join web_muestras_revision wmr on m.ID_MUESTRA = wmr.MUESTRA_ID 
where m.TIPO_MUESTRA_ID  in (2,6)
  and m.anulada = 0
  and m.cerrada = 1
  and c.airbus = 1
  and m.ANNO = 2013
order by 1,2,3,4,5,6,7,8,9,10,11