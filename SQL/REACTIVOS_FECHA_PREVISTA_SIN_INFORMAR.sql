select b.NOMBRE, d.NOMBRE, a.CODIGO_PEDIDO_PROVEEDOR, concat(substr(a.CONFIRMADO,7,4),'-',substr(a.CONFIRMADO,4,2),'-',substr(a.CONFIRMADO,1,2)) AS FECHA_PREVISTA
     , b.DIAS_ENTREGA,  current_date, date_add(a.FECHA_PEDIDO,interval b.DIAS_ENTREGA DAY)
	  , if (date_add(a.FECHA_PEDIDO,interval b.DIAS_ENTREGA DAY) < current_date,1,0) AS CADUCADO
 from pedidos_bote_ex a, proveedores b, tipos_bote_ex c, tipos_reactivo_ex d
where a.recibido <> 1 and year(a.fecha) = 2016 
and a.ANULADO = 0
and a.confirmado = ''
and a.PROVEEDOR_ID = b.ID_PROVEEDOR 
and a.TIPO_BOTE_EX_ID = c.ID_TIPO_BOTE_EX 
and c.TIPO_REACTIVO_EX_ID = d.ID_TIPO_REACTIVO_EX 
