select distinct b.ID_CALIBRACION,a.NUMERO_EQUIPO_CLIENTE,a.NOMBRE,a.DESCRIPCION,c.NOMBRE,b.FECHA_ACTUAL
,d.*
from equipos a
inner join eq_calibracion_equipos b on a.ID_EQUIPO = b.EQUIPO_ID and year(b.FECHA_ACTUAL) >= 2018
inner join clientes c on a.CLIENTE_ID = c.ID_CLIENTE
inner join docs_pago e on e.TIPO = 2 and e.FECHA_FACTURA between date_sub(b.FECHA_ACTUAL, interval 1 month) and date_add(b.FECHA_ACTUAL, interval 1 month) and e.CLIENTE_ID = e.CLIENTE_ID 
inner join docs_pago_conceptos d on d.DOC_ID = e.ID_DOC and d.DESCRIPCION like concat('%',a.NUMERO_EQUIPO_CLIENTE,'%') 
where a.CLIENTE_ID <> 0
  and a.NUMERO_EQUIPO_CLIENTE = 'SM4A09293'
--  and year(b.FECHA_ACTUAL) >= 2018
--  and not isnull(d.ID_CONCEPTO)
 