select distinct bb.*,cc.* 
from docs_pago aa, (
 select a.NUMERO_EQUIPO_CLIENTE c3,a.NOMBRE as EQUIPO,a.DESCRIPCION,c.NOMBRE,b.FECHA_ACTUAL c1,a.CLIENTE_ID c2
  from equipos a
  left join eq_calibracion_equipos b on a.ID_EQUIPO = b.EQUIPO_ID
  left join clientes c on a.CLIENTE_ID = c.ID_CLIENTE
 where a.CLIENTE_ID <> 0
--  and a.NUMERO_EQUIPO_CLIENTE = 'SM4A09592'
  and year(b.FECHA_ACTUAL) >= 2018) bb,
  docs_pago_conceptos cc 
where aa.TIPO = 2 and aa.FECHA_FACTURA between date_sub(bb.c1, interval 2 month) and date_add(bb.c1, interval 2 month) 
-- and aa.CLIENTE_ID = bb.c2
and aa.ID_DOC = cc.DOC_ID 
 and cc.DESCRIPCION like concat('%',bb.c3,'%') 
 and bb.c1 = cc.FECHA 