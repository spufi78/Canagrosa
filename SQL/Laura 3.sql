select f.nombre,c.nombre, month(d.FECHA_FACTURA) , sum(dpc.precio)
from docs_pago d, docs_pago_conceptos dpc, clientes c, FAMILIAS f
where d.ID_DOC = dpc.doc_id
and d.FECHA_FACTURA >= '2011-01-01' and d.FECHA_FACTURA <= '2011-12-31'
and d.ANULADO = 0
and d.cliente_id = c.ID_CLIENTE 
and dpc.FAMILIA_ID  = f.id_familia
group by f.nombre,c.nombre, month(d.FECHA_FACTURA)