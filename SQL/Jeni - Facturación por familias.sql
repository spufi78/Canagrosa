select d.nombre,c.NOMBRE, sum(b.PRECIO)
from docs_pago a, docs_pago_conceptos b, familias c, clientes d
where a.ID_DOC = b.DOC_ID 
  and b.FAMILIA_ID  = c.ID_FAMILIA 
  and a.ANULADO = 0 and a.TIPO = 2 
  and a.FECHA_FACTURA >= '2013-01-01' and a.FECHA_FACTURA <= '2013-12-31'
  and a.CLIENTE_ID = d.ID_CLIENTE 
  and d.AIRBUS = 1
group by d.nombre,c.NOMBRE
  

