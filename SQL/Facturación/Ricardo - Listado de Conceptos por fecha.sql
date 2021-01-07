select f.NOMBRE,cli.NOMBRE, b.FECHA,
concat(c.NUMERO,' (',year(c.FECHA_FACTURA),')') as NUMERO_FACTURA, 
CASE c.TIPO WHEN 1 THEN 'ALBARAN' WHEN 2 THEN 'FACTURA' ELSE 'ABONO' END  AS TIPO,
c.FECHA_FACTURA,b.TOTAL 
from docs_pago_conceptos b, docs_pago c, familias f, clientes cli
where c.FECHA_FACTURA >= '2017-01-01' and c.ANULADO = 0
and b.DOC_ID = c.ID_DOC 
and c.TIPO > 1
and b.FAMILIA_ID = f.ID_FAMILIA 
and c.CLIENTE_ID_FACTURA  = cli.ID_CLIENTE 
order by c.FECHA_FACTURA 
