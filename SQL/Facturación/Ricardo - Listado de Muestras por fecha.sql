select a.ID_MUESTRA, a.ID_GENERAL,a.ANNO, f.NOMBRE,tm.NOMBRE, cli.NOMBRE, a.FECHA_RECEPCION,a.FECHA_CIERRE,
concat(c.NUMERO,'/',year(c.FECHA_FACTURA)) as NUMERO_FACTURA, 
CASE c.TIPO WHEN 1 THEN 'ALBARAN' WHEN 2 THEN 'FACTURA' ELSE 'ABONO' END  AS TIPO,
c.FECHA_FACTURA,b.PRECIO 
from muestras a, docs_pago_muestras b, docs_pago c, tipos_muestra tm, familias f, clientes cli
where c.FECHA_FACTURA >= '2017-01-01' and c.ANULADO = 0
and a.ID_MUESTRA = b.MUESTRA_ID 
and b.MUESTRA_ID <> 0
and b.DOC_ID = c.ID_DOC 
and c.TIPO > 1
and a.TIPO_MUESTRA_ID = tm.ID_TIPO_MUESTRA 
and tm.FAMILIA_ID = f.ID_FAMILIA 
and c.CLIENTE_ID_FACTURA  = cli.ID_CLIENTE 
order by a.ANNO,a.ID_GENERAL