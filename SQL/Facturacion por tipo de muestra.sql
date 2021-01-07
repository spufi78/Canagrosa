select d.NOMBRE , sum(b.PRECIO) 
from docs_pago a, docs_pago_muestras b, muestras c, tipos_muestra d
where numero in (1854, 1529, 1530,1776,1774,1531)
and fecha_factura >= '2013-01-01'
and a.ID_DOC = b.DOC_ID 
and b.MUESTRA_ID  = c.ID_MUESTRA 
and c.TIPO_MUESTRA_ID = d.ID_TIPO_MUESTRA 
and b.MUESTRA_ID <> 0 
group by d.NOMBRE 