select dp.*
-- select sum(total - ((total * descuento) / 100) + ((total * iva) / 100))
from docs_pago dp
inner join clientes c on dp.CLIENTE_ID = c.ID_CLIENTE 
left join docs_pago_cobros dpc on dp.ID_DOC = dpc.DOC_ID 
where c.AIRBUS = 0
and dp.FECHA_FACTURA >= '2010-01-01' and dp.FECHA_FACTURA < '2014-12-31'
and dp.TIPO = 2 and dp.ANULADO = 0
and isnull(dpc.DOC_ID)