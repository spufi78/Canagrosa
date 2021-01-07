select c.NOMBRE, b.NOMBRE, sum(d.PRECIO) 
from muestras a, tipos_muestra b, clientes c, docs_pago_muestras d
where b.nombre like '%inmersion%'
  and a.TIPO_MUESTRA_ID = b.ID_TIPO_MUESTRA 
  and a.CLIENTE_ID = c.id_cliente
  and c.ANULADO = 0 and c.AIRBUS = 1
  and a.ID_MUESTRA = d.MUESTRA_ID 
  and a.FECHA_RECEPCION >= '2013-07-01' and a.FECHA_RECEPCION <= '2013-12-30'
group by c.NOMBRE, b.nombre