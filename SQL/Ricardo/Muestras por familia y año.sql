select month(a.FECHA_RECEPCION) AS MES,d.NOMBRE AS FAMILIA,count(*) AS CANTIDAD,sum(a.PRECIO) AS IMPORTE
from muestras a, clientes b, familias c, tipos_muestra d
where a.CLIENTE_ID = b.ID_CLIENTE 
--  and b.IBERIA = 1
  and a.CENTRO_ID = 3
  and a.ANNO = 2017
  and a.ANULADA = 0
  and a.TIPO_MUESTRA_ID <> 253
  and a.TIPO_MUESTRA_ID = d.ID_TIPO_MUESTRA 
  and c.ID_FAMILIA = d.FAMILIA_ID 
group by 1,2