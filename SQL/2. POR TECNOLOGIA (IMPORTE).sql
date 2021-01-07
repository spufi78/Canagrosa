select c.NOMBRE,sum(d.PRECIO)
  from muestras a, clientes b , tipos_muestra c, docs_pago_muestras d
where a.CLIENTE_ID = b.ID_CLIENTE 
  and a.TIPO_MUESTRA_ID  = c.ID_TIPO_MUESTRA 
  and a.ANNO = 2011
  and b.AIRBUS = 1
  and a.ANULADA = 0
  and a.ID_MUESTRA = d.MUESTRA_ID 
group by c.NOMBRE 
--  and a.ANALISIS_MODIFICADO = 0
 