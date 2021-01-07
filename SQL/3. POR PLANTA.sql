select b.nombre,count(*), sum(d.PRECIO)
  from muestras a, clientes b, docs_pago_muestras d, docs_pago e
where a.CLIENTE_ID = b.ID_CLIENTE 
--  and a.TIPO_MUESTRA_ID  = c.ID_TIPO_MUESTRA 
  and a.ANNO = 2012
  and b.AIRBUS = 1
  and a.ANULADA = 0
  AND a.ID_MUESTRA = d.MUESTRA_ID 
  AND d.DOC_ID = e.ID_DOC
  AND e.TIPO = 2
  group by b.NOMBRE 
--  and a.ANALISIS_MODIFICADO = 0