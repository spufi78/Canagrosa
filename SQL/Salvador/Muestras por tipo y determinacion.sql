select a.anno,b.nombre as CLIENTE,c.NOMBRE as TIPO_MUESTRA,d.NOMBRE as DETERMINACION, count(*),sum(e.PRECIO)
  from muestras a, clientes b , tipos_muestra c, tipos_analisis d, docs_pago_muestras e, docs_pago f
where a.CLIENTE_ID = b.ID_CLIENTE 
  and a.TIPO_MUESTRA_ID  = c.ID_TIPO_MUESTRA 
  and a.TIPO_ANALISIS_ID = d.ID_TIPO_ANALISIS 
  and a.ANNO in (2019,2020)
--  and b.AIRBUS = 1
  and a.ANULADA = 0
  AND a.ID_MUESTRA = e.MUESTRA_ID 
  AND e.DOC_ID = f.ID_DOC
  AND f.TIPO = 2
 group by 1,2,3,4
--  and a.ANALISIS_MODIFICADO = 0