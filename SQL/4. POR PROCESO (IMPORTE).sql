select count(*)
  from muestras a, clientes b , tipos_muestra c, tipos_analisis d
where a.CLIENTE_ID = b.ID_CLIENTE 
  and a.TIPO_MUESTRA_ID  = c.ID_TIPO_MUESTRA 
  and a.TIPO_ANALISIS_ID = d.ID_TIPO_ANALISIS 
  and a.ANNO = 2011
  and b.AIRBUS = 1
  and a.ANULADA = 0
-- group by c.NOMBRE, d.nombre
--  and a.ANALISIS_MODIFICADO = 0