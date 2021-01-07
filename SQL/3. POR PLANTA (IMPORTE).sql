select b.nombre,sum(c.precio)
  from muestras a, clientes b , docs_pago_muestras c
where a.CLIENTE_ID = b.ID_CLIENTE 
  and a.id_muestra = c.MUESTRA_ID 
  and a.ANNO = 2010
  and b.AIRBUS = 1
  and a.ANULADA = 0
 group by b.NOMBRE 
--  and a.ANALISIS_MODIFICADO = 0