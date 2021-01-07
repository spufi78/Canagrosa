select c.NOMBRE,ta.NOMBRE, count(*),sum(d.PRECIO)
  from muestras a 
   inner join clientes b on a.CLIENTE_ID = b.ID_CLIENTE 
	inner join tipos_muestra c on a.TIPO_MUESTRA_ID  = c.ID_TIPO_MUESTRA 
	inner join docs_pago_muestras d on a.ID_MUESTRA = d.MUESTRA_ID 
	inner join docs_pago e on d.DOC_ID = e.ID_DOC
	left join tipos_analisis ta on a.TIPO_ANALISIS_ID  = ta.ID_TIPO_ANALISIS 
where a.FECHA_RECEPCION >='2011-06-01' and a.FECHA_RECEPCION <= '2012-05-31'
  and b.AIRBUS = 1
  and a.ANULADA = 0  
  AND e.TIPO = 2
 group by c.NOMBRE,ta.NOMBRE 
--  and a.ANALISIS_MODIFICADO = 0
 