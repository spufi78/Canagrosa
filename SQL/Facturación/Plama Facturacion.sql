select d.NOMBRE,'MICROESTRUCTURA BOND', count(distinct a.ID_MUESTRA)
 from muestras a, plasma_recepcion b, plasma_resultados c, clientes d
 where a.analisis_modificado = 5
  and a.ANULADA = 0
  and a.REFERENCIA_CLIENTE not like '%IRR%'
  and a.TIPO_MUESTRA_ID = 222
  and a.DOCUMENTO_PAGO = 0
  and d.IBERIA = 1
  and month(a.FECHA_RECEPCION) = 4 and a.anno = 2016
  and a.ID_MUESTRA = b.MUESTRA_ID 
  and a.ID_MUESTRA = c.MUESTRA_ID 
  and c.BATCH <> 'N/A'
  and a.CLIENTE_ID = d.ID_CLIENTE 
  and (c.MICROESTRUCTURA1 <> 2 or c.MICROESTRUCTURA2 <> 2 or c.MICROESTRUCTURA3 <> 2 or c.MICROESTRUCTURA4 <> 2 or c.MICROESTRUCTURA5 <> 2 or c.MICROESTRUCTURA6 <> 2)
  and c.TIPO = 1
union
select d.NOMBRE,'MICROESTRUCTURA TOP', count(distinct a.ID_MUESTRA)
 from muestras a, plasma_recepcion b, plasma_resultados c, clientes d
 where a.analisis_modificado = 5
  and a.ANULADA = 0
  and a.REFERENCIA_CLIENTE not like '%IRR%'
  and a.TIPO_MUESTRA_ID = 222
  and a.DOCUMENTO_PAGO = 0
  and d.IBERIA = 1
  and month(a.FECHA_RECEPCION) = 4 and a.anno = 2016
  and a.ID_MUESTRA = b.MUESTRA_ID 
  and a.ID_MUESTRA = c.MUESTRA_ID 
  and c.BATCH <> 'N/A'
  and a.CLIENTE_ID = d.ID_CLIENTE 
  and (c.MICROESTRUCTURA1 <> 2 or c.MICROESTRUCTURA2 <> 2 or c.MICROESTRUCTURA3 <> 2 or c.MICROESTRUCTURA4 <> 2 or c.MICROESTRUCTURA5 <> 2 or c.MICROESTRUCTURA6 <> 2)
  and c.TIPO = 2
union
select d.NOMBRE,'TRACCION BOND',count(*)
 from muestras a, plasma_recepcion b, plasma_resultados c, clientes d
 where a.analisis_modificado = 5
  and a.ANULADA = 0
  and a.REFERENCIA_CLIENTE not like '%IRR%'
  and a.TIPO_MUESTRA_ID = 222
  and a.DOCUMENTO_PAGO = 0
  and d.IBERIA = 1
  and month(a.FECHA_RECEPCION) = 4 and a.anno = 2016
  and a.ID_MUESTRA = b.MUESTRA_ID 
  and a.ID_MUESTRA = c.MUESTRA_ID 
  and c.BATCH <> 'N/A'
  and a.CLIENTE_ID = d.ID_CLIENTE 
  and c.TRACCION_RES  <> ''
  and c.TIPO = 1
group by 1
UNION
select d.NOMBRE,'TRACCION TOP',count(*)
 from muestras a, plasma_recepcion b, plasma_resultados c, clientes d
 where a.analisis_modificado = 5
  and a.ANULADA = 0
  and a.REFERENCIA_CLIENTE not like '%IRR%'
  and a.TIPO_MUESTRA_ID = 222
  and a.DOCUMENTO_PAGO = 0
  and d.IBERIA = 1
  and month(a.FECHA_RECEPCION) = 4 and a.anno = 2016
  and a.ID_MUESTRA = b.MUESTRA_ID 
  and a.ID_MUESTRA = c.MUESTRA_ID 
  and c.BATCH <> 'N/A'
  and a.CLIENTE_ID = d.ID_CLIENTE 
  and c.TRACCION_RES  <> ''
  and c.TIPO = 2
group by 1
UNION
select d.NOMBRE,'MACRO DUREZA BOND',count(*)
 from muestras a, plasma_recepcion b, plasma_resultados c, clientes d
 where a.analisis_modificado = 5
  and a.ANULADA = 0
  and a.REFERENCIA_CLIENTE not like '%IRR%'
  and a.TIPO_MUESTRA_ID = 222
  and a.DOCUMENTO_PAGO = 0
  and d.IBERIA = 1
  and month(a.FECHA_RECEPCION) = 4 and a.anno = 2016
  and a.ID_MUESTRA = b.MUESTRA_ID 
  and a.ID_MUESTRA = c.MUESTRA_ID 
  and c.BATCH <> 'N/A'
  and a.CLIENTE_ID = d.ID_CLIENTE 
  and c.MACRO_DUREZA_RES  <> ''
  and c.TIPO = 1
group by 1
UNION
select d.NOMBRE,'MACRO DUREZA TOP',count(*)
 from muestras a, plasma_recepcion b, plasma_resultados c, clientes d
 where a.analisis_modificado = 5
  and a.ANULADA = 0
  and a.REFERENCIA_CLIENTE not like '%IRR%'
  and a.TIPO_MUESTRA_ID = 222
  and a.DOCUMENTO_PAGO = 0
  and d.IBERIA = 1
  and month(a.FECHA_RECEPCION) = 4 and a.anno = 2016
  and a.ID_MUESTRA = b.MUESTRA_ID 
  and a.ID_MUESTRA = c.MUESTRA_ID 
  and c.BATCH <> 'N/A'
  and a.CLIENTE_ID = d.ID_CLIENTE 
  and c.MACRO_DUREZA_RES  <> ''
  and c.TIPO = 2
group by 1
UNION
select d.NOMBRE,'MICRO DUREZA BOND',count(*)
 from muestras a, plasma_recepcion b, plasma_resultados c, clientes d
 where a.analisis_modificado = 5
  and a.ANULADA = 0
  and a.REFERENCIA_CLIENTE not like '%IRR%'
  and a.TIPO_MUESTRA_ID = 222
  and a.DOCUMENTO_PAGO = 0
  and d.IBERIA = 1
  and month(a.FECHA_RECEPCION) = 4 and a.anno = 2016
  and a.ID_MUESTRA = b.MUESTRA_ID 
  and a.ID_MUESTRA = c.MUESTRA_ID 
  and c.BATCH <> 'N/A'
  and a.CLIENTE_ID = d.ID_CLIENTE 
  and c.MICRO_DUREZA_RES  <> ''
  and c.TIPO = 1
group by 1
UNION
select d.NOMBRE,'MICRO DUREZA TOP',count(*)
 from muestras a, plasma_recepcion b, plasma_resultados c, clientes d
 where a.analisis_modificado = 5
  and a.ANULADA = 0
  and a.REFERENCIA_CLIENTE not like '%IRR%'
  and a.TIPO_MUESTRA_ID = 222
  and a.DOCUMENTO_PAGO = 0
  and d.IBERIA = 1
  and month(a.FECHA_RECEPCION) = 4 and a.anno = 2016
  and a.ID_MUESTRA = b.MUESTRA_ID 
  and a.ID_MUESTRA = c.MUESTRA_ID 
  and c.BATCH <> 'N/A'
  and a.CLIENTE_ID = d.ID_CLIENTE 
  and c.MICRO_DUREZA_RES  <> ''
  and c.TIPO = 2
group by 1
UNION
select d.NOMBRE,'ESPESOR BOND',count(*)
 from muestras a, plasma_recepcion b, plasma_resultados c, clientes d
 where a.analisis_modificado = 5
  and a.ANULADA = 0
  and a.REFERENCIA_CLIENTE not like '%IRR%'
  and a.TIPO_MUESTRA_ID = 222
  and a.DOCUMENTO_PAGO = 0
  and d.IBERIA = 1
  and month(a.FECHA_RECEPCION) = 4 and a.anno = 2016
  and a.ID_MUESTRA = b.MUESTRA_ID 
  and a.ID_MUESTRA = c.MUESTRA_ID 
  and c.BATCH <> 'N/A'
  and a.CLIENTE_ID = d.ID_CLIENTE 
  and c.ESPESOR_RES <> ''
  and c.TIPO = 1
group by 1
UNION
select d.NOMBRE,'ESPESOR TOP',count(*)
 from muestras a, plasma_recepcion b, plasma_resultados c, clientes d
 where a.analisis_modificado = 5
  and a.ANULADA = 0
  and a.REFERENCIA_CLIENTE not like '%IRR%'
  and a.TIPO_MUESTRA_ID = 222
  and a.DOCUMENTO_PAGO = 0
  and d.IBERIA = 1
  and month(a.FECHA_RECEPCION) = 4 and a.anno = 2016
  and a.ID_MUESTRA = b.MUESTRA_ID 
  and a.ID_MUESTRA = c.MUESTRA_ID 
  and c.BATCH <> 'N/A'
  and a.CLIENTE_ID = d.ID_CLIENTE 
  and c.ESPESOR_RES <> ''
  and c.TIPO = 2
group by 1