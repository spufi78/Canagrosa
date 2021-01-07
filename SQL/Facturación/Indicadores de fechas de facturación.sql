Select distinct m.ID_MUESTRA,m.ID_GENERAL,m.ANNO, f.NOMBRE, c.NOMBRE , m.FECHA_RECEPCION,m.FECHA_CIERRE,DATEDIFF(m.FECHA_CIERRE,m.FECHA_RECEPCION) AS DIAS,
                max(dp.FECHA_FACTURA) AS FECHA_FACTURA,DATEDIFF(max(dp.FECHA_FACTURA),m.FECHA_CIERRE) AS DIAS,
                dpc.FECHA as FECHA_COBRO,DATEDIFF(dpc.FECHA,max(dp.FECHA_FACTURA)) AS DIAS
  from muestras m
  inner join docs_pago_muestras dpm ON m.ID_MUESTRA = dpm.MUESTRA_ID 
  inner join docs_pago dp ON dpm.DOC_ID = dp.ID_DOC
  inner join tipos_muestra tm ON m.TIPO_MUESTRA_ID = tm.ID_TIPO_MUESTRA 
  inner join familias f ON tm.FAMILIA_ID = f.ID_FAMILIA 
  inner join clientes c ON dp.CLIENTE_ID_FACTURA = c.ID_CLIENTE 
  left join docs_pago_cobros dpc on dp.ID_DOC = dpc.DOC_ID 
 where m.FECHA_RECEPCION >= '2016-03-01'
   and dp.FECHA_FACTURA >= '2016-03-01'
   and m.ANULADA = 0
   and m.REVISION_USUARIO <> 0
   and dpm.MUESTRA_ID <> 0
--   and not isnull(dpc.DOC_ID)
group by 1
