SELECT sect.NOMBRE,fam.NOMBRE,area.DESCRIPCION,linea.DESCRIPCION,dp.NUMERO,dp.FECHA_FACTURA,c.NOMBRE, tm.nombre,m.ID_GENERAL, ta.nombre, dpm.PRECIO,m.BANO_ID
FROM muestras m
INNER JOIN tipos_muestra tm ON m.tipo_muestra_id = tm.id_tipo_muestra 
INNER JOIN tipos_analisis ta ON m.tipo_analisis_id = ta.id_tipo_analisis
INNER JOIN clientes c ON m.cliente_id = c.id_cliente 
INNER JOIN docs_pago_muestras dpm ON m.ID_MUESTRA = dpm.MUESTRA_ID 
INNER JOIN docs_pago dp ON dp.ID_DOC = dpm.DOC_ID AND dp.TIPO = 2
LEFT JOIN banos b ON m.BANO_ID = b.ID_BANO 
LEFT JOIN decodificadora area ON area.CODIGO = 76 AND b.AIRBUS_AREA_ID = area.VALOR  
LEFT JOIN decodificadora linea ON linea.CODIGO = 77 AND b.AIRBUS_LINEA_ID = linea.VALOR 
LEFT JOIN sectores sect ON tm.SECTOR_ID = sect.ID_SECTOR 
LEFT JOIN familias fam ON tm.FAMILIA_ID = fam.ID_FAMILIA 
WHERE m.fecha_recepcion >= '2011-01-01' 
AND m.fecha_recepcion <= '2012-02-07' 
AND m.anulada = 0 
AND m.cliente_id = 1656 
-- AND m.ANALISIS_MODIFICADO = 0
