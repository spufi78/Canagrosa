SELECT distinct a.ANNO,a.ID_GENERAL,CONCAT(d.CODIGO,'-',lpad(a.ID_PARTICULAR,4,0)) AS PARTICULAR, a.FECHA_RECEPCION, d.NOMBRE AS TIPO_MUESTRA
, e.NOMBRE AS TIPO_ANALISIS,f.PRODUCTO,i.NOMBRE AS ENSAYO,b.LOTE,h.RESULTADO 
FROM muestras a
INNER JOIN sellantes_recepcion b ON a.ID_MUESTRA = b.MUESTRA_ID
INNER JOIN clientes c ON a.CLIENTE_ID = c.ID_CLIENTE
INNER JOIN tipos_muestra d ON a.TIPO_MUESTRA_ID = d.ID_TIPO_MUESTRA 
INNER JOIN tipos_analisis e ON a.TIPO_ANALISIS_ID = e.ID_TIPO_ANALISIS 
INNER JOIN sellantes f ON b.SELLANTE_ID = f.ID_SELLANTE 
INNER JOIN sellantes_determinaciones g ON g.MUESTRA_ID = a.ID_MUESTRA 
INNER JOIN sellantes_resultados h ON h.MUESTRA_ID = a.ID_MUESTRA 
INNER JOIN tipos_determinacion i ON i.ID_TIPO_DETERMINACION = h.TIPO_DETERMINACION_ID 
-- WHERE b.SELLANTE_ID IN (17,42,100,109)
WHERE b.SELLANTE_ID IN (42,109) -- SIN LOS B2
AND a.ANNO >= 2018
AND i.NOMBRE LIKE '%DUREZA%'
AND c.AIRBUS = 1

