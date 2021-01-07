SELECT S.NOMBRE, TM.NOMBRE, COUNT(*), SUM(DPM.PRECIO)
FROM MUESTRAS M, TIPOS_MUESTRA TM, SECTORES S, DOCS_PAGO_MUESTRAS DPM, DOCS_PAGO DP
WHERE ANNO = 2011 AND M.ANULADA = 0 AND M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA
AND TM.SECTOR_ID = S.ID_SECTOR
AND DPM.MUESTRA_ID = M.ID_MUESTRA
AND DP.ID_DOC = DPM.DOC_ID
AND DP.TIPO = 2
GROUP BY S.NOMBRE, TM.NOMBRE


