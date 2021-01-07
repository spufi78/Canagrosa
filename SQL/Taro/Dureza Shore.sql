SELECT distinct m.ID_GENERAL, m.FECHA_RECEPCION, m.ID_PARTICULAR , 
sd1.VALOR_1, sd1.VALOR_2,
sd2.VALOR_1, sd2.VALOR_2,
sd3.VALOR_1, sd3.VALOR_2,
sd4.VALOR_1, sd4.VALOR_2,
sd5.VALOR_1, sd5.VALOR_2,
sd6.VALOR_1, sd6.VALOR_2
FROM muestras m, sellantes_resultados sr, 
sellantes_determinaciones sd1,
sellantes_determinaciones sd2,
sellantes_determinaciones sd3,
sellantes_determinaciones sd4,
sellantes_determinaciones sd5,
sellantes_determinaciones sd6
where m.ID_MUESTRA = sr.MUESTRA_ID 
and m.ANNO = 2012 and m.ANULADA = 0 and m.CERRADA = 1
and sr.TIPO_DETERMINACION_ID = 215
and m.ID_MUESTRA = sd1.MUESTRA_ID and sd1.CAMPO_ID = 3080
and m.ID_MUESTRA = sd2.MUESTRA_ID and sd2.CAMPO_ID = 3081
and m.ID_MUESTRA = sd3.MUESTRA_ID and sd3.CAMPO_ID = 3082
and m.ID_MUESTRA = sd4.MUESTRA_ID and sd4.CAMPO_ID = 3083
and m.ID_MUESTRA = sd5.MUESTRA_ID and sd5.CAMPO_ID = 3084
and m.ID_MUESTRA = sd6.MUESTRA_ID and sd6.CAMPO_ID = 3085