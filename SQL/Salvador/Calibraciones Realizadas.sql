SELECT year(a.FECHA_ACTUAL) AS ANNO,MONTH(a.FECHA_ACTUAL) Mes,
                         SUM(IF(a.TIPO_ID=1 AND b.CLIENTE_ID=0,1,0)) CALIBRACIONES_INTERNAS_CANAGROSA,
                         SUM(IF(a.TIPO_ID=1 AND b.CLIENTE_ID<>0,1,0)) CALIBRACIONES_INTERNAS_CLIENTE,
                 		 SUM(IF(a.TIPO_ID=2 AND b.CLIENTE_ID=0,1,0)) CALIBRACIONES_EXTERNAS_CANAGROSA, 
                		 SUM(IF(a.TIPO_ID=2 AND b.CLIENTE_ID<>0,1,0)) CALIBRACIONES_EXTERNAS_CLIENTE, 
                        COUNT(*) TOTAL_CALIBRACIONES
                        from geslab_canagrosa.eq_calibracion_equipos a left join geslab_canagrosa.equipos b on a.EQUIPO_ID = b.ID_EQUIPO
                        where year(a.FECHA_ACTUAL) in (2019,2020)
                      AND a.ESTADO <> 3 
GROUP BY 1,2
