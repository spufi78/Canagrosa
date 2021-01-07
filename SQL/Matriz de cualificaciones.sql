SELECT ca.NOMBRE, CA.CODIGO, e.NOMBRE, 
		 ec.FECHA_FIRMA_TECNICO, IF(ec.FECHA_FIRMA_TECNICO='1900-01-01',0,1) AS CUALIFICADO,
		 ec.FECHA_ULTIMA_RECUALIFICACION, 
       CASE 
		    WHEN LEFT(ca.nombre,5) = 'PNT C' THEN date_add(ec.FECHA_ULTIMA_RECUALIFICACION, interval 3 year)
		    ELSE date_add(ec.FECHA_ULTIMA_RECUALIFICACION, interval 2 year)
		 END AS FECHA_CADUCIDAD,
       CASE 
		    WHEN LEFT(ca.nombre,5) = 'PNT C' THEN IF(date_add(ec.FECHA_ULTIMA_RECUALIFICACION, interval 3 year) >= CURRENT_DATE,0,1)
		    ELSE IF(date_add(ec.FECHA_ULTIMA_RECUALIFICACION, interval 2 year) >= CURRENT_DATE,0,1)
		 END AS RECUALIFICACION_CADUCADA,
       ec.EN_HISTORICO,  
       ec.ES_FORMADOR, ec.FORMADOR_NO_CUALIFICADO
FROM empleados_cualificaciones ec
INNER JOIN empleados e ON ec.EMPLEADO_ID = e.ID_EMPLEADO
INNER JOIN ca_documentos ca ON ec.DOCUMENTO_ID = ca.ID_DOCUMENTO
INNER JOIN usuarios usu ON e.USUARIO_ID = usu.ID_EMPLEADO
WHERE 1 = 1 AND e.ESTADO_ID = 0 AND ca.anulado = 0 
-- AND ca.nombre like '%. Absorción%'
ORDER BY ca.NOMBRE,e.nombre 
