select a.ID_PAQUETE,a.CODIGO_SC,e.NOMBRE, a.FECHA_CREACION,d.NOMBRE 
from sc_paquetes a, sc_paquetes_detalle b, determinaciones c, tipos_determinacion d, proveedores e
where a.FECHA_CREACION >= '2013-01-01'
and a.ID_PAQUETE = b.PAQUETE_ID 
and b.DETERMINACION_ID = c.ID_DETERMINACION 
and c.TIPO_DETERMINACION_ID = d.ID_TIPO_DETERMINACION 
and a.SUBCONTRATA_ID = e.ID_PROVEEDOR 