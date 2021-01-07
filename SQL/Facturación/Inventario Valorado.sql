select d.NOMBRE,sum(c.PRECIO/c.CANTIDAD_UNIDAD_PEDIDO)
from rex_inventarios_botes a, botes_ex b, tipos_bote_ex c, tipos_reactivo_ex d
where inventario_id = 49
and a.BOTE_EX_ID = b.ID_BOTE_EX 
and b.TIPO_BOTE_EX_ID = c.ID_TIPO_BOTE_EX 
and c.TIPO_REACTIVO_EX_ID = d.ID_TIPO_REACTIVO_EX 
group by d.NOMBRE 