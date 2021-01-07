select b.NOMBRE,d.NOMBRE, c.PRECIO, c.CANTIDAD, sum(a.CANTIDAD), sum(a.CANTIDAD_RECIBIDA)
from pedidos_bote_ex a, proveedores b, tipos_bote_ex c, tipos_reactivo_ex d
where a.TIPO_BOTE_EX_ID = c.ID_TIPO_BOTE_EX 
  and c.TIPO_REACTIVO_EX_ID = d.ID_TIPO_REACTIVO_EX 
  and c.PROVEEDOR_ID = b.ID_PROVEEDOR 
  and year(a.FECHA) = 2012
  and a.ANULADO = 0
group by b.NOMBRE,d.NOMBRE, c.PRECIO, c.CANTIDAD