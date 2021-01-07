select year(a.FECHA_APERTURA),c.NOMBRE,b.CANTIDAD, count(*) as N_BOTES, sum(b.PRECIO / b.CANTIDAD_UNIDAD_PEDIDO) as PRECIO
 from botes_ex a, tipos_bote_ex b, tipos_reactivo_ex c
where year(a.FECHA_APERTURA) >= 2018 -- between '2019-01-01' and '2019-12-31'
and a.TIPO_BOTE_EX_ID = b.ID_TIPO_BOTE_EX 
and b.TIPO_REACTIVO_EX_ID = c.ID_TIPO_REACTIVO_EX 
and b.TIPO_M_REFERENCIA_ID = 6
group by 1,2,3