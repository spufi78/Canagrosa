select * from (
select c.ID_GENERAL,concat(f.CODIGO,'-',c.ID_PARTICULAR),'BOND',c.FECHA_RECEPCION,a.MACRO_DUREZA,a.MACRO_DUREZA_RES, a.MACRO_DUREZA_SD,a.MACRO_DUREZA_POR,a.MACRO_DUREZA_PASS,e.MACRO_DUREZA_REQ 
from plasma_resultados a, plasma_recepcion b, muestras c, plasma_procesos d, plasma_ficha e, tipos_muestra f
where a.macro_dureza_res <> ''
and a.MUESTRA_ID = b.MUESTRA_ID 
and b.MUESTRA_ID = c.ID_MUESTRA 
and c.TIPO_MUESTRA_ID = f.ID_TIPO_MUESTRA 
and c.ANNO = 2016 and c.ANULADA = 0
and b.PROCESO_ID = d.ID_PROCESO 
and a.TIPO = 1
and d.BOND_COAT_FICHA_ID = e.ID_FICHA 
union
select c.ID_GENERAL,concat(f.CODIGO,'-',c.ID_PARTICULAR),'TOP',c.FECHA_RECEPCION,a.MACRO_DUREZA,a.MACRO_DUREZA_RES, a.MACRO_DUREZA_SD,a.MACRO_DUREZA_POR,a.MACRO_DUREZA_PASS,e.MACRO_DUREZA_REQ 
from plasma_resultados a, plasma_recepcion b, muestras c, plasma_procesos d, plasma_ficha e, tipos_muestra f
where a.macro_dureza_res <> ''
and a.MUESTRA_ID = b.MUESTRA_ID 
and b.MUESTRA_ID = c.ID_MUESTRA 
and c.TIPO_MUESTRA_ID = f.ID_TIPO_MUESTRA 
and c.ANNO = 2016 and c.ANULADA = 0
and b.PROCESO_ID = d.ID_PROCESO 
and a.TIPO = 2
and d.TOP_COAT_FICHA_ID = e.ID_FICHA 
) aa
order by 1,2,3