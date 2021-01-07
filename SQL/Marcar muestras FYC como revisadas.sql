insert into web_muestras_revision 
select m.id_muestra,m.ULT_EDICION_IMP,1,'',18,current_timestamp,0,'',18,current_timestamp,0,'',18,current_timestamp,0,'',18,current_timestamp from muestras m
inner join tipos_muestra tm on m.TIPO_MUESTRA_ID = tm.ID_TIPO_MUESTRA 
inner join clientes c on m.CLIENTE_ID = c.ID_CLIENTE 
left join web_muestras_revision wm on m.ID_MUESTRA = wm.MUESTRA_ID 
where m.ANNO = 2013
and m.ANULADA = 0
and tm.TIPO_ESPECIAL_ID = 4
and wm.MUESTRA_ID is null
and m.FECHA_RECEPCION < '2013-04-18'