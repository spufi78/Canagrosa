delete aa.* from eq_operaciones_pendientes aa, (
select a.ID_CVM c1 from eq_operaciones_pendientes a
left join eq_calibracion_equipos b on a.ID_CVM=b.ID_CALIBRACION  and a.EQUIPO_ID = b.EQUIPO_ID 
where isnull(b.ID_CALIBRACION)
and a.TIPO_CVM_ID = 0
) bb
where aa.ID_CVM = bb.c1 AND aa.TIPO_CVM_ID = 0;
