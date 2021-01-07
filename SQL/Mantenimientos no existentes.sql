SELECT *
FROM eq_mantenimiento_equipos a 
left join eq_operaciones_pendientes b on a.ID_MANTENIMIENTO = b.ID_CVM and b.TIPO_CVM_ID = 2
WHERE a.estado = 0 
-- AND equipo_id IN (57,61,568)
and isnull(b.ID_CVM)
ORDER BY id_mantenimiento