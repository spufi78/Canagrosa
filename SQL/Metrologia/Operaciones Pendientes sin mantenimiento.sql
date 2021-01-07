delete aa.* from eq_operaciones_pendientes aa, (
select a.ID_CVM c1,a.TIPO_CVM_ID c2 from eq_operaciones_pendientes a
left join eq_mantenimiento_equipos b on a.EQUIPO_ID = b.EQUIPO_ID and a.ID_CVM = b.ID_MANTENIMIENTO and b.ESTADO = 0
where a.TIPO_CVM_ID = 2
  and isnull(b.ID_MANTENIMIENTO)
  ) bb
where aa.ID_CVM = bb.c1
and aa.TIPO_CVM_ID = bb.c2 AND aa.TIPO_CVM_ID = 2;

delete aa.* from eq_operaciones_pendientes aa, (
select a.ID_CVM c1 from eq_operaciones_pendientes a
left join eq_mantenimiento_equipos b on a.ID_CVM =b.ID_MANTENIMIENTO   and a.EQUIPO_ID = b.EQUIPO_ID 
where isnull(b.ID_MANTENIMIENTO)
and a.TIPO_CVM_ID = 2
) bb
where aa.ID_CVM = bb.c1 AND aa.TIPO_CVM_ID = 2;

delete aa.* from eq_mantenimiento_equipos_acciones aa, (
select distinct a.MANTENIMIENTO_ID c1 from eq_mantenimiento_equipos_acciones a
left join eq_mantenimiento_equipos b on a.MANTENIMIENTO_ID =b.ID_MANTENIMIENTO 
where isnull(b.ID_MANTENIMIENTO)
) bb
where aa.MANTENIMIENTO_ID = bb.c1;


