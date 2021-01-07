update eq_operaciones_pendientes aa, (
SELECT a.ID_CVM c1,a.EQUIPO_ID c2,a.TIPO_CVM_ID c3, fecha_prevista, fecha_preaviso,a.PERIODICIDAD_ID,b.DIAS_PREAVISO,date_sub(fecha_prevista, interval b.DIAS_PREAVISO day) c4
FROM eq_operaciones_pendientes a, eq_periodicidad b
where a.periodicidad_id = b.ID_PERIODICIDAD 
) bb
set aa.FECHA_PREAVISO = bb.c4
where aa.ID_CVM = bb.c1
  and aa.EQUIPO_ID = bb.c2
  and aa.TIPO_CVM_ID = bb.c3;