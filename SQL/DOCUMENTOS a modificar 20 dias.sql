select * from (
select 'PENDIENTE APROBACION MAS DE 20 DÍAS (EDICION > 1)',aa.CODIGO,aa.DESCRIPCION,aa.EDICION,aa.MODIFICACION_FECHA AS FECHA,concat(cc.NOMBRE,' ',cc.APELLIDOS) AS RESPONSABLE
 from ca_pnt aa, 
	(
	select a.id_documento c1,a.codigo c2,a.edicion c3,max(b.EDICION) c4
	from ca_documentos as a
	left join ca_pnt as b on a.id_documento = b.documento_id
	where estado_id = 22 and anulado = 0 and uso = 1
	group by 1,2,3
	) bb, usuarios cc
where aa.documento_id = bb.c1
  and aa.EDICION = bb.c4 
  and bb.c3 < bb.c4
  and aa.MODIFICACION_USUARIO_REVISION = cc.ID_EMPLEADO 
  and current_date > date_add(aa.MODIFICACION_FECHA, interval 20 day)
UNION
select 'PENDIENTE APROBACION MAS DE 1 AÑO (EDICION 1)',aa.CODIGO,aa.DESCRIPCION,aa.EDICION,aa.MODIFICACION_FECHA,concat(cc.NOMBRE,' ',cc.APELLIDOS)  from ca_pnt aa, 
	(
	select a.id_documento c1,a.codigo c2,a.edicion c3,max(b.EDICION) c4
	from ca_documentos as a
	left join ca_pnt as b on a.id_documento = b.documento_id
	where estado_id = 22 and anulado = 0 and uso = 1
	group by 1,2,3
	) bb, usuarios cc
where aa.documento_id = bb.c1
  and aa.EDICION = bb.c4 
  and bb.c3 = bb.c4
  and aa.MODIFICACION_USUARIO_REVISION = cc.ID_EMPLEADO 
  and current_date > date_add(aa.FECHA_CREACION , interval 1 YEAR)
UNION
select 'PENDIENTE DE REVISIÓN MAS DE 20 DIAS',aa.CODIGO,aa.DESCRIPCION,aa.EDICION,aa.MODIFICACION_FECHA AS FECHA,concat(cc.NOMBRE,' ',cc.APELLIDOS) AS RESPONSABLE
 from ca_pnt aa, 
	(
	select a.id_documento c1,a.codigo c2,a.edicion c3,max(b.EDICION) c4
	from ca_documentos as a
	left join ca_pnt as b on a.id_documento = b.documento_id
	where estado_id = 21 and anulado = 0 and uso = 1
	group by 1,2,3
	) bb, usuarios cc
where aa.documento_id = bb.c1
  and aa.EDICION = bb.c4 
  and bb.c3 = bb.c4
  and aa.MODIFICACION_USUARIO_REVISION = cc.ID_EMPLEADO 
  and current_date > date_add(aa.MODIFICACION_FECHA, interval 20 day)
UNION
select 'PENDIENTES DE MODIFICACION MAS DE 20 DIAS (EDICION > 1)',aa.CODIGO,aa.DESCRIPCION,aa.EDICION,aa.MODIFICACION_FECHA AS FECHA,concat(cc.NOMBRE,' ',cc.APELLIDOS) AS RESPONSABLE
 from ca_pnt aa, 
	(
	select a.id_documento c1,a.codigo c2,a.edicion c3,max(b.EDICION) c4
	from ca_documentos as a
	left join ca_pnt as b on a.id_documento = b.documento_id
	where estado_id = 20 and anulado = 0 and uso = 1
	group by 1,2,3
	) bb, usuarios cc
where aa.documento_id = bb.c1
  and aa.EDICION = bb.c4 
  and bb.c3 < bb.c4
  and aa.MODIFICACION_USUARIO_REVISION = cc.ID_EMPLEADO 
  and current_date > date_add(aa.MODIFICACION_FECHA, interval 20 day)
UNION
select 'PENDIENTES DE MODIFICACION MAS DE 1 AÑO (EDICION 1)',aa.CODIGO,aa.DESCRIPCION,aa.EDICION,aa.MODIFICACION_FECHA AS FECHA,concat(cc.NOMBRE,' ',cc.APELLIDOS) AS RESPONSABLE
 from ca_pnt aa, 
	(
	select a.id_documento c1,a.codigo c2,a.edicion c3,max(b.EDICION) c4
	from ca_documentos as a
	left join ca_pnt as b on a.id_documento = b.documento_id
	where estado_id = 20 and anulado = 0 and uso = 1
	group by 1,2,3
	) bb, usuarios cc
where aa.documento_id = bb.c1
  and aa.EDICION = bb.c4 
  and bb.c3 = bb.c4
  and aa.MODIFICACION_USUARIO_REVISION = cc.ID_EMPLEADO 
  and current_date > date_add(aa.MODIFICACION_FECHA, interval 1 YEAR)
UNION
select 'PENDIENTE DE CREACION MAS DE 1 AÑO (EDICION 1)',aa.CODIGO,aa.DESCRIPCION,aa.EDICION,aa.MODIFICACION_FECHA AS FECHA,concat(cc.NOMBRE,' ',cc.APELLIDOS) AS RESPONSABLE
 from ca_pnt aa, 
	(
	select a.id_documento c1,a.codigo c2,a.edicion c3,max(b.EDICION) c4
	from ca_documentos as a
	left join ca_pnt as b on a.id_documento = b.documento_id
	where estado_id = 19 and anulado = 0 and uso = 1
	group by 1,2,3
	) bb, usuarios cc
where aa.documento_id = bb.c1
  and aa.EDICION = bb.c4 
  and bb.c3 = bb.c4
  and aa.MODIFICACION_USUARIO_REVISION = cc.ID_EMPLEADO 
  and current_date > date_add(aa.MODIFICACION_FECHA, interval 1 YEAR)
) aaa
order by responsable