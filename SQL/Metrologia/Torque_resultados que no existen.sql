select a.*,d.*
 from torque_previa a
 inner join geslab_canagrosa.eq_calibracion_equipos b on a.CALIBRACION_ID = b.ID_CALIBRACION 
 inner join geslab_canagrosa.equipos c on b.EQUIPO_ID = c.ID_EQUIPO 
 left join torque_resultados d on convert(c.NUMERO_EQUIPO_CLIENTE USING latin1) = convert(d.CODIGO USING latin1) and b.FECHA_ACTUAL = d.FECHA_CALIBRACION 
where ISNULL(d.CODIGO) and b.ESTADO = 2
  and a.TIPO_ID = 0