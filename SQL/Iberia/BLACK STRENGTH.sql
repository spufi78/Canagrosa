select concat(c.CODIGO,'-',a.ID_PARTICULAR), b.* from muestras a, plasma_traccion_p b, tipos_muestra c
where a.ANNO = 2016 and a.ANULADA = 0
  and a.ID_MUESTRA = b.MUESTRA_ID 
  and a.TIPO_MUESTRA_ID = c.ID_TIPO_MUESTRA 
  and b.IDENTIFICATION = 'B (Blank Strength)'
  
  
  