select a.ID_GENERAL,a.ANNO,concat(c.CODIGO,'-',a.ID_particular) as codigo, a.FECHA_RECEPCION
-- ,a.TIPO_MUESTRA_ID
       ,c.NOMBRE as TIPO_MUESTRA,ta.NOMBRE as TIPO_ANALISIS,
       group_concat(distinct tdet.PROC_REF_EADS) as PROC_REF_EADS,
		 b.NOMBRE as CLIENTE, a.REFERENCIA_CLIENTE, a.HORA_RECEPCION, a.FECHA_FINALIZACION, a.HORA_CIERRE,
       if(e_bond.MICROESTRUCTURA1=1,1,0) AS BOND_MICROESTRUCTURA,
		 if(e_bond.TRACCION<>'' AND e_bond.TRACCION<>'--',1,0) AS BOND_TRACCION,
		 if(e_bond.MACRO_DUREZA<>'' AND e_bond.MACRO_DUREZA<>'--',1,0) AS BOND_MACRO_DUREZA,
		 if(e_bond.MICRO_DUREZA<>'' and e_bond.MICRO_DUREZA<>'--',1,0) AS BOND_MICRO_DUREZA,
	  	 if(e_bond.ESPESOR<>'' AND e_bond.ESPESOR<>'--',1,0) AS BOND_ESPESOR,
       if(e_top.MICROESTRUCTURA1=1,1,0) AS TOP_MICROESTRUCTURA,
		 if(e_top.TRACCION<>'' AND e_top.TRACCION<>'--',1,0) AS TOP_TRACCION,
		 if(e_top.MACRO_DUREZA<>'' AND e_top.MACRO_DUREZA<>'--',1,0) AS TOP_MACRO_DUREZA,
		 if(e_top.MICRO_DUREZA<>'' AND e_top.MICRO_DUREZA<>'--',1,0) AS TOP_MICRO_DUREZA,
	  	 if(e_top.ESPESOR<>'' AND e_top.ESPESOR<>'--',1,0) AS TOP_ESPESOR,
	  	 if(a.TIPO_MUESTRA_ID = 47,1,0) AS COMBUSTIBLE,
	  	 if(a.TIPO_MUESTRA_ID = 200,1,0) AS COMBUSTIBLE_AGRUPADOS,
	  	 if(a.TIPO_MUESTRA_ID IN (47),count(tdet.ID_TIPO_DETERMINACION),0) AS COMBUSTIBLE_DETERMINACIONES,
	  	 if(a.TIPO_MUESTRA_ID IN (200),count(tdet.ID_TIPO_DETERMINACION),0) AS COMBUSTIBLE_AGRUPADO_DETERMINACIONES,
	  	 if(a.TIPO_MUESTRA_ID = 24,1,0) AS AGUAS,
	  	 if(a.TIPO_MUESTRA_ID = 6,1,0) AS BAÑOS,
	  	 if(a.TIPO_MUESTRA_ID = 160,1,0) AS DUREZA,
	  	 if(a.TIPO_MUESTRA_ID = 80,1,0) AS FLUIDO_HIDRAULICO,
	  	 if(a.TIPO_MUESTRA_ID = 98,1,0) AS GLICOL,
	  	 if(a.TIPO_MUESTRA_ID = 118,1,0) AS TRACCION,
	  	 if(a.TIPO_MUESTRA_ID = 55,1,0) AS PESO_PELICULA,
	  	 if(a.TIPO_MUESTRA_ID = 100,1,0) AS ACEITE_MOTOR,
	  	 
	  	 if(a.TIPO_MUESTRA_ID NOT IN(222,47,200,24,6,160,80,98,118,55,100),1,0) AS OTROS
 from muestras a 
inner join clientes b on a.CLIENTE_ID = b.ID_CLIENTE 
inner join tipos_muestra c on a.TIPO_MUESTRA_ID = c.ID_TIPO_MUESTRA 
inner join tipos_analisis ta on a.TIPO_ANALISIS_ID  = ta.ID_TIPO_ANALISIS 
left join plasma_recepcion d on a.ID_MUESTRA = d.MUESTRA_ID 
left join plasma_resultados e_bond on a.ID_MUESTRA = e_bond.MUESTRA_ID and e_bond.TIPO = 1 
left join plasma_resultados e_top on a.ID_MUESTRA = e_top.MUESTRA_ID and e_top.TIPO = 2
left join determinaciones det on a.ID_MUESTRA = det.MUESTRA_ID AND det.RESULTADO <> '-' AND det.RESULTADO <> '--' AND det.RESULTADO <> ''
left join tipos_determinacion tdet on det.TIPO_DETERMINACION_ID = tdet.ID_TIPO_DETERMINACION 
where b.IBERIA = 1
and a.anno = 2019
-- AND a.ID_MUESTRA = 336127
-- and month(a.FECHA_RECEPCION) in (9,10)
-- and b.NOMBRE like '%HEA%'
and a.ANULADA = 0
group by a.ID_MUESTRA 