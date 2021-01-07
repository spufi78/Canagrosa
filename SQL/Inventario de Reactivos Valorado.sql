
 SELECT `decodificadora1`.`CODIGO`, `tipos_reactivo_ex1`.`NOMBRE`, `decodificadora1`.`DESCRIPCION`, `tipos_bote_ex1`.`CANTIDAD`, `tipos_bote_ex1`.`CODIGO_PROVEEDOR`, 
 `tipos_bote_ex1`.`PRECIO` / `tipos_bote_ex1`.CANTIDAD_UNIDAD_PEDIDO as PRECIO,
 `botes_ex1`.`ID_BOTE_EX`, `botes_ex1`.`LOTE`, `botes_ex1`.`ANULADO`, `botes_ex1`.`FINALIZADO`, `botes_ex1`.`FECHA_RECEPCION`, `botes_ex1`.`FECHA_CADUCIDAD`, `tipos_bote_ex1`.`CANTIDAD_UNIDAD_PEDIDO`,ce.NOMBRE 
 FROM   ((`geslab_canagrosa`.`botes_ex` `botes_ex1` 
 INNER JOIN `geslab_canagrosa`.`tipos_bote_ex` `tipos_bote_ex1` ON `botes_ex1`.`TIPO_BOTE_EX_ID`=`tipos_bote_ex1`.`ID_TIPO_BOTE_EX`) 
 INNER JOIN `geslab_canagrosa`.`tipos_reactivo_ex` `tipos_reactivo_ex1` ON `tipos_bote_ex1`.`TIPO_REACTIVO_EX_ID`=`tipos_reactivo_ex1`.`ID_TIPO_REACTIVO_EX`) INNER JOIN `geslab_canagrosa`.`decodificadora` `decodificadora1` ON `tipos_bote_ex1`.`TIPO_M_REFERENCIA_ID`=`decodificadora1`.`VALOR`
 inner join centros ce on botes_ex1.centro_id = ce.ID_CENTRO 
 WHERE  `decodificadora1`.`CODIGO`=30 AND `botes_ex1`.`ANULADO`=0 AND `botes_ex1`.`FINALIZADO`=0 
   and (botes_ex1.finalizado = 0 or botes_ex1.FECHA_FIN <= '2017-03-31')
	and botes_ex1.FECHA_RECEPCION <= '2017-03-31' 
	and `tipos_bote_ex1`.TIPO_M_REFERENCIA_ID in (1,2,3,6,7)
 ORDER BY `decodificadora1`.`DESCRIPCION`, `tipos_reactivo_ex1`.`NOMBRE`


