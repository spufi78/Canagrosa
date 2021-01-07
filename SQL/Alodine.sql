ALTER TABLE `clientes`
	ADD COLUMN `IDIOMA_FACTURA` INT(1) NOT NULL DEFAULT '0' AFTER `INTRA`;



INSERT INTO usuarios_firmas (id,file) VALUES (1,LOAD_FILE('Z:\Compartida\JGM.JPG'))


--CARGAR DECODIFICADORA.SPECIMENID

ALTER TABLE `muestras`
	ADD COLUMN `URGENTE` INT(1) NULL DEFAULT '0' AFTER `NADCAP`;

INSERT INTO `geslab_canagrosa`.`parametros` (`ID_PARAMETRO`, `DESCRIPCION`, `VALOR`) VALUES ('27', 'Iberia Recepción Plasma Datos Defecto', '3136;2;2;2;3;1;1;93');
UPDATE `geslab_canagrosa`.`parametros` SET `VALOR`='3136;2;2;3;3;1;1;93' WHERE  `ID_PARAMETRO`=27 AND `USUARIO`='';


ALTER TABLE `alodine_clientes`
	ADD COLUMN `PEDIDO_ID` INT(11) NULL DEFAULT '0' AFTER `PEDIDO`;

ALTER TABLE `alodine_planificacion`
	ADD COLUMN `PEDIDO_ID` INT(11) NULL DEFAULT '0' AFTER `PEDIDO`;

update alodine_clientes aa, (
select b.alodine_id,b.CLIENTE_ID,b.PEDIDO,c.ID_PEDIDO 
from alodine a
left join alodine_clientes b on a.ID_ALODINE = b.ALODINE_ID 
left join clientes_pedidos c on c.CODIGO like concat('%',b.PEDIDO,'%') and b.CLIENTE_ID = c.CLIENTE_ID 
left join clientes d on b.CLIENTE_ID = d.ID_CLIENTE 
where b.pedido <> '' and b.PEDIDO_ID = 0
  and not isnull(c.ID_PEDIDO)) bb
set aa.PEDIDO_ID = bb.ID_PEDIDO
where aa.ALODINE_ID = bb.ALODINE_ID
and aa.CLIENTE_ID = bb.CLIENTE_ID
and aa.PEDIDO = bb.PEDIDO;

update alodine_planificacion aa, (
select b.LOTE_ID,b.CLIENTE_ID,b.PEDIDO,c.ID_PEDIDO 
from alodine_lotes a
left join alodine_planificacion b on a.ID_LOTE = b.LOTE_ID 
left join clientes_pedidos c on c.CODIGO like concat('%',b.PEDIDO,'%') and b.CLIENTE_ID = c.CLIENTE_ID 
left join clientes d on b.CLIENTE_ID = d.ID_CLIENTE 
left join alodine e on a.ALODINE_ID  = e.ID_ALODINE 
where b.pedido <> '' and b.PEDIDO_ID  = 0
  and not isnull(c.ID_PEDIDO)) bb
set aa.PEDIDO_ID = bb.ID_PEDIDO
where aa.LOTE_ID = bb.LOTE_ID
and aa.CLIENTE_ID = bb.CLIENTE_ID
and aa.PEDIDO = bb.PEDIDO;

-- SIN FACTURAR
select *
from alodine_lotes a
left join alodine_planificacion b on a.ID_LOTE = b.LOTE_ID 
left join clientes_pedidos c on c.CODIGO like concat('%',b.PEDIDO,'%') and b.CLIENTE_ID = c.CLIENTE_ID 
left join clientes d on b.CLIENTE_ID = d.ID_CLIENTE 
left join alodine e on a.ALODINE_ID  = e.ID_ALODINE 
where b.pedido <> '' and b.PEDIDO_ID  = 0 and b.DOC_ID = 0
  and isnull(c.ID_PEDIDO)