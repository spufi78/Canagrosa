SELECT distinct d.*,a.cliente_id,c.nombre,a.pedido fROM ALODINE_CLIENTES a
inner join clientes c on a.cliente_id = c.id_cliente
inner join alodine d on a.alodine_id = d.ID_ALODINE 
left join clientes_pedidos b on a.cliente_id = b.cliente_id and a.pedido = b.codigo
WHERE PEDIDO <> ''
and a.cliente_id = 2813
and isnull(b.ID_PEDIDO);

select * from clientes_pedidos where cliente_id = 2813;