select a.ID_MUESTRA,a.FECHA_RECEPCION,a.HORA_RECEPCION,a.FECHA_CIERRE,a.HORA_CIERRE, 
CONCAT(datediff(a.FECHA_CIERRE, a.FECHA_RECEPCION), ' días, ',DATE_FORMAT(sec_to_time(timestampdiff(second,concat(a.FECHA_RECEPCION,' ',a.HORA_RECEPCION),concat(a.FECHA_CIERRE ,' ',a.HORA_CIERRE ))),'%H:%i:%s'), ' horas')
from muestras a 
where a.FECHA_CIERRE <> '0000-00-00' 
order by id_muestra desc
limit 100 


ALTER TABLE `ca_documentos`
	ADD COLUMN `MTL` INT(1) NOT NULL DEFAULT '0' AFTER `NADCAP`;
ALTER TABLE `ca_normas`
	ADD COLUMN `MTL` INT(1) NOT NULL DEFAULT '0' AFTER `NADCAP`;
