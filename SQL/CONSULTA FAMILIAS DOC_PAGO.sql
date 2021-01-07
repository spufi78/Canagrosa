ALTER TABLE `docs_pago_cobros`
	ADD COLUMN `FECHA_PREVISTA` DATE NOT NULL DEFAULT '0000-00-00' AFTER `OBSERVACIONES`;

ALTER TABLE `docs_pago_cobros`
	CHANGE COLUMN `FECHA` `FECHA` DATE NULL DEFAULT NULL AFTER `DOC_ID`,
	CHANGE COLUMN `HORA` `HORA` TIME NULL AFTER `FECHA`,
	CHANGE COLUMN `FECHA_PREVISTA` `FECHA_PREVISTA` DATE NULL DEFAULT NULL AFTER `OBSERVACIONES`;



select a.ID_DOC,a.NUMERO as c1,c.NOMBRE as c2,c.CC as c3,sum(b.TOTAL) as c4 from docs_pago a,docs_pago_conceptos b, familias c
where a.ID_DOC = b.DOC_ID 
and b.FAMILIA_ID = c.ID_FAMILIA 
and a.NUMERO in (289,290,291,292,293,295,296,297,298,302,303,304)
-- and a.NUMERO in (298)
and year(a.FECHA_FACTURA) = 2017
and a.TIPO = 1
group by 1,2,3
union
select a.ID_DOC,a.NUMERO as c1,c.NOMBRE as c2,c.CC as c3,sum(b.PRECIO)  as c4
from docs_pago a,docs_pago_muestras b, familias c, muestras d, tipos_muestra tm
where a.ID_DOC = b.DOC_ID 
and b.MUESTRA_ID = d.ID_MUESTRA 
and d.TIPO_MUESTRA_ID = tm.ID_TIPO_MUESTRA 
and tm.FAMILIA_ID = c.ID_FAMILIA 
and a.NUMERO in (289,290,291,292,293,295,296,297,298,302,303,304)
-- and a.NUMERO in (298)
and year(a.FECHA_FACTURA) = 2017
and a.TIPO = 1
and b.muestra_id <> 0
group by 1,2,3