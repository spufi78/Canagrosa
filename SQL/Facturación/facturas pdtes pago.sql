 SELECT fp_id,b.NOMBRE , sum((a.total - (a.total * a.descuento) /100)+ (((a.total - ((a.total * a.descuento) /100)) * a.IVA) /100)) AS TOTAL
-- SELECT a.ID_DOC,a.NUMERO,a.FECHA_FACTURA,b.DIAS,date_add(a.FECHA_FACTURA, INTERVAL b.DIAS DAY) as FECHA_VENCIMIENTO,
--       a.TOTAL,a.DESCUENTO,a.IVA,(a.total * a.descuento) /100 as DTO,
--      (a.total - (a.total * a.descuento) /100)+ (((a.total - ((a.total * a.descuento) /100)) * a.IVA) /100) as TOTAL
 FROM docs_pago a
 left join formas_pago b on a.FP_ID = b.ID_FP 
where a.TIPO = 2
  and a.PAGADO = 0
  and a.ANULADO = 0
  and year(a.FECHA_FACTURA) >= 2015
  and date_add(a.FECHA_FACTURA, INTERVAL b.DIAS DAY) between '2015-09-01' and '2015-09-30'
--  and a.ID_DOC = 18924
 group by fp_id