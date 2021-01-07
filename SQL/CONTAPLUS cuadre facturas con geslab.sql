select CAST(replace(a.concepto,'Factura de ventas N?','') as SIGNED),sum(eurohaber), cast(b.total - ((b.total * b.DESCUENTO) / 100) as SIGNED)
 from diario a left join geslab_canagrosa.docs_pago b on CAST(replace(a.concepto,'Factura de ventas N?','') as SIGNED) = b.NUMERO and b.TIPO = 2 and b.FECHA_FACTURA >='2015-01-01'
where a.subcta between 7000000 and 7999999
and a.concepto like 'Fact%'
group by CAST(replace(concepto,'Factura de ventas N?','') as SIGNED)
order by CAST(replace(concepto,'Factura de ventas N?','') as SIGNED)