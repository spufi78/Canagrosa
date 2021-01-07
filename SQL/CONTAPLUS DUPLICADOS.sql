select * from diario a, diario b
where a.ASIEN <> b.ASIEN 
and a.SUBCTA = b.SUBCTA 
and a.EURODEBE = b.EURODEBE 
and a.FECHA = b.FECHA 
 and a.CONCEPTO = b.CONCEPTO 
and a.SUBCTA between 6000000 and 6999999