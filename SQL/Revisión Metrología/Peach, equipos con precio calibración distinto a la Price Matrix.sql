select b.CODIGO,a.FECHA_CALIBRACION,b.PRECIO AS PRICE_MATRIX,a.INSITU,a.AJUSTE,a.PRECIO 
from calibraciones a
inner join equipos b on a.EQUIPO_ID = b.ID_EQUIPO 
where year(a.FECHA_CALIBRACION) = 2020
  and insitu = 0 and ajuste = 0 and a.PRECIO <> b.PRECIO 
UNION
select b.CODIGO,a.FECHA_CALIBRACION,b.PRECIO_AJUSTE as PRICE_MATRIX,a.INSITU,a.AJUSTE,a.PRECIO 
from calibraciones a
inner join equipos b on a.EQUIPO_ID = b.ID_EQUIPO 
where year(a.FECHA_CALIBRACION) = 2020
  and insitu = 0 and ajuste = 1 and a.PRECIO <> b.PRECIO_AJUSTE
UNION
select b.CODIGO,a.FECHA_CALIBRACION,b.PRECIO_INSITU as PRICE_MATRIX,a.INSITU,a.AJUSTE,a.PRECIO 
from calibraciones a
inner join equipos b on a.EQUIPO_ID = b.ID_EQUIPO 
where year(a.FECHA_CALIBRACION) = 2020
  and insitu = 1 and ajuste = 0 and a.PRECIO <> b.PRECIO_INSITU
UNION
select b.CODIGO,a.FECHA_CALIBRACION,b.PRECIO_INSITU_AJUSTE as PRICE_MATRIX,a.INSITU,a.AJUSTE,a.PRECIO 
from calibraciones a
inner join equipos b on a.EQUIPO_ID = b.ID_EQUIPO 
where year(a.FECHA_CALIBRACION) = 2020
  and insitu = 1 and ajuste = 0 and a.PRECIO <> b.PRECIO_INSITU_AJUSTE  