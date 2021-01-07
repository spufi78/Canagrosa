VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHenkel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ListadoPendienteFacturar(fdesde As String, fhasta As String, cerradas As Boolean, ID_MUESTRA As Long) As ADODB.Recordset
    Dim consulta As String
    Dim filtro As String
   On Error GoTo ListadoPendienteFacturar_Error

    filtro = " WHERE 1 = 1"
    filtro = filtro & " and c.ANULADA = 0 and c.DOCUMENTO_PAGO = 0 "
    Dim oParametros As New clsParametros
    oParametros.Carga parametros.PARAM_HENKEL_TM, ""
    filtro = filtro & " AND c.TIPO_MUESTRA_ID IN (" & oParametros.getVALOR & ")"
    oParametros.Carga parametros.PARAM_HENKEL_TA, ""
    filtro = filtro & " AND c.TIPO_ANALISIS_ID IN (" & oParametros.getVALOR & ")"
    oParametros.Carga parametros.PARAM_HENKEL_CLIENTE, ""
    filtro = filtro & " AND c.CLIENTE_ID IN (" & oParametros.getVALOR & ")"
        
    If fdesde <> "" Then
        filtro = filtro & " AND c.FECHA_RECEPCION >='" & Format(fdesde, "yyyy-mm-dd") & "'"
    End If
    If fhasta <> "" Then
        filtro = filtro & " AND c.FECHA_RECEPCION <='" & Format(fhasta, "yyyy-mm-dd") & "'"
    End If
    If cerradas = False Then
        filtro = filtro & " AND c.CERRADA <> 0"
    End If
    If ID_MUESTRA <> 0 Then
        filtro = filtro & " AND c.ID_MUESTRA = " & ID_MUESTRA
    End If
'    consulta = "SELECT DISTINCT c.ID_MUESTRA,concat(d.CODIGO,'-',lpad(c.ID_PARTICULAR,4,'0')) as CODIGO,c.REFERENCIA_CLIENTE,c.FECHA_RECEPCION,c.ID_GENERAL  " & _
'               ", (select count(*) from ce_resultados aa where aa.muestra_id = c.ID_MUESTRA and aa.DOC_ID = 0) AS PROBETAS " & _
'               ", (select count(*) from ce_resultados aa where aa.muestra_id = c.ID_MUESTRA and aa.CONFORME = 1 and aa.DOC_ID = 0) AS CONFORMES " & _
'               ", (select count(*) from ce_resultados aa where aa.muestra_id = c.ID_MUESTRA and aa.CONFORME = 0 and aa.DOC_ID = 0) AS NO_CONFORMES " & _
'                " , (select IFNULL(sum(IF(aa.CONFORME = 0,0,IF(substr(aa.identificacion_cliente,14,1)='E', " & _
'                "                 IF(substr(aa.identificacion_cliente,12,1)='2',IFNULL(cc.PRECIO_E,0),IFNULL(bb.PRECIO_E,0)), " & _
'                "                      IF(substr(aa.identificacion_cliente,12,1)='2',IFNULL(cc.PRECIO_ABCD,0),IFNULL(bb.PRECIO_ABCD,0)) " & _
'                "                  ))),0) " & _
'                "   from ce_resultados aa " & _
'                "    left join henkel_precios bb on bb.CODIGO = substr(aa.identificacion_cliente,10,4) " & _
'                "    left join henkel_sets cc on cc.CODIGO = substr(aa.identificacion_cliente,10,4) and cc.DIMENSION = aa.DIMENSION " & _
'                "    where aa.MUESTRA_ID = c.ID_MUESTRA and aa.DOC_ID = 0 " & _
'                "  ) AS IMPORTE " & _
'               ", (select count(*) from ce_resultados aa where aa.MUESTRA_ID = c.ID_MUESTRA and aa.DOC_ID = 0 and substr(aa.identificacion_cliente,14,1)='E') AS PROBETAS_E " & _
'               ", (select count(*) from ce_resultados aa where aa.MUESTRA_ID = c.ID_MUESTRA and aa.DOC_ID = 0 and substr(aa.identificacion_cliente,14,1)<>'E') AS PROBETAS_ABCD " & _
'               ", d.NOMBRE " & _
'               "  FROM muestras c " & _
'               "  INNER JOIN tipos_muestra d ON c.TIPO_MUESTRA_ID = d.ID_TIPO_MUESTRA " & _
'               filtro & _
'               " ORDER BY c.ID_MUESTRA  "
    consulta = "SELECT DISTINCT c.ID_MUESTRA,concat(d.CODIGO,'-',lpad(c.ID_PARTICULAR,4,'0')) as CODIGO,c.REFERENCIA_CLIENTE,c.FECHA_RECEPCION,c.ID_GENERAL  " & _
               ", (select count(*) from ce_resultados aa where aa.muestra_id = c.ID_MUESTRA and aa.DOC_ID = 0) AS PROBETAS " & _
               ", (select count(*) from ce_resultados aa where aa.muestra_id = c.ID_MUESTRA and aa.CONFORME = 1 and aa.DOC_ID = 0) AS CONFORMES " & _
               ", (select count(*) from ce_resultados aa where aa.muestra_id = c.ID_MUESTRA and aa.CONFORME = 0 and aa.DOC_ID = 0) AS NO_CONFORMES " & _
               ", (SELECT IFNULL(SUM(IF(aa.CONFORME = 0,0, IFNULL(bb.PRECIO,0))),0) " & _
               "     FROM ce_resultados aa " & _
               "    INNER JOIN henkel_price bb ON aa.IDENTIFICACION_CLIENTE = bb.CODIGO and aa.DIMENSION = bb.DIMENSION " & _
               "    WHERE aa.MUESTRA_ID = c.ID_MUESTRA AND aa.DOC_ID = 0) AS IMPORTE " & _
               ", (SELECT COUNT(*) FROM ce_resultados aa " & _
               "    INNER JOIN henkel_price bb ON aa.IDENTIFICACION_CLIENTE = bb.CODIGO and aa.DIMENSION = bb.DIMENSION and bb.PRIMER = 0 " & _
               "    WHERE aa.MUESTRA_ID = c.ID_MUESTRA AND aa.DOC_ID = 0) AS PROBETAS_E " & _
               ", (SELECT COUNT(*) FROM ce_resultados aa " & _
               "    INNER JOIN henkel_price bb ON aa.IDENTIFICACION_CLIENTE = bb.CODIGO and aa.DIMENSION = bb.DIMENSION and bb.PRIMER = 1 " & _
               "    WHERE aa.MUESTRA_ID = c.ID_MUESTRA AND aa.DOC_ID = 0) AS PROBETAS_ABCD " & _
               ", d.NOMBRE " & _
               "  FROM muestras c " & _
               "  INNER JOIN tipos_muestra d ON c.TIPO_MUESTRA_ID = d.ID_TIPO_MUESTRA " & _
               filtro & _
               " ORDER BY c.ID_MUESTRA  "
    Set ListadoPendienteFacturar = datos_bd(consulta)

   On Error GoTo 0
   Exit Function

ListadoPendienteFacturar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListadoPendienteFacturar of Módulo de clase clsHenkel"
End Function

Public Function ListadoProbetas(MUESTRA_ID As Long) As ADODB.Recordset
    Dim consulta As String
    Dim filtro As String

   On Error GoTo ListadoProbetas_Error

    filtro = " WHERE a.MUESTRA_ID = " & MUESTRA_ID & " AND a.DOC_ID <= 0 "
'    consulta = "SELECT a.MUESTRA_ID,a.DESIGNACION,a.PROBETA,a.AREA,a.MATERIAL,a.DIMENSION,a.IDENTIFICACION_CLIENTE,a.IDENTIFICACION_CANAGROSA,a.CRITERIO_ACEPTACION,a.RESULTADO,a.FECHA,a.CONFORME,a.DOC_ID,IF(a.CONFORME = 0,0,IF(substr(identificacion_cliente,14,1)='E', " & _
'               " IF(substr(a.identificacion_cliente,12,1)='2',ifnull(c.PRECIO_E,0),b.PRECIO_E), " & _
'               " IF(substr(a.identificacion_cliente,12,1)='2',ifnull(c.PRECIO_ABCD,0),b.PRECIO_ABCD) " & _
'               " )) AS PRECIO " & _
'               "  FROM ce_resultados a " & _
'               " left join henkel_precios b on b.CODIGO = substr(identificacion_cliente,10,4) " & _
'               " left join henkel_sets c on c.CODIGO = substr(identificacion_cliente,10,4) and c.dimension = a.dimension " & _
'               filtro & _
'               " GROUP BY a.IDENTIFICACION_CLIENTE " & _
'               " ORDER BY a.IDENTIFICACION_CLIENTE "
    consulta = "SELECT a.MUESTRA_ID,a.DESIGNACION,a.PROBETA,a.AREA,a.MATERIAL,a.DIMENSION,a.IDENTIFICACION_CLIENTE,a.IDENTIFICACION_CANAGROSA,a.CRITERIO_ACEPTACION,a.RESULTADO,a.FECHA,a.CONFORME,a.DOC_ID " & _
               "    , IF(a.CONFORME = 0,0,IFNULL(b.PRECIO,0)) AS PRECIO " & _
               "  FROM ce_resultados a " & _
               " LEFT JOIN henkel_price b on a.IDENTIFICACION_CLIENTE = b.CODIGO and a.DIMENSION = b.dimension " & _
               filtro & _
               " ORDER BY 1,2,3 "
    Set ListadoProbetas = datos_bd(consulta)

   On Error GoTo 0
   Exit Function

ListadoProbetas_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListadoProbetas of Módulo de clase clsHenkel"
End Function


