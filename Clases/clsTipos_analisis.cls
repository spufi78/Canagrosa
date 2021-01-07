VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_analisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE TIPOS_ANALISIS      *
'************************************************************
Private ID_TIPO_ANALISIS As Integer
Private nombre As String
Private NOMBRE_INGLES As String
Private NOMBRE_SPA As String
Private NORMALIZADO As Integer
Private PRECIO As String
Private TIPO_MUESTRA_ID As Integer
Private NORMATIVA As String
Private METODO As String
Private PARAMETRO As String
Private TARIFA_CODIGO_ID As Integer
Private revisar_factura As Integer
Private FACTURA_DETERMINACIONES As Integer
Private TIPO_TRIGO As Integer
Private DIAS_TRABAJO As Integer
Private ADJUNTOS As String
Private anulado As Integer
Public Property Let setID_TIPO_ANALISIS(ByVal dato As Integer)
    ID_TIPO_ANALISIS = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Let setNOMBRE_INGLES(ByVal dato As String)
    NOMBRE_INGLES = Trim(dato)
End Property
Public Property Let setNOMBRE_SPA(ByVal dato As String)
    NOMBRE_SPA = Trim(dato)
End Property
Public Property Let setNORMALIZADO(ByVal dato As Integer)
    NORMALIZADO = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
    PRECIO = Trim(dato)
End Property
Public Property Let setTIPO_MUESTRA_ID(ByVal dato As Integer)
    TIPO_MUESTRA_ID = dato
End Property
Public Property Let setANULADO(ByVal dato As Integer)
    anulado = dato
End Property
Public Property Let setNORMATIVA(ByVal dato As String)
    NORMATIVA = Trim(dato)
End Property
Public Property Let setMETODO(ByVal dato As String)
    METODO = Trim(dato)
End Property
Public Property Let setTARIFA_CODIGO_ID(ByVal dato As Integer)
    TARIFA_CODIGO_ID = Trim(dato)
End Property
Public Property Let setREVISAR_FACTURA(ByVal dato As Integer)
    revisar_factura = Trim(dato)
End Property
Public Property Let setFACTURA_DETERMINACIONES(ByVal dato As Integer)
    FACTURA_DETERMINACIONES = Trim(dato)
End Property
Public Property Let setTIPO_TRIGO(ByVal dato As Integer)
    TIPO_TRIGO = Trim(dato)
End Property
Public Property Let setDIAS_TRABAJO(ByVal dato As Integer)
    DIAS_TRABAJO = Trim(dato)
End Property
Public Property Let setADJUNTOS(ByVal dato As String)
    ADJUNTOS = Trim(dato)
End Property
Public Property Let setPARAMETRO(ByVal dato As String)
    PARAMETRO = Trim(dato)
End Property
Public Property Get getID_TIPO_ANALISIS() As Integer
    getID_TIPO_ANALISIS = ID_TIPO_ANALISIS
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getNOMBRE_INGLES() As String
    getNOMBRE_INGLES = NOMBRE_INGLES
End Property
Public Property Get getNOMBRE_SPA() As String
    getNOMBRE_SPA = NOMBRE_SPA
End Property
Public Property Get getNORMALIZADO() As Integer
    getNORMALIZADO = NORMALIZADO
End Property
Public Property Get getPRECIO() As String
    getPRECIO = PRECIO
End Property
Public Property Get getNORMATIVA() As String
    getNORMATIVA = NORMATIVA
End Property
Public Property Get getMETODO() As String
    getMETODO = METODO
End Property
Public Property Get getPARAMETRO() As String
    getPARAMETRO = PARAMETRO
End Property
Public Property Get getTIPO_MUESTRA_ID() As Integer
    getTIPO_MUESTRA_ID = TIPO_MUESTRA_ID
End Property
Public Property Get getANULADO() As Integer
    getANULADO = anulado
End Property
Public Property Get getTARIFA_CODIGO_ID() As Integer
    getTARIFA_CODIGO_ID = TARIFA_CODIGO_ID
End Property
Public Property Get getREVISAR_FACTURA() As Integer
    getREVISAR_FACTURA = revisar_factura
End Property
Public Property Get getFACTURA_DETERMINACIONES() As Integer
    getFACTURA_DETERMINACIONES = FACTURA_DETERMINACIONES
End Property
Public Property Get getTIPO_TRIGO() As Integer
    getTIPO_TRIGO = TIPO_TRIGO
End Property
Public Property Get getDIAS_TRABAJO() As Integer
    getDIAS_TRABAJO = DIAS_TRABAJO
End Property
Public Property Get getADJUNTOS() As String
    getADJUNTOS = ADJUNTOS
End Property
Public Function PrecioDelAnalisis(ByVal id_analisis As Integer) As Currency
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT precio FROM tipos_analisis WHERE id_tipo_analisis=" & id_analisis 'el id_analisis de tipo de analisis
    PrecioDelAnalisis = datos_bd(consulta).Fields("precio")
    Exit Function
fallo:
    MsgBox "No localizo el precio del análisis.", vbCritical, Err.Description
End Function
Public Function AnalisisAsociadosMuestra(ByVal ID_MUESTRA As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT id_tipo_analisis,nombre " & _
               "FROM tipos_analisis " & _
               "WHERE tipo_muestra_id=" & ID_MUESTRA & " and anulado = 0 order by nombre"
    Set AnalisisAsociadosMuestra = datos_bd(consulta)
End Function
Public Function NombreAnalisis(id_analisis As Integer) As String
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT nombre FROM tipos_analisis WHERE id_tipo_analisis=" & id_analisis
    NombreAnalisis = datos_bd(consulta).Fields("nombre")
    Exit Function
fallo:
    NombreAnalisis = ""
End Function
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM tipos_analisis where anulado = 0 and tipo_muestra_id <> 0 order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT id_tipo_analisis,nombre FROM tipos_analisis where anulado = 0 order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Listado_CE() As ADODB.Recordset
    Dim consulta As String
    'CONTROL_PROCESOS
    consulta = "SELECT A.ID_TIPO_ANALISIS,A.NOMBRE " & _
               "  FROM TIPOS_ANALISIS A, TIPOS_MUESTRA B " & _
               " WHERE A.ANULADO = 0 " & _
               "   AND A.TIPO_MUESTRA_ID = B.ID_TIPO_MUESTRA " & _
               "   AND B.TIPO_ESPECIAL_ID IN (" & tipo_especial.control_eficacia & "," & tipo_especial.CONTROLES_PROCESOS & ")" & _
               " ORDER BY A.NOMBRE"
    Set Listado_CE = datos_bd(consulta)
End Function
Public Function CARGAR(TIPO_ANALISIS As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM tipos_analisis WHERE id_tipo_analisis=" & TIPO_ANALISIS
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        ID_TIPO_ANALISIS = rs("id_tipo_analisis")
        nombre = rs("NOMBRE")
        NOMBRE_INGLES = rs("NOMBRE_INGLES")
        NOMBRE_SPA = rs("NOMBRE_SPA")
        NORMALIZADO = rs("NORMALIZADO")
        PRECIO = rs("precio")
        TIPO_MUESTRA_ID = rs("TIPO_MUESTRA_ID")
        NORMATIVA = rs("NORMATIVA")
        METODO = rs("METODO")
        anulado = rs("ANULADO")
        PARAMETRO = rs("PARAMETRO")
        TARIFA_CODIGO_ID = rs("TARIFA_CODIGO_ID")
        revisar_factura = rs("REVISAR_FACTURA")
        FACTURA_DETERMINACIONES = rs("FACTURA_DETERMINACIONES")
        TIPO_TRIGO = rs("TIPO_TRIGO")
        DIAS_TRABAJO = rs("DIAS_TRABAJO")
        ADJUNTOS = rs("ADJUNTOS")
        CARGAR = True
    Else
        CARGAR = False
    End If
    Exit Function
fallo:
    CARGAR = False
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_TIPO_ANALISIS) FROM TIPOS_ANALISIS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_TIPO_ANALISIS = 1
    Else
        ID_TIPO_ANALISIS = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearID
    consulta = "Insert into TIPOS_ANALISIS " & _
               " values(" & _
               ID_TIPO_ANALISIS & "," & TIPO_MUESTRA_ID & ",'" & nombre & "','" & NOMBRE_INGLES & "','" & NOMBRE_SPA & "'," & NORMALIZADO & "," & _
               IIf(Trim(PRECIO) = "", "0", Replace(PRECIO, ",", ".")) & ",'" & NORMATIVA & "','" & METODO & "','" & PARAMETRO & "'," & TARIFA_CODIGO_ID & "," & revisar_factura & "," & _
               FACTURA_DETERMINACIONES & "," & TIPO_TRIGO & "," & DIAS_TRABAJO & ",'" & ADJUNTOS & "',0)"
    execute_bd consulta
    Insertar = ID_TIPO_ANALISIS
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar los TIPOS_ANALISIS (TIPOS_ANALISIS)", vbCritical, Err.Description
End Function
Public Function Eliminar(TIPO_ANALISIS As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    execute_bd "UPDATE TIPOS_ANALISIS SET ANULADO = 2 WHERE ID_TIPO_ANALISIS=" & TIPO_ANALISIS
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar el tipo de análisis (TIPOS_ANALISIS)", vbCritical, Err.Description
End Function

Public Sub Imprimir_Listado(ByVal prm_desc As String, ByVal prm_TM As Long, ByVal prm_evitar_dia As Boolean)

    Dim objRpt As New frmReport
    Dim arrNom() As String
    Dim arrVal() As String
    Dim strCriterio As String
    
    
    strCriterio = " {tipos_analisis.ANULADO} = 0 AND {tipos_analisis.TIPO_MUESTRA_ID} <> 0"
    
    If prm_evitar_dia Then strCriterio = strCriterio & " AND {tipos_muestra.ID_TIPO_MUESTRA} <> 56 "
    If Trim(prm_desc) <> "" Then strCriterio = strCriterio & " AND {tipos_analisis.NOMBRE} like " & Chr(34) & "*" & prm_desc & "*" & Chr(34)
    If prm_TM > 0 Then strCriterio = strCriterio & " AND {tipos_muestra.ID_TIPO_MUESTRA} = " & CStr(prm_TM)
                
    ReDim arrNom(1)
    ReDim arrVal(1)

    arrNom(1) = "impresor"
    arrVal(1) = USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
        
    With objRpt
        .iniciar
        .informe = "General\rptListadoTA"
        .criterio = strCriterio
        .imprimir = False
        
        .ParametrosNombre = arrNom
        .ParametrosValores = arrVal

        .generar
        '.Visible = True
        .Show 1
    End With
    Unload objRpt
    Set objRpt = Nothing

End Sub
Public Function Eliminar_BD(TIPO_ANALISIS As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    execute_bd "DELETE FROM TIPOS_ANALISIS WHERE ID_TIPO_ANALISIS=" & TIPO_ANALISIS
    execute_bd "DELETE FROM DETERMINACIONES_ANALISIS WHERE TIPO_ANALISIS_ID=" & TIPO_ANALISIS
    execute_bd "DELETE FROM TARIFAS_PRECIOS WHERE BANO_ID = 0 AND TIPO_DETERMINACION_ID = 0 AND TIPO_ANALISIS_ID=" & TIPO_ANALISIS
    execute_bd "DELETE FROM TIPOS_DATOS_ANALISIS WHERE TIPO_ANALISIS_ID=" & TIPO_ANALISIS
'    execute_bd "DELETE FROM TIPOS_DOCUMENTOS_ANALISIS WHERE TIPO_ANALISIS_ID=" & TIPO_ANALISIS
    Eliminar_BD = True
    Exit Function
fallo:
    Eliminar_BD = False
    MsgBox "Error al eliminar el tipo de análisis (TIPOS_ANALISIS)", vbCritical, Err.Description
End Function

Public Function Modificar(ID As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearID
    consulta = " UPDATE TIPOS_ANALISIS " & _
               "  SET TIPO_MUESTRA_ID= " & TIPO_MUESTRA_ID & "," & _
               "               NOMBRE='" & nombre & "'," & _
               "        NOMBRE_INGLES='" & NOMBRE_INGLES & "'," & _
               "           NOMBRE_SPA='" & NOMBRE_SPA & "'," & _
               "         NORMALIZADO = " & NORMALIZADO & "," & _
               "         NORMATIVA   ='" & NORMATIVA & "'," & _
               "         METODO      ='" & METODO & "'," & _
               " TARIFA_CODIGO_ID    = " & TARIFA_CODIGO_ID & "," & _
               " REVISAR_FACTURA     = " & revisar_factura & "," & _
               " FACTURA_DETERMINACIONES = " & FACTURA_DETERMINACIONES & "," & _
               "          TIPO_TRIGO = " & TIPO_TRIGO & "," & _
               "        DIAS_TRABAJO = " & DIAS_TRABAJO & "," & _
               "            ADJUNTOS ='" & ADJUNTOS & "'," & _
               "         PARAMETRO   ='" & PARAMETRO & "'," & _
               "              PRECIO = " & IIf(Trim(PRECIO) = "", "0", Replace(PRECIO, ",", ".")) & "" & _
               " WHERE ID_TIPO_ANALISIS=" & ID
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al modificar el TIPOS_ANALISIS (TIPOS_ANALISIS)", vbCritical, Err.Description
End Function

Public Function Listado_Muestra(MUESTRA As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM tipos_analisis where tipo_muestra_id = " & MUESTRA & " order by nombre"
    Set Listado_Muestra = datos_bd(consulta)
End Function
Public Function lista(filtro As String, dia As Boolean) As ADODB.Recordset
    Dim consulta As String
    Dim sdia As String
    If dia = True Then
        sdia = " and tm.id_tipo_muestra <> 56 "
    End If
    consulta = "SELECT ta.nombre,tm.nombre,ta.normalizado,ta.precio,ta.id_tipo_analisis " & _
               "  FROM tipos_analisis ta, tipos_muestra tm " & _
               " where ta.anulado = 0 " & _
               "   and ta.tipo_muestra_id = tm.id_tipo_muestra " & _
               "   and ta.tipo_muestra_id <> 0 " & _
               "   and ta.nombre like '%" & filtro & "%' " & _
               sdia & _
               " order by ta.nombre"
    Set lista = datos_bd(consulta)
End Function
Public Function lista_tipo_muestra(tm As Integer, filtro As String, dia As Boolean) As ADODB.Recordset
    Dim consulta As String
    Dim sdia As String
    If dia = True Then
        sdia = " and tm.id_tipo_muestra <> 56 "
    End If
    consulta = "SELECT ta.nombre,tm.nombre,ta.normalizado,ta.precio,ta.id_tipo_analisis " & _
               "  FROM tipos_analisis ta, tipos_muestra tm " & _
               " where ta.anulado = 0 " & _
               "   and ta.tipo_muestra_id = tm.id_tipo_muestra " & _
               "   and ta.tipo_muestra_id <> 0 " & _
               "   and tm.id_tipo_muestra = " & tm & _
               "   and ta.nombre like '%" & filtro & "%' " & _
               sdia & _
               " order by ta.nombre"
    Set lista_tipo_muestra = datos_bd(consulta)
End Function

Public Function Listado_Tarifa(TARIFA As Long, cliente As Long, CODIGO_TARIFA As Long) As ADODB.Recordset
    Dim consulta As String
    Dim s1 As String
    If CODIGO_TARIFA <> 0 Then
        s1 = " AND a.TARIFA_CODIGO_ID = " & CODIGO_TARIFA
    End If
    If cliente = 0 Then
        consulta = "SELECT distinct a.id_tipo_analisis,a.nombre,c.codigo,d.precio,b.precio,a.tarifa_codigo_id " & _
                   "  FROM tipos_analisis a " & _
                   "  LEFT JOIN tarifas_precios b on a.id_tipo_analisis = b.tipo_analisis_id " & _
                   "  LEFT JOIN tarifas_precios d on a.id_tipo_analisis = d.tipo_analisis_id " & _
                   "  LEFT JOIN tarifas_codigos c on a.tarifa_codigo_id = c.id_codigo " & _
                   " INNER JOIN tarifas       tar on d.tarifa_id = tar.tarifa_origen_id  " & _
                   " WHERE a.anulado = 0 and b.tarifa_id = " & TARIFA & _
                   "   AND a.tipo_muestra_id <> 0 " & _
                   "   AND tar.id_tarifa = " & TARIFA & _
                   s1 & _
                   " ORDER BY a.nombre"
    Else
        consulta = "SELECT distinct a.id_tipo_analisis,a.nombre,c.codigo,d.precio,b.precio,a.tarifa_codigo_id " & _
                   "  FROM tipos_analisis a " & _
                   "  LEFT JOIN tarifas_precios b on a.id_tipo_analisis = b.tipo_analisis_id " & _
                   "  LEFT JOIN tarifas_precios d on a.id_tipo_analisis = d.tipo_analisis_id " & _
                   "  LEFT JOIN tarifas_codigos c on a.tarifa_codigo_id = c.id_codigo " & _
                   " INNER JOIN tarifas       tar on d.tarifa_id = tar.tarifa_origen_id  " & _
                   " INNER JOIN muestras        m on m.tipo_analisis_id = a.id_tipo_analisis  " & _
                   " WHERE a.anulado = 0 and b.tarifa_id = " & TARIFA & _
                   "   AND m.cliente_id = " & cliente & _
                   "   AND a.tipo_muestra_id <> 0 " & _
                   "   AND tar.id_tarifa = " & TARIFA & _
                   s1 & _
                   " ORDER BY a.nombre"
    End If
    Set Listado_Tarifa = datos_bd(consulta)
End Function
Public Function lista_por_determinacion(DETERMINACION As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ta.nombre,tm.nombre,ta.normalizado,ta.precio,ta.id_tipo_analisis " & _
               "  FROM tipos_analisis ta, tipos_muestra tm, determinaciones_analisis da " & _
               " where ta.anulado = 0 " & _
               "   and ta.tipo_muestra_id = tm.id_tipo_muestra " & _
               "   and ta.tipo_muestra_id <> 0 " & _
               "   and ta.id_tipo_analisis = da.tipo_analisis_id " & _
               "   and da.tipo_determinacion_id = " & DETERMINACION & _
               " order by ta.nombre"
    Set lista_por_determinacion = datos_bd(consulta)
End Function
Public Function Modificar_Codigo_Tarifa(ID As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = " UPDATE TIPOS_ANALISIS " & _
               "  SET TARIFA_CODIGO_ID= " & TARIFA_CODIGO_ID & _
               " WHERE ID_TIPO_ANALISIS=" & ID
    execute_bd consulta
    Modificar_Codigo_Tarifa = True
    Exit Function
fallo:
    Modificar_Codigo_Tarifa = False
    MsgBox "Error al Modificar_Codigo_Tarifa el TIPOS_ANALISIS (TIPOS_ANALISIS)", vbCritical, Err.Description
End Function

Public Function revisar_facturacion(TIPO_ANALISIS As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT revisar_factura FROM tipos_analisis WHERE id_tipo_analisis=" & TIPO_ANALISIS
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        revisar_facturacion = rs("REVISAR_FACTURA")
    Else
        revisar_facturacion = False
    End If
    Exit Function
fallo:
    revisar_facturacion = False
End Function

Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = "TIPO_MUESTRA_ID"
        .setFK_VALOR = FK
        .setTABLA = "TIPOS_ANALISIS"
        .setDESCRIPCION = "Tipos Análisis"
        .setPK = "ID_TIPO_ANALISIS"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = True
        .setFILTRO = filtro
        Set .FORMULARIO = FORMULARIO
    End With
End Function
