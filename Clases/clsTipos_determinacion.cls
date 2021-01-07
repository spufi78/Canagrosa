VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_determinacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*****************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE tipos_determinacion      *
'*****************************************************************
Private ID_TIPO_DETERMINACION As Integer
Private nombre As String
Private NOMBRE_INGLES As String
Private PNT As String
Private PROC_REF_EADS As String
Private DESCRIPCION As String
Private FORMULA_ID As Integer
Private FAMILIA_ID As Integer
Private PRECIO As String
Private TARIFA_CODIGO_ID As Integer
Private revisar_factura As Integer
Private PARTICULAS As Integer
Private anulado As Integer
Private ES_SUBCONTRATABLE As Integer
Private PNT_VINCULADO As Integer
Private METODO As String
Private NORMA_ID As Integer
'E0115-F
Public Property Let setID_TIPO_DETERMINACION(ByVal dato As Integer)
    ID_TIPO_DETERMINACION = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Let setNOMBRE_INGLES(ByVal dato As String)
    NOMBRE_INGLES = Trim(dato)
End Property
Public Property Let setPNT(ByVal dato As String)
    PNT = Trim(dato)
End Property
Public Property Let setPROC_REF_EADS(ByVal dato As String)
    PROC_REF_EADS = Trim(dato)
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
    DESCRIPCION = Trim(dato)
End Property
Public Property Let setFORMULA_ID(ByVal dato As Integer)
    FORMULA_ID = dato
End Property
Public Property Let setFAMILIA_ID(ByVal dato As Integer)
    FAMILIA_ID = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
    PRECIO = Trim(dato)
End Property
Public Property Let setTARIFA_CODIGO_ID(ByVal dato As Integer)
    TARIFA_CODIGO_ID = Trim(dato)
End Property
Public Property Let setREVISAR_FACTURA(ByVal dato As Integer)
    revisar_factura = Trim(dato)
End Property
Public Property Let setPARTICULAS(ByVal dato As Integer)
    PARTICULAS = dato
End Property
Public Property Let setANULADO(ByVal dato As Integer)
    anulado = dato
End Property
'E0116-I
Public Property Let setES_SUBCONTRATABLE(ByVal dato As Integer)
        ES_SUBCONTRATABLE = dato
End Property
Public Property Let setMETODO(ByVal dato As String)
    METODO = dato
End Property
Public Property Let setPNT_VINCULADO(ByVal dato As Integer)
    PNT_VINCULADO = dato
End Property
Public Property Let setNORMA_ID(ByVal dato As Integer)
    NORMA_ID = dato
End Property
'E0116-F
''''''''''''''''''''''''
'   PROPIEDADES GET
''''''''''''''''''''''''
Public Property Get getID_TIPO_DETERMINACION() As Integer
    getID_TIPO_DETERMINACION = ID_TIPO_DETERMINACION
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getNOMBRE_INGLES() As String
    getNOMBRE_INGLES = NOMBRE_INGLES
End Property
Public Property Get getPNT() As String
    getPNT = PNT
End Property
Public Property Get getPROC_REF_EADS() As String
    getPROC_REF_EADS = PROC_REF_EADS
End Property
Public Property Get getDESCRIPCION() As String
    getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getFORMULA_ID() As Integer
    getFORMULA_ID = FORMULA_ID
End Property
Public Property Get getFAMILIA_ID() As Integer
    getFAMILIA_ID = FAMILIA_ID
End Property
Public Property Get getPRECIO() As String
    getPRECIO = PRECIO
End Property
Public Property Get getTARIFA_CODIGO_ID() As Integer
    getTARIFA_CODIGO_ID = TARIFA_CODIGO_ID
End Property
Public Property Get getREVISAR_FACTURA() As Integer
    getREVISAR_FACTURA = revisar_factura
End Property
Public Property Get getPARTICULAS() As Integer
    getPARTICULAS = PARTICULAS
End Property
Public Property Get getANULADO() As Integer
    getANULADO = anulado
End Property
'E0117-I
Public Property Get getES_SUBCONTRATABLE() As Integer
        getES_SUBCONTRATABLE = ES_SUBCONTRATABLE
End Property
Public Property Get getMETODO() As String
    getMETODO = METODO
End Property
Public Property Get getPNT_VINCULADO() As Integer
    getPNT_VINCULADO = PNT_VINCULADO
End Property
Public Property Get getNORMA_ID() As Integer
    getNORMA_ID = NORMA_ID
End Property

Public Function DeterminacionesPorDefecto(id_analisis As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT de.id_tipo_determinacion as tipo, " & _
               "       de.pnt, de.nombre, de.descripcion, an.minimo, an.maximo, an.requerida " & _
               "  FROM tipos_determinacion de, determinaciones_analisis an " & _
               " WHERE tipo_analisis_id =" & id_analisis & _
               " AND id_tipo_determinacion = tipo_determinacion_id" & _
               " AND de.anulado = 0" & _
               " ORDER BY orden"
    Set DeterminacionesPorDefecto = datos_bd(consulta)
End Function
Public Function DeterminacionesPorDefectoBano(BANO As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT de.id_tipo_determinacion as tipo, " & _
               "       de.pnt, de.nombre, de.descripcion, an.minimo, an.maximo, an.requerida " & _
               "  FROM tipos_determinacion de, determinaciones_analisis an " & _
               " WHERE bano_id =" & BANO & _
               " AND id_tipo_determinacion = tipo_determinacion_id" & _
               " AND de.anulado = 0" & _
               " ORDER BY orden"
    Set DeterminacionesPorDefectoBano = datos_bd(consulta)
End Function
Public Function CargarTipoDeterminacion(ByVal ID As Integer) As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "select * from tipos_determinacion where id_tipo_determinacion = " & ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        CargarTipoDeterminacion = False
    Else
        CargarTipoDeterminacion = True
        ID_TIPO_DETERMINACION = rs("ID_TIPO_DETERMINACION")
        nombre = rs("NOMBRE")
        NOMBRE_INGLES = rs("NOMBRE_INGLES")
        PNT = rs("PNT")
        PROC_REF_EADS = rs("PROC_REF_EADS")
        DESCRIPCION = rs("DESCRIPCION")
        FORMULA_ID = rs("FORMULA_ID")
        FAMILIA_ID = rs("FAMILIA_ID")
        PRECIO = rs("PRECIO")
        TARIFA_CODIGO_ID = rs("TARIFA_CODIGO_ID")
        revisar_factura = rs("REVISAR_FACTURA")
        PARTICULAS = rs("PARTICULAS")
        anulado = rs("ANULADO")
        ES_SUBCONTRATABLE = rs("ES_SUBCONTRATABLE")
        METODO = rs("METODO")
        PNT_VINCULADO = rs("PNT_VINCULADO")
        NORMA_ID = entero(rs("NORMA_ID"))
    End If
    Set rs = Nothing
    Exit Function
fallo:
    CargarTipoDeterminacion = False
    MsgBox "Error al cargar el tipo determinacion (CargarTipoDeterminacion)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT id_tipo_determinacion, concat(nombre,' ',descripcion) as tipo FROM tipos_determinacion where anulado = 0 order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function DeterminacionesBano(id_bano As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT de.id_tipo_determinacion as tipo, " & _
               " de.pnt, de.nombre, de.descripcion,rb.requerida " & _
               " FROM tipos_determinacion de, DETERMINACIONES_ANALISIS rb" & _
               " WHERE bano_id =" & id_bano & _
               " AND id_tipo_determinacion = tipo_determinacion_id" & _
               " AND de.anulado = 0" & _
               " order by rb.orden"
    Set DeterminacionesBano = datos_bd(consulta)
End Function
Public Function Listado_completo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM tipos_determinacion order by nombre"
    Set Listado_completo = datos_bd(consulta)
End Function
Public Function Listado_por_Formula(Formula As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM tipos_determinacion WHERE FORMULA_ID = " & Formula & " order by nombre"
    Set Listado_por_Formula = datos_bd(consulta)
End Function
Public Function lista(filtro As String, PNT As String, NORMA As String, strSubcontratable As String, familia As String, anuladas As Integer) As ADODB.Recordset
    Dim consulta As String
    Dim s As String
    If familia <> "0" Then
        s = s & " AND td.familia_id = " & CLng(familia)
    End If
    If anuladas = 1 Then
        s = s & " AND td.ANULADO > 0"
    Else
        s = s & " AND td.ANULADO = 0"
    End If
    If PNT <> "" Then
        s = s & " AND td.PNT like '%" & PNT & "%'"
    End If
    If NORMA <> "" Then
        s = s & " AND td.PROC_REF_EADS like '%" & NORMA & "%'"
    End If
    consulta = "SELECT td.nombre,f.nombre as FORMULA,td.pnt,td.id_tipo_determinacion,td.precio,tdf.nombre as FAMILIA,td.familia_id, td.ES_SUBCONTRATABLE, td.descripcion " & _
               "  FROM tipos_determinacion td, formulas f,tipos_determinacion_familias tdf " & _
               " WHERE td.formula_id = f.id_formula " & _
               "   and td.familia_id = tdf.id_familia " & _
               "   AND concat(td.NOMBRE,' ',td.descripcion) LIKE '%" & Trim(filtro) & "%' " & _
               IIf(strSubcontratable = "1", " AND td.ES_SUBCONTRATABLE = 1 ", "") & _
               s & _
               " order by td.nombre"
    Set lista = datos_bd(consulta)
End Function

Public Sub Imprimir_Listado(ByVal prm_desc As String, ByVal prm_Familia As Long, ByVal prm_dia As Boolean, ByVal prm_no_usadas As Boolean, ByVal prm_subcontratables As Boolean)

    Dim objRpt As New frmReport
    Dim arrNom() As String
    Dim arrVal() As String
    Dim strCriterio As String
        
    strCriterio = " {tipos_determinacion.ANULADO} = 0"
        
    If prm_Familia > 0 Then strCriterio = strCriterio & " AND {tipos_determinacion.FAMILIA_ID} = " & CStr(prm_Familia)
    If Trim(prm_desc) <> "" Then strCriterio = strCriterio & " AND {tipos_determinacion.NOMBRE} like " & Chr(34) & "*" & prm_desc & "*" & Chr(34)
    If prm_subcontratables Then strCriterio = strCriterio & " AND {tipos_determinacion.ES_SUBCONTRATABLE} = 1"
                
    ReDim arrNom(1)
    ReDim arrVal(1)

    arrNom(1) = "impresor"
    arrVal(1) = USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
        
    With objRpt
        .iniciar
        .informe = "General\rptListadoTD" & IIf(prm_no_usadas, "_nousados", "")
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

Public Function lista_no_utilizadas(filtro As String, PNT As String, NORMA As String, strSubcontratable, familia As String, anuladas As Integer) As ADODB.Recordset
    Dim consulta As String
    Dim s As String
    If familia <> "" Then
        s = s & " AND td.familia_id = " & CLng(familia)
    End If
    If anuladas = 1 Then
        s = s & " AND td.ANULADO = 1"
    End If
    If PNT <> "" Then
        s = s & " AND td.PNT like '%" & PNT & "%'"
    End If
    If NORMA <> "" Then
        s = s & " AND td.PROC_REF_EADS like '%" & NORMA & "%'"
    End If
    
    consulta = "SELECT td.nombre,f.nombre as FORMULA,td.pnt,td.id_tipo_determinacion,td.precio,tdf.nombre as FAMILIA,td.familia_id, td.ES_SUBCONTRATABLE, td.descripcion  " & _
               "  FROM tipos_determinacion td " & _
               "  LEFT JOIN determinaciones_analisis da on td.id_tipo_determinacion = da.tipo_determinacion_id " & _
               "  LEFT JOIN formulas f on td.formula_id = f.id_formula " & _
               "  LEFT JOIN tipos_determinacion_familias tdf on td.familia_id = tdf.id_familia " & _
               " WHERE (da.tipo_determinacion_id is null or f.id_formula is null) " & _
               "   and td.anulado = 0 " & _
               "   AND concat(td.NOMBRE,' ',td.descripcion) LIKE '%" & Trim(filtro) & "%' " & _
               IIf(strSubcontratable = "1", " AND td.ES_SUBCONTRATABLE = 1 ", "") & _
               s & _
               " order by td.nombre"
    Set lista_no_utilizadas = datos_bd(consulta)
End Function

Public Function Modificar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    'E0119-I
    consulta = " UPDATE TIPOS_DETERMINACION " & _
               "  SET NOMBRE= '" & nombre & "'," & _
               "   NOMBRE_INGLES ='" & NOMBRE_INGLES & "'," & _
               "          PNT='" & PNT & "'," & _
               " PROC_REF_EADS ='" & PROC_REF_EADS & "'," & _
               "   DESCRIPCION ='" & DESCRIPCION & "'," & _
               "   PRECIO ='" & PRECIO & "'," & _
               "   TARIFA_CODIGO_ID =" & TARIFA_CODIGO_ID & "," & _
               "   REVISAR_FACTURA = " & revisar_factura & "," & _
               "   FAMILIA_ID = " & FAMILIA_ID & "," & "   PARTICULAS = " & PARTICULAS & "," & _
               "   FORMULA_ID = " & FORMULA_ID & ","
    consulta = consulta & _
               "   ES_SUBCONTRATABLE = " & ES_SUBCONTRATABLE & "," & _
               "   METODO = '" & METODO & "'," & _
               "   PNT_VINCULADO = " & PNT_VINCULADO & _
               "  ,NORMA_ID = " & IIf(NORMA_ID = 0, "null", NORMA_ID) & _
               " WHERE ID_TIPO_DETERMINACION=" & ID
    'E0119-F
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al modificar la determinacion (Mod_Determinacion)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearID
    consulta = "Insert into TIPOS_DETERMINACION " & _
               " values(" & _
               ID_TIPO_DETERMINACION & ",'" & nombre & "','" & NOMBRE_INGLES & "'," & FORMULA_ID & "," & FAMILIA_ID & ",'" & PNT & "','" & _
               PROC_REF_EADS & "','" & DESCRIPCION & "','" & _
               PRECIO & "'," & TARIFA_CODIGO_ID & "," & revisar_factura & "," & _
               PARTICULAS & ",0, " & _
               ES_SUBCONTRATABLE & ",'" & METODO & "'," & PNT_VINCULADO & "," & IIf(NORMA_ID = 0, "null", NORMA_ID) & ")"
    execute_bd consulta
    Insertar = ID_TIPO_DETERMINACION
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar los TIPOS_DETERMINACION", vbCritical, Err.Description
End Function

Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_TIPO_DETERMINACION) FROM TIPOS_DETERMINACION"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_TIPO_DETERMINACION = 1
    Else
        ID_TIPO_DETERMINACION = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub

Public Function Eliminar(TIPO_DETERMINACION As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "UPDATE TIPOS_DETERMINACION SET ANULADO = " & USUARIO.getID_EMPLEADO & " WHERE ID_TIPO_DETERMINACION=" & TIPO_DETERMINACION
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar los TIPOS_DETERMINACION", vbCritical, Err.Description
End Function
Public Function DeterminacionesPorMuestra(TIPO_MUESTRA As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT distinct de.id_tipo_determinacion, concat(de.nombre,' ',de.descripcion) as tipo " & _
               "FROM tipos_determinacion de, determinaciones_analisis an, tipos_analisis ta " & _
               " WHERE ta.tipo_muestra_id = " & TIPO_MUESTRA & _
               "   AND ta.id_tipo_analisis = an.tipo_analisis_id" & _
               "   AND de.id_tipo_determinacion = an.tipo_determinacion_id" & _
               "   AND de.ANULADO = 0 " & _
               " ORDER BY de.nombre"
    Set DeterminacionesPorMuestra = datos_bd(consulta)
End Function
Public Function DeterminacionesBanoNombre() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT DISTINCT id_tipo_determinacion, concat(nombre,' ',descripcion) as tipo " & _
               " FROM tipos_determinacion de, determinaciones_analisis rb" & _
               " WHERE id_tipo_determinacion = tipo_determinacion_id " & _
               "   AND de.ANULADO = 0 " & _
               " order by de.nombre"
    Set DeterminacionesBanoNombre = datos_bd(consulta)
End Function
Public Function Determinaciones_por_bano(BANO As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT DISTINCT id_tipo_determinacion, concat(nombre,' ',descripcion) as tipo " & _
               " FROM tipos_determinacion de, determinaciones_analisis rb" & _
               " WHERE id_tipo_determinacion = tipo_determinacion_id " & _
               "   AND rb.BANO_ID = " & BANO & _
               "   AND de.ANULADO = 0 " & _
               " order by de.nombre"
    Set Determinaciones_por_bano = datos_bd(consulta)
End Function
Public Function DeterminacionesTodas() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT distinct de.id_tipo_determinacion, concat(de.nombre,' ',de.descripcion, ' (Id:',de.id_tipo_determinacion,')') as tipo " & _
               "FROM tipos_determinacion de " & _
               "WHERE ANULADO = 0 " & _
               " ORDER BY de.nombre"
    Set DeterminacionesTodas = datos_bd(consulta)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setQUERY = ""
        .setTABLA = "TIPOS_DETERMINACION"
        .setDESCRIPCION = "Tipos de determinaciones"
        .setPK = "ID_TIPO_DETERMINACION"
        .setCAMPO = "concat(NOMBRE,' ',DESCRIPCION)"
'        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = True
        .setFILTRO = "ANULADO=0"
        Set .FORMULARIO = FORMULARIO
    End With
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT id_tipo_determinacion, concat(nombre,' ',descripcion) as tipo FROM tipos_determinacion where anulado = 0 order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function

Public Function Listado_Tarifa(TARIFA As Long, cliente As Long, solucion As Long, CODIGO As Long) As ADODB.Recordset
    Dim consulta As String
    Dim s1 As String
    Dim s2 As String
    If CODIGO <> 0 Then
        s1 = " AND a.TARIFA_CODIGO_ID = " & CODIGO
    End If
    If solucion <> 0 Then
        s2 = "   AND bano.solucion_id = " & solucion
    End If
    If cliente = 0 Then
        consulta = "SELECT DISTINCT a.id_tipo_determinacion,concat(a.nombre,' (',a.descripcion,')') as tipo,c.codigo,d.precio,b.precio,a.tarifa_codigo_id " & _
                   "  FROM tipos_determinacion a " & _
                   "  LEFT JOIN tarifas_precios b on a.id_tipo_determinacion = b.tipo_determinacion_id " & _
                   "  LEFT JOIN tarifas_precios d on a.id_tipo_determinacion = d.tipo_determinacion_id and d.bano_id = 0 and d.tipo_analisis_id = 0 " & _
                   "  LEFT JOIN tarifas_codigos c on a.tarifa_codigo_id = c.id_codigo " & _
                   " INNER JOIN tarifas       tar on d.tarifa_id = tar.tarifa_origen_id  " & _
                   " INNER JOIN banos        bano on d.tarifa_id = tar.tarifa_origen_id  " & _
                   " WHERE a.anulado = 0 and b.tarifa_id = " & TARIFA & _
                   "   AND b.tarifa_id = tar.id_tarifa " & _
                   s1 & s2 & _
                   " ORDER BY a.nombre"
    Else
        consulta = "SELECT DISTINCT a.id_tipo_determinacion,concat(a.nombre,' (',a.descripcion,')') as tipo,c.codigo,d.precio,b.precio,a.tarifa_codigo_id " & _
                   "  FROM tipos_determinacion a " & _
                   "  LEFT JOIN tarifas_precios b on a.id_tipo_determinacion = b.tipo_determinacion_id " & _
                   "  LEFT JOIN tarifas_precios d on a.id_tipo_determinacion = d.tipo_determinacion_id and d.bano_id = 0 and d.tipo_analisis_id = 0 " & _
                   "  LEFT JOIN tarifas_codigos c on a.tarifa_codigo_id = c.id_codigo " & _
                   " INNER JOIN tarifas       tar on d.tarifa_id = tar.tarifa_origen_id  " & _
                   " INNER JOIN banos        bano on d.tarifa_id = tar.tarifa_origen_id  " & _
                   " INNER JOIN determinaciones_analisis da ON a.id_tipo_determinacion = da.tipo_determinacion_id AND da.bano_id = bano.id_bano " & _
                   " WHERE a.anulado = 0 and b.tarifa_id = " & TARIFA & _
                   "   AND b.tarifa_id = tar.id_tarifa " & _
                   " and bano.cliente_id = " & cliente & _
                   s1 & s2 & _
                   " ORDER BY a.nombre"
    End If
    Set Listado_Tarifa = datos_bd(consulta)
End Function

Public Function Modificar_Familia(TD As Integer, familia As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = " UPDATE TIPOS_DETERMINACION " & _
               "  SET " & _
               "    FAMILIA_ID = " & familia & _
               " WHERE ID_TIPO_DETERMINACION=" & TD
    execute_bd consulta
    Modificar_Familia = True
    Exit Function
fallo:
    Modificar_Familia = False
    MsgBox "Error al modificar la determinacion (Modificar_Familia)", vbCritical, Err.Description
End Function

'Public Function Modificar_Tarifa(TD As Integer, TARIFA As String) As Boolean
'    On Error GoTo fallo
'    Dim consulta As String
'    consulta = " UPDATE TIPOS_DETERMINACION " & _
'               "  SET " & _
'               "    TARIFA_CODIGO_ID = '" & TARIFA & "' " & _
'               " WHERE ID_TIPO_DETERMINACION=" & TD
'    execute_bd consulta
'    Modificar_Tarifa = True
'    Exit Function
'fallo:
'    Modificar_Tarifa = False
'    MsgBox "Error al modificar la tarifa (Modificar_Familia)", vbCritical, Err.Description
'End Function
Public Function Modificar_Precio(TD As Integer, lPrecio As String) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = " UPDATE TIPOS_DETERMINACION " & _
               "  SET " & _
               "    PRECIO = '" & lPrecio & "' " & _
               " WHERE ID_TIPO_DETERMINACION=" & TD
    execute_bd consulta
    Modificar_Precio = True
    Exit Function
fallo:
    Modificar_Precio = False
    MsgBox "Error al modificar el precio (Modificar_Familia)", vbCritical, Err.Description
End Function
Public Function Modificar_Codigo_Tarifa(ID As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = " UPDATE TIPOS_DETERMINACION " & _
               "  SET TARIFA_CODIGO_ID= " & TARIFA_CODIGO_ID & _
               " WHERE ID_TIPO_DETERMINACION=" & ID
    execute_bd consulta
    Modificar_Codigo_Tarifa = True
    Exit Function
fallo:
    Modificar_Codigo_Tarifa = False
    MsgBox "Error al Modificar_Codigo_Tarifa el TIPOS_ANALISIS (ID_TIPO_DETERMINACION)", vbCritical, Err.Description
End Function
Public Function Modificar_FORMULA(FORMULA_ANTIGUA As Integer, FORMULA_NUEVA As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = " UPDATE TIPOS_DETERMINACION " & _
               "    SET FORMULA_ID = " & FORMULA_NUEVA & _
               "  WHERE FORMULA_ID =" & FORMULA_ANTIGUA
    execute_bd consulta
    Modificar_FORMULA = True
    Exit Function
fallo:
    Modificar_FORMULA = False
    MsgBox "Error al modificar las formulas (clsTD : Modificar_FORMULA)", vbCritical, Err.Description
End Function
Public Function duplicar(ID_TIPO_DETERMINACION As Long) As Long
    Dim oDeter As New clsTipos_determinacion
   On Error GoTo duplicar_Error
    Dim DETERMINACION As Long
    If oDeter.CargarTipoDeterminacion(ID_TIPO_DETERMINACION) = True Then
        With oDeter
          .setNOMBRE = oDeter.getNOMBRE & " (Duplicado)"
          .setPRECIO = moneda_bd(oDeter.getPRECIO)
          DETERMINACION = .Insertar
          If DETERMINACION = 0 Then
              MsgBox "Error al insertar el tipo de determinación duplicado.", vbCritical, App.Title
              duplicar = DETERMINACION
              Exit Function
          End If
        End With
        Dim c As String
        ' Botes_ex
        c = "INSERT INTO tipos_determinacion_botes_ex "
        c = c & " SELECT " & DETERMINACION & ",BOTE_EX_ID,TIPO,ORDEN FROM tipos_determinacion_botes_ex "
        c = c & " WHERE TIPO_DETERMINACION_ID = " & ID_TIPO_DETERMINACION
        execute_bd c
        ' Equipos
        c = "INSERT INTO tipos_determinacion_equipos "
        c = c & " SELECT " & DETERMINACION & ",EQUIPO_ID,ORDEN FROM tipos_determinacion_equipos"
        c = c & " WHERE TIPO_DETERMINACION_ID = " & ID_TIPO_DETERMINACION
        execute_bd c
        ' Tipo_dependientes
        c = "INSERT INTO tipos_determinacion_dep "
        c = c & " SELECT " & DETERMINACION & ",CAMPO_ID,TIPO_DETERMINACION_ID_DEP FROM tipos_determinacion_dep"
        c = c & " WHERE TIPO_DETERMINACION_ID = " & ID_TIPO_DETERMINACION
        execute_bd c
        ' Contratas
        Dim oTDC As New clsTipos_determinacion_contratas
        Dim rsC As New ADODB.Recordset
        Set rsC = oTDC.Listado_Tipo(ID_TIPO_DETERMINACION)
        If rsC.RecordCount > 0 Then
            Do
                With oTDC
                    .setTIPO_DETERMINACION_ID = DETERMINACION
                    .setCONTRATA_ID = rsC("CONTRATA_ID")
                    .setVALOR_REFERENCIA = rsC("VALOR_REFERENCIA")
                    .setPRECIO = moneda_bd(rsC("PRECIO"))
                    .setNORMATIVA_APLICABLE = rsC("NORMATIVA_APLICABLE")
                    .Insertar
                End With
                rsC.MoveNext
            Loop Until rsC.EOF
        End If
        ' Tarifas
        Dim oTP As New clsTarifas_precios
        Dim rs As ADODB.Recordset
        Set rs = oTP.Listado_por_determinacion(ID_TIPO_DETERMINACION)
        If rs.RecordCount > 0 Then
          Do
              With oTP
                  .setTIPO_DETERMINACION_ID = DETERMINACION
                  .setBANO_ID = 0
                  .setTIPO_ANALISIS_ID = 0
                  .setPRECIO = moneda_bd(rs("PRECIO"))
                  .setTARIFA_ID = rs("TARIFA_ID")
                  .Insertar
              End With
              rs.MoveNext
          Loop Until rs.EOF
        End If
    End If
    duplicar = DETERMINACION
   On Error GoTo 0
   Exit Function

duplicar_Error:
    duplicar = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure duplicar of Módulo de clase clsTipos_determinacion"
End Function
