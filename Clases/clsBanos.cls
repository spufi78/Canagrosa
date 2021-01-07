VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBanos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*****************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE MUESTRA      *
'*****************************************************
Private id_bano As Long
Private nombre As String
Private CLIENTE_ID As Integer
Private SOLUCION_ID As Integer
Private PROCESO_BASE_ID As Integer
Private TIPO_FRECUENCIA_ID As Integer
Private LINEA_ID As Integer
Private INSTALACION_ID As Integer
Private FORMATO_ID As Integer
Private CONSERVACION As String
Private TOMA_FIN As String
Private VOLUMEN As String
Private OBSERVACIONES As String
Private SOLUCION_PROCEDENCIA_ID As Integer
Private TIPO_MUESTRA_ID As Integer
Private TARIFA_CODIGO_ID As Integer
Private FACTURA_DETERMINACIONES As Integer
Private revisar_factura As Integer
Private FICHA_ID As Long
Private CENTRO_ID As Long
Private AIRBUS_AREA_ID As Long
Private AIRBUS_LINEA_ID As Long
Private PRECIO As String
'BUG-834-I
Private anulado As Integer
'BUG-834-F

'*****************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE MUESTRA    *
'*****************************************************
Public Property Let setID_BANO(ByVal dato As Long)
    id_bano = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Let setID_LINEA(ByVal dato As Integer)
    LINEA_ID = dato
End Property
Public Property Let setINSTALACION_ID(ByVal dato As Integer)
    INSTALACION_ID = dato
End Property

Public Property Let setID_SOLUCION(ByVal dato As Integer)
    SOLUCION_ID = dato
End Property
Public Property Let setFORMATO_ID(ByVal dato As Integer)
    FORMATO_ID = dato
End Property
Public Property Let setID_PROCESO_BASE(ByVal dato As Integer)
    PROCESO_BASE_ID = dato
End Property
Public Property Let setTIPO_FRECUENCIA_ID(ByVal dato As Integer)
    TIPO_FRECUENCIA_ID = dato
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Integer)
    CLIENTE_ID = dato
End Property
Public Property Let setTIPO_MUESTRA_ID(ByVal dato As Integer)
    TIPO_MUESTRA_ID = dato
End Property
Public Property Let setCONSERVACION(ByVal dato As String)
    CONSERVACION = Trim(dato)
End Property
Public Property Let setTOMA_FIN(ByVal dato As String)
    TOMA_FIN = Trim(dato)
End Property
Public Property Let setVOLUMEN(ByVal dato As String)
    VOLUMEN = Trim(dato)
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
    OBSERVACIONES = Trim(dato)
End Property
Public Property Let setSOLUCION_PROCEDENCIA_ID(ByVal dato As Integer)
    SOLUCION_PROCEDENCIA_ID = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
    PRECIO = Trim(dato)
End Property
Public Property Let setTARIFA_CODIGO_ID(ByVal dato As Integer)
    TARIFA_CODIGO_ID = dato
End Property
Public Property Let setFACTURA_DETERMINACIONES(ByVal dato As Integer)
    FACTURA_DETERMINACIONES = dato
End Property
Public Property Let setREVISAR_FACTURA(ByVal dato As Integer)
    revisar_factura = dato
End Property
Public Property Let setFICHA_ID(ByVal dato As Long)
    FICHA_ID = dato
End Property
Public Property Let setCENTRO_ID(ByVal dato As Long)
    CENTRO_ID = dato
End Property
Public Property Let setAIRBUS_AREA_ID(ByVal dato As Long)
    AIRBUS_AREA_ID = dato
End Property
Public Property Let setAIRBUS_LINEA_ID(ByVal dato As Long)
    AIRBUS_LINEA_ID = dato
End Property
'BUG-834-I
Public Property Let setANULADO(ByVal dato As Integer)
    anulado = dato
End Property
'BUG-834-F
'*****************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE MUESTRA      *
'*****************************************************
Public Property Get getID_BANO() As Long
    getID_BANO = id_bano
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getID_LINEA() As Integer
    getID_LINEA = LINEA_ID
End Property
Public Property Get getINSTALACION_ID() As Integer
    getINSTALACION_ID = INSTALACION_ID
End Property
Public Property Get getFORMATO_ID() As Integer
    getFORMATO_ID = FORMATO_ID
End Property
Public Property Get getID_SOLUCION() As Integer
    getID_SOLUCION = SOLUCION_ID
End Property
Public Property Get getID_PROCESO_BASE() As Integer
    getID_PROCESO_BASE = PROCESO_BASE_ID
End Property
Public Property Get getTIPO_FRECUENCIA_ID() As Integer
    getTIPO_FRECUENCIA_ID = TIPO_FRECUENCIA_ID
End Property
Public Property Get getCLIENTE_ID() As Integer
    getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getTIPO_MUESTRA_ID() As Integer
    getTIPO_MUESTRA_ID = TIPO_MUESTRA_ID
End Property
Public Property Get getCONSERVACION() As String
    getCONSERVACION = CONSERVACION
End Property
Public Property Get getTOMA_FIN() As String
    getTOMA_FIN = TOMA_FIN
End Property
Public Property Get getSOLUCION_PROCEDENCIA_ID() As Integer
    getSOLUCION_PROCEDENCIA_ID = SOLUCION_PROCEDENCIA_ID
End Property
Public Property Get getVOLUMEN() As String
    getVOLUMEN = VOLUMEN
End Property
Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
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
Public Property Get getFACTURA_DETERMINACIONES() As Integer
    getFACTURA_DETERMINACIONES = FACTURA_DETERMINACIONES
End Property
Public Property Get getFICHA_ID() As Long
    getFICHA_ID = FICHA_ID
End Property
Public Property Get getCENTRO_ID() As Long
    getCENTRO_ID = CENTRO_ID
End Property
Public Property Get getAIRBUS_AREA_ID() As Long
    getAIRBUS_AREA_ID = AIRBUS_AREA_ID
End Property
Public Property Get getAIRBUS_LINEA_ID() As Long
    getAIRBUS_LINEA_ID = AIRBUS_LINEA_ID
End Property
'BUG-834-I
Public Property Get getANULADO() As Long
    getANULADO = anulado
End Property
'BUG-834-F
Public Function banos_cliente(cliente As Long, TIPO_MUESTRA As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    If TIPO_MUESTRA = 2 Or TIPO_MUESTRA = 6 Then
'BUG-834-I
'      consulta = "SELECT id_bano, nombre,precio, solucion_id FROM banos" & _
'                 " WHERE cliente_id=" & cliente & _
'                 " AND tipo_muestra_id=" & TIPO_MUESTRA & _
'                 " ORDER BY nombre"
      consulta = "SELECT id_bano, nombre,precio, solucion_id FROM banos" & _
                 " WHERE cliente_id=" & cliente & _
                 " AND tipo_muestra_id=" & TIPO_MUESTRA & " AND anulado <> 1" & _
                 " ORDER BY nombre"
'BUG-834-F
    Else
'BUG-834-I
'      consulta = "SELECT id_bano, nombre,precio, solucion_id FROM banos" & _
'                 " WHERE cliente_id=" & cliente & _
'                 " ORDER BY nombre"
      consulta = "SELECT id_bano, nombre,precio, solucion_id FROM banos" & _
                 " WHERE cliente_id=" & cliente & " AND anulado <> 1" & _
                 " ORDER BY nombre"
'BUG-834-F
    End If

    Set banos_cliente = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error en la clase 'banos_cliente'. ", vbCritical, Err.Description
End Function
Public Function revisar_facturacion(codbano As Long) As Boolean
    On Error GoTo fallo
    Dim RS As New ADODB.Recordset
    Dim consulta As String
    consulta = "Select * from banos where id_bano = " & codbano
    Set RS = datos_bd(consulta)
    If RS.RecordCount > 0 Then
        revisar_facturacion = RS("REVISAR_FACTURA")
    Else
        revisar_facturacion = False
    End If
    Exit Function
fallo:
    revisar_facturacion = False
End Function

Public Function cargar_bano(codbano As Long) As Boolean
    On Error GoTo fallo
    Dim RS As New ADODB.Recordset
    Dim consulta As String
    consulta = "Select * from banos where id_bano = " & codbano
    Set RS = datos_bd(consulta)
    If RS.RecordCount <> 0 Then
        id_bano = RS("id_bano")
        nombre = RS("NOMBRE")
        LINEA_ID = RS("LINEA_ID")
        INSTALACION_ID = RS("INSTALACION_ID")
        FORMATO_ID = RS("FORMATO_ID")
        SOLUCION_ID = RS("SOLUCION_ID")
        PROCESO_BASE_ID = RS("PROCESO_BASE_ID")
        TIPO_FRECUENCIA_ID = RS("TIPO_FRECUENCIA_ID")
        CLIENTE_ID = RS("CLIENTE_ID")
        TIPO_MUESTRA_ID = RS("TIPO_MUESTRA_ID")
        CONSERVACION = RS("CONSERVACION")
        TOMA_FIN = RS("TOMA_FIN")
        VOLUMEN = RS("VOLUMEN")
        OBSERVACIONES = RS("OBSERVACIONES")
        SOLUCION_PROCEDENCIA_ID = RS("SOLUCION_PROCEDENCIA_ID")
        PRECIO = RS("precio")
        TARIFA_CODIGO_ID = RS("TARIFA_CODIGO_ID")
        revisar_factura = RS("REVISAR_FACTURA")
        FACTURA_DETERMINACIONES = RS("FACTURA_DETERMINACIONES")
        FICHA_ID = RS("FICHA_ID")
        CENTRO_ID = RS("CENTRO_ID")
        AIRBUS_AREA_ID = RS("AIRBUS_AREA_ID")
        AIRBUS_LINEA_ID = RS("AIRBUS_LINEA_ID")
'BUG-834-I
        anulado = RS("ANULADO")
'BUG-834-F
        cargar_bano = True
    Else
        cargar_bano = False
    End If
    Exit Function
fallo:
    cargar_bano = False
    MsgBox "Error al cargar el baño (cargar_bano)", vbCritical, Err.Description
End Function
Public Function cargar_linea_bano(codbano As Long) As Boolean
    On Error GoTo fallo
    Dim RS As New ADODB.Recordset
    Dim consulta As String
    consulta = "Select LINEA_ID from banos where id_bano = " & codbano
    Set RS = datos_bd(consulta)
    If RS.RecordCount <> 0 Then
        LINEA_ID = RS("LINEA_ID")
        cargar_linea_bano = True
    Else
        cargar_linea_bano = False
    End If
    Exit Function
fallo:
    cargar_linea_bano = False
    MsgBox "Error al cargar el baño (cargar_bano)", vbCritical, Err.Description
End Function
Public Function tipo_solucion(codbano As Long) As Integer
    On Error GoTo fallo
    Dim consulta As String
    consulta = "Select solucion_id from banos where id_bano = " & codbano
    tipo_solucion = datos_bd(consulta).Fields(0)
    Exit Function
fallo:
    tipo_solucion = 0
End Function
'BUG-834-I
'Public Function Listado_Filtro(filtro As String, cliente As Long, pb As Long, solucion As Long, filtroCE As Boolean, Centro As Long) As ADODB.RecordSet
Public Function Listado_Filtro(filtro As String, cliente As Long, pb As Long, solucion As Long, filtroCE As Boolean, Centro As Long, ANULADA As Boolean) As ADODB.Recordset
'BUG-834-F
    On Error GoTo fallo
    Dim consulta As String
    Dim sCliente As String
    Dim spb As String
    Dim ssolucion As String
    Dim sce As String
    Dim scentro As String
'BUG-834-I
    Dim sAnulada As String
'BUG-834-F
    If cliente <> 0 Then
        sCliente = " and b.cliente_id = " & cliente
    End If
    If pb <> 0 Then
        spb = " and b.proceso_base_id = " & pb
    End If
    If solucion <> 0 Then
        ssolucion = " and b.solucion_id = " & solucion
    End If
    If filtroCE = True Then
        sce = " and b.FICHA_ID <> 0 "
    End If
    If Centro <> 0 Then
        scentro = " and b.CENTRO_ID = " & Centro
    End If
'BUG-834-I
    If ANULADA = False Then
        sAnulada = " and b.anulado = 0 "
    Else
        sAnulada = " "
    End If
'BUG-834-F
    consulta = "select b.id_bano, b.nombre as bano, cli.nombre, b.anulado " & _
        "  from banos b, clientes cli " & _
        " where b.cliente_id = cli.id_cliente " & _
        "   and b.nombre like '%" & Trim(filtro) & "%' " & _
        sCliente & _
        spb & _
        ssolucion & _
        sce & scentro & _
        sAnulada & _
        " order by b.nombre"
    Set Listado_Filtro = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al listar los baños (listado).", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "select b.id_bano, b.nombre as bano, cli.nombre,b.precio " & _
        "from banos b, clientes cli " & _
        "where b.cliente_id = cli.id_cliente " & _
        " order by b.nombre"
    Set Listado = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al listar los baños (listado).", vbCritical, Err.Description
End Function
Public Sub CrearID()
    Dim RS As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_BANO) FROM BANOS"
    Set RS = datos_bd(consulta)
    If IsNull(RS.Fields(0)) Or (RS.EOF And RS.BOF) Then  'si es nulo No se recupero ninguno
        id_bano = 1
    Else
        id_bano = RS.Fields(0) + 1
    End If
    Set RS = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearID
    consulta = "Insert into BANOS " & _
               " values(" & _
               id_bano & ",'" & nombre & "'," & _
               CLIENTE_ID & "," & SOLUCION_ID & "," & _
               PROCESO_BASE_ID & "," & TIPO_FRECUENCIA_ID & "," & _
               LINEA_ID & "," & INSTALACION_ID & "," & FORMATO_ID & ",'" & _
               CONSERVACION & "','" & TOMA_FIN & "','" & _
               VOLUMEN & "','" & OBSERVACIONES & "'," & SOLUCION_PROCEDENCIA_ID & "," & _
               TIPO_MUESTRA_ID & "," & TARIFA_CODIGO_ID & "," & FACTURA_DETERMINACIONES & "," & revisar_factura & "," & FICHA_ID & "," & CENTRO_ID & "," & _
               AIRBUS_AREA_ID & "," & AIRBUS_LINEA_ID & ",'" & PRECIO & "'," & anulado & ")"
    execute_bd consulta
    Insertar = id_bano
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el baño (Insertar)", vbCritical, Err.Description
End Function
Public Function Eliminar(BANO As Integer) As Boolean
    On Error GoTo fallo
    Dim RS As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM MUESTRAS WHERE BANO_ID = " & BANO
    Set RS = datos_bd(consulta)
    If RS.RecordCount <> 0 Then
        MsgBox "No puede eliminar el BAÑO. Existen muestras vinculadas.", vbCritical, App.Title
        Eliminar = False
        Exit Function
    End If
    
    execute_bd ("DELETE FROM DETERMINACIONES_ANALISIS WHERE BANO_ID =" & BANO)
    execute_bd ("DELETE FROM TIPOS_DATOS_ANALISIS WHERE BANO_ID =" & BANO)
    execute_bd ("DELETE FROM BANOS WHERE ID_BANO=" & BANO)
    execute_bd ("DELETE FROM TARIFAS_PRECIOS WHERE BANO_ID=" & BANO)
    
    execute_bd ("DELETE FROM CE_BANOS_ENSAYOS WHERE BANO_ID=" & BANO)
    execute_bd ("DELETE FROM CE_BANOS_PROBETAS WHERE BANO_ID=" & BANO)
    execute_bd ("DELETE FROM CE_BANOS_PROBETAS_MATERIALES WHERE BANO_ID=" & BANO)
    
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function Modificar(BANO As Long) As Long
    On Error GoTo fallo
    Dim consulta As String
    consulta = "Update BANOS set " & _
                  " NOMBRE = '" & nombre & "'," & _
                  " CLIENTE_ID = " & CLIENTE_ID & "," & _
                  " SOLUCION_ID = " & SOLUCION_ID & "," & _
                  " PROCESO_BASE_ID = " & PROCESO_BASE_ID & "," & _
                  " TIPO_FRECUENCIA_ID = " & TIPO_FRECUENCIA_ID & "," & _
                  " LINEA_ID = " & LINEA_ID & "," & _
                  " INSTALACION_ID = " & INSTALACION_ID & "," & _
                  " FORMATO_ID = '" & FORMATO_ID & "'," & _
                  " CONSERVACION = '" & CONSERVACION & "'," & _
                  " TOMA_FIN = '" & TOMA_FIN & "'," & _
                  " VOLUMEN = '" & VOLUMEN & "'," & _
                  " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                  " TARIFA_CODIGO_ID = " & TARIFA_CODIGO_ID & "," & _
                  " FICHA_ID = " & FICHA_ID & "," & _
                  " CENTRO_ID = " & CENTRO_ID & "," & _
                  " AIRBUS_AREA_ID = " & AIRBUS_AREA_ID & "," & _
                  " AIRBUS_LINEA_ID = " & AIRBUS_LINEA_ID & "," & _
                  " FACTURA_DETERMINACIONES = " & FACTURA_DETERMINACIONES & "," & _
                  " REVISAR_FACTURA = " & revisar_factura & "," & _
                  " ANULADO = " & anulado & "," & _
                  " SOLUCION_PROCEDENCIA_ID = " & SOLUCION_PROCEDENCIA_ID & "," & _
                  " TIPO_MUESTRA_ID = " & TIPO_MUESTRA_ID & "," & _
                  " precio = '" & PRECIO & "'" & _
                " WHERE ID_BANO = " & BANO
    execute_bd consulta
    Modificar = BANO
    Exit Function
fallo:
    Modificar = 0
    MsgBox "Error al Modificar el baño (Modificar)", vbCritical, Err.Description
End Function
Public Function Modificar_Codigo_Tarifa(BANO As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = " UPDATE BANOS " & _
               "  SET TARIFA_CODIGO_ID= " & TARIFA_CODIGO_ID & _
               " WHERE ID_BANO=" & BANO
    execute_bd consulta
    Modificar_Codigo_Tarifa = True
    Exit Function
fallo:
    Modificar_Codigo_Tarifa = False
    MsgBox "Error al Modificar_Codigo_Tarifa (BANO)", vbCritical, Err.Description
End Function

'BUG-834-I
Public Function Anular(BANO As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = " UPDATE BANOS " & _
               "  SET ANULADO= 1" & _
               " WHERE ID_BANO=" & BANO
    execute_bd consulta
    Anular = True
    Exit Function
fallo:
    Anular = False
    MsgBox "Error al Anular (BANO)", vbCritical, Err.Description
End Function
'BUG-834-F

Public Function Listado_Lineas(linea As Integer) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "select b.id_bano, b.nombre as bano, cli.nombre " & _
        "from banos b, clientes cli " & _
        "where b.cliente_id = cli.id_cliente " & _
        "  and b.linea_id = " & linea & _
        " order by b.nombre"
    Set Listado_Lineas = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al listar los baños (Listado_Lineas).", vbCritical, Err.Description
End Function
Public Function Listado_Lineas_Controladas(linea As Integer) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "select DISTINCT b.id_bano, b.nombre as bano, cli.nombre " & _
        "from banos b, clientes cli,banos_control c " & _
        "where b.cliente_id = cli.id_cliente " & _
        "  and b.linea_id = " & linea & _
        "  and b.id_bano = c.bano_id " & _
        " order by b.nombre"
    Set Listado_Lineas_Controladas = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al listar los baños (Listado_Lineas_Controladas).", vbCritical, Err.Description
End Function
Public Function Resultados_Banos(ID As Long, ANNO As Integer) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
'    consulta = "SELECT M.ID_MUESTRA,CONCAT(TM.CODIGO,'-',CAST(M.ID_PARTICULAR AS CHAR)),M.FECHA_MUESTREO, D.RESULTADO, D.TIPO_DETERMINACION_ID,M.ID_GENERAL,YEAR(M.FECHA_RECEPCION),M.ULT_EDICION_IMP " & _
'               "  FROM MUESTRAS AS M, DETERMINACIONES AS D,TIPOS_MUESTRA AS TM " & _
'               " WHERE M.ID_MUESTRA = D.MUESTRA_ID " & _
'               "   AND M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
'               "   AND M.BANO_ID  = " & ID & _
'               "   AND M.ANNO >= " & ANNO & _
'               "   AND M.ANULADA = 0 " & _
'               "   AND M.CERRADA = 1 " & _
'               " ORDER BY M.ID_MUESTRA DESC"
    consulta = "SELECT M.ID_MUESTRA,CONCAT(TM.CODIGO,'-',CAST(M.ID_PARTICULAR AS CHAR)),M.FECHA_MUESTREO, D.RESULTADO, D.TIPO_DETERMINACION_ID,M.ID_GENERAL,YEAR(M.FECHA_RECEPCION),M.ULT_EDICION_IMP " & _
               "      ,DV1.VALOR,DV2.VALOR,DV3.VALOR,DV4.VALOR, RE.AN_RUTA " & _
               "  FROM MUESTRAS AS M " & _
               "  LEFT JOIN DATOS_VALORES DV1 ON M.ID_MUESTRA = DV1.MUESTRA_ID AND DV1.TIPO_DATO_ID = " & TIPOS_DATOS.TIPOS_DATOS_OBSERVACIONES & _
               "  LEFT JOIN DATOS_VALORES DV2 ON M.ID_MUESTRA = DV2.MUESTRA_ID AND DV2.TIPO_DATO_ID = " & TIPOS_DATOS.TIPOS_DATOS_TEMPERATURA & _
               "  LEFT JOIN DATOS_VALORES DV3 ON M.ID_MUESTRA = DV3.MUESTRA_ID AND DV3.TIPO_DATO_ID = " & TIPOS_DATOS.TIPOS_DATOS_HORA_TOMA_MUESTRA & _
               "  LEFT JOIN DATOS_VALORES DV4 ON M.ID_MUESTRA = DV4.MUESTRA_ID AND DV4.TIPO_DATO_ID = " & TIPOS_DATOS.TIPOS_DATOS_NIVEL & _
               "  LEFT JOIN RECARGAS RE ON M.ID_MUESTRA = RE.MUESTRA_ID " & _
               " INNER JOIN DETERMINACIONES D ON M.ID_MUESTRA = D.MUESTRA_ID " & _
               " INNER JOIN TIPOS_MUESTRA  TM ON M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
               " WHERE M.BANO_ID  = " & ID & _
               "   AND M.ANNO >= " & ANNO & _
               "   AND M.ANULADA = 0 " & _
               "   AND M.CERRADA = 1 " & _
               " ORDER BY M.ID_MUESTRA DESC"
    Set Resultados_Banos = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al buscar los resultados de los baños (Resultados_Banos).", vbCritical, Err.Description
End Function
Public Function Resultados_Banos_fecha(BANO As Long, fecha As String) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    Dim oBANO As New clsBanos
    oBANO.cargar_bano BANO
    consulta = "SELECT M.ID_MUESTRA,CONCAT(TM.CODIGO,'-',CAST(M.ID_PARTICULAR AS CHAR)),M.FECHA_MUESTREO, D.RESULTADO, D.TIPO_DETERMINACION_ID,M.ID_GENERAL,YEAR(M.FECHA_RECEPCION),M.ULT_EDICION_IMP " & _
               "  FROM MUESTRAS AS M, DETERMINACIONES AS D,TIPOS_MUESTRA AS TM, DETERMINACIONES_ANALISIS DA " & _
               " WHERE M.ID_MUESTRA = D.MUESTRA_ID " & _
               "   AND M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
               "   AND M.BANO_ID  = " & BANO & _
               "   AND M.CLIENTE_ID = " & oBANO.getCLIENTE_ID & _
               "   AND M.BANO_ID = DA.BANO_ID " & _
               "   AND DA.TIPO_DETERMINACION_ID = D.TIPO_DETERMINACION_ID " & _
               "   AND DA.TIPO_ANALISIS_ID = 0 " & _
               "   AND M.FECHA_RECEPCION > '" & Format(fecha, "yyyy-mm-dd") & "'" & _
               "   AND M.ANNO >= " & Year(fecha) & _
               "   AND M.ANULADA = 0 " & _
               "   AND M.CERRADA = 1 " & _
               " ORDER BY M.ID_MUESTRA ASC,DA.ORDEN"
'               " AND DA.REQUERIDA = 1 "
    Set oBANO = Nothing
    Set Resultados_Banos_fecha = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al buscar los resultados de los baños (Resultados_Banos).", vbCritical, Err.Description
End Function
Public Function Resultados_Bano_por_fecha(BANO As Long, fecha As String) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT M.ID_MUESTRA,CONCAT(TM.CODIGO,'-',CAST(M.ID_PARTICULAR AS CHAR)),M.FECHA_MUESTREO, D.RESULTADO, D.TIPO_DETERMINACION_ID,M.ID_GENERAL,YEAR(M.FECHA_RECEPCION),M.ULT_EDICION_IMP " & _
               "  FROM MUESTRAS AS M, DETERMINACIONES AS D,TIPOS_MUESTRA AS TM " & _
               " WHERE M.ID_MUESTRA = D.MUESTRA_ID " & _
               "   AND M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
               "   AND M.BANO_ID  = " & BANO & _
               "   AND M.FECHA_RECEPCION = '" & Format(fecha, "yyyy-mm-dd") & "'" & _
               "   AND M.ANULADA = 0 " & _
               "   AND M.CERRADA = 1 " & _
               " ORDER BY M.ID_MUESTRA ASC"
    Set Resultados_Bano_por_fecha = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al buscar los resultados de los baños (Resultados_Bano_por_fecha).", vbCritical, Err.Description
End Function

Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_BANO, NOMBRE FROM BANOS " & _
               " ORDER BY NOMBRE"
    Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Listado_por_Cliente(ID As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_BANO, NOMBRE FROM BANOS " & _
               " WHERE CLIENTE_ID = " & ID & _
               " ORDER BY NOMBRE"
    Set Listado_por_Cliente = datos_bd(consulta)
End Function
Public Function Listado_con_CE() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_BANO, NOMBRE FROM BANOS " & _
               " WHERE FICHA_ID <> 0 " & _
               " ORDER BY NOMBRE"
    Set Listado_con_CE = datos_bd(consulta)
End Function

Public Function Listado_por_Cliente_con_CE(ID As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_BANO, NOMBRE FROM BANOS " & _
               " WHERE CLIENTE_ID = " & ID & _
               "   AND FICHA_ID <> 0 " & _
               " ORDER BY NOMBRE"
    Set Listado_por_Cliente_con_CE = datos_bd(consulta)
End Function
Public Function Listado_Clientes() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT DISTINCT C.ID_CLIENTE AS C1,C.NOMBRE AS C2 " & _
               "  FROM CLIENTES C, BANOS B " & _
               " WHERE B.CLIENTE_ID = C.ID_CLIENTE " & _
               "   AND C.ANULADO = 0 " & _
               " ORDER BY C.NOMBRE "
    Set Listado_Clientes = datos_bd(consulta)
End Function
Public Function Listado_Clientes_CE() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT DISTINCT C.ID_CLIENTE AS C1,C.NOMBRE AS C2 " & _
               "  FROM CLIENTES C, BANOS B " & _
               " WHERE B.CLIENTE_ID = C.ID_CLIENTE " & _
               "   AND C.ANULADO = 0 " & _
               "   AND B.FICHA_ID <> 0 " & _
               " ORDER BY C.NOMBRE "
    Set Listado_Clientes_CE = datos_bd(consulta)
End Function
Public Function Listado_Soluciones() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT DISTINCT S.ID_SOLUCION AS C1, S.NOMBRE AS C2 " & _
               "  FROM BANOS B,SOLUCIONES S " & _
               " WHERE B.SOLUCION_ID = S.ID_SOLUCION " & _
               " ORDER BY S.NOMBRE"
    Set Listado_Soluciones = datos_bd(consulta)
End Function
Public Function Listado_Soluciones_por_proceso(PROCESO As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT DISTINCT S.ID_SOLUCION AS C1, S.NOMBRE AS C2 " & _
               "  FROM BANOS B,SOLUCIONES S " & _
               " WHERE B.SOLUCION_ID = S.ID_SOLUCION " & _
               "   AND B.PROCESO_BASE_ID = " & PROCESO & _
               " ORDER BY S.NOMBRE"
    Set Listado_Soluciones_por_proceso = datos_bd(consulta)
End Function

Public Function Listado_Soluciones_por_Cliente(ID As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT DISTINCT S.ID_SOLUCION AS C1, S.NOMBRE AS C2 " & _
               "  FROM BANOS B,SOLUCIONES S " & _
               " WHERE B.CLIENTE_ID = " & ID & _
               "   AND B.SOLUCION_ID = S.ID_SOLUCION " & _
               " ORDER BY S.NOMBRE"
    Set Listado_Soluciones_por_Cliente = datos_bd(consulta)
End Function
Public Function Listado_Banos_por_Solucion(ID As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT DISTINCT B.ID_BANO AS C1, B.NOMBRE AS C2 " & _
               "  FROM BANOS B " & _
               " WHERE B.SOLUCION_ID = " & ID & _
               " ORDER BY B.NOMBRE"
    Set Listado_Banos_por_Solucion = datos_bd(consulta)
End Function
Public Function Listado_Tarifa(TARIFA As Long, cliente As Long, solucion As Long, CODIGO_TARIFA As Long) As ADODB.Recordset
    Dim consulta As String
    Dim s1 As String
    Dim s2 As String
    If CODIGO_TARIFA <> 0 Then
        s1 = " AND a.TARIFA_CODIGO_ID = " & CODIGO_TARIFA
    End If
    If solucion <> 0 Then
        s2 = "   AND a.solucion_id = " & solucion
    End If
    If cliente = 0 Then
        consulta = "SELECT distinct a.id_bano,a.nombre,c.codigo,d.precio,b.precio,a.tarifa_codigo_id " & _
                   "  FROM banos a " & _
                   "  LEFT JOIN tarifas_precios b on a.id_bano = b.bano_id and b.tipo_determinacion_id = 0 and b.tipo_analisis_id = 0 " & _
                   "  LEFT JOIN tarifas_precios d on a.id_bano = d.bano_id and d.tipo_determinacion_id = 0 and d.tipo_analisis_id = 0 " & _
                   "  LEFT JOIN tarifas_codigos c on a.tarifa_codigo_id = c.id_codigo " & _
                   " INNER JOIN tarifas       tar on d.tarifa_id = tar.tarifa_origen_id  " & _
                   " WHERE b.tarifa_id = " & TARIFA & _
                   "   AND b.tarifa_id = tar.id_tarifa " & _
                   s1 & s2 & _
                   " ORDER BY a.nombre"
    Else
        consulta = "SELECT distinct a.id_bano,a.nombre,c.codigo,d.precio,b.precio,a.tarifa_codigo_id " & _
                   "  FROM banos a " & _
                   "  LEFT JOIN tarifas_precios b on a.id_bano = b.bano_id and b.tipo_determinacion_id = 0 and b.tipo_analisis_id = 0 " & _
                   "  LEFT JOIN tarifas_precios d on a.id_bano = d.bano_id and d.tipo_determinacion_id = 0 and d.tipo_analisis_id = 0 " & _
                   "  LEFT JOIN tarifas_codigos c on a.tarifa_codigo_id = c.id_codigo " & _
                   " INNER JOIN tarifas       tar on d.tarifa_id = tar.tarifa_origen_id  " & _
                   " WHERE a.cliente_id = " & cliente & _
                   " and b.tarifa_id = " & TARIFA & _
                   "   AND b.tarifa_id = tar.id_tarifa " & _
                   s1 & s2 & _
                   " ORDER BY a.nombre"
    End If
    Set Listado_Tarifa = datos_bd(consulta)
End Function
Public Function Listado_por_determinacion(DETERMINACION As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "select b.id_bano, b.nombre as bano, cli.nombre " & _
               "  from banos b, clientes cli,determinaciones_analisis da " & _
               " where b.cliente_id = cli.id_cliente " & _
               "   and b.id_bano = da.bano_id " & _
               "   and da.tipo_determinacion_id = " & DETERMINACION & _
               " order by b.nombre"
    Set Listado_por_determinacion = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al listar los baños (Listado_por_determinacion).", vbCritical, Err.Description
End Function
Public Function Listado_por_solucion_donde(solucion As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "select b.id_bano, b.nombre as bano, cli.nombre " & _
               "  from banos b, clientes cli " & _
               " where b.cliente_id = cli.id_cliente " & _
               "   and b.solucion_id = " & solucion & _
               " order by b.nombre"
    Set Listado_por_solucion_donde = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al listar los baños (Listado_por_solucion_donde).", vbCritical, Err.Description
End Function
Public Function Listado_por_PB_donde(pb As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "select b.id_bano, b.nombre as bano, cli.nombre " & _
               "  from banos b, clientes cli " & _
               " where b.cliente_id = cli.id_cliente " & _
               "   and b.proceso_base_id = " & pb & _
               " order by b.nombre"
    Set Listado_por_PB_donde = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al listar los baños (Listado_por_PB_donde).", vbCritical, Err.Description
End Function
Public Function Listado_por_FICHA_CE_donde(FICHA As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "select b.id_bano, b.nombre as bano, cli.nombre " & _
               "  from banos b, clientes cli " & _
               " where b.cliente_id = cli.id_cliente " & _
               "   and b.FICHA_id = " & FICHA & _
               " order by b.nombre"
    Set Listado_por_FICHA_CE_donde = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al listar los baños (Listado_por_FICHA_CE_donde).", vbCritical, Err.Description
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = "CLIENTE_ID"
        .setFK_VALOR = FK
        .setTABLA = "BANOS"
        .setDESCRIPCION = "Baños"
        .setPK = "ID_BANO"
        .setCAMPO = "NOMBRE"
        .setFILTRO = filtro
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = FORMULARIO
    End With
End Function

Public Function Modificar_Ficha(BANO As Integer, FICHA As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = " UPDATE BANOS " & _
               "  SET FICHA_ID= " & FICHA & _
               " WHERE ID_BANO=" & BANO
    execute_bd consulta
    Modificar_Ficha = True
    Exit Function
fallo:
    Modificar_Ficha = False
    MsgBox "Error al Modificar_Ficha (BANO)", vbCritical, Err.Description
End Function
