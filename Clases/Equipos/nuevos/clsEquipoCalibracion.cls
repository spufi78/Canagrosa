VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipoCalibracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQ_CALIBRACION_EQUIPOS
'***************************************************************
Private ID_CALIBRACION As Long
Private MUESTRA_ID As Long
Private EQUIPO_ID As Long
Private TIPO_ID As Long
Private PERIODICIDAD_ID As Long
Private PROCEDIMIENTO_ID As Long
Private CALIBRADOR_EXTERNO_ID As Long
Private CALIBRADOR_INTERNO_ID As Long
Private fecha_actual As String
Private fecha_proxima As String
Private RANGO_MIN As String
Private RANGO_MAX As String
Private UNIDADES_ID As Long
Private RUTA_PLANTILLA As String
Private RUTA_CERTIFICADO As String
Private RUTA_EVALUACION As String
Private ESTADO As Integer
Private RESULTADO As Integer
Private CUSERID As Long
Private MUSERID As Long
Private TS As String
Private ID_AUX As Integer
Private responsable As String
Private mvarobjLimitacionesUso As New clsGenericCollection
Private mvarobjAccesoriosEquipo As New clsGenericCollection
Private mvarobjAccesoriosCalibracion As New clsGenericCollection
Private PROCEDIMIENTO As String
Private REACTIVOS As String
Private INCIDENCIAS As String
Private UBICACION_ID As Long
Private BUSQUEDAS As Integer
Private ENVIADO As Integer
Private EDICION As Integer
Private INCERTIDUMBRE_USO As String
Private INCERTIDUMBRE As String
Private SEGREGADO As Integer
Private ETIQUETADO As Integer

Private Const CONST_DESCRIPCION As Integer = 0
Private Const CONST_RANGO_MIN As Integer = 1
Private Const CONST_RANGO_MAX As Integer = 2
Private Const CONST_unidad As Integer = 3
Private Const CONST_RESULTADO_CAL As Integer = 4
Private Const CONST_TOLERANCIA As Integer = 5
Private Const CONST_INCERTIDUMBRE As Integer = 6
Private Const CONST_CORRECCION As Integer = 7
Private Const CONST_ID_RESULTADO As Integer = 8
Private Const CONST_ID_UNIDAD As Integer = 9

'JVG.2010.08.30
Private fecha_prevista As String



Private Function comprobar_anterior_cal_no_conforme() As Boolean
Dim rs As ADODB.Recordset

    comprobar_anterior_cal_no_conforme = False

    Set rs = datos_bd("select * from eq_calibracion_equipos where equipo_id = " & EQUIPO_ID & " AND id_calibracion <> " & ID_CALIBRACION & " order by id_calibracion desc")
        
    
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
'        If CInt(rs!ESTADO) = 2 Then
        If CInt(rs!RESULTADO) = CAL_RESULTADOS.CAL_RESULTADO_NO_CONFORME Then
            comprobar_anterior_cal_no_conforme = True
            Set rs = Nothing
        End If
    End If
    
    Set rs = Nothing
    
    
End Function

Public Function crear_calibracion_pendiente(ByVal prm_id_cal As Long, ByVal prm_id_equipo As Long) As Boolean

    '1 crea una calibracion a partir de esta
    '2 la guarda
    '3 crea el registro de operacion pendiente
        
    Dim oCal As New clsEquipoCalibracion
    Dim oPer As New clsEquiposPeriodicidad
    Dim oOP As clsEquiposOperacionesPendientes
    Dim rs As ADODB.Recordset, consulta As String
    
   On Error GoTo crear_calibracion_pendiente_Error

    oCal.Carga prm_id_cal
    
    ' Si es un equipo de cliente, no crear la siguiente calibracion
    Dim oEquipo As New clsEquipos
    oEquipo.Carga oCal.getEQUIPO_ID
    If Not IsNull(oEquipo.getCLIENTE_ID) Then
        If oEquipo.getCLIENTE_ID <> 0 Then
            log "No se crea siguiente calibración, es un equipo de cliente."
            Exit Function
        End If
    End If
    
    
    ' cambia las propiedades
    
    ' calcula la fecha en que se va a realizar la operacion
    ' cuando el estado es cerrado NO CONFORME ,
    If ESTADO = CAL_ESTADOS.CAL_ESTADO_REALIZADA Then
        oCal.setFECHA_ACTUAL = oPer.calcular_fecha(oCal.getFECHA_ACTUAL, oCal.getPERIODICIDAD_ID)
    End If
    
    ' una vez conocida la fecha Prevista, estima la fecha de la siguiente para cuando se cierre
    oCal.setFECHA_PREVISTA = oCal.getFECHA_ACTUAL
    oCal.setFECHA_PROXIMA = oPer.calcular_fecha(oCal.getFECHA_ACTUAL, oCal.getPERIODICIDAD_ID)
    oCal.setESTADO = CAL_ESTADOS.CAL_ESTADO_PREVISTA
    oCal.setRESULTADO = CAL_RESULTADOS.CAL_RESULTADO_CONFORME
    oCal.setRUTA_CERTIFICADO = ""
    oCal.setRUTA_EVALUACION = ""
    oCal.setRUTA_PLANTILLA = ""
        
    ' Los Parametros no se insertan automáticamente, con lo que habrá que hacerlo posteriormente
    ' Se eliminan los Accesorios y reactivos para que no se insertan
    Set oCal.AccesoriosCalibracion = New clsGenericCollection
    oCal.setREACTIVOS = ""
    oCal.setINCIDENCIAS = ""
    oCal.setUBICACION_ID = 0
    oCal.setBUSQUEDAS = 0
    oCal.setEDICION = 0
    oCal.setINCERTIDUMBRE_USO = 0
    oCal.setINCERTIDUMBRE = ""
    oCal.setSEGREGADO = 0
    oCal.setETIQUETADO = 0
    oCal.setENVIADO = 0
    oCal.Insertar False
    
    'Una vez insertado el clon previsto, se insertan Accesorios, Reactivos y Parámetros con Resultados a 0
    ' Accesorios
    consulta = "insert into eq_calibracion_equipos_accesorios (ID_CALIBRACION, EQUIPO_ID, ACCESORIO_ID)" & _
                " select " & oCal.getID_CALIBRACION & " as ID_CALIBRACION, " & oCal.getEQUIPO_ID & " AS EQUIPO_ID, eca.ACCESORIO_ID from eq_calibracion_equipos_accesorios eca WHERE ECA.ID_CALIBRACION = " & prm_id_cal
    execute_bd consulta
    
    'Reactivos
    consulta = "insert into eq_calibracion_reactivos (CALIBRACION_ID, EQUIPO_ID, REACTIVO_ID) " & _
                " select " & oCal.getID_CALIBRACION & " as CALIBRACION_ID, " & oCal.getEQUIPO_ID & " AS EQUIPO_ID, e.REACTIVO_ID from eq_calibracion_reactivos e WHERE E.CALIBRACION_ID = " & prm_id_cal
    execute_bd consulta
    
    ' Parámetros
    Set rs = datos_bd("select * from eq_calibracion_parametros_resultados where calibracion_id = " & prm_id_cal)
    
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        While Not rs.EOF
            consulta = "insert into eq_calibracion_parametros_resultados (equipo_id, id_resultado, calibracion_id, unidad_id, "
            consulta = consulta & "rango_min, rango_max, resultado, tolerancia_max, correccion, incertidumbre, descripcion) "
            consulta = consulta & " SELECT "
            consulta = consulta & prm_id_equipo & " as equipo_id, "
            consulta = consulta & " cast(coalesce(max(id_resultado),0)+ 1 as unsigned) as id_resultado, "
            consulta = consulta & oCal.getID_CALIBRACION & " as calibracion_id, "
            consulta = consulta & rs!UNIDAD_ID & " as unidad_id , "
            consulta = consulta & Replace(Format(rs!RANGO_MIN, "0.000000"), ",", ".") & " as rango_min , "
            consulta = consulta & Replace(Format(rs!RANGO_MAX, "0.000000"), ",", ".") & " as rango_max , "
            consulta = consulta & " 0.000000 as resultado , "
            consulta = consulta & Replace(Format(rs!TOLERANCIA_MAX, "0.000000"), ",", ".") & " as tolerancia_max , "
            consulta = consulta & Replace(Format(rs!CORRECCION, "0.000000"), ",", ".") & " as correccion , "
            consulta = consulta & Replace(Format(rs!INCERTIDUMBRE, "0.000000"), ",", ".") & " as incertidumbre , "
            consulta = consulta & "'" & rs!DESCRIPCION & "' as descripcion "
            consulta = consulta & " from eq_calibracion_parametros_resultados"
            
            execute_bd consulta
            rs.MoveNext
        Wend
    End If
    
    Set oOP = New clsEquiposOperacionesPendientes
    oOP.crear_calibracion_pendiente prm_id_cal, oCal.getID_CALIBRACION, prm_id_equipo, oPer
    
    Set oCal = Nothing
    Set oPer = Nothing
    Set oOP = Nothing
    Set rs = Nothing

   On Error GoTo 0
   Exit Function

crear_calibracion_pendiente_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure crear_calibracion_pendiente of Módulo de clase clsEquipoCalibracion"
    
End Function

Public Property Get getREACTIVOS() As String
    getREACTIVOS = REACTIVOS
End Property
Public Property Get getINCIDENCIAS() As String
    getINCIDENCIAS = INCIDENCIAS
End Property
Public Property Get getUBICACION_ID() As Long
    getUBICACION_ID = UBICACION_ID
End Property
Public Property Get getBUSQUEDAS() As Integer
    getBUSQUEDAS = BUSQUEDAS
End Property
Public Property Get getENVIADO() As Integer
    getENVIADO = ENVIADO
End Property
Public Property Get getEDICION() As Integer
    getEDICION = EDICION
End Property
Public Property Get getINCERTIDUMBRE_USO() As String
    getINCERTIDUMBRE_USO = INCERTIDUMBRE_USO
End Property
Public Property Get getINCERTIDUMBRE() As String
    getINCERTIDUMBRE = INCERTIDUMBRE
End Property
Public Property Get getSEGREGADO() As Integer
    getSEGREGADO = SEGREGADO
End Property
Public Property Get getETIQUETADO() As Integer
    getETIQUETADO = ETIQUETADO
End Property

Public Property Let setFECHA_PREVISTA(dato As String)
    fecha_prevista = dato
End Property
Public Property Get getFECHA_PREVISTA() As String
    getFECHA_PREVISTA = fecha_prevista
End Property

Public Property Let setREACTIVOS(ByVal dato As String)
    REACTIVOS = dato
End Property
Public Property Let setINCIDENCIAS(ByVal dato As String)
    INCIDENCIAS = dato
End Property
Public Property Let setUBICACION_ID(ByVal dato As Long)
    UBICACION_ID = dato
End Property
Public Property Let setBUSQUEDAS(ByVal dato As Integer)
    BUSQUEDAS = dato
End Property
Public Property Let setENVIADO(ByVal dato As Integer)
    ENVIADO = dato
End Property
Public Property Let setEDICION(ByVal dato As Integer)
    EDICION = dato
End Property
Public Property Let setINCERTIDUMBRE_USO(ByVal dato As String)
    INCERTIDUMBRE_USO = dato
End Property
Public Property Let setINCERTIDUMBRE(ByVal dato As String)
    INCERTIDUMBRE = dato
End Property
Public Property Let setSEGREGADO(ByVal dato As Integer)
    SEGREGADO = dato
End Property
Public Property Let setETIQUETADO(ByVal dato As Integer)
    ETIQUETADO = dato
End Property
Public Property Let setPROCEDIMIENTO(ByVal dato As String)
    PROCEDIMIENTO = dato
End Property
Public Property Get getPROCEDIMIENTO() As String
    getPROCEDIMIENTO = PROCEDIMIENTO
End Property
Private Sub Class_Initialize()
    mvarobjLimitacionesUso.KeyName = "setID_LIMITACION_USO"
    mvarobjAccesoriosEquipo.KeyName = "setID_ACCESORIO"
    mvarobjAccesoriosCalibracion.KeyName = "setID_ACCESORIO"
End Sub

Private Sub Class_Terminate()
    Set mvarobjLimitacionesUso = Nothing
    Set mvarobjAccesoriosEquipo = Nothing
    Set mvarobjAccesoriosCalibracion = Nothing
End Sub
Public Property Get LimitacionesUso() As clsGenericCollection
    Set LimitacionesUso = mvarobjLimitacionesUso
End Property

Public Property Set LimitacionesUso(objLimitacionesUso As clsGenericCollection)
    Set mvarobjLimitacionesUso = objLimitacionesUso
End Property

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQ_CALIBRACION_EQUIPOS
'***************************************************************
Public Property Let setID_CALIBRACION(ByVal dato As Long)
        ID_CALIBRACION = dato
End Property
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setTIPO_ID(ByVal dato As Long)
        TIPO_ID = dato
End Property
Public Property Let setPERIODICIDAD_ID(ByVal dato As Long)
        PERIODICIDAD_ID = dato
End Property
Public Property Let setPROCEDIMIENTO_ID(ByVal dato As Long)
        PROCEDIMIENTO_ID = dato
End Property
Public Property Let setCALIBRADOR_EXTERNO_ID(ByVal dato As Long)
        CALIBRADOR_EXTERNO_ID = dato
End Property
Public Property Let setCALIBRADOR_INTERNO_ID(ByVal dato As Long)
        CALIBRADOR_INTERNO_ID = dato
End Property
Public Property Let setFECHA_ACTUAL(ByVal dato As String)
        fecha_actual = dato
End Property
Public Property Let setFECHA_PROXIMA(ByVal dato As String)
        fecha_proxima = dato
End Property
Public Property Let setRANGO_MIN(ByVal dato As String)
        RANGO_MIN = dato
End Property
Public Property Let setRANGO_MAX(ByVal dato As String)
        RANGO_MAX = dato
End Property
Public Property Let setUNIDADES_ID(ByVal dato As Long)
        UNIDADES_ID = dato
End Property
Public Property Let setRUTA_PLANTILLA(ByVal dato As String)
        RUTA_PLANTILLA = dato
End Property
Public Property Let setRUTA_CERTIFICADO(ByVal dato As String)
        RUTA_CERTIFICADO = dato
End Property
Public Property Let setRUTA_EVALUACION(ByVal dato As String)
        RUTA_EVALUACION = dato
End Property
Public Property Let setESTADO(ByVal dato As Integer)
        ESTADO = dato
End Property
Public Property Let setRESULTADO(ByVal dato As Integer)
        RESULTADO = dato
End Property
Public Property Let setCUSERID(ByVal dato As Long)
        CUSERID = dato
End Property
Public Property Let setMUSERID(ByVal dato As Long)
        MUSERID = dato
End Property
Public Property Let setTS(ByVal dato As String)
        TS = dato
End Property
Public Property Let setID_AUX(ByVal dato As Integer)
        ID_AUX = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQ_CALIBRACION_EQUIPOS
'***************************************************************
Public Property Get getID_CALIBRACION() As Long
        getID_CALIBRACION = ID_CALIBRACION
End Property
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getTIPO_ID() As Long
        getTIPO_ID = TIPO_ID
End Property
Public Property Get getPERIODICIDAD_ID() As Long
        getPERIODICIDAD_ID = PERIODICIDAD_ID
End Property
Public Property Get getPROCEDIMIENTO_ID() As Long
        getPROCEDIMIENTO_ID = PROCEDIMIENTO_ID
End Property
Public Property Get getCALIBRADOR_EXTERNO_ID() As Long
        getCALIBRADOR_EXTERNO_ID = CALIBRADOR_EXTERNO_ID
End Property
Public Property Get getRESPONSABLE() As String
        getRESPONSABLE = responsable
End Property

Public Property Let setRESPONSABLE(ByVal dato As String)
        responsable = dato
End Property
Public Property Get getCALIBRADOR_INTERNO_ID() As Long
        getCALIBRADOR_INTERNO_ID = CALIBRADOR_INTERNO_ID
End Property
Public Property Get getFECHA_ACTUAL() As String
        getFECHA_ACTUAL = fecha_actual
End Property
Public Property Get getFECHA_PROXIMA() As String
        getFECHA_PROXIMA = fecha_proxima
End Property
Public Property Get getRANGO_MIN() As String
        getRANGO_MIN = RANGO_MIN
End Property
Public Property Get getRANGO_MAX() As String
        getRANGO_MAX = RANGO_MAX
End Property
Public Property Get getUNIDADES_ID() As Long
        getUNIDADES_ID = UNIDADES_ID
End Property
Public Property Get getRUTA_PLANTILLA() As String
        getRUTA_PLANTILLA = RUTA_PLANTILLA
End Property
Public Property Get getRUTA_CERTIFICADO() As String
        getRUTA_CERTIFICADO = RUTA_CERTIFICADO
End Property
Public Property Get getRUTA_EVALUACION() As String
        getRUTA_EVALUACION = RUTA_EVALUACION
End Property
Public Property Get getESTADO() As Integer
        getESTADO = ESTADO
End Property
Public Property Get getRESULTADO() As Integer
        getRESULTADO = RESULTADO
End Property
Public Property Get getCUSERID() As Long
        getCUSERID = CUSERID
End Property
Public Property Get getMUSERID() As Long
        getMUSERID = MUSERID
End Property
Public Property Get getTS() As String
        getTS = TS
End Property

Public Property Get getID_AUX() As Integer
        getID_AUX = ID_AUX
End Property

'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQ_CALIBRACION_EQUIPOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        
        
        consulta = " SELECT t1.*, COALESCE(ca_documentos.NOMBRE, '') as PROCEDIMIENTO FROM (select eq_calibracion_equipos.*, proveedores.NOMBRE as RESPONSABLE"
        consulta = consulta & " From eq_calibracion_equipos"
        consulta = consulta & " inner join proveedores on eq_calibracion_equipos.CALIBRADOR_EXTERNO_ID = proveedores.ID_PROVEEDOR"
        consulta = consulta & " where eq_calibracion_equipos.tipo_id = 2 and eq_calibracion_equipos.ID_CALIBRACION = " & ID
        consulta = consulta & " Union"
        consulta = consulta & " select eq_calibracion_equipos.*, concat(usuarios.NOMBRE, ' ', usuarios.APELLIDOS) as RESPONSABLE"
        consulta = consulta & " From eq_calibracion_equipos"
        consulta = consulta & " inner join usuarios on eq_calibracion_equipos.CALIBRADOR_INTERNO_ID = usuarios.ID_EMPLEADO"
        consulta = consulta & " where eq_calibracion_equipos.tipo_id = 1 and eq_calibracion_equipos.ID_CALIBRACION = " & ID
        consulta = consulta & " ) as t1 "
        consulta = consulta & " LEFT join ca_documentos ON T1.PROCEDIMIENTO_ID = CA_DOCUMENTOS.ID_DOCUMENTO"
        consulta = consulta & " WHERE t1.ID_CALIBRACION = " & ID
                
        Set rs = datos_bd(consulta)
        
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_CALIBRACION = rs("ID_CALIBRACION")
                MUESTRA_ID = rs("MUESTRA_ID")
                EQUIPO_ID = rs("EQUIPO_ID")
                TIPO_ID = rs("TIPO_ID")
                PERIODICIDAD_ID = rs("PERIODICIDAD_ID")
                PROCEDIMIENTO_ID = rs("PROCEDIMIENTO_ID")
                PROCEDIMIENTO = rs("PROCEDIMIENTO")
                CALIBRADOR_EXTERNO_ID = rs("CALIBRADOR_EXTERNO_ID")
                CALIBRADOR_INTERNO_ID = rs("CALIBRADOR_INTERNO_ID")
                fecha_actual = rs("FECHA_ACTUAL")
                fecha_proxima = rs("FECHA_PROXIMA")
                fecha_prevista = rs("FECHA_prevista")
                RANGO_MIN = rs("RANGO_MIN")
                RANGO_MAX = rs("RANGO_MAX")
                UNIDADES_ID = rs("UNIDADES_ID")
                RUTA_PLANTILLA = rs("RUTA_PLANTILLA")
                RUTA_CERTIFICADO = rs("RUTA_CERTIFICADO")
                RUTA_EVALUACION = rs("RUTA_EVALUACION")
                ESTADO = rs("ESTADO")
                RESULTADO = rs("RESULTADO")
                CUSERID = rs("CUSERID")
                MUSERID = rs("MUSERID")
                responsable = rs("RESPONSABLE")
                TS = rs("TS")
                If Not IsNull(rs("INCIDENCIAS")) Then
                    INCIDENCIAS = rs("INCIDENCIAS")
                End If
                If Not IsNull(rs("UBICACION_ID")) Then
                    UBICACION_ID = rs("UBICACION_ID")
                Else
                    UBICACION_ID = 0
                End If
                If Not IsNull(rs("BUSQUEDAS")) Then
                    BUSQUEDAS = rs("BUSQUEDAS")
                Else
                    BUSQUEDAS = 0
                End If
                If Not IsNull(rs("ENVIADO")) Then
                    ENVIADO = rs("ENVIADO")
                Else
                    ENVIADO = 0
                End If
                If Not IsNull(rs("EDICION")) Then
                    EDICION = rs("EDICION")
                Else
                    EDICION = 0
                End If
                If Not IsNull(rs("INCERTIDUMBRE_USO")) Then
                    INCERTIDUMBRE_USO = rs("INCERTIDUMBRE_USO")
                Else
                    INCERTIDUMBRE_USO = 0
                End If
                If Not IsNull(rs("INCERTIDUMBRE")) Then
                    INCERTIDUMBRE = rs("INCERTIDUMBRE")
                Else
                    INCERTIDUMBRE = ""
                End If
                If Not IsNull(rs("SEGREGADO")) Then
                    SEGREGADO = rs("SEGREGADO")
                Else
                    SEGREGADO = 0
                End If
                If Not IsNull(rs("ETIQUETADO")) Then
                    ETIQUETADO = rs("ETIQUETADO")
                Else
                    ETIQUETADO = 0
                End If
                Carga = True
                
                Call Carga_AccesoriosCalibracion
                Call Carga_Reactivos
                
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquipoCalibracion)", vbCritical, Err.Description
End Function
Public Function CargaPorMuestraId(idMuestra As Long) As Long
        Dim rs As ADODB.Recordset
        Dim c As String
        c = "select id_calibracion from eq_calibracion_equipos " & _
            " where muestra_id = " & idMuestra
        Set rs = datos_bd(c)
        If rs.RecordCount = 0 Then
            CargaPorMuestraId = 0
        Else
            CargaPorMuestraId = rs("ID_CALIBRACION")
        End If
        Set rs = Nothing
End Function

Public Sub Carga_Reactivos()
    Dim strSql As String, rs As ADODB.Recordset
    
    strSql = "SELECT * FROM eq_calibracion_reactivos WHERE CALIBRACION_ID = " & CStr(ID_CALIBRACION) & " AND EQUIPO_ID = " & CStr(EQUIPO_ID)
    
    Set rs = datos_bd(strSql)
    
    If rs.RecordCount = 0 Then
        REACTIVOS = ""
        Set rs = Nothing
        Exit Sub
    End If
    
    
    rs.MoveFirst
    While Not rs.EOF
        REACTIVOS = REACTIVOS & rs("REACTIVO_ID") & ","
        rs.MoveNext
    Wend
    
    ' quita la última coma
    REACTIVOS = Left(REACTIVOS, Len(REACTIVOS) - 1)
    
    
End Sub

Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_CALIBRACION) FROM EQ_CALIBRACION_EQUIPOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_CALIBRACION = 1
        Else
                ID_CALIBRACION = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar(Optional ByVal prmActualizarFechaProxima As Boolean = False, Optional ByRef xres As XArrayDB = Nothing, Optional ByVal filas As Long = 0) As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        
        consulta = "INSERT INTO EQ_CALIBRACION_EQUIPOS (ID_CALIBRACION, RUTA_PLANTILLA, RUTA_CERTIFICADO, CUSERID) VALUES (" & CStr(ID_CALIBRACION) & ", '', '', " & CStr(USUARIO.getID_EMPLEADO) & ")"
        
        execute_bd consulta
        
        Insertar = ID_CALIBRACION
        
        Modificar ID_CALIBRACION, True, True, xres, filas
                
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEquipoCalibracion)", vbCritical, Err.Description
End Function
Public Function informarMuestraId(idCalibracion As Long, idMuestra As Long) As Long
        Dim consulta As String
   On Error GoTo informarMuestraId_Error

        consulta = "UPDATE EQ_CALIBRACION_EQUIPOS SET MUESTRA_ID=" & idMuestra & " WHERE ID_CALIBRACION = " & idCalibracion
        execute_bd consulta

   On Error GoTo 0
   Exit Function

informarMuestraId_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure informarMuestraId of Módulo de clase clsEquipoCalibracion"
End Function

Public Function Modificar(ID As Long, Optional ByVal prmActualizarFechaProxima As Boolean = False, Optional prmEs_Alta As Boolean = False, Optional ByRef xres As XArrayDB = Nothing, Optional ByVal filas As Long = 0) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    Dim strCert As String, strHoja As String, strEval As String
    
    consulta = "UPDATE EQ_CALIBRACION_EQUIPOS SET " & _
                    " EQUIPO_ID = " & EQUIPO_ID & "," & _
                    " TIPO_ID = " & TIPO_ID & "," & _
                    " PERIODICIDAD_ID = " & PERIODICIDAD_ID & "," & _
                    " PROCEDIMIENTO_ID = " & PROCEDIMIENTO_ID & "," & _
                    " CALIBRADOR_EXTERNO_ID = " & CALIBRADOR_EXTERNO_ID & "," & _
                    " CALIBRADOR_INTERNO_ID = " & CALIBRADOR_INTERNO_ID & "," & _
                    " FECHA_ACTUAL = '" & Format(CDate(fecha_actual), "yyyy/mm/dd") & "'," & _
                    " FECHA_PROXIMA = '" & Format(CDate(fecha_proxima), "yyyy/mm/dd") & "'," & _
                    " FECHA_PREVISTA = '" & Format(CDate(fecha_prevista), "yyyy/mm/dd") & "'," & _
                    " RANGO_MIN = " & moneda_bd(RANGO_MIN) & "," & _
                    " RANGO_MAX = " & moneda_bd(RANGO_MAX) & "," & _
                    " UNIDADES_ID = " & UNIDADES_ID & "," & _
                    " RUTA_PLANTILLA = '" & RUTA_PLANTILLA & "'," & _
                    " RUTA_CERTIFICADO = '" & RUTA_CERTIFICADO & "'," & _
                    " RUTA_EVALUACION = '" & RUTA_EVALUACION & "'," & _
                    " INCIDENCIAS = '" & INCIDENCIAS & "'," & _
                    " INCERTIDUMBRE = '" & INCERTIDUMBRE & "'," & _
                    " UBICACION_ID = " & UBICACION_ID & "," & _
                    " SEGREGADO = " & SEGREGADO & "," & " ETIQUETADO = " & ETIQUETADO & "," & " ESTADO = " & ESTADO & "," & _
                    " RESULTADO = " & RESULTADO & "," & _
                    " MUSERID = " & CStr(USUARIO.getID_EMPLEADO) & "," & _
                    " TS = LOCALTIMESTAMP" & _
                " WHERE ID_CALIBRACION = " & ID_CALIBRACION
                
        execute_bd consulta
        
        If prmActualizarFechaProxima And ESTADO > CAL_ESTADOS.CAL_ESTADO_PREVISTA Then
            consulta = "UPDATE EQUIPOS SET FECHA_PROX_CALIBRACION = '" & Format(CDate(fecha_proxima), "yyyy-mm-dd") & "' WHERE ID_EQUIPO = " & EQUIPO_ID & " AND FECHA_PROX_CALIBRACION < '" & Format(CDate(fecha_proxima), "yyyy-mm-dd") & "'"
            Call execute_bd(consulta)
            consulta = "UPDATE EQUIPOS SET FECHA_ULT_CALIBRACION = '" & Format(CDate(fecha_actual), "yyyy-mm-dd") & "' WHERE ID_EQUIPO = " & EQUIPO_ID & " AND FECHA_ULT_CALIBRACION < '" & Format(CDate(fecha_actual), "yyyy-mm-dd") & "'"
            Call execute_bd(consulta)
        End If
        If ESTADO = CAL_ESTADOS.CAL_ESTADO_REALIZADA Then
            Dim oEq As New clsEquipos
            oEq.PonerEquipoEnM2 CStr(EQUIPO_ID)
            Set oEq = Nothing
        End If
        Call ActualizarDatos_Accesorios
        Call ActualizarDatos_Reactivos
        
        ' Guarda los datos de Parámetros
        If filas <> 0 And Not xres Is Nothing Then
            GuardarParametrosCalibracion EQUIPO_ID, ID_CALIBRACION, xres, filas
        End If
        
        ' genera un evento siempre que se cierre la Calibracion (ya sea conforme o no conforme)
        If ESTADO = CAL_ESTADOS.CAL_ESTADO_REALIZADA Then ' Se ha cerrado la Calibracion
            Dim obj As New clsEquipos
            obj.setID_EQUIPO = EQUIPO_ID
            
            If RESULTADO = CAL_RESULTADOS.CAL_RESULTADO_CONFORME Then
                obj.generar_evento EQUIPO_ID, EVT_CALIBRACION_REALIZADA, EVTRZ_RESULTADO_OK, , ID
                crear_calibracion_pendiente ID, EQUIPO_ID
            Else ' CERRADO NO CONFORME
                If Not comprobar_anterior_cal_no_conforme Then
                    obj.generar_evento EQUIPO_ID, EVT_CALIBRACION_REALIZADA, EVTRZ_RESULTADO_NEG_1A_VEZ, , ID
                    crear_calibracion_pendiente ID, EQUIPO_ID
                Else
                    obj.generar_evento EQUIPO_ID, EVT_CALIBRACION_REALIZADA, EVTRZ_RESULTADO_NEG, , ID
'                    oOP.Eliminar ID_CALIBRACION, EQUIPO_ID, EQUIPOS_CVM_TIPOS.CVM_TIPO_CALIBRACION
                End If
            End If
            Set obj = Nothing
'        Else
            ' cuando el estado es 0, puede que haya cambiado algo (fecha prevista, responsable, periodicidad)
'            If Not prmEs_Alta Then
'                oOP.revisar_tarea ID_CALIBRACION, EQUIPO_ID, EQUIPOS_CVM_TIPOS.CVM_TIPO_CALIBRACION
'            End If
        End If
        Dim oOP As New clsEquiposOperacionesPendientes
        oOP.revisar_tarea ID_CALIBRACION, EQUIPO_ID, EQUIPOS_CVM_TIPOS.CVM_TIPO_CALIBRACION
        Set oOP = Nothing
        
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEquipoCalibracion)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        execute_bd "DELETE FROM eq_calibracion_parametros_resultados WHERE CALIBRACION_ID = " & ID
        execute_bd "DELETE FROM eq_calibracion_equipos_accesorios WHERE ID_CALIBRACION = " & ID
        execute_bd "DELETE FROM EQ_CALIBRACION_EQUIPOS WHERE ID_CALIBRACION = " & ID
        
        Dim oD As New clsDocumentacion
        oD.EliminarEquipoTipo EQUIPO_ID, 0, ID
        Set oD = Nothing
        
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquipoCalibracion)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EQ_CALIBRACION_EQUIPOS ORDER BY ID_CALIBRACION"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_CALIBRACION,EQUIPO_ID FROM EQ_CALIBRACION_EQUIPOS ORDER BY EQUIPO_ID"
    Set Listado_Combo = datos_bd(consulta)
End Function
Public Property Get AccesoriosEquipo() As clsGenericCollection
    Set AccesoriosEquipo = mvarobjAccesoriosEquipo
End Property

Public Property Set AccesoriosEquipo(objAccesoriosEquipo As clsGenericCollection)
    Set mvarobjAccesoriosEquipo = objAccesoriosEquipo
End Property

Public Property Get AccesoriosCalibracion() As clsGenericCollection

    Set AccesoriosCalibracion = mvarobjAccesoriosCalibracion

End Property

Public Property Set AccesoriosCalibracion(objAccesoriosCalibracion As clsGenericCollection)

    Set mvarobjAccesoriosCalibracion = objAccesoriosCalibracion

End Property

Public Sub Carga_AccesoriosCalibracion()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim objAcc As clsEquipoAccesorios
    
    consulta = "SELECT eq_calibracion_equipos_accesorios.*, EQUIPOS.NOMBRE AS NOMBRE_ACCESORIO, EQUIPOS.DESCRIPCION AS DESCRIPCION_ACCESORIO "
    consulta = consulta & "FROM eq_calibracion_equipos_accesorios "
    consulta = consulta & " INNER JOIN EQUIPOS ON  eq_calibracion_equipos_accesorios.ACCESORIO_ID = EQUIPOS.ID_EQUIPO "
    consulta = consulta & "WHERE eq_calibracion_equipos_accesorios.ID_CALIBRACION = " & ID_CALIBRACION
        
    Set rs = datos_bd(consulta)
    
    If rs.RecordCount = 0 Then Exit Sub

    rs.MoveFirst
    While Not rs.EOF
    
        Set objAcc = New clsEquipoAccesorios
        objAcc.setID_ACCESORIO = rs("ACCESORIO_ID")
        objAcc.setNOMBRE = rs("NOMBRE_ACCESORIO")
        Call mvarobjAccesoriosCalibracion.Add(objAcc, objAcc.getID_ACCESORIO)
        
        rs.MoveNext
    Wend
End Sub

Private Sub ActualizarDatos_Accesorios()
On Error GoTo ActualizarDatos_Accesorios_Error

Dim strSql() As String
Dim intCont As Integer
Dim objAcc As clsEquipoAccesorios
intCont = 0

    'coge las acciones inmediata y realiza las siguiente modificaciones.
    ' 1.- Elimina todas los registros con Id_aux =-1
    ' 2.- Hace Update de los registro con Id_aux = 1
    ' 3.- Inserta los registros con id_aux = -2

    If AccesoriosCalibracion.Count = 0 Then Exit Sub
    
    For Each objAcc In AccesoriosCalibracion.Iterator
        Select Case objAcc.getID_AUX
            Case Is = enumIdAux.ID_AUX_NUEVO
                intCont = intCont + 1
                ReDim Preserve strSql(intCont)
                
                strSql(intCont) = "INSERT INTO eq_calibracion_equipos_accesorios ("
                strSql(intCont) = strSql(intCont) & "equipo_id, "
                strSql(intCont) = strSql(intCont) & "id_calibracion, "
                strSql(intCont) = strSql(intCont) & "accesorio_id "
                strSql(intCont) = strSql(intCont) & ") "
                strSql(intCont) = strSql(intCont) & "SELECT " & CStr(EQUIPO_ID) & " as equipo_id,"
                strSql(intCont) = strSql(intCont) & CStr(ID_CALIBRACION) & " as id_calibracion, "
                strSql(intCont) = strSql(intCont) & CStr(objAcc.getID_ACCESORIO) & " as accesorio_id "
                                                          
            Case Is = enumIdAux.ID_AUX_MODIFICADO
                ' no va a pasar nunca
                
            Case Is = enumIdAux.ID_AUX_ELIMINADO
                intCont = intCont + 1
                ReDim Preserve strSql(intCont)
                strSql(intCont) = "DELETE FROM eq_calibracion_equipos_accesorios "
                strSql(intCont) = strSql(intCont) & " WHERE accesorio_id= " & CStr(objAcc.getID_ACCESORIO)
                strSql(intCont) = strSql(intCont) & " AND id_calibracion = " & CStr(ID_CALIBRACION)
        End Select
    Next objAcc
    
    For intCont = 1 To intCont
        execute_bd strSql(intCont)
    Next intCont

On Error GoTo 0
    Exit Sub
ActualizarDatos_Accesorios_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ActualizarDatos_Accesorios of Módulo de clase clsEquipos"
End Sub
Public Function DevolverParametrosResultados(ByVal prmId As String) As ADODB.Recordset
    Dim sql As String
    sql = "SELECT eq_calibracion_parametros_resultados.*, coalesce(unidades.nombre, 'N/A') as unidad from eq_calibracion_parametros_resultados left outer join unidades on unidades.id_unidad = eq_calibracion_parametros_resultados.unidad_id where eq_calibracion_parametros_resultados.calibracion_id = " & prmId
    Set DevolverParametrosResultados = datos_bd(sql)
End Function

Public Function GuardarParametrosCalibracion(ByVal prmid_equipo As Long, ByVal prmid_calibracion As Long, ByRef xres As XArrayDB, ByVal filas As Long) As Boolean
    Dim i As Long, sql As String, Id_resultado As String
    Dim blnCrearParametro As Boolean
    ' Elimina primero todos los paramtros guardados hasta el momento
    sql = "delete from eq_calibracion_parametros_resultados WHERE CALIBRACION_ID = " & CStr(prmid_calibracion) & " AND EQUIPO_ID = " & CStr(prmid_equipo)
    
    execute_bd sql

    For i = 0 To filas
        If Not IsEmpty(xres(i, 0)) Then ' mira se existe la descripción
            If Trim(xres(i, 0)) <> "" Then
                Id_resultado = "-1"
                If Not IsEmpty(xres(i, CONST_ID_RESULTADO)) Then
                    If xres(i, CONST_ID_RESULTADO) = "" Or xres(i, CONST_ID_RESULTADO) = "-1" Then ' Alta o modificacion de Registro
                        Id_resultado = "coalesce(max(id_resultado),0) + 1 as id_resultado FROM eq_calibracion_parametros_resultados"
                    Else
                        Id_resultado = xres(i, CONST_ID_RESULTADO)
                    End If
                Else
                    Id_resultado = "coalesce(max(id_resultado),0) + 1 as id_resultado FROM eq_calibracion_parametros_resultados"
                End If
                            
                sql = "INSERT INTO eq_calibracion_parametros_resultados (EQUIPO_ID, unidad_id, calibracion_id, descripcion, rango_min, rango_max, resultado, tolerancia_max, correccion, incertidumbre, ID_RESULTADO)"
                sql = sql & " SELECT " & CStr(prmid_equipo) & " AS ID_EQUIPO"
                
                If IsEmpty(xres(i, CONST_ID_UNIDAD)) Then
                    sql = sql & ", -1 AS UNIDAD_ID"
                ElseIf Trim(xres(i, CONST_ID_UNIDAD)) = "" Then
                    sql = sql & ", -1 AS UNIDAD_ID"
                Else
                    sql = sql & ", " & xres(i, CONST_ID_UNIDAD) & " AS UNIDAD_ID"
                End If
                
                sql = sql & ", " & CStr(prmid_calibracion) & " AS CALIBRACION_ID"
                sql = sql & ", '" & xres(i, CONST_DESCRIPCION) & "' AS DESCRIPCION"
                
                
                If IsEmpty(xres(i, CONST_RANGO_MIN)) Then
                    sql = sql & ", -1 AS RANGO_MIN"
                ElseIf Trim(xres(i, CONST_RANGO_MIN)) = "" Then
                    sql = sql & ", -1 AS RANGO_MIN"
                Else
                    sql = sql & ", " & Replace(xres(i, CONST_RANGO_MIN), ",", ".") & " AS RANGO_MIN"
                End If
                
                If IsEmpty(xres(i, CONST_RANGO_MAX)) Then
                    sql = sql & ", 0.000000 AS RANGO_MAX"
                ElseIf Trim(xres(i, CONST_RANGO_MAX)) = "" Then
                    sql = sql & ", 0.000000 AS RANGO_MAX"
                Else
                    sql = sql & ", " & Replace(xres(i, CONST_RANGO_MAX), ",", ".") & " AS RANGO_MAX"
                End If
                
                If IsEmpty(xres(i, CONST_RESULTADO_CAL)) Then
                    sql = sql & ", 0.000000 AS RESULTADO"
                ElseIf Trim(xres(i, CONST_RESULTADO_CAL)) = "" Then
                    sql = sql & ", 0.000000 AS RESULTADO"
                Else
                    sql = sql & ", " & Replace(xres(i, CONST_RESULTADO_CAL), ",", ".") & " AS RESULTADO"
                End If
                
                ' TOLERANCIA
                If IsEmpty(xres(i, CONST_TOLERANCIA)) Then
                    sql = sql & ", 0.000000 AS TOLERANCIA_MAX"
                ElseIf Trim(xres(i, CONST_TOLERANCIA)) = "" Then
                    sql = sql & ", 0.000000 AS TOLERANCIA_MAX"
                Else
                    sql = sql & ", " & Replace(xres(i, CONST_TOLERANCIA), ",", ".") & " AS TOLERANCIA_MAX"
                End If
                
                'CORRECCION
                If IsEmpty(xres(i, CONST_CORRECCION)) Then
                    sql = sql & ", 0.000000 AS CORRECCION"
                ElseIf Trim(xres(i, CONST_CORRECCION)) = "" Then
                    sql = sql & ", 0.000000 AS CORRECCION"
                Else
                    sql = sql & ", " & Replace(xres(i, CONST_CORRECCION), ",", ".") & " AS CORRECCION"
                End If
                
                ' INCERTIDUMBRE
                If IsEmpty(xres(i, CONST_INCERTIDUMBRE)) Then
                    sql = sql & ", 0.000000 AS INCERTIDUMBRE"
                ElseIf Trim(xres(i, CONST_INCERTIDUMBRE)) = "" Then
                    sql = sql & ", 0.000000 AS INCERTIDUMBRE"
                Else
                    sql = sql & ", " & Replace(xres(i, CONST_INCERTIDUMBRE), ",", ".") & " AS INCERTIDUMBRE"
                End If
                
                sql = sql & ", " & Id_resultado
                            
                execute_bd sql
            End If
        End If
        
        
        
    Next i


End Function

'Public Property Get Evaluacion() As clsArchivoAdjunto
'    Set Evaluacion = mvarobjEvaluacion
'End Property
'Public Property Set Evaluacion(objEvaluacion As clsArchivoAdjunto)
'    Set mvarobjEvaluacion = objEvaluacion
'End Property

Private Sub ActualizarDatos_Reactivos()

    Dim arrRex() As String, x As Integer, sel As String

    sel = "delete from eq_calibracion_reactivos where equipo_id = " & CStr(EQUIPO_ID) & " AND calibracion_id = " & CStr(ID_CALIBRACION)
    execute_bd sel
    

    If Trim(REACTIVOS) <> "" Then
        arrRex = Split(REACTIVOS, ",")
    Else
        Exit Sub
    End If

    For x = 0 To UBound(arrRex)
        sel = "insert into eq_calibracion_reactivos (equipo_id, calibracion_id, reactivo_id) VALUES ("
        sel = sel & CStr(EQUIPO_ID) & ",  " & CStr(ID_CALIBRACION) & ",  " & CStr(arrRex(x))
        sel = sel & ")"
        execute_bd sel
    Next x
End Sub

Public Sub imprimir_etiqueta(CALIBRACION_ID As Long)
    Dim prnPrinter As Printer
    ' se mira si el equipo tiene impresora de etiquetas
    Dim oParametro As New clsParametros

   On Error GoTo imprimir_etiqueta_Error

    log ("Impresion etiqueta calibracion")
    Dim oEC As New clsEquipoCalibracion
    oEC.Carga CALIBRACION_ID
    Dim PC As String
'    If EQUIPO <> "" Then
'        PC = EQUIPO
'    Else
        PC = USUARIO.getUSO
'    End If
    log ("PC : " & PC)
    If PC = "IBERIA" Then
        Dim oEtiqueta As New clsEtiquetas
        With oEtiqueta
           .setTIPO_ID = C_ETIQUETAS_TIPOS.C_ETIQUETAS_TIPOS_EQ_CAL
           .setCENTRO_ID = CENTROS.CENTRO_MADRID
           .setID = CALIBRACION_ID
'           .setUSUARIO_ID = oEC.getCALIBRADOR_INTERNO_ID
           .setUSUARIO_ID = oEC.getCUSERID
           .Insertar
        End With
        MsgBox "Se han enviado a imprimir las copias.", vbInformation, App.Title
        Exit Sub
    End If
    
    If Not oParametro.Carga(parametros.IMPRESORA_ETIQUETAS_PEQUENA, PC) Then
        MsgBox "Este equipo no tiene asignada impresora de etiquetas.", vbCritical, App.Title
        Exit Sub
    End If
    log ("Parametro impresora : " & Replace(oParametro.getVALOR, "/", "\"))
    
    Dim impresora_encontrada As String
    Dim IMPRESORA As Printer
    impresora_encontrada = ""
    For Each prnPrinter In Printers
'        log "Impresora instalada : " & prnPrinter.DeviceName
        ' Quitamos del nombre de la impresora el #00X para las impresoras TSN del remoto
        Dim impresoraArray() As String
        Dim impresoraNombre As String
        impresoraArray = Split(UCase(prnPrinter.DeviceName), "#")
        impresoraNombre = Trim(impresoraArray(0))
        
'        If UCase(prnPrinter.DeviceName) = Replace(UCase(oParametro.getVALOR), "/", "\") Then
        If impresoraNombre = Replace(UCase(oParametro.getVALOR), "/", "\") Then
            Set Printer = prnPrinter
            Set IMPRESORA = prnPrinter
            impresora_encontrada = prnPrinter.DeviceName
            Exit For
        End If
    Next
    If impresora_encontrada <> "" Then
        Dim imp_ant As String
        imp_ant = Impresora_Predeterminada
        Establecer_Impresora impresora_encontrada
        
        Dim P1() As String
        Dim P2() As String
        ReDim P1(1) As String
        ReDim P2(1) As String
        P1(1) = "FIRMA"
        Dim oUsuario As New clsUsuarios
'        Dim oEC As New clsEquipoCalibracion
'        oEC.Carga CALIBRACION_ID
        oUsuario.CARGAR oEC.getCUSERID
        If ReadINI(App.Path + "\config.ini", "Documentos", "firmas") = "c:\geslab\recursos\firmas" Then
            Dim firma As String
            Dim Lista() As String
            If oUsuario.getFIRMA <> "" Then
                Lista = Split(oUsuario.getFIRMA, "/")
                firma = ReadINI(App.Path + "\config.ini", "Documentos", "firmas") & "\" & Lista(UBound(Lista))
            End If
            P2(1) = firma
        Else
            P2(1) = Replace(oUsuario.getFIRMA, "/", "\")
        End If
        
        With frmReport
            .iniciar
            .informe = "\Equipos\rptEquipos_ETIQUETA_CalibracionPEQ"
            .criterio = "{eq_calibracion_equipos.ID_CALIBRACION} =" & CALIBRACION_ID
            .ParametrosNombre = P1
            .ParametrosValores = P2
            .imprimir = True
            .generar
        End With
        Set frmReport = Nothing
        Establecer_Impresora imp_ant
    Else
        MsgBox "No se localiza la impresora de etiquetas.", vbExclamation, App.Title
    End If

   On Error GoTo 0
   Exit Sub

imprimir_etiqueta_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imprimir_etiqueta of Módulo de clase clsEquipoCalibracion"
    
End Sub
Public Sub Reabrir(ByVal prmId As String)
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE EQ_CALIBRACION_EQUIPOS SET " & _
                    " ESTADO = 0," & _
                    " RESULTADO = 0," & _
                    " MUSERID = " & CStr(USUARIO.getID_EMPLEADO) & "," & _
                    " TS = CURRENT_TIMESTAMP" & _
            " WHERE ID_CALIBRACION = " & prmId
            
    execute_bd consulta
    Exit Sub
fallo:
    MsgBox "Error al reabrir (clsEquipoCalibracion)", vbCritical, Err.Description
End Sub


