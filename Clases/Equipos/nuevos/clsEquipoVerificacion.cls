VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipoVerificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQ_VERIFICACION_EQUIPOS
'***************************************************************
Private ID_VERIFICACION As Long
Private EQUIPO_ID As Long
Private TIPO_ID As Long
Private PERIODICIDAD_ID As Long
Private PROCEDIMIENTO_ID As Long
Private VERIFICADOR_EXTERNO_ID As Long
Private VERIFICADOR_INTERNO_ID As Long
Private fecha_actual As String
Private fecha_proxima As String
Private fecha_prevista As String
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
Private mvarobjLimitacionesUso As New clsGenericCollection
Private responsable As String
Private PROCEDIMIENTO As String
Private UBICACION_ID As Long
Private INCIDENCIAS As String
Private BUSQUEDAS As Integer
Private ENVIADO As Integer
Private EDICION As Integer
Private INCERTIDUMBRE_USO As String

Private Const CONST_DESCRIPCION As Integer = 0
Private Const CONST_RANGO_MIN As Integer = 1
Private Const CONST_RANGO_MAX As Integer = 2
Private Const CONST_unidad As Integer = 3
Private Const CONST_RESULTADO_VER As Integer = 4
Private Const CONST_TOLERANCIA As Integer = 5
Private Const CONST_INCERTIDUMBRE As Integer = 6
Private Const CONST_CORRECCION As Integer = 7
Private Const CONST_ID_RESULTADO As Integer = 8
Private Const CONST_ID_UNIDAD As Integer = 9

Private Const COL_Descripcion As Integer = 0
Private Const COL_Unidad As Integer = 1
Private Const COL_RANGO_MIN As Integer = 2
Private Const COL_RANGO_MAX As Integer = 3
Private Const COL_RESULTADO_MEDIA As Integer = 4
Private Const COL_id_tipo As Integer = 5
Private Const COL_Id_resultado As Integer = 6
Private Const COL_RESULTADO_CUALIDAD As Integer = 7
Private Const COL_RESULTADOS_PATRON As Integer = 8
Private Const COL_id_unidad As Integer = 9
Private Const COL_id_patron As Integer = 10
Private Const COL_n_medidas As Integer = 11
Private Const COL_LEQUIPOS As Integer = 12
Private Const COL_LREACTIVOS As Integer = 13
Private Const COL_LREACTIVOS_PROPIOS As Integer = 14

Public Property Let setPROCEDIMIENTO(ByVal dato As String)
    PROCEDIMIENTO = dato
End Property
Public Property Get getPROCEDIMIENTO() As String
    getPROCEDIMIENTO = PROCEDIMIENTO
End Property
Public Property Let setUBICACION_ID(ByVal dato As Long)
    UBICACION_ID = dato
End Property
Public Property Get getUBICACION_ID() As Long
    getUBICACION_ID = UBICACION_ID
End Property
Public Property Let setINCIDENCIAS(ByVal dato As String)
    INCIDENCIAS = dato
End Property
Public Property Get getINCIDENCIAS() As String
    getINCIDENCIAS = INCIDENCIAS
End Property
Public Property Let setBUSQUEDAS(ByVal dato As Integer)
    BUSQUEDAS = dato
End Property
Public Property Get getBUSQUEDAS() As Integer
    getBUSQUEDAS = BUSQUEDAS
End Property
Public Property Let setENVIADO(ByVal dato As Integer)
    ENVIADO = dato
End Property
Public Property Get getENVIADO() As Integer
    getENVIADO = ENVIADO
End Property
Public Property Let setEDICION(ByVal dato As Integer)
    EDICION = dato
End Property
Public Property Get getEDICION() As Integer
    getEDICION = EDICION
End Property
Public Property Let setINCERTIDUMBRE_USO(ByVal dato As String)
    INCERTIDUMBRE_USO = dato
End Property
Public Property Get getINCERTIDUMBRE_USO() As String
    getINCERTIDUMBRE_USO = INCERTIDUMBRE_USO
End Property

Private Sub Class_Initialize()
    mvarobjLimitacionesUso.KeyName = "setID_LIMITACION_USO"
End Sub

Private Sub Class_Terminate()
    Set mvarobjLimitacionesUso = Nothing
End Sub
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQ_VERIFICACION_EQUIPOS
'***************************************************************
Public Property Let setID_VERIFICACION(ByVal dato As Long)
        ID_VERIFICACION = dato
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
Public Property Let setVERIFICADOR_EXTERNO_ID(ByVal dato As Long)
        VERIFICADOR_EXTERNO_ID = dato
End Property
Public Property Let setVERIFICADOR_INTERNO_ID(ByVal dato As Long)
        VERIFICADOR_INTERNO_ID = dato
End Property
Public Property Let setFECHA_ACTUAL(ByVal dato As String)
        fecha_actual = dato
End Property
Public Property Let setFECHA_PROXIMA(ByVal dato As String)
        fecha_proxima = dato
End Property
Public Property Let setFECHA_PREVISTA(ByVal dato As String)
        fecha_prevista = dato
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
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQ_VERIFICACION_EQUIPOS
'***************************************************************
Public Property Get getID_VERIFICACION() As Long
        getID_VERIFICACION = ID_VERIFICACION
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
Public Property Get getVERIFICADOR_EXTERNO_ID() As Long
        getVERIFICADOR_EXTERNO_ID = VERIFICADOR_EXTERNO_ID
End Property
Public Property Get getVERIFICADOR_INTERNO_ID() As Long
        getVERIFICADOR_INTERNO_ID = VERIFICADOR_INTERNO_ID
End Property
Public Property Get getFECHA_ACTUAL() As String
        getFECHA_ACTUAL = fecha_actual
End Property
Public Property Get getFECHA_PROXIMA() As String
        getFECHA_PROXIMA = fecha_proxima
End Property
Public Property Get getFECHA_PREVISTA() As String
        getFECHA_PREVISTA = fecha_prevista
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
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQ_VERIFICACION_EQUIPOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = " SELECT t1.*, COALESCE(ca_documentos.NOMBRE, '') as PROCEDIMIENTO FROM (select eq_verificacion_equipos.*, proveedores.NOMBRE as RESPONSABLE"
        consulta = consulta & " From eq_verificacion_equipos"
        consulta = consulta & " inner join proveedores on eq_verificacion_equipos.VERIFICADOR_EXTERNO_ID = proveedores.ID_PROVEEDOR"
        consulta = consulta & " where eq_verificacion_equipos.tipo_id = 2 and eq_verificacion_equipos.ID_VERIFICACION = " & ID
        consulta = consulta & " Union"
        consulta = consulta & " select eq_verificacion_equipos.*, concat(usuarios.NOMBRE, ' ', usuarios.APELLIDOS) as RESPONSABLE"
        consulta = consulta & " From eq_verificacion_equipos"
        consulta = consulta & " inner join usuarios on eq_verificacion_equipos.VERIFICADOR_INTERNO_ID = usuarios.ID_EMPLEADO"
        consulta = consulta & " where eq_verificacion_equipos.tipo_id = 1 and eq_verificacion_equipos.ID_VERIFICACION = " & ID
        consulta = consulta & " ) as t1 "
        consulta = consulta & " LEFT join ca_documentos ON T1.PROCEDIMIENTO_ID = CA_DOCUMENTOS.ID_DOCUMENTO"
        consulta = consulta & " WHERE t1.ID_VERIFICACION = " & ID
        
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_VERIFICACION = rs("ID_VERIFICACION")
                EQUIPO_ID = rs("EQUIPO_ID")
                TIPO_ID = rs("TIPO_ID")
                PERIODICIDAD_ID = rs("PERIODICIDAD_ID")
                PROCEDIMIENTO_ID = rs("PROCEDIMIENTO_ID")
                PROCEDIMIENTO = rs("PROCEDIMIENTO")
                VERIFICADOR_EXTERNO_ID = rs("VERIFICADOR_EXTERNO_ID")
                VERIFICADOR_INTERNO_ID = rs("VERIFICADOR_INTERNO_ID")
                fecha_actual = rs("FECHA_ACTUAL")
                fecha_proxima = rs("FECHA_PROXIMA")
                fecha_prevista = rs("FECHA_PREVISTA")
                RANGO_MIN = rs("RANGO_MIN")
                RANGO_MAX = rs("RANGO_MAX")
                responsable = rs("RESPONSABLE")
                UNIDADES_ID = rs("UNIDADES_ID")
                RUTA_PLANTILLA = rs("RUTA_PLANTILLA")
                RUTA_CERTIFICADO = rs("RUTA_CERTIFICADO")
                RUTA_EVALUACION = rs("RUTA_EVALUACION")
                ESTADO = rs("ESTADO")
                RESULTADO = rs("RESULTADO")
                CUSERID = rs("CUSERID")
                MUSERID = rs("MUSERID")
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
                
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquipoVerificacion)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_VERIFICACION) FROM EQ_VERIFICACION_EQUIPOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_VERIFICACION = 1
        Else
                ID_VERIFICACION = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar(Optional ByVal prmActualizarFechaProxima As Boolean = False, Optional ByRef lista_parametros As ListView = Nothing) As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        
        consulta = "INSERT INTO EQ_VERIFICACION_EQUIPOS(ID_VERIFICACION, RUTA_PLANTILLA, RUTA_CERTIFICADO, CUSERID) VALUES (" & CStr(ID_VERIFICACION) & ", '', '', " & CStr(USUARIO.getID_EMPLEADO) & ")"
        
        execute_bd consulta
        
        Insertar = ID_VERIFICACION
        
        'GuardarParametrosVerificacion
        
        Call Modificar(ID_VERIFICACION, prmActualizarFechaProxima, True, lista_parametros)
        
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEquipoVerificacion)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long, Optional prmActualizarFechaProxima As Boolean = False, Optional ByVal prmEs_Alta As Boolean = False, Optional ByRef lista_parametros As ListView = Nothing) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    
    consulta = "UPDATE EQ_VERIFICACION_EQUIPOS SET " & _
                    " EQUIPO_ID = " & EQUIPO_ID & "," & _
                    " TIPO_ID = " & TIPO_ID & "," & _
                    " PERIODICIDAD_ID = " & PERIODICIDAD_ID & "," & _
                    " PROCEDIMIENTO_ID = " & PROCEDIMIENTO_ID & "," & _
                    " UBICACION_ID = " & UBICACION_ID & "," & _
                    " INCIDENCIAS = '" & INCIDENCIAS & "'," & _
                    " VERIFICADOR_EXTERNO_ID = " & VERIFICADOR_EXTERNO_ID & "," & _
                    " VERIFICADOR_INTERNO_ID = " & VERIFICADOR_INTERNO_ID & "," & _
                    " FECHA_ACTUAL = '" & Format(CDate(fecha_actual), "yyyy/mm/dd") & "'," & _
                    " FECHA_PROXIMA = '" & Format(CDate(fecha_proxima), "yyyy/mm/dd") & "'," & _
                    " FECHA_PREVISTA = '" & Format(CDate(fecha_prevista), "yyyy/mm/dd") & "'," & _
                    " RANGO_MIN = " & moneda_bd(RANGO_MIN) & "," & _
                    " RANGO_MAX = " & moneda_bd(RANGO_MAX) & "," & _
                    " UNIDADES_ID = " & UNIDADES_ID & "," & _
                    " RUTA_PLANTILLA = '" & RUTA_PLANTILLA & "'," & " RUTA_CERTIFICADO = '" & RUTA_CERTIFICADO & "'," & " RUTA_EVALUACION = '" & RUTA_EVALUACION & "'," & _
                    " ESTADO = " & ESTADO & "," & _
                    " RESULTADO = " & RESULTADO & "," & _
                    " MUSERID = " & CStr(USUARIO.getID_EMPLEADO) & "," & _
                    " TS = LOCALTIMESTAMP" & _
            " WHERE ID_VERIFICACION = " & ID

        execute_bd consulta
        If PERIODICIDAD_ID <> ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_ENSAYO And _
           PERIODICIDAD_ID <> ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_CADA_ENSAYO Then
            ' Actualiza la fecha Proxima en el equipo
            If prmActualizarFechaProxima And ESTADO > VER_ESTADOS.VER_ESTADO_PREVISTA Then
                consulta = "UPDATE EQUIPOS SET FECHA_ULT_VERIFICACION = '" & Format(CDate(fecha_actual), "yyyy-mm-dd") & "' WHERE ID_EQUIPO = " & EQUIPO_ID & " AND FECHA_ULT_VERIFICACION < '" & Format(CDate(fecha_actual), "yyyy-mm-dd") & "'"
                consulta = "UPDATE EQUIPOS SET FECHA_PROX_VERIFICACION = '" & Format(CDate(fecha_proxima), "yyyy-mm-dd") & "' WHERE ID_EQUIPO = " & EQUIPO_ID & " AND FECHA_PROX_VERIFICACION < '" & Format(CDate(fecha_proxima), "yyyy-mm-dd") & "'"
                Call execute_bd(consulta)
            End If
        End If
        ' Guarda los parámetros
        If Not lista_parametros Is Nothing Then
            GuardarParametrosVerificacion EQUIPO_ID, ID_VERIFICACION, lista_parametros
        End If
        
        'genera un evento siempre que se cierre el Verificacion (ya sea conforme o no conforme)
        If ESTADO = VER_ESTADOS.VER_ESTADO_REALIZADA Then ' Se ha cerrado la VERIFICACION
            Dim obj As New clsEquipos
            obj.setID_EQUIPO = EQUIPO_ID
            
            If RESULTADO = VER_RESULTADOS.VER_RESULTADO_CONFORME Then
                obj.generar_evento EQUIPO_ID, EVT_VERIFICACION_REALIZADA, EVTRZ_RESULTADO_OK, , ID
                If PERIODICIDAD_ID <> ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_ENSAYO And _
                   PERIODICIDAD_ID <> ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_CADA_ENSAYO Then
                    crear_verificacion_pendiente ID, EQUIPO_ID, False
                End If
            Else ' CERRADO NO CONFORME
                If Not comprobar_anterior_ver_no_conforme Then
                    obj.generar_evento EQUIPO_ID, EVT_VERIFICACION_REALIZADA, EVTRZ_RESULTADO_NEG_1A_VEZ, , ID
                    If PERIODICIDAD_ID <> ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_ENSAYO And _
                       PERIODICIDAD_ID <> ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_CADA_ENSAYO Then
                        crear_verificacion_pendiente ID, EQUIPO_ID, False
                    End If
                Else
                    obj.generar_evento EQUIPO_ID, EVT_VERIFICACION_REALIZADA, EVTRZ_RESULTADO_NEG, , ID
                    ' elimina la operacion pendiente directamente
                    
'                    oOP.Eliminar ID_VERIFICACION, EQUIPO_ID, EQUIPOS_CVM_TIPOS.CVM_TIPO_VERIFICACION
'                    Set oOP = Nothing
                End If
            End If
            
        Else
            ' cuando el estado es 0, puede que haya cambiado algo (fecha prevista, responsable, periodicidad)
'OOP            If Not prmEs_Alta Then
'                oOP.revisar_tarea ID_VERIFICACION, EQUIPO_ID, EQUIPOS_CVM_TIPOS.CVM_TIPO_VERIFICACION
'OOP            End If
        End If
        
        Dim oOP As New clsEquiposOperacionesPendientes
        oOP.revisar_tarea ID_VERIFICACION, EQUIPO_ID, EQUIPOS_CVM_TIPOS.CVM_TIPO_VERIFICACION
        Set oOP = Nothing
        
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEquipoVerificacion)", vbCritical, Err.Description
End Function
Private Function comprobar_anterior_ver_no_conforme() As Boolean
Dim rs As ADODB.Recordset

    comprobar_anterior_ver_no_conforme = False

    Set rs = datos_bd("select * from eq_verificacion_equipos where equipo_id = " & EQUIPO_ID & " AND id_verificacion <> " & ID_VERIFICACION & " order by id_verificacion desc")
            
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        If CInt(rs!RESULTADO) = VER_RESULTADOS.VER_RESULTADO_NO_CONFORME Then
            comprobar_anterior_ver_no_conforme = True
            Set rs = Nothing
        End If
    End If
    
    Set rs = Nothing
    
    
End Function
Public Function buscar_verificacion_anterior_misma_periodicidad(EQUIPO As Long, PERIODICIDAD As String) As Long
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim filtro As String
    If PERIODICIDAD <> "" Then
        filtro = " AND PERIODICIDAD_ID IN (" & PERIODICIDAD & ")"
    End If
    consulta = "select id_verificacion from eq_verificacion_equipos " & _
               " where equipo_id = " & EQUIPO & _
               filtro & _
               " order by id_verificacion desc"
    
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        buscar_verificacion_anterior_misma_periodicidad = rs(0)
    Else
        buscar_verificacion_anterior_misma_periodicidad = 0
    End If
    
    Set rs = Nothing
    
    
End Function
'MANTIS-810-F

Public Function crear_verificacion_pendiente(ByVal prm_id_ver As Long, ByVal prm_id_equipo As Long, hoy As Boolean) As Boolean

    '1 crea una verificacion a partir de esta
    '2 la guarda
    '3 crea el registro de operacion pendiente
        
    Dim over As New clsEquipoVerificacion
    Dim oPer As New clsEquiposPeriodicidad
    Dim oOP As clsEquiposOperacionesPendientes
    Dim rs As ADODB.Recordset, consulta As String
    
   On Error GoTo crear_verificacion_pendiente_Error

    over.Carga prm_id_ver
    
    ' cambia las propiedades
    
    ' calcula la fecha en que se va a realizar la operacion
    ' cuando el estado es cerrado NO CONFORME ,
    If ESTADO = VER_ESTADOS.VER_ESTADO_REALIZADA Then
        If hoy Then
            over.setFECHA_ACTUAL = Date
        Else
            over.setFECHA_ACTUAL = oPer.calcular_fecha(over.getFECHA_ACTUAL, over.getPERIODICIDAD_ID)
        End If
    End If
    
    ' una vez conocida la fecha Prevista, estima la fecha de la siguiente para cuando se cierre
    over.setFECHA_PREVISTA = over.getFECHA_ACTUAL
    over.setFECHA_PROXIMA = oPer.calcular_fecha(over.getFECHA_ACTUAL, over.getPERIODICIDAD_ID)
    over.setESTADO = VER_ESTADOS.VER_ESTADO_PREVISTA
    over.setRESULTADO = VER_RESULTADOS.VER_RESULTADO_CONFORME
    'M0496-I
    ' Si las fechas para la nueva verificación son las mismas no genera la verificación
    oPer.Carga over.getPERIODICIDAD_ID
    If oPer.getMESES <> 0 Or oPer.getDIAS <> 0 Then
        If over.getFECHA_PREVISTA = over.getFECHA_PROXIMA Then
            MsgBox "No se puede crear la verificación, la fecha previsa y la proxima es la misma.", vbCritical, App.Title
            Exit Function
        End If
    End If
    'M0496-F
    
    ' Le quita los archivos adjuntos
    over.setRUTA_CERTIFICADO = ""
    over.setRUTA_EVALUACION = ""
    over.setRUTA_PLANTILLA = ""
    over.setBUSQUEDAS = 0
    over.setENVIADO = 0
    over.setEDICION = 0
    over.setINCERTIDUMBRE_USO = 0
    ' Los Parametros no se insertan automáticamente, con lo que habrá que hacerlo posteriormente
    over.Insertar False
    'Una vez insertado el clon previsto, se insertan Parámetros con Resultados a 0
    ' Parámetros
    Set rs = datos_bd("select * from eq_verificacion_parametros_resultados where verificacion_id = " & prm_id_ver)
    
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        While Not rs.EOF
        
            consulta = "insert into eq_verificacion_parametros_resultados (equipo_id, realizado, id_resultado, verificacion_id, unidad_id, "
            consulta = consulta & "rango_min, rango_max, resultado, tolerancia_max, correccion, incertidumbre, descripcion, tipo_id, N_MEDIDAS, RESULTADO_CUALIDAD, PATRON_ID,DATOS_PARAMETRO,LEQUIPOS,LREACTIVOS,LREACTIVOS_PROPIOS) "
            consulta = consulta & " SELECT "
            consulta = consulta & prm_id_equipo & " as equipo_id, "
            consulta = consulta & " 1 as realizado, "
            consulta = consulta & " cast(coalesce(max(id_resultado),0)+ 1 as unsigned) as id_resultado, "
            consulta = consulta & over.getID_VERIFICACION & " as verificacion_id, "
            consulta = consulta & rs!UNIDAD_ID & " as unidad_id , "
            consulta = consulta & Replace(Format(rs!RANGO_MIN, "0.000000"), ",", ".") & " as rango_min , "
            consulta = consulta & Replace(Format(rs!RANGO_MAX, "0.000000"), ",", ".") & " as rango_max , "
            consulta = consulta & " 0.000000 as resultado , "
            consulta = consulta & Replace(Format(rs!TOLERANCIA_MAX, "0.000000"), ",", ".") & " as tolerancia_max , "
            consulta = consulta & Replace(Format(rs!CORRECCION, "0.000000"), ",", ".") & " as correccion , "
            consulta = consulta & Replace(Format(rs!INCERTIDUMBRE, "0.000000"), ",", ".") & " as incertidumbre , "
            consulta = consulta & "'" & rs!DESCRIPCION & "' as descripcion, "
            'tipo_id, N_MEDIDAS, RESULTADO_CUALIDAD, PATRON_ID,DATOS_PARAMETRO
            consulta = consulta & rs!TIPO_ID & " as tipo_id, "
            consulta = consulta & rs!n_medidas & " as N_MEDIDAS, "
            consulta = consulta & rs!RESULTADO_CUALIDAD & " as RESULTADO_CUALIDAD, "
            consulta = consulta & rs!patron_id & " as PATRON_ID, "
            If rs!TIPO_ID = 0 Then
                consulta = consulta & "'" & rs!DATOS_PARAMETRO & "' as DATOS_PARAMETRO "
            Else
                consulta = consulta & "'0' as DATOS_PARAMETRO "
            End If
            consulta = consulta & ", '" & rs!LEQUIPOS & "'"
            consulta = consulta & ", '" & rs!lReactivos & "'"
            consulta = consulta & ", '" & rs!LREACTIVOS_PROPIOS & "'"
            
            consulta = consulta & " from eq_verificacion_parametros_resultados"
            
            execute_bd consulta
            rs.MoveNext
        Wend
    End If
    Set oOP = New clsEquiposOperacionesPendientes
    oOP.crear_verificacion_pendiente prm_id_ver, over.getID_VERIFICACION, prm_id_equipo, oPer
    
    
    Set over = Nothing
    Set oPer = Nothing
    Set oOP = Nothing
    Set rs = Nothing
    crear_verificacion_pendiente = True

   On Error GoTo 0
   Exit Function

crear_verificacion_pendiente_Error:
    crear_verificacion_pendiente = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure crear_verificacion_pendiente of Módulo de clase clsEquipoVerificacion"
    
End Function


Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        execute_bd "DELETE FROM eq_verificacion_parametros_resultados WHERE VERIFICACION_ID = " & ID
        execute_bd "DELETE FROM EQ_VERIFICACION_EQUIPOS WHERE ID_VERIFICACION = " & ID
        
        Dim oD As New clsDocumentacion
        oD.EliminarEquipoTipo EQUIPO_ID, 1, ID
        Set oD = Nothing
        
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquipoVerificacion)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM EQ_VERIFICACION_EQUIPOS ORDER BY ID_VERIFICACION"
        Set Listado = datos_bd(consulta)
End Function
Public Function pendienteVerificacion(ID_EQUIPO As Long, fecha As String) As String
        Dim consulta As String
        Dim rs As ADODB.Recordset
        Dim salida As String
        consulta = "SELECT FECHA_ACTUAL FROM EQ_VERIFICACION_EQUIPOS WHERE EQUIPO_ID = " & ID_EQUIPO & " AND ESTADO = 0 AND FECHA_ACTUAL <= '" & Format(fecha, "yyyy-mm-dd") & "' ORDER BY ID_VERIFICACION LIMIT 1"
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            salida = "El equipo " & ID_EQUIPO & " tiene verificación pendiente con fecha " & Format(rs("FECHA_ACTUAL"), "dd/mm/yyyy")
        End If
        pendienteVerificacion = salida
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_VERIFICACION,EQUIPO_ID FROM EQ_VERIFICACION_EQUIPOS ORDER BY EQUIPO_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function

Public Property Get LimitacionesUso() As clsGenericCollection

    Set LimitacionesUso = mvarobjLimitacionesUso

End Property

Public Property Set LimitacionesUso(objLimitacionesUso As clsGenericCollection)

    Set mvarobjLimitacionesUso = objLimitacionesUso

End Property

Public Property Let setRESPONSABLE(ByVal strRESPONSABLE As String)

    responsable = strRESPONSABLE

End Property

Public Property Get getRESPONSABLE() As String

    getRESPONSABLE = responsable

End Property

Public Function DevolverParametrosResultados(ByVal prmId As String) As ADODB.Recordset
    Dim sql As String
    sql = "SELECT eq_verificacion_parametros_resultados.*, coalesce(unidades.nombre, 'N/A') as unidad from eq_verificacion_parametros_resultados left outer join unidades on unidades.id_unidad = eq_verificacion_parametros_resultados.unidad_id where eq_verificacion_parametros_resultados.verificacion_id = " & prmId
    Set DevolverParametrosResultados = datos_bd(sql)
End Function

Public Function GuardarParametrosVerificacion(ByVal prmid_equipo As Long, prmid_verificacion As Long, ByVal lista_parametros As ListView) As Boolean
    Dim i As Long, sql As String, Id_resultado As String
    Dim blnCrearParametro As Boolean
    Dim filas As Long
    filas = lista_parametros.ListItems.Count
    ' Elimina primero todos los paramtros guardados hasta el momento
    sql = "delete from eq_verificacion_parametros_resultados WHERE VERIFICACION_ID = " & CStr(prmid_verificacion) & " AND EQUIPO_ID = " & CStr(prmid_equipo)
    execute_bd sql
    
    
    ' si las ha eliminado todas, fuera.
    If filas = 0 Then Exit Function

    For i = 1 To filas
        Id_resultado = "-1"
        With lista_parametros.ListItems(i)
        
            If Trim(.SubItems(COL_Id_resultado)) = "" Or Trim(.SubItems(COL_Id_resultado)) = "0" Then
                Id_resultado = "coalesce(max(id_resultado),0) + 1 as id_resultado FROM eq_verificacion_parametros_resultados"
            Else
                Id_resultado = .SubItems(COL_Id_resultado)
            End If
            
            ' JONATHAN 2010.09.07 NUEVOS CAMPOS EN PARAMETROS
            ' TIPO_ID, N_MEDIDAS, RESULTADO_CUALIDAD, PATRON_ID, DATOS_PARAMETRO
            
            sql = "INSERT INTO eq_verificacion_parametros_resultados (EQUIPO_ID, realizado, unidad_id, verificacion_id, descripcion, rango_min, rango_max, resultado, tolerancia_max, correccion, incertidumbre, TIPO_ID, N_MEDIDAS, RESULTADO_CUALIDAD, PATRON_ID, DATOS_PARAMETRO, LEQUIPOS, LREACTIVOS, LREACTIVOS_PROPIOS,ID_RESULTADO)"
            sql = sql & " SELECT " & CStr(prmid_equipo) & " AS ID_EQUIPO"
            ' si se realiza o no
            sql = sql & ", " & IIf(.Checked, "1", "0") & " AS REALIZADO"
            
            If Trim(.SubItems(COL_id_unidad)) = "" Or Trim(.SubItems(COL_id_unidad)) = "N/A" Then
                sql = sql & ",0 AS UNIDAD_ID"
            Else
                sql = sql & ", " & .SubItems(COL_id_unidad) & " AS UNIDAD_ID"
            End If
                
            sql = sql & ", " & CStr(prmid_verificacion) & " AS VERIFICACION_ID"
            sql = sql & ", '" & .Text & "' AS DESCRIPCION"
                                
            If Trim(.SubItems(COL_RANGO_MIN)) = "" Or Trim(.SubItems(COL_RANGO_MIN)) = "N/A" Then
                sql = sql & ",0 AS RANGO_MIN"
            Else
                sql = sql & ", " & Replace(.SubItems(COL_RANGO_MIN), ",", ".") & " AS RANGO_MIN"
            End If
            
            If Trim(.SubItems(COL_RANGO_MAX)) = "" Or Trim(.SubItems(COL_RANGO_MAX)) = "N/A" Then
                sql = sql & ",0 AS RANGO_MAX"
            Else
                sql = sql & ", " & Replace(.SubItems(COL_RANGO_MAX), ",", ".") & " AS RANGO_MAX"
            End If
                
            If Trim(.SubItems(COL_RESULTADO_MEDIA)) = "" Then
                sql = sql & ", 0.000000 AS RESULTADO"
            Else
                If Not IsNumeric(.SubItems(COL_RESULTADO_MEDIA)) Then
                    sql = sql & ", 0.000000 AS RESULTADO"
                Else
                    sql = sql & ", " & Replace(.SubItems(COL_RESULTADO_MEDIA), ",", ".") & " AS RESULTADO"
                End If
            End If
            
            sql = sql & ", 0.000000 AS TOLERANCIA_MAX"
            sql = sql & ", 0.000000 AS CORRECCION"
            sql = sql & ", 0.000000 AS INCERTIDUMBRE"
            
            ' TIPO_ID, N_MEDIDAS, RESULTADO_CUALIDAD, PATRON_ID, DATOS_PARAMETRO
            sql = sql & ", " & .SubItems(COL_id_tipo) & " AS TIPO_ID"
            sql = sql & ", " & .SubItems(COL_n_medidas) & " AS N_MEDIDAS"
            sql = sql & ", " & .SubItems(COL_RESULTADO_CUALIDAD) & " AS RESULTADO_CUALIDAD"
            sql = sql & ", " & .SubItems(COL_id_patron) & " AS PATRON_ID"
            sql = sql & ", '" & .SubItems(COL_RESULTADOS_PATRON) & "' AS DATOS_PARAMETRO"
            ' EQUIPOS, REACTIVOS, REACTIVOS_PROPIOS
            sql = sql & ", '" & .SubItems(COL_LEQUIPOS) & "'"
            sql = sql & ", '" & .SubItems(COL_LREACTIVOS) & "'"
            sql = sql & ", '" & .SubItems(COL_LREACTIVOS_PROPIOS) & "'"
            ' ID_RESULTADO
            sql = sql & ", " & Id_resultado
                            
            execute_bd sql
            
        End With

        
    Next i
 '   execute_bd "delete from eq_verificacion_parametros_resultados WHERE VERIFICACION_ID = 0 AND ID_RESULTADO=0 AND EQUIPO_ID = 0"


End Function

Public Function GuardarParametrosVerificacion_old(ByVal prmid_equipo As Long, ByVal prmid_verificacion As Long, ByRef xres As XArrayDB, ByVal filas As Long) As Boolean
    Dim i As Long, sql As String, Id_resultado As String
    Dim blnCrearParametro As Boolean

    ' Elimina primero todos los paramtros guardados hasta el momento
    sql = "delete from eq_verificacion_parametros_resultados WHERE VERIFICACION_ID = " & CStr(prmid_verificacion) & " AND EQUIPO_ID = " & CStr(prmid_equipo)
    
    execute_bd sql

    For i = 0 To filas
        If Not IsEmpty(xres(i, 0)) Then ' mira se existe la descripción
            If Trim(xres(i, 0)) <> "" Then
                Id_resultado = "-1"
                If Not IsEmpty(xres(i, CONST_ID_RESULTADO)) Then
                    If xres(i, CONST_ID_RESULTADO) = "" Or xres(i, CONST_ID_RESULTADO) = "-1" Then ' Alta o modificacion de Registro
                        Id_resultado = "coalesce(max(id_resultado),0) + 1 as id_resultado FROM eq_verificacion_parametros_resultados"
                    Else
                        Id_resultado = xres(i, CONST_ID_RESULTADO)
                    End If
                Else
                    Id_resultado = "coalesce(max(id_resultado),0) + 1 as id_resultado FROM eq_verificacion_parametros_resultados"
                End If
                            
                sql = "INSERT INTO eq_verificacion_parametros_resultados (EQUIPO_ID, unidad_id, verificacion_id, descripcion, rango_min, rango_max, resultado, tolerancia_max, correccion, incertidumbre, ID_RESULTADO)"
                sql = sql & " SELECT " & CStr(prmid_equipo) & " AS ID_EQUIPO"
                
                If IsEmpty(xres(i, CONST_ID_UNIDAD)) Then
                    sql = sql & ", -1 AS UNIDAD_ID"
                ElseIf Trim(xres(i, CONST_ID_UNIDAD)) = "" Then
                    sql = sql & ", -1 AS UNIDAD_ID"
                Else
                    sql = sql & ", " & xres(i, CONST_ID_UNIDAD) & " AS UNIDAD_ID"
                End If
                
                sql = sql & ", " & CStr(prmid_verificacion) & " AS VERIFICACION_ID"
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
                
                If IsEmpty(xres(i, CONST_RESULTADO_VER)) Then
                    sql = sql & ", 0.000000 AS RESULTADO"
                ElseIf Trim(xres(i, CONST_RESULTADO_VER)) = "" Then
                    sql = sql & ", 0.000000 AS RESULTADO"
                Else
                    sql = sql & ", " & Replace(xres(i, CONST_RESULTADO_VER), ",", ".") & " AS RESULTADO"
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

Public Sub imprimir_etiqueta(ID As Long)
    Dim prnPrinter As Printer
    ' se mira si el equipo tiene impresora de etiquetas
    Dim oParametro As New clsParametros
   On Error GoTo imprimir_etiqueta_Error

    log ("Impresion etiqueta verificacion")
    Dim oEV As New clsEquipoVerificacion
    oEV.Carga ID
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
           .setTIPO_ID = C_ETIQUETAS_TIPOS.C_ETIQUETAS_TIPOS_EQ_VER
           .setCENTRO_ID = CENTROS.CENTRO_MADRID
           .setID = ID
'           .setUSUARIO_ID = oEV.getVERIFICADOR_INTERNO_ID
           .setUSUARIO_ID = oEV.getCUSERID
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
'        Dim oEV As New clsEquipoVerificacion
'        oEV.Carga ID
        oUsuario.CARGAR oEV.getCUSERID
        If ReadINI(App.Path + "\config.ini", "Documentos", "firmas") = "c:\geslab\recursos\firmas" Then
            Dim firma As String
            Dim lista() As String
            If oUsuario.getFIRMA <> "" Then
                lista = Split(oUsuario.getFIRMA, "/")
                firma = ReadINI(App.Path + "\config.ini", "Documentos", "firmas") & "\" & lista(UBound(lista))
            End If
            P2(1) = firma
        Else
            P2(1) = Replace(oUsuario.getFIRMA, "/", "\")
        End If
        With frmReport
            .iniciar
            .informe = "\Equipos\rptEquipos_ETIQUETA_VerificacionPEQ"
            .criterio = "{eq_verificacion_equipos.ID_VERIFICACION} =" & ID
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imprimir_etiqueta of Módulo de clase clsEquipoVerificacion"

End Sub
Public Sub Reabrir(ByVal prmId As String)
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE EQ_VERIFICACION_EQUIPOS SET " & _
                    " ESTADO = 0," & _
                    " RESULTADO = 0," & _
                    " MUSERID = " & CStr(USUARIO.getID_EMPLEADO) & "," & _
                    " TS = CURRENT_TIMESTAMP" & _
            " WHERE ID_VERIFICACION = " & prmId
            
    execute_bd consulta
    Exit Sub
fallo:
    MsgBox "Error al reabrir (clsEquiposVerificacion)", vbCritical, Err.Description
End Sub

Public Function duplicar(ID As Long) As Boolean
    Dim oV As New clsEquipoVerificacion
    Dim idVer As Boolean
    If oV.Carga(ID) Then
        idVer = oV.crear_verificacion_pendiente(ID, oV.getEQUIPO_ID, True)
    End If
    duplicar = idVer
End Function

