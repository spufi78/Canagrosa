VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipoMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQ_MANTENIMIENTO_EQUIPOS
'***************************************************************
Private id_mantenimiento As Long
Private EQUIPO_ID As Long
Private PLANMTO_ID As Long
Private PROCEDIMIENTO_ID As Long
Private MANTENEDOR_ID As Long
Private fecha_actual As String
Private OBSERVACIONES As String
Private ESTADO As Integer
Private ID_AUX As Integer
Private CUSERID As Long
Private MUSERID As Long
Private TS As String
Private responsable As String
Private PROCEDIMIENTO As String
Private PLAN_MANTENIMIENTO As String
Private PERIODICIDAD As String
Private RUTA_CERTIFICADO As String

'Private mvarobjCertificado As New clsArchivoAdjunto

Private ACCIONES As String
Private mvararrSQLInsercion() As String
Private mvarlngIndiceSQLInsercion As Long

Private PERIODICIDAD_ID As Long
Private Function devolver_id_prox_mantenimiento(Optional de_otros_planes As Boolean = False) As Long
Dim rs As ADODB.Recordset
Dim consulta As String

    de_otros_planes = False

    consulta = "SELECT em.id_mantenimiento, em.fecha_actual from eq_mantenimiento_equipos em " & _
    "where em.id_mantenimiento <> " & CStr(id_mantenimiento) & _
    " and em.PLANMTO_ID = " & CStr(PLANMTO_ID) & _
    " and em.equipo_id = " & CStr(EQUIPO_ID) & _
    " and em.estado = 0 " & _
    " order by em.fecha_actual asc, em.ID_MANTENIMIENTO asc"
    
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        devolver_id_prox_mantenimiento = 0
        
        ' consulta de otros mantenimientos pertenecientes a otros planes de mant.
        consulta = "SELECT em.* from eq_mantenimiento_equipos em " & _
        "where em.PLANMTO_ID <> " & CStr(PLANMTO_ID) & _
        " and em.equipo_id = " & CStr(EQUIPO_ID) & _
        " and em.estado = 0 " & _
        " order by em.ID_MANTENIMIENTO asc"
        
        Set rs = datos_bd(consulta)
        If rs.RecordCount <> 0 Then
            de_otros_planes = True
        End If
    Else
        rs.MoveFirst
        devolver_id_prox_mantenimiento = rs!id_mantenimiento
    End If
    
    
    
    
    
    
    
    
    
    Set rs = Nothing

End Function


Public Function devolver_listado_equipos_sin_mto_previstos() As ADODB.Recordset

Dim consulta As String, consulta2 As String, consulta_fin, ANNO As Integer, fecha_aviso As String
Dim oParam As New clsParametros

    ANNO = Year(Now)

    If Not oParam.carga(parametros.EQ_FECHA_AVISO_MTO_ANNO_SIGUIENTE, "") Then
        fecha_aviso = ""
    Else
        fecha_aviso = oParam.getVALOR & "/" & CStr(ANNO)
    End If

    consulta = " SELECT * FROM ("
    consulta = consulta & " select distinct "
    consulta = consulta & " epm.EQUIPO_id as id_equipo"
    consulta = consulta & " , e.NOMBRE as nombre_equipo"
    consulta = consulta & " , epm.PLAN_MANTENIMIENTO_ID"
    consulta = consulta & " , plan.NOMBRE as nombre_plan"
    consulta = consulta & " , plan.FRECUENCIA_ID as periodicidad_id"
    consulta = consulta & " , per.DESCRIPCION as nombre_periodicicad"
    consulta = consulta & " , cast(coalesce(date_format(last_mto.fecha_actual, '%d/%m/%Y'), '--') as char(10)) as fecha_ult_mto"
    consulta = consulta & " , cast(" & CStr(ANNO) & " as unsigned) as anno"
    consulta = consulta & " from eq_planes_mantenimiento_equipos epm"
    consulta = consulta & " inner join equipos e on epm.EQUIPO_ID = e.ID_EQUIPO and e.ALTA_BAJA = 0 and e.CON_MANTENIMIENTO = 1"
    consulta = consulta & " inner join eq_planes_mantenimiento plan on epm.PLAN_MANTENIMIENTO_ID = plan.ID_PLAN_MTO"
    consulta = consulta & " left outer join (select equipo_id, planmto_id, fecha_actual from eq_mantenimiento_equipos as eqmto where fecha_actual = (select max(fecha_actual) from eq_mantenimiento_equipos where equipo_id = eqmto.equipo_id and estado <> 0  and planmto_id = eqmto.PLANMTO_ID)) as last_mto on epm.EQUIPO_ID = last_mto.equipo_id and epm.PLAN_MANTENIMIENTO_ID = last_mto.planmto_id "
    consulta = consulta & " inner join eq_periodicidad per on plan.FRECUENCIA_ID = per.ID_PERIODICIDAD"
    consulta = consulta & " where concat(epm.EQUIPO_ID,':',epm.PLAN_MANTENIMIENTO_ID) not in (select concat(m.equipo_id,':',m.PLANMTO_ID) from eq_mantenimiento_equipos m where estado=0 and year(fecha_actual) = " & CStr(ANNO) & ")"
    
    
    consulta_fin = " ) AS CONSULTA order by anno, id_equipo, plan_mantenimiento_id"


    If CDate(fecha_aviso) <= CDate(Format(Now, "dd/mm/yyyy 00:00:00")) Then
        ANNO = ANNO + 1
        consulta2 = " UNION"
        consulta2 = consulta2 & " select distinct "
        consulta2 = consulta2 & " epm.EQUIPO_id as id_equipo"
        consulta2 = consulta2 & " , e.NOMBRE as nombre_equipo"
        consulta2 = consulta2 & " , epm.PLAN_MANTENIMIENTO_ID"
        consulta2 = consulta2 & " , plan.NOMBRE as nombre_plan"
        consulta2 = consulta2 & " , plan.FRECUENCIA_ID as periodicidad_id"
        consulta2 = consulta2 & " , per.DESCRIPCION as nombre_periodicicad"
        consulta2 = consulta2 & " , cast(coalesce(date_format(last_mto.fecha_actual, '%d/%m/%Y'), '--') as char(10)) as fecha_ult_mto"
        consulta2 = consulta2 & " , cast(" & CStr(ANNO) & " as unsigned) as anno"
        consulta2 = consulta2 & " from eq_planes_mantenimiento_equipos epm"
        consulta2 = consulta2 & " inner join equipos e on epm.EQUIPO_ID = e.ID_EQUIPO and e.ALTA_BAJA = 0 and e.CON_MANTENIMIENTO = 1"
        consulta2 = consulta2 & " inner join eq_planes_mantenimiento plan on epm.PLAN_MANTENIMIENTO_ID = plan.ID_PLAN_MTO"
        consulta2 = consulta2 & " left outer join (select equipo_id, planmto_id, fecha_actual from eq_mantenimiento_equipos as eqmto where fecha_actual = (select max(fecha_actual) from eq_mantenimiento_equipos where equipo_id = eqmto.equipo_id and estado <> 0  and planmto_id = eqmto.PLANMTO_ID)) as last_mto on epm.EQUIPO_ID = last_mto.equipo_id and epm.PLAN_MANTENIMIENTO_ID = last_mto.planmto_id "
        consulta2 = consulta2 & " inner join eq_periodicidad per on plan.FRECUENCIA_ID = per.ID_PERIODICIDAD"
        consulta2 = consulta2 & " where concat(epm.EQUIPO_ID,':',epm.PLAN_MANTENIMIENTO_ID) not in (select concat(m.equipo_id,':',m.PLANMTO_ID) from eq_mantenimiento_equipos m where estado=0 and year(fecha_actual) = " & CStr(ANNO) & ")"
    Else
        consulta2 = ""
    End If

    Set devolver_listado_equipos_sin_mto_previstos = datos_bd(consulta & consulta2 & consulta_fin)
    
    Set oParam = Nothing

End Function

Public Property Get getPERIODICIDAD() As String
    getPERIODICIDAD = PERIODICIDAD
End Property


Public Property Get getPERIODICIDAD_ID() As Long
    getPERIODICIDAD_ID = PERIODICIDAD_ID
End Property

Public Property Get getPLAN_MANTENIMIENTO() As String
    getPLAN_MANTENIMIENTO = PLAN_MANTENIMIENTO
End Property

'Public Function revisar_mantenimiento_pendiente() As Boolean
'
'    ' Recoge los mantenimientos como operaciones pendientes, según mismo equipo y mismo plan
'    ' verifica con los que hay pendientes, si
'
'    ' IMPORTANTE: SOLO EJECUTA ESTO CUANDO SE DA DE ALTA UN MANTENIMIENTO COMO PREVISTO
'
'    Dim id_previsto As Long
'    Dim oOP As New clsEquiposOperacionesPendientes
'    Dim rs As ADODB.Recordset
'    Dim consulta As String
'
'
'    id_previsto = devolver_id_prox_mantenimiento
'
'
'    If id_previsto = 0 Then
'        ' no hay mantenimientos pendientes para ese plan/equipo, por lo que hay que crear una operacion pendiente para este mantenimiento
'                                         ' pone 0 porque es el primero
'        oOP.crear_mantenimiento_pendiente 0, id_mantenimiento, EQUIPO_ID, New clsEquiposPeriodicidad
'    Else
'
'        ' recupera todas los operaciones pendientes de mantenimiento para este equipo, que pertenezcan al mismo plan
'
'        Set rs = oOP.devolver_registros_mantenimiento_segun_plan(EQUIPO_ID, PLANMTO_ID)
'
'        If rs.RecordCount <> 0 Then
'            ' comprueba que fecha es anterior, y si es el caso, modifica el registro de operaciones pendientes
'            rs.MoveFirst
'            If CDate(rs!fecha_actual) > CDate(fecha_actual) Then
'                ' En este caso, el mantenimiento que se da de alta es anterior a cualquier otro mantenimiento de este equipo del mismo plan, así que hay que modificarlo
'                oOP.crear_mantenimiento_pendiente rs!id_mantenimiento, id_mantenimiento, EQUIPO_ID, New clsEquiposPeriodicidad
'            End If
'        End If
'    End If
'
'
'
'
'End Function

Public Function revisar_mantenimiento_pendiente_cuaderno(ByVal prm_id_equipo As Long, ByVal prm_id_plan As Long, ByVal prm_ID_MANTENIMIENTO As Long) As Boolean

    ' Recoge los mantenimientos como operaciones pendientes, según mismo equipo y mismo plan
    ' verifica con los que hay pendientes, si
    
    ' IMPORTANTE: SOLO EJECUTA ESTO CUANDO SE DA DE ALTA UN MANTENIMIENTO COMO PREVISTO
    
    Dim id_previsto As Long
    Dim oOP As New clsEquiposOperacionesPendientes
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim existe_tarea_pendiente As Boolean
    
    
    existe_tarea_pendiente = oOP.carga(prm_ID_MANTENIMIENTO, EQUIPOS_CVM_TIPOS.CVM_TIPO_MANTENIMIENTO)
    ' Recoge los mantenimientos pendientes del mismo equipo y plan de mantenimiento
    Set rs = datos_bd("select id_mantenimiento, fecha_actual from eq_mantenimiento_equipos where estado = 0 and equipo_id = " & prm_id_equipo & " and planmto_id = " & prm_id_plan & " order by fecha_actual asc")
    
    If existe_tarea_pendiente Then
        ' si esta tarea estaba como tarea pendientes
        If rs.RecordCount <> 0 Then
            ' recoge el de fecha más temprana
            rs.MoveFirst
            oOP.crear_mantenimiento_pendiente prm_ID_MANTENIMIENTO, rs!id_mantenimiento, prm_id_equipo, New clsEquiposPeriodicidad
        End If
    End If
    
    ' Si no existe el mantenimiento cerrado como tarea pendiente, no hace nada.

End Function


Public Property Let setPERIODICIDAD(ByVal dato As String)
    PERIODICIDAD = dato
End Property

Public Property Let setPLAN_MANTENIMIENTO(ByVal dato As String)
    PLAN_MANTENIMIENTO = dato
End Property

Public Property Let SQLGenesisFechasMto(ByVal strSQLGenesisFechasMto As String)

    mvarlngIndiceSQLInsercion = mvarlngIndiceSQLInsercion + 1
    ReDim Preserve mvararrSQLInsercion(mvarlngIndiceSQLInsercion)
    mvararrSQLInsercion(mvarlngIndiceSQLInsercion) = strSQLGenesisFechasMto

End Property

Private Sub Class_Initialize()
mvarlngIndiceSQLInsercion = -1
End Sub

Private Sub Class_Terminate()

'    Set mvarobjCertificado = Nothing

End Sub



'Public Property Get CERTIFICADO() As clsArchivoAdjunto
'
'    Set CERTIFICADO = mvarobjCertificado
'
'End Property
'
'Public Property Set CERTIFICADO(objCertificado As clsArchivoAdjunto)
'
'    Set mvarobjCertificado = objCertificado
'
'End Property

Public Property Let setPROCEDIMIENTO(ByVal dato As String)
    PROCEDIMIENTO = dato
End Property
Public Property Get getPROCEDIMIENTO() As String
    getPROCEDIMIENTO = PROCEDIMIENTO
End Property


Property Get getRESPONSABLE() As String
    getRESPONSABLE = responsable
End Property

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQ_MANTENIMIENTO_EQUIPOS
'***************************************************************
Public Property Let setID_MANTENIMIENTO(ByVal dato As Long)
        id_mantenimiento = dato
End Property
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setPLANMTO_ID(ByVal dato As Long)
        PLANMTO_ID = dato
End Property
Public Property Let setPROCEDIMIENTO_ID(ByVal dato As Long)
        PROCEDIMIENTO_ID = dato
End Property
Public Property Let setMANTENEDOR_ID(ByVal dato As Long)
        MANTENEDOR_ID = dato
End Property
Public Property Let setFECHA_ACTUAL(ByVal dato As String)
        fecha_actual = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setESTADO(ByVal dato As Integer)
        ESTADO = dato
End Property

Public Property Let setID_AUX(ByVal dato As Integer)
        ID_AUX = dato
End Property

Public Property Let setCUSERID(ByVal dato As Long)
        CUSERID = dato
End Property
Public Property Let setACCIONES(ByVal dato As String)
        ACCIONES = dato
End Property
Public Property Let setMUSERID(ByVal dato As Long)
        MUSERID = dato
End Property
Public Property Let setRESPONSABLE(ByVal dato As String)
    responsable = dato
End Property
Public Property Let setTS(ByVal dato As String)
        TS = dato
End Property
Public Property Let setRUTA_CERTIFICADO(ByVal dato As String)
        RUTA_CERTIFICADO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQ_MANTENIMIENTO_EQUIPOS
'***************************************************************
Public Property Get getID_MANTENIMIENTO() As Long
        getID_MANTENIMIENTO = id_mantenimiento
End Property
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getPLANMTO_ID() As Long
        getPLANMTO_ID = PLANMTO_ID
End Property
Public Property Get getPROCEDIMIENTO_ID() As Long
        getPROCEDIMIENTO_ID = PROCEDIMIENTO_ID
End Property
Public Property Get getMANTENEDOR_ID() As Long
        getMANTENEDOR_ID = MANTENEDOR_ID
End Property
Public Property Get getFECHA_ACTUAL() As String
        getFECHA_ACTUAL = fecha_actual
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getESTADO() As Integer
        getESTADO = ESTADO
End Property
Public Property Get getID_AUX() As Integer
        getID_AUX = ID_AUX
End Property

Public Property Get getCUSERID() As Long
        getCUSERID = CUSERID
End Property
Public Property Get getACCIONES() As String
        getACCIONES = ACCIONES
End Property

Public Property Get getMUSERID() As Long
        getMUSERID = MUSERID
End Property
Public Property Get getTS() As String
        getTS = TS
End Property
Public Property Get getRUTA_CERTIFICADO() As String
        getRUTA_CERTIFICADO = RUTA_CERTIFICADO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQ_MANTENIMIENTO_EQUIPOS
'***************************************************************
Public Function carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        
        'consulta = "SELECT * FROM EQ_MANTENIMIENTO_EQUIPOS WHERE ID_MANTENIMIENTO = " & ID
        consulta = " select eq_mantenimiento_equipos.*, COALESCE(ca_documentos.NOMBRE, '') as PROCEDIMIENTO, concat(usuarios.NOMBRE, ' ', usuarios.APELLIDOS) as RESPONSABLE, planes.PLAN, planes.PERIODICIDAD, planes.FRECUENCIA_ID"
        consulta = consulta & " From eq_mantenimiento_equipos"
        consulta = consulta & " inner join usuarios on eq_mantenimiento_equipos.MANTENEDOR_ID = usuarios.ID_EMPLEADO"
        consulta = consulta & " left join (select eq_planes_mantenimiento.ID_PLAN_MTO, eq_planes_mantenimiento.NOMBRE AS 'PLAN', decodificadora.descripcion as 'PERIODICIDAD', eq_planes_mantenimiento.FRECUENCIA_ID From eq_planes_mantenimiento inner join decodificadora on eq_planes_mantenimiento.frecuencia_id = decodificadora.valor and decodificadora.CODIGO = " & DECODIFICADORA.EQ_periodicidad & ") as planes on eq_mantenimiento_equipos.planmto_id = planes.ID_PLAN_MTO "
        consulta = consulta & " LEFT join ca_documentos ON eq_mantenimiento_equipos.PROCEDIMIENTO_ID = CA_DOCUMENTOS.ID_DOCUMENTO"
        consulta = consulta & " where eq_mantenimiento_equipos.ID_MANTENIMIENTO = " & CStr(ID)
        
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                carga = False
        Else
                id_mantenimiento = rs("ID_MANTENIMIENTO")
                EQUIPO_ID = rs("EQUIPO_ID")
                PLANMTO_ID = rs("PLANMTO_ID")
                PROCEDIMIENTO_ID = rs("PROCEDIMIENTO_ID")
                PROCEDIMIENTO = rs("PROCEDIMIENTO")
                If IsNull(rs("PLAN")) Then
                PLAN_MANTENIMIENTO = 0
                Else
                PLAN_MANTENIMIENTO = rs("PLAN")
                End If
                If IsNull(rs("PERIODICIDAD")) Then
                PERIODICIDAD = 0
                Else
                PERIODICIDAD = rs("PERIODICIDAD")
                End If
                MANTENEDOR_ID = rs("MANTENEDOR_ID")
                responsable = rs("RESPONSABLE")
                fecha_actual = rs("FECHA_ACTUAL")
                OBSERVACIONES = rs("OBSERVACIONES")
                ESTADO = rs("ESTADO")
                CUSERID = rs("CUSERID")
                MUSERID = rs("MUSERID")
                TS = rs("TS")
                If IsNull(rs("FRECUENCIA_ID")) Then
                PERIODICIDAD_ID = 0
                Else
                PERIODICIDAD_ID = rs("FRECUENCIA_ID")
                End If
                RUTA_CERTIFICADO = rs("RUTA_CERTIFICADO")
                carga = True
                
                Call cargar_acciones_realizadas(rs("ID_MANTENIMIENTO"))
'                Call Carga_ArchivosAdjuntos(rs("RUTA_CERTIFICADO"))
        End If
        Set rs = Nothing
        Exit Function
fallo:
        carga = False
        MsgBox "Error al cargar los datos(clsEquipoMantenimiento)", vbCritical, Err.Description
End Function

Public Function Carga_ultimo_mto_realizado(ByVal prm_id_equipo As Long, ByVal prm_id_plan As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        
        'consulta = "SELECT * FROM EQ_MANTENIMIENTO_EQUIPOS WHERE ID_MANTENIMIENTO = " & ID
        consulta = " select eqmto.*, COALESCE(ca_documentos.NOMBRE, '') as PROCEDIMIENTO, concat(usuarios.NOMBRE, ' ', usuarios.APELLIDOS) as RESPONSABLE, planes.PLAN, planes.PERIODICIDAD"
        consulta = consulta & " From eq_mantenimiento_equipos eqmto"
        consulta = consulta & " inner join usuarios on eqmto.MANTENEDOR_ID = usuarios.ID_EMPLEADO"
        consulta = consulta & " inner join (select eq_planes_mantenimiento.ID_PLAN_MTO, eq_planes_mantenimiento.NOMBRE AS 'PLAN', decodificadora.descripcion as 'PERIODICIDAD' From eq_planes_mantenimiento inner join decodificadora on eq_planes_mantenimiento.frecuencia_id = decodificadora.valor and decodificadora.CODIGO = " & DECODIFICADORA.EQ_periodicidad & ") as planes on eqmto.planmto_id = planes.ID_PLAN_MTO "
        consulta = consulta & " LEFT join ca_documentos ON eqmto.PROCEDIMIENTO_ID = CA_DOCUMENTOS.ID_DOCUMENTO"
        'consulta = consulta & " where eq_mantenimiento_equipos.ID_MANTENIMIENTO = " & CStr(ID)
        consulta = consulta & " where eqmto.fecha_actual = (select max(fecha_actual) from eq_mantenimiento_equipos where equipo_id = eqmto.equipo_id and estado <> 0  and planmto_id = eqmto.PLANMTO_ID)"
        consulta = consulta & " and eqmto.equipo_id = " & prm_id_equipo & " and eqmto.planmto_id = " & prm_id_plan
        
        
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga_ultimo_mto_realizado = False
        Else
                id_mantenimiento = rs("ID_MANTENIMIENTO")
                EQUIPO_ID = rs("EQUIPO_ID")
                PLANMTO_ID = rs("PLANMTO_ID")
                PROCEDIMIENTO_ID = rs("PROCEDIMIENTO_ID")
                PROCEDIMIENTO = rs("PROCEDIMIENTO")
                PLAN_MANTENIMIENTO = rs("PLAN")
                PERIODICIDAD = rs("PERIODICIDAD")
                MANTENEDOR_ID = rs("MANTENEDOR_ID")
                responsable = rs("RESPONSABLE")
                fecha_actual = rs("FECHA_ACTUAL")
                OBSERVACIONES = rs("OBSERVACIONES")
                ESTADO = rs("ESTADO")
                CUSERID = rs("CUSERID")
                MUSERID = rs("MUSERID")
                TS = rs("TS")
                RUTA_CERTIFICADO = rs("RUTA_CERTIFICADO")
                Carga_ultimo_mto_realizado = True
                
                Call cargar_acciones_realizadas(rs("ID_MANTENIMIENTO"))
'                Call Carga_ArchivosAdjuntos(rs("RUTA_CERTIFICADO"))
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga_ultimo_mto_realizado = False
        MsgBox "Error al cargar los datos(clsEquipoMantenimiento)", vbCritical, Err.Description
End Function

'Private Sub Carga_ArchivosAdjuntos(ruta_cert)
'    Dim RUTA_DOCS As String
'
'    Set mvarobjCertificado = New clsArchivoAdjunto
'    RUTA_DOCS = ReadINI(App.Path & "\config.ini", "Documentos", "Ruta")
'
'    mvarobjCertificado.setSUBRUTA_SERVIDOR = "EQUIPOS\" & CStr(EQUIPO_ID) & "\MTO\" & id_mantenimiento & "\CERT"
'    mvarobjCertificado.setRUTA = RUTA_DOCS & "\" & mvarobjCertificado.getSUBRUTA_SERVIDOR & "\" & ruta_cert
'
'
'End Sub

Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_MANTENIMIENTO) FROM EQ_MANTENIMIENTO_EQUIPOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                id_mantenimiento = 1
        Else
                id_mantenimiento = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        
        consulta = "INSERT INTO EQ_MANTENIMIENTO_EQUIPOS (ID_MANTENIMIENTO, CUSERID, TS) VALUES " & _
        "(" & id_mantenimiento & ", " & USUARIO.getID_EMPLEADO & ", LOCALTIMESTAMP)"
                
        execute_bd consulta
        
        Insertar = id_mantenimiento
        
        Call Me.Modificar(id_mantenimiento, True)
        
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEquipoMantenimiento)", vbCritical, Err.Description
End Function


Public Function InsertarFechasMtoGeneradas(Optional ByVal prmlng_ID_Equipo As Long = 0) As Boolean
Dim lngCont As Long
Dim oEq As New clsEquipos

    For lngCont = 0 To mvarlngIndiceSQLInsercion
        'Debug.Print mvararrSQLInsercion(lngCont) & vbCrLf
        execute_bd mvararrSQLInsercion(lngCont)
        Debug.Print mvararrSQLInsercion(lngCont) & ";"
    Next lngCont
   
    oEq.Actualizar_Fecha_Prox_Mantenimiento prmlng_ID_Equipo, CDate("31/12/1899")
    Set oEq = Nothing
   
End Function

Public Function Modificar(ID As Long, Optional ByVal prmEs_Alta As Boolean = False) As Boolean
    On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE EQ_MANTENIMIENTO_EQUIPOS SET " & _
                        " EQUIPO_ID = " & EQUIPO_ID & "," & _
                        " PLANMTO_ID = " & PLANMTO_ID & "," & _
                        " PROCEDIMIENTO_ID = " & PROCEDIMIENTO_ID & "," & _
                        " MANTENEDOR_ID = " & MANTENEDOR_ID & "," & _
                        " FECHA_ACTUAL = '" & Format(fecha_actual, "yyyy/mm/dd") & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " ESTADO = " & ESTADO & "," & _
                        " RUTA_CERTIFICADO = '" & RUTA_CERTIFICADO & "'," & _
                        " MUSERID = " & MUSERID & "," & _
                        " TS = LOCALTIMESTAMP " & _
                " WHERE ID_MANTENIMIENTO = " & ID
        
        execute_bd consulta
        
        id_mantenimiento = ID
        comprobar_cierre_mantenimiento EQUIPO_ID, ESTADO
        
                
        Dim lng_id_mant_previsto As Long, ANNO As Long, de_otros_planes As Boolean
        
        
        If ESTADO = 1 Then
            
            lng_id_mant_previsto = devolver_id_prox_mantenimiento(de_otros_planes)
            
            If lng_id_mant_previsto = 0 Then
                If de_otros_planes Then
                    MsgBox "No hay más mantenimientos Previstos para este equipo, en el presente año, según el plan de mantenimiento: " & PLAN_MANTENIMIENTO & "." & vbCrLf & "Aunque sí existen mantenimientos previstos para este equipo, según otros planes de Mantenimiento. ", vbInformation, "Mantenimiento de Equipos"
                Else
                    MsgBox "No hay más mantenimientos Previstos para este equipo, en el presente año, según el plan de mantenimiento: " & PLAN_MANTENIMIENTO, vbCritical, "Mantenimiento de Equipos"
                End If
                ' Al no haber más mantenimiento para este plan, elimina la tarea pendiente
                
'                oOP.Eliminar id_mantenimiento, EQUIPO_ID, EQUIPOS_CVM_TIPOS.CVM_TIPO_MANTENIMIENTO
            Else
                
'                oOP.crear_mantenimiento_pendiente id_mantenimiento, lng_id_mant_previsto, EQUIPO_ID, New clsEquiposPeriodicidad
                
            End If
        Else
            ' cuando el estado es 0, puede que haya cambiado algo (fecha prevista, responsable, periodicidad)
'OOP            If Not prmEs_Alta Then
'                oOP.revisar_tarea id_mantenimiento, EQUIPO_ID, EQUIPOS_CVM_TIPOS.CVM_TIPO_MANTENIMIENTO
'OOP            End If
        End If
                
        Dim oOP  As New clsEquiposOperacionesPendientes
        oOP.revisar_tarea id_mantenimiento, EQUIPO_ID, EQUIPOS_CVM_TIPOS.CVM_TIPO_MANTENIMIENTO
        Set oOP = Nothing

        'Comprobacion de mantenimientos previstos
        
        Call actualizar_acciones
                
        Dim obj As New clsEquipos
        
        'genera un evento siempre que se cierre el Mantenimiento (ya sea conforme o no conforme)
        If ESTADO > 0 Then ' Se ha cerrado el Mantenimiento
            obj.setID_EQUIPO = EQUIPO_ID
            If ESTADO = 1 Then
                obj.generar_evento EQUIPO_ID, EVT_MANTENIMIENTO_REALIZADO, EVTRZ_RESULTADO_OK, , ID
                
            'Else
            '    obj.generar_evento EQUIPO_ID, EVT_MANTENIMIENTO_REALIZADO, EVTRZ_RESULTADO_NEG, , ID
            End If
        
            ' Por si acaso, se actualiza la fecha de proximo mantenimiento
            obj.Actualizar_Fecha_Prox_Mantenimiento EQUIPO_ID
        End If
                        
        
        Set obj = Nothing
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEquipoMantenimiento)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        carga ID
        
        consulta = "DELETE FROM EQ_MANTENIMIENTO_EQUIPOS " & _
                "    WHERE ID_MANTENIMIENTO = " & ID
        execute_bd consulta
        
        consulta = "DELETE FROM EQ_MANTENIMIENTO_EQUIPOS_ACCIONES " & _
                "    WHERE MANTENIMIENTO_ID = " & ID
        execute_bd consulta
        
        Dim oD As New clsDocumentacion
        oD.EliminarEquipoTipo EQUIPO_ID, 0, ID
        Set oD = Nothing
        
'        mvarobjCertificado.EliminarDeUbicacionFinal
        
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquipoMantenimiento)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM EQ_MANTENIMIENTO_EQUIPOS ORDER BY ID_MANTENIMIENTO"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_MANTENIMIENTO,EQUIPO_ID FROM EQ_MANTENIMIENTO_EQUIPOS ORDER BY EQUIPO_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function

Public Function cargar_acciones_realizadas(ByVal id_mantenimiento As Long, Optional ByVal reemplazar_existentes As Boolean = True) As ADODB.Recordset
Dim consulta As String

    Dim rs As ADODB.Recordset
    
    
    consulta = "select eq_planes_acciones.*, acc_plan.realizado, acc_plan.orden" & _
            " from eq_planes_acciones," & _
            " (select eq_planes_mantenimiento_detalle.*, 0 as realizado" & _
            " from eq_planes_mantenimiento_detalle  where accion_id not in" & _
            " (select accion_id from eq_mantenimiento_equipos_acciones where mantenimiento_id = " & CStr(id_mantenimiento) & ")" & _
            " Union" & _
            " select eq_planes_mantenimiento_detalle.*, 1 as realizado" & _
            " from eq_planes_mantenimiento_detalle  where accion_id in" & _
            " (select accion_id from eq_mantenimiento_equipos_acciones where mantenimiento_id = " & CStr(id_mantenimiento) & ")" & _
            " ) as acc_plan" & _
            " where acc_plan.accion_id = eq_planes_acciones.id_accion " & _
            " and acc_plan.plan_mto_id in (select planmto_id as plan_mto_id from eq_mantenimiento_equipos where id_mantenimiento = " & CStr(id_mantenimiento) & ")" & _
            " order by acc_plan.orden"

    Set rs = datos_bd(consulta)
    
    If rs.RecordCount = 0 Then Exit Function
    
    If reemplazar_existentes Then
        rs.MoveFirst
        ACCIONES = ""
        While Not rs.EOF
            If rs("realizado") = 1 Then
               ACCIONES = ACCIONES & rs("ID_ACCION") & ";"
            End If
            rs.MoveNext
        Wend
        If Trim(ACCIONES) <> "" Then
            ACCIONES = Left(ACCIONES, Len(ACCIONES) - 1)
        End If
    End If
    
    Set cargar_acciones_realizadas = rs
    


End Function

Private Sub actualizar_acciones()
    
    Dim arrAcciones() As String, sql As String
    Dim x As Long, total As Long
    
    sql = "DELETE FROM eq_mantenimiento_equipos_acciones WHERE MANTENIMIENTO_ID = " & id_mantenimiento
    execute_bd sql
    
    If Trim(ACCIONES) <> "" Then
        arrAcciones = Split(ACCIONES, ";")
    
        total = UBound(arrAcciones)
        
        For x = 0 To total
            sql = "insert into eq_mantenimiento_equipos_acciones (plan_mto_id, accion_id, mantenimiento_id) values("
            sql = sql & PLANMTO_ID & ", " & arrAcciones(x) & ", " & id_mantenimiento & ")"
            execute_bd sql
        Next x
    
    End If
    
    

End Sub

Private Sub comprobar_cierre_mantenimiento(ByVal ID_EQUIPO As Long, ByVal ESTADO As Integer)


    Dim strCad As String, rs As ADODB.Recordset
    
    If ESTADO = 0 Then Exit Sub
    
    ' Comprueba si queda algun mantenimiento pendiente
    
    strCad = "select me.equipo_id From eq_mantenimiento_equipos me where me.estado=0 and me.equipo_id = " & CStr(ID_EQUIPO)
    Set rs = datos_bd(strCad)
    
    If rs.RecordCount = 0 Then
        ' Comprueba si no queda algun mantenimiento pendiente, pone la fecha de próximo a inicial
        strCad = "UPDATE equipos SET fecha_prox_mantenimiento = null where id_equipo = " & CStr(EQUIPO_ID)
        execute_bd strCad
        Exit Sub
    End If
    
    strCad = "select me.equipo_id, min(me.fecha_actual) fa From eq_mantenimiento_equipos me where me.estado=0 and me.equipo_id = " & CStr(ID_EQUIPO)
    Set rs = datos_bd(strCad)
    
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        strCad = "UPDATE equipos SET fecha_prox_mantenimiento = '" & Format(rs("fa"), "yyyy-mm-dd") & "' where id_equipo = " & CStr(EQUIPO_ID)
        execute_bd strCad
        Exit Sub
    End If
    
    ' cuando no existe un próximo mantenimiento, no hace nada, pone en blanco los mantenimientos.
    strCad = "UPDATE equipos SET fecha_prox_mantenimiento = '' where id_equipo = " & CStr(EQUIPO_ID)
    execute_bd strCad


End Sub

Public Function comprobar_fin_mantenimientos_previstos(ByVal prm_id_equipo As Long, ByVal prm_id_plan As Long) As Boolean

Dim existen_previstos As Boolean, de_otros_planes As Boolean
Dim rs As New ADODB.Recordset
Dim str_plan As String, consulta As String

'comprueba que no existan más mantenimientos previstos para

consulta = "select * from eq_mantenimiento_equipos em "
consulta = consulta & " where em.estado = 0 and em.equipo_id = " & prm_id_equipo & " and em.planmto_id = " & prm_id_plan
Set rs = datos_bd(consulta)

existen_previstos = (rs.RecordCount <> 0)

If Not existen_previstos Then
    consulta = "select * from eq_mantenimiento_equipos where estado = 0 and equipo_id = " & prm_id_equipo
    Set rs = datos_bd(consulta)
    
    de_otros_planes = (rs.RecordCount <> 0)
End If


If Not existen_previstos Then
    consulta = "select * from eq_planes_mantenimiento where id_plan_mto = " & CStr(prm_id_plan)
    Set rs = datos_bd(consulta)
    
    str_plan = "--"
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        str_plan = rs!nombre
    End If
    
    If de_otros_planes Then
        MsgBox "No hay más Mantenimientos Previstos para este equipo, en el presente año, según el plan de mantenimiento: " & str_plan & "," & vbCrLf & "aunque sí existen mantenimientos previstos para este equipo, pero según otros Planes de Mantenimiento. ", vbInformation, "Mantenimiento de Equipos"
    Else
        MsgBox "No hay más Mantenimientos Previstos para este equipo, en el presente año, según el plan de mantenimiento: " & str_plan, vbCritical, "Mantenimiento de Equipos"
    End If
End If

Set rs = Nothing

End Function
Public Function Cerrar(ID As Long) As Boolean
    On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE EQ_MANTENIMIENTO_EQUIPOS SET " & _
                        " MANTENEDOR_ID = " & MANTENEDOR_ID & "," & _
                        " FECHA_ACTUAL = '" & Format(fecha_actual, "yyyy/mm/dd") & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " ESTADO = " & ESTADO & "," & _
                        " RUTA_CERTIFICADO = '" & RUTA_CERTIFICADO & "'," & _
                        " MUSERID = " & USUARIO.getID_EMPLEADO & "," & _
                        " TS = LOCALTIMESTAMP " & _
                " WHERE ID_MANTENIMIENTO = " & ID
        execute_bd consulta
        id_mantenimiento = ID
        comprobar_cierre_mantenimiento EQUIPO_ID, ESTADO
        Dim lng_id_mant_previsto As Long, ANNO As Long, de_otros_planes As Boolean
        Dim oOP  As New clsEquiposOperacionesPendientes
        oOP.revisar_tarea id_mantenimiento, EQUIPO_ID, EQUIPOS_CVM_TIPOS.CVM_TIPO_MANTENIMIENTO
        Set oOP = Nothing
        'Comprobacion de mantenimientos previstos
'        Call actualizar_acciones
        Dim obj As New clsEquipos
        'genera un evento siempre que se cierre el Mantenimiento (ya sea conforme o no conforme)
        If ESTADO > 0 Then ' Se ha cerrado el Mantenimiento
            obj.setID_EQUIPO = EQUIPO_ID
            If ESTADO = 1 Then
                obj.generar_evento EQUIPO_ID, EVT_MANTENIMIENTO_REALIZADO, EVTRZ_RESULTADO_OK, , ID
            End If
        End If
        ' Por si acaso, se actualiza la fecha de proximo mantenimiento
        obj.Actualizar_Fecha_Prox_Mantenimiento EQUIPO_ID
        Set obj = Nothing
        Cerrar = True
        Exit Function
fallo:
        Cerrar = False
        MsgBox "Error al Cerrar (clsEquipoMantenimiento)", vbCritical, Err.Description
End Function

Public Function CerrarFicticio(ID As Long, EQUIPO_ID As Long) As Boolean
    On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE EQ_MANTENIMIENTO_EQUIPOS SET " & _
                        " ESTADO = 1 " & _
                        ",MUSERID = MANTENEDOR_ID " & _
                        ",CUSERID = MANTENEDOR_ID " & _
                        ",TS = concat(fecha_actual, ' ',date_format(ts,'%h:%m:%s')) " & _
                " WHERE ID_MANTENIMIENTO = " & ID
        execute_bd consulta
        id_mantenimiento = ID
        ' Crear acciones marcadas
        consulta = "insert into eq_mantenimiento_equipos_acciones " & _
        " select b.PLAN_MTO_ID,b.ACCION_ID,a.ID_MANTENIMIENTO from eq_mantenimiento_equipos a, eq_planes_mantenimiento_detalle b where id_mantenimiento = " & ID & " and a.PLANMTO_ID = b.PLAN_MTO_ID"
        execute_bd consulta
'        comprobar_cierre_mantenimiento EQUIPO_ID, ESTADO
        Dim lng_id_mant_previsto As Long, ANNO As Long, de_otros_planes As Boolean
        Dim oOP  As New clsEquiposOperacionesPendientes
        oOP.revisar_tarea id_mantenimiento, EQUIPO_ID, EQUIPOS_CVM_TIPOS.CVM_TIPO_MANTENIMIENTO
        Set oOP = Nothing
        'Comprobacion de mantenimientos previstos
'        Call actualizar_acciones
        Dim obj As New clsEquipos
        obj.Actualizar_Fecha_Prox_Mantenimiento EQUIPO_ID
        ' Evento
        Dim objEvento As New clsEquipoEventos
        Dim IdEvento As Long
        With objEvento
            .setCVM_ID = id_mantenimiento
            .setEQUIPO_ID = EQUIPO_ID
            .setEVENTO_ID = EVT_MANTENIMIENTO_REALIZADO
            .setOBSERVACIONES = ""
            .setRAZON_ID = EVTRZ_RESULTADO_OK
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            IdEvento = .Insertar
        End With
        Dim oem As New clsEquipoMantenimiento
        oem.carga id_mantenimiento
        execute_bd "update eq_eventos set usuario_id = " & oem.getMANTENEDOR_ID & ", ts = '" & Format(oem.getTS, "yyyy-mm-dd hh:mm:ss") & "' where id_eventoequipo = " & IdEvento
        
        Set obj = Nothing
        CerrarFicticio = True
        Exit Function
fallo:
        CerrarFicticio = False
        MsgBox "Error al Cerrar (clsEquipoMantenimiento)", vbCritical, Err.Description
End Function

