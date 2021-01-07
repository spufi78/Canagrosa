VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcNcAccionCorrectora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPROCNCACCIONESCORRECTORAS
'***************************************************************
Private ID_ACCION As Long
Private PROCNC_ID As Long
Private TIPO_ID As Integer
Private titulo As String
Private DESCRIPCION As String
Private FECHA_PREVISTA As Date
Private FECHA_RESOLUCION As Date
Private FECHA_PUESTA_EN_MARCHA As Date
Private RESPONSABLE_ID As Long
Private responsable As String
Private RESOLUCION_EFECTIVA As Integer
Private RESOLUCION_PUESTA_MARCHA_EN_PLAZO As Integer
Private RESOLUCION_EVIDENCIAS As Integer
Private RESOLUCION_COMUNICADO_MODIFICACIONES As Integer
Private ESTADO_ID As Long
Private ESTADO As String
Private mvarintID_AUX As Integer
'JGM-I
'Public Function EnviarMensaje_AccTramitada() As Boolean
'Dim idRespCalidad() As Long
'Dim ASUNTO As String
'Dim texto As String
'
'
'On Error GoTo EnviarMensaje_AccTramitada_Error
'
'    ReDim Preserve idRespCalidad(1)
'
'    idRespCalidad(0) = RESPONSABLE_ID
'
'    ASUNTO = "Accion Asignada. Incidencia/No Conformidad Nº: " & Format(PROCNC_ID, "00000")
'    texto = texto & "El usuario " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & ", como Responsable de Calidad, le ha asignado como responsable de la tramitación de la Accion siguiente: " & vbNewLine & vbNewLine
'    texto = texto & "Titulo: " & titulo & vbNewLine
'    texto = texto & "Descripcion: " & DESCRIPCION & vbNewLine & vbNewLine
'    texto = texto & "Fecha de Puesta en Marcha : " & Format(FECHA_PUESTA_EN_MARCHA, "dd-mm-yyyy") & vbNewLine
'    texto = texto & "Fecha de Prevista Finalización: " & Format(FECHA_PREVISTA, "dd-mm-yyyy") & vbNewLine & vbNewLine
'
'
'    Call EnviarMensajeInterno(idRespCalidad, ASUNTO, texto, "frmProcNCEdicion_AccionCorrectiva;" & CStr(ID_ACCION), Now)
'
'
'
'On Error GoTo 0
'    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.EnviarMensaje_AccTramitada"
'    Exit Function
'EnviarMensaje_AccTramitada_Error:
'    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.EnviarMensaje_AccTramitada"
'    error_grave Err.Number & " (" & Err.Description & ") in procedure EnviarMensaje_AccTramitada of Módulo de clase clsProcNcAccionCorrectora" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
'    G_TRAZABILIDAD_ERROR = ""
'
'End Function

'Public Function EnviarMensaje_Exencion_Acc_Correctiva(ByVal prmResponsable_id_anterior As Long) As Boolean
'Dim idRespCalidad() As Long
'Dim ASUNTO As String
'Dim texto As String
'
'
'On Error GoTo EnviarMensaje_Exencion_Acc_Correctiva_Error
'
'    ReDim Preserve idRespCalidad(1)
'
'    idRespCalidad(0) = prmResponsable_id_anterior
'
'    ASUNTO = "Exento de Accion. Incidencia/No Conformidad Nº: " & Format(PROCNC_ID, "00000")
'    texto = texto & "El Responsable de Calidad le ha desestimado como Responsable de ejecución de la siguiente Acción, por lo que está exento de ejecutar dicha Acción: " & vbNewLine & vbNewLine
'    texto = texto & "Titulo: " & titulo & vbNewLine
'    texto = texto & "Descripcion: " & DESCRIPCION & vbNewLine & vbNewLine
'    texto = texto & "Fecha de Puesta en Marcha : " & Format(FECHA_PUESTA_EN_MARCHA, "dd-mm-yyyy") & vbNewLine
'    texto = texto & "Fecha de Prevista Finalización: " & Format(FECHA_PREVISTA, "dd-mm-yyyy") & vbNewLine & vbNewLine
'
'
'    Call EnviarMensajeInterno(idRespCalidad, ASUNTO, texto, "frmProcNCEdicion_AccionCorrectiva;" & CStr(ID_ACCION), Now)
'
'
'On Error GoTo 0
'    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.EnviarMensaje_Exencion_Acc_Correctiva"
'    Exit Function
'EnviarMensaje_Exencion_Acc_Correctiva_Error:
'    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.EnviarMensaje_Exencion_Acc_Correctiva"
'    error_grave Err.Number & " (" & Err.Description & ") in procedure EnviarMensaje_Exencion_Acc_Correctiva of Módulo de clase clsProcNcAccionCorrectora" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
'    G_TRAZABILIDAD_ERROR = ""
'
'End Function
'JGM-F



Public Function EnviarMensaje_AccCerradaOk() As Boolean

Dim objUsu As New clsUsuarios
Dim idRespCalidad() As Long, indexResp As Long
Dim ASUNTO As String
Dim texto As String
Dim blnYaExistente As Boolean, x As Integer
Dim objPnc As New clsProcNc

On Error GoTo EnviarMensaje_AccCerradaOk_Error

    blnYaExistente = False
    
    Call objPnc.devolverListaResponsablesDepartamentos(idRespCalidad)
    indexResp = UBound(idRespCalidad)
    
    objUsu.CARGAR RESPONSABLE_ID
    
    'Avisa a los responsables de calidad
    For x = 0 To indexResp - 1
        If idRespCalidad(x) = RESPONSABLE_ID Then
            blnYaExistente = True
        End If
    Next x
    If Not blnYaExistente Then
        indexResp = indexResp + 1
        ReDim Preserve idRespCalidad(indexResp)
        idRespCalidad(indexResp - 1) = RESPONSABLE_ID
    End If
        
    ASUNTO = "Accion Con Éxito. Incidencia/No Conformidad Nº: " & Format(PROCNC_ID, "00000")
    texto = texto & "El usuario " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & ", como Responsable de Calidad, ha dado por cerrada con éxito la siguiente acción asignada a " & objUsu.getNOMBRE & " " & objUsu.getAPELLIDOS & ": " & vbNewLine & vbNewLine
    texto = texto & "Titulo: " & titulo & vbNewLine
    texto = texto & "Descripcion: " & DESCRIPCION & vbNewLine & vbNewLine
    texto = texto & "Fecha de Puesta en Marcha : " & Format(FECHA_PUESTA_EN_MARCHA, "dd-mm-yyyy") & vbNewLine
    texto = texto & "Fecha de Prevista Finalización: " & Format(FECHA_PREVISTA, "dd-mm-yyyy") & vbNewLine & vbNewLine
    texto = texto & "Fecha de Real Finalización: " & Format(FECHA_RESOLUCION, "dd-mm-yyyy") & vbNewLine & vbNewLine
        
    Call EnviarMensajeInterno(idRespCalidad, ASUNTO, texto, "frmProcNCEdicion_AccionCorrectiva;" & CStr(ID_ACCION), Now)

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.EnviarMensaje_AccCerradaOk"
    Exit Function
EnviarMensaje_AccCerradaOk_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.EnviarMensaje_AccCerradaOk"
    error_grave Err.Number & " (" & Err.Description & ") in procedure EnviarMensaje_AccCerradaOk of Módulo de clase clsProcNcAccionCorrectora" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Function


Public Function EnviarMensaje_AccFinalizadaPorResponsable(ByVal prmstrObservaciones As String) As Boolean

Dim objUsu As New clsUsuarios
Dim idRespCalidad() As Long, indexResp As Long
Dim ASUNTO As String
Dim texto As String
Dim blnYaExistente As Boolean, x As Integer
Dim objPnc As New clsProcNc

On Error GoTo EnviarMensaje_AccFinalizadaPorResponsable_Error

    blnYaExistente = False
    
    Call objPnc.devolverListaResponsablesCalidad(idRespCalidad)
    indexResp = UBound(idRespCalidad)
    
    'Avisa a los responsables de calidad
    For x = 0 To indexResp - 1
        If idRespCalidad(x) = RESPONSABLE_ID Then
            blnYaExistente = True
        End If
    Next x
    If Not blnYaExistente Then
        indexResp = indexResp + 1
        ReDim Preserve idRespCalidad(indexResp)
        idRespCalidad(indexResp - 1) = RESPONSABLE_ID
    End If
        
    ASUNTO = "Accion Correctiva Pre-Finalizada. Incidencia/No Conformidad Nº: " & Format(PROCNC_ID, "00000")
    texto = texto & "El usuario " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & ", asignado a la siguiente Accion Correctiva, informa a los responsables de calidad que da por finalizada la Acción Correctiva: " & vbNewLine & vbNewLine
    texto = texto & "Titulo: " & titulo & vbNewLine
    texto = texto & "Descripcion: " & DESCRIPCION & vbNewLine & vbNewLine
    texto = texto & "Fecha de Puesta en Marcha : " & Format(FECHA_PUESTA_EN_MARCHA, "dd-mm-yyyy") & vbNewLine
    texto = texto & "Fecha de Prevista Finalización: " & Format(FECHA_PREVISTA, "dd-mm-yyyy") & vbNewLine & vbNewLine
    texto = texto & "Fecha de Finalización según Usuario Asignado: " & Format(Now, "dd-mm-yyyy") & vbNewLine & vbNewLine
    If Trim(prmstrObservaciones) = "" Then
        texto = texto & "El Usuario Asignado no consigna observación alguna." & vbNewLine & vbNewLine
    Else
        texto = texto & "Observaciones Usuario Asignado: " & prmstrObservaciones & vbNewLine & vbNewLine
    End If
    
        
    Call EnviarMensajeInterno(idRespCalidad, ASUNTO, texto, "frmProcNCEdicion_AccionCorrectiva;" & CStr(ID_ACCION), Now)

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.EnviarMensaje_AccFinalizadaPorResponsable"
    Exit Function
EnviarMensaje_AccFinalizadaPorResponsable_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.EnviarMensaje_AccFinalizadaPorResponsable"
    error_grave Err.Number & " (" & Err.Description & ") in procedure EnviarMensaje_AccFinalizadaPorResponsable of Módulo de clase clsProcNcAccionCorrectora" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Function



Private Sub EnviarMensajeInterno(ByRef prmListaUsuarioDestinatario() As Long, ByVal prmAsunto As String, ByVal prmCuerpo As String, ByVal prmEnlaceForm As String, prmFechaEnvio As Date)
Dim objMensaje As New clsMensajes
Dim mens As Long
Dim omu As New clsMensajes_usuarios
Dim x As Integer

On Error GoTo EnviarMensajeInterno_Error

With objMensaje
        .setASUNTO = prmAsunto
        .setTEXTO = prmCuerpo
        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        .setFECHA_INICIO = Format(prmFechaEnvio, "yyyy-mm-dd")
        .setFECHA_FIN = Format(prmFechaEnvio + 7, "yyyy-mm-dd")
        .setACCION = prmEnlaceForm
        
        .setHORA_INICIO = Format(Time, "hh:mm:ss")
        .setHORA_FIN = Format(Time, "hh:mm:ss")
        .setCATEGORIA = MENSAJES_CATEGORIAS.MENSAJES_CATEGORIAS_PROCN
        .setDURACION = 0
        
        mens = .Insertar
        If mens > 0 Then
            For x = 0 To UBound(prmListaUsuarioDestinatario()) - 1
                omu.setEMPLEADO_ID = prmListaUsuarioDestinatario(x)
                omu.setMENSAJE_ID = mens
                omu.Insertar
            Next x
            frmCalendario.cargar_eventos
        End If
End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.EnviarMensajeInterno"
    Exit Sub
EnviarMensajeInterno_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.EnviarMensajeInterno"
    error_grave Err.Number & " (" & Err.Description & ") in procedure EnviarMensajeInterno of Módulo de clase clsProcNcAccionCorrectora" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub
Public Property Get getID_AUX() As Integer
    getID_AUX = mvarintID_AUX
End Property
Public Property Let setID_AUX(ByVal dato As Integer)
    mvarintID_AUX = dato
End Property

Public Property Get getTITULO() As String
    getTITULO = titulo
End Property
Public Property Let setTITULO(ByVal dato As String)
    titulo = dato
End Property

Public Property Get getRESPONSABLE() As String
    getRESPONSABLE = responsable
End Property
Public Property Let setRESPONSABLE(ByVal dato As String)
    responsable = dato
End Property
Public Property Get getTIPO_ID() As String
    getTIPO_ID = TIPO_ID
End Property
Public Property Let setTIPO_ID(ByVal dato As String)
    TIPO_ID = dato
End Property

Public Property Get getESTADO() As String
    getESTADO = ESTADO
End Property
Public Property Let setESTADO(ByVal dato As String)
    ESTADO = dato
End Property

Public Property Get getRESOLUCION_EFECTIVA() As Integer
    getRESOLUCION_EFECTIVA = RESOLUCION_EFECTIVA
End Property
Public Property Let setRESOLUCION_EFECTIVA(ByVal dato As Integer)
     RESOLUCION_EFECTIVA = dato
End Property

Public Property Get getRESOLUCION_COMUNICADO_MODIFICACIONES() As Integer
    getRESOLUCION_COMUNICADO_MODIFICACIONES = RESOLUCION_COMUNICADO_MODIFICACIONES
End Property
Public Property Let setRESOLUCION_COMUNICADO_MODIFICACIONES(ByVal dato As Integer)
     RESOLUCION_COMUNICADO_MODIFICACIONES = dato
End Property

Public Property Get getRESOLUCION_EVIDENCIAS() As Integer
    getRESOLUCION_EVIDENCIAS = RESOLUCION_EVIDENCIAS
End Property
Public Property Let setRESOLUCION_EVIDENCIAS(ByVal dato As Integer)
     RESOLUCION_EVIDENCIAS = dato
End Property

Public Property Get getRESOLUCION_PUESTA_MARCHA_EN_PLAZO() As Integer
    getRESOLUCION_PUESTA_MARCHA_EN_PLAZO = RESOLUCION_PUESTA_MARCHA_EN_PLAZO
End Property
Public Property Let setRESOLUCION_PUESTA_MARCHA_EN_PLAZO(ByVal dato As Integer)
     RESOLUCION_PUESTA_MARCHA_EN_PLAZO = dato
End Property


Public Property Get ID_AUX() As Integer
    ID_AUX = mvarintID_AUX
End Property

Public Property Let ID_AUX(ByVal intID_AUX As Integer)
    mvarintID_AUX = intID_AUX
End Property

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSNC_ACCIONES
'***************************************************************
Public Property Let setID_ACCION(ByVal dato As Long)
        ID_ACCION = dato
End Property
Public Property Let setPROCNC_ID(ByVal dato As Long)
        PROCNC_ID = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setFECHA_PREVISTA(ByVal dato As Date)
        FECHA_PREVISTA = dato
End Property
Public Property Let setFECHA_PUESTA_EN_MARCHA(ByVal dato As Date)
        FECHA_PUESTA_EN_MARCHA = dato
End Property
Public Property Let setFECHA_RESOLUCION(ByVal dato As Date)
        FECHA_RESOLUCION = dato
End Property
Public Property Let setRESPONSABLE_ID(ByVal dato As Long)
        RESPONSABLE_ID = dato
End Property
Public Property Let setESTADO_ID(ByVal dato As Long)
        ESTADO_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSNC_ACCIONES
'***************************************************************
Public Property Get getID_ACCION() As Long
        getID_ACCION = ID_ACCION
End Property
Public Property Get getPROCNC_ID() As Long
        getPROCNC_ID = PROCNC_ID
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getFECHA_PREVISTA() As Date
        getFECHA_PREVISTA = FECHA_PREVISTA
End Property
Public Property Get getFECHA_PUESTA_EN_MARCHA() As Date
        getFECHA_PUESTA_EN_MARCHA = FECHA_PUESTA_EN_MARCHA
End Property
Public Property Get getFECHA_RESOLUCION() As Date
        getFECHA_RESOLUCION = FECHA_RESOLUCION
End Property
Public Property Get getRESPONSABLE_ID() As Long
        getRESPONSABLE_ID = RESPONSABLE_ID
End Property
Public Property Get getESTADO_ID() As Long
        getESTADO_ID = ESTADO_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSNC_ACCIONES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
On Error GoTo Carga_Error

        Dim rs As ADODB.Recordset
        Dim strSql As String
        
        'strSql = "SELECT procnc_accionescorrectivas.*, concat(usuarios.nombre, ' ', usuarios.apellidos) as RESPONSABLE, d2.descripcion as ESTADO" & _
        " FROM procnc_accionescorrectivas " & _
        " join usuarios on procnc_accionescorrectivas.responsable_id = usuarios.id_empleado" & _
        " join decodificadora as d2 on procnc_accionescorrectivas.estado_id = d2.valor and d2.codigo=111" & _
        " WHERE id_accion_correctiva = " & CStr(id)

        strSql = "SELECT procnc_accionescorrectivas.*, concat(usuarios.nombre, ' ', usuarios.apellidos) as RESPONSABLE, d2.descripcion as ESTADO" & _
        " FROM procnc_accionescorrectivas " & _
        " join usuarios on procnc_accionescorrectivas.responsable_id = usuarios.id_empleado" & _
        " join decodificadora as d2 on procnc_accionescorrectivas.estado_id = d2.valor and d2.codigo=111" & _
        " WHERE procnc_accionescorrectivas.id_accion_correctiva = " & CStr(ID)
        
        Set rs = datos_bd(strSql)
        If rs.RecordCount = 0 Then
            Carga = False
        Else
            TIPO_ID = rs("TIPO_ID")
            titulo = rs("TITULO")
            DESCRIPCION = rs("DESCRIPCION")
            ID_ACCION = rs("id_accion_correctiva")
            PROCNC_ID = rs("id_procnc")
            ESTADO_ID = rs("estado_id")
            ESTADO = rs("ESTADO")
            FECHA_PREVISTA = rs("fecha_prevista")
            FECHA_PUESTA_EN_MARCHA = rs("fecha_puesta_en_marcha")
            FECHA_RESOLUCION = rs("fecha_resolucion")
            RESOLUCION_COMUNICADO_MODIFICACIONES = rs("resolucion_comunicado_modificaciones")
            RESOLUCION_EFECTIVA = rs("resolucion_efectiva")
            RESOLUCION_EVIDENCIAS = rs("resolucion_evidencias")
            RESOLUCION_PUESTA_MARCHA_EN_PLAZO = rs("resolucion_puesta_marcha_en_plazo")
            responsable = rs("RESPONSABLE")
            RESPONSABLE_ID = rs("responsable_id")
            ID_AUX = 0
        
'            ID_ACCION = rs("id_accion_correctiva")
'            PROCNC_ID = rs("id_procnc")
'            Titulo = rs("titulo")
'            DESCRIPCION = rs("DESCRIPCION")
'            FECHA_PREVISTA = rs("FECHA_PREVISTA")
'            FECHA_RESOLUCION = rs("FECHA_RESOLUCION")
'            FECHA_PUESTA_EN_MARCHA = rs("fecha_puesta_en_marcha")
'            RESPONSABLE_ID = rs("RESPONSABLE_ID")
'            Responsable = rs("RESPONSABLE")
'            ESTADO_ID = rs("ESTADO_ID")
'            ESTADO = rs("ESTADO")
'            RESOLUCION_COMUNICADO_MODIFICACIONES = rs("resolucion_comunicado_modificaciones")
'            RESOLUCION_EFECTIVA = rs("resolucion_efectiva")
'            RESOLUCION_EVIDENCIAS = rs("resolucion_evidencias")
'            RESOLUCION_PUESTA_MARCHA_EN_PLAZO = rs("resolucion_puesta_marcha_en_plazo")
            
            Carga = True
        End If
        Set rs = Nothing
        Exit Function

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.Carga"
    Exit Function
Carga_Error:
    Carga = False
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.Carga"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Carga of Módulo de clase clsProcNcAccionCorrectora" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
On Error GoTo CrearID_Error

        consulta = "SELECT MAX(ID_ACCION_CORRECTIVA) FROM procnc_accionescorrectivas"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_ACCION = 1
        Else
                ID_ACCION = rs.Fields(0) + 1
        End If
        Set rs = Nothing

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.CrearID"
    Exit Function
CrearID_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.CrearID"
    error_grave Err.Number & " (" & Err.Description & ") in procedure CrearID of Módulo de clase clsProcNcAccionCorrectora" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Function
Public Function Insertar() As Long
On Error GoTo Insertar_Error

        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        
        ESTADO_ID = C_PROCNC_ESTADOS.ABIERTA

        consulta = "INSERT INTO procnc_accionescorrectivas ("
        consulta = consulta & "id_procnc, "
        consulta = consulta & "id_accion_correctiva, "
        consulta = consulta & "tipo_id, "
        consulta = consulta & "titulo, "
        consulta = consulta & "descripcion, "
        consulta = consulta & "fecha_prevista, "
        consulta = consulta & "fecha_puesta_en_marcha, "
        consulta = consulta & "fecha_resolucion, "
        consulta = consulta & "responsable_id, "
        consulta = consulta & "estado_id, "
        consulta = consulta & "resolucion_efectiva, "
        consulta = consulta & "resolucion_puesta_marcha_en_plazo, "
        consulta = consulta & "resolucion_evidencias, "
        consulta = consulta & "resolucion_comunicado_modificaciones "
        consulta = consulta & ") VALUES ("
        consulta = consulta & CStr(PROCNC_ID) & ", "
        consulta = consulta & CStr(ID_ACCION) & ", "
        consulta = consulta & "0, "
        consulta = consulta & "'', "
        consulta = consulta & "'', "
        consulta = consulta & "'1900-01-01 00:00:00', "
        consulta = consulta & "'1900-01-01 00:00:00', "
        consulta = consulta & "'1900-01-01 00:00:00', "
        consulta = consulta & "0, "
        consulta = consulta & "1, "
        consulta = consulta & "-1, "
        consulta = consulta & "-1, "
        consulta = consulta & "-1, "
        consulta = consulta & "-1)"
                       
        execute_bd consulta
        
        Modificar
        
        Insertar = ID_ACCION
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsProcNcAccionesCorrectoras)", vbCritical, Err.Description

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.Insertar"
    Exit Function
Insertar_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.Insertar"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Insertar of Módulo de clase clsProcNcAccionCorrectora" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Function
Public Function Modificar() As Boolean
        On Error GoTo fallo
        Dim consulta As String, DIAS_PREAVISO As String
        Dim rs As ADODB.Recordset
        Dim lng_id_resp_anterior As Long

        lng_id_resp_anterior = 0

        consulta = "select responsable_id from procnc_accionescorrectivas "
        consulta = consulta & " WHERE id_accion_correctiva = " & CStr(ID_ACCION)
        consulta = consulta & " AND id_procnc = " & CStr(PROCNC_ID)
        
        Set rs = datos_bd(consulta)

        If rs.RecordCount <> 0 Then
            rs.MoveFirst
            lng_id_resp_anterior = rs("responsable_id")
            Set rs = Nothing
        End If

        consulta = "update procnc_accionescorrectivas SET "
        consulta = consulta & "tipo_id = " & TIPO_ID
        consulta = consulta & ",titulo = '" & titulo & "'"
        consulta = consulta & ",descripcion = '" & DESCRIPCION & "'"
        consulta = consulta & ",fecha_prevista = '" & Format(FECHA_PREVISTA, "yyyy-mm-dd 00:00:00") & "'"
        consulta = consulta & ",fecha_puesta_en_marcha = '" & Format(FECHA_PUESTA_EN_MARCHA, "yyyy-mm-dd 00:00:00") & "'"
        consulta = consulta & ",fecha_resolucion = '" & Format(FECHA_RESOLUCION, "yyyy-mm-dd 00:00:00") & "'"
        consulta = consulta & ",responsable_id = " & RESPONSABLE_ID & ""
        consulta = consulta & ",estado_id = " & ESTADO_ID & ""
        consulta = consulta & ",resolucion_efectiva = " & RESOLUCION_EFECTIVA
        consulta = consulta & ",resolucion_puesta_marcha_en_plazo = " & RESOLUCION_PUESTA_MARCHA_EN_PLAZO
        consulta = consulta & ",resolucion_evidencias = " & RESOLUCION_EVIDENCIAS
        consulta = consulta & ",resolucion_comunicado_modificaciones = " & RESOLUCION_COMUNICADO_MODIFICACIONES
        consulta = consulta & " WHERE id_accion_correctiva = " & CStr(ID_ACCION)
        consulta = consulta & " AND id_procnc = " & CStr(PROCNC_ID)

        execute_bd consulta
'JGM-I
'        consulta = "delete from procnc_accionescorrectivas_avisos_pendiente where accion_id = " & CStr(ID_ACCION)
'        execute_bd consulta
'
'        If ESTADO_ID = C_PROCNC_ESTADOS.EN_TRAMITACION Then
'            ' DA DE ALTA EL PREAVISO
'            Dim oParam As New clsParametros
'            oParam.Carga parametros.ACCIONES_CORRECTIVAS_DIAS_PREAVISO, ""
'            DIAS_PREAVISO = oParam.getVALOR
'            Set oParam = Nothing
'            consulta = "insert into procnc_accionescorrectivas_avisos_pendiente (ACCION_ID, RESPONSABLE_ID, FECHA_AVISO) VALUES (" & CStr(ID_ACCION) & ", " & RESPONSABLE_ID & ", '" & Format(DateAdd("d", -1 * CLng(DIAS_PREAVISO), FECHA_PREVISTA), "yyyy-mm-dd") & "')"
'            execute_bd consulta
'
'            ' al estar en tramitacion, se comprueba si se ha cambiado el responsable
'            If lng_id_resp_anterior > 0 And lng_id_resp_anterior <> RESPONSABLE_ID Then
'                ' envia un mensaje al responsable anterior diciendo que está exento de responsabilidad para esta accion correctiva
'                EnviarMensaje_Exencion_Acc_Correctiva lng_id_resp_anterior
'                ' envia un mensaje al nuevo responsable de la tramitacion
'                EnviarMensaje_AccTramitada
'            End If
'
'        End If
'JGM-F
        
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsProcNcAccionesCorrectoras)", vbCritical, Err.Description
        
End Function
Public Function Eliminar(ID As Long) As Boolean
On Error GoTo Eliminar_Error

        Dim consulta As String
        consulta = "DELETE FROM procnc_accionescorrectivas " & _
                "    WHERE ID_ACCION_CORRECTIVA = " & ID
        execute_bd consulta
        
'JGM-I
'        consulta = "delete from procnc_accionescorrectivas_avisos_pendiente where id_accion = " & CStr(ID)
'        execute_bd consulta
'JGM-F
        Eliminar = True
        Exit Function

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.Eliminar"
    Exit Function
Eliminar_Error:
    Eliminar = False
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.Eliminar"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Eliminar of Módulo de clase clsProcNcAccionCorrectora" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
On Error GoTo Listado_Error

        consulta = "SELECT A.FECHA_PREVISTA,A.DESCRIPCION,B.APELLIDOS, B.NOMBRE,C.DESCRIPCION,A.FECHA_RESOLUCION,A.ID_ACCION,A.RESPONSABLE_ID,A.ESTADO_ID " & _
                   " FROM NC_ACCIONES A,usuarios B,DECODIFICADORA C " & _
                   " WHERE A.PROCNC_ID = " & ID & _
                   "   AND A.RESPONSABLE_ID = B.ID_EMPLEADO " & _
                   "   AND A.ESTADO_ID = C.VALOR " & _
                   "   AND C.CODIGO = " & DECODIFICADORA.NC_ESTADOS & _
                   " ORDER BY A.ID_ACCION"
        Set Listado = datos_bd(consulta)

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.Listado"
    Exit Function
Listado_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.Listado"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Listado of Módulo de clase clsProcNcAccionCorrectora" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Function
Public Function ListadoAccionesPendientes() As ADODB.Recordset
                
        Dim consulta As String
        
        consulta = "select procnc_accionescorrectivas.*, concat(usuarios.NOMBRE, ' ', usuarios.APELLIDOS) as RESPONSABLE"
        consulta = consulta & " From procnc_accionescorrectivas"
        consulta = consulta & " inner join usuarios on procnc_accionescorrectivas.responsable_id = usuarios.ID_EMPLEADO"
        consulta = consulta & " where 1 = 1 "
        If USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.Calidad) = vbUnchecked Then
            consulta = consulta & "   and procnc_accionescorrectivas.responsable_id = " & CStr(USUARIO.getID_EMPLEADO)
        End If
        consulta = consulta & "   and procnc_accionescorrectivas.estado_id <= " & C_PROCNC_ESTADOS_ACCIONES.EN_TRAMITACION
        consulta = consulta & " order by procnc_accionescorrectivas.FECHA_PREVISTA"
                                      
        Set ListadoAccionesPendientes = datos_bd(consulta)
        
End Function


Public Function ListadoAcciones(DESCRIPCION As String, responsable As Long, checkPrevista As Integer, fprevistaI As String, fprevistaF As String, checkPuesta As Integer, fPuestaI As String, fPuestaF As String, checkResolucion As Integer, fResolucionI As String, fresolucionF As String, idEstado As String, idTipo As String, CNCi As String, CNCf As String) As ADODB.Recordset
                
        Dim consulta As String
        Dim strDescripcion As String
        Dim strRESPONSABLE As String
        Dim strFechaPrevista As String
        Dim strFechaPuesta As String
        Dim strFechaResolucion As String
        Dim strEstado As String
        Dim strTipo As String
        Dim strCNC As String
                           
On Error GoTo ListadoAccionesPendientesRespCalidad_Error

        If Trim(DESCRIPCION) <> "" Then
            strDescripcion = " AND procnc_accionescorrectivas.TITULO like '%" & DESCRIPCION & "%'"
        End If
        If responsable <> 0 Then
            strRESPONSABLE = "AND procnc_accionescorrectivas.responsable_id = " & responsable
        Else
            strRESPONSABLE = ""
        End If
        If checkPrevista = 1 Then
            strFechaPrevista = " AND procnc_accionescorrectivas.fecha_prevista >= '" & Format(fprevistaI, "yyyy-mm-dd") & "' AND procnc_accionescorrectivas.fecha_prevista <= '" & Format(fprevistaF, "yyyy-mm-dd") & "'"
        End If
        If checkPuesta = 1 Then
            strFechaPuesta = " AND procnc_accionescorrectivas.fecha_puesta_en_marcha >= '" & Format(fPuestaI, "yyyy-mm-dd") & "' AND procnc_accionescorrectivas.fecha_puesta_en_marcha <= '" & Format(fPuestaF, "yyyy-mm-dd") & "'"
        End If
        If checkResolucion = 1 Then
            strFechaResolucion = " AND procnc_accionescorrectivas.fecha_resolucion >= '" & Format(fResolucionI, "yyyy-mm-dd") & "' AND procnc_accionescorrectivas.fecha_resolucion <= '" & Format(fresolucionF, "yyyy-mm-dd") & "'"
        End If
        strEstado = ""
        If IsNumeric(idEstado) Then
            If CInt(idEstado) <> 0 Then
                strEstado = " AND procnc_accionescorrectivas.estado_id = " & idEstado
            End If
        End If
        If IsNumeric(idTipo) Then
            If CInt(idTipo) <> 2 Then
                strTipo = " AND procnc_accionescorrectivas.tipo_id = " & idTipo
            End If
        End If
        If IsNumeric(Trim(CNCi)) And IsNumeric(Trim(CNCf)) Then
            strCNC = " AND procnc_accionescorrectivas.id_procnc >= " & CLng(CNCi) & " AND procnc_accionescorrectivas.id_procnc <= " & CLng(CNCf)
        End If
        
        consulta = "select procnc_accionescorrectivas.*, concat(usuarios.NOMBRE, ' ', usuarios.APELLIDOS) as RESPONSABLE, deco1.DESCRIPCION AS ESTADO, deco2.DESCRIPCION as TIPO"
        consulta = consulta & " From procnc_accionescorrectivas"
        consulta = consulta & " inner join usuarios on procnc_accionescorrectivas.responsable_id = usuarios.ID_EMPLEADO"
        consulta = consulta & " inner join decodificadora as deco1 on procnc_accionescorrectivas.estado_id = deco1.valor AND deco1.codigo = " & DECODIFICADORA.PROCNC_ESTADOS_ACCIONES_CORRECTIVAS
        consulta = consulta & " inner join decodificadora as deco2 on procnc_accionescorrectivas.tipo_id = deco2.valor AND deco2.codigo = " & DECODIFICADORA.PROCNC_ACCIONES_TIPOS
        consulta = consulta & " WHERE 1 = 1 "
        consulta = consulta & strDescripcion
        consulta = consulta & strRESPONSABLE
        consulta = consulta & strFechaPrevista
        consulta = consulta & strFechaPuesta
        consulta = consulta & strFechaResolucion
        consulta = consulta & strEstado
        consulta = consulta & strTipo
        consulta = consulta & strCNC
        consulta = consulta & " order by procnc_accionescorrectivas.FECHA_PREVISTA"
        
                                      
        Set ListadoAcciones = datos_bd(consulta)

    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.ListadoAccionesPendientesRespCalidad"
    Exit Function
ListadoAccionesPendientesRespCalidad_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcAccionCorrectora.ListadoAccionesPendientesRespCalidad"
    MsgBox Err.Number & " (" & Err.Description & ") in procedure ListadoAccionesPendientesRespCalidad of Módulo de clase clsProcNcAccionCorrectora" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR, vbInformation
    G_TRAZABILIDAD_ERROR = ""
        
End Function
