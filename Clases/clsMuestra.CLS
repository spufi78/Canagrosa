VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMuestra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*****************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE MUESTRA      *
'*****************************************************
Private ID_MUESTRA As Long
Private TIPO_MUESTRA_ID As Long
Private ID_GENERAL As Long
Private ID_PARTICULAR As Long
Private ANNO As Long
Private TIPO_ANALISIS_ID As Long
Private ANALISIS_MODIFICADO As Integer
Private FECHA_MUESTREO As String
Private ENTIDAD_MUESTREO_ID As Long
Private DETALLE_MUESTREO As String
Private OBSERVACIONES_MUESTREO As String
Private FECHA_RECEPCION As String
Private HORA_RECEPCION As String
Private EMPLEADO_ID As Long
Private FORMATO_ID As Integer
Private ENTIDAD_ENTREGA_ID As Long
Private DETALLE_ENTREGA As String
Private OBSERVACIONES_ENTREGA As String
Private CLIENTE_ID As Long
Private CENTRO_ID As Long
Private REPLACEMENT_ID As Long
Private REFERENCIA_CLIENTE As String
Private PRECIO As String
Private PRECIO_FIJADO As Integer
Private FECHA_PREV_FIN As String
Private OBSERVACIONES As String
Private ANALISIS_DUPLICADO As Integer
Private ANULADA As Integer
Private precinto As String
Private BANO_ID As Long
Private IPA As Integer
'M1054-I
Private INFORME_MANUAL As Integer
'M1054-F
Private FECHA_RECOGIDA As String

Private FECHA_COMIENZO As String
Private FECHA_FINALIZACION As String
Private fecha_cierre As String

Private HORA_CIERRE As String
Private FECHA_ENVIO As String
Private MOTIVO_RETRASO_ID As Integer
Private RETRASO_OBSERVACION As String

Private CERRADA As Integer
Private revision_usuario As Integer
Private CERRADA_USUARIO As Integer
Private DOCUMENTO_PAGO As Integer
Private IMPRESA As Integer
Private enviado_correo As Integer
Private ULT_EDICION_IMP As Integer
Private OP_VUELO As Integer
Private OP_INSITU As Integer
Private OP_LABMOVIL As Integer
Private OP_NORUTINARIA As Integer
Private OP_REPETICION As Integer
Private situacion As Integer
Private firma As String
Private producto As String
Private RESPONSABLE_ID As Integer
Private PEDIDO_ID As Integer
'M1009-I
Private OFERTA_ID As Long
'M1009-F
Private PAQUETE_ID As Long
Private ENAC As Long
Private nadcap As Integer
Private urgente As Integer
Private ajuste As Integer
Private consulta As Integer
Private CONSULTA_OBSERVACIONES As String

Private AGRUPADA As Integer
Private AGRUPADA_MUESTRA_ID As Long

Private TITULO_MUESTRA As String

'*****************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE MUESTRA    *
'*****************************************************
Public Property Let setID_MUESTRA(ByVal dato As Long)
    ID_MUESTRA = dato
End Property
Public Property Let setTIPO_MUESTRA_ID(ByVal dato As Long)
    TIPO_MUESTRA_ID = dato
End Property
Public Property Let setID_GENERAL(ByVal dato As Long)
    ID_GENERAL = dato
End Property
Public Property Let setID_PARTICULAR(ByVal dato As Long)
    ID_PARTICULAR = dato
End Property
Public Property Let setTIPO_ANALISIS_ID(ByVal dato As Long)
    TIPO_ANALISIS_ID = dato
End Property
Public Property Let setANNO(ByVal dato As Long)
    ANNO = dato
End Property
Public Property Let setANALISIS_MODIFICADO(ByVal dato As Integer)
    ANALISIS_MODIFICADO = dato
End Property
Public Property Let setFECHA_MUESTREO(ByVal dato As String)
    FECHA_MUESTREO = Trim(dato)
End Property
Public Property Let setENTIDAD_MUESTREO_ID(ByVal dato As Long)
    ENTIDAD_MUESTREO_ID = dato
End Property
Public Property Let setDETALLE_MUESTREO(ByVal dato As String)
    DETALLE_MUESTREO = Trim(dato)
End Property
Public Property Let setOBSERVACIONES_MUESTREO(ByVal dato As String)
    OBSERVACIONES_MUESTREO = Trim(dato)
End Property
Public Property Let setFECHA_RECEPCION(ByVal dato As String)
    FECHA_RECEPCION = Trim(dato)
End Property
Public Property Let setHORA_RECEPCION(ByVal dato As String)
    HORA_RECEPCION = Trim(dato)
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Long)
    EMPLEADO_ID = dato
End Property
Public Property Let setFORMATO_ID(ByVal dato As Integer)
    FORMATO_ID = dato
End Property
Public Property Let setENTIDAD_ENTREGA_ID(ByVal dato As Long)
    ENTIDAD_ENTREGA_ID = dato
End Property
Public Property Let setDETALLE_ENTREGA(ByVal dato As String)
    DETALLE_ENTREGA = Trim(dato)
End Property
Public Property Let setOBSERVACIONES_ENTREGA(ByVal dato As String)
    OBSERVACIONES_ENTREGA = Trim(dato)
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
    CLIENTE_ID = dato
End Property
Public Property Let setCENTRO_ID(ByVal dato As Long)
    CENTRO_ID = dato
End Property
Public Property Let setREPLACEMENT_ID(ByVal dato As Long)
    REPLACEMENT_ID = dato
End Property
Public Property Let setREFERENCIA_CLIENTE(ByVal dato As String)
    REFERENCIA_CLIENTE = Trim(dato)
End Property
Public Property Let setPRECIO(ByVal dato As String)
    PRECIO = Trim(dato)
End Property
Public Property Let setPRECIO_FIJADO(ByVal dato As Integer)
    PRECIO_FIJADO = Trim(dato)
End Property
Public Property Let setFECHA_PREV_FIN(ByVal dato As String)
    FECHA_PREV_FIN = Trim(dato)
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
    OBSERVACIONES = Trim(dato)
End Property
Public Property Let setANALISIS_DUPLICADO(ByVal dato As Integer)
    ANALISIS_DUPLICADO = dato
End Property
Public Property Let setANULADA(ByVal dato As Integer)
    ANULADA = dato
End Property
Public Property Let setPRECINTO(ByVal dato As String)
    precinto = Trim(dato)
End Property
Public Property Let setBANO_ID(ByVal dato As Long)
    BANO_ID = dato
End Property
Public Property Let setIPA(ByVal dato As Integer)
    IPA = dato
End Property
'M1054-I
Public Property Let setINFORME_MANUAL(ByVal dato As Integer)
    INFORME_MANUAL = dato
End Property
'M1054-F
Public Property Let setFECHA_RECOGIDA(ByVal dato As String)
    FECHA_RECOGIDA = dato
End Property
Public Property Let setFECHA_COMIENZO(ByVal dato As String)
    FECHA_COMIENZO = Trim(dato)
End Property
Public Property Let setFECHA_FINALIZACION(ByVal dato As String)
    FECHA_FINALIZACION = Trim(dato)
End Property
Public Property Let setFECHA_CIERRE(ByVal dato As String)
    fecha_cierre = Trim(dato)
End Property
Public Property Let setHORA_CIERRE(ByVal dato As String)
    HORA_CIERRE = Trim(dato)
End Property
Public Property Let setFECHA_ENVIO(ByVal dato As String)
    FECHA_ENVIO = Trim(dato)
End Property
Public Property Let setMOTIVO_RETRASO_ID(ByVal dato As Integer)
    MOTIVO_RETRASO_ID = Trim(dato)
End Property
Public Property Let setRETRASO_OBSERVACION(ByVal dato As String)
    RETRASO_OBSERVACION = Trim(dato)
End Property
Public Property Let setCERRADA(ByVal dato As Integer)
    CERRADA = dato
End Property
Public Property Let setREVISION_USUARIO(ByVal dato As Integer)
    revision_usuario = dato
End Property
Public Property Let setCERRADA_USUARIO(ByVal dato As Integer)
    CERRADA_USUARIO = dato
End Property
Public Property Let setDOCUMENTO_PAGO(ByVal dato As Integer)
    DOCUMENTO_PAGO = dato
End Property
Public Property Let setIMPRESA(ByVal dato As Integer)
    IMPRESA = dato
End Property
Public Property Let setENVIADO_CORREO(ByVal dato As Integer)
    enviado_correo = dato
End Property
Public Property Let setULT_EDICION_IMP(ByVal dato As Integer)
    ULT_EDICION_IMP = dato
End Property
Public Property Let setOP_VUELO(ByVal dato As Integer)
    OP_VUELO = dato
End Property
Public Property Let setOP_INSITU(ByVal dato As Integer)
    OP_INSITU = dato
End Property
Public Property Let setOP_LABMOVIL(ByVal dato As Integer)
    OP_LABMOVIL = dato
End Property
Public Property Let setOP_NORUTINARIA(ByVal dato As Integer)
    OP_NORUTINARIA = dato
End Property
Public Property Let setOP_REPETICION(ByVal dato As Integer)
    OP_REPETICION = dato
End Property
Public Property Let setSITUACION(ByVal dato As Integer)
    situacion = dato
End Property
Public Property Let setFIRMA(ByVal dato As String)
    firma = dato
End Property
Public Property Let setPRODUCTO(ByVal dato As String)
    producto = dato
End Property
Public Property Let setRESPONSABLE_ID(ByVal dato As Integer)
    RESPONSABLE_ID = dato
End Property
Public Property Let setPEDIDO_ID(ByVal dato As Integer)
    PEDIDO_ID = dato
End Property
'M1009-I
Public Property Let setOFERTA_ID(ByVal dato As Long)
    OFERTA_ID = dato
End Property
'M1009-F
Public Property Let setPAQUETE_ID(ByVal dato As Long)
    PAQUETE_ID = dato
End Property
Public Property Let setENAC(ByVal dato As Long)
    ENAC = dato
End Property
Public Property Let setNADCAP(ByVal dato As Integer)
    nadcap = dato
End Property
Public Property Let setURGENTE(ByVal dato As Integer)
    urgente = dato
End Property
Public Property Let setAJUSTE(ByVal dato As Integer)
    ajuste = dato
End Property
Public Property Let setCONSULTA(ByVal dato As Integer)
    consulta = dato
End Property
Public Property Let setCONSULTA_OBSERVACIONES(ByVal dato As String)
    CONSULTA_OBSERVACIONES = dato
End Property
Public Property Let setAGRUPADA(ByVal dato As Integer)
    AGRUPADA = dato
End Property
Public Property Let setAGRUPADA_MUESTRA_ID(ByVal dato As Long)
    AGRUPADA_MUESTRA_ID = dato
End Property
'*****************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE MUESTRA      *
'*****************************************************

Public Property Get getID_MUESTRA() As Long
    getID_MUESTRA = ID_MUESTRA
End Property
Public Property Get getTIPO_MUESTRA_ID() As Long
    getTIPO_MUESTRA_ID = TIPO_MUESTRA_ID
End Property
Public Property Get getID_GENERAL() As Long
    getID_GENERAL = ID_GENERAL
End Property
Public Property Get getID_PARTICULAR() As Long
    getID_PARTICULAR = ID_PARTICULAR
End Property
Public Property Get getANNO() As Long
    getANNO = ANNO
End Property
Public Property Get getTIPO_ANALISIS_ID() As Long
    getTIPO_ANALISIS_ID = TIPO_ANALISIS_ID
End Property
Public Property Get getANALISIS_MODIFICADO() As Integer
    getANALISIS_MODIFICADO = ANALISIS_MODIFICADO
End Property
Public Property Get getFECHA_MUESTREO() As String
    getFECHA_MUESTREO = FECHA_MUESTREO
End Property
Public Property Get getENTIDAD_MUESTREO_ID() As Long
    getENTIDAD_MUESTREO_ID = ENTIDAD_MUESTREO_ID
End Property
Public Property Get getDETALLE_MUESTREO() As String
    getDETALLE_MUESTREO = DETALLE_MUESTREO
End Property
Public Property Get getOBSERVACIONES_MUESTREO() As String
    getOBSERVACIONES_MUESTREO = OBSERVACIONES_MUESTREO
End Property
Public Property Get getFECHA_RECEPCION() As String
    getFECHA_RECEPCION = FECHA_RECEPCION
End Property
Public Property Get getHORA_RECEPCION() As String
    getHORA_RECEPCION = HORA_RECEPCION
End Property
Public Property Get getEMPLEADO_ID() As Long
    getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getFORMATO_ID() As Integer
    getFORMATO_ID = FORMATO_ID
End Property
Public Property Get getENTIDAD_ENTREGA_ID() As Long
    getENTIDAD_ENTREGA_ID = ENTIDAD_ENTREGA_ID
End Property
Public Property Get getDETALLE_ENTREGA() As String
    getDETALLE_ENTREGA = DETALLE_ENTREGA
End Property
Public Property Get getOBSERVACIONES_ENTREGA() As String
    getOBSERVACIONES_ENTREGA = OBSERVACIONES_ENTREGA
End Property
Public Property Get getCLIENTE_ID() As Long
    getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getCENTRO_ID() As Long
    getCENTRO_ID = CENTRO_ID
End Property
Public Property Get getREPLACEMENT_ID() As Long
    getREPLACEMENT_ID = REPLACEMENT_ID
End Property
Public Property Get getREFERENCIA_CLIENTE() As String
    getREFERENCIA_CLIENTE = REFERENCIA_CLIENTE
End Property
Public Property Get getPRECIO() As String
    getPRECIO = PRECIO
End Property
Public Property Get getPRECIO_FIJADO() As Integer
    getPRECIO_FIJADO = PRECIO_FIJADO
End Property
Public Property Get getFECHA_PREV_FIN() As String
    getFECHA_PREV_FIN = FECHA_PREV_FIN
End Property
Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getANALISIS_DUPLICADO() As Integer
    getANALISIS_DUPLICADO = ANALISIS_DUPLICADO
End Property
Public Property Get getANULADA() As Integer
    getANULADA = ANULADA
End Property
Public Property Get getPRECINTO() As String
    getPRECINTO = precinto
End Property
Public Property Get getBANO_ID() As Long
    getBANO_ID = BANO_ID
End Property
Public Property Get getIPA() As Integer
    getIPA = IPA
End Property
'M1054-I
Public Property Get getINFORME_MANUAL() As Integer
    getINFORME_MANUAL = INFORME_MANUAL
End Property
'M1054-F
Public Property Get getFECHA_RECOGIDA() As String
    getFECHA_RECOGIDA = FECHA_RECOGIDA
End Property
Public Property Get getFECHA_COMIENZO() As String
    getFECHA_COMIENZO = FECHA_COMIENZO
End Property
Public Property Get getFECHA_FINALIZACION() As String
    getFECHA_FINALIZACION = FECHA_FINALIZACION
End Property
Public Property Get getFECHA_CIERRE() As String
    getFECHA_CIERRE = fecha_cierre
End Property
Public Property Get getHORA_CIERRE() As String
    getHORA_CIERRE = HORA_CIERRE
End Property
Public Property Get getFECHA_ENVIO() As String
    getFECHA_ENVIO = FECHA_ENVIO
End Property
Public Property Get getMOTIVO_RETRASO_ID() As Integer
    getMOTIVO_RETRASO_ID = MOTIVO_RETRASO_ID
End Property
Public Property Get getRETRASO_OBSERVACION() As String
    getRETRASO_OBSERVACION = RETRASO_OBSERVACION
End Property
Public Property Get getCERRADA() As Integer
    getCERRADA = CERRADA
End Property
Public Property Get getREVISION_USUARIO() As Integer
    getREVISION_USUARIO = revision_usuario
End Property
Public Property Get getCERRADA_USUARIO() As Integer
    getCERRADA_USUARIO = CERRADA_USUARIO
End Property
Public Property Get getDOCUMENTO_PAGO() As Integer
    getDOCUMENTO_PAGO = DOCUMENTO_PAGO
End Property
Public Property Get getIMPRESA() As Integer
    getIMPRESA = IMPRESA
End Property
Public Property Get getENVIADO_CORREO() As Integer
    getENVIADO_CORREO = enviado_correo
End Property
Public Property Get getULT_EDICION_IMP() As Integer
    getULT_EDICION_IMP = ULT_EDICION_IMP
End Property
Public Property Get getOP_VUELO() As Integer
    getOP_VUELO = OP_VUELO
End Property
Public Property Get getOP_INSITU() As Integer
    getOP_INSITU = OP_INSITU
End Property
Public Property Get getOP_LABMOVIL() As Integer
    getOP_LABMOVIL = OP_LABMOVIL
End Property
Public Property Get getOP_NORUTINARIA() As Integer
    getOP_NORUTINARIA = OP_NORUTINARIA
End Property
Public Property Get getOP_REPETICION() As Integer
    getOP_REPETICION = OP_REPETICION
End Property
Public Property Get getSITUACION() As Integer
    getSITUACION = situacion
End Property
Public Property Get getFIRMA() As String
    getFIRMA = firma
End Property
Public Property Get getPRODUCTO() As String
    getPRODUCTO = producto
End Property
Public Property Get getRESPONSABLE_ID() As Integer
    getRESPONSABLE_ID = RESPONSABLE_ID
End Property
Public Property Get getPEDIDO_ID() As Integer
    getPEDIDO_ID = PEDIDO_ID
End Property
'M1009-I
Public Property Get getOFERTA_ID() As Long
    getOFERTA_ID = OFERTA_ID
End Property
'M1009-F
Public Property Get getPAQUETE_ID() As Long
    getPAQUETE_ID = PAQUETE_ID
End Property
Public Property Get getENAC() As Long
    getENAC = ENAC
End Property
Public Property Get getNADCAP() As Integer
    getNADCAP = nadcap
End Property
Public Property Get getURGENTE() As Integer
    getURGENTE = urgente
End Property
Public Property Get getAJUSTE() As Integer
    getAJUSTE = ajuste
End Property
Public Property Get getCONSULTA() As Integer
    getCONSULTA = consulta
End Property
Public Property Get getCONSULTA_OBSERVACIONES() As String
    getCONSULTA_OBSERVACIONES = CONSULTA_OBSERVACIONES
End Property
Public Property Get getAGRUPADA() As Integer
    getAGRUPADA = AGRUPADA
End Property
Public Property Get getAGRUPADA_MUESTRA_ID() As Long
    getAGRUPADA_MUESTRA_ID = AGRUPADA_MUESTRA_ID
End Property
Public Property Get getTITULO_MUESTRA() As String
    getTITULO_MUESTRA = TITULO_MUESTRA
End Property

'*****************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE MUESTRA  *
'*****************************************************
Public Function esBano(ID_MUESTRA As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
'    consulta = "SELECT tipo_especial_id,codigo FROM tipos_muestra" & _
'               " WHERE id_tipo_muestra=" & ID_MUESTRA
    consulta = "SELECT bano FROM tipos_muestra a, tipos_especial b WHERE a.tipo_especial_id = b.id_tipo_especial and a.id_tipo_muestra=" & ID_MUESTRA
    esBano = False
    If datos_bd(consulta).Fields(0) <> 0 Then
       esBano = True
    End If
    Exit Function
fallo:
    MsgBox "Error al determinar el tipo de muestra (esBano)", vbCritical, Err.Description
End Function
Private Sub CrearIdCodigoGeneral()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    On Error GoTo fallo
    ' Recuperamos el maximo secuencial de muestra
    consulta = "SELECT MAX(ID_GENERAL) FROM muestras " & _
               " WHERE ANNO = " & Year(Me.getFECHA_RECEPCION)
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_GENERAL = 1
    Else
        ID_GENERAL = rs.Fields(0) + 1
    End If
    Set rs = Nothing
    Exit Sub
fallo:
    MsgBox "Error en CrearIdCodigoGeneral : " & Err.Description, vbCritical, App.Title
End Sub
Public Sub CrearIdCodigoParticular(muestra As Long)
    Dim otm As New clsTipos_muestra
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    On Error GoTo fallo
    otm.CARGAR (muestra)
    ' Recuperamos el maximo secuencial de muestra
    consulta = "SELECT MAX(ID_PARTICULAR) FROM muestras a, tipos_muestra b " & _
               " WHERE a.tipo_muestra_id=b.id_tipo_muestra " & _
               "   AND b.codigo = '" & UCase(otm.getCODIGO) & "'" & _
               "   and a.ANNO = " & Year(Me.getFECHA_RECEPCION)
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_PARTICULAR = 1
    Else
        ID_PARTICULAR = rs.Fields(0) + 1
    End If
    Set rs = Nothing
    Exit Sub
fallo:
    MsgBox "Error en CrearIdCodigoParticular : " & Err.Description, vbCritical, App.Title
End Sub

Public Sub CrearIdMuestra()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_MUESTRA) FROM muestras"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_MUESTRA = 1
    Else
        ID_MUESTRA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Private Sub RECUPERAR_RESPONSABLE_MUESTRA()
    Dim otm As New clsTipos_muestra
    otm.CARGAR Me.getTIPO_MUESTRA_ID
    RESPONSABLE_ID = otm.getRESPONSABLE_ID
End Sub
Public Function guardarMuestra() As Long
    Dim sql As String
    On Error GoTo fallo:
    Me.setANNO = Year(Me.getFECHA_RECEPCION)
    Me.CrearIdCodigoParticular (TIPO_MUESTRA_ID)
    If ANALISIS_DUPLICADO = 0 Then
        If Me.getID_PARTICULAR Mod 10 = 0 Then
            ANALISIS_DUPLICADO = 1
        Else
            ANALISIS_DUPLICADO = 0
        End If
    End If
    ' Si es un sellante, nunca por duplicado (Mari Cruz 14/07/2015)
    If TIPO_MUESTRA_ID = TIPOS_MUESTRAS.SELLANTE Then
        ANALISIS_DUPLICADO = 0
    End If
    If TIPO_MUESTRA_ID = TIPOS_MUESTRAS.TM_AGUA Then
        ANALISIS_DUPLICADO = 1
    End If
    
    HORA_RECEPCION = Format(Time, "hh:mm:ss")
    RECUPERAR_RESPONSABLE_MUESTRA
    CrearIdCodigoGeneral
    Me.CrearIdMuestra
    'M1054 : INCLUIR INFORME_MANUAL
    'M1009 : OFERTA_ID
    'M1105 : OP_VUELO, OP_INSITU, OP_LABMOVIL, OP_NORUTINARIA, FECHA_RECOGIDA
    Dim frecogida As String
    If FECHA_RECOGIDA = "" Then
        frecogida = "NULL"
    Else
        frecogida = "'" & FECHA_RECOGIDA & "'"
    End If
    ' ENAC
    If ANALISIS_MODIFICADO <> 2 Then
        Dim otm As New clsTipos_muestra
        otm.CARGAR TIPO_MUESTRA_ID
        ENAC = otm.getENAC
        nadcap = otm.getNADCAP
    End If
    If CENTRO_ID = 0 Then
        CENTRO_ID = 1
    End If
    If firma = "" Then
        firma = "0"
    End If
    ajuste = 0
'    CONSULTA = 0
'    CONSULTA_OBSERVACIONES = ""
    sql = "INSERT INTO muestras(ID_MUESTRA,TIPO_MUESTRA_ID,ID_GENERAL,ID_PARTICULAR,ANNO," & _
               "TIPO_ANALISIS_ID,ANALISIS_MODIFICADO,FECHA_MUESTREO," & _
               "ENTIDAD_MUESTREO_ID,DETALLE_MUESTREO,OBSERVACIONES_MUESTREO," & _
               "FECHA_RECEPCION,HORA_RECEPCION,EMPLEADO_ID,FORMATO_ID," & _
               "ENTIDAD_ENTREGA_ID,DETALLE_ENTREGA,OBSERVACIONES_ENTREGA," & _
               "CLIENTE_ID,CENTRO_ID,REPLACEMENT_ID,REFERENCIA_CLIENTE,PRECIO," & _
               "FECHA_PREV_FIN,OBSERVACIONES,ANALISIS_DUPLICADO,ANULADA," & _
               "PRECINTO,BANO_ID,IPA,INFORME_MANUAL,FECHA_COMIENZO,FECHA_FINALIZACION,FECHA_CIERRE,CERRADA,REVISION_USUARIO," & _
               "DOCUMENTO_PAGO,IMPRESA,ENVIADO_CORREO,ULT_EDICION_IMP,FIRMA,PRODUCTO,PEDIDO_ID,OFERTA_ID,PAQUETE_ID,ENAC,NADCAP,URGENTE,AJUSTE,RESPONSABLE_ID,CERRADA_USUARIO," & _
               "OP_VUELO, OP_INSITU, OP_LABMOVIL, OP_NORUTINARIA,OP_REPETICION, FECHA_RECOGIDA)" & _
               "values(" & _
               ID_MUESTRA & "," & TIPO_MUESTRA_ID & "," & ID_GENERAL & "," & ID_PARTICULAR & "," & ANNO & "," & TIPO_ANALISIS_ID & "," & ANALISIS_MODIFICADO & ",'" & FECHA_MUESTREO & "'," & _
               ENTIDAD_MUESTREO_ID & ",'" & DETALLE_MUESTREO & "','" & OBSERVACIONES_MUESTREO & "','" & FECHA_RECEPCION & "',CURRENT_TIME," & EMPLEADO_ID & "," & FORMATO_ID & "," & _
               ENTIDAD_ENTREGA_ID & ",'" & DETALLE_ENTREGA & "','" & OBSERVACIONES_ENTREGA & "'," & CLIENTE_ID & "," & CENTRO_ID & "," & REPLACEMENT_ID & ",'" & REFERENCIA_CLIENTE & "'," & moneda_bd(PRECIO) & ",'" & _
               FECHA_PREV_FIN & "','" & OBSERVACIONES & "'," & ANALISIS_DUPLICADO & "," & ANULADA & ",'" & precinto & "'," & BANO_ID & "," & IPA & "," & INFORME_MANUAL & ",'" & FECHA_COMIENZO & "','" & FECHA_FINALIZACION & "','" & fecha_cierre & "'," & CERRADA & "," & revision_usuario & "," & _
               DOCUMENTO_PAGO & "," & IMPRESA & "," & enviado_correo & "," & ULT_EDICION_IMP & ",'" & firma & "','" & producto & "'," & PEDIDO_ID & "," & OFERTA_ID & "," & PAQUETE_ID & "," & ENAC & "," & nadcap & "," & urgente & "," & ajuste & "," & RESPONSABLE_ID & ",1," & _
               OP_VUELO & "," & OP_INSITU & "," & OP_LABMOVIL & "," & OP_NORUTINARIA & "," & OP_REPETICION & "," & frecogida & ")"
    execute_bd sql
    ANALISIS_DUPLICADO = 0
    guardarMuestra = ID_MUESTRA
    Exit Function
fallo:
    guardarMuestra = 0
    MsgBox "Error al guardar la muestra en la bd : " & Err.Description
End Function

Public Function CodigoParticular(muestra As Long) As String
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "select CONCAT(tp.codigo,'-',CAST(Me.id_particular AS CHAR)) AS CODIGO" & _
               "  from muestras Me, tipos_muestra tp" & _
               " Where Me.tipo_muestra_id = tp.id_tipo_muestra" & _
               "   and Me.id_muestra=" & muestra
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then 'si es nulo No se recupero ninguno
       CodigoParticular = ""
    Else
       CodigoParticular = rs("CODIGO")
    End If
    Set rs = Nothing
    Exit Function
fallo:
    MsgBox "Error al recuperar el codigo particular de la muestra", vbCritical, Err.Description
End Function
Public Function ProximoCodigo(TIPO_ANALISIS As Long) As String
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "select CONCAT(tp.codigo,'-',CAST(Me.id_particular + 1 AS CHAR)) AS CODIGO" & _
               "  from muestras Me, tipos_muestra tp,tipos_analisis ta " & _
               " Where Me.tipo_muestra_id = tp.id_tipo_muestra " & _
               "   and ta.tipo_muestra_id=tp.id_tipo_muestra " & _
               "   and ta.id_tipo_analisis=" & TIPO_ANALISIS & _
               "   and Me.anno = " & Year(Date) & _
               " order by Me.id_particular desc "
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then 'si es nulo No se recupero ninguno
       ProximoCodigo = ""
    Else
       ProximoCodigo = rs.Fields(0)
    End If
    Set rs = Nothing
    Exit Function
fallo:
    MsgBox "Error al recuperar el codigo particular de la muestra", vbCritical, Err.Description
End Function
Public Function CargaCodigoBano(N_ensayo As Long) As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT BANO_ID FROM muestras WHERE id_muestra=" & N_ensayo
    Set rs = datos_bd(consulta)
        
    If IsNull(rs.Fields(0)) Or (rs.RecordCount = 0) Then
        CargaCodigoBano = False
    Else
        CargaCodigoBano = True
        BANO_ID = rs.Fields("BANO_ID")
    End If
End Function
Public Function CargaAnno(idMuestra As Long) As Integer
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ANNO FROM muestras WHERE id_muestra=" & idMuestra
    Set rs = datos_bd(consulta)
        
    If IsNull(rs.Fields(0)) Or (rs.RecordCount = 0) Then
        CargaAnno = 0
    Else
        CargaAnno = rs.Fields("ANNO")
    End If
End Function

Public Function CargaMuestra(N_ensayo As Long) As Boolean
    On Error GoTo fallo:
    Dim rs As New ADODB.Recordset
    Dim sql As String
    sql = "SELECT * FROM muestras WHERE id_muestra=" & N_ensayo
    Set rs = datos_bd(sql)
        
    If IsNull(rs.Fields(0)) Or (rs.RecordCount = 0) Then
        CargaMuestra = False
    Else
        ID_MUESTRA = rs.Fields("ID_MUESTRA")
        TIPO_MUESTRA_ID = rs.Fields("TIPO_MUESTRA_ID")
        ID_GENERAL = rs.Fields("ID_GENERAL")
        ID_PARTICULAR = rs.Fields("ID_PARTICULAR")
        ANNO = rs.Fields("ANNO")
        ANALISIS_MODIFICADO = rs.Fields("ANALISIS_MODIFICADO") & ""
        TIPO_ANALISIS_ID = rs.Fields("TIPO_ANALISIS_ID") & ""
        ANALISIS_DUPLICADO = rs.Fields("ANALISIS_DUPLICADO") & ""
        If IsNull(rs.Fields("ANULADA")) Then
            ANULADA = 0
        Else
            ANULADA = rs.Fields("ANULADA")
        End If
        ULT_EDICION_IMP = rs.Fields("ULT_EDICION_IMP") & ""
        ENTIDAD_MUESTREO_ID = rs.Fields("ENTIDAD_MUESTREO_ID") & ""
        EMPLEADO_ID = rs.Fields("EMPLEADO_ID") & ""
        FORMATO_ID = rs.Fields("FORMATO_ID") & ""
        ENTIDAD_ENTREGA_ID = rs.Fields("ENTIDAD_ENTREGA_ID") & ""
        CLIENTE_ID = rs.Fields("CLIENTE_ID") & ""
        CENTRO_ID = rs.Fields("CENTRO_ID") & ""
        REPLACEMENT_ID = rs.Fields("REPLACEMENT_ID") & ""
        If IsNull(rs.Fields("BANO_ID")) Then
            BANO_ID = 0
        Else
            BANO_ID = rs.Fields("BANO_ID") & ""
        End If
        IPA = rs("IPA")
        'M1054-I
        INFORME_MANUAL = rs("INFORME_MANUAL")
        'M1054-F
        If Not IsNull(rs("FECHA_RECOGIDA")) Then
            FECHA_RECOGIDA = rs("FECHA_RECOGIDA")
        Else
            FECHA_RECOGIDA = ""
        End If
        If IsNull(rs.Fields("CERRADA")) Then
            CERRADA = 0
        Else
            CERRADA = rs.Fields("CERRADA")
        End If
        revision_usuario = rs.Fields("REVISION_USUARIO")
        CERRADA_USUARIO = rs("CERRADA_USUARIO")
        DOCUMENTO_PAGO = rs.Fields("DOCUMENTO_PAGO") & ""
        DETALLE_MUESTREO = rs.Fields("DETALLE_MUESTREO") & ""
        OBSERVACIONES_MUESTREO = rs.Fields("OBSERVACIONES_MUESTREO") & ""
        DETALLE_ENTREGA = rs.Fields("DETALLE_ENTREGA") & ""
        OBSERVACIONES_ENTREGA = rs.Fields("OBSERVACIONES_ENTREGA") & ""
        OBSERVACIONES = rs.Fields("OBSERVACIONES") & ""
        precinto = rs.Fields("PRECINTO") & ""
        REFERENCIA_CLIENTE = rs.Fields("REFERENCIA_CLIENTE") & ""
        FECHA_MUESTREO = rs.Fields("FECHA_MUESTREO") & ""
        FECHA_RECEPCION = rs.Fields("FECHA_RECEPCION") & ""
        HORA_RECEPCION = rs.Fields("HORA_RECEPCION") & ""
        FECHA_PREV_FIN = rs.Fields("FECHA_PREV_FIN") & ""
        FECHA_COMIENZO = rs.Fields("FECHA_COMIENZO") & ""
        FECHA_FINALIZACION = rs.Fields("FECHA_FINALIZACION") & ""
        fecha_cierre = rs.Fields("FECHA_CIERRE") & ""
        HORA_CIERRE = rs("HORA_CIERRE")
        If Not IsNull(rs("FECHA_ENVIO")) Then
            FECHA_ENVIO = rs("FECHA_ENVIO")
        Else
            FECHA_ENVIO = ""
        End If
        If Not IsNull(rs("RETRASO_OBSERVACION")) Then
            RETRASO_OBSERVACION = rs("RETRASO_OBSERVACION")
        Else
            RETRASO_OBSERVACION = ""
        End If
        MOTIVO_RETRASO_ID = rs("MOTIVO_RETRASO_ID")
        PRECIO = rs.Fields("PRECIO")
        PRECIO_FIJADO = rs.Fields("PRECIO_FIJADO")
        IMPRESA = rs.Fields("IMPRESA")
        enviado_correo = rs.Fields("ENVIADO_CORREO")
        firma = rs("FIRMA")
        producto = rs("PRODUCTO")
        RESPONSABLE_ID = rs("RESPONSABLE_ID")
        PEDIDO_ID = rs("PEDIDO_ID")
        'M1009-I
        OFERTA_ID = rs("OFERTA_ID")
        'M1009-F
        PAQUETE_ID = rs("PAQUETE_ID")
        ENAC = rs("ENAC")
        nadcap = rs("NADCAP")
        urgente = rs("URGENTE")
        ajuste = rs("AJUSTE")
        consulta = rs("CONSULTA")
        If Not IsNull(rs("CONSULTA_OBSERVACIONES")) Then
            CONSULTA_OBSERVACIONES = rs("CONSULTA_OBSERVACIONES")
        End If
        OP_VUELO = rs("OP_VUELO")
        OP_INSITU = rs("OP_INSITU")
        OP_LABMOVIL = rs("OP_LABMOVIL")
        OP_NORUTINARIA = rs("OP_NORUTINARIA")
        OP_REPETICION = rs("OP_REPETICION")
        situacion = rs("SITUACION")
        
        AGRUPADA = rs("AGRUPADA")
        AGRUPADA_MUESTRA_ID = rs("AGRUPADA_MUESTRA_ID")
        
        TITULO_MUESTRA = ID_GENERAL & "/" & ANNO & " (" & Me.CodigoParticular(N_ensayo) & ") Edición : " & IIf(CERRADA = 0, ULT_EDICION_IMP + 1, ULT_EDICION_IMP)
        CargaMuestra = True
    End If
    Set rs = Nothing
    Exit Function
fallo:
    CargaMuestra = False
    MsgBox "Error al recuperar la muestra (clsMuestra.CargaMuestra)", vbCritical, Err.Description
End Function
Public Function NombreMuestra(TIPO_MUESTRA As Integer) As String
     Dim consulta As String
     consulta = "SELECT nombre FROM tipos_muestra WHERE id_tipo_muestra=" & TIPO_MUESTRA
     NombreMuestra = datos_bd(consulta).Fields("nombre")
End Function
Public Function Informar_Documento_Pago(muestra As Long, tipo_doc As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE muestras SET documento_pago=" & tipo_doc & _
               " WHERE id_muestra=" & muestra
    execute_bd consulta
    Informar_Documento_Pago = True
    Exit Function
fallo:
    Informar_Documento_Pago = False
    MsgBox "Error al modificar el documento de pago de la muestra", vbCritical, Err.Description
End Function
Public Function Informar_Cliente(muestra As Long, cliente As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE MUESTRAS SET CLIENTE_ID=" & cliente & _
               " WHERE ID_MUESTRA=" & muestra
    execute_bd consulta
    Informar_Cliente = True
    Exit Function
fallo:
    Informar_Cliente = False
    MsgBox "Error al modificar el cliente de la muestra.", vbCritical, Err.Description
End Function
Public Function informar_pedido(muestra As Long, pedido As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE MUESTRAS SET PEDIDO_ID=" & pedido & _
               " WHERE ID_MUESTRA=" & muestra
    execute_bd consulta
    informar_pedido = True
    Exit Function
fallo:
    informar_pedido = False
    MsgBox "Error al modificar el pedido de la muestra.", vbCritical, Err.Description
End Function

Public Function Informar_Documentos_Pago(grupo As String, tipo_doc As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE muestras SET documento_pago=" & tipo_doc & _
               " WHERE id_muestra in (" & grupo & ")"
    execute_bd consulta
    Informar_Documentos_Pago = True
    Exit Function
fallo:
    Informar_Documentos_Pago = False
    MsgBox "Error al modificar los documentos de pago de la muestra", vbCritical, Err.Description
End Function
Public Function Informar_Documento_Pago_general(GENERAL As Long, tipo_doc As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE muestras SET documento_pago=" & tipo_doc & _
               " WHERE id_general=" & GENERAL & " and anno = " & Year(Date)
    execute_bd consulta
    Informar_Documento_Pago_general = True
    Exit Function
fallo:
    Informar_Documento_Pago_general = False
    MsgBox "Error al modificar el documento de pago de la muestra", vbCritical, Err.Description
End Function

Public Function Modificar_Datos_Adicionales(codmuestra As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
'    consulta = "UPDATE muestras SET" & _
'               " REFERENCIA_CLIENTE='" & REFERENCIA_CLIENTE & "'," & _
'               " PRECINTO='" & precinto & "'," & _
'               " PRECIO='" & PRECIO & "'," & _
'               " FIRMA='" & firma & "'," & _
'               " ANALISIS_DUPLICADO=" & ANALISIS_DUPLICADO & _
'               " WHERE id_muestra=" & codmuestra
    If firma = "" Then
        firma = "0"
    End If
    consulta = "UPDATE muestras SET" & _
               " REFERENCIA_CLIENTE='" & REFERENCIA_CLIENTE & "'," & _
               " PRECINTO='" & precinto & "'," & _
               " FIRMA='" & firma & "'," & _
               " ANALISIS_DUPLICADO=" & ANALISIS_DUPLICADO & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta
    Modificar_Datos_Adicionales = True
    Exit Function
fallo:
    Modificar_Datos_Adicionales = False
    MsgBox "Error al modificar los datos adicionales de la muestra", vbCritical, Err.Description
End Function
Public Function Modificar_Otros_datos(codmuestra As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    Dim fEnvio As String
    If FECHA_ENVIO = "" Then
        fEnvio = " FECHA_ENVIO = NULL,"
    Else
        fEnvio = " FECHA_ENVIO = '" & FECHA_ENVIO & "',"
    
    End If
    consulta = "UPDATE muestras SET" & _
               "  ULT_EDICION_IMP=" & ULT_EDICION_IMP & "," & _
               "  FECHA_COMIENZO='" & Format(FECHA_COMIENZO, "YYYY-MM-DD") & "'," & _
               "  FECHA_FINALIZACION='" & Format(FECHA_FINALIZACION, "YYYY-MM-DD") & "'," & _
               "  FECHA_CIERRE='" & Format(fecha_cierre, "YYYY-MM-DD") & "'," & _
               "  ANULADA=" & ANULADA & "," & _
               "  INFORME_MANUAL=" & INFORME_MANUAL & "," & _
               "  HORA_CIERRE='" & HORA_CIERRE & "'," & _
               fEnvio & _
               "  CERRADA_USUARIO=" & CERRADA_USUARIO & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta
    Modificar_Otros_datos = True
    Exit Function
fallo:
    Modificar_Otros_datos = False
    MsgBox "Error al modificar los datos adicionales de la muestra", vbCritical, Err.Description
End Function

Public Function Modificar_datos_adicionales_bano(codmuestra As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    'M1081-I
    ' RECALCULO FECHA PREVISTA DE ENTREGA
    Dim oMuestra As New clsMuestra
    Dim oTA As New clsTipos_analisis
    oMuestra.CargaMuestra codmuestra
    oTA.CARGAR TIPO_ANALISIS_ID
    Dim fecha As Date
    If IsNumeric(oTA.getDIAS_TRABAJO) Then
        fecha = calcularFechaFinalizacion(oMuestra.getFECHA_RECEPCION, oTA.getDIAS_TRABAJO)
    Else
        fecha = oMuestra.getFECHA_RECEPCION
    End If
    'M1081-F
    consulta = "UPDATE muestras SET" & _
               " TIPO_ANALISIS_ID=" & TIPO_ANALISIS_ID & "," & _
               " REFERENCIA_CLIENTE='" & REFERENCIA_CLIENTE & "'," & _
               " BANO_ID=" & BANO_ID & "," & _
               " ANALISIS_DUPLICADO=" & ANALISIS_DUPLICADO & "," & _
               " FECHA_PREV_FIN = '" & Format(fecha, "yyyy-mm-dd") & "'" & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta
    Modificar_datos_adicionales_bano = True
    Exit Function
fallo:
    Modificar_datos_adicionales_bano = False
    MsgBox "Error al modificar los datos adicionales del baño.", vbCritical, Err.Description
End Function
Public Function Anular(codmuestra As Long, oB As String) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE muestras SET" & _
               " OBSERVACIONES = concat(OBSERVACIONES,'" & "- Anulada por : " & USUARIO.getNOMBRE & ". Motivo: " & Trim(oB) & "')," & _
               " ANULADA=1" & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta
    consulta = "INSERT into muestras_anuladas VALUES(" & _
            codmuestra & ",'" & Trim(oB) & "','" & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & "','" & _
            Format(Date, "yyyy-mm-dd") & "')"
    execute_bd consulta
    Anular = True
    Exit Function
fallo:
    Anular = False
    MsgBox "Error al Anular la muestra.", vbCritical, Err.Description
End Function
'Public Function Nueva_Edicion(codmuestra As Long, EDICION As Integer, oB As String, fecha As String) As Boolean
'    On Error GoTo fallo
'    Dim consulta As String
''    execute_bd "DELETE FROM MUESTRAS_NUEVA_EDICION WHERE MUESTRA_ID = " & codmuestra & " AND EDICION = " & EDICION
'    consulta = "INSERT into muestras_nueva_edicion VALUES(" & _
'            codmuestra & "," & EDICION & ",'" & Trim(oB) & "','" & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & "','" & _
'            Format(fecha, "yyyy-mm-dd") & "')"
'    execute_bd consulta
'    Nueva_Edicion = True
'    Exit Function
'fallo:
'    Nueva_Edicion = False
'    MsgBox "Error al generar el motivo de nueva edición.", vbCritical, Err.Description
'End Function
Public Sub actualizar_fecha_cierre(codmuestra As Long)
    Dim consulta As String
    consulta = "UPDATE muestras SET" & _
               " CERRADA=1,FECHA_CIERRE ='" & Format(Date, "yyyy-mm-dd") & "'" & _
               " ,HORA_CIERRE ='" & Format(Time, "hh:mm:ss") & "'" & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta
End Sub
Public Sub actualizarPaqueteId(codmuestra As Long, PAQUETE_ID As Long)
    Dim consulta As String
    consulta = "UPDATE muestras " & _
               "   SET PAQUETE_ID = " & PAQUETE_ID & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta
End Sub
Public Sub actualizar_precio(codmuestra As Long, PRECIO As String)
    Dim consulta As String
    consulta = "UPDATE muestras SET" & _
               " PRECIO='" & PRECIO & "'" & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta
End Sub
Private Sub actualizar_fecha_plasma(codmuestra As Long)
    Dim consulta As String
   On Error GoTo actualizar_fecha_plasma_Error

    consulta = "UPDATE muestras SET" & _
               "  FECHA_COMIENZO = FECHA_RECEPCION" & _
               " ,FECHA_FINALIZACION = CURRENT_DATE " & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta

   On Error GoTo 0
   Exit Sub

actualizar_fecha_plasma_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure actualizar_fecha_plasma of Módulo de clase clsMuestra"
End Sub
Private Sub actualizar_fecha_comienzo(codmuestra As Long, tabla As String)
    Dim consulta As String
    Dim rs As ADODB.Recordset
    Dim fecha As Date
   On Error GoTo actualizar_fecha_comienzo_Error

    If UCase(tabla) = "DETERMINACIONES" Then
        consulta = "select min(b.fecha) " & _
                   "  from determinaciones a, determinaciones_historico b " & _
                   " Where a.MUESTRA_ID = " & codmuestra & _
                   "   and a.ID_DETERMINACION = b.DETERMINACION_ID"
    Else
        consulta = "SELECT MIN(FECHA) FROM " & tabla & " WHERE MUESTRA_ID = " & codmuestra
    End If
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        fecha = Format(rs(0), "yyyy-mm-dd")
    Else
        fecha = Format(Date, "yyyy-mm-dd")
    End If
    consulta = "UPDATE muestras SET" & _
               " FECHA_COMIENZO ='" & Format(fecha, "yyyy-mm-dd") & "'" & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta

   On Error GoTo 0
   Exit Sub

actualizar_fecha_comienzo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure actualizar_fecha_comienzo of Módulo de clase clsMuestra"
End Sub
Private Sub actualizar_fecha_fin(codmuestra As Long, tabla As String)
    Dim consulta As String
    Dim rs As ADODB.Recordset
    Dim fecha As Date

   On Error GoTo actualizar_fecha_fin_Error

    If UCase(tabla) = "DETERMINACIONES" Then
        consulta = "select max(b.fecha) " & _
                   "  from determinaciones a, determinaciones_historico b " & _
                   " Where a.MUESTRA_ID = " & codmuestra & _
                   "   and a.ID_DETERMINACION = b.DETERMINACION_ID"
    Else
        consulta = "SELECT MAX(FECHA) FROM " & tabla & " WHERE MUESTRA_ID = " & codmuestra
    End If
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        fecha = Format(rs(0), "yyyy-mm-dd")
    Else
        fecha = Format(Date, "yyyy-mm-dd")
    End If
    consulta = "UPDATE muestras SET" & _
               " FECHA_FINALIZACION ='" & Format(fecha, "yyyy-mm-dd") & "'" & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta

   On Error GoTo 0
   Exit Sub

actualizar_fecha_fin_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure actualizar_fecha_fin of Módulo de clase clsMuestra"

End Sub

Public Function Cerrar(codmuestra As Long) As Boolean
    On Error GoTo fallo
    ' Cerramos la muestra
    Dim consulta As String
    ' Se limpia el ENVIADO_CORREO para las ediciones 2
    consulta = "UPDATE MUESTRAS SET " & _
               " ENVIADO_CORREO=0," & _
               " CERRADA=1," & _
               " FECHA_CIERRE ='" & Format(Date, "yyyy-mm-dd") & "'," & _
               " HORA_CIERRE ='" & Format(Time, "hh:mm:ss") & "'," & _
               " CERRADA_USUARIO = " & USUARIO.getID_EMPLEADO & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta
    ' Generacion de informe
    Dim generar As Boolean
    generar = True
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra (codmuestra)
    ' Solo calculamos fecha de comienzo y finalizacion si es primera edicion
'    If oMuestra.getULT_EDICION_IMP = 0 Then
'        consulta = "UPDATE MUESTRAS SET" & _
'                   " FECHA_FINALIZACION ='" & Format(Date, "yyyy-mm-dd") & "' " & _
'                   " WHERE id_muestra=" & codmuestra
'        execute_bd consulta
        ' Recalculamos el precio de la muestra
        If oMuestra.getPRECIO_FIJADO = 0 Then
            oMuestra.informar_precio_muestra codmuestra
        End If
        ' Si es un agua, validamos que para generar el informe,
        ' todas sus muestras esten cerradas
        Select Case oMuestra.getANALISIS_MODIFICADO
            Case 2 ' CONTROL DE EFICACIA
                Dim fecha_ce As String
                fecha_ce = Format(oMuestra.getFECHA_RECEPCION, "yyyy-mm-dd")
                Dim oce_recepcion As New clsCe_recepcion
                With oce_recepcion
                 If .Carga(codmuestra) Then
                    If .getDURACION_FECHA_DESDE <> "" And InStr(1, .getDURACION_FECHA_DESDE, "1900") < 1 Then
                        fecha_ce = Format(.getDURACION_FECHA_DESDE, "yyyy-mm-dd")
                        consulta = "UPDATE muestras SET" & _
                                   " FECHA_COMIENZO ='" & fecha_ce & "'" & _
                                   " WHERE id_muestra=" & codmuestra
                        execute_bd consulta
                        If IsDate(.getDURACION_FECHA_HASTA) Then
                            fecha_ce = Format(.getDURACION_FECHA_HASTA, "yyyy-mm-dd")
                            consulta = "UPDATE muestras SET" & _
                                       " FECHA_FINALIZACION ='" & fecha_ce & "'" & _
                                       " WHERE id_muestra=" & codmuestra
                            execute_bd consulta
                        End If
                    Else
                        actualizar_fecha_comienzo codmuestra, "CE_RESULTADOS"
                        actualizar_fecha_fin codmuestra, "CE_RESULTADOS"
                    End If
                 End If
                End With
                ' SPDA : Modificación 07/11/2019. Si no tiene reactivo propio SPDA no debe actualizar
                If oMuestra.getULT_EDICION_IMP = 0 And Trim(oce_recepcion.getREACTIVOS_PROPIOS) <> "" Then
                    Dim op As New clsParametros
                    Dim total As Single
                    total = 0
                    op.Carga parametros.SPDA_TIPO_MUESTRA, ""
                    If oMuestra.getTIPO_MUESTRA_ID = CInt(op.getVALOR) Then
                        Dim rsSPDA As ADODB.Recordset
                        Set rsSPDA = datos_bd("SELECT RESULTADO FROM CE_RESULTADOS WHERE MUESTRA_ID = " & codmuestra)
                        If rsSPDA.RecordCount > 0 Then
                            If IsNumeric(rsSPDA(0)) Then
                                Do
                                    total = total + Replace(rsSPDA(0), ".", ",")
                                    rsSPDA.MoveNext
                                Loop Until rsSPDA.EOF
                                op.Carga parametros.SPDA_CANTIDAD, ""
                                op.setVALOR = CSng(op.getVALOR) + total
                                op.actualizar_valor parametros.SPDA_CANTIDAD, ""
                            End If
                        End If
                        Set rsSPDA = Nothing
                    End If
                End If
            Case 3 ' SELLANTE
                actualizar_fecha_comienzo codmuestra, "SELLANTES_RESULTADOS"
                actualizar_fecha_fin codmuestra, "SELLANTES_RESULTADOS"
            Case 5 ' plasma
                actualizar_fecha_plasma codmuestra
            Case Else ' RESTO DE MUESTRAS
                actualizar_fecha_comienzo codmuestra, "DETERMINACIONES"
                actualizar_fecha_fin codmuestra, "DETERMINACIONES"
        End Select
'    End If
'   SI LA FECHA DE FINALIZACION ES 0000-00-00, LA ACTUALIZAMOS CON LA DE CIERRE
    consulta = "UPDATE MUESTRAS SET FECHA_FINALIZACION = FECHA_CIERRE WHERE id_muestra=" & codmuestra & " and FECHA_FINALIZACION  = '0000-00-00'"
    execute_bd consulta

    ' Para aguas, verificar si pertenece a un agua agrupada
    If generar = True Then
        If oMuestra.getBANO_ID <> 0 Then
'M00687-I
'            If oMuestra.getTIPO_MUESTRA_ID = 2 Then
'                Dim rs_aux As ADODB.RecordSet
'                Dim muestra_agua As Long
'                Dim olb As New clsLineas_Banos
'                Set rs_aux = olb.Buscar_Documento(olb.Buscar_Bano(oMuestra.getBANO_ID))
'                If rs_aux.RecordCount <> 0 Then ' Es una agrupación en lineas_banos
'                   Do
'                      muestra_agua = oMuestra.Cargar_Agua(oMuestra.getFECHA_RECEPCION, rs_aux("bano_id"))
'                      If muestra_agua > 0 Then
'                        oMuestra.CargaMuestra (muestra_agua)
'                        If oMuestra.getCERRADA = 0 Or oMuestra.getCERRADA = 2 Then
'                           generar = False
'                        End If
'                      End If
'                      rs_aux.MoveNext
'                   Loop Until rs_aux.EOF
'                End If
'            Else
                ' Verificamos si es un fluido agrupado
                Dim otm As New clsTipos_muestra
                otm.CARGAR oMuestra.getTIPO_MUESTRA_ID
                If otm.getTIPO_ESPECIAL_ID = tipo_especial.FLUIDO Then
                    Dim oFF As New clsFluidos_ficha
                    Dim rs_fluido As ADODB.Recordset
                    Set rs_fluido = oFF.Es_Fluido_En_Columna(codmuestra)
                    If rs_fluido.RecordCount > 0 Then
                        Do
                            oMuestra.CargaMuestra (rs_fluido("ID_MUESTRA"))
                            If oMuestra.getCERRADA = 0 Or oMuestra.getCERRADA = 2 Then
                                generar = False
                                MsgBox "RECUERDE: Debe cerrar todos los fluidos de la misma linea para generar el informe.", vbInformation, App.Title
                            End If
                            rs_fluido.MoveNext
                        Loop Until rs_fluido.EOF Or generar = False
                    End If
                End If
'            End If
'M00687-F
        End If
    End If
    ' Generar edición
    If generar = True And oMuestra.getANULADA = 0 Then
'C001       If omuestra.getTIPO_MUESTRA_ID <> 71 And omuestra.getTIPO_MUESTRA_ID <> 72 Then
        imprimir codmuestra, 1, False
'C001       End If
    End If
    ' Informamos el estado de la muestra
    Me.informar_situacion codmuestra

    Cerrar = True
    Exit Function
fallo:
    Cerrar = False
    MsgBox "Error al cerrar la muestra.", vbCritical, Err.Description
End Function
Public Function pte_cierre(codmuestra As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE muestras SET" & _
               " CERRADA=2 " & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta
    pte_cierre = True
    Exit Function
fallo:
    pte_cierre = False
    MsgBox "Error al cerrar la muestra.", vbCritical, Err.Description
End Function
Public Function Abrir(codmuestra As Long, mantenerEnviada As Boolean) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    If mantenerEnviada = True Then
        consulta = "UPDATE muestras SET" & _
                   " CERRADA=0,REVISION_USUARIO=0,IMPRESA=0 " & _
                   " WHERE id_muestra=" & codmuestra
    Else
        consulta = "UPDATE muestras SET" & _
                   " CERRADA=0,ENVIADO_CORREO=0,REVISION_USUARIO=0,IMPRESA=0 " & _
                   " WHERE id_muestra=" & codmuestra
    End If
    execute_bd consulta
    Abrir = True
    Exit Function
fallo:
    Abrir = False
    MsgBox "Error al abrir la muestra.", vbCritical, Err.Description
End Function

Public Function aumentar_edicion_impresa(codmuestra As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE muestras SET" & _
               " ULT_EDICION_IMP=ULT_EDICION_IMP+1" & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta
    aumentar_edicion_impresa = True
    Exit Function
fallo:
    aumentar_edicion_impresa = False
    MsgBox "Error al aumentar la edicion de la muestra.", vbCritical, Err.Description
End Function
Public Function disminuir_edicion_impresa(codmuestra As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE muestras SET" & _
               " ULT_EDICION_IMP=ULT_EDICION_IMP-1" & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta
    disminuir_edicion_impresa = True
    Exit Function
fallo:
    disminuir_edicion_impresa = False
    MsgBox "Error al disminuir la edicion de la muestra.", vbCritical, Err.Description
End Function
Public Function modificar_edicion_impresa(codmuestra As Long, EDICION As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE muestras SET" & _
               " ULT_EDICION_IMP=" & EDICION & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta
    modificar_edicion_impresa = True
    Exit Function
fallo:
    modificar_edicion_impresa = False
    MsgBox "Error al modificar la edicion de la muestra.", vbCritical, Err.Description
End Function
'M1054-I
Public Function modificarInformeManual(codmuestra As Long, MANUAL As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE muestras " & _
               "   SET INFORME_MANUAL =" & MANUAL & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta
    modificarInformeManual = True
    Exit Function
fallo:
    modificarInformeManual = False
    MsgBox "Error al modificarInformeManual de la muestra.", vbCritical, Err.Description
End Function
'M1054-F
'Public Function datos_especiales(TIPO_ANALISIS As Long) As ADODB.Recordset
'    On Error GoTo fallo
'    Dim consulta As String
'    consulta = "select td.id_tipo_dato, td.nombre, td.unidad_id " & _
'               " from tipos_datos_analisis db, tipos_dato td " & _
'               " where db.tipo_dato_id = td.id_tipo_dato " & _
'               "   and db.tipo_muestra_id = " & TIPO_ANALISIS
'    Set datos_especiales = datos_bd(consulta)
'    Exit Function
'fallo:
'    MsgBox "Error al recuperar los datos especiales.", vbCritical, Err.Description
'End Function
'Public Function datos_cabecera_documento(MUESTRA As Long) As ADODB.Recordset
'    On Error GoTo fallo
'    Dim consulta As String
'    consulta = "select cli.nombre,cli.informes_direccion,cli.informes_cod_postal,mun.nombre,prov.nombre," & _
'               "       cli.centro,cli.responsable, " & _
'               "       mue.fecha_recepcion,mue.fecha_muestreo,mue.FECHA_CIERRE,mue.ult_edicion_imp,   " & _
'               "       tm.nombre, ta.nombre, mue.referencia_cliente, ee.descripcion, em.descripcion,mue.observaciones_muestreo,mue.id_general, " & _
'               "       cli.id_cliente,cli.cif,cli.telefono,cli.fax,mue.fecha_prev_fin,f.descripcion,mue.observaciones,mue.tipo_analisis_id,mue.hora_recepcion,cli.direccion,mue.observaciones_entrega,cli.direccion,mue.fecha_comienzo,mue.bano_id,mue.fecha_finalizacion " & _
'               " from muestras as mue, clientes as cli, municipios as mun, provincias as prov, " & _
'               "      tipos_muestra as tm, tipos_analisis as ta, " & _
'               "      entidades_entrega as ee, entidades_muestreo as em, formatos as f " & _
'               " where mue.cliente_id = cli.id_cliente " & _
'               "   and cli.informes_municipio_id = mun.id_municipio " & _
'               "   and cli.informes_provincia_id = prov.id_provincia " & _
'               "   and mue.tipo_muestra_id = tm.id_tipo_muestra " & _
'               "   and mue.tipo_analisis_id = ta.id_tipo_analisis " & _
'               "   and mue.entidad_entrega_id = ee.id_entidad_entrega " & _
'               "   and mue.entidad_muestreo_id = em.id_entidad_muestreo " & _
'               "   and mue.formato_id = f.id_formato " & _
'               "   and mue.id_muestra = " & MUESTRA
'    Set datos_cabecera_documento = datos_bd(consulta)
'    log ("Datos de cabecera recuperados.")
'    Exit Function
'fallo:
'    MsgBox "Error al recuperar los datos de cabecera del documento.", vbCritical, Err.Description
'End Function
'Public Function datos_cabecera_bano(MUESTRA As Long) As ADODB.Recordset
'    On Error GoTo fallo
'    Dim consulta As String
'    consulta = "select li.nombre,pb.nombre,tf.nombre,sol.nombre,mue.observaciones_entrega,ba.nombre,ba.volumen,ba.id_bano,mue.referencia_cliente " & _
'               " from muestras as mue, banos as ba, lineas as li, procesos_base as pb, " & _
'               "      tipos_frecuencia as tf, soluciones as sol " & _
'               " where mue.bano_id = ba.id_bano " & _
'               "   and ba.linea_id = li.id_linea " & _
'               "   and ba.proceso_base_id = pb.id_proceso_base " & _
'               "   and ba.tipo_frecuencia_id = tf.id_tipo_frecuencia " & _
'               "   and ba.solucion_id = sol.id_solucion " & _
'               "   and mue.id_muestra = " & MUESTRA
'    Set datos_cabecera_bano = datos_bd(consulta)
'    Exit Function
'fallo:
'    MsgBox "Error al recuperar los datos de cabecera del baño.", vbCritical, Err.Description
'End Function
Public Function obtener_tipo_muestra(muestra As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "select tm.codigo, tm.nombre " & _
               "  from muestras as mue, tipos_muestra as tm " & _
               " where mue.tipo_muestra_id = tm.id_tipo_muestra " & _
               "   and mue.id_muestra = " & muestra
    Set obtener_tipo_muestra = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al obtener el tipo de muestra.", vbCritical, Err.Description
End Function
Public Function obtener_tipo_documento(muestra As Long) As Integer
    On Error GoTo fallo
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "select tm.tipo_documento_id " & _
               "  from muestras as mue, tipos_muestra as tm " & _
               " where mue.tipo_muestra_id = tm.id_tipo_muestra " & _
               "   and mue.id_muestra = " & muestra
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        obtener_tipo_documento = 1
    Else
        obtener_tipo_documento = rs.Fields(0)
    End If
    Set rs = Nothing
    Exit Function
fallo:
    MsgBox "Error al obtener el tipo de documento de la muestra.", vbCritical, Err.Description
End Function
Public Function obtener_plantilla(muestra As Long) As String
    On Error GoTo fallo
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "select td.plantilla " & _
               "  from muestras as mue, tipos_muestra as tm, tipos_documentos td " & _
               " where mue.tipo_muestra_id = tm.id_tipo_muestra " & _
               "   and tm.tipo_documento_id = td.id_tipo_documento " & _
               "   and mue.id_muestra = " & muestra
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        obtener_plantilla = ""
    Else
        obtener_plantilla = rs.Fields(0)
    End If
    Set rs = Nothing
    Exit Function
fallo:
    obtener_plantilla = ""
    MsgBox "Error al obtener la plantilla asociada a la muestra.", vbCritical, Err.Description
End Function
Public Function Modificar(codmuestra As Long) As Boolean
    On Error GoTo fallo
    Dim sql As String
    ' Verificar al modificar la muestra si se ha tocado el precio
    Dim oParametro As New clsParametros
    oParametro.Carga PARAM_USUARIO_VIGILADO, ""
    If USUARIO.getID_EMPLEADO = oParametro.getVALOR Then
        Dim oMuestra As New clsMuestra
        oMuestra.CargaMuestra codmuestra
        If moneda(oMuestra.getPRECIO) <> moneda(PRECIO) Then
            Dim ASUNTO As String
            Dim DETALLE As String
            ASUNTO = "El usuario " & USUARIO.getUSUARIO & " ha modificado manualmente el precio de la muestra."
            DETALLE = "" & vbNewLine
            DETALLE = DETALLE & " Fecha : " & Format(Date, "dd-mm-yyyy") & vbNewLine
            DETALLE = DETALLE & " Hora  : " & Time & vbNewLine & vbNewLine
            DETALLE = DETALLE & " Muestra : " & oMuestra.getID_GENERAL & "/" & oMuestra.getANNO & vbNewLine
            Dim oCliente As New clsCliente
            oCliente.CargaCliente oMuestra.getCLIENTE_ID
            DETALLE = DETALLE & " Cliente : " & oCliente.getNOMBRE & vbNewLine & vbNewLine
            DETALLE = DETALLE & " Precio Anterior : " & moneda(oMuestra.getPRECIO) & vbNewLine
            DETALLE = DETALLE & " Precio Nuevo : " & moneda(PRECIO) & vbNewLine
            
            oParametro.Carga PARAM_USUARIO_VIGILADO_CORREO, ""
            ret = Enviar_Mail_CDO(oParametro.getVALOR, ASUNTO, DETALLE, vbNullString)
        End If
    End If
    Dim frecogida As String
    If FECHA_RECOGIDA = "" Then
        frecogida = "FECHA_RECOGIDA = NULL,"
    Else
        frecogida = "FECHA_RECOGIDA = '" & FECHA_RECOGIDA & "',"
    End If
    'M1009 : OFERTA_ID
    sql = "update muestras " & _
               "   set referencia_cliente = '" & REFERENCIA_CLIENTE & "'," & _
               " fecha_recepcion = '" & FECHA_RECEPCION & "'," & _
               IIf(HORA_RECEPCION <> "", "hora_recepcion = '" & HORA_RECEPCION & "',", "") & _
               " formato_id = " & FORMATO_ID & "," & " entidad_entrega_id = " & ENTIDAD_ENTREGA_ID & "," & _
               " detalle_entrega = '" & DETALLE_ENTREGA & "'," & _
               " observaciones_entrega = '" & OBSERVACIONES_ENTREGA & "'," & _
               " fecha_muestreo = '" & FECHA_MUESTREO & "'," & _
               " entidad_muestreo_id = " & ENTIDAD_MUESTREO_ID & "," & _
               " cliente_id = " & CLIENTE_ID & "," & _
               " CENTRO_id = " & CENTRO_ID & "," & _
               " REPLACEMENT_id = " & REPLACEMENT_ID & "," & _
               " PEDIDO_ID = " & PEDIDO_ID & "," & " OFERTA_ID = " & OFERTA_ID & "," & _
               " detalle_muestreo = '" & DETALLE_MUESTREO & "'," & " observaciones_muestreo = '" & OBSERVACIONES_MUESTREO & "'," & _
               " fecha_prev_fin = '" & FECHA_PREV_FIN & "',AJUSTE = " & ajuste & "," & _
               " CONSULTA = " & consulta & "," & _
               " CONSULTA_OBSERVACIONES = '" & CONSULTA_OBSERVACIONES & "'," & _
               " precio = '" & PRECIO & "',precio_fijado=" & PRECIO_FIJADO & ",ENAC = " & ENAC & "," & _
               " NADCAP = " & nadcap & ",URGENTE = " & urgente & ", precinto = '" & precinto & "'," & " producto = '" & producto & "'," & _
               " ANALISIS_DUPLICADO = " & ANALISIS_DUPLICADO & "," & _
               " EMPLEADO_ID = " & EMPLEADO_ID & "," & " RESPONSABLE_ID = " & RESPONSABLE_ID & "," & " IPA = " & IPA & "," & frecogida & _
               " MOTIVO_RETRASO_ID = " & MOTIVO_RETRASO_ID & ",RETRASO_OBSERVACION='" & RETRASO_OBSERVACION & "'," & _
               " OP_VUELO = " & OP_VUELO & ",OP_INSITU =" & OP_INSITU & ", OP_LABMOVIL=" & OP_LABMOVIL & ",OP_NORUTINARIA=" & OP_NORUTINARIA & ",OP_REPETICION=" & OP_REPETICION & "," & _
               " observaciones = '" & OBSERVACIONES & "'" & _
               " where id_muestra = " & codmuestra
    execute_bd sql
    ' Modificamos las determinaciones para la duplicidad
    sql = "update determinaciones set es_duplicado = " & ANALISIS_DUPLICADO & " where muestra_id = " & codmuestra
    execute_bd sql
    ' Modificamos el ajuste si es una calibracion
    sql = "update eq_calibracion_equipos set ajuste = " & ajuste & " where muestra_id = " & codmuestra
    execute_bd sql
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al modificar la muestra.", vbCritical, Err.Description
End Function
Public Function obtener_muestras_anteriores(muestra As Long, DETERMINACION As Long, UNICA As Boolean, Optional fdesde As String, Optional fhasta As String) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    Dim filtro As String
    If fdesde <> "" Then
        filtro = filtro & " AND mue.FECHA_RECEPCION >='" & Format(fdesde, "yyyy-mm-dd") & "'"
    End If
    If fhasta <> "" Then
        filtro = filtro & " AND mue.FECHA_RECEPCION <='" & Format(fhasta, "yyyy-mm-dd") & "'"
    End If
    consulta = "select mue.id_general,mue.bano_id,det.tipo_determinacion_id,det.resultado,td.nombre,mue.fecha_recepcion,mue.id_muestra " & _
               "  from muestras as m, determinaciones as d, muestras as mue, determinaciones as det, tipos_determinacion as td" & _
               " WHERE m.ID_MUESTRA = " & muestra & _
               "   AND d.ID_DETERMINACION = " & DETERMINACION & _
               "   AND mue.tipo_muestra_id = m.TIPO_MUESTRA_ID " & _
               "   AND mue.cliente_id = m.CLIENTE_ID " & _
               "   AND mue.bano_id = m.BANO_ID " & _
               "   and mue.id_muestra = det.muestra_id " & _
               "   and det.tipo_determinacion_id = d.TIPO_DETERMINACION_ID " & _
               "   and det.tipo_determinacion_id = td.id_tipo_determinacion " & _
               "   and mue.anulada = 0 " & _
               "   and det.resultado <> '' " & _
               "   and mue.id_muestra <= m.ID_MUESTRA " & _
               filtro & _
               " order by mue.id_muestra desc "
'               "   AND mue.tipo_analisis_id = m.TIPO_ANALISIS_ID "
    If UNICA = True Then
        consulta = consulta & " limit 1 "
    End If
    Set obtener_muestras_anteriores = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al obtener las muestras anteriores (obtener_muestras_anteriores).", vbCritical, Err.Description
End Function


'Public Function obtener_bano_anteriores(MUESTRA As Long, DETERMINACION As Long) As ADODB.Recordset
'    On Error GoTo fallo
'    Dim consulta As String
'    Dim oMuestra As New clsMuestra
'    Dim TIPO_DETERMINACION As Integer
'    Dim oDET As New clsDeterminaciones
'    oDET.CargarDeterminacion (DETERMINACION)
'    TIPO_DETERMINACION = oDET.getTIPO_DETERMINACION_ID
'    oMuestra.CargaMuestra (MUESTRA)
'    consulta = "select mue.id_general,mue.bano_id,det.tipo_determinacion_id,det.resultado,td.nombre,mue.fecha_recepcion,mue.id_muestra " & _
'               "  from muestras as mue, determinaciones as det, tipos_determinacion as td" & _
'               " where mue.tipo_muestra_id = " & oMuestra.getTIPO_MUESTRA_ID & _
'               "   and mue.tipo_analisis_id = " & oMuestra.getTIPO_ANALISIS_ID & _
'               "   and mue.bano_id = " & oMuestra.getBANO_ID & _
'               "   and mue.id_muestra = det.muestra_id " & _
'               "   and det.tipo_determinacion_id = " & TIPO_DETERMINACION & _
'               "   and det.tipo_determinacion_id = td.id_tipo_determinacion " & _
'               "   and mue.anulada = 0 " & _
'               "   and det.resultado <> '' " & _
'               "   and mue.cliente_id = " & oMuestra.getCLIENTE_ID & _
'               "   and mue.id_muestra <= " & MUESTRA & _
'               " order by mue.id_muestra desc"
'    Set obtener_bano_anteriores = datos_bd(consulta)
'    Exit Function
'fallo:
'    MsgBox "Error al obtener el tipo de muestra.", vbCritical, Err.Description
'End Function
Public Function historico_determinaciones(TIPO_DETERMINACION As Long, fdesde As Date, fhasta As Date) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "select mue.id_muestra,det.ID_DETERMINACION,id_general,CONCAT(tp.codigo,'-',CAST(mue.id_particular AS CHAR)) AS CODIGO,det.resultado,det.DIF_DUPLICADOS, td.nombre,det.FECHA, usu.USUARIO,det.USUARIO_ID_REVISION, revisor.USUARIO, det.FECHA_REVISION " & _
               "  from muestras as mue " & _
               " inner join determinaciones as det on mue.ID_MUESTRA = det.MUESTRA_ID " & _
               " inner join tipos_determinacion as td on det.TIPO_DETERMINACION_ID = td.ID_TIPO_DETERMINACION " & _
               " inner join usuarios as usu on det.EMPLEADO_ID = usu.ID_EMPLEADO " & _
               " inner join tipos_muestra as tp on mue.tipo_muestra_id = tp.id_tipo_muestra " & _
               "  left join usuarios as revisor on det.USUARIO_ID_REVISION = revisor.ID_EMPLEADO " & _
               " where det.TIPO_DETERMINACION_ID  = " & TIPO_DETERMINACION & _
               "   and det.fecha >= '" & Format(fdesde, "yyyy-mm-dd") & "' " & _
               "   and det.fecha <= '" & Format(fhasta, "yyyy-mm-dd") & "' " & _
               "   and mue.ANALISIS_DUPLICADO = 1 " & _
               "   and mue.CERRADA <> 0 " & _
               "   and mue.ANULADA = 0 " & _
               "   and det.DIF_DUPLICADOS <> '' " & _
               " order by mue.ID_MUESTRA desc "
    Set historico_determinaciones = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al obtener el historico_determinaciones.", vbCritical, Err.Description
End Function

Public Function buscar_ultimo_codigo_general(ANNO) As Long
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_GENERAL) FROM muestras where anno = " & ANNO
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        buscar_ultimo_codigo_general = 1
    Else
        buscar_ultimo_codigo_general = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Function

Public Function Cargar_Agua(fecha As Date, agua As Integer) As Long
    On Error GoTo fallo:
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT id_muestra FROM muestras WHERE bano_id=" & agua & " AND FECHA_RECEPCION='" & Format(fecha, "yyyy-mm-dd") & "' AND ANULADA=0" ' AND CERRADA=1"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.RecordCount = 0) Then
        Cargar_Agua = 0
    Else
        Cargar_Agua = rs("id_muestra")
    End If
    Set rs = Nothing
    Exit Function
fallo:
    Cargar_Agua = 0
    MsgBox "Error al recuperar la muestra_Agua (clsMuestra.CargaMuestra)", vbCritical, Err.Description
End Function
Public Function obtener_id_muestra(GENERAL As Long, ANNO As Integer) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "select id_muestra " & _
               "  from muestras " & _
               " where id_general = " & GENERAL & _
               "   and anno = " & ANNO
    Set obtener_id_muestra = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al obtener_id_muestra de muestra.", vbCritical, Err.Description
End Function
Public Function comprobar_cierre(muestra As Long) As Integer
    ' Cerrar Muestra
    On Error GoTo fallo
    Dim rs As ADODB.Recordset
    Dim Cerrar As Boolean
    Dim oMuestra As New clsMuestra
    Dim oDeter As New clsDeterminaciones
    Cerrar = True
    comprobar_cierre = 0
    If oMuestra.CargaMuestra(muestra) = True Then
        If oMuestra.getCERRADA = 1 Then
            Exit Function
        End If
        Select Case oMuestra.getANALISIS_MODIFICADO
            Case 2 ' Es un control de eficacia
                Dim oCe_resultados As New clsCe_resultados
                Cerrar = oCe_resultados.Cerrar(muestra)
            Case 3 ' Es un sellante
                Dim oSellante_resultado As New clsSellantes_resultados
                Cerrar = oSellante_resultado.Cerrar(muestra)
            Case 5 ' PLASMA
                Cerrar = False
            Case Else
                Set rs = oDeter.lista_determinaciones(muestra)
                If rs.RecordCount <> 0 Then
                    Do
                        oDeter.CargarDeterminacion (rs("id_determinacion"))
                        If oDeter.getRESULTADO = "" Or IsNull(oDeter.getRESULTADO) Then
                            Cerrar = False
                        End If
                        rs.MoveNext
                    Loop Until rs.EOF Or Cerrar = False
                End If
        End Select
        If Cerrar = True Then
          If USUARIO.getPER_CIERRE = True Then
            If MsgBox("Muestra con resultados completos. ¿Desea cerrar la muestra?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                'M1370-I
                'oMuestra.Cerrar (MUESTRA)
                'comprobar_cierre = 1
                If MsgBox("¿Se han verificado los equipos y las condiciones ambientales del ensayo?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                     oMuestra.Cerrar (muestra)
                     comprobar_cierre = 1
                Else
                    oMuestra.pte_cierre (muestra)
                    comprobar_cierre = 2
                    Cerrar = False
                End If
                'M1370-F
            Else
                oMuestra.pte_cierre (muestra)
                comprobar_cierre = 2
                Cerrar = False
            End If
          Else
            oMuestra.pte_cierre (muestra)
            comprobar_cierre = 2
          End If
          Set oMuestra = Nothing
        End If
    End If
    Set rs = Nothing
    Exit Function
fallo:
    comprobar_cierre = 0
    MsgBox "Error al cerrar la muestra.", vbCritical, App.Title
End Function
Public Function ImporteMuestraPorDeterminaciones(muestra As Long, cliente As Long) As Currency
    On Error GoTo fallo
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    Dim oCliente As New clsCliente
    Dim TARIFA As Integer
    TARIFA = oCliente.Recupera_Tarifa(cliente)
    If TARIFA = 0 Then
'MANTIS-929-I
'        consulta = " SELECT SUM(td.PRECIO) " & _
'                   "  FROM determinaciones d,tipos_determinacion td " & _
'                   " WHERE d.tipo_determinacion_id = td.id_tipo_determinacion" & _
'                   "   AND d.muestra_id=" & MUESTRA
'    Else
'        consulta = " SELECT SUM(tp.PRECIO) " & _
'                   "  FROM determinaciones d,tipos_determinacion td,tarifas_precios tp " & _
'                   " WHERE d.tipo_determinacion_id = td.id_tipo_determinacion" & _
'                   "   AND tp.tipo_determinacion_id = td.id_tipo_determinacion" & _
'                   "   AND tp.tarifa_id = " & TARIFA & _
'                   "   AND d.muestra_id=" & MUESTRA
        consulta = " SELECT SUM(td.PRECIO) " & _
                   "  FROM determinaciones d,tipos_determinacion td " & _
                   " WHERE d.tipo_determinacion_id = td.id_tipo_determinacion" & _
                   "   AND d.muestra_id=" & muestra & _
                   "   AND d.resultado <> '--'"
    Else
        consulta = " SELECT SUM(tp.PRECIO) " & _
                   "  FROM determinaciones d,tipos_determinacion td,tarifas_precios tp " & _
                   " WHERE d.tipo_determinacion_id = td.id_tipo_determinacion" & _
                   "   AND tp.tipo_determinacion_id = td.id_tipo_determinacion" & _
                   "   AND tp.tarifa_id = " & TARIFA & _
                   "   AND d.muestra_id=" & muestra & _
                   "   AND d.resultado <> '--'"
'MANTIS-929-F
    End If
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
        ImporteMuestraPorDeterminaciones = 0
    Else
        ImporteMuestraPorDeterminaciones = rs.Fields(0)
    End If
    Set rs = Nothing
    Exit Function
fallo:
    ImporteMuestraPorDeterminaciones = 0
    MsgBox "Error al calcular el precio por determinacioens (ImporteMuestraPorDeterminaciones)", vbCritical, Err.Description
End Function
Public Function informar_impresion(codmuestra As Long, USUARIO As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE muestras SET" & _
               " IMPRESA=" & USUARIO & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta
    informar_impresion = True
    Exit Function
fallo:
    informar_impresion = False
    MsgBox "Error al informar la impresión de la muestra.", vbCritical, Err.Description
End Function
Public Function informar_urgente(codmuestra As Long, urgente As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE muestras SET" & _
               " URGENTE=" & urgente & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta
    informar_urgente = True
    Exit Function
fallo:
    informar_urgente = False
    MsgBox "Error al informar la muestra (informar_urgente).", vbCritical, Err.Description
End Function
Public Function informar_ajuste(codmuestra As Long, ajuste As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE muestras SET" & _
               " AJUSTE=" & ajuste & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta
    informar_ajuste = True
    Exit Function
fallo:
    informar_ajuste = False
    MsgBox "Error al informar la muestra (informar_ajuste).", vbCritical, Err.Description
End Function
Public Function informar_situacion(ID As Long) As Boolean
    On Error GoTo fallo
    ' situacion
    ' 0 : CONFORME
    ' 1 : CERCA DE FUERA DE RANGO
    ' 2 : NO CONFORME
    Dim consulta As String
    Dim situacion As Integer
    Dim rs As ADODB.Recordset
    situacion = C_SITUACION.S_EN_RANGO
    consulta = "SELECT ANALISIS_MODIFICADO FROM MUESTRAS WHERE ID_MUESTRA = " & ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        Select Case rs(0)
        Case 2 ' CE
            consulta = "SELECT MIN(CONFORME) FROM CE_RESULTADOS WHERE MUESTRA_ID = " & ID
            Set rs = datos_bd(consulta)
            If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then 'si es nulo No se recupero ninguno
               situacion = C_SITUACION.S_EN_RANGO
            Else
                If CInt(rs.Fields(0).Value) = 0 Then
                    situacion = C_SITUACION.S_FUERA_RANGO
                End If
            End If
        Case 3 ' Sellante
            consulta = "SELECT MIN(CONFORME) FROM SELLANTES_RESULTADOS WHERE MUESTRA_ID = " & ID
            Set rs = datos_bd(consulta)
            If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then 'si es nulo No se recupero ninguno
               situacion = C_SITUACION.S_EN_RANGO
            Else
                If CInt(rs.Fields(0).Value) = 0 Then
                    situacion = C_SITUACION.S_FUERA_RANGO
                End If
            End If
        Case 5 ' Plasma
            Dim oPR As New clsPlasma_recepcion
            oPR.Carga ID
            If oPR.getRESULT = 0 Then
                situacion = C_SITUACION.S_FUERA_RANGO
            Else
                situacion = C_SITUACION.S_EN_RANGO
            End If
        Case Else
            consulta = "SELECT MAX(SITUACION) FROM DETERMINACIONES WHERE MUESTRA_ID = " & ID
            Set rs = datos_bd(consulta)
            If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then 'si es nulo No se recupero ninguno
               situacion = C_SITUACION.S_EN_RANGO
            Else
               situacion = rs.Fields(0)
            End If
        End Select
    End If

    consulta = "UPDATE muestras SET" & _
               " SITUACION=" & situacion & _
               " WHERE id_muestra=" & ID
    execute_bd consulta
    informar_situacion = True
    Exit Function
fallo:
    informar_situacion = False
    MsgBox "Error al informar la situacion de la muestra.", vbCritical, Err.Description
End Function
Public Function informar_producto(idMuestra As Long, producto As String) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE muestras " & _
               "   set PRODUCTO='" & producto & "'" & _
               " WHERE id_muestra=" & idMuestra
    execute_bd consulta
    informar_producto = True
    Exit Function
fallo:
    informar_producto = False
    MsgBox "Error al informar el producto.", vbCritical, Err.Description
End Function

Public Function informar_correo(codmuestra As Long, USUARIO As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE muestras SET" & _
               " ENVIADO_CORREO=" & USUARIO & _
               " ,FECHA_ENVIO = CURRENT_TIMESTAMP " & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta
    informar_correo = True
    Exit Function
fallo:
    informar_correo = False
    MsgBox "Error al informar el envio de correo de la muestra.", vbCritical, Err.Description
End Function

Public Function informar_adjunto(codmuestra As Long, texto As String) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE MUESTRAS " & _
               "   SET OBSERVACIONES_ENTREGA=CONCAT(OBSERVACIONES_ENTREGA,' DOCUMENTO ADJUNTO:" & texto & "') " & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta
    informar_adjunto = True
    Exit Function
fallo:
    informar_adjunto = False
    MsgBox "Error al informar el adjunto en las observaciones.", vbCritical, Err.Description
End Function
Public Function informar_firma(codmuestra As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    Dim firma As String
'    firma = codmuestra & ".jpg"
    firma = codmuestra
    If firma = "" Then
        firma = "0"
    End If
    consulta = "UPDATE MUESTRAS " & _
               "   SET FIRMA='" & firma & "'" & _
               " WHERE id_muestra=" & codmuestra
    execute_bd consulta
    informar_firma = True
    Exit Function
fallo:
    informar_firma = False
    MsgBox "Error al guardar la firma de la muestra.", vbCritical, Err.Description
End Function
Public Function cargar_firma(codmuestra As Long) As String
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT FIRMA FROM MUESTRAS " & _
               " WHERE id_muestra=" & codmuestra
    Dim rs As ADODB.Recordset
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        cargar_firma = Replace(rs(0), "/", "\")
    Else
        cargar_firma = ""
    End If
    Exit Function
fallo:
    cargar_firma = ""
    MsgBox "Error al cargar la firma de la muestra.", vbCritical, Err.Description
End Function

Public Function Eliminar(codmuestra As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    ' Fluidos
    execute_bd "DELETE FROM FLUIDOS_RECEPCION WHERE MUESTRA_ID = " & codmuestra
    execute_bd "DELETE FROM FLUIDOS_RESULTADOS WHERE MUESTRA_ID = " & codmuestra
    execute_bd "DELETE FROM FLUIDOS_RESULTADOS_GRID WHERE MUESTRA_ID = " & codmuestra
    execute_bd "DELETE FROM FLUIDOS_IP WHERE MUESTRA_ID = " & codmuestra
    ' Plasma
    execute_bd "DELETE FROM plasma_equipos WHERE MUESTRA_ID = " & codmuestra
    execute_bd "DELETE FROM plasma_reactivos WHERE MUESTRA_ID = " & codmuestra
    execute_bd "DELETE FROM plasma_recepcion WHERE MUESTRA_ID = " & codmuestra
    execute_bd "DELETE FROM plasma_resultados WHERE MUESTRA_ID = " & codmuestra
    ' Ce
    execute_bd "DELETE FROM CE_RECEPCION WHERE MUESTRA_ID = " & codmuestra
    execute_bd "DELETE FROM CE_RESULTADOS WHERE MUESTRA_ID = " & codmuestra
    ' Sellantes
    execute_bd "DELETE FROM SELLANTES_RECEPCION WHERE MUESTRA_ID = " & codmuestra
    execute_bd "DELETE FROM SELLANTES_RESULTADOS WHERE MUESTRA_ID = " & codmuestra
    execute_bd "DELETE FROM SELLANTES_DETERMINACIONES WHERE MUESTRA_ID = " & codmuestra
    
    consulta = "DELETE FROM MUESTRAS WHERE ID_MUESTRA = " & codmuestra
    execute_bd consulta
    consulta = "DELETE FROM MUESTRAS_ANULADAS WHERE MUESTRA_ID = " & codmuestra
    execute_bd consulta
    consulta = "DELETE FROM MUESTRAS_SOLUCIONES WHERE MUESTRA_ID = " & codmuestra
    execute_bd consulta
    consulta = "DELETE FROM Datos_valores WHERE MUESTRA_ID = " & codmuestra
    execute_bd consulta
    
    consulta = "UPDATE eq_calibracion_equipos SET MUESTRA_ID = 0 WHERE MUESTRA_ID = " & codmuestra
    execute_bd consulta
    
    Dim oDeterminacion As New clsDeterminaciones
    oDeterminacion.Eliminar_Por_Muestra (codmuestra)
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al Eliminar la muestra.", vbCritical, Err.Description
End Function

Public Function esControlEficacia(muestra As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT ANALISIS_MODIFICADO FROM MUESTRAS" & _
               " WHERE ID_MUESTRA=" & muestra
    esControlEficacia = False
    If datos_bd(consulta).Fields(0) = 2 Then
       esControlEficacia = True
    End If
    Exit Function
fallo:
    log ("Error al determinar si es un control de eficacia (esControlEficacia)")
End Function
Public Function esControlEficaciaCargaSostenida(muestra As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    esControlEficaciaCargaSostenida = False
    consulta = "select tipo_muestra_id " & _
               "  from muestras " & _
               " where id_muestra = " & muestra
    If datos_bd(consulta).Fields(0) = TIPOS_MUESTRAS.ecs Then
       esControlEficaciaCargaSostenida = True
    End If
    Exit Function
fallo:
    log ("Error al determinar si es un control de eficacia (esControlEficaciaCargaSostenida)")
End Function

Public Function esSellante(muestra As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT ANALISIS_MODIFICADO FROM MUESTRAS" & _
               " WHERE ID_MUESTRA=" & muestra
    esSellante = False
    If datos_bd(consulta).Fields(0) = 3 Then
       esSellante = True
    End If
    Exit Function
fallo:
    log ("Error al determinar si es un Sellante (esSellante)")
End Function

Public Function informar_precio_muestra(muestra As Long) As Boolean
        On Error GoTo fallo
        Dim total As Currency
        Dim rs_defecto As ADODB.Recordset
        Dim rs As ADODB.Recordset
        Dim odd As New clsDeterminaciones_analisis
        Dim oDeter As New clsDeterminaciones
        Dim oCliente As New clsCliente
        Dim oTarifa As New clsTarifas_precios
        Dim oMuestra As New clsMuestra
        If oMuestra.CargaMuestra(muestra) Then
            If oMuestra.getPRECIO_FIJADO = 0 Then
                oCliente.CargaCliente (oMuestra.getCLIENTE_ID)
                ' Recuperamos el precio del analisis o bano por tarifa
                ' Miramos si se factura por tipo analisis o control de eficacia
                If oMuestra.getBANO_ID = 0 Or oMuestra.getANALISIS_MODIFICADO = 2 Then
                    If oTarifa.Carga_por_TA(oMuestra.getTIPO_ANALISIS_ID, oCliente.getTARIFA_ID) Then
                        total = CCur(Replace(oTarifa.getPRECIO, ".", ","))
                    End If
                Else
                    If oTarifa.Carga_por_BANO(oMuestra.getBANO_ID, oCliente.getTARIFA_ID) Then
                        total = CCur(Replace(oTarifa.getPRECIO, ".", ","))
                    End If
                End If
                ' Recuperamos los datos por defecto del analisis o bano
                If oMuestra.getBANO_ID = 0 Then
                    total = total + odd.Precio_determinaciones_por_tipo_analisis(oMuestra.getTIPO_ANALISIS_ID, muestra, oCliente.getTARIFA_ID)
                Else
                    total = total + odd.Precio_determinaciones_por_bano(oMuestra.getBANO_ID, muestra, oCliente.getTARIFA_ID)
                End If
                ' Actualizamos el precio de la muestra
                oMuestra.actualizar_precio muestra, Replace(total, ",", ".")
            End If
        End If
    Set rs = Nothing
    Set rs_defecto = Nothing
    Set odd = Nothing
    Set oDeter = Nothing
    Set oMuestra = Nothing
    Exit Function
fallo:
    MsgBox "Error al obtener el tipo de documento de la muestra.", vbCritical, Err.Description
End Function
Public Function recalcular_precios_muestras_sin_facturar() As Boolean

   On Error GoTo recalcular_precios_muestras_sin_facturar_Error
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "select * from muestras where documento_pago = 0 and anulada = 0 and precio_fijado = 0"
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        Do
            Me.informar_precio_muestra rs("id_muestra")
            rs.MoveNext
        Loop Until rs.EOF
    End If
    recalcular_precios_muestras_sin_facturar = True
   On Error GoTo 0
   Exit Function

recalcular_precios_muestras_sin_facturar_Error:
    recalcular_precios_muestras_sin_facturar = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure recalcular_precios_muestras_sin_facturar of Módulo de clase clsMuestra"
End Function
'M1051-I
'Public Sub Imprimir_listadomuestras_pendientes(fecha_desde As String, fecha_hasta As String, cliente As Long, cerradas As Boolean, INCLUYE_FECHAS As Boolean)
'    On Error GoTo fallo
'    Dim criterio As String
'    Dim fh As Date, fd As Date, todos_los_clientes As String
'    Dim objfrm As New frmReport
'    Dim strCad As String
'    Dim arrNom() As String
'    Dim arrVal() As String
'
'    If cerradas = False Then criterio = " {muestras.cerrada} = 1"
'
'    If INCLUYE_FECHAS Then
'        fd = CDate(fecha_desde)
'        fh = CDate(fecha_hasta)
'        criterio = criterio & " AND {muestras.fecha_recepcion} >= date(" & Year(fd) & "," & Month(fd) & " ," & Day(fd) & ")"
'        criterio = criterio & " AND {muestras.fecha_recepcion} <= date(" & Year(fh) & "," & Month(fh) & " ," & Day(fh) & ")"
'    End If
'
'    If cliente <> 0 Then
'        criterio = criterio & " and {muestras.cliente_id} = " & CStr(cliente)
'    End If
'
'    criterio = criterio & " AND {muestras.DOCUMENTO_PAGO} = 0"
'    criterio = criterio & " AND (ISNULL({muestras.ANULADA}) OR {muestras.ANULADA}=0)"
'
'    With objfrm
'        .iniciar
'        .informe = "General\rptListadoMuestrasPendientes"
'
'        strCad = criterio
'
'
'        ReDim arrNom(3)
'        ReDim arrVal(3)
'
'        arrNom(1) = "fecha_desde"
'        arrVal(1) = Format(fecha_desde, "dd/mm/yyyy")
'
'        arrNom(2) = "fecha_hasta"
'        arrVal(2) = Format(fecha_hasta, "dd/mm/yyyy")
'
'        arrNom(3) = "todos_los_clientes"
'        arrVal(3) = IIf(cliente <> 0, "1", "0")
'
'
'        .ParametrosNombre = arrNom
'        .ParametrosValores = arrVal
'
'        .criterio = strCad
'
'        .imprimir = False
'        .generar
'        .Visible = True
'
'    End With
''    Unload objfrm
'    Set objfrm = Nothing
'
'
'Exit Sub
'fallo:
'    MsgBox "Se ha producido un error al imprimir las muestras pendientes (Imprimir_listadomuestras_pendientes.", vbCritical, Err.Description
'End Sub
'M1051-F
Public Function Muestras_pendientes_facturar(MUESTRA_ID As Long, fecha_desde As String, fecha_hasta As String, cliente As Long, cerradas As Boolean, pedidos As Long, INCLUYE_FECHAS As Boolean, NO_FACTURABLES As Boolean, AIRBUS As Boolean, IBERIA As Boolean) As ADODB.Recordset
    Dim consulta As String
    On Error GoTo fallo
    Dim strCerradas As String
    Dim strfechas As String
    If cerradas = False Then
        strCerradas = " AND mu.cerrada = 1 "
    End If
    Dim strPedidos As String
    If pedidos <> 0 Then
        strPedidos = " AND mu.pedido_id = " & pedidos
    End If
    Dim strCliente As String
    Dim strAirbus As String
    Dim strIberia As String
    If cliente <> 0 Then
        strCliente = " AND mu.cliente_id = " & cliente
    End If
    Dim strMuestra As String
    If MUESTRA_ID <> 0 Then
        strMuestra = " AND mu.id_muestra = " & MUESTRA_ID
    End If
    If INCLUYE_FECHAS Then
        strfechas = " AND mu.fecha_recepcion>='" & Format(fecha_desde, "yyyy-mm-dd") & "' " & _
                     " AND mu.fecha_recepcion<='" & Format(fecha_hasta, "yyyy-mm-dd") & "' "

    End If
    Dim strNF As String
    If NO_FACTURABLES Then
      strNF = " AND (mu.documento_pago=0 or mu.documento_pago=99) "
    Else
      strNF = " AND mu.documento_pago=0 "
    End If
    If AIRBUS Then
        strAirbus = " AND cl.AIRBUS = 1 "
    End If
    If IBERIA Then
        strIberia = " AND cl.IBERIA = 1 "
    End If
    consulta = "SELECT cl.id_cliente, concat(tm.codigo,'-',CAST(mu.id_particular AS CHAR)),cl.nombre,mu.tipo_analisis_id,mu.referencia_cliente, " & _
                    "mu.fecha_recepcion,mu.id_general,mu.precio,mu.id_muestra,cl.factura_determinaciones, " & _
                    "ta.nombre, tm.id_tipo_muestra,cl.fp_id,ta.factura_determinaciones,mu.bano_id, " & _
                    "mu.analisis_modificado,cl.tarifa_id,cli_ped.id_pedido,concat(cli_ped.codigo,' (',cli_ped.descripcion,')'),b.FACTURA_DETERMINACIONES, " & _
                    "mu.enviado_correo,mu.anulada,mu.cerrada,mu.revision_usuario,f.nombre,mu.fecha_cierre,mu.urgente,mu.ajuste " & _
                   "FROM clientes as cl " & _
                   " INNER JOIN muestras as mu ON mu.cliente_id=cl.id_cliente" & _
                   " INNER JOIN tipos_muestra as tm ON mu.tipo_muestra_id=tm.id_tipo_muestra " & _
                   " INNER JOIN tipos_analisis as ta ON mu.tipo_analisis_id=ta.id_tipo_analisis " & _
                   " INNER JOIN clientes_pedidos cli_ped ON mu.pedido_id = cli_ped.id_pedido " & _
                   " INNER JOIN banos AS b ON mu.BANO_ID = b.ID_BANO " & _
                   " LEFT JOIN familias f ON tm.familia_id = f.ID_FAMILIA " & _
                   "WHERE  1 = 1 " & _
                   strfechas & strMuestra & strCliente & strCerradas & strPedidos & strNF & strAirbus & strIberia & _
                   " AND (mu.anulada Is Null or mu.anulada = 0) " & _
                   " order by cl.id_cliente, mu.id_muestra"
    Set Muestras_pendientes_facturar = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Se ha producido un error al buscar las muestras (Muestras_pendientes_facturar.", vbCritical, Err.Description
End Function

Public Function Informar_Datos_documento(muestra As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
'    consulta = "UPDATE MUESTRAS " & _
'               "   SET DOCUMENTO_PAGO=" & DOCUMENTO_PAGO & _
'               "     , PRECIO = '" & PRECIO & "'" & _
'               "     , CLIENTE_ID = " & CLIENTE_ID & _
'               "     , PEDIDO_ID = " & PEDIDO_ID & _
'               " WHERE ID_MUESTRA=" & MUESTRA
    consulta = "UPDATE MUESTRAS " & _
               "   SET DOCUMENTO_PAGO=" & DOCUMENTO_PAGO & _
               "     , PRECIO = '" & PRECIO & "'" & _
               "     , PEDIDO_ID = " & PEDIDO_ID & _
               " WHERE ID_MUESTRA=" & muestra
    execute_bd consulta
    Informar_Datos_documento = True
    Exit Function
fallo:
    Informar_Datos_documento = False
    MsgBox "Error al modificar el documento de pago de la muestra (Informar_Datos_documento)", vbCritical, Err.Description
End Function

Public Function Informe_Recepcion(muestra As Long, imprimir As Boolean) As Boolean
    
   On Error GoTo Informe_Recepcion_Error
'    Dim sRecepcionCE As String
'    Dim sRecepcion As String
'    Dim s() As String
'    Dim oMuestra As New clsMuestra
'    Dim i As Integer
'    s = Split(MUESTRA, ",")
'    For i = LBound(s) To UBound(s)
'        If oMuestra.esControlEficacia(CLng(s(i))) Then
'            If sRecepcionCE <> "" Then sRecepcionCE = sRecepcionCE & ","
'            sRecepcionCE = sRecepcionCE & s(i)
'        Else
'            If sRecepcion <> "" Then sRecepcion = sRecepcion & ","
'            sRecepcion = sRecepcion & s(i)
'        End If
'    Next
    
    ' CE
'    If sRecepcionCE <> "" Then
        Dim objrep As New frmReport
        With objrep
            .iniciar
            .criterio = "{muestras.ID_MUESTRA} = " & muestra
            If Me.esControlEficacia(muestra) Then
                .informe = "\Informes\rptRecepcion_CE"
            Else
                .informe = "\Informes\rptRecepcion"
            End If
            .imprimir = imprimir
            .pdf = ""
            .generar
            If Not imprimir Then
                objrep.Show vbModal
                Unload objrep
            End If
            Set objrep = Nothing
        End With
'    End If
    ' MUESTRAS
'    If sRecepcion <> "" Then
'        Dim objrep2 As New frmReport
'        With objrep2
'            .iniciar
'            .CRITERIO = "{muestras.ID_MUESTRA} IN [" & sRecepcion & "]"
'            .informe = "\Informes\rptRecepcion"
'            .imprimir = imprimir
'            .PDF = ""
'            .generar
'            If Not imprimir Then
'                objrep2.Show vbModal
'                Unload objrep2
'            End If
'            Set objrep2 = Nothing
'        End With
'    End If
    

   On Error GoTo 0
   Exit Function

Informe_Recepcion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Informe_Recepcion of Módulo de clase clsMuestra"
End Function
Public Function Genera_Informe(muestra As Long, pdf As String, imprimir As Boolean, parametros As String, tipo As Integer) As Boolean
'    Dim objrep As New frmReport
    Dim objrep As New frmReportSinVisor
   On Error GoTo Informe_Recepcion_Error
    Dim ret As Boolean
    ret = False
    
    With objrep
        .iniciar
        .criterio = "{muestras.ID_MUESTRA}=" & muestra & " " & parametros
        .informe = Me.obtener_plantilla(muestra)
        ' RTM
        If .informe = "Informes\rptRTM" Then
            ' Si hay alguna determinacion fuera de rango, genero el informe _NC
            Dim c As String
            c = "SELECT ID_DETERMINACION FROM determinaciones D " & _
                " WHERE D.MUESTRA_ID=" & muestra & _
                "   AND D.SITUACION = " & C_SITUACION.S_FUERA_RANGO
            Dim rs As ADODB.Recordset
            Set rs = datos_bd(c)
            If rs.RecordCount > 0 Then
                .informe = "Informes\rptRTM_NC"
            End If
        End If
        'M0687-I
        ' Si es de TIPO = 4 (Web MTQM)
        If .informe = "Informes\rptBano" Or .informe = "Informes\rptCO" Or .informe = "Informes\rptFluido" Or .informe = "Informes\rptFluidoGrado" Or .informe = "Informes\rptCOAgrupado" Or .informe = "Informes\rptSellante" Then
            Dim arrNom() As String
            Dim arrVal() As String
            ReDim arrNom(1)
            ReDim arrVal(1)
            arrNom(1) = "MTQM"
            If tipo = 4 Then
                arrVal(1) = "S"
            Else
                arrVal(1) = "N"
            End If
            .ParametrosNombre = arrNom
            .ParametrosValores = arrVal
        End If
        'M0687-F
        .imprimir = imprimir
        .pdf = pdf
        ret = .generar
'        If Not imprimir Then
'            objrep.Show vbModal
'        End If
    End With
    Unload objrep
    Set objrep = Nothing
    Genera_Informe = ret
   On Error GoTo 0
   Exit Function

Informe_Recepcion_Error:
    Genera_Informe = False
    enviar_informe_error muestra, "Genera_Informe : " & "Error " & Err.Number & " (" & Err.Description & ") in procedure Genera_Informe of Módulo de clase clsMuestra"
End Function

Public Function Informe_Recepcion_multiple(LISTA_MUESTRAS As String) As Boolean
    Dim objrep As New frmReport
   On Error GoTo Informe_Recepcion_multiple_Error

    With objrep
        .iniciar
        .criterio = "{muestras.ID_MUESTRA} in [" & LISTA_MUESTRAS & "]"
        .informe = "\Informes\rptRecepcion"
        .imprimir = False
        .pdf = ""
        .generar
        objrep.Show vbModal
        Unload objrep
        Set objrep = Nothing
    End With

   On Error GoTo 0
   Exit Function

Informe_Recepcion_multiple_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Informe_Recepcion_multiple of Módulo de clase clsMuestra"

End Function

'Public Function Listado_ediciones(MUESTRA As Long) As ADODB.Recordset
'    On Error GoTo fallo
'    Dim consulta As String
'    consulta = "SELECT * " & _
'               "  FROM MUESTRAS_NUEVA_EDICION " & _
'               " WHERE MUESTRA_ID = " & MUESTRA & _
'               " ORDER BY EDICION DESC"
'    Set Listado_ediciones = datos_bd(consulta)
'    Exit Function
'fallo:
'    MsgBox "Error al obtener el Listado_ediciones.", vbCritical, Err.Description
'End Function

Public Function esCerrada(codmuestra As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    Dim rs As ADODB.Recordset
    consulta = "select cerrada " & _
               "  from muestras " & _
               " where id_muestra = " & codmuestra
    Set rs = datos_bd(consulta)
    esCerrada = True
    If rs.RecordCount > 0 Then
        If rs(0) = 0 Then
            esCerrada = False
        Else
            esCerrada = True
        End If
    End If
    Exit Function
fallo:
    esCerrada = False
    MsgBox "Error al esCerrada de muestra.", vbCritical, Err.Description
End Function

Public Function imprimir_listadomuestras(ByVal criterio As String, ByVal fecha_desde As String, ByVal fecha_hasta As String, ByVal todos_clientes As Boolean, ByVal todas_muestras As Boolean, Optional ByVal titulo_informe As String = "Listado de Muestras")

    Dim objfrm As New frmReport
    Dim strCad As String
    Dim arrNom() As String
    Dim arrVal() As String
    
    With objfrm
        .iniciar
        .informe = "General\rptListadoMuestras"
        
        strCad = criterio
        
        
        ReDim arrNom(5)
        ReDim arrVal(5)
        
        arrNom(1) = "fecha_desde"
        arrVal(1) = fecha_desde
        
        arrNom(2) = "fecha_hasta"
        arrVal(2) = fecha_hasta
        
        arrNom(3) = "todos_los_clientes"
        arrVal(3) = IIf(todos_clientes, "1", "0")
        
        arrNom(4) = "todas_las_muestras"
        arrVal(4) = IIf(todas_muestras, "1", "0")
        
        arrNom(5) = "titulo"
        arrVal(5) = titulo_informe
                
        .ParametrosNombre = arrNom
        .ParametrosValores = arrVal
        
        .criterio = strCad
        
        .imprimir = False
        .generar
        .visible = True
        
    End With
    Set objfrm = Nothing


End Function
Public Function CerrarSinInforme(muestra As Long) As Boolean
    Dim consulta As String
   On Error GoTo CerrarSinInforme_Error

    consulta = "UPDATE MUESTRAS SET" & _
               " CERRADA=3,FECHA_CIERRE ='" & Format(Date, "yyyy-mm-dd") & "'," & _
               " HORA_CIERRE ='" & Format(Time, "hh:mm:ss") & "'," & _
               " CERRADA_USUARIO = " & USUARIO.getID_EMPLEADO & _
               " WHERE id_muestra=" & muestra
    execute_bd consulta
    
    Me.informar_situacion muestra
    CerrarSinInforme = True
   On Error GoTo 0
   Exit Function

CerrarSinInforme_Error:
    CerrarSinInforme = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CerrarSinInforme of Módulo de clase clsMuestra"
End Function
Public Function Revisar(muestra As Long) As Boolean
    Dim consulta As String
   On Error GoTo CerrarSinInforme_Error

    consulta = "UPDATE MUESTRAS " & _
               "   SET REVISION_USUARIO = " & USUARIO.getID_EMPLEADO & _
               " WHERE ID_MUESTRA=" & muestra
    execute_bd consulta
    'M1054-I
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra muestra
    If oMuestra.getINFORME_MANUAL = 0 Then
        imprimir muestra, 2, False
    Else
        imprimir muestra, 70, False
    End If
    'M1054-F
    
    Revisar = True
   On Error GoTo 0
   Exit Function

CerrarSinInforme_Error:
    Revisar = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Revisar of Módulo de clase clsMuestra"
End Function
Public Function listadoMuestrasFueraPlazo(Centro As String) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    Dim parametros As New clsParametros
    Dim clientes As String
    parametros.Carga CLIENTES_MUESTRAS_PLAZOS, ""
    clientes = parametros.getVALOR
    
    ' TM QUE NO NECESITAN REVISION
    parametros.Carga PARAM_TM_NO_FUERA_PLAZO, ""
    Dim listaTM As String
    If Trim(parametros.getVALOR) <> "" Then
        listaTM = " and M.tipo_muestra_id not in (" & parametros.getVALOR & ") "
    End If
    
    
    consulta = "SELECT M.ID_MUESTRA,M.ID_GENERAL,CONCAT(TM.CODIGO,'-',CAST(M.ID_PARTICULAR AS CHAR)) AS CODIGO, " & _
               "       C.NOMBRE,M.REFERENCIA_CLIENTE,M.FECHA_RECEPCION,M.FECHA_PREV_FIN,DATEDIFF(CURRENT_DATE,M.FECHA_PREV_FIN),M.IPA " & _
               "  FROM MUESTRAS M, CLIENTES C,TIPOS_MUESTRA TM " & _
               " where M.CLIENTE_ID = C.ID_CLIENTE " & _
               "   AND M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
               "   AND M.ANULADA = 0 " & _
               "   AND M.CERRADA = 0 " & _
               IIf(Centro <> "", " AND M.CENTRO_ID = " & Centro, "") & _
               "   AND M.ANNO >= YEAR(CURRENT_DATE) - 1 " & _
               "   AND M.FECHA_PREV_FIN <= CURRENT_DATE " & _
               IIf(clientes <> "", " AND M.CLIENTE_ID NOT IN (" & clientes & ")", "") & _
               listaTM & _
               " ORDER BY DATEDIFF(CURRENT_DATE,M.FECHA_PREV_FIN) DESC"
    Set listadoMuestrasFueraPlazo = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al obtener el listadoMuestrasFueraPlazo.", vbCritical, Err.Description
End Function
Public Function listadoMuestrasProximasPlazo(DIAS As Integer) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT M.ID_MUESTRA,M.ID_GENERAL,CONCAT(TM.CODIGO,'-',CAST(M.ID_PARTICULAR AS CHAR)) AS CODIGO," & _
               "       C.NOMBRE,M.REFERENCIA_CLIENTE,M.FECHA_RECEPCION,M.FECHA_PREV_FIN,DATEDIFF(CURRENT_DATE,M.FECHA_PREV_FIN),M.IPA " & _
               "  FROM MUESTRAS M, CLIENTES C,TIPOS_MUESTRA TM " & _
               " where M.CLIENTE_ID = C.ID_CLIENTE " & _
               "   AND M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
               "   AND M.ANULADA = 0 " & _
               "   AND M.CERRADA = 0 " & _
               "   AND M.ANNO >= 2012 " & _
               "   AND M.FECHA_PREV_FIN > CURRENT_DATE " & _
               "   AND M.FECHA_PREV_FIN <= DATE_ADD(CURRENT_DATE ,INTERVAL " & DIAS & "  DAY) " & _
               " ORDER BY DATEDIFF(CURRENT_DATE,M.FECHA_PREV_FIN) DESC"
    Set listadoMuestrasProximasPlazo = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al obtener el listadoMuestrasProximasPlazo.", vbCritical, Err.Description
End Function
'M1278-I
Public Function listadoFacturacionSectores(familia As Long, SECTOR As Long, tipo As String, fechaI As String, fechaF As String, cliente As Long, AIRBUS As Boolean) As ADODB.Recordset
On Error GoTo fallo
    Dim consulta As String
    Dim strFamilia1 As String
    Dim strFamilia2 As String
    Dim strSECTOR As String
    Dim strTipo As String
    Dim filtro As String
    
    If familia <> 0 Then
        strFamilia1 = " AND F.ID_FAMILIA = " & familia
'        strFamilia2 = " AND C.ID_FAMILIA = " & familia
    End If
    
    If SECTOR <> 0 Then
        strSECTOR = " AND S.ID_SECTOR = " & SECTOR
    End If
    If cliente <> 0 Then
        filtro = filtro & " AND CLIE.ID_CLIENTE = " & cliente
    End If
    If AIRBUS Then
        filtro = filtro & " AND CLIE.AIRBUS = 1 "
    End If
    
    'JOIN PARA LOS ALBARANES DE TCTSKY
    Dim join As String
    join = " SELECT DISTINCT(A.ALBARAN_ID) " & _
           " FROM DOCS_PAGO_CONCEPTOS A, DOCS_PAGO B, FAMILIAS F, SECTORES S, CLIENTES CLIE " & _
           " WHERE a.ALBARAN_ID <> 0 " & _
           "   AND A.DOC_ID = B.ID_DOC " & _
           "   AND A.FAMILIA_ID = F.ID_FAMILIA " & _
           "   AND F.SECTOR_ID = S.ID_SECTOR " & _
           "   AND B.CLIENTE_ID_FACTURA = CLIE.ID_CLIENTE " & _
           "   AND B.ANULADO = 0 " & _
           "   AND B.FECHA_FACTURA >= '" & fechaI & "' and B.FECHA_FACTURA <= '" & fechaF & "'" & _
           "   AND CLIE.IBERIA = 0 " & _
           strFamilia1 & strSECTOR & filtro
    'MDET
    consulta = "SELECT AA.C1 AS SECTOR, AA.C2 AS FAMILIA, SUM(AA.C3) AS SUMA, GROUP_CONCAT(DISTINCT AA.DOCS ORDER BY 1) AS DOCS FROM ( "
    consulta = consulta & " SELECT DP.NUMERO AS C4, S.NOMBRE AS C1, F.NOMBRE AS C2, SUM(DPM.PRECIO) - ((sum(DPM.PRECIO) * DP.descuento) / 100) AS C3, DP.TIPO AS TIPO, GROUP_CONCAT(DISTINCT DP.ID_DOC) AS DOCS" & _
               "   FROM MUESTRAS M, TIPOS_MUESTRA TM, SECTORES S, DOCS_PAGO_MUESTRAS DPM, DOCS_PAGO DP, FAMILIAS F, CLIENTES CLIE" & _
               "  WHERE M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA" & _
               "    AND DP.FECHA_FACTURA >= '" & fechaI & "' and DP.FECHA_FACTURA <= '" & fechaF & "'" & _
               "    AND TM.SECTOR_ID = S.ID_SECTOR" & _
               "    AND TM.FAMILIA_ID = F.ID_FAMILIA" & _
               "    AND DPM.MUESTRA_ID = M.ID_MUESTRA AND DPM.DETERMINACION_ID = 0 " & _
               "    AND DP.ID_DOC = DPM.DOC_ID" & _
               "    AND DP.CLIENTE_ID_FACTURA = CLIE.ID_CLIENTE " & _
               "    AND DP.TIPO IN " & tipo & _
               "    AND DPM.MUESTRA_ID <> 0 " & strFamilia1 & strSECTOR & filtro & _
               "   AND CLIE.IBERIA = 0 " & _
               "  GROUP BY DP.NUMERO,S.NOMBRE, F.NOMBRE"
    ' Incluir apartado = 0 y suma del total no del precio
    consulta = consulta & " UNION " & _
               " select a.NUMERO AS C4, s.nombre AS C1,F.NOMBRE AS C2, sum(b.total) - ((sum(b.total) * a.descuento) / 100) AS C3, a.TIPO as TIPO2, GROUP_CONCAT(DISTINCT a.ID_DOC) AS DOCS " & _
               "   from docs_pago a, docs_pago_conceptos b, familias F, SECTORES s, CLIENTES CLIE" & _
               "  where a.ID_DOC = b.DOC_ID" & _
               "    AND b.FAMILIA_ID = F.ID_FAMILIA" & _
               "    AND F.SECTOR_ID = s.ID_SECTOR" & _
               "    AND a.CLIENTE_ID_FACTURA = CLIE.ID_CLIENTE " & _
               "    AND a.ANULADO <> 1 and a.TIPO IN " & tipo & _
               "    AND b.ALBARAN_ID = 0 " & _
               "    AND b.APARTADO = 0 " & _
               "    AND a.FECHA_FACTURA >= '" & fechaI & "' and a.FECHA_FACTURA <= '" & fechaF & "' " & strFamilia1 & strSECTOR & filtro & _
               "    AND CLIE.IBERIA = 0 " & _
               " group by a.NUMERO, s.nombre,F.NOMBRE"
    ' Conceptos de las facturas de TCTSKY
    ' " SELECT a.NUMERO AS C4, s.nombre AS C1,F.NOMBRE AS C2, SUM(b.PRECIO) - ((SUM(b.PRECIO) * a.descuento) / 100) AS C3, a.TIPO AS TIPO2, GROUP_CONCAT(DISTINCT a.ID_DOC) AS DOCS "
    consulta = consulta & " UNION " & _
               " SELECT a.NUMERO AS C4, s.nombre AS C1,F.NOMBRE AS C2, SUM(b.total) - ((SUM(b.total) * a.descuento) / 100) AS C3, a.TIPO AS TIPO2, GROUP_CONCAT(DISTINCT a.ID_DOC) AS DOCS " & _
               "   FROM docs_pago a, docs_pago_conceptos b, familias F, SECTORES s, CLIENTES CLIE " & _
               "  where a.ID_DOC = b.DOC_ID And b.FAMILIA_ID = F.ID_FAMILIA And F.SECTOR_ID = s.ID_SECTOR " & _
               "    AND a.CLIENTE_ID_FACTURA = CLIE.ID_CLIENTE AND a.ANULADO <> 1 " & _
               "    AND a.ID_DOC IN (" & join & ") " & _
               "    AND b.APARTADO = 0 " & _
               "   AND CLIE.IBERIA = 0 " & _
               "  GROUP BY a.NUMERO, s.nombre,F.NOMBRE "
'               "    AND a.FECHA_FACTURA >= '" & fechaI & "' and a.FECHA_FACTURA <= '" & fechaF & "' "
    ' Muestras de las facturas de TCTSKY
    'MDET
    consulta = consulta & " UNION " & _
               " SELECT DP.NUMERO AS C4, S.NOMBRE AS C1, F.NOMBRE AS C2, SUM(DPM.PRECIO) - ((SUM(DPM.PRECIO) * DP.descuento) / 100) AS C3, DP.TIPO AS TIPO, GROUP_CONCAT(DISTINCT DP.ID_DOC) AS DOCS " & _
               "   FROM MUESTRAS M, TIPOS_MUESTRA TM, SECTORES S, DOCS_PAGO_MUESTRAS DPM, DOCS_PAGO DP, FAMILIAS F, CLIENTES CLIE " & _
               "  WHERE m.TIPO_MUESTRA_ID = tm.id_tipo_muestra " & _
               "    AND TM.SECTOR_ID = S.ID_SECTOR " & _
               "    AND TM.FAMILIA_ID = F.ID_FAMILIA " & _
               "    AND DPM.MUESTRA_ID = M.ID_MUESTRA AND DPM.DETERMINACION_ID = 0 " & _
               "    AND DP.ID_DOC = DPM.DOC_ID " & _
               "    AND DP.CLIENTE_ID_FACTURA = CLIE.ID_CLIENTE " & _
               "    AND DP.TIPO = 1 " & _
               "    AND DPM.MUESTRA_ID <> 0 " & _
               "    AND DP.ID_DOC IN (" & join & ") " & _
               "    AND CLIE.IBERIA = 0 " & _
               "  GROUP BY DP.NUMERO,S.NOMBRE, F.NOMBRE "
'               "    AND DP.FECHA_FACTURA >= '" & fechaI & "' and DP.FECHA_FACTURA <= '" & fechaF & "' "
    ' IBERIA
    If strFamilia1 = "" Then
        consulta = consulta & " UNION " & _
                   " SELECT a.NUMERO AS C4, 'AERONAUTICO' AS C1,'IBERIA' AS C2, SUM(a.total) - ((SUM(a.total) * a.descuento) / 100) AS C3, a.TIPO AS TIPO2, GROUP_CONCAT(DISTINCT a.ID_DOC) AS DOCS " & _
                   "   FROM docs_pago a, CLIENTES CLIE " & _
                   "  WHERE a.CLIENTE_ID_FACTURA = CLIE.id_cliente " & _
                   "    AND a.FECHA_FACTURA >= '" & fechaI & "' and a.FECHA_FACTURA <= '" & fechaF & "' " & _
                   "    AND CLIE.IBERIA = 1 " & _
                   "    AND a.TIPO IN " & tipo & _
                   filtro & _
                   "  GROUP BY a.numero "
    '               strFamilia1 & strSector & filtro
    End If
    consulta = consulta & " ) AS AA " & _
               "  GROUP BY AA.C1, AA.C2     "
               
    Set listadoFacturacionSectores = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al obtener el listadoFacturacionSectores.", vbCritical, Err.Description
End Function
Public Function agruparMuestra(ID_MUESTRA As Long, lista As String) As Boolean
    Dim c As String
    ' Limpiar las agrupadas
   On Error GoTo agruparMuestra_Error

    c = "UPDATE muestras SET AGRUPADA = 0,AGRUPADA_MUESTRA_ID = 0 WHERE ID_MUESTRA = " & ID_MUESTRA & " OR AGRUPADA_MUESTRA_ID = " & ID_MUESTRA
    execute_bd c
    ' Asignar las nuevas
    If lista <> "" Then
        c = "UPDATE muestras SET AGRUPADA = 1,AGRUPADA_MUESTRA_ID = " & ID_MUESTRA & " where ID_MUESTRA IN (" & lista & ")"
        execute_bd c
        c = "UPDATE muestras SET AGRUPADA = 1,AGRUPADA_MUESTRA_ID = 0 where ID_MUESTRA = " & ID_MUESTRA
        execute_bd c
    End If
   On Error GoTo 0
   Exit Function

agruparMuestra_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure agruparMuestra of Módulo de clase clsMuestra"
End Function

Public Function agrupadasListadoMuestra(ID_MUESTRA As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT M.ID_MUESTRA,CONCAT(TM.CODIGO,'-',CAST(M.ID_PARTICULAR AS CHAR)) AS CODIGO " & _
               "      ,M.REFERENCIA_CLIENTE,M.ID_GENERAL " & _
               "  FROM MUESTRAS M, TIPOS_MUESTRA TM " & _
               " WHERE M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
               "   AND M.AGRUPADA_MUESTRA_ID = " & ID_MUESTRA & _
               " ORDER BY ID_MUESTRA "
    Set agrupadasListadoMuestra = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al obtener el agrupadasListadoMuestra.", vbCritical, Err.Description
End Function
Public Function agrupadasListadoMuestraNoAsignadas(ID_MUESTRA As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra ID_MUESTRA
    consulta = "SELECT M.ID_MUESTRA,CONCAT(TM.CODIGO,'-',CAST(M.ID_PARTICULAR AS CHAR)) AS CODIGO " & _
               "      ,M.REFERENCIA_CLIENTE,M.ID_GENERAL " & _
               "  FROM MUESTRAS M, TIPOS_MUESTRA TM " & _
               " WHERE M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
               "   AND M.AGRUPADA = 0 " & _
               "   AND M.CLIENTE_ID = " & oMuestra.getCLIENTE_ID & _
               "   AND M.TIPO_MUESTRA_ID = " & oMuestra.getTIPO_MUESTRA_ID & _
               "   AND M.TIPO_ANALISIS_ID = " & oMuestra.getTIPO_ANALISIS_ID & _
               "   AND M.FECHA_RECEPCION = '" & Format(oMuestra.getFECHA_RECEPCION, "yyyy-mm-dd") & "'" & _
               "   AND M.ANULADA = 0 " & _
               "   AND M.ID_MUESTRA <> " & ID_MUESTRA & _
               " ORDER BY ID_MUESTRA "
    Set agrupadasListadoMuestraNoAsignadas = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al obtener el agrupadasListadoMuestraNoAsignadas.", vbCritical, Err.Description
End Function

Public Function MuestrasEnFactura(ID_DOC As String, ID_MUESTRA As Long, CLIENTE_ID As Long, TIPO_ANALISIS_ID As Long, bPlant As Boolean, bEnsayo As Boolean, bPrograma As Boolean, bSection As Boolean, bFluid As Boolean, bFacility As Boolean) As ADODB.Recordset
    Dim consulta As String
    Dim rs As ADODB.Recordset
   On Error GoTo MuestrasEnFactura_Error

    consulta = "SELECT GROUP_CONCAT(DISTINCT ALBARAN_ID) FROM DOCS_PAGO_CONCEPTOS " & _
               " WHERE DOC_ID = " & ID_DOC & _
               "   AND ALBARAN_ID <> 0;"
    Set rs = datos_bd(consulta)
    Dim docs As String
    If IsNull(rs(0)) Then
        docs = ID_DOC
    Else
        docs = rs(0)
    End If
    Dim filtro As String
    If ID_MUESTRA <> 0 Then
        filtro = filtro & " and mu.id_muestra = " & ID_MUESTRA
    End If
    If CLIENTE_ID <> 0 Then
        filtro = filtro & " and mu.cliente_id = " & CLIENTE_ID
    End If
    If TIPO_ANALISIS_ID <> 0 Then
        filtro = filtro & " and mu.tipo_analisis_id = " & TIPO_ANALISIS_ID
    End If
    If bPlant = True Then
        filtro = filtro & " and cl.PLANT_ID = 0"
    End If
    If bEnsayo = True Then
        filtro = filtro & " and isnull(d1.valor)"
    End If
    If bPrograma = True Then
        filtro = filtro & " and isnull(d2.valor)"
    End If
    If bSection = True Then
        filtro = filtro & " and isnull(d3.valor)"
    End If
    If bFluid = True Then
        filtro = filtro & " and isnull(d4.valor)"
    End If
    If bFacility = True Then
        filtro = filtro & " and isnull(d5.valor)"
    End If
    'MDET
    consulta = "SELECT mu.id_muestra, concat(tm.codigo,'-',cast(mu.id_particular as char)), cl.nombre,ta.nombre,mu.referencia_cliente,mu.id_general " & _
               "      ,d0.DESCRIPCION,d1.DESCRIPCION,d2.DESCRIPCION,d3.DESCRIPCION,d4.DESCRIPCION,d5.DESCRIPCION " & _
               " FROM muestras as mu  " & _
               " INNER JOIN tipos_muestra as tm ON mu.tipo_muestra_id=tm.id_tipo_muestra  " & _
               " INNER JOIN tipos_analisis as ta ON mu.tipo_analisis_id=ta.id_tipo_analisis " & _
               " INNER JOIN centros as ce ON mu.centro_id = ce.id_centro " & _
               " INNER JOIN clientes as cl ON mu.cliente_id=cl.id_cliente " & _
               " INNER JOIN docs_pago_muestras dpm ON mu.id_muestra = dpm.muestra_id and dpm.determinacion_id = 0 " & _
               " INNER JOIN docs_pago dp ON dpm.DOC_ID = dp.ID_DOC " & _
               " LEFT JOIN muestras_airbus ma ON mu.id_muestra = ma.muestra_id " & _
               " LEFT JOIN decodificadora d0 ON d0.codigo = " & DECODIFICADORA.AIRBUS_PLANT & " and d0.valor = cl.PLANT_ID " & _
               " LEFT JOIN decodificadora d1 ON d1.codigo = " & DECODIFICADORA.AIRBUS_TIPOS_ENSAYOS & " and d1.valor = ma.ENSAYO_ID and d1.parametros = cl.PLANT_ID " & _
               " LEFT JOIN decodificadora d2 ON d2.codigo = " & DECODIFICADORA.AIRBUS_PROGRAMAS & " and d2.valor = ma.PROGRAMA_ID and d2.parametros = cl.PLANT_ID " & _
               " LEFT JOIN decodificadora d3 ON d3.codigo = " & DECODIFICADORA.AIRBUS_SECTION & " and d3.valor = ma.SECTION_ID and d3.parametros = cl.PLANT_ID " & _
               " LEFT JOIN decodificadora d4 ON d4.codigo = " & DECODIFICADORA.AIRBUS_FLUID & " and d4.valor = ma.FLUID_ID and d4.parametros = cl.PLANT_ID " & _
               " LEFT JOIN decodificadora d5 ON d5.codigo = " & DECODIFICADORA.AIRBUS_FACILITY & " and d5.valor = ma.FACILITY_ID and d5.parametros = cl.PLANT_ID " & _
               " WHERE dpm.muestra_id <> 0 " & _
               "   AND dp.id_doc in (" & docs & ") " & _
               filtro & _
               " ORDER BY dp.ID_DOC, dpm.MUESTRA_ID asc"
    Set MuestrasEnFactura = datos_bd(consulta)

   On Error GoTo 0
   Exit Function

MuestrasEnFactura_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MuestrasEnFactura of Módulo de clase clsMuestra"
End Function

Public Function MuestrasAirbus(ID_MUESTRA As String, CLIENTE_ID As Long, TIPO_ANALISIS_ID As Long, bPlant As Boolean, bEnsayo As Boolean, bPrograma As Boolean, bSection As Boolean, bFluid As Boolean, bFacility As Boolean) As ADODB.Recordset
    Dim consulta As String
    Dim rs As ADODB.Recordset
    Dim filtro As String
    If ID_MUESTRA <> "" Then
        filtro = filtro & " and mu.id_muestra in (" & ID_MUESTRA & ")"
    End If
    If CLIENTE_ID <> 0 Then
        filtro = filtro & " and mu.cliente_id = " & CLIENTE_ID
    End If
    If TIPO_ANALISIS_ID <> 0 Then
        filtro = filtro & " and mu.tipo_analisis_id = " & TIPO_ANALISIS_ID
    End If
    If bPlant = True Then
        filtro = filtro & " and cl.PLANT_ID = 0"
    End If
    If bEnsayo = True Then
        filtro = filtro & " and isnull(d1.valor)"
    End If
    If bPrograma = True Then
        filtro = filtro & " and isnull(d2.valor)"
    End If
    If bSection = True Then
        filtro = filtro & " and isnull(d3.valor)"
    End If
    If bFluid = True Then
        filtro = filtro & " and isnull(d4.valor)"
    End If
    If bFacility = True Then
        filtro = filtro & " and isnull(d5.valor)"
    End If
    
    consulta = "SELECT mu.id_muestra, concat(tm.codigo,'-',cast(mu.id_particular as char)), cl.nombre,ta.nombre,mu.referencia_cliente,mu.id_general " & _
               "      ,d0.DESCRIPCION,d1.DESCRIPCION,d2.DESCRIPCION,d3.DESCRIPCION,d4.DESCRIPCION,d5.DESCRIPCION " & _
               " FROM muestras as mu  " & _
               " INNER JOIN tipos_muestra as tm ON mu.tipo_muestra_id=tm.id_tipo_muestra  " & _
               " INNER JOIN tipos_analisis as ta ON mu.tipo_analisis_id=ta.id_tipo_analisis " & _
               " INNER JOIN centros as ce ON mu.centro_id = ce.id_centro " & _
               " INNER JOIN clientes as cl ON mu.cliente_id=cl.id_cliente " & _
               " LEFT JOIN muestras_airbus ma ON mu.id_muestra = ma.muestra_id " & _
               " LEFT JOIN decodificadora d0 ON d0.codigo = " & DECODIFICADORA.AIRBUS_PLANT & " and d0.valor = cl.PLANT_ID " & _
               " LEFT JOIN decodificadora d1 ON d1.codigo = " & DECODIFICADORA.AIRBUS_TIPOS_ENSAYOS & " and d1.valor = ma.ENSAYO_ID and d1.parametros = cl.PLANT_ID " & _
               " LEFT JOIN decodificadora d2 ON d2.codigo = " & DECODIFICADORA.AIRBUS_PROGRAMAS & " and d2.valor = ma.PROGRAMA_ID and d2.parametros = cl.PLANT_ID " & _
               " LEFT JOIN decodificadora d3 ON d3.codigo = " & DECODIFICADORA.AIRBUS_SECTION & " and d3.valor = ma.SECTION_ID and d3.parametros = cl.PLANT_ID " & _
               " LEFT JOIN decodificadora d4 ON d4.codigo = " & DECODIFICADORA.AIRBUS_FLUID & " and d4.valor = ma.FLUID_ID and d4.parametros = cl.PLANT_ID " & _
               " LEFT JOIN decodificadora d5 ON d5.codigo = " & DECODIFICADORA.AIRBUS_FACILITY & " and d5.valor = ma.FACILITY_ID and d5.parametros = cl.PLANT_ID " & _
               " WHERE 1 = 1 " & _
               filtro & _
               " ORDER BY mu.id_muestra asc"
    Set MuestrasAirbus = datos_bd(consulta)

End Function


