VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMuestras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSMUESTRAS
'***************************************************************
Private ID_MUESTRA As Long
Private TIPO_MUESTRA_ID As Long
Private ID_GENERAL As Long
Private ID_PARTICULAR As Long
Private ANNO As Long
Private TIPO_ANALISIS_ID As Long
Private ANALISIS_MODIFICADO As Long
Private FECHA_MUESTREO As String
Private ENTIDAD_MUESTREO_ID As Long
Private DETALLE_MUESTREO As String
Private OBSERVACIONES_MUESTREO As String
Private FECHA_RECEPCION As String
Private HORA_RECEPCION As String
Private EMPLEADO_ID As Long
Private FORMATO_ID As Long
Private ENTIDAD_ENTREGA_ID As Long
Private DETALLE_ENTREGA As String
Private OBSERVACIONES_ENTREGA As String
Private CLIENTE_ID As Long
Private CENTRO_ID As Long
Private REPLACEMENT_ID As Long
Private REFERENCIA_CLIENTE As String
Private PRECIO As Long
Private FECHA_PREV_FIN As String
Private OBSERVACIONES As String
Private ANALISIS_DUPLICADO As Long
Private ANULADA As Long
Private PRECINTO As String
Private BANO_ID As Long
Private IPA As Long
Private INFORME_MANUAL As Long
Private FECHA_RECOGIDA As String
Private FECHA_COMIENZO As String
Private FECHA_FINALIZACION As String
Private FECHA_CIERRE As String
Private HORA_CIERRE As String
Private FECHA_ENVIO As String
Private MOTIVO_RETRASO_ID As Long
Private RETRASO_OBSERVACION As String
Private CERRADA As Long
Private REVISION_USUARIO As Long
Private CERRADA_USUARIO As Long
Private DOCUMENTO_PAGO As Long
Private IMPRESA As Long
Private ENVIADO_CORREO As Long
Private ULT_EDICION_IMP As Long
Private OP_VUELO As Long
Private OP_INSITU As Long
Private OP_LABMOVIL As Long
Private OP_NORUTINARIA As Long
Private OP_REPETICION As Long
Private SITUACION As Long
Private FIRMA As Long
Private PRODUCTO As String
Private RESPONSABLE_ID As Long
Private PEDIDO_ID As Long
Private OFERTA_ID As Long
Private PAQUETE_ID As Long
Private ENAC As Long
Private NADCAP As Long
Private URGENTE As Long
Private AJUSTE As Long
Private AGRUPADA As Long
Private AGRUPADA_MUESTRA_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSMUESTRAS
'***************************************************************
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
Public Property Let setANNO(ByVal dato As Long)
        ANNO = dato
End Property
Public Property Let setTIPO_ANALISIS_ID(ByVal dato As Long)
        TIPO_ANALISIS_ID = dato
End Property
Public Property Let setANALISIS_MODIFICADO(ByVal dato As Long)
        ANALISIS_MODIFICADO = dato
End Property
Public Property Let setFECHA_MUESTREO(ByVal dato As String)
        FECHA_MUESTREO = dato
End Property
Public Property Let setENTIDAD_MUESTREO_ID(ByVal dato As Long)
        ENTIDAD_MUESTREO_ID = dato
End Property
Public Property Let setDETALLE_MUESTREO(ByVal dato As String)
        DETALLE_MUESTREO = dato
End Property
Public Property Let setOBSERVACIONES_MUESTREO(ByVal dato As String)
        OBSERVACIONES_MUESTREO = dato
End Property
Public Property Let setFECHA_RECEPCION(ByVal dato As String)
        FECHA_RECEPCION = dato
End Property
Public Property Let setHORA_RECEPCION(ByVal dato As String)
        HORA_RECEPCION = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Long)
        EMPLEADO_ID = dato
End Property
Public Property Let setFORMATO_ID(ByVal dato As Long)
        FORMATO_ID = dato
End Property
Public Property Let setENTIDAD_ENTREGA_ID(ByVal dato As Long)
        ENTIDAD_ENTREGA_ID = dato
End Property
Public Property Let setDETALLE_ENTREGA(ByVal dato As String)
        DETALLE_ENTREGA = dato
End Property
Public Property Let setOBSERVACIONES_ENTREGA(ByVal dato As String)
        OBSERVACIONES_ENTREGA = dato
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
        REFERENCIA_CLIENTE = dato
End Property
Public Property Let setPRECIO(ByVal dato As Long)
        PRECIO = dato
End Property
Public Property Let setFECHA_PREV_FIN(ByVal dato As String)
        FECHA_PREV_FIN = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setANALISIS_DUPLICADO(ByVal dato As Long)
        ANALISIS_DUPLICADO = dato
End Property
Public Property Let setANULADA(ByVal dato As Long)
        ANULADA = dato
End Property
Public Property Let setPRECINTO(ByVal dato As String)
        PRECINTO = dato
End Property
Public Property Let setBANO_ID(ByVal dato As Long)
        BANO_ID = dato
End Property
Public Property Let setIPA(ByVal dato As Long)
        IPA = dato
End Property
Public Property Let setINFORME_MANUAL(ByVal dato As Long)
        INFORME_MANUAL = dato
End Property
Public Property Let setFECHA_RECOGIDA(ByVal dato As String)
        FECHA_RECOGIDA = dato
End Property
Public Property Let setFECHA_COMIENZO(ByVal dato As String)
        FECHA_COMIENZO = dato
End Property
Public Property Let setFECHA_FINALIZACION(ByVal dato As String)
        FECHA_FINALIZACION = dato
End Property
Public Property Let setFECHA_CIERRE(ByVal dato As String)
        FECHA_CIERRE = dato
End Property
Public Property Let setHORA_CIERRE(ByVal dato As String)
        HORA_CIERRE = dato
End Property
Public Property Let setFECHA_ENVIO(ByVal dato As String)
        FECHA_ENVIO = dato
End Property
Public Property Let setMOTIVO_RETRASO_ID(ByVal dato As Long)
        MOTIVO_RETRASO_ID = dato
End Property
Public Property Let setRETRASO_OBSERVACION(ByVal dato As String)
        RETRASO_OBSERVACION = dato
End Property
Public Property Let setCERRADA(ByVal dato As Long)
        CERRADA = dato
End Property
Public Property Let setREVISION_USUARIO(ByVal dato As Long)
        REVISION_USUARIO = dato
End Property
Public Property Let setCERRADA_USUARIO(ByVal dato As Long)
        CERRADA_USUARIO = dato
End Property
Public Property Let setDOCUMENTO_PAGO(ByVal dato As Long)
        DOCUMENTO_PAGO = dato
End Property
Public Property Let setIMPRESA(ByVal dato As Long)
        IMPRESA = dato
End Property
Public Property Let setENVIADO_CORREO(ByVal dato As Long)
        ENVIADO_CORREO = dato
End Property
Public Property Let setULT_EDICION_IMP(ByVal dato As Long)
        ULT_EDICION_IMP = dato
End Property
Public Property Let setOP_VUELO(ByVal dato As Long)
        OP_VUELO = dato
End Property
Public Property Let setOP_INSITU(ByVal dato As Long)
        OP_INSITU = dato
End Property
Public Property Let setOP_LABMOVIL(ByVal dato As Long)
        OP_LABMOVIL = dato
End Property
Public Property Let setOP_NORUTINARIA(ByVal dato As Long)
        OP_NORUTINARIA = dato
End Property
Public Property Let setOP_REPETICION(ByVal dato As Long)
        OP_REPETICION = dato
End Property
Public Property Let setSITUACION(ByVal dato As Long)
        SITUACION = dato
End Property
Public Property Let setFIRMA(ByVal dato As Long)
        FIRMA = dato
End Property
Public Property Let setPRODUCTO(ByVal dato As String)
        PRODUCTO = dato
End Property
Public Property Let setRESPONSABLE_ID(ByVal dato As Long)
        RESPONSABLE_ID = dato
End Property
Public Property Let setPEDIDO_ID(ByVal dato As Long)
        PEDIDO_ID = dato
End Property
Public Property Let setOFERTA_ID(ByVal dato As Long)
        OFERTA_ID = dato
End Property
Public Property Let setPAQUETE_ID(ByVal dato As Long)
        PAQUETE_ID = dato
End Property
Public Property Let setENAC(ByVal dato As Long)
        ENAC = dato
End Property
Public Property Let setNADCAP(ByVal dato As Long)
        NADCAP = dato
End Property
Public Property Let setURGENTE(ByVal dato As Long)
        URGENTE = dato
End Property
Public Property Let setAJUSTE(ByVal dato As Long)
        AJUSTE = dato
End Property
Public Property Let setAGRUPADA(ByVal dato As Long)
        AGRUPADA = dato
End Property
Public Property Let setAGRUPADA_MUESTRA_ID(ByVal dato As Long)
        AGRUPADA_MUESTRA_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSMUESTRAS
'***************************************************************
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
Public Property Get getANALISIS_MODIFICADO() As Long
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
Public Property Get getFORMATO_ID() As Long
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
Public Property Get getPRECIO() As Long
        getPRECIO = PRECIO
End Property
Public Property Get getFECHA_PREV_FIN() As String
        getFECHA_PREV_FIN = FECHA_PREV_FIN
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getANALISIS_DUPLICADO() As Long
        getANALISIS_DUPLICADO = ANALISIS_DUPLICADO
End Property
Public Property Get getANULADA() As Long
        getANULADA = ANULADA
End Property
Public Property Get getPRECINTO() As String
        getPRECINTO = PRECINTO
End Property
Public Property Get getBANO_ID() As Long
        getBANO_ID = BANO_ID
End Property
Public Property Get getIPA() As Long
        getIPA = IPA
End Property
Public Property Get getINFORME_MANUAL() As Long
        getINFORME_MANUAL = INFORME_MANUAL
End Property
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
        getFECHA_CIERRE = FECHA_CIERRE
End Property
Public Property Get getHORA_CIERRE() As String
        getHORA_CIERRE = HORA_CIERRE
End Property
Public Property Get getFECHA_ENVIO() As String
        getFECHA_ENVIO = FECHA_ENVIO
End Property
Public Property Get getMOTIVO_RETRASO_ID() As Long
        getMOTIVO_RETRASO_ID = MOTIVO_RETRASO_ID
End Property
Public Property Get getRETRASO_OBSERVACION() As String
        getRETRASO_OBSERVACION = RETRASO_OBSERVACION
End Property
Public Property Get getCERRADA() As Long
        getCERRADA = CERRADA
End Property
Public Property Get getREVISION_USUARIO() As Long
        getREVISION_USUARIO = REVISION_USUARIO
End Property
Public Property Get getCERRADA_USUARIO() As Long
        getCERRADA_USUARIO = CERRADA_USUARIO
End Property
Public Property Get getDOCUMENTO_PAGO() As Long
        getDOCUMENTO_PAGO = DOCUMENTO_PAGO
End Property
Public Property Get getIMPRESA() As Long
        getIMPRESA = IMPRESA
End Property
Public Property Get getENVIADO_CORREO() As Long
        getENVIADO_CORREO = ENVIADO_CORREO
End Property
Public Property Get getULT_EDICION_IMP() As Long
        getULT_EDICION_IMP = ULT_EDICION_IMP
End Property
Public Property Get getOP_VUELO() As Long
        getOP_VUELO = OP_VUELO
End Property
Public Property Get getOP_INSITU() As Long
        getOP_INSITU = OP_INSITU
End Property
Public Property Get getOP_LABMOVIL() As Long
        getOP_LABMOVIL = OP_LABMOVIL
End Property
Public Property Get getOP_NORUTINARIA() As Long
        getOP_NORUTINARIA = OP_NORUTINARIA
End Property
Public Property Get getOP_REPETICION() As Long
        getOP_REPETICION = OP_REPETICION
End Property
Public Property Get getSITUACION() As Long
        getSITUACION = SITUACION
End Property
Public Property Get getFIRMA() As Long
        getFIRMA = FIRMA
End Property
Public Property Get getPRODUCTO() As String
        getPRODUCTO = PRODUCTO
End Property
Public Property Get getRESPONSABLE_ID() As Long
        getRESPONSABLE_ID = RESPONSABLE_ID
End Property
Public Property Get getPEDIDO_ID() As Long
        getPEDIDO_ID = PEDIDO_ID
End Property
Public Property Get getOFERTA_ID() As Long
        getOFERTA_ID = OFERTA_ID
End Property
Public Property Get getPAQUETE_ID() As Long
        getPAQUETE_ID = PAQUETE_ID
End Property
Public Property Get getENAC() As Long
        getENAC = ENAC
End Property
Public Property Get getNADCAP() As Long
        getNADCAP = NADCAP
End Property
Public Property Get getURGENTE() As Long
        getURGENTE = URGENTE
End Property
Public Property Get getAJUSTE() As Long
        getAJUSTE = AJUSTE
End Property
Public Property Get getAGRUPADA() As Long
        getAGRUPADA = AGRUPADA
End Property
Public Property Get getAGRUPADA_MUESTRA_ID() As Long
        getAGRUPADA_MUESTRA_ID = AGRUPADA_MUESTRA_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSMUESTRAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM geslab_canagrosa.MUESTRAS WHERE ID_MUESTRA = " & ID
        Set rs = datos_bd_metrologia(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_MUESTRA = rs("ID_MUESTRA")
                TIPO_MUESTRA_ID = rs("TIPO_MUESTRA_ID")
                ID_GENERAL = rs("ID_GENERAL")
                ID_PARTICULAR = rs("ID_PARTICULAR")
                ANNO = rs("ANNO")
                TIPO_ANALISIS_ID = rs("TIPO_ANALISIS_ID")
                ANALISIS_MODIFICADO = rs("ANALISIS_MODIFICADO")
                FECHA_MUESTREO = rs("FECHA_MUESTREO")
                ENTIDAD_MUESTREO_ID = rs("ENTIDAD_MUESTREO_ID")
                DETALLE_MUESTREO = rs("DETALLE_MUESTREO")
                OBSERVACIONES_MUESTREO = rs("OBSERVACIONES_MUESTREO")
                FECHA_RECEPCION = rs("FECHA_RECEPCION")
                HORA_RECEPCION = rs("HORA_RECEPCION")
                EMPLEADO_ID = rs("EMPLEADO_ID")
                FORMATO_ID = rs("FORMATO_ID")
                ENTIDAD_ENTREGA_ID = rs("ENTIDAD_ENTREGA_ID")
                DETALLE_ENTREGA = rs("DETALLE_ENTREGA")
                OBSERVACIONES_ENTREGA = rs("OBSERVACIONES_ENTREGA")
                CLIENTE_ID = rs("CLIENTE_ID")
                CENTRO_ID = rs("CENTRO_ID")
                REPLACEMENT_ID = rs("REPLACEMENT_ID")
                REFERENCIA_CLIENTE = rs("REFERENCIA_CLIENTE")
                PRECIO = rs("PRECIO")
'                FECHA_PREV_FIN = rs("FECHA_PREV_FIN")
'                OBSERVACIONES = rs("OBSERVACIONES")
'                ANALISIS_DUPLICADO = rs("ANALISIS_DUPLICADO")
'                ANULADA = rs("ANULADA")
'                PRECINTO = rs("PRECINTO")
'                BANO_ID = rs("BANO_ID")
'                IPA = rs("IPA")
'                INFORME_MANUAL = rs("INFORME_MANUAL")
'                FECHA_RECOGIDA = rs("FECHA_RECOGIDA")
'                FECHA_COMIENZO = rs("FECHA_COMIENZO")
'                FECHA_FINALIZACION = rs("FECHA_FINALIZACION")
'                FECHA_CIERRE = rs("FECHA_CIERRE")
'                HORA_CIERRE = rs("HORA_CIERRE")
'                FECHA_ENVIO = rs("FECHA_ENVIO")
'                MOTIVO_RETRASO_ID = rs("MOTIVO_RETRASO_ID")
'                RETRASO_OBSERVACION = rs("RETRASO_OBSERVACION")
'                CERRADA = rs("CERRADA")
'                REVISION_USUARIO = rs("REVISION_USUARIO")
'                CERRADA_USUARIO = rs("CERRADA_USUARIO")
'                DOCUMENTO_PAGO = rs("DOCUMENTO_PAGO")
'                IMPRESA = rs("IMPRESA")
'                ENVIADO_CORREO = rs("ENVIADO_CORREO")
'                ULT_EDICION_IMP = rs("ULT_EDICION_IMP")
'                OP_VUELO = rs("OP_VUELO")
'                OP_INSITU = rs("OP_INSITU")
'                OP_LABMOVIL = rs("OP_LABMOVIL")
'                OP_NORUTINARIA = rs("OP_NORUTINARIA")
'                OP_REPETICION = rs("OP_REPETICION")
'                SITUACION = rs("SITUACION")
'                FIRMA = rs("FIRMA")
'                PRODUCTO = rs("PRODUCTO")
'                RESPONSABLE_ID = rs("RESPONSABLE_ID")
'                PEDIDO_ID = rs("PEDIDO_ID")
'                OFERTA_ID = rs("OFERTA_ID")
'                PAQUETE_ID = rs("PAQUETE_ID")
'                ENAC = rs("ENAC")
'                NADCAP = rs("NADCAP")
'                URGENTE = rs("URGENTE")
'                AJUSTE = rs("AJUSTE")
'                AGRUPADA = rs("AGRUPADA")
'                AGRUPADA_MUESTRA_ID = rs("AGRUPADA_MUESTRA_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsMuestras)", vbCritical, Err.Description
End Function
