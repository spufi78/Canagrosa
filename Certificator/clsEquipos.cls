VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS
'***************************************************************
Private ID_EQUIPO As Long
Private PARTICULAR As Long
Private NOMBRE As String
Private CONDICIONES_AMBIENTALES As Long
Private PROVEEDOR_ID As Long
Private FECHA_RECEPCION As String
Private FECHA_SERVICIO As String
Private ALTA_BAJA As Integer
Private ESTADO_ID As String
Private SERIE As String
Private FAMILIA_ID As Long
Private MODELO As String
Private FABRICANTE As String
Private ES_NADCAP As Long
Private ES_MTL As Long
Private ES_CP As Long
Private ES_ENAC As Long
Private RANGO_MEDIDA_MIN As String
Private RANGO_MEDIDA_MAX As String
Private UNIDAD_ID As Long
Private TOLERANCIA_MAXIMA As String
Private INCERTIDUMBRE_MAXIMA As String
Private PRECISIONN As String
Private TEMPERATURA_MIN As String
Private TEMPERATURA_MAX As String
Private HUMEDAD_MIN As String
Private HUMEDAD_MAX As String
Private COND_AMBIENTALES_OTRAS As String
Private RANGO_TRABAJO_MIN As String
Private RANGO_TRABAJO_MAX As String
Private LIMITACIONES_USO As String
Private CON_CALIBRACION As Long
Private PERIODICIDAD_CALIBRACION_ID As Long
Private FECHA_PROX_CALIBRACION As String
Private CON_VERIFICACION As Integer
Private PERIODICIDAD_VERIFICACION_ID As Long
Private FECHA_PROX_VERIFICACION As String
Private CON_MANTENIMIENTO As Integer
Private PERIODICIDAD_MANTENIMIENTO_ID As Long
Private FECHA_PROX_MANTENIMIENTO As String
Private SITUACION_ID As Long
Private TIPO_CALIBRACION_ID As Long
Private CALIBRADOR_ID As Long
Private CALIBRADOR_INTERNO_ID As Long
Private NOTAS As String
Private REVISADO As Long
Private CUSERID As String
Private FTIMESTP As String
Private MOSTRAR_EN_PLANO As Integer
Private PRIORITARIO As Integer
Private INSITU As Integer
Private RESPONSABLE_ID As Long
Private PROCEDENCIA_ID As Long
Private SITUACION_DESCRIPCION As String
Private COORDX As Long
Private COORDY As Long
Private VERIFICADOR_ID As Long
Private VERIFICADOR_INTERNO_ID As Long
Private TIPO_VERIFICACION_ID As Long
Private PLAN_MANTENIMIENTO_ID As Long
Private DESCRIPCION As String
Private OBSERVACIONES As String
Private MANTENEDOR_ID As Long
Private TIPO_EQUIPO_ID As Long
Private PROCEDIMIENTO_CALIBRACION_ID As Long
Private FUERA_SERVICIO As Integer
Private PROCEDIMIENTO_VERIFICACION_ID As Long
Private ES_ACCESORIO As Integer
Private NUMERO_USOS_MAXIMO As Long
Private NUMERO_USOS_CONTADOR As Long
Private UNIDADES_LOTE As Long
Private CLIENTE_ID As Long
Private TRAZMET_ID As Long
Private NUMERO_EQUIPO_CLIENTE As String
Private AV As Long
Private SPR As Long
Private RPR As Long
Private CRITICO As Long
Private CENTRO_ID As Long

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS
'***************************************************************
Public Property Let setID_EQUIPO(ByVal dato As Long)
        ID_EQUIPO = dato
End Property
Public Property Let setPARTICULAR(ByVal dato As Long)
        PARTICULAR = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
Public Property Let setCONDICIONES_AMBIENTALES(ByVal dato As Long)
        CONDICIONES_AMBIENTALES = dato
End Property
Public Property Let setPROVEEDOR_ID(ByVal dato As Long)
        PROVEEDOR_ID = dato
End Property
Public Property Let setFECHA_RECEPCION(ByVal dato As String)
        FECHA_RECEPCION = dato
End Property
Public Property Let setFECHA_SERVICIO(ByVal dato As String)
        FECHA_SERVICIO = dato
End Property
Public Property Let setALTA_BAJA(ByVal dato As Integer)
        ALTA_BAJA = dato
End Property
Public Property Let setESTADO_ID(ByVal dato As String)
        ESTADO_ID = dato
End Property
Public Property Let setSERIE(ByVal dato As String)
        SERIE = dato
End Property
Public Property Let setFAMILIA_ID(ByVal dato As Long)
        FAMILIA_ID = dato
End Property
Public Property Let setMODELO(ByVal dato As String)
        MODELO = dato
End Property
Public Property Let setFABRICANTE(ByVal dato As String)
        FABRICANTE = dato
End Property
Public Property Let setES_NADCAP(ByVal dato As Long)
        ES_NADCAP = dato
End Property
Public Property Let setES_MTL(ByVal dato As Long)
        ES_MTL = dato
End Property
Public Property Let setES_CP(ByVal dato As Long)
        ES_CP = dato
End Property
Public Property Let setES_ENAC(ByVal dato As Long)
        ES_ENAC = dato
End Property
Public Property Let setRANGO_MEDIDA_MIN(ByVal dato As String)
        RANGO_MEDIDA_MIN = dato
End Property
Public Property Let setRANGO_MEDIDA_MAX(ByVal dato As String)
        RANGO_MEDIDA_MAX = dato
End Property
Public Property Let setUNIDAD_ID(ByVal dato As Long)
        UNIDAD_ID = dato
End Property
Public Property Let setTOLERANCIA_MAXIMA(ByVal dato As String)
        TOLERANCIA_MAXIMA = dato
End Property
Public Property Let setINCERTIDUMBRE_MAXIMA(ByVal dato As String)
        INCERTIDUMBRE_MAXIMA = dato
End Property
Public Property Let setPRECISIONN(ByVal dato As String)
        PRECISIONN = dato
End Property
Public Property Let setTEMPERATURA_MIN(ByVal dato As String)
        TEMPERATURA_MIN = dato
End Property
Public Property Let setTEMPERATURA_MAX(ByVal dato As String)
        TEMPERATURA_MAX = dato
End Property
Public Property Let setHUMEDAD_MIN(ByVal dato As String)
        HUMEDAD_MIN = dato
End Property
Public Property Let setHUMEDAD_MAX(ByVal dato As String)
        HUMEDAD_MAX = dato
End Property
Public Property Let setCOND_AMBIENTALES_OTRAS(ByVal dato As String)
        COND_AMBIENTALES_OTRAS = dato
End Property
Public Property Let setRANGO_TRABAJO_MIN(ByVal dato As String)
        RANGO_TRABAJO_MIN = dato
End Property
Public Property Let setRANGO_TRABAJO_MAX(ByVal dato As String)
        RANGO_TRABAJO_MAX = dato
End Property
Public Property Let setLIMITACIONES_USO(ByVal dato As String)
        LIMITACIONES_USO = dato
End Property
Public Property Let setCON_CALIBRACION(ByVal dato As Long)
        CON_CALIBRACION = dato
End Property
Public Property Let setPERIODICIDAD_CALIBRACION_ID(ByVal dato As Long)
        PERIODICIDAD_CALIBRACION_ID = dato
End Property
Public Property Let setFECHA_PROX_CALIBRACION(ByVal dato As String)
        FECHA_PROX_CALIBRACION = dato
End Property
Public Property Let setCON_VERIFICACION(ByVal dato As Integer)
        CON_VERIFICACION = dato
End Property
Public Property Let setPERIODICIDAD_VERIFICACION_ID(ByVal dato As Long)
        PERIODICIDAD_VERIFICACION_ID = dato
End Property
Public Property Let setFECHA_PROX_VERIFICACION(ByVal dato As String)
        FECHA_PROX_VERIFICACION = dato
End Property
Public Property Let setCON_MANTENIMIENTO(ByVal dato As Integer)
        CON_MANTENIMIENTO = dato
End Property
Public Property Let setPERIODICIDAD_MANTENIMIENTO_ID(ByVal dato As Long)
        PERIODICIDAD_MANTENIMIENTO_ID = dato
End Property
Public Property Let setFECHA_PROX_MANTENIMIENTO(ByVal dato As String)
        FECHA_PROX_MANTENIMIENTO = dato
End Property
Public Property Let setSITUACION_ID(ByVal dato As Long)
        SITUACION_ID = dato
End Property
Public Property Let setTIPO_CALIBRACION_ID(ByVal dato As Long)
        TIPO_CALIBRACION_ID = dato
End Property
Public Property Let setCALIBRADOR_ID(ByVal dato As Long)
        CALIBRADOR_ID = dato
End Property
Public Property Let setCALIBRADOR_INTERNO_ID(ByVal dato As Long)
        CALIBRADOR_INTERNO_ID = dato
End Property
Public Property Let setNOTAS(ByVal dato As String)
        NOTAS = dato
End Property
Public Property Let setREVISADO(ByVal dato As Long)
        REVISADO = dato
End Property
Public Property Let setCUSERID(ByVal dato As String)
        CUSERID = dato
End Property
Public Property Let setFTIMESTP(ByVal dato As String)
        FTIMESTP = dato
End Property
Public Property Let setMOSTRAR_EN_PLANO(ByVal dato As Integer)
        MOSTRAR_EN_PLANO = dato
End Property
Public Property Let setPRIORITARIO(ByVal dato As Integer)
        PRIORITARIO = dato
End Property
Public Property Let setINSITU(ByVal dato As Integer)
        INSITU = dato
End Property
Public Property Let setRESPONSABLE_ID(ByVal dato As Long)
        RESPONSABLE_ID = dato
End Property
Public Property Let setPROCEDENCIA_ID(ByVal dato As Long)
        PROCEDENCIA_ID = dato
End Property
Public Property Let setSITUACION_DESCRIPCION(ByVal dato As String)
        SITUACION_DESCRIPCION = dato
End Property
Public Property Let setCOORDX(ByVal dato As Long)
        COORDX = dato
End Property
Public Property Let setCOORDY(ByVal dato As Long)
        COORDY = dato
End Property
Public Property Let setVERIFICADOR_ID(ByVal dato As Long)
        VERIFICADOR_ID = dato
End Property
Public Property Let setVERIFICADOR_INTERNO_ID(ByVal dato As Long)
        VERIFICADOR_INTERNO_ID = dato
End Property
Public Property Let setTIPO_VERIFICACION_ID(ByVal dato As Long)
        TIPO_VERIFICACION_ID = dato
End Property
Public Property Let setPLAN_MANTENIMIENTO_ID(ByVal dato As Long)
        PLAN_MANTENIMIENTO_ID = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setMANTENEDOR_ID(ByVal dato As Long)
        MANTENEDOR_ID = dato
End Property
Public Property Let setTIPO_EQUIPO_ID(ByVal dato As Long)
        TIPO_EQUIPO_ID = dato
End Property
Public Property Let setPROCEDIMIENTO_CALIBRACION_ID(ByVal dato As Long)
        PROCEDIMIENTO_CALIBRACION_ID = dato
End Property
Public Property Let setFUERA_SERVICIO(ByVal dato As Integer)
        FUERA_SERVICIO = dato
End Property
Public Property Let setPROCEDIMIENTO_VERIFICACION_ID(ByVal dato As Long)
        PROCEDIMIENTO_VERIFICACION_ID = dato
End Property
Public Property Let setES_ACCESORIO(ByVal dato As Integer)
        ES_ACCESORIO = dato
End Property
Public Property Let setNUMERO_USOS_MAXIMO(ByVal dato As Long)
        NUMERO_USOS_MAXIMO = dato
End Property
Public Property Let setNUMERO_USOS_CONTADOR(ByVal dato As Long)
        NUMERO_USOS_CONTADOR = dato
End Property
Public Property Let setUNIDADES_LOTE(ByVal dato As Long)
        UNIDADES_LOTE = dato
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setTRAZMET_ID(ByVal dato As Long)
        TRAZMET_ID = dato
End Property
Public Property Let setNUMERO_EQUIPO_CLIENTE(ByVal dato As String)
        NUMERO_EQUIPO_CLIENTE = dato
End Property
Public Property Let setAV(ByVal dato As Long)
        AV = dato
End Property
Public Property Let setSPR(ByVal dato As Long)
        SPR = dato
End Property
Public Property Let setRPR(ByVal dato As Long)
        RPR = dato
End Property
Public Property Let setCRITICO(ByVal dato As Long)
        CRITICO = dato
End Property
Public Property Let setCENTRO_ID(ByVal dato As Long)
        CENTRO_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS
'***************************************************************
Public Property Get getID_EQUIPO() As Long
        getID_EQUIPO = ID_EQUIPO
End Property
Public Property Get getPARTICULAR() As Long
        getPARTICULAR = PARTICULAR
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
Public Property Get getCONDICIONES_AMBIENTALES() As Long
        getCONDICIONES_AMBIENTALES = CONDICIONES_AMBIENTALES
End Property
Public Property Get getPROVEEDOR_ID() As Long
        getPROVEEDOR_ID = PROVEEDOR_ID
End Property
Public Property Get getFECHA_RECEPCION() As String
        getFECHA_RECEPCION = FECHA_RECEPCION
End Property
Public Property Get getFECHA_SERVICIO() As String
        getFECHA_SERVICIO = FECHA_SERVICIO
End Property
Public Property Get getALTA_BAJA() As Integer
        getALTA_BAJA = ALTA_BAJA
End Property
Public Property Get getESTADO_ID() As String
        getESTADO_ID = ESTADO_ID
End Property
Public Property Get getSERIE() As String
        getSERIE = SERIE
End Property
Public Property Get getFAMILIA_ID() As Long
        getFAMILIA_ID = FAMILIA_ID
End Property
Public Property Get getMODELO() As String
        getMODELO = MODELO
End Property
Public Property Get getFABRICANTE() As String
        getFABRICANTE = FABRICANTE
End Property
Public Property Get getES_NADCAP() As Long
        getES_NADCAP = ES_NADCAP
End Property
Public Property Get getES_MTL() As Long
        getES_MTL = ES_MTL
End Property
Public Property Get getES_CP() As Long
        getES_CP = ES_CP
End Property
Public Property Get getES_ENAC() As Long
        getES_ENAC = ES_ENAC
End Property
Public Property Get getRANGO_MEDIDA_MIN() As String
        getRANGO_MEDIDA_MIN = RANGO_MEDIDA_MIN
End Property
Public Property Get getRANGO_MEDIDA_MAX() As String
        getRANGO_MEDIDA_MAX = RANGO_MEDIDA_MAX
End Property
Public Property Get getUNIDAD_ID() As Long
        getUNIDAD_ID = UNIDAD_ID
End Property
Public Property Get getTOLERANCIA_MAXIMA() As String
        getTOLERANCIA_MAXIMA = TOLERANCIA_MAXIMA
End Property
Public Property Get getINCERTIDUMBRE_MAXIMA() As String
        getINCERTIDUMBRE_MAXIMA = INCERTIDUMBRE_MAXIMA
End Property
Public Property Get getPRECISIONN() As String
        getPRECISIONN = PRECISIONN
End Property
Public Property Get getTEMPERATURA_MIN() As String
        getTEMPERATURA_MIN = TEMPERATURA_MIN
End Property
Public Property Get getTEMPERATURA_MAX() As String
        getTEMPERATURA_MAX = TEMPERATURA_MAX
End Property
Public Property Get getHUMEDAD_MIN() As String
        getHUMEDAD_MIN = HUMEDAD_MIN
End Property
Public Property Get getHUMEDAD_MAX() As String
        getHUMEDAD_MAX = HUMEDAD_MAX
End Property
Public Property Get getCOND_AMBIENTALES_OTRAS() As String
        getCOND_AMBIENTALES_OTRAS = COND_AMBIENTALES_OTRAS
End Property
Public Property Get getRANGO_TRABAJO_MIN() As String
        getRANGO_TRABAJO_MIN = RANGO_TRABAJO_MIN
End Property
Public Property Get getRANGO_TRABAJO_MAX() As String
        getRANGO_TRABAJO_MAX = RANGO_TRABAJO_MAX
End Property
Public Property Get getLIMITACIONES_USO() As String
        getLIMITACIONES_USO = LIMITACIONES_USO
End Property
Public Property Get getCON_CALIBRACION() As Long
        getCON_CALIBRACION = CON_CALIBRACION
End Property
Public Property Get getPERIODICIDAD_CALIBRACION_ID() As Long
        getPERIODICIDAD_CALIBRACION_ID = PERIODICIDAD_CALIBRACION_ID
End Property
Public Property Get getFECHA_PROX_CALIBRACION() As String
        getFECHA_PROX_CALIBRACION = FECHA_PROX_CALIBRACION
End Property
Public Property Get getCON_VERIFICACION() As Integer
        getCON_VERIFICACION = CON_VERIFICACION
End Property
Public Property Get getPERIODICIDAD_VERIFICACION_ID() As Long
        getPERIODICIDAD_VERIFICACION_ID = PERIODICIDAD_VERIFICACION_ID
End Property
Public Property Get getFECHA_PROX_VERIFICACION() As String
        getFECHA_PROX_VERIFICACION = FECHA_PROX_VERIFICACION
End Property
Public Property Get getCON_MANTENIMIENTO() As Integer
        getCON_MANTENIMIENTO = CON_MANTENIMIENTO
End Property
Public Property Get getPERIODICIDAD_MANTENIMIENTO_ID() As Long
        getPERIODICIDAD_MANTENIMIENTO_ID = PERIODICIDAD_MANTENIMIENTO_ID
End Property
Public Property Get getFECHA_PROX_MANTENIMIENTO() As String
        getFECHA_PROX_MANTENIMIENTO = FECHA_PROX_MANTENIMIENTO
End Property
Public Property Get getSITUACION_ID() As Long
        getSITUACION_ID = SITUACION_ID
End Property
Public Property Get getTIPO_CALIBRACION_ID() As Long
        getTIPO_CALIBRACION_ID = TIPO_CALIBRACION_ID
End Property
Public Property Get getCALIBRADOR_ID() As Long
        getCALIBRADOR_ID = CALIBRADOR_ID
End Property
Public Property Get getCALIBRADOR_INTERNO_ID() As Long
        getCALIBRADOR_INTERNO_ID = CALIBRADOR_INTERNO_ID
End Property
Public Property Get getNOTAS() As String
        getNOTAS = NOTAS
End Property
Public Property Get getREVISADO() As Long
        getREVISADO = REVISADO
End Property
Public Property Get getCUSERID() As String
        getCUSERID = CUSERID
End Property
Public Property Get getFTIMESTP() As String
        getFTIMESTP = FTIMESTP
End Property
Public Property Get getMOSTRAR_EN_PLANO() As Integer
        getMOSTRAR_EN_PLANO = MOSTRAR_EN_PLANO
End Property
Public Property Get getPRIORITARIO() As Integer
        getPRIORITARIO = PRIORITARIO
End Property
Public Property Get getINSITU() As Integer
        getINSITU = INSITU
End Property
Public Property Get getRESPONSABLE_ID() As Long
        getRESPONSABLE_ID = RESPONSABLE_ID
End Property
Public Property Get getPROCEDENCIA_ID() As Long
        getPROCEDENCIA_ID = PROCEDENCIA_ID
End Property
Public Property Get getSITUACION_DESCRIPCION() As String
        getSITUACION_DESCRIPCION = SITUACION_DESCRIPCION
End Property
Public Property Get getCOORDX() As Long
        getCOORDX = COORDX
End Property
Public Property Get getCOORDY() As Long
        getCOORDY = COORDY
End Property
Public Property Get getVERIFICADOR_ID() As Long
        getVERIFICADOR_ID = VERIFICADOR_ID
End Property
Public Property Get getVERIFICADOR_INTERNO_ID() As Long
        getVERIFICADOR_INTERNO_ID = VERIFICADOR_INTERNO_ID
End Property
Public Property Get getTIPO_VERIFICACION_ID() As Long
        getTIPO_VERIFICACION_ID = TIPO_VERIFICACION_ID
End Property
Public Property Get getPLAN_MANTENIMIENTO_ID() As Long
        getPLAN_MANTENIMIENTO_ID = PLAN_MANTENIMIENTO_ID
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getMANTENEDOR_ID() As Long
        getMANTENEDOR_ID = MANTENEDOR_ID
End Property
Public Property Get getTIPO_EQUIPO_ID() As Long
        getTIPO_EQUIPO_ID = TIPO_EQUIPO_ID
End Property
Public Property Get getPROCEDIMIENTO_CALIBRACION_ID() As Long
        getPROCEDIMIENTO_CALIBRACION_ID = PROCEDIMIENTO_CALIBRACION_ID
End Property
Public Property Get getFUERA_SERVICIO() As Integer
        getFUERA_SERVICIO = FUERA_SERVICIO
End Property
Public Property Get getPROCEDIMIENTO_VERIFICACION_ID() As Long
        getPROCEDIMIENTO_VERIFICACION_ID = PROCEDIMIENTO_VERIFICACION_ID
End Property
Public Property Get getES_ACCESORIO() As Integer
        getES_ACCESORIO = ES_ACCESORIO
End Property
Public Property Get getNUMERO_USOS_MAXIMO() As Long
        getNUMERO_USOS_MAXIMO = NUMERO_USOS_MAXIMO
End Property
Public Property Get getNUMERO_USOS_CONTADOR() As Long
        getNUMERO_USOS_CONTADOR = NUMERO_USOS_CONTADOR
End Property
Public Property Get getUNIDADES_LOTE() As Long
        getUNIDADES_LOTE = UNIDADES_LOTE
End Property
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getTRAZMET_ID() As Long
        getTRAZMET_ID = TRAZMET_ID
End Property
Public Property Get getNUMERO_EQUIPO_CLIENTE() As String
        getNUMERO_EQUIPO_CLIENTE = NUMERO_EQUIPO_CLIENTE
End Property
Public Property Get getAV() As Long
        getAV = AV
End Property
Public Property Get getSPR() As Long
        getSPR = SPR
End Property
Public Property Get getRPR() As Long
        getRPR = RPR
End Property
Public Property Get getCRITICO() As Long
        getCRITICO = CRITICO
End Property
Public Property Get getCENTRO_ID() As Long
        getCENTRO_ID = CENTRO_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM geslab_canagrosa.EQUIPOS WHERE ID_EQUIPO = " & ID
        Set rs = datos_bd_metrologia(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_EQUIPO = rs("ID_EQUIPO")
                PARTICULAR = rs("PARTICULAR")
                NOMBRE = rs("NOMBRE")
                CONDICIONES_AMBIENTALES = rs("CONDICIONES_AMBIENTALES")
                PROVEEDOR_ID = rs("PROVEEDOR_ID")
                FECHA_RECEPCION = rs("FECHA_RECEPCION")
                FECHA_SERVICIO = rs("FECHA_SERVICIO")
                ALTA_BAJA = rs("ALTA_BAJA")
                ESTADO_ID = rs("ESTADO_ID")
                SERIE = rs("SERIE")
                FAMILIA_ID = rs("FAMILIA_ID")
                MODELO = rs("MODELO")
                FABRICANTE = rs("FABRICANTE")
                ES_NADCAP = rs("ES_NADCAP")
                ES_MTL = rs("ES_MTL")
                ES_CP = rs("ES_CP")
                ES_ENAC = rs("ES_ENAC")
                RANGO_MEDIDA_MIN = rs("RANGO_MEDIDA_MIN")
                RANGO_MEDIDA_MAX = rs("RANGO_MEDIDA_MAX")
                UNIDAD_ID = rs("UNIDAD_ID")
                TOLERANCIA_MAXIMA = rs("TOLERANCIA_MAXIMA")
                INCERTIDUMBRE_MAXIMA = rs("INCERTIDUMBRE_MAXIMA")
                PRECISIONN = rs("PRECISIONN")
                TEMPERATURA_MIN = rs("TEMPERATURA_MIN")
                TEMPERATURA_MAX = rs("TEMPERATURA_MAX")
                HUMEDAD_MIN = rs("HUMEDAD_MIN")
                HUMEDAD_MAX = rs("HUMEDAD_MAX")
                COND_AMBIENTALES_OTRAS = rs("COND_AMBIENTALES_OTRAS")
                RANGO_TRABAJO_MIN = rs("RANGO_TRABAJO_MIN")
                RANGO_TRABAJO_MAX = rs("RANGO_TRABAJO_MAX")
                LIMITACIONES_USO = rs("LIMITACIONES_USO")
                CON_CALIBRACION = rs("CON_CALIBRACION")
                PERIODICIDAD_CALIBRACION_ID = rs("PERIODICIDAD_CALIBRACION_ID")
                FECHA_PROX_CALIBRACION = rs("FECHA_PROX_CALIBRACION")
                CON_VERIFICACION = rs("CON_VERIFICACION")
                PERIODICIDAD_VERIFICACION_ID = rs("PERIODICIDAD_VERIFICACION_ID")
                FECHA_PROX_VERIFICACION = rs("FECHA_PROX_VERIFICACION")
                CON_MANTENIMIENTO = rs("CON_MANTENIMIENTO")
                PERIODICIDAD_MANTENIMIENTO_ID = rs("PERIODICIDAD_MANTENIMIENTO_ID")
                FECHA_PROX_MANTENIMIENTO = rs("FECHA_PROX_MANTENIMIENTO")
                SITUACION_ID = rs("SITUACION_ID")
                TIPO_CALIBRACION_ID = rs("TIPO_CALIBRACION_ID")
                CALIBRADOR_ID = rs("CALIBRADOR_ID")
                CALIBRADOR_INTERNO_ID = rs("CALIBRADOR_INTERNO_ID")
                NOTAS = rs("NOTAS")
                REVISADO = rs("REVISADO")
                CUSERID = rs("CUSERID")
                FTIMESTP = rs("FTIMESTP")
                MOSTRAR_EN_PLANO = rs("MOSTRAR_EN_PLANO")
                PRIORITARIO = rs("PRIORITARIO")
                INSITU = rs("INSITU")
                RESPONSABLE_ID = rs("RESPONSABLE_ID")
                PROCEDENCIA_ID = rs("PROCEDENCIA_ID")
                SITUACION_DESCRIPCION = rs("SITUACION_DESCRIPCION")
                COORDX = rs("COORDX")
                COORDY = rs("COORDY")
                VERIFICADOR_ID = rs("VERIFICADOR_ID")
                VERIFICADOR_INTERNO_ID = rs("VERIFICADOR_INTERNO_ID")
                TIPO_VERIFICACION_ID = rs("TIPO_VERIFICACION_ID")
                PLAN_MANTENIMIENTO_ID = rs("PLAN_MANTENIMIENTO_ID")
                DESCRIPCION = rs("DESCRIPCION")
                OBSERVACIONES = rs("OBSERVACIONES")
                MANTENEDOR_ID = rs("MANTENEDOR_ID")
                TIPO_EQUIPO_ID = rs("TIPO_EQUIPO_ID")
                PROCEDIMIENTO_CALIBRACION_ID = rs("PROCEDIMIENTO_CALIBRACION_ID")
                FUERA_SERVICIO = rs("FUERA_SERVICIO")
                PROCEDIMIENTO_VERIFICACION_ID = rs("PROCEDIMIENTO_VERIFICACION_ID")
                ES_ACCESORIO = rs("ES_ACCESORIO")
                NUMERO_USOS_MAXIMO = rs("NUMERO_USOS_MAXIMO")
                NUMERO_USOS_CONTADOR = rs("NUMERO_USOS_CONTADOR")
                UNIDADES_LOTE = rs("UNIDADES_LOTE")
                CLIENTE_ID = rs("CLIENTE_ID")
                TRAZMET_ID = rs("TRAZMET_ID")
                NUMERO_EQUIPO_CLIENTE = rs("NUMERO_EQUIPO_CLIENTE")
                AV = rs("AV")
                SPR = rs("SPR")
                RPR = rs("RPR")
                CRITICO = rs("CRITICO")
                CENTRO_ID = rs("CENTRO_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquipos)", vbCritical, Err.Description
End Function
