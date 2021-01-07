VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEq_calibracion_equipos"
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
Private FECHA_ACTUAL As String
Private FECHA_PROXIMA As String
Private FECHA_PREVISTA As String
Private RANGO_MIN As Long
Private RANGO_MAX As Long
Private UNIDADES_ID As Long
Private RUTA_PLANTILLA As String
Private RUTA_CERTIFICADO As String
Private ESTADO As Integer
Private RESULTADO As Integer
Private CUSERID As Long
Private MUSERID As Long
Private TS As String
Private RUTA_EVALUACION As String
Private UBICACION_ID As Long
Private INCIDENCIAS As String
Private BUSQUEDAS As Long
Private ENVIADO As Long
Private EDICION As Long
Private INCERTIDUMBRE_USO As Long
Private INCERTIDUMBRE As String
Private SEGREGADO As Long
Private ETIQUETADO As Long
Private FECHA_CIERRE As String
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
        FECHA_ACTUAL = dato
End Property
Public Property Let setFECHA_PROXIMA(ByVal dato As String)
        FECHA_PROXIMA = dato
End Property
Public Property Let setFECHA_PREVISTA(ByVal dato As String)
        FECHA_PREVISTA = dato
End Property
Public Property Let setRANGO_MIN(ByVal dato As Long)
        RANGO_MIN = dato
End Property
Public Property Let setRANGO_MAX(ByVal dato As Long)
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
Public Property Let setRUTA_EVALUACION(ByVal dato As String)
        RUTA_EVALUACION = dato
End Property
Public Property Let setUBICACION_ID(ByVal dato As Long)
        UBICACION_ID = dato
End Property
Public Property Let setINCIDENCIAS(ByVal dato As String)
        INCIDENCIAS = dato
End Property
Public Property Let setBUSQUEDAS(ByVal dato As Long)
        BUSQUEDAS = dato
End Property
Public Property Let setENVIADO(ByVal dato As Long)
        ENVIADO = dato
End Property
Public Property Let setEDICION(ByVal dato As Long)
        EDICION = dato
End Property
Public Property Let setINCERTIDUMBRE_USO(ByVal dato As Long)
        INCERTIDUMBRE_USO = dato
End Property
Public Property Let setINCERTIDUMBRE(ByVal dato As String)
        INCERTIDUMBRE = dato
End Property
Public Property Let setSEGREGADO(ByVal dato As Long)
        SEGREGADO = dato
End Property
Public Property Let setETIQUETADO(ByVal dato As Long)
        ETIQUETADO = dato
End Property
Public Property Let setFECHA_CIERRE(ByVal dato As String)
        FECHA_CIERRE = dato
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
Public Property Get getCALIBRADOR_INTERNO_ID() As Long
        getCALIBRADOR_INTERNO_ID = CALIBRADOR_INTERNO_ID
End Property
Public Property Get getFECHA_ACTUAL() As String
        getFECHA_ACTUAL = FECHA_ACTUAL
End Property
Public Property Get getFECHA_PROXIMA() As String
        getFECHA_PROXIMA = FECHA_PROXIMA
End Property
Public Property Get getFECHA_PREVISTA() As String
        getFECHA_PREVISTA = FECHA_PREVISTA
End Property
Public Property Get getRANGO_MIN() As Long
        getRANGO_MIN = RANGO_MIN
End Property
Public Property Get getRANGO_MAX() As Long
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
Public Property Get getRUTA_EVALUACION() As String
        getRUTA_EVALUACION = RUTA_EVALUACION
End Property
Public Property Get getUBICACION_ID() As Long
        getUBICACION_ID = UBICACION_ID
End Property
Public Property Get getINCIDENCIAS() As String
        getINCIDENCIAS = INCIDENCIAS
End Property
Public Property Get getBUSQUEDAS() As Long
        getBUSQUEDAS = BUSQUEDAS
End Property
Public Property Get getENVIADO() As Long
        getENVIADO = ENVIADO
End Property
Public Property Get getEDICION() As Long
        getEDICION = EDICION
End Property
Public Property Get getINCERTIDUMBRE_USO() As Long
        getINCERTIDUMBRE_USO = INCERTIDUMBRE_USO
End Property
Public Property Get getINCERTIDUMBRE() As String
        getINCERTIDUMBRE = INCERTIDUMBRE
End Property
Public Property Get getSEGREGADO() As Long
        getSEGREGADO = SEGREGADO
End Property
Public Property Get getETIQUETADO() As Long
        getETIQUETADO = ETIQUETADO
End Property
Public Property Get getFECHA_CIERRE() As String
        getFECHA_CIERRE = FECHA_CIERRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQ_CALIBRACION_EQUIPOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM geslab_canagrosa.EQ_CALIBRACION_EQUIPOS WHERE ID_CALIBRACION = " & ID
        Set rs = datos_bd_metrologia(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_CALIBRACION = rs("ID_CALIBRACION")
                MUESTRA_ID = rs("MUESTRA_ID")
                EQUIPO_ID = rs("EQUIPO_ID")
                TIPO_ID = rs("TIPO_ID")
                PERIODICIDAD_ID = rs("PERIODICIDAD_ID")
                PROCEDIMIENTO_ID = rs("PROCEDIMIENTO_ID")
                CALIBRADOR_EXTERNO_ID = rs("CALIBRADOR_EXTERNO_ID")
                CALIBRADOR_INTERNO_ID = rs("CALIBRADOR_INTERNO_ID")
                FECHA_ACTUAL = rs("FECHA_ACTUAL")
                FECHA_PROXIMA = rs("FECHA_PROXIMA")
                FECHA_PREVISTA = rs("FECHA_PREVISTA")
                RANGO_MIN = rs("RANGO_MIN")
                RANGO_MAX = rs("RANGO_MAX")
                UNIDADES_ID = rs("UNIDADES_ID")
                RUTA_PLANTILLA = rs("RUTA_PLANTILLA")
                RUTA_CERTIFICADO = rs("RUTA_CERTIFICADO")
                ESTADO = rs("ESTADO")
                RESULTADO = rs("RESULTADO")
                CUSERID = rs("CUSERID")
                MUSERID = rs("MUSERID")
                TS = rs("TS")
                RUTA_EVALUACION = rs("RUTA_EVALUACION")
                UBICACION_ID = rs("UBICACION_ID")
                INCIDENCIAS = rs("INCIDENCIAS")
                BUSQUEDAS = rs("BUSQUEDAS")
                If Not IsNull(rs("ENVIADO")) Then
                    ENVIADO = rs("ENVIADO")
                End If
                If Not IsNull(rs("EDICION")) Then
                    EDICION = rs("EDICION")
                End If
                If Not IsNull(rs("INCERTIDUMBRE_USO")) Then
                    INCERTIDUMBRE_USO = rs("INCERTIDUMBRE_USO")
                End If
                If Not IsNull(rs("INCERTIDUMBRE")) Then
                    INCERTIDUMBRE = rs("INCERTIDUMBRE")
                End If
                If Not IsNull(rs("SEGREGADO")) Then
                    SEGREGADO = rs("SEGREGADO")
                End If
                If Not IsNull(rs("ETIQUETADO")) Then
                    ETIQUETADO = rs("ETIQUETADO")
                End If
                If Not IsNull(rs("FECHA_CIERRE")) Then
                    FECHA_CIERRE = rs("FECHA_CIERRE")
                End If
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEq_calibracion_equipos)", vbCritical, Err.Description
End Function
