VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsManometro_vacuometro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSMANOMETRO_VACUOMETRO
'***************************************************************
Private EQUIPO_ID As String
Private FLUIDO_DE_CAL As String
Private FLUIDO_CAL_INGLES As String
Private CRITERIO_ACEPTACION As String
Private ACCEPTANCE_CRITERIA As String
Private PUNTO1 As String
Private PUNTO2 As String
Private PUNTO3 As String
Private PUNTO4 As String
Private PUNTO5 As String
Private PUNTO6 As String
Private PUNTO7 As String
Private PUNTO8 As String
Private PUNTO9 As String
Private PUNTO10 As String
Private PNT_F_ASOCIADO As String
Private CMC_PUNTO1 As String
Private CMC_PUNTO2 As String
Private CMC_PUNTO3 As String
Private CMC_PUNTO4 As String
Private CMC_PUNTO5 As String
Private CMC_PUNTO6 As String
Private CMC_PUNTO7 As String
Private CMC_PUNTO8 As String
Private CMC_PUNTO9 As String
Private CMC_PUNTO10 As String
Private RANGO_UTIL_CAL As String
Private TOLERANCIA_PUNTO1 As String
Private TOLERANCIA_PUNTO2 As String
Private TOLERANCIA_PUNTO3 As String
Private TOLERANCIA_PUNTO4 As String
Private TOLERANCIA_PUNTO5 As String
Private TOLERANCIA_PUNTO6 As String
Private TOLERANCIA_PUNTO7 As String
Private TOLERANCIA_PUNTO8 As String
Private TOLERANCIA_PUNTO9 As String
Private TOLERANCIA_PUNTO10 As String
Private UNIDAD_TOLERANCIA As String
Private EVALUACION_TOLERANCIA As String
Private CRITERIO_DE_AJUSTE_PUNTO1 As String
Private CRITERIO_DE_AJUSTE_PUNTO2 As String
Private CRITERIO_DE_AJUSTE_PUNTO3 As String
Private CRITERIO_DE_AJUSTE_PUNTO4 As String
Private CRITERIO_DE_AJUSTE_PUNTO5 As String
Private CRITERIO_DE_AJUSTE_PUNTO6 As String
Private CRITERIO_DE_AJUSTE_PUNTO7 As String
Private CRITERIO_DE_AJUSTE_PUNTO8 As String
Private CRITERIO_DE_AJUSTE_PUNTO9 As String
Private CRITERIO_DE_AJUSTE_PUNTO10 As String
Private UNIDAD_AJUSTE As String
Private EVALUACION_AJUSTE As String
Private TRAZMET As String
Private AUXILIAR2 As String
Private AUXILIAR3 As String
Private INTERVALO_CALIBRACION_MESES As String
Private I_USO As String
Private NOMBRE_BANCO As String
Private IDENTIFICACION_BANCO As String
Private CMC_TOL_PUNTO_1 As String
Private CMC_TOL_PUNTO_2 As String
Private CMC_TOL_PUNTO_3 As String
Private CMC_TOL_PUNTO_4 As String
Private CMC_TOL_PUNTO_5 As String
Private CMC_TOL_PUNTO_6 As String
Private CMC_TOL_PUNTO_7 As String
Private CMC_TOL_PUNTO_8 As String
Private CMC_TOL_PUNTO_9 As String
Private CMC_TOL_PUNTO_10 As String
Private TRANSD_PUNTO1 As String
Private TRANSD_PUNTO2 As String
Private TRANSD_PUNTO3 As String
Private TRANSD_PUNTO4 As String
Private TRANSD_PUNTO5 As String
Private TRANSD_PUNTO6 As String
Private TRANSD_PUNTO7 As String
Private TRANSD_PUNTO8 As String
Private TRANSD_PUNTO9 As String
Private TRANSD_PUNTO10 As String
Private CLASE_DE_INFORME As String
Private TRANS1_CERT As String
Private TRANS2_CERT As String
Private TRANS3_CERT As String

'***************************************************************
'*   VARIABLES POR ENUMERACION DE LA CLASE CLSMANOMETRO_VACUOMETRO
'***************************************************************
Public Enum LISTA_CLSMANOMETRO_VACUOMETRO_COLS
        ENUM_EQUIPO_ID = 1
        ENUM_FLUIDO_DE_CAL = 2
        ENUM_FLUIDO_CAL_INGLES = 3
        ENUM_CRITERIO_ACEPTACION = 4
        ENUM_ACCEPTANCE_CRITERIA = 5
        ENUM_PUNTO1 = 6
        ENUM_PUNTO2 = 7
        ENUM_PUNTO3 = 8
        ENUM_PUNTO4 = 9
        ENUM_PUNTO5 = 10
        ENUM_PUNTO6 = 11
        ENUM_PUNTO7 = 12
        ENUM_PUNTO8 = 13
        ENUM_PUNTO9 = 14
        ENUM_PUNTO10 = 15
        ENUM_PNT_F_ASOCIADO = 16
        ENUM_CMC_PUNTO1 = 17
        ENUM_CMC_PUNTO2 = 18
        ENUM_CMC_PUNTO3 = 19
        ENUM_CMC_PUNTO4 = 20
        ENUM_CMC_PUNTO5 = 21
        ENUM_CMC_PUNTO6 = 22
        ENUM_CMC_PUNTO7 = 23
        ENUM_CMC_PUNTO8 = 24
        ENUM_CMC_PUNTO9 = 25
        ENUM_CMC_PUNTO10 = 26
        ENUM_RANGO_UTIL_CAL = 27
        ENUM_TOLERANCIA_PUNTO1 = 28
        ENUM_TOLERANCIA_PUNTO2 = 29
        ENUM_TOLERANCIA_PUNTO3 = 30
        ENUM_TOLERANCIA_PUNTO4 = 31
        ENUM_TOLERANCIA_PUNTO5 = 32
        ENUM_TOLERANCIA_PUNTO6 = 33
        ENUM_TOLERANCIA_PUNTO7 = 34
        ENUM_TOLERANCIA_PUNTO8 = 35
        ENUM_TOLERANCIA_PUNTO9 = 36
        ENUM_TOLERANCIA_PUNTO10 = 37
        ENUM_UNIDAD_TOLERANCIA = 38
        ENUM_EVALUACION_TOLERANCIA = 39
        ENUM_CRITERIO_DE_AJUSTE_PUNTO1 = 40
        ENUM_CRITERIO_DE_AJUSTE_PUNTO2 = 41
        ENUM_CRITERIO_DE_AJUSTE_PUNTO3 = 42
        ENUM_CRITERIO_DE_AJUSTE_PUNTO4 = 43
        ENUM_CRITERIO_DE_AJUSTE_PUNTO5 = 44
        ENUM_CRITERIO_DE_AJUSTE_PUNTO6 = 45
        ENUM_CRITERIO_DE_AJUSTE_PUNTO7 = 46
        ENUM_CRITERIO_DE_AJUSTE_PUNTO8 = 47
        ENUM_CRITERIO_DE_AJUSTE_PUNTO9 = 48
        ENUM_CRITERIO_DE_AJUSTE_PUNTO10 = 49
        ENUM_UNIDAD_AJUSTE = 50
        ENUM_EVALUACION_AJUSTE = 51
        ENUM_TRAZMET = 52
        ENUM_AUXILIAR2 = 53
        ENUM_AUXILIAR3 = 54
        ENUM_INTERVALO_CALIBRACION_MESES = 55
        ENUM_I_USO = 56
        ENUM_NOMBRE_BANCO = 57
        ENUM_IDENTIFICACION_BANCO = 58
        ENUM_CMC_TOL_PUNTO_1 = 59
        ENUM_CMC_TOL_PUNTO_2 = 60
        ENUM_CMC_TOL_PUNTO_3 = 61
        ENUM_CMC_TOL_PUNTO_4 = 62
        ENUM_CMC_TOL_PUNTO_5 = 63
        ENUM_CMC_TOL_PUNTO_6 = 64
        ENUM_CMC_TOL_PUNTO_7 = 65
        ENUM_CMC_TOL_PUNTO_8 = 66
        ENUM_CMC_TOL_PUNTO_9 = 67
        ENUM_CMC_TOL_PUNTO_10 = 68
        ENUM_TRANSD_PUNTO1 = 69
        ENUM_TRANSD_PUNTO2 = 70
        ENUM_TRANSD_PUNTO3 = 71
        ENUM_TRANSD_PUNTO4 = 72
        ENUM_TRANSD_PUNTO5 = 73
        ENUM_TRANSD_PUNTO6 = 74
        ENUM_TRANSD_PUNTO7 = 75
        ENUM_TRANSD_PUNTO8 = 76
        ENUM_TRANSD_PUNTO9 = 77
        ENUM_TRANSD_PUNTO10 = 78
        ENUM_CLASE_DE_INFORME = 79
        ENUM_TRANS1_CERT = 80
        ENUM_TRANS2_CERT = 81
        ENUM_TRANS3_CERT = 82
End Enum

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSMANOMETRO_VACUOMETRO
'***************************************************************
Public Property Let setEQUIPO_ID(ByVal dato As String)
        EQUIPO_ID = dato
End Property
Public Property Let setFLUIDO_DE_CAL(ByVal dato As String)
        FLUIDO_DE_CAL = dato
End Property
Public Property Let setFLUIDO_CAL_INGLES(ByVal dato As String)
        FLUIDO_CAL_INGLES = dato
End Property
Public Property Let setCRITERIO_ACEPTACION(ByVal dato As String)
        CRITERIO_ACEPTACION = dato
End Property
Public Property Let setACCEPTANCE_CRITERIA(ByVal dato As String)
        ACCEPTANCE_CRITERIA = dato
End Property
Public Property Let setPUNTO1(ByVal dato As String)
        PUNTO1 = dato
End Property
Public Property Let setPUNTO2(ByVal dato As String)
        PUNTO2 = dato
End Property
Public Property Let setPUNTO3(ByVal dato As String)
        PUNTO3 = dato
End Property
Public Property Let setPUNTO4(ByVal dato As String)
        PUNTO4 = dato
End Property
Public Property Let setPUNTO5(ByVal dato As String)
        PUNTO5 = dato
End Property
Public Property Let setPUNTO6(ByVal dato As String)
        PUNTO6 = dato
End Property
Public Property Let setPUNTO7(ByVal dato As String)
        PUNTO7 = dato
End Property
Public Property Let setPUNTO8(ByVal dato As String)
        PUNTO8 = dato
End Property
Public Property Let setPUNTO9(ByVal dato As String)
        PUNTO9 = dato
End Property
Public Property Let setPUNTO10(ByVal dato As String)
        PUNTO10 = dato
End Property
Public Property Let setPNT_F_ASOCIADO(ByVal dato As String)
        PNT_F_ASOCIADO = dato
End Property
Public Property Let setCMC_PUNTO1(ByVal dato As String)
        CMC_PUNTO1 = dato
End Property
Public Property Let setCMC_PUNTO2(ByVal dato As String)
        CMC_PUNTO2 = dato
End Property
Public Property Let setCMC_PUNTO3(ByVal dato As String)
        CMC_PUNTO3 = dato
End Property
Public Property Let setCMC_PUNTO4(ByVal dato As String)
        CMC_PUNTO4 = dato
End Property
Public Property Let setCMC_PUNTO5(ByVal dato As String)
        CMC_PUNTO5 = dato
End Property
Public Property Let setCMC_PUNTO6(ByVal dato As String)
        CMC_PUNTO6 = dato
End Property
Public Property Let setCMC_PUNTO7(ByVal dato As String)
        CMC_PUNTO7 = dato
End Property
Public Property Let setCMC_PUNTO8(ByVal dato As String)
        CMC_PUNTO8 = dato
End Property
Public Property Let setCMC_PUNTO9(ByVal dato As String)
        CMC_PUNTO9 = dato
End Property
Public Property Let setCMC_PUNTO10(ByVal dato As String)
        CMC_PUNTO10 = dato
End Property
Public Property Let setRANGO_UTIL_CAL(ByVal dato As String)
        RANGO_UTIL_CAL = dato
End Property
Public Property Let setTOLERANCIA_PUNTO1(ByVal dato As String)
        TOLERANCIA_PUNTO1 = dato
End Property
Public Property Let setTOLERANCIA_PUNTO2(ByVal dato As String)
        TOLERANCIA_PUNTO2 = dato
End Property
Public Property Let setTOLERANCIA_PUNTO3(ByVal dato As String)
        TOLERANCIA_PUNTO3 = dato
End Property
Public Property Let setTOLERANCIA_PUNTO4(ByVal dato As String)
        TOLERANCIA_PUNTO4 = dato
End Property
Public Property Let setTOLERANCIA_PUNTO5(ByVal dato As String)
        TOLERANCIA_PUNTO5 = dato
End Property
Public Property Let setTOLERANCIA_PUNTO6(ByVal dato As String)
        TOLERANCIA_PUNTO6 = dato
End Property
Public Property Let setTOLERANCIA_PUNTO7(ByVal dato As String)
        TOLERANCIA_PUNTO7 = dato
End Property
Public Property Let setTOLERANCIA_PUNTO8(ByVal dato As String)
        TOLERANCIA_PUNTO8 = dato
End Property
Public Property Let setTOLERANCIA_PUNTO9(ByVal dato As String)
        TOLERANCIA_PUNTO9 = dato
End Property
Public Property Let setTOLERANCIA_PUNTO10(ByVal dato As String)
        TOLERANCIA_PUNTO10 = dato
End Property
Public Property Let setUNIDAD_TOLERANCIA(ByVal dato As String)
        UNIDAD_TOLERANCIA = dato
End Property
Public Property Let setEVALUACION_TOLERANCIA(ByVal dato As String)
        EVALUACION_TOLERANCIA = dato
End Property
Public Property Let setCRITERIO_DE_AJUSTE_PUNTO1(ByVal dato As String)
        CRITERIO_DE_AJUSTE_PUNTO1 = dato
End Property
Public Property Let setCRITERIO_DE_AJUSTE_PUNTO2(ByVal dato As String)
        CRITERIO_DE_AJUSTE_PUNTO2 = dato
End Property
Public Property Let setCRITERIO_DE_AJUSTE_PUNTO3(ByVal dato As String)
        CRITERIO_DE_AJUSTE_PUNTO3 = dato
End Property
Public Property Let setCRITERIO_DE_AJUSTE_PUNTO4(ByVal dato As String)
        CRITERIO_DE_AJUSTE_PUNTO4 = dato
End Property
Public Property Let setCRITERIO_DE_AJUSTE_PUNTO5(ByVal dato As String)
        CRITERIO_DE_AJUSTE_PUNTO5 = dato
End Property
Public Property Let setCRITERIO_DE_AJUSTE_PUNTO6(ByVal dato As String)
        CRITERIO_DE_AJUSTE_PUNTO6 = dato
End Property
Public Property Let setCRITERIO_DE_AJUSTE_PUNTO7(ByVal dato As String)
        CRITERIO_DE_AJUSTE_PUNTO7 = dato
End Property
Public Property Let setCRITERIO_DE_AJUSTE_PUNTO8(ByVal dato As String)
        CRITERIO_DE_AJUSTE_PUNTO8 = dato
End Property
Public Property Let setCRITERIO_DE_AJUSTE_PUNTO9(ByVal dato As String)
        CRITERIO_DE_AJUSTE_PUNTO9 = dato
End Property
Public Property Let setCRITERIO_DE_AJUSTE_PUNTO10(ByVal dato As String)
        CRITERIO_DE_AJUSTE_PUNTO10 = dato
End Property
Public Property Let setUNIDAD_AJUSTE(ByVal dato As String)
        UNIDAD_AJUSTE = dato
End Property
Public Property Let setEVALUACION_AJUSTE(ByVal dato As String)
        EVALUACION_AJUSTE = dato
End Property
Public Property Let setTRAZMET(ByVal dato As String)
        TRAZMET = dato
End Property
Public Property Let setAUXILIAR2(ByVal dato As String)
        AUXILIAR2 = dato
End Property
Public Property Let setAUXILIAR3(ByVal dato As String)
        AUXILIAR3 = dato
End Property
Public Property Let setINTERVALO_CALIBRACION_MESES(ByVal dato As String)
        INTERVALO_CALIBRACION_MESES = dato
End Property
Public Property Let setI_USO(ByVal dato As String)
        I_USO = dato
End Property
Public Property Let setNOMBRE_BANCO(ByVal dato As String)
        NOMBRE_BANCO = dato
End Property
Public Property Let setIDENTIFICACION_BANCO(ByVal dato As String)
        IDENTIFICACION_BANCO = dato
End Property
Public Property Let setCMC_TOL_PUNTO_1(ByVal dato As String)
        CMC_TOL_PUNTO_1 = dato
End Property
Public Property Let setCMC_TOL_PUNTO_2(ByVal dato As String)
        CMC_TOL_PUNTO_2 = dato
End Property
Public Property Let setCMC_TOL_PUNTO_3(ByVal dato As String)
        CMC_TOL_PUNTO_3 = dato
End Property
Public Property Let setCMC_TOL_PUNTO_4(ByVal dato As String)
        CMC_TOL_PUNTO_4 = dato
End Property
Public Property Let setCMC_TOL_PUNTO_5(ByVal dato As String)
        CMC_TOL_PUNTO_5 = dato
End Property
Public Property Let setCMC_TOL_PUNTO_6(ByVal dato As String)
        CMC_TOL_PUNTO_6 = dato
End Property
Public Property Let setCMC_TOL_PUNTO_7(ByVal dato As String)
        CMC_TOL_PUNTO_7 = dato
End Property
Public Property Let setCMC_TOL_PUNTO_8(ByVal dato As String)
        CMC_TOL_PUNTO_8 = dato
End Property
Public Property Let setCMC_TOL_PUNTO_9(ByVal dato As String)
        CMC_TOL_PUNTO_9 = dato
End Property
Public Property Let setCMC_TOL_PUNTO_10(ByVal dato As String)
        CMC_TOL_PUNTO_10 = dato
End Property
Public Property Let setTRANSD_PUNTO1(ByVal dato As String)
        TRANSD_PUNTO1 = dato
End Property
Public Property Let setTRANSD_PUNTO2(ByVal dato As String)
        TRANSD_PUNTO2 = dato
End Property
Public Property Let setTRANSD_PUNTO3(ByVal dato As String)
        TRANSD_PUNTO3 = dato
End Property
Public Property Let setTRANSD_PUNTO4(ByVal dato As String)
        TRANSD_PUNTO4 = dato
End Property
Public Property Let setTRANSD_PUNTO5(ByVal dato As String)
        TRANSD_PUNTO5 = dato
End Property
Public Property Let setTRANSD_PUNTO6(ByVal dato As String)
        TRANSD_PUNTO6 = dato
End Property
Public Property Let setTRANSD_PUNTO7(ByVal dato As String)
        TRANSD_PUNTO7 = dato
End Property
Public Property Let setTRANSD_PUNTO8(ByVal dato As String)
        TRANSD_PUNTO8 = dato
End Property
Public Property Let setTRANSD_PUNTO9(ByVal dato As String)
        TRANSD_PUNTO9 = dato
End Property
Public Property Let setTRANSD_PUNTO10(ByVal dato As String)
        TRANSD_PUNTO10 = dato
End Property
Public Property Let setCLASE_DE_INFORME(ByVal dato As String)
        CLASE_DE_INFORME = dato
End Property
Public Property Let setTRANS1_CERT(ByVal dato As String)
        TRANS1_CERT = dato
End Property
Public Property Let setTRANS2_CERT(ByVal dato As String)
        TRANS2_CERT = dato
End Property
Public Property Let setTRANS3_CERT(ByVal dato As String)
        TRANS3_CERT = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSMANOMETRO_VACUOMETRO
'***************************************************************
Public Property Get getEQUIPO_ID() As String
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getFLUIDO_DE_CAL() As String
        getFLUIDO_DE_CAL = FLUIDO_DE_CAL
End Property
Public Property Get getFLUIDO_CAL_INGLES() As String
        getFLUIDO_CAL_INGLES = FLUIDO_CAL_INGLES
End Property
Public Property Get getCRITERIO_ACEPTACION() As String
        getCRITERIO_ACEPTACION = CRITERIO_ACEPTACION
End Property
Public Property Get getACCEPTANCE_CRITERIA() As String
        getACCEPTANCE_CRITERIA = ACCEPTANCE_CRITERIA
End Property
Public Property Get getPUNTO1() As String
        getPUNTO1 = PUNTO1
End Property
Public Property Get getPUNTO2() As String
        getPUNTO2 = PUNTO2
End Property
Public Property Get getPUNTO3() As String
        getPUNTO3 = PUNTO3
End Property
Public Property Get getPUNTO4() As String
        getPUNTO4 = PUNTO4
End Property
Public Property Get getPUNTO5() As String
        getPUNTO5 = PUNTO5
End Property
Public Property Get getPUNTO6() As String
        getPUNTO6 = PUNTO6
End Property
Public Property Get getPUNTO7() As String
        getPUNTO7 = PUNTO7
End Property
Public Property Get getPUNTO8() As String
        getPUNTO8 = PUNTO8
End Property
Public Property Get getPUNTO9() As String
        getPUNTO9 = PUNTO9
End Property
Public Property Get getPUNTO10() As String
        getPUNTO10 = PUNTO10
End Property
Public Property Get getPNT_F_ASOCIADO() As String
        getPNT_F_ASOCIADO = PNT_F_ASOCIADO
End Property
Public Property Get getCMC_PUNTO1() As String
        getCMC_PUNTO1 = CMC_PUNTO1
End Property
Public Property Get getCMC_PUNTO2() As String
        getCMC_PUNTO2 = CMC_PUNTO2
End Property
Public Property Get getCMC_PUNTO3() As String
        getCMC_PUNTO3 = CMC_PUNTO3
End Property
Public Property Get getCMC_PUNTO4() As String
        getCMC_PUNTO4 = CMC_PUNTO4
End Property
Public Property Get getCMC_PUNTO5() As String
        getCMC_PUNTO5 = CMC_PUNTO5
End Property
Public Property Get getCMC_PUNTO6() As String
        getCMC_PUNTO6 = CMC_PUNTO6
End Property
Public Property Get getCMC_PUNTO7() As String
        getCMC_PUNTO7 = CMC_PUNTO7
End Property
Public Property Get getCMC_PUNTO8() As String
        getCMC_PUNTO8 = CMC_PUNTO8
End Property
Public Property Get getCMC_PUNTO9() As String
        getCMC_PUNTO9 = CMC_PUNTO9
End Property
Public Property Get getCMC_PUNTO10() As String
        getCMC_PUNTO10 = CMC_PUNTO10
End Property
Public Property Get getRANGO_UTIL_CAL() As String
        getRANGO_UTIL_CAL = RANGO_UTIL_CAL
End Property
Public Property Get getTOLERANCIA_PUNTO1() As String
        getTOLERANCIA_PUNTO1 = TOLERANCIA_PUNTO1
End Property
Public Property Get getTOLERANCIA_PUNTO2() As String
        getTOLERANCIA_PUNTO2 = TOLERANCIA_PUNTO2
End Property
Public Property Get getTOLERANCIA_PUNTO3() As String
        getTOLERANCIA_PUNTO3 = TOLERANCIA_PUNTO3
End Property
Public Property Get getTOLERANCIA_PUNTO4() As String
        getTOLERANCIA_PUNTO4 = TOLERANCIA_PUNTO4
End Property
Public Property Get getTOLERANCIA_PUNTO5() As String
        getTOLERANCIA_PUNTO5 = TOLERANCIA_PUNTO5
End Property
Public Property Get getTOLERANCIA_PUNTO6() As String
        getTOLERANCIA_PUNTO6 = TOLERANCIA_PUNTO6
End Property
Public Property Get getTOLERANCIA_PUNTO7() As String
        getTOLERANCIA_PUNTO7 = TOLERANCIA_PUNTO7
End Property
Public Property Get getTOLERANCIA_PUNTO8() As String
        getTOLERANCIA_PUNTO8 = TOLERANCIA_PUNTO8
End Property
Public Property Get getTOLERANCIA_PUNTO9() As String
        getTOLERANCIA_PUNTO9 = TOLERANCIA_PUNTO9
End Property
Public Property Get getTOLERANCIA_PUNTO10() As String
        getTOLERANCIA_PUNTO10 = TOLERANCIA_PUNTO10
End Property
Public Property Get getUNIDAD_TOLERANCIA() As String
        getUNIDAD_TOLERANCIA = UNIDAD_TOLERANCIA
End Property
Public Property Get getEVALUACION_TOLERANCIA() As String
        getEVALUACION_TOLERANCIA = EVALUACION_TOLERANCIA
End Property
Public Property Get getCRITERIO_DE_AJUSTE_PUNTO1() As String
        getCRITERIO_DE_AJUSTE_PUNTO1 = CRITERIO_DE_AJUSTE_PUNTO1
End Property
Public Property Get getCRITERIO_DE_AJUSTE_PUNTO2() As String
        getCRITERIO_DE_AJUSTE_PUNTO2 = CRITERIO_DE_AJUSTE_PUNTO2
End Property
Public Property Get getCRITERIO_DE_AJUSTE_PUNTO3() As String
        getCRITERIO_DE_AJUSTE_PUNTO3 = CRITERIO_DE_AJUSTE_PUNTO3
End Property
Public Property Get getCRITERIO_DE_AJUSTE_PUNTO4() As String
        getCRITERIO_DE_AJUSTE_PUNTO4 = CRITERIO_DE_AJUSTE_PUNTO4
End Property
Public Property Get getCRITERIO_DE_AJUSTE_PUNTO5() As String
        getCRITERIO_DE_AJUSTE_PUNTO5 = CRITERIO_DE_AJUSTE_PUNTO5
End Property
Public Property Get getCRITERIO_DE_AJUSTE_PUNTO6() As String
        getCRITERIO_DE_AJUSTE_PUNTO6 = CRITERIO_DE_AJUSTE_PUNTO6
End Property
Public Property Get getCRITERIO_DE_AJUSTE_PUNTO7() As String
        getCRITERIO_DE_AJUSTE_PUNTO7 = CRITERIO_DE_AJUSTE_PUNTO7
End Property
Public Property Get getCRITERIO_DE_AJUSTE_PUNTO8() As String
        getCRITERIO_DE_AJUSTE_PUNTO8 = CRITERIO_DE_AJUSTE_PUNTO8
End Property
Public Property Get getCRITERIO_DE_AJUSTE_PUNTO9() As String
        getCRITERIO_DE_AJUSTE_PUNTO9 = CRITERIO_DE_AJUSTE_PUNTO9
End Property
Public Property Get getCRITERIO_DE_AJUSTE_PUNTO10() As String
        getCRITERIO_DE_AJUSTE_PUNTO10 = CRITERIO_DE_AJUSTE_PUNTO10
End Property
Public Property Get getUNIDAD_AJUSTE() As String
        getUNIDAD_AJUSTE = UNIDAD_AJUSTE
End Property
Public Property Get getEVALUACION_AJUSTE() As String
        getEVALUACION_AJUSTE = EVALUACION_AJUSTE
End Property
Public Property Get getTRAZMET() As String
        getTRAZMET = TRAZMET
End Property
Public Property Get getAUXILIAR2() As String
        getAUXILIAR2 = AUXILIAR2
End Property
Public Property Get getAUXILIAR3() As String
        getAUXILIAR3 = AUXILIAR3
End Property
Public Property Get getINTERVALO_CALIBRACION_MESES() As String
        getINTERVALO_CALIBRACION_MESES = INTERVALO_CALIBRACION_MESES
End Property
Public Property Get getI_USO() As String
        getI_USO = I_USO
End Property
Public Property Get getNOMBRE_BANCO() As String
        getNOMBRE_BANCO = NOMBRE_BANCO
End Property
Public Property Get getIDENTIFICACION_BANCO() As String
        getIDENTIFICACION_BANCO = IDENTIFICACION_BANCO
End Property
Public Property Get getCMC_TOL_PUNTO_1() As String
        getCMC_TOL_PUNTO_1 = CMC_TOL_PUNTO_1
End Property
Public Property Get getCMC_TOL_PUNTO_2() As String
        getCMC_TOL_PUNTO_2 = CMC_TOL_PUNTO_2
End Property
Public Property Get getCMC_TOL_PUNTO_3() As String
        getCMC_TOL_PUNTO_3 = CMC_TOL_PUNTO_3
End Property
Public Property Get getCMC_TOL_PUNTO_4() As String
        getCMC_TOL_PUNTO_4 = CMC_TOL_PUNTO_4
End Property
Public Property Get getCMC_TOL_PUNTO_5() As String
        getCMC_TOL_PUNTO_5 = CMC_TOL_PUNTO_5
End Property
Public Property Get getCMC_TOL_PUNTO_6() As String
        getCMC_TOL_PUNTO_6 = CMC_TOL_PUNTO_6
End Property
Public Property Get getCMC_TOL_PUNTO_7() As String
        getCMC_TOL_PUNTO_7 = CMC_TOL_PUNTO_7
End Property
Public Property Get getCMC_TOL_PUNTO_8() As String
        getCMC_TOL_PUNTO_8 = CMC_TOL_PUNTO_8
End Property
Public Property Get getCMC_TOL_PUNTO_9() As String
        getCMC_TOL_PUNTO_9 = CMC_TOL_PUNTO_9
End Property
Public Property Get getCMC_TOL_PUNTO_10() As String
        getCMC_TOL_PUNTO_10 = CMC_TOL_PUNTO_10
End Property
Public Property Get getTRANSD_PUNTO1() As String
        getTRANSD_PUNTO1 = TRANSD_PUNTO1
End Property
Public Property Get getTRANSD_PUNTO2() As String
        getTRANSD_PUNTO2 = TRANSD_PUNTO2
End Property
Public Property Get getTRANSD_PUNTO3() As String
        getTRANSD_PUNTO3 = TRANSD_PUNTO3
End Property
Public Property Get getTRANSD_PUNTO4() As String
        getTRANSD_PUNTO4 = TRANSD_PUNTO4
End Property
Public Property Get getTRANSD_PUNTO5() As String
        getTRANSD_PUNTO5 = TRANSD_PUNTO5
End Property
Public Property Get getTRANSD_PUNTO6() As String
        getTRANSD_PUNTO6 = TRANSD_PUNTO6
End Property
Public Property Get getTRANSD_PUNTO7() As String
        getTRANSD_PUNTO7 = TRANSD_PUNTO7
End Property
Public Property Get getTRANSD_PUNTO8() As String
        getTRANSD_PUNTO8 = TRANSD_PUNTO8
End Property
Public Property Get getTRANSD_PUNTO9() As String
        getTRANSD_PUNTO9 = TRANSD_PUNTO9
End Property
Public Property Get getTRANSD_PUNTO10() As String
        getTRANSD_PUNTO10 = TRANSD_PUNTO10
End Property
Public Property Get getCLASE_DE_INFORME() As String
        getCLASE_DE_INFORME = CLASE_DE_INFORME
End Property
Public Property Get getTRANS1_CERT() As String
        getTRANS1_CERT = TRANS1_CERT
End Property
Public Property Get getTRANS2_CERT() As String
        getTRANS2_CERT = TRANS2_CERT
End Property
Public Property Get getTRANS3_CERT() As String
        getTRANS3_CERT = TRANS3_CERT
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSMANOMETRO_VACUOMETRO
'***************************************************************
Public Function Carga(ID As String) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM MANOMETRO_VACUOMETRO WHERE EQUIPO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                EQUIPO_ID = rs("EQUIPO_ID")
                FLUIDO_DE_CAL = rs("FLUIDO_DE_CAL")
                FLUIDO_CAL_INGLES = rs("FLUIDO_CAL_INGLES")
                CRITERIO_ACEPTACION = rs("CRITERIO_ACEPTACION")
                ACCEPTANCE_CRITERIA = rs("ACCEPTANCE_CRITERIA")
                PUNTO1 = rs("PUNTO1")
                PUNTO2 = rs("PUNTO2")
                PUNTO3 = rs("PUNTO3")
                PUNTO4 = rs("PUNTO4")
                PUNTO5 = rs("PUNTO5")
                PUNTO6 = rs("PUNTO6")
                PUNTO7 = rs("PUNTO7")
                PUNTO8 = rs("PUNTO8")
                PUNTO9 = rs("PUNTO9")
                PUNTO10 = rs("PUNTO10")
                PNT_F_ASOCIADO = rs("PNT_F_ASOCIADO")
                CMC_PUNTO1 = rs("CMC_PUNTO1")
                CMC_PUNTO2 = rs("CMC_PUNTO2")
                CMC_PUNTO3 = rs("CMC_PUNTO3")
                CMC_PUNTO4 = rs("CMC_PUNTO4")
                CMC_PUNTO5 = rs("CMC_PUNTO5")
                CMC_PUNTO6 = rs("CMC_PUNTO6")
                CMC_PUNTO7 = rs("CMC_PUNTO7")
                CMC_PUNTO8 = rs("CMC_PUNTO8")
                CMC_PUNTO9 = rs("CMC_PUNTO9")
                CMC_PUNTO10 = rs("CMC_PUNTO10")
                RANGO_UTIL_CAL = rs("RANGO_UTIL_CAL")
                TOLERANCIA_PUNTO1 = rs("TOLERANCIA_PUNTO1")
                TOLERANCIA_PUNTO2 = rs("TOLERANCIA_PUNTO2")
                TOLERANCIA_PUNTO3 = rs("TOLERANCIA_PUNTO3")
                TOLERANCIA_PUNTO4 = rs("TOLERANCIA_PUNTO4")
                TOLERANCIA_PUNTO5 = rs("TOLERANCIA_PUNTO5")
                TOLERANCIA_PUNTO6 = rs("TOLERANCIA_PUNTO6")
                TOLERANCIA_PUNTO7 = rs("TOLERANCIA_PUNTO7")
                TOLERANCIA_PUNTO8 = rs("TOLERANCIA_PUNTO8")
                TOLERANCIA_PUNTO9 = rs("TOLERANCIA_PUNTO9")
                TOLERANCIA_PUNTO10 = rs("TOLERANCIA_PUNTO10")
                UNIDAD_TOLERANCIA = rs("UNIDAD_TOLERANCIA")
                EVALUACION_TOLERANCIA = rs("EVALUACION_TOLERANCIA")
                CRITERIO_DE_AJUSTE_PUNTO1 = rs("CRITERIO_DE_AJUSTE_PUNTO1")
                CRITERIO_DE_AJUSTE_PUNTO2 = rs("CRITERIO_DE_AJUSTE_PUNTO2")
                CRITERIO_DE_AJUSTE_PUNTO3 = rs("CRITERIO_DE_AJUSTE_PUNTO3")
                CRITERIO_DE_AJUSTE_PUNTO4 = rs("CRITERIO_DE_AJUSTE_PUNTO4")
                CRITERIO_DE_AJUSTE_PUNTO5 = rs("CRITERIO_DE_AJUSTE_PUNTO5")
                CRITERIO_DE_AJUSTE_PUNTO6 = rs("CRITERIO_DE_AJUSTE_PUNTO6")
                CRITERIO_DE_AJUSTE_PUNTO7 = rs("CRITERIO_DE_AJUSTE_PUNTO7")
                CRITERIO_DE_AJUSTE_PUNTO8 = rs("CRITERIO_DE_AJUSTE_PUNTO8")
                CRITERIO_DE_AJUSTE_PUNTO9 = rs("CRITERIO_DE_AJUSTE_PUNTO9")
                CRITERIO_DE_AJUSTE_PUNTO10 = rs("CRITERIO_DE_AJUSTE_PUNTO10")
                UNIDAD_AJUSTE = rs("UNIDAD_AJUSTE")
                EVALUACION_AJUSTE = rs("EVALUACION_AJUSTE")
                TRAZMET = rs("TRAZMET")
                AUXILIAR2 = rs("AUXILIAR2")
                AUXILIAR3 = rs("AUXILIAR3")
                INTERVALO_CALIBRACION_MESES = rs("INTERVALO_CALIBRACION_MESES")
                I_USO = rs("I_USO")
                NOMBRE_BANCO = rs("NOMBRE_BANCO")
                IDENTIFICACION_BANCO = rs("IDENTIFICACION_BANCO")
                CMC_TOL_PUNTO_1 = rs("CMC_TOL_PUNTO_1")
                CMC_TOL_PUNTO_2 = rs("CMC_TOL_PUNTO_2")
                CMC_TOL_PUNTO_3 = rs("CMC_TOL_PUNTO_3")
                CMC_TOL_PUNTO_4 = rs("CMC_TOL_PUNTO_4")
                CMC_TOL_PUNTO_5 = rs("CMC_TOL_PUNTO_5")
                CMC_TOL_PUNTO_6 = rs("CMC_TOL_PUNTO_6")
                CMC_TOL_PUNTO_7 = rs("CMC_TOL_PUNTO_7")
                CMC_TOL_PUNTO_8 = rs("CMC_TOL_PUNTO_8")
                CMC_TOL_PUNTO_9 = rs("CMC_TOL_PUNTO_9")
                CMC_TOL_PUNTO_10 = rs("CMC_TOL_PUNTO_10")
                TRANSD_PUNTO1 = rs("TRANSD_PUNTO1")
                TRANSD_PUNTO2 = rs("TRANSD_PUNTO2")
                TRANSD_PUNTO3 = rs("TRANSD_PUNTO3")
                TRANSD_PUNTO4 = rs("TRANSD_PUNTO4")
                TRANSD_PUNTO5 = rs("TRANSD_PUNTO5")
                TRANSD_PUNTO6 = rs("TRANSD_PUNTO6")
                TRANSD_PUNTO7 = rs("TRANSD_PUNTO7")
                TRANSD_PUNTO8 = rs("TRANSD_PUNTO8")
                TRANSD_PUNTO9 = rs("TRANSD_PUNTO9")
                TRANSD_PUNTO10 = rs("TRANSD_PUNTO10")
                CLASE_DE_INFORME = rs("CLASE_DE_INFORME")
                TRANS1_CERT = rs("TRANS1_CERT")
                TRANS2_CERT = rs("TRANS2_CERT")
                TRANS3_CERT = rs("TRANS3_CERT")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsManometro_vacuometro)", vbCritical, Err.Description
End Function
Public Function Insertar() As String
        On Error GoTo fallo
        Dim consulta As String
        consulta = "REPLACE INTO MANOMETRO_VACUOMETRO " & _
                   "  VALUES (" & _
                        EQUIPO_ID & "," & "'" & FLUIDO_DE_CAL & "'" & "," & "'" & FLUIDO_CAL_INGLES & "'" & "," & _
                        "'" & CRITERIO_ACEPTACION & "'" & "," & "'" & ACCEPTANCE_CRITERIA & "'" & ","
        consulta = consulta & IIf(PUNTO1 = "", "null", PUNTO1) & ","
        consulta = consulta & IIf(PUNTO2 = "", "null", PUNTO2) & ","
        consulta = consulta & IIf(PUNTO3 = "", "null", PUNTO3) & ","
        consulta = consulta & IIf(PUNTO4 = "", "null", PUNTO4) & ","
        consulta = consulta & IIf(PUNTO5 = "", "null", PUNTO5) & ","
        consulta = consulta & IIf(PUNTO6 = "", "null", PUNTO6) & ","
        consulta = consulta & IIf(PUNTO7 = "", "null", PUNTO7) & ","
        consulta = consulta & IIf(PUNTO8 = "", "null", PUNTO8) & ","
        consulta = consulta & IIf(PUNTO9 = "", "null", PUNTO9) & ","
        consulta = consulta & IIf(PUNTO10 = "", "null", PUNTO10) & ","
        consulta = consulta & IIf(PNT_F_ASOCIADO = "", "null", PNT_F_ASOCIADO) & ","
        consulta = consulta & IIf(CMC_PUNTO1 = "", "null", CMC_PUNTO1) & ","
        consulta = consulta & IIf(CMC_PUNTO2 = "", "null", CMC_PUNTO2) & ","
        consulta = consulta & IIf(CMC_PUNTO3 = "", "null", CMC_PUNTO3) & ","
        consulta = consulta & IIf(CMC_PUNTO4 = "", "null", CMC_PUNTO4) & ","
        consulta = consulta & IIf(CMC_PUNTO5 = "", "null", CMC_PUNTO5) & ","
        consulta = consulta & IIf(CMC_PUNTO6 = "", "null", CMC_PUNTO6) & ","
        consulta = consulta & IIf(CMC_PUNTO7 = "", "null", CMC_PUNTO7) & ","
        consulta = consulta & IIf(CMC_PUNTO8 = "", "null", CMC_PUNTO8) & ","
        consulta = consulta & IIf(CMC_PUNTO9 = "", "null", CMC_PUNTO9) & ","
        consulta = consulta & IIf(CMC_PUNTO10 = "", "null", CMC_PUNTO10) & ","
        consulta = consulta & IIf(RANGO_UTIL_CAL = "", "null", RANGO_UTIL_CAL) & ","
        consulta = consulta & IIf(TOLERANCIA_PUNTO1 = "", "null", TOLERANCIA_PUNTO1) & ","
        consulta = consulta & IIf(TOLERANCIA_PUNTO2 = "", "null", TOLERANCIA_PUNTO2) & ","
        consulta = consulta & IIf(TOLERANCIA_PUNTO3 = "", "null", TOLERANCIA_PUNTO3) & ","
        consulta = consulta & IIf(TOLERANCIA_PUNTO4 = "", "null", TOLERANCIA_PUNTO4) & ","
        consulta = consulta & IIf(TOLERANCIA_PUNTO5 = "", "null", TOLERANCIA_PUNTO5) & ","
        consulta = consulta & IIf(TOLERANCIA_PUNTO6 = "", "null", TOLERANCIA_PUNTO6) & ","
        consulta = consulta & IIf(TOLERANCIA_PUNTO7 = "", "null", TOLERANCIA_PUNTO7) & ","
        consulta = consulta & IIf(TOLERANCIA_PUNTO8 = "", "null", TOLERANCIA_PUNTO8) & ","
        consulta = consulta & IIf(TOLERANCIA_PUNTO9 = "", "null", TOLERANCIA_PUNTO9) & ","
        consulta = consulta & IIf(TOLERANCIA_PUNTO10 = "", "null", TOLERANCIA_PUNTO10) & ","
        consulta = consulta & "'" & UNIDAD_TOLERANCIA & "',"
        consulta = consulta & "'" & EVALUACION_TOLERANCIA & "',"
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO1 = "", "null", CRITERIO_DE_AJUSTE_PUNTO1) & ","
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO2 = "", "null", CRITERIO_DE_AJUSTE_PUNTO2) & ","
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO3 = "", "null", CRITERIO_DE_AJUSTE_PUNTO3) & ","
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO4 = "", "null", CRITERIO_DE_AJUSTE_PUNTO4) & ","
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO5 = "", "null", CRITERIO_DE_AJUSTE_PUNTO5) & ","
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO6 = "", "null", CRITERIO_DE_AJUSTE_PUNTO6) & ","
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO7 = "", "null", CRITERIO_DE_AJUSTE_PUNTO7) & ","
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO8 = "", "null", CRITERIO_DE_AJUSTE_PUNTO8) & ","
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO9 = "", "null", CRITERIO_DE_AJUSTE_PUNTO9) & ","
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO10 = "", "null", CRITERIO_DE_AJUSTE_PUNTO10) & ","
        consulta = consulta & "'" & UNIDAD_AJUSTE & "',"
        consulta = consulta & "'" & EVALUACION_AJUSTE & "',"
        consulta = consulta & "'" & TRAZMET & "',"
        consulta = consulta & "'" & AUXILIAR2 & "',"
        consulta = consulta & "'" & AUXILIAR3 & "',"
        consulta = consulta & "'" & IIf(INTERVALO_CALIBRACION_MESES = "", "", INTERVALO_CALIBRACION_MESES) & "',"
        consulta = consulta & "'" & I_USO & "',"
        consulta = consulta & "'" & NOMBRE_BANCO & "',"
        consulta = consulta & "'" & IDENTIFICACION_BANCO & "',"
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO1 = "", "null", CRITERIO_DE_AJUSTE_PUNTO1) & ","
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO2 = "", "null", CRITERIO_DE_AJUSTE_PUNTO2) & ","
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO3 = "", "null", CRITERIO_DE_AJUSTE_PUNTO3) & ","
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO4 = "", "null", CRITERIO_DE_AJUSTE_PUNTO4) & ","
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO5 = "", "null", CRITERIO_DE_AJUSTE_PUNTO5) & ","
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO6 = "", "null", CRITERIO_DE_AJUSTE_PUNTO6) & ","
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO7 = "", "null", CRITERIO_DE_AJUSTE_PUNTO7) & ","
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO8 = "", "null", CRITERIO_DE_AJUSTE_PUNTO8) & ","
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO9 = "", "null", CRITERIO_DE_AJUSTE_PUNTO9) & ","
        consulta = consulta & IIf(CRITERIO_DE_AJUSTE_PUNTO10 = "", "null", CRITERIO_DE_AJUSTE_PUNTO10) & ","
        consulta = consulta & IIf(TRANSD_PUNTO1 = "", "null", TRANSD_PUNTO1) & ","
        consulta = consulta & IIf(TRANSD_PUNTO2 = "", "null", TRANSD_PUNTO2) & ","
        consulta = consulta & IIf(TRANSD_PUNTO3 = "", "null", TRANSD_PUNTO3) & ","
        consulta = consulta & IIf(TRANSD_PUNTO4 = "", "null", TRANSD_PUNTO4) & ","
        consulta = consulta & IIf(TRANSD_PUNTO5 = "", "null", TRANSD_PUNTO5) & ","
        consulta = consulta & IIf(TRANSD_PUNTO6 = "", "null", TRANSD_PUNTO6) & ","
        consulta = consulta & IIf(TRANSD_PUNTO7 = "", "null", TRANSD_PUNTO7) & ","
        consulta = consulta & IIf(TRANSD_PUNTO8 = "", "null", TRANSD_PUNTO8) & ","
        consulta = consulta & IIf(TRANSD_PUNTO9 = "", "null", TRANSD_PUNTO9) & ","
        consulta = consulta & IIf(TRANSD_PUNTO10 = "", "null", TRANSD_PUNTO10) & ","
        consulta = consulta & "'" & CLASE_DE_INFORME & "',"
        consulta = consulta & IIf(TRANS1_CERT = "", "null", TRANS1_CERT) & ","
        consulta = consulta & IIf(TRANS2_CERT = "", "null", TRANS2_CERT) & ","
        consulta = consulta & IIf(TRANS3_CERT = "", "null", TRANS3_CERT) & ")"
        
        execute_bd consulta
        Insertar = EQUIPO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsManometro_vacuometro)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM MANOMETRO_VACUOMETRO " & _
                "    WHERE EQUIPO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsManometro_vacuometro)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM MANOMETRO_VACUOMETRO ORDER BY EQUIPO_ID"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSMANOMETRO_VACUOMETRO
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

