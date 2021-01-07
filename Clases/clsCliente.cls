VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*****************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLIENTE      *
'*****************************************************
Private id_cliente As Long
Private nombre As String
Private DIRECCION As String
Private COD_POSTAL As Long
Private CIF As String
Private TELEFONO As String
Private FAX As String
Private responsable As String
Private responsable_OTROS As String
Private RESPONSABLE_METROLOGIA As String
Private CARGO As String
Private tipo As String
Private EMAIL As String
Private EMAIL2 As String
Private EMAIL_FACTURACION As String
Private EMAIL_METROLOGIA As String
Private banco As String
Private cuenta As String
Private OBSERVACIONES As String
Private WEB As String
Private anulado As Long
Private Centro As String
Private SECCION As String
Private FACTURA_DETERMINACIONES As Integer
Private EADS As Integer
Private AIRBUS As Integer
Private IBERIA As Integer
Private AGROALIMENTARIO As Integer
Private EXTRANJERO As Integer
Private INTRA As Integer
Private IDIOMA_FACTURA As Integer
Private MUNICIPIO_ID As Long
Private PROVINCIA_ID As Long
Private PAIS_ID As Long
Private TARIFA_ID As Long
Private CLAVEWEB As String
Private CC As String
Private FP_ID As Integer
Private ENVIO_DIRECCION As String
Private ENVIO_COD_POSTAL As Long
Private ENVIO_PAIS_ID As Long
Private ENVIO_PROVINCIA_ID As Long
Private ENVIO_MUNICIPIO_ID As Long
Private INFORMES_DIRECCION As String
Private INFORMES_COD_POSTAL As Long
Private INFORMES_PAIS_ID As Long
Private INFORMES_PROVINCIA_ID As Long
Private INFORMES_MUNICIPIO_ID As Long
'M1275-I
Private FACTURA_ELECTRONICA As Integer
Private PLANT_ID As Integer
'M1275-F
'E0104-I
'E0200-I
'Private ES_SUBCONTRATA As Integer
'E0200-F
'E0104-F
'*****************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE MUESTRA    *
'*****************************************************
Public Property Let setID_CLIENTE(ByVal dato As Long)
    id_cliente = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Let setDIRECCION(ByVal dato As String)
    DIRECCION = Trim(dato)
End Property
Public Property Let setCOD_POSTAL(ByVal dato As Long)
    COD_POSTAL = dato
End Property
Public Property Let setCIF(ByVal dato As String)
    CIF = Trim(dato)
End Property
Public Property Let setTELEFONO(ByVal dato As String)
    TELEFONO = Trim(dato)
End Property
Public Property Let setFAX(ByVal dato As String)
    FAX = Trim(dato)
End Property
Public Property Let setRESPONSABLE(ByVal dato As String)
    responsable = Trim(dato)
End Property
Public Property Let setRESPONSABLE_OTROS(ByVal dato As String)
    responsable_OTROS = Trim(dato)
End Property
Public Property Let setRESPONSABLE_METROLOGIA(ByVal dato As String)
    RESPONSABLE_METROLOGIA = Trim(dato)
End Property
Public Property Let setCARGO(ByVal dato As String)
    CARGO = Trim(dato)
End Property
Public Property Let setTIPO(ByVal dato As String)
    tipo = Trim(dato)
End Property
Public Property Let setEMAIL(ByVal dato As String)
    EMAIL = Trim(dato)
End Property
Public Property Let setEMAIL2(ByVal dato As String)
    EMAIL2 = Trim(dato)
End Property
Public Property Let setEMAIL_FACTURACION(ByVal dato As String)
    EMAIL_FACTURACION = Trim(dato)
End Property
Public Property Let setEMAIL_METROLOGIA(ByVal dato As String)
    EMAIL_METROLOGIA = Trim(dato)
End Property
Public Property Let setBANCO(ByVal dato As String)
    banco = Trim(dato)
End Property
Public Property Let setCUENTA(ByVal dato As String)
    cuenta = Trim(dato)
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
    OBSERVACIONES = Trim(dato)
End Property
Public Property Let setWEB(ByVal dato As String)
    WEB = Trim(dato)
End Property
Public Property Let setANULADO(ByVal dato As Long)
    anulado = dato
End Property
Public Property Let setCENTRO(ByVal dato As String)
    Centro = Trim(dato)
End Property
Public Property Let setSECCION(ByVal dato As String)
    SECCION = Trim(dato)
End Property
Public Property Let setFACTURA_DETERMINACIONES(ByVal dato As Integer)
    FACTURA_DETERMINACIONES = dato
End Property
Public Property Let setEADS(ByVal dato As Integer)
    EADS = dato
End Property
Public Property Let setAIRBUS(ByVal dato As Integer)
    AIRBUS = dato
End Property
Public Property Let setIBERIA(ByVal dato As Integer)
    IBERIA = dato
End Property
Public Property Let setAGROALIMENTARIO(ByVal dato As Integer)
    AGROALIMENTARIO = dato
End Property
Public Property Let setEXTRANJERO(ByVal dato As Integer)
    EXTRANJERO = dato
End Property
Public Property Let setINTRA(ByVal dato As Integer)
    INTRA = dato
End Property
Public Property Let setIDIOMA_FACTURA(ByVal dato As Integer)
    IDIOMA_FACTURA = dato
End Property
Public Property Let setCC(ByVal dato As String)
    CC = dato
End Property
Public Property Let setFP_ID(ByVal dato As Integer)
    FP_ID = dato
End Property
Public Property Let setMUNICIPIO_ID(ByVal dato As Long)
    MUNICIPIO_ID = dato
End Property
Public Property Let setPROVINCIA_ID(ByVal dato As Long)
    PROVINCIA_ID = dato
End Property
Public Property Let setPAIS_ID(ByVal dato As Long)
    PAIS_ID = dato
End Property
Public Property Let setTARIFA_ID(ByVal dato As Long)
    TARIFA_ID = dato
End Property
Public Property Let setCLAVEWEB(ByVal dato As String)
    CLAVEWEB = Trim(dato)
End Property

Public Property Let setENVIO_DIRECCION(ByVal dato As String)
    ENVIO_DIRECCION = Trim(dato)
End Property
Public Property Let setENVIO_COD_POSTAL(ByVal dato As Long)
    ENVIO_COD_POSTAL = dato
End Property
Public Property Let setENVIO_MUNICIPIO_ID(ByVal dato As Long)
    ENVIO_MUNICIPIO_ID = dato
End Property
Public Property Let setENVIO_PROVINCIA_ID(ByVal dato As Long)
    ENVIO_PROVINCIA_ID = dato
End Property
Public Property Let setENVIO_PAIS_ID(ByVal dato As Long)
    ENVIO_PAIS_ID = dato
End Property
Public Property Let setINFORMES_DIRECCION(ByVal dato As String)
    INFORMES_DIRECCION = Trim(dato)
End Property
Public Property Let setINFORMES_COD_POSTAL(ByVal dato As Long)
    INFORMES_COD_POSTAL = dato
End Property
Public Property Let setINFORMES_MUNICIPIO_ID(ByVal dato As Long)
    INFORMES_MUNICIPIO_ID = dato
End Property
Public Property Let setINFORMES_PROVINCIA_ID(ByVal dato As Long)
    INFORMES_PROVINCIA_ID = dato
End Property
Public Property Let setINFORMES_PAIS_ID(ByVal dato As Long)
    INFORMES_PAIS_ID = dato
End Property
'M1275-I
Public Property Let setFACTURA_ELECTRONICA(ByVal dato As Integer)
    FACTURA_ELECTRONICA = dato
End Property
'M1275-F
Public Property Let setPLANT_ID(ByVal dato As Integer)
    PLANT_ID = dato
End Property

'E0105-I
'E0200-I
'Public Property Let setES_SUBCONTRATA(ByVal dato As Integer)
'    ES_SUBCONTRATA = dato
'End Property
'E0200-F
'E0105-F

'*****************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE MUESTRA      *
'*****************************************************
Public Property Get getID_CLIENTE() As Long
    getID_CLIENTE = id_cliente
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getDIRECCION() As String
    getDIRECCION = DIRECCION
End Property
Public Property Get getCOD_POSTAL() As Long
    getCOD_POSTAL = COD_POSTAL
End Property
Public Property Get getCIF() As String
    getCIF = CIF
End Property
Public Property Get getTELEFONO() As String
    getTELEFONO = TELEFONO
End Property
Public Property Get getFAX() As String
    getFAX = FAX
End Property
Public Property Get getRESPONSABLE() As String
    getRESPONSABLE = responsable
End Property
Public Property Get getRESPONSABLE_OTROS() As String
    getRESPONSABLE_OTROS = responsable_OTROS
End Property
Public Property Get getRESPONSABLE_METROLOGIA() As String
    getRESPONSABLE_METROLOGIA = RESPONSABLE_METROLOGIA
End Property
Public Property Get getCARGO() As String
    getCARGO = CARGO
End Property
Public Property Get getTIPO() As String
    getTIPO = tipo
End Property
Public Property Get getEMAIL() As String
    getEMAIL = EMAIL
End Property
Public Property Get getEMAIL2() As String
    getEMAIL2 = EMAIL2
End Property
Public Property Get getEMAIL_FACTURACION() As String
    getEMAIL_FACTURACION = EMAIL_FACTURACION
End Property
Public Property Get getEMAIL_METROLOGIA() As String
    getEMAIL_METROLOGIA = EMAIL_METROLOGIA
End Property
Public Property Get getBANCO() As String
    getBANCO = banco
End Property
Public Property Get getCUENTA() As String
    getCUENTA = cuenta
End Property
Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getWEB() As String
    getWEB = WEB
End Property
Public Property Get getANULADO() As Long
    getANULADO = anulado
End Property
Public Property Get getCENTRO() As String
    getCENTRO = Centro
End Property
Public Property Get getSECCION() As String
    getSECCION = SECCION
End Property
Public Property Get getFACTURA_DETERMINACIONES() As Integer
    getFACTURA_DETERMINACIONES = FACTURA_DETERMINACIONES
End Property
Public Property Get getEADS() As Integer
    getEADS = EADS
End Property
Public Property Get getAIRBUS() As Integer
    getAIRBUS = AIRBUS
End Property
Public Property Get getIBERIA() As Integer
    getIBERIA = IBERIA
End Property
Public Property Get getAGROALIMENTARIO() As Integer
    getAGROALIMENTARIO = AGROALIMENTARIO
End Property
Public Property Get getEXTRANJERO() As Integer
    getEXTRANJERO = EXTRANJERO
End Property
Public Property Get getINTRA() As Integer
    getINTRA = INTRA
End Property
Public Property Get getIDIOMA_FACTURA() As Integer
    getIDIOMA_FACTURA = IDIOMA_FACTURA
End Property
Public Property Get getFP_ID() As Integer
    getFP_ID = FP_ID
End Property
Public Property Get getMUNICIPIO_ID() As Long
    getMUNICIPIO_ID = MUNICIPIO_ID
End Property
Public Property Get getPROVINCIA_ID() As Long
    getPROVINCIA_ID = PROVINCIA_ID
End Property
Public Property Get getPAIS_ID() As Long
    getPAIS_ID = PAIS_ID
End Property
Public Property Get getTARIFA_ID() As Long
    getTARIFA_ID = TARIFA_ID
End Property
Public Property Get getCLAVEWEB() As String
    getCLAVEWEB = CLAVEWEB
End Property
Public Property Get getCC() As String
    getCC = CC
End Property
Public Property Get getENVIO_DIRECCION() As String
    getENVIO_DIRECCION = ENVIO_DIRECCION
End Property
Public Property Get getENVIO_COD_POSTAL() As Long
    getENVIO_COD_POSTAL = ENVIO_COD_POSTAL
End Property
Public Property Get getENVIO_MUNICIPIO_ID() As Long
    getENVIO_MUNICIPIO_ID = ENVIO_MUNICIPIO_ID
End Property
Public Property Get getENVIO_PROVINCIA_ID() As Long
    getENVIO_PROVINCIA_ID = ENVIO_PROVINCIA_ID
End Property
Public Property Get getENVIO_PAIS_ID() As Long
    getENVIO_PAIS_ID = ENVIO_PAIS_ID
End Property
Public Property Get getINFORMES_DIRECCION() As String
    getINFORMES_DIRECCION = INFORMES_DIRECCION
End Property
Public Property Get getINFORMES_COD_POSTAL() As Long
    getINFORMES_COD_POSTAL = INFORMES_COD_POSTAL
End Property
Public Property Get getINFORMES_MUNICIPIO_ID() As Long
    getINFORMES_MUNICIPIO_ID = INFORMES_MUNICIPIO_ID
End Property
Public Property Get getINFORMES_PROVINCIA_ID() As Long
    getINFORMES_PROVINCIA_ID = INFORMES_PROVINCIA_ID
End Property
Public Property Get getINFORMES_PAIS_ID() As Long
    getINFORMES_PAIS_ID = INFORMES_PAIS_ID
End Property
'M1275-I
Public Property Get getFACTURA_ELECTRONICA() As Integer
    getFACTURA_ELECTRONICA = FACTURA_ELECTRONICA
End Property
'M1275-F
Public Property Get getPLANT_ID() As Integer
    getPLANT_ID = PLANT_ID
End Property
'*****************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE MUESTRA  *
'*****************************************************
Public Function CargaCliente(id_cliente As Long) As Boolean
        On Error GoTo fallo
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM clientes WHERE ID_CLIENTE=" & id_cliente
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            CargaCliente = False
        Else
            id_cliente = rs.Fields("ID_CLIENTE")
            nombre = rs.Fields("NOMBRE") & ""
            DIRECCION = rs.Fields("DIRECCION") & ""
 '           POBLACION = rs.Fields("POBLACION") & ""
 '           provincia = rs.Fields("PROVINCIA") & ""
            If Not IsNull(rs.Fields("COD_POSTAL")) Then
                COD_POSTAL = rs.Fields("COD_POSTAL")
            Else
                COD_POSTAL = 0
            End If
            CIF = rs.Fields("CIF") & ""
            TELEFONO = rs.Fields("TELEFONO") & ""
            FAX = rs.Fields("FAX") & ""
            responsable = rs.Fields("RESPONSABLE") & ""
            responsable_OTROS = rs.Fields("RESPONSABLE_OTROS") & ""
            RESPONSABLE_METROLOGIA = rs.Fields("RESPONSABLE_METROLOGIA") & ""
            CARGO = rs.Fields("CARGO") & ""
            tipo = rs.Fields("TIPO") & ""
            EMAIL = rs.Fields("EMAIL") & ""
            EMAIL2 = rs.Fields("EMAIL2") & ""
            EMAIL_FACTURACION = rs.Fields("EMAIL_FACTURACION") & ""
            EMAIL_METROLOGIA = rs.Fields("EMAIL_METROLOGIA") & ""
            banco = rs.Fields("BANCO") & ""
            cuenta = rs.Fields("CUENTA") & ""
            OBSERVACIONES = rs.Fields("OBSERVACIONES") & ""
            WEB = rs.Fields("WEB") & ""
            If IsNull(rs.Fields("ANULADO")) Then
               anulado = 0
            Else
               anulado = rs.Fields("ANULADO")
            End If
            Centro = rs.Fields("CENTRO") & ""
            SECCION = rs.Fields("SECCION") & ""
            FACTURA_DETERMINACIONES = rs("FACTURA_DETERMINACIONES")
            EADS = rs("EADS")
            AIRBUS = rs("AIRBUS")
            IBERIA = rs("IBERIA")
            AGROALIMENTARIO = rs("AGROALIMENTARIO")
            EXTRANJERO = rs("EXTRANJERO")
            INTRA = rs("INTRA")
            IDIOMA_FACTURA = rs("IDIOMA_FACTURA")
            If IsNull(rs.Fields("MUNICIPIO_ID")) Then
               MUNICIPIO_ID = 0
            Else
               MUNICIPIO_ID = rs.Fields("MUNICIPIO_ID")
            End If
            If IsNull(rs.Fields("PROVINCIA_ID")) Then
               PROVINCIA_ID = 0
            Else
               PROVINCIA_ID = rs.Fields("PROVINCIA_ID")
            End If
            If IsNull(rs.Fields("PAIS_ID")) Then
               PAIS_ID = 0
            Else
               PAIS_ID = rs.Fields("PAIS_ID")
            End If
            FP_ID = rs.Fields("FP_ID")
            TARIFA_ID = rs("TARIFA_ID")
            CLAVEWEB = rs("CLAVEWEB")
            CC = rs("CC")
            ' ENVIOS
            ENVIO_DIRECCION = rs.Fields("ENVIO_DIRECCION") & ""
            If Not IsNull(rs.Fields("ENVIO_COD_POSTAL")) Then
                ENVIO_COD_POSTAL = rs.Fields("ENVIO_COD_POSTAL")
            Else
                ENVIO_COD_POSTAL = 0
            End If
            If IsNull(rs.Fields("ENVIO_MUNICIPIO_ID")) Then
               ENVIO_MUNICIPIO_ID = 0
            Else
               ENVIO_MUNICIPIO_ID = rs.Fields("ENVIO_MUNICIPIO_ID")
            End If
            If IsNull(rs.Fields("ENVIO_PROVINCIA_ID")) Then
               ENVIO_PROVINCIA_ID = 0
            Else
               ENVIO_PROVINCIA_ID = rs.Fields("ENVIO_PROVINCIA_ID")
            End If
            If IsNull(rs.Fields("ENVIO_PAIS_ID")) Then
               ENVIO_PAIS_ID = 0
            Else
               ENVIO_PAIS_ID = rs.Fields("ENVIO_PAIS_ID")
            End If
            ' INFORMES
            INFORMES_DIRECCION = rs.Fields("INFORMES_DIRECCION") & ""
            If Not IsNull(rs.Fields("INFORMES_COD_POSTAL")) Then
                INFORMES_COD_POSTAL = rs.Fields("INFORMES_COD_POSTAL")
            Else
                INFORMES_COD_POSTAL = 0
            End If
            If IsNull(rs.Fields("INFORMES_MUNICIPIO_ID")) Then
               INFORMES_MUNICIPIO_ID = 0
            Else
               INFORMES_MUNICIPIO_ID = rs.Fields("INFORMES_MUNICIPIO_ID")
            End If
            If IsNull(rs.Fields("INFORMES_PROVINCIA_ID")) Then
               INFORMES_PROVINCIA_ID = 0
            Else
               INFORMES_PROVINCIA_ID = rs.Fields("INFORMES_PROVINCIA_ID")
            End If
            If IsNull(rs.Fields("INFORMES_PAIS_ID")) Then
               INFORMES_PAIS_ID = 0
            Else
               INFORMES_PAIS_ID = rs.Fields("INFORMES_PAIS_ID")
            End If
            'M1275-I
            If IsNull(rs.Fields("FACTURA_ELECTRONICA")) Then
               FACTURA_ELECTRONICA = 0
            Else
                FACTURA_ELECTRONICA = rs.Fields("FACTURA_ELECTRONICA")
            End If
            'M1275-F
            If IsNull(rs.Fields("PLANT_ID")) Then
               PLANT_ID = 0
            Else
                PLANT_ID = rs.Fields("PLANT_ID")
            End If
            'E0107-I
            'E0200-I
'            ES_SUBCONTRATA = rs("ES_SUBCONTRATA")
            'E0200-F
            'E0107-F
            CargaCliente = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        CargaCliente = False
        MsgBox "Error al cargar los datos del cliente (CargaCliente)", vbCritical, Err.Description
End Function
Public Function duplicado() As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    If id_cliente <> 0 Then
        consulta = "SELECT ID_CLIENTE FROM CLIENTES WHERE CIF = '" & CIF & "' AND ANULADO = 0 AND ID_CLIENTE <> " & id_cliente
    Else
        consulta = "SELECT ID_CLIENTE FROM CLIENTES WHERE CIF = '" & CIF & "' AND ANULADO = 0"
    End If
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        duplicado = False
    Else
        MsgBox "El CIF ya existe. Cliente número : " & rs(0), vbExclamation, App.Title
        duplicado = True
    End If
    Exit Function
fallo:
    duplicado = False
End Function

Public Function insertar_cliente() As Long
    Dim consulta As String
    On Error GoTo fallo:
'M1278-i
'    If duplicado = True Then
'        Exit Function
'    End If
'M1278-F
    Me.CrearId_Cliente
    '**************** TABLA MUESTRAS *********************
    'E0108-I
    'E0200-I
    consulta = "INSERT INTO CLIENTES(ID_CLIENTE,NOMBRE,DIRECCION," & _
               "COD_POSTAL,CIF,TELEFONO," & _
               "FAX,RESPONSABLE,RESPONSABLE_OTROS,RESPONSABLE_METROLOGIA,CARGO," & _
               "TIPO,EMAIL,EMAIL2,EMAIL_FACTURACION,EMAIL_METROLOGIA,BANCO," & _
               "CUENTA,OBSERVACIONES,WEB," & _
               "ANULADO,CENTRO,SECCION,FACTURA_DETERMINACIONES,EADS,AIRBUS,IBERIA,AGROALIMENTARIO,EXTRANJERO,INTRA,IDIOMA_FACTURA," & _
               "MUNICIPIO_ID,PROVINCIA_ID,PAIS_ID,FP_ID,TARIFA_ID,CLAVEWEB,CC," & _
               "ENVIO_DIRECCION,ENVIO_COD_POSTAL,ENVIO_PAIS_ID,ENVIO_PROVINCIA_ID,ENVIO_MUNICIPIO_ID," & _
               "INFORMES_DIRECCION,INFORMES_COD_POSTAL,INFORMES_PAIS_ID,INFORMES_PROVINCIA_ID,INFORMES_MUNICIPIO_ID, FACTURA_ELECTRONICA, PLANT_ID" & _
               ")" & _
               "values(" & _
               id_cliente & ",'" & nombre & "','" & DIRECCION & "'," & _
               COD_POSTAL & ",'" & CIF & "','" & TELEFONO & "','" & _
               FAX & "','" & responsable & "','" & responsable_OTROS & "','" & RESPONSABLE_METROLOGIA & "','" & CARGO & "','" & _
               tipo & "','" & EMAIL & "','" & EMAIL2 & "','" & EMAIL_FACTURACION & "','" & EMAIL_METROLOGIA & "','" & banco & "','" & _
               cuenta & "','" & OBSERVACIONES & "','" & WEB & "'," & _
               anulado & ",'" & Centro & "','" & SECCION & "'," & _
               FACTURA_DETERMINACIONES & "," & EADS & "," & AIRBUS & "," & IBERIA & "," & AGROALIMENTARIO & "," & EXTRANJERO & "," & INTRA & "," & IDIOMA_FACTURA & "," & _
               MUNICIPIO_ID & "," & PROVINCIA_ID & "," & PAIS_ID & "," & FP_ID & "," & TARIFA_ID & ",'" & CLAVEWEB & "','" & CC & "'," & _
               "'" & DIRECCION & "'," & COD_POSTAL & "," & PAIS_ID & "," & PROVINCIA_ID & "," & MUNICIPIO_ID & "," & _
               "'" & DIRECCION & "'," & COD_POSTAL & "," & PAIS_ID & "," & PROVINCIA_ID & "," & MUNICIPIO_ID & "," & FACTURA_ELECTRONICA & "," & PLANT_ID & _
               ")"
    'E0200-F
    'E0108-F
    execute_bd consulta
    insertar_cliente = id_cliente
    MsgBox "Cliente almacenado correctamente. Código : " & id_cliente, vbInformation, App.Title
    Exit Function
fallo:
    insertar_cliente = 0
    MsgBox "Error al guardar el Cliente en la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Sub CrearId_Cliente()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_CLIENTE) FROM clientes"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        id_cliente = 1
    Else
        id_cliente = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function modificar_cliente() As Boolean
    Dim consulta As String
    On Error GoTo fallo:
'    If duplicado = True Then
'        Exit Function
'    End If
    'E0109-I
    'E0200-I
    consulta = "UPDATE CLIENTES SET " & _
               "NOMBRE = '" & nombre & "'" & _
               ",DIRECCION = '" & DIRECCION & "',COD_POSTAL = " & COD_POSTAL & ",CIF = '" & CIF & "',TELEFONO = '" & TELEFONO & "'" & _
               ",FAX = '" & FAX & "',RESPONSABLE = '" & responsable & "'" & _
               ",RESPONSABLE_OTROS = '" & responsable_OTROS & "'" & _
               ",RESPONSABLE_METROLOGIA = '" & RESPONSABLE_METROLOGIA & "'" & _
               ",CARGO = '" & CARGO & "',TIPO = '" & tipo & "'" & ",EMAIL = '" & EMAIL & "',BANCO = '" & banco & "'" & _
               ",CUENTA = '" & cuenta & "',OBSERVACIONES = '" & OBSERVACIONES & "'" & _
               ",WEB = '" & WEB & "',ANULADO = " & anulado & _
               ",CENTRO = '" & Centro & "',SECCION = '" & SECCION & "'" & _
               ",FACTURA_DETERMINACIONES = " & FACTURA_DETERMINACIONES & _
               ",EADS = " & EADS & ",AIRBUS=" & AIRBUS & ",IBERIA=" & IBERIA & ",AGROALIMENTARIO=" & AGROALIMENTARIO & ",EXTRANJERO=" & EXTRANJERO & ",EMAIL2='" & EMAIL2 & "'" & _
               ",FP_ID = " & FP_ID & ",INTRA = " & INTRA & _
               ",IDIOMA_FACTURA = " & IDIOMA_FACTURA & ",TARIFA_ID = " & TARIFA_ID & _
               ",MUNICIPIO_ID = " & MUNICIPIO_ID & ",PROVINCIA_ID = " & PROVINCIA_ID & _
               ",PAIS_ID = " & PAIS_ID & ",CC = '" & CC & "'" & _
               ",CLAVEWEB = '" & CLAVEWEB & "'" & _
               ", FACTURA_ELECTRONICA = " & FACTURA_ELECTRONICA & _
               ", PLANT_ID = " & PLANT_ID & _
               ", EMAIL_FACTURACION = '" & EMAIL_FACTURACION & "'" & _
               ", EMAIL_METROLOGIA = '" & EMAIL_METROLOGIA & "'" & _
               " WHERE " & _
               " id_cliente = " & id_cliente
    'E0200-F
    'E0109-F
               execute_bd consulta
    modificar_cliente = True
    MsgBox "Cliente modificado correctamente. Código : " & id_cliente, vbInformation, App.Title
    Exit Function
fallo:
    modificar_cliente = False
    MsgBox "Error al modificar el Cliente en la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Function eliminar_cliente() As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "UPDATE CLIENTES SET ANULADO=1 " & _
               " WHERE " & _
               " id_cliente = " & id_cliente
               execute_bd consulta
    eliminar_cliente = True
    MsgBox "Cliente eliminado correctamente. Código : " & id_cliente, vbInformation, App.Title
    Exit Function
fallo:
    eliminar_cliente = False
    MsgBox "Error al eliminar el Cliente de la bd : " & Err.Description, vbCritical, Err.Number
End Function

'E0113-I
'E0200-I
'Public Function listado_subcontratas() As ADODB.RecordSet
'    Dim consulta As String
'
'    consulta = "SELECT ID_CLIENTE, NOMBRE " & _
'               "  FROM CLIENTES " & _
'               " WHERE ANULADO = 0 " & _
'               "   AND ES_SUBCONTRATA = 1 " & _
'               " ORDER BY NOMBRE "
'    Set listado_subcontratas = datos_bd(consulta)
'End Function
'E0200-F

' Filtro para mostrar sólo los clientes que también son contratas
'E0200-I
'Public Function Listado(p1 As String, p2 As String, p3 As String, Optional strSoloContratas As String = "") As ADODB.RecordSet
'E0200-F
'M1275: Factura electrónica
Public Function Listado(P1 As String, P2 As String, P3 As String, Optional EXTRANJERO As Boolean, Optional INTRA As Boolean, Optional AIRBUS As Boolean, Optional IBERIA As Boolean, Optional AGROALIMENTARIO As Boolean, Optional RESPONSABLE_ID As Long, Optional FACTURAELECTRONICA As Boolean) As ADODB.Recordset
    Dim consulta As String
    Dim filtro As String
    If EXTRANJERO = True Then
        filtro = filtro & " AND EXTRANJERO = 1 "
    End If
    If INTRA = True Then
        filtro = filtro & " AND INTRA = 1 "
    End If
    If AIRBUS = True Then
        filtro = filtro & " AND AIRBUS = 1 "
    End If
    If IBERIA = True Then
        filtro = filtro & " AND IBERIA = 1 "
    End If
    If AGROALIMENTARIO = True Then
        filtro = filtro & " AND AGROALIMENTARIO = 1 "
    End If
    'M1275-I
    If FACTURAELECTRONICA = True Then
        filtro = filtro & " AND FACTURA_ELECTRONICA = 0"
    End If
    'M1275-F
    
    If RESPONSABLE_ID <> 0 Then
        filtro = filtro & " AND cr.responsable_id = " & RESPONSABLE_ID
    End If
    
    consulta = "SELECT DISTINCT c.* FROM clientes c left join clientes_responsables cr on c.id_cliente = cr.cliente_id " & _
               " WHERE anulado=0 and " & _
               " nombre like '%" & Trim(P1) & "%' and " & _
               " cif like '%" & Trim(P2) & "%' and " & _
               " telefono like '%" & Trim(P3) & "%'" & _
               filtro & _
               " order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Sub Imprimir_Listado(ByVal prm_nombre As String, ByVal prm_cif As String, ByVal prm_telefono As String, booEads As Boolean, booAirbus As Boolean, booIberia As Boolean, booExtranjero As Boolean, booAgroalimentario As Boolean, booSinFacturaElectronica As Boolean, responsable As Long)

    Dim objRpt As New frmReport
    Dim arrNom() As String
    Dim arrVal() As String
    Dim strCriterio As String


    If prm_nombre <> "" Then strCriterio = strCriterio & " AND {clientes.NOMBRE} like " & Chr(34) & "*" & prm_nombre & "*" & Chr(34)
    If prm_cif <> "" Then strCriterio = strCriterio & " AND {clientes.CIF} like " & Chr(34) & "*" & prm_cif & "*" & Chr(34)
    If prm_telefono <> "" Then strCriterio = strCriterio & " AND {clientes.TELEFONO} like " & Chr(34) & "*" & prm_telefono & "*" & Chr(34)
    If booEads Then
        strCriterio = strCriterio & " AND {clientes.EADS} = 1 "
    End If
    If booAirbus Then
        strCriterio = strCriterio & " AND {clientes.AIRBUS} = 1 "
    End If
    If booIberia Then
        strCriterio = strCriterio & " AND {clientes.IBERIA} = 1 "
    End If
    If booExtranjero Then
        strCriterio = strCriterio & " AND {clientes.EXTRANJERO} = 1 "
    End If
    If booAgroalimentario Then
        strCriterio = strCriterio & " AND {clientes.AGROALIMENTARIO} = 1 "
    End If
    If booSinFacturaElectronica Then
        strCriterio = strCriterio & " AND {clientes.FACTURA_ELECTRONICA} = 0 "
    End If
    If responsable <> 0 Then
        strCriterio = strCriterio & " AND {clientes_responsables.RESPONSABLE_ID} = " & responsable
    End If
    

    strCriterio = Mid(strCriterio, 5)

    ReDim arrNom(1)
    ReDim arrVal(1)

    arrNom(1) = "impresor"
    arrVal(1) = USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS

    With objRpt
        .iniciar
        .informe = "General\rptListadoClientes"
        .criterio = strCriterio
        .imprimir = False

        .ParametrosNombre = arrNom
        .ParametrosValores = arrVal

        .generar
        '.Visible = True
        .Show 1
    End With
    Unload objRpt
    Set objRpt = Nothing

End Sub


Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_CLIENTE, NOMBRE FROM clientes " & _
               " WHERE anulado=0 " & _
               " order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Clientes_muestras_especiales() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT distinct c.* FROM clientes c,banos b " & _
               " WHERE c.id_cliente = b.cliente_id " & _
               "   and c.anulado = 0 " & _
               " order by c.nombre"
    Set Clientes_muestras_especiales = datos_bd(consulta)
End Function

Public Function modificar_direcciones(ID As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "UPDATE CLIENTES SET " & _
               " ENVIO_DIRECCION = '" & ENVIO_DIRECCION & "'" & _
               ",ENVIO_COD_POSTAL = " & ENVIO_COD_POSTAL & _
               ",ENVIO_PAIS_ID = " & ENVIO_PAIS_ID & _
               ",ENVIO_PROVINCIA_ID = " & ENVIO_PROVINCIA_ID & _
               ",ENVIO_MUNICIPIO_ID = " & ENVIO_MUNICIPIO_ID & _
               ",INFORMES_DIRECCION = '" & INFORMES_DIRECCION & "'" & _
               ",INFORMES_COD_POSTAL = " & INFORMES_COD_POSTAL & _
               ",INFORMES_PAIS_ID = " & INFORMES_PAIS_ID & _
               ",INFORMES_PROVINCIA_ID = " & INFORMES_PROVINCIA_ID & _
               ",INFORMES_MUNICIPIO_ID = " & INFORMES_MUNICIPIO_ID & _
               " WHERE " & _
               " id_cliente = " & ID
    execute_bd consulta
    modificar_direcciones = True
    Exit Function
fallo:
    modificar_direcciones = False
    MsgBox "Error al modificar el Cliente en la bd : " & Err.Description, vbCritical, Err.Number
End Function

Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    Dim f As String
    If filtro = "" Then
        f = "ANULADO=0"
    Else
        f = filtro
    End If
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "CLIENTES"
        .setDESCRIPCION = "Clientes"
        .setPK = "ID_CLIENTE"
'        .setFILTRO = filtro
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = True
'        .setFILTRO = filtro
        .setFILTRO = f
        Set .FORMULARIO = FORMULARIO
    End With
End Function
Public Function llenar_combo_Banos(combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    Dim consulta As String
    consulta = "select DISTINCT a.ID_CLIENTE,a.NOMBRE from clientes a, banos b where a.ID_CLIENTE = b.CLIENTE_ID "
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        With combo
            .setCONN = conn
            .setFK_CAMPO = ""
            .setFK_VALOR = FK
            .setTABLA = "CLIENTES"
            .setDESCRIPCION = "Clientes"
            .setPK = "a.ID_CLIENTE"
            .setFILTRO = filtro
            .setCAMPO = "a.NOMBRE"
            .setQUERY = consulta
            .setMUESTRA_DETALLE = True
            .setFILTRO = ""
            Set .FORMULARIO = FORMULARIO
        End With
    End If
End Function


Public Function Recupera_Tarifa(id_cliente As Long) As Integer
        On Error GoTo fallo
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT TARIFA_ID FROM clientes WHERE ID_CLIENTE=" & id_cliente
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Recupera_Tarifa = 0
        Else
            Recupera_Tarifa = rs(0)
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Recupera_Tarifa = 0
End Function
Public Function Modificar_Tarifa(id_cliente As Long, TARIFA As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE clientes SET TARIFA_ID = " & TARIFA & " WHERE ID_CLIENTE=" & id_cliente
        execute_bd consulta
        Modificar_Tarifa = True
        Exit Function
fallo:
        Modificar_Tarifa = False

End Function
'M1273-I
Public Function duplicarCliente(id_cliente As Long) As Long
On Error GoTo fallo
        Me.CargaCliente id_cliente
        Me.setNOMBRE = Me.getNOMBRE & " (DUPLICADO)"
        Dim cliente As Long
        cliente = Me.insertar_cliente
        If cliente > 0 Then
            Dim oCR As New clsClientes_responsables
            oCR.duplicar id_cliente, cliente
            Set oCR = Nothing
        End If
        duplicarCliente = cliente
        Exit Function
fallo:
        duplicarCliente = 0
End Function
'M1273-F
