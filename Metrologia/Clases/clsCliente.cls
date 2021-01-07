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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'*****************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLIENTE      *
'*****************************************************
Private id_cliente As Long
Private NOMBRE As String
Private DIRECCION As String
Private CP As String
Private PROVINCIA_ID As Long
Private MUNICIPIO_ID As Long
Private CIF As String
Private Telefono As String
Private movil As String
Private FAX As String
Private EMAIL As String
Private OBSERVACIONES As String
Private RAZON As String
Private FORMA_PAGO As Integer
Private anulado As Integer
Private iva As Integer
Private COPIAS_FACTURA As Integer
Private RIESGO As String
Private RIESGO_REAL As String
Private CCONTABLE As String

'*****************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLIENTES   *
'*****************************************************
Public Property Let setID_CLIENTE(ByVal dato As Long)
    id_cliente = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Let setDIRECCION(ByVal dato As String)
    DIRECCION = Trim(dato)
End Property
Public Property Let setPROVINCIA_ID(ByVal dato As Long)
    PROVINCIA_ID = dato
End Property
Public Property Let setCP(ByVal dato As String)
    CP = dato
End Property
Public Property Let setCIF(ByVal dato As String)
    CIF = Trim(dato)
End Property
Public Property Let setTelefono(ByVal dato As String)
    Telefono = Trim(dato)
End Property
Public Property Let setMOVIL(ByVal dato As String)
    movil = Trim(dato)
End Property
Public Property Let setFAX(ByVal dato As String)
    FAX = Trim(dato)
End Property
Public Property Let setEMAIL(ByVal dato As String)
    EMAIL = Trim(dato)
End Property
Public Property Let setMUNICIPIO_ID(ByVal dato As Long)
    MUNICIPIO_ID = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
    OBSERVACIONES = Trim(dato)
End Property
Public Property Let setRAZON(ByVal dato As String)
    RAZON = Trim(dato)
End Property
Public Property Let setFORMA_PAGO(ByVal dato As Integer)
    FORMA_PAGO = dato
End Property
Public Property Let setANULADO(ByVal dato As Integer)
    anulado = Trim(dato)
End Property
Public Property Let setIVA(ByVal dato As Integer)
    iva = Trim(dato)
End Property
Public Property Let setCOPIAS_FACTURA(ByVal dato As Integer)
    COPIAS_FACTURA = Trim(dato)
End Property
Public Property Let setRIESGO(ByVal dato As String)
    RIESGO = Trim(dato)
End Property
Public Property Let setRIESGO_REAL(ByVal dato As String)
    RIESGO_REAL = Trim(dato)
End Property
Public Property Let setCCONTABLE(ByVal dato As String)
    CCONTABLE = Trim(dato)
End Property
'*****************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLIENTES   *
'*****************************************************
Public Property Get getID_CLIENTE() As Long
    getID_CLIENTE = id_cliente
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property
Public Property Get getDIRECCION() As String
    getDIRECCION = DIRECCION
End Property
Public Property Get getPROVINCIA_ID() As Long
    getPROVINCIA_ID = PROVINCIA_ID
End Property
Public Property Get getCP() As String
    getCP = CP
End Property
Public Property Get getCIF() As String
    getCIF = CIF
End Property
Public Property Get getTelefono() As String
    getTelefono = Telefono
End Property
Public Property Get getMOVIL() As String
    getMOVIL = movil
End Property
Public Property Get getFAX() As String
    getFAX = FAX
End Property
Public Property Get getEMAIL() As String
    getEMAIL = EMAIL
End Property
Public Property Get getMUNICIPIO_ID() As Long
    getMUNICIPIO_ID = MUNICIPIO_ID
End Property
Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getRAZON() As String
    getRAZON = RAZON
End Property
Public Property Get getFORMA_PAGO() As Integer
    getFORMA_PAGO = FORMA_PAGO
End Property
Public Property Get getANULADO() As Integer
    getANULADO = anulado
End Property
Public Property Get getIVA() As Integer
    getIVA = iva
End Property
Public Property Get getCOPIAS_FACTURA() As Integer
    getCOPIAS_FACTURA = COPIAS_FACTURA
End Property
Public Property Get getCCONTABLE() As String
    getCCONTABLE = CCONTABLE
End Property
Public Property Get getRIESGO() As String
    getRIESGO = RIESGO
End Property
Public Property Get getRIESGO_REAL() As String
    getRIESGO_REAL = RIESGO_REAL
End Property

Public Function CargaCliente(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM clientes WHERE ID_CLIENTE=" & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            CargaCliente = False
        Else
            id_cliente = rs.Fields("ID_CLIENTE")
            NOMBRE = rs.Fields("NOMBRE")
            DIRECCION = rs.Fields("DIRECCION")
            PROVINCIA_ID = rs.Fields("PROVINCIA_ID")
'            If Not IsNull(rs.Fields("CP")) Then
                CP = rs.Fields("CP")
'            Else
'                CP = 0
'            End If
            CIF = rs.Fields("CIF")
            Telefono = rs.Fields("TELEFONO")
            movil = rs.Fields("MOVIL")
            FAX = rs.Fields("FAX")
            EMAIL = rs.Fields("EMAIL")
            MUNICIPIO_ID = rs.Fields("MUNICIPIO_ID")
            OBSERVACIONES = rs.Fields("OBSERVACIONES")
            RAZON = rs.Fields("RAZON")
            FORMA_PAGO = rs.Fields("FORMA_PAGO")
            anulado = rs("ANULADO")
            iva = rs("IVA")
            COPIAS_FACTURA = rs("COPIAS_FACTURA")
            CCONTABLE = rs("CCONTABLE")
            RIESGO = rs("RIESGO")
            RIESGO_REAL = rs("RIESGO_REAL")
            CargaCliente = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        CargaCliente = False
        MsgBox "Error al cargar los datos del cliente (CargaCliente)", vbCritical, Err.Description
End Function

Public Function insertar_cliente() As Long
    Dim consulta As String
    On Error GoTo fallo:
    If id_cliente = 0 Then
        Me.CrearId_Cliente
    End If
    consulta = "INSERT INTO CLIENTES" & _
               " VALUES(" & _
               id_cliente & ",'" & NOMBRE & "','" & DIRECCION & "','" & _
               CP & "'," & PROVINCIA_ID & "," & MUNICIPIO_ID & ",'" & _
               CIF & "','" & Telefono & "','" & movil & "','" & FAX & "','" & EMAIL & "','" & _
               OBSERVACIONES & "','" & RAZON & "'," & FORMA_PAGO & "," & _
               anulado & "," & _
               iva & "," & COPIAS_FACTURA & ",'" & RIESGO & "','" & RIESGO_REAL & "','" & CCONTABLE & "')"
    execute_bd consulta
    If NOMBRE <> "" Then
        Dim oagenda As New clsAgenda
        With oagenda
            .setNOMBRE = NOMBRE
            .setTelefono = Telefono
            .setMOVIL = movil
            .setFAX = FAX
            .setCLIENTE_ID = id_cliente
            .Insertar
        End With
    End If
'    Dim oCliente_FP As New clsCLIENTES_FP
'    With oCliente_FP
'        .setCLIENTE_ID = id_cliente
'        .setFORMA_PAGO_ID = 0
'        .setORDEN = 1
'        .Insertar
'    End With
'    Dim oCliente_Direccion As New clsCLIENTES_DIRECCIONES
'    With oCliente_Direccion
'        .setCLIENTE_ID = id_cliente
'        .setID_DIRECCION = 0
'        .setNOMBRE = "Fiscal"
'        .setDIRECCION = DIRECCION
'        .setCP = CP
'        .setMUNICIPIO_ID = MUNICIPIO_ID
'        .setPROVINCIA_ID = PROVINCIA_ID
'        .Insertar
'    End With
    insertar_cliente = id_cliente
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
Public Function BuscarId_ClienteLibre(Optional anterior As Long) As Long
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    Dim c As String
    Dim numero As Long
    If anterior > 0 Then
        numero = anterior + 1
        c = " WHERE ID_CLIENTE >= " & anterior
    Else
        numero = 4300001
    End If
    consulta = "SELECT ID_CLIENTE " & _
               "  FROM clientes " & _
               c & _
               " ORDER BY ID_CLIENTE"
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        Do
            If rs(0) > numero Then
                BuscarId_ClienteLibre = numero
                Exit Do
            End If
            numero = numero + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
'    BuscarId_ClienteLibre = numero
    Set rs = Nothing
End Function

Public Function Verificar_Codigo(ID As Long) As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_CLIENTE FROM clientes where ID_CLIENTE = " & ID
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        Verificar_Codigo = False
    Else
        Verificar_Codigo = True
    End If
    Set rs = Nothing
End Function

Public Function modificar_cliente() As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "UPDATE CLIENTES SET" & _
               " NOMBRE = '" & NOMBRE & "'," & _
               " DIRECCION = '" & DIRECCION & "'," & _
               " CP = '" & CP & "'," & _
               " PROVINCIA_ID = " & PROVINCIA_ID & "," & " MUNICIPIO_ID = " & MUNICIPIO_ID & "," & _
               " CIF = '" & CIF & "'," & _
               " TELEFONO = '" & Telefono & "'," & _
               " MOVIL = '" & movil & "'," & _
               " FAX = '" & FAX & "'," & _
               " EMAIL = '" & EMAIL & "'," & _
               " RAZON = '" & RAZON & "'," & _
               " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
               " FORMA_PAGO = " & FORMA_PAGO & "," & _
               " IVA = " & iva & "," & _
               " COPIAS_FACTURA = " & COPIAS_FACTURA & "," & _
               " RIESGO = '" & RIESGO & "',RIESGO_REAL='" & RIESGO_REAL & "'," & _
               " CCONTABLE = '" & CCONTABLE & "'" & _
               " WHERE " & _
               " id_cliente = " & id_cliente
    execute_bd consulta
'   LP001 Modificamos la tabla cliente->direcciones
'    Dim oCliente_Direccion As New clsCLIENTES_DIRECCIONES
'    With oCliente_Direccion
'        .setNOMBRE = "Fiscal"
'        .setDIRECCION = DIRECCION
'        .setCP = CP
'        .setMUNICIPIO_ID = MUNICIPIO_ID
'        .setPROVINCIA_ID = PROVINCIA_ID
'        .setTELEFONO = telefono
'        .Modificar id_cliente, 0
'    End With
    modificar_cliente = True
    Exit Function
fallo:
    modificar_cliente = False
    MsgBox "Error al modificar el Cliente en la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Function anular_cliente() As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "UPDATE CLIENTES SET ANULADO = 1 WHERE ID_CLIENTE = " & id_cliente
    execute_bd consulta
    anular_cliente = True
    Exit Function
fallo:
    anular_cliente = False
    MsgBox "Error al anular el Cliente de la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Function restablecer_cliente() As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "UPDATE CLIENTES SET ANULADO = 0 WHERE ID_CLIENTE = " & id_cliente
    execute_bd consulta
    restablecer_cliente = True
    Exit Function
fallo:
    restablecer_cliente = False
    MsgBox "Error al restablecer el Cliente de la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Function eliminar_cliente() As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    eliminar_cliente = False
    consulta = "SELECT * FROM OBRAS WHERE CLIENTE_ID = " & id_cliente
    Dim rs As ADODB.Recordset
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        MsgBox "No se puede eliminar el cliente. Tiene obras relacionados.", vbExclamation, App.Title
        Exit Function
    End If
    consulta = "DELETE FROM CLIENTES_DIRECCIONES WHERE CLIENTE_ID = " & id_cliente
    execute_bd consulta
    consulta = "DELETE FROM CLIENTES_FP WHERE CLIENTE_ID = " & id_cliente
    execute_bd consulta
    consulta = "DELETE FROM CLIENTES WHERE ID_CLIENTE = " & id_cliente
    execute_bd consulta
    eliminar_cliente = True
    Exit Function
fallo:
    eliminar_cliente = False
    MsgBox "Error al eliminar el Cliente de la bd : " & Err.Description, vbCritical, Err.Number
End Function

Public Function Listado(lNOMBRE As String, lNIF As String, lTELEFONO As String, lPOBLACION As String, RECUPERAR_ANULADOS As Boolean) As ADODB.Recordset
    Dim consulta As String
    Dim s As String
    If Trim(lNOMBRE) <> "" Then
        s = s & " AND A.NOMBRE LIKE '%" & lNOMBRE & "%'"
    End If
    If Trim(lNIF) <> "" Then
        s = s & " AND A.CIF LIKE '%" & lNIF & "%'"
    End If
    If Trim(lTELEFONO) <> "" Then
        s = s & " AND A.TELEFONO LIKE '%" & lTELEFONO & "%'"
    End If
    If Trim(lPOBLACION) <> "" Then
        s = s & " AND C.NOMBRE LIKE '%" & lPOBLACION & "%'"
    End If
    If RECUPERAR_ANULADOS = True Then
        s = s & " AND A.ANULADO = 1 "
    Else
        s = s & " AND A.ANULADO = 0 "
    End If
    consulta = "SELECT A.ID_CLIENTE,A.NOMBRE,A.DIRECCION,B.NOMBRE,C.NOMBRE,A.ANULADO " & _
               "  FROM CLIENTES A " & _
               "  LEFT JOIN PROVINCIAS B ON A.PROVINCIA_ID = B.ID_PROVINCIA " & _
               "  LEFT JOIN MUNICIPIOS C ON A.MUNICIPIO_ID = C.ID_MUNICIPIO " & _
               " WHERE 1=1 " & _
               s & _
               " ORDER BY A.NOMBRE"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_CLIENTE, NOMBRE FROM clientes " & _
               " WHERE ANULADO = 0 " & _
               " ORDER BY NOMBRE"
    Set Listado_Combo = datos_bd(consulta)
End Function
Public Function modificar_direccion(ID As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "UPDATE CLIENTES SET" & _
               " DIRECCION = '" & DIRECCION & "'," & _
               " CP = '" & CP & "'," & _
               " PROVINCIA_ID = " & PROVINCIA_ID & "," & _
               " TELEFONO = '" & Telefono & "'," & _
               " MUNICIPIO_ID = " & MUNICIPIO_ID & _
               " WHERE " & _
               " id_cliente = " & ID
    execute_bd consulta
    modificar_direccion = True
    Exit Function
fallo:
    modificar_direccion = False
    MsgBox "Error al modificar el Cliente en la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Function llenar_combo(combo As miComboBCA, FK As Long, FORMULARIO As Form, FILTRO As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "CLIENTES"
        .setDESCRIPCION = "Clientes"
        .setPK = "ID_CLIENTE"
        .setCAMPO = "CONCAT(cast(ID_CLIENTE as char),' - ',NOMBRE)"
        .setMUESTRA_DETALLE = True
        .setFILTRO = FILTRO
        Set .FORMULARIO = FORMULARIO
    End With
End Function

