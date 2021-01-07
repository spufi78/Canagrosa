VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*****************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLIENTE      *
'*****************************************************
Private ID_EMPLEADO As Long
Private USUARIO_ID As Long
Private nombre As String
Private APODO As String
Private DIRECCION As String
Private cp As Long
Private PROVINCIA_ID As Long
Private MUNICIPIO_ID As Long
Private CIF As String
Private TELEFONO As String
Private movil As String
Private foto As String
Private OBSERVACIONES As String
Private ESTADO_ID As Integer
Private CODIGO_INTERNO As String
Private ccc As String
Private FECHA_NACIMIENTO As String
Private FECHA_INCORPORACION As String
Private FECHA_BAJA As String
Private DEPARTAMENTOS As String
Private PNTS As String
'M0521-I
Private external As Integer
'M0521-F
'MANTIS-XXX-I
Private empresa_id As Integer
'MANTIS-XXX-F
'*****************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE MUESTRA    *
'*****************************************************
Public Property Let setID_EMPLEADO(ByVal dato As Long)
    ID_EMPLEADO = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
    USUARIO_ID = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Let setAPODO(ByVal dato As String)
    APODO = Trim(dato)
End Property
Public Property Let setDIRECCION(ByVal dato As String)
    DIRECCION = Trim(dato)
End Property
Public Property Let setPROVINCIA_ID(ByVal dato As Long)
    PROVINCIA_ID = dato
End Property
Public Property Let setCP(ByVal dato As Long)
    cp = dato
End Property
Public Property Let setCIF(ByVal dato As String)
    CIF = Trim(dato)
End Property
Public Property Let setTELEFONO(ByVal dato As String)
    TELEFONO = Trim(dato)
End Property
Public Property Let setMOVIL(ByVal dato As String)
    movil = Trim(dato)
End Property
Public Property Let setMUNICIPIO_ID(ByVal dato As Long)
    MUNICIPIO_ID = dato
End Property
Public Property Let setfoto(ByVal dato As String)
    foto = Trim(dato)
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
    OBSERVACIONES = Trim(dato)
End Property
Public Property Let setESTADO_ID(ByVal dato As Integer)
    ESTADO_ID = Trim(dato)
End Property
Public Property Let setCODIGO_INTERNO(ByVal dato As String)
    CODIGO_INTERNO = Trim(dato)
End Property
Public Property Let setCCC(ByVal dato As String)
    ccc = Trim(dato)
End Property
Public Property Let setFECHA_NACIMIENTO(ByVal dato As String)
    FECHA_NACIMIENTO = Trim(dato)
End Property
Public Property Let setFECHA_INCORPORACION(ByVal dato As String)
    FECHA_INCORPORACION = Trim(dato)
End Property
Public Property Let setFECHA_BAJA(ByVal dato As String)
    FECHA_BAJA = Trim(dato)
End Property
Public Property Let setDEPARTAMENTOS(ByVal dato As String)
    DEPARTAMENTOS = Trim(dato)
End Property
Public Property Let setPNTS(ByVal dato As String)
    PNTS = Trim(dato)
End Property
'M0521-I
Public Property Let setEXTERNAL(ByVal dato As Integer)
    external = Trim(dato)
End Property
'M0521-F
'MANTIS-XXX-I
Public Property Let setEMPRESA_ID(ByVal dato As Integer)
    empresa_id = Trim(dato)
End Property
'MANTIS-XXX-F
Public Property Get getID_EMPLEADO() As Long
    getID_EMPLEADO = ID_EMPLEADO
End Property
Public Property Get getUSUARIO_ID() As Long
    getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getAPODO() As String
    getAPODO = APODO
End Property
Public Property Get getDIRECCION() As String
    getDIRECCION = DIRECCION
End Property
Public Property Get getPROVINCIA_ID() As Long
    getPROVINCIA_ID = PROVINCIA_ID
End Property
Public Property Get getCP() As Long
    getCP = cp
End Property
Public Property Get getCIF() As String
    getCIF = CIF
End Property
Public Property Get getTELEFONO() As String
    getTELEFONO = TELEFONO
End Property
Public Property Get getMOVIL() As String
    getMOVIL = movil
End Property
Public Property Get getMUNICIPIO_ID() As Long
    getMUNICIPIO_ID = MUNICIPIO_ID
End Property
Public Property Get getfoto() As String
    getfoto = foto
End Property
Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getESTADO_ID() As String
    getESTADO_ID = ESTADO_ID
End Property
Public Property Get getCODIGO_INTERNO() As String
    getCODIGO_INTERNO = CODIGO_INTERNO
End Property
Public Property Get getCCC() As String
    getCCC = ccc
End Property
Public Property Get getFECHA_NACIMIENTO() As String
    getFECHA_NACIMIENTO = FECHA_NACIMIENTO
End Property
Public Property Get getFECHA_INCORPORACION() As String
    getFECHA_INCORPORACION = FECHA_INCORPORACION
End Property
Public Property Get getFECHA_BAJA() As String
    getFECHA_BAJA = FECHA_BAJA
End Property
Public Property Get getDEPARTAMENTOS() As String
    getDEPARTAMENTOS = DEPARTAMENTOS
End Property
Public Property Get getPNTS() As String
    getPNTS = PNTS
End Property
'M0521-I
Public Property Get getEXTERNAL() As Integer
    getEXTERNAL = external
End Property
'M0521-F
'MANTIS-XXX-I
Public Property Get getEMPRESA_ID() As Integer
    getEMPRESA_ID = empresa_id
End Property
'MANTIS-XXX-F
Public Function CARGAR(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM empleados WHERE ID_EMPLEADO=" & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            CARGAR = False
        Else
            ID_EMPLEADO = rs.Fields("ID_EMPLEADO")
            USUARIO_ID = rs.Fields("USUARIO_ID")
            nombre = rs.Fields("NOMBRE")
            APODO = rs.Fields("APODO")
            DIRECCION = rs.Fields("DIRECCION")
            PROVINCIA_ID = rs.Fields("PROVINCIA_ID")
            If Not IsNull(rs.Fields("CP")) Then
                cp = rs.Fields("CP")
            Else
                cp = 0
            End If
            CIF = rs.Fields("CIF")
            TELEFONO = rs.Fields("TELEFONO")
            movil = rs.Fields("MOVIL")
            MUNICIPIO_ID = rs.Fields("MUNICIPIO_ID")
            foto = rs.Fields("FOTO")
            OBSERVACIONES = rs.Fields("OBSERVACIONES")
            ESTADO_ID = rs.Fields("ESTADO_ID")
            CODIGO_INTERNO = rs.Fields("CODIGO_INTERNO")
            ccc = rs.Fields("CCC")
            FECHA_NACIMIENTO = rs("FECHA_NACIMIENTO")
            FECHA_INCORPORACION = rs("FECHA_INCORPORACION")
            FECHA_BAJA = rs("FECHA_BAJA")
            DEPARTAMENTOS = rs("DEPARTAMENTOS")
            PNTS = rs("PNTS")
            'M0521-I
            external = rs("EXTERNAL")
            'M0521-F
            'MANTIS-XXX-I
            empresa_id = rs("EMPRESA_ID")
            'MANTIS-XXX-F
            
            CARGAR = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        CARGAR = False
        MsgBox "Error al cargar los datos del empleado (Cargar)", vbCritical, Err.Description
End Function
Public Function CARGAR_POR_USUARIO(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM empleados WHERE USUARIO_ID=" & ID
        Set rs = datos_bd(consulta, True)
        If rs.RecordCount = 0 Then
            CARGAR_POR_USUARIO = False
        Else
            ID_EMPLEADO = rs.Fields("ID_EMPLEADO")
            USUARIO_ID = rs.Fields("USUARIO_ID")
            nombre = rs.Fields("NOMBRE")
            APODO = rs.Fields("APODO")
            DIRECCION = rs.Fields("DIRECCION")
            PROVINCIA_ID = rs.Fields("PROVINCIA_ID")
            If Not IsNull(rs.Fields("CP")) Then
                cp = rs.Fields("CP")
            Else
                cp = 0
            End If
            CIF = rs.Fields("CIF")
            TELEFONO = rs.Fields("TELEFONO")
            movil = rs.Fields("MOVIL")
            MUNICIPIO_ID = rs.Fields("MUNICIPIO_ID")
            foto = rs.Fields("FOTO")
            OBSERVACIONES = rs.Fields("OBSERVACIONES")
            ESTADO_ID = rs.Fields("ESTADO_ID")
            CODIGO_INTERNO = rs.Fields("CODIGO_INTERNO")
            ccc = rs.Fields("CCC")
            FECHA_NACIMIENTO = rs("FECHA_NACIMIENTO")
            FECHA_INCORPORACION = rs("FECHA_INCORPORACION")
            FECHA_BAJA = rs("FECHA_BAJA")
            DEPARTAMENTOS = rs("DEPARTAMENTOS")
            PNTS = rs("PNTS")
            'M0521-I
            external = rs("EXTERNAL")
            'M0521-F
            'MANTIS-XXX-I
            empresa_id = rs("EMPRESA_ID")
            'MANTIS-XXX-F
            CARGAR_POR_USUARIO = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        CARGAR_POR_USUARIO = False
        MsgBox "Error al cargar los datos del empleado (Cargar)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long
    Dim consulta As String
    On Error GoTo fallo:
    Me.CrearID_EMPLEADO
    consulta = "INSERT INTO EMPLEADOS" & _
               " VALUES(" & _
               ID_EMPLEADO & "," & USUARIO_ID & ",'" & nombre & "','" & APODO & "','" & DIRECCION & "'," & _
               cp & "," & PROVINCIA_ID & "," & MUNICIPIO_ID & ",'" & _
               CIF & "','" & TELEFONO & "','" & movil & "','" & _
               foto & "','" & OBSERVACIONES & "'," & ESTADO_ID & ",'" & CODIGO_INTERNO & "','" & _
               ccc & "','" & FECHA_INCORPORACION & "','" & FECHA_INCORPORACION & "','" & FECHA_BAJA & "','" & DEPARTAMENTOS & "','" & PNTS & "','" & external & "'," & empresa_id & ")"
    execute_bd consulta
    Insertar = ID_EMPLEADO
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al guardar el empleado en la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Sub CrearID_EMPLEADO()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_EMPLEADO) FROM EMPLEADOS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_EMPLEADO = 1
    Else
        ID_EMPLEADO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Modificar(ID As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "UPDATE empleados SET" & _
               "  USUARIO_ID = " & USUARIO_ID & "," & _
               "  NOMBRE = '" & nombre & "'," & _
               "  APODO = '" & APODO & "'," & _
               "  DIRECCION = '" & DIRECCION & "'," & _
               "  CP = " & cp & "," & _
               "  PROVINCIA_ID = " & PROVINCIA_ID & "," & _
               "  MUNICIPIO_ID = " & MUNICIPIO_ID & "," & _
               "  CIF = '" & CIF & "'," & _
               "  TELEFONO = '" & TELEFONO & "'," & _
               "  MOVIL = '" & movil & "'," & _
               "  FOTO = '" & foto & "'," & _
               "  OBSERVACIONES = '" & OBSERVACIONES & "'," & _
               "  CODIGO_INTERNO = '" & CODIGO_INTERNO & "'," & _
               "  CCC = '" & ccc & "'," & _
               "  FECHA_NACIMIENTO = '" & FECHA_NACIMIENTO & "'," & _
               "  FECHA_INCORPORACION = '" & FECHA_INCORPORACION & "'," & _
               "  FECHA_BAJA = '" & FECHA_BAJA & "'," & _
               "  DEPARTAMENTOS = '" & DEPARTAMENTOS & "'," & _
               "  PNTS = '" & PNTS & "'," & _
               "  EXTERNAL = " & external & "," & _
               "  ESTADO_ID = " & ESTADO_ID & "," & _
               "  EMPRESA_ID = " & empresa_id & _
               " WHERE ID_EMPLEADO = " & ID
               execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al modificar el empleado en la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Function Eliminar(ID As Integer) As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "DELETE FROM empleados " & _
               " WHERE ID_EMPLEADO = " & ID
               execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar el empleado de la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Function Listado(nombre As String, cod As String, ESTADO As String, external As Long, empresa_id As Long, departamento As Long, DNI As String) As ADODB.Recordset
    Dim consulta As String
    Dim filtro As String
    If nombre <> "" Then
        filtro = filtro & " AND NOMBRE LIKE '%" & nombre & "%'"
    End If
    If DNI <> "" Then
        filtro = filtro & " AND CIF LIKE '%" & DNI & "%'"
    End If
    If cod <> "" Then
        filtro = filtro & " AND CODIGO_INTERNO LIKE '%" & cod & "%'"
    End If
    If empresa_id <> 0 Then
        filtro = filtro & " AND EMPRESA_ID = " & empresa_id
    End If
    'JGM : CARGA DE EMPRESAS DESDE LA QUERY
    consulta = "SELECT * " & _
               "  FROM empleados A LEFT JOIN EMPLEADOS_EMPRESAS B ON A.EMPRESA_ID = B.ID_EMPRESA " & _
               " WHERE 1 = 1  " & _
               filtro & _
               IIf(Trim(ESTADO) <> "", " AND ESTADO_ID = " & ESTADO, "") & _
               " AND EXTERNAL = " & external & _
               " order by nombre"
    'M0972-F
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Apodos() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT apodo,ID_EMPLEADO,ESTADO_ID FROM empleados order by apodo"
    Set Listado_Apodos = datos_bd(consulta)
End Function
Public Function Listado_Apodos_Activos() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT apodo,ID_EMPLEADO FROM empleados WHERE ESTADO_ID = 0 order by apodo"
    Set Listado_Apodos_Activos = datos_bd(consulta)
End Function
Public Function Listado_Nominas() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM empleados where ESTADO_ID <> 2 order by codigo_interno asc"
    Set Listado_Nominas = datos_bd(consulta)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "empleados"
        .setDESCRIPCION = "Empleados"
        .setPK = "ID_EMPLEADO"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = False
        .setFILTRO = filtro
        Set .FORMULARIO = FORMULARIO
    End With
End Function

