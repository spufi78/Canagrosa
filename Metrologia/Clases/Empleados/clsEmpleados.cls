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
Private NOMBRE As String
Private APODO As String
Private DIRECCION As String
Private CP As Long
Private PROVINCIA_ID As Long
Private MUNICIPIO_ID As Long
Private CIF As String
Private telefono As String
Private movil As String
Private foto As String
Private OBSERVACIONES As String
Private ESTADO_ID As Integer
Private CODIGO_INTERNO As String
Private ccc As String
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
    NOMBRE = Trim(dato)
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
    CP = dato
End Property
Public Property Let setCIF(ByVal dato As String)
    CIF = Trim(dato)
End Property
Public Property Let setTELEFONO(ByVal dato As String)
    telefono = Trim(dato)
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
Public Property Get getID_EMPLEADO() As Long
    getID_EMPLEADO = ID_EMPLEADO
End Property
Public Property Get getUSUARIO_ID() As Long
    getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
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
    getCP = CP
End Property
Public Property Get getCIF() As String
    getCIF = CIF
End Property
Public Property Get getTELEFONO() As String
    getTELEFONO = telefono
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
Public Function cargar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim RS As New ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM EMPLEADOS WHERE ID_EMPLEADO=" & ID
        Set RS = datos_bd(consulta)
        If RS.RecordCount = 0 Then
            cargar = False
        Else
            ID_EMPLEADO = RS.Fields("ID_EMPLEADO")
            USUARIO_ID = RS.Fields("USUARIO_ID")
            NOMBRE = RS.Fields("NOMBRE")
            APODO = RS.Fields("APODO")
            DIRECCION = RS.Fields("DIRECCION")
            PROVINCIA_ID = RS.Fields("PROVINCIA_ID")
            If Not IsNull(RS.Fields("CP")) Then
                CP = RS.Fields("CP")
            Else
                CP = 0
            End If
            CIF = RS.Fields("CIF")
            telefono = RS.Fields("TELEFONO")
            movil = RS.Fields("MOVIL")
            MUNICIPIO_ID = RS.Fields("MUNICIPIO_ID")
            foto = RS.Fields("FOTO")
            OBSERVACIONES = RS.Fields("OBSERVACIONES")
            ESTADO_ID = RS.Fields("ESTADO_ID")
            CODIGO_INTERNO = RS.Fields("CODIGO_INTERNO")
            ccc = RS.Fields("CCC")
            cargar = True
        End If
        Set RS = Nothing
        Exit Function
fallo:
        cargar = False
        MsgBox "Error al cargar los datos del empleado (Cargar)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long
    Dim consulta As String
    On Error GoTo fallo:
    Me.CrearID_EMPLEADO
    consulta = "INSERT INTO EMPLEADOS" & _
               " VALUES(" & _
               ID_EMPLEADO & "," & USUARIO_ID & ",'" & NOMBRE & "','" & APODO & "','" & DIRECCION & "'," & _
               CP & "," & PROVINCIA_ID & "," & MUNICIPIO_ID & ",'" & _
               CIF & "','" & telefono & "','" & movil & "','" & _
               foto & "','" & OBSERVACIONES & "'," & ESTADO_ID & ",'" & CODIGO_INTERNO & "','" & _
               ccc & "')"
    execute_bd consulta
    Insertar = ID_EMPLEADO
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al guardar el empleado en la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Sub CrearID_EMPLEADO()
    Dim RS As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_EMPLEADO) FROM EMPLEADOS"
    Set RS = datos_bd(consulta)
    If IsNull(RS.Fields(0)) Or (RS.EOF And RS.BOF) Then  'si es nulo No se recupero ninguno
        ID_EMPLEADO = 1
    Else
        ID_EMPLEADO = RS.Fields(0) + 1
    End If
    Set RS = Nothing
End Sub
Public Function Modificar(ID As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "UPDATE EMPLEADOS SET" & _
               "  USUARIO_ID = " & USUARIO_ID & "," & _
               "  NOMBRE = '" & NOMBRE & "'," & _
               "  APODO = '" & APODO & "'," & _
               "  DIRECCION = '" & DIRECCION & "'," & _
               "  CP = " & CP & "," & _
               "  PROVINCIA_ID = " & PROVINCIA_ID & "," & _
               "  MUNICIPIO_ID = " & MUNICIPIO_ID & "," & _
               "  CIF = '" & CIF & "'," & _
               "  TELEFONO = '" & telefono & "'," & _
               "  MOVIL = '" & movil & "'," & _
               "  FOTO = '" & foto & "'," & _
               "  OBSERVACIONES = '" & OBSERVACIONES & "'," & _
               "  CODIGO_INTERNO = '" & CODIGO_INTERNO & "'," & _
               "  CCC = '" & ccc & "'," & _
               "  ESTADO_ID = " & ESTADO_ID & _
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
    consulta = "DELETE FROM EMPLEADOS " & _
               " WHERE ID_EMPLEADO = " & ID
               execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar el empleado de la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Estado(estado As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS where ESTADO_ID = " & estado & " order by nombre"
    Set Listado_Estado = datos_bd(consulta)
End Function
Public Function Listado_Apodos() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT apodo,ID_EMPLEADO,ESTADO_ID FROM EMPLEADOS order by apodo"
    Set Listado_Apodos = datos_bd(consulta)
End Function
Public Function Listado_Apodos_Activos() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT apodo,ID_EMPLEADO FROM EMPLEADOS WHERE ESTADO_ID = 0 order by apodo"
    Set Listado_Apodos_Activos = datos_bd(consulta)
End Function
Public Function Listado_Nominas() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS where ESTADO_ID <> 2 order by codigo_interno asc"
    Set Listado_Nominas = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_EMPLEADO, NOMBRE FROM EMPLEADOS " & _
               " order by NOMBRE"
    Set Listado_Combo = datos_bd(consulta)
End Function

