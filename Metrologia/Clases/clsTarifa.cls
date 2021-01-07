VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTarifa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsTARIFAS
'***************************************************************
Private ID_TARIFA As Integer
Private DESCRIPCION As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsTARIFAS
'***************************************************************
Public Property Let setID_TARIFA(ByVal dato As Integer)
        ID_TARIFA = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsTARIFAS
'***************************************************************
Public Property Get getID_TARIFA() As Integer
        getID_TARIFA = ID_TARIFA
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsTARIFAS
'***************************************************************
Public Function Carga(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TARIFAS WHERE ID_TARIFA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_TARIFA = rs("ID_TARIFA")
                DESCRIPCION = rs("DESCRIPCION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTARIFAS)", vbCritical, Err.Description
End Function
Public Function Insertar() As Integer
        On Error GoTo fallo
        Dim consulta As String
        If Me.duplicado = False Then
            Me.CrearID
            consulta = "INSERT INTO TARIFAS " & _
                       "  VALUES (" & _
                            ID_TARIFA & "," & "'" & DESCRIPCION & "'" & _
                    ")"
            execute_bd consulta
            Insertar = ID_TARIFA
        Else
            MsgBox "Ya existe una familia con esa descripción.", vbExclamation, App.Title
        End If
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTARIFAS)", vbCritical, Err.Description
End Function
Public Function duplicado() As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_TARIFA FROM TARIFAS WHERE DESCRIPCION = '" & Me.getDESCRIPCION & "'"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        duplicado = False
    Else
        duplicado = True
    End If
    Set rs = Nothing
End Function
Public Function Modificar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE TARIFAS SET " & _
                        " DESCRIPCION = '" & DESCRIPCION & "'" & _
                " WHERE ID_TARIFA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTARIFAS)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TARIFAS ORDER BY DESCRIPCION"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_TARIFA, DESCRIPCION FROM TARIFAS " & _
               " order by DESCRIPCION"
    Set Listado_Combo = datos_bd(consulta)
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_TARIFA) FROM TARIFAS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_TARIFA = 1
    Else
        ID_TARIFA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub



