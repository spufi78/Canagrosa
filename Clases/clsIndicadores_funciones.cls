VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIndicadores_funciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsIndicadores_funciones
'***************************************************************
Private ID_FUNCION As Long
Private NOMBRE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsIndicadores_funciones
'***************************************************************
Public Property Let setID_FUNCION(ByVal dato As Long)
        ID_FUNCION = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsIndicadores_funciones
'***************************************************************
Public Property Get getID_FUNCION() As Long
        getID_FUNCION = ID_FUNCION
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsIndicadores_funciones
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM Indicadores_funciones WHERE ID_FUNCION = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_FUNCION = rs("ID_FUNCION")
                NOMBRE = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsIndicadores_funciones)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_FUNCION) FROM Indicadores_funciones"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_FUNCION = 1
        Else
                ID_FUNCION = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO Indicadores_funciones " & _
                   "  VALUES (" & _
                        ID_FUNCION & "," & "'" & NOMBRE & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_FUNCION
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsIndicadores_funciones)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE Indicadores_funciones SET " & _
                        " NOMBRE = '" & NOMBRE & "'" & _
                " WHERE ID_FUNCION = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsIndicadores_funciones)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM Indicadores_funciones " & _
                "    WHERE ID_FUNCION = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsIndicadores_funciones)", vbCritical, Err.Description
End Function
Public Function Listado() As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM Indicadores_funciones ORDER BY ID_FUNCION"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As adodb.Recordset
    Dim consulta As String
    consulta = "SELECT ID_FUNCION,NOMBRE FROM indicadores_funciones order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function

