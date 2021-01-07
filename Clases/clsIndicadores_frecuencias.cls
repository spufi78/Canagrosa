VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIndicadores_frecuencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsIndicadores_frecuencias
'***************************************************************
Private ID_FRECUENCIA As Long
Private NOMBRE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsIndicadores_frecuencias
'***************************************************************
Public Property Let setID_FRECUENCIA(ByVal dato As Long)
        ID_FRECUENCIA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsIndicadores_frecuencias
'***************************************************************
Public Property Get getID_FRECUENCIA() As Long
        getID_FRECUENCIA = ID_FRECUENCIA
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsIndicadores_frecuencias
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM Indicadores_frecuencias WHERE ID_FRECUENCIA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_FRECUENCIA = rs("ID_FRECUENCIA")
                NOMBRE = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsIndicadores_frecuencias)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_FRECUENCIA) FROM Indicadores_frecuencias"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_FRECUENCIA = 1
        Else
                ID_FRECUENCIA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO Indicadores_frecuencias " & _
                   "  VALUES (" & _
                        ID_FRECUENCIA & "," & "'" & NOMBRE & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_FRECUENCIA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsIndicadores_frecuencias)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE Indicadores_frecuencias SET " & _
                        " NOMBRE = '" & NOMBRE & "'" & _
                " WHERE ID_FRECUENCIA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsIndicadores_frecuencias)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM Indicadores_frecuencias " & _
                "    WHERE ID_FRECUENCIA = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsIndicadores_frecuencias)", vbCritical, Err.Description
End Function
Public Function Listado() As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM Indicadores_frecuencias ORDER BY ID_FRECUENCIA"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT ID_FRECUENCIA,NOMBRE FROM Indicadores_frecuencias order by nombre"
        Set Listado_Combo = datos_bd(consulta)
End Function
