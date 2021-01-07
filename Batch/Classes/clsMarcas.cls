VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMarcas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSMARCAS
'***************************************************************
Private ID_MARCA As Long
Private nombre As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSMARCAS
'***************************************************************
Public Property Let setID_MARCA(ByVal dato As Long)
        ID_MARCA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        nombre = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSMARCAS
'***************************************************************
Public Property Get getID_MARCA() As Long
        getID_MARCA = ID_MARCA
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSMARCAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM MARCAS WHERE ID_MARCA = " & ID
        Set rs = datos_bd(consulta, conn2)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_MARCA = rs("ID_MARCA")
                nombre = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsMarcas)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_MARCA) FROM MARCAS"
        Set rs = datos_bd(consulta, conn2)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_MARCA = 1
        Else
                ID_MARCA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        
        Dim rs As ADODB.Recordset
        consulta = "SELECT * FROM MARCAS WHERE NOMBRE = '" & nombre & "'"
        Set rs = datos_bd(consulta, conn2)
        If rs.RecordCount = 0 Then
            Me.CrearID
            consulta = "INSERT INTO MARCAS " & _
                       "  VALUES (" & _
                            ID_MARCA & "," & "'" & nombre & "'" & _
                    ")"
            execute_bd consulta, conn2
        Else
            Me.Modificar rs("ID_MARCA")
        End If
        Insertar = ID_MARCA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsMarcas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE MARCAS SET " & _
                        " NOMBRE = '" & nombre & "'" & _
                " WHERE ID_MARCA = " & ID
        execute_bd consulta, conn2
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsMarcas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM MARCAS " & _
                "    WHERE ID_MARCA = " & ID
        execute_bd consulta, conn2
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsMarcas)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM MARCAS ORDER BY ID_MARCA"
        Set Listado = datos_bd(consulta, conn2)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_MARCA,NOMBRE FROM MARCAS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta, conn2)
End Function
