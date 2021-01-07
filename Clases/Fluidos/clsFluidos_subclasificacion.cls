VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFluidos_subclasificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSFLUIDOS_SUBCLASIFICACION
'***************************************************************
Private ID_SUBCLASIFICACION As Long
Private NOMBRE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSFLUIDOS_SUBCLASIFICACION
'***************************************************************
Public Property Let setID_SUBCLASIFICACION(ByVal dato As Long)
        ID_SUBCLASIFICACION = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSFLUIDOS_SUBCLASIFICACION
'***************************************************************
Public Property Get getID_SUBCLASIFICACION() As Long
        getID_SUBCLASIFICACION = ID_SUBCLASIFICACION
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSFLUIDOS_SUBCLASIFICACION
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM FLUIDOS_SUBCLASIFICACION WHERE ID_SUBCLASIFICACION = " & ID
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_SUBCLASIFICACION = rs("ID_SUBCLASIFICACION")
                NOMBRE = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsFluidos_subclasificacion)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT MAX(ID_SUBCLASIFICACION) FROM FLUIDOS_SUBCLASIFICACION"
        Set rs = datos_bd(CONSULTA)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_SUBCLASIFICACION = 1
        Else
                ID_SUBCLASIFICACION = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim CONSULTA As String
        Me.CrearID
        CONSULTA = "INSERT INTO FLUIDOS_SUBCLASIFICACION " & _
                   "  VALUES (" & _
                        ID_SUBCLASIFICACION & "," & "'" & NOMBRE & "'" & _
                ")"
        execute_bd CONSULTA
        Insertar = ID_SUBCLASIFICACION
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsFluidos_subclasificacion)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE FLUIDOS_SUBCLASIFICACION SET " & _
                        " NOMBRE = '" & NOMBRE & "'" & _
                " WHERE ID_SUBCLASIFICACION = " & ID
        execute_bd CONSULTA
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFluidos_subclasificacion)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "DELETE FROM FLUIDOS_SUBCLASIFICACION " & _
                "    WHERE ID_SUBCLASIFICACION = " & ID
        execute_bd CONSULTA
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFluidos_subclasificacion)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM FLUIDOS_SUBCLASIFICACION ORDER BY ID_SUBCLASIFICACION"
        Set Listado = datos_bd(CONSULTA)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT ID_SUBCLASIFICACION,NOMBRE FROM FLUIDOS_SUBCLASIFICACION ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(CONSULTA)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "FLUIDOS_SUBCLASIFICACION"
        .setDESCRIPCION = "SubClasificación de Fluidos"
        .setPK = "ID_SUBCLASIFICACION"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function
