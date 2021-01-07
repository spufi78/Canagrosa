VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos_Proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS_PROVEEDORES
'***************************************************************
Private ID_PROVEEDOR As Long
Private DESCRIPCION As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS_PROVEEDORES
'***************************************************************
Public Property Let setID_PROVEEDOR(ByVal dato As Long)
        ID_PROVEEDOR = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS_PROVEEDORES
'***************************************************************
Public Property Get getID_PROVEEDOR() As Long
        getID_PROVEEDOR = ID_PROVEEDOR
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS_PROVEEDORES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM EQUIPOS_PROVEEDORES WHERE ID_PROVEEDOR = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PROVEEDOR = rs("ID_PROVEEDOR")
                DESCRIPCION = rs("DESCRIPCION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEQUIPOS_PROVEEDORES)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT MAX(ID_PROVEEDOR) FROM EQUIPOS_PROVEEDORES"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_PROVEEDOR = 1
        Else
                ID_PROVEEDOR = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO EQUIPOS_PROVEEDORES " & _
                   "  VALUES (" & _
                        ID_PROVEEDOR & "," & "'" & DESCRIPCION & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_PROVEEDOR
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEQUIPOS_PROVEEDORES)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE EQUIPOS_PROVEEDORES SET " & _
                        " DESCRIPCION = '" & DESCRIPCION & "'" & _
                " WHERE ID_PROVEEDOR = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEQUIPOS_PROVEEDORES)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        consulta = "SELECT * FROM EQUIPOS WHERE PROVEEDOR_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            MsgBox "No se puede eliminar. Existen equipos con este proveedor.", vbCritical, App.Title
            Eliminar = False
            Exit Function
        End If
        consulta = "DELETE FROM EQUIPOS_PROVEEDORES " & _
                "    WHERE ID_PROVEEDOR = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEQUIPOS_PROVEEDORES)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM EQUIPOS_PROVEEDORES ORDER BY DESCRIPCION"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT ID_PROVEEDOR,DESCRIPCION FROM EQUIPOS_PROVEEDORES ORDER BY DESCRIPCION"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function llenar_combo(combo As miCombo, FK As Long, FORMULARIO As Form, FILTRO As String)
    With combo
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "EQUIPOS_PROVEEDORES"
        .setDESCRIPCION = "Proveedores Equipos"
        .setPK = "ID_PROVEEDOR"
        .setCAMPO = "DESCRIPCION"
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function


