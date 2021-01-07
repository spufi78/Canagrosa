VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos_Familias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS_PROVEEDORES
'***************************************************************
Private ID_FAMILIA As Long
Private NOMBRE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS_PROVEEDORES
'***************************************************************
Public Property Let setID_FAMILIA(ByVal dato As Long)
        ID_FAMILIA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS_PROVEEDORES
'***************************************************************
Public Property Get getID_FAMILIA() As Long
        getID_FAMILIA = ID_FAMILIA
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS_FAMILIAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT * FROM EQUIPOS_FAMILIAS WHERE ID_FAMILIA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_FAMILIA = rs("ID_FAMILIA")
                NOMBRE = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEQUIPOS_FAMILIAS)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT MAX(ID_FAMILIA) FROM EQUIPOS_FAMILIAS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_FAMILIAS = 1
        Else
                ID_FAMILIAS = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        
        Me.CrearID
        consulta = "INSERT INTO EQUIPOS_FAMILIAS " & _
                   "VALUES (" & ID_FAMILIAS & "," & "'" & NOMBRE & "'" & ")"
        execute_bd consulta
        Insertar = ID_FAMILIA
        
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEQUIPOS_FAMILIAS)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE EQUIPOS_FAMILIAS SET " & _
                   "       NOMBRE = '" & NOMBRE & "'" & _
                   " WHERE ID_FAMILIA = " & ID
        execute_bd consulta
        Modificar = True
        
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEQUIPOS_FAMILIAS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        
        consulta = "SELECT * FROM EQUIPOS WHERE FAMILIA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            MsgBox "No se puede eliminar. Existen equipos con esta familia.", vbCritical, App.Title
            Eliminar = False
            Exit Function
        End If
        consulta = "DELETE FROM EQUIPOS_FAMILIAS " & _
                   " WHERE ID_FAMILIA = " & ID
        execute_bd consulta
        Eliminar = True
        
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEQUIPOS_FAMILIAS)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT * FROM EQUIPOS_FAMILIAS ORDER BY NOMBRE"
        
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT ID_FAMILIA, NOMBRE " & _
                   "  FROM EQUIPOS_FAMILIAS " & _
                   " ORDER BY NOMBRE"
        
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function llenar_combo(combo As miCombo, FK As Long, FORMULARIO As Form)
    With combo
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "EQUIPOS_FAMILIAS"
        .setDESCRIPCION = "Familias Equipos"
        .setPK = "ID_FAMILIA"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
    
End Function




