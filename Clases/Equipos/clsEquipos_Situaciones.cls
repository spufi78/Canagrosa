VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos_Situaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS_SITUACIONES
'***************************************************************
Private ID_SITUACION As Long
Private NOMBRE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS_SITUACIONES
'***************************************************************
Public Property Let setID_SITUACION(ByVal dato As Long)
        ID_SITUACION = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS_SITUACIONES
'***************************************************************
Public Property Get getID_SITUACION() As Long
        getID_SITUACION = ID_SITUACION
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS_SITUACIONES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT * FROM EQUIPOS_SITUACIONES WHERE ID_SITUACION = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_SITUACION = rs("ID_SITUACION")
                NOMBRE = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEQUIPOS_SITUACIONES)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT MAX(ID_SITUACION) FROM EQUIPOS_SITUACIONES"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_SITUACION = 1
        Else
                ID_SITUACION = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        
        Me.CrearID
        consulta = "INSERT INTO EQUIPOS_SITUACIONES " & _
                   "  VALUES (" & ID_SITUACION & "," & "'" & NOMBRE & "'" & ")"
        execute_bd consulta
        Insertar = ID_SITUACION
        
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEQUIPOS_SITUACIONES)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE EQUIPOS_SITUACIONES SET " & _
                   "       NOMBRE = '" & NOMBRE & "'" & _
                   " WHERE ID_SITUACION = " & ID
        execute_bd consulta
        Modificar = True
        
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEQUIPOS_SITUACIONES)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        
        consulta = "SELECT * FROM EQUIPOS WHERE SITUACION_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            MsgBox "No se puede eliminar. Existen equipos con esta situación.", vbCritical, App.Title
            Eliminar = False
            Exit Function
        End If
        consulta = "DELETE FROM EQUIPOS_SITUACIONES " & _
                   " WHERE ID_SITUACION = " & ID
        execute_bd consulta
        Eliminar = True
        
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEQUIPOS_SITUACIONES)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT * FROM EQUIPOS_SITUACIONES ORDER BY NOMBRE"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT ID_SITUACION, NOMBRE " & _
                   "  FROM EQUIPOS_SITUACIONES " & _
                   " ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function llenar_combo(combo As miCombo, FK As Long, FORMULARIO As Form, FILTRO As String)
    With combo
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "EQUIPOS_SITUACIONES"
        .setDESCRIPCION = "Situaciones Equipos"
        .setPK = "ID_SITUACION"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function
