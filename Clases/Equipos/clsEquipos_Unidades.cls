VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos_Unidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS_UNIDADES
'***************************************************************
Private ID_UNIDAD As Long
Private DESCRIPCION As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS_UNIDADES
'***************************************************************
Public Property Let setID_UNIDAD(ByVal dato As Long)
        ID_UNIDAD = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS_UNIDADES
'***************************************************************
Public Property Get getID_UNIDAD() As Long
        getID_UNIDAD = ID_UNIDAD
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS_UNIDADES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT * FROM EQUIPOS_UNIDADES WHERE ID_UNIDAD = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_UNIDAD = rs("ID_UNIDAD")
                DESCRIPCION = rs("DESCRIPCION")
                Carga = True
        End If
        Set rs = Nothing
        
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEQUIPOS_UNIDADES)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT MAX(ID_UNIDAD) FROM EQUIPOS_UNIDADES"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_UNIDAD = 1
        Else
                ID_UNIDAD = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        
        Me.CrearID
        consulta = "INSERT INTO EQUIPOS_UNIDADES " & _
                   "  VALUES (" & ID_UNIDAD & "," & "'" & DESCRIPCION & "'" & ")"
        execute_bd consulta
        Insertar = ID_UNIDAD
        
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEQUIPOS_UNIDADES)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE EQUIPOS_UNIDADES SET " & _
                   "       DESCRIPCION = '" & DESCRIPCION & "'" & _
                   " WHERE ID_UNIDAD = " & ID
        execute_bd consulta
        Modificar = True
        
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEQUIPOS_UNIDADES)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        
        consulta = "SELECT * FROM EQUIPOS WHERE UNIDAD_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            MsgBox "No se puede eliminar. Existen equipos con esta unidad.", vbCritical, App.Title
            Eliminar = False
            Exit Function
        End If
        consulta = "DELETE FROM EQUIPOS_UNIDADES " & _
                   " WHERE ID_UNIDAD = " & ID
        execute_bd consulta
        Eliminar = True
        
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEQUIPOS_UNIDADES)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT * FROM EQUIPOS_UNIDADES ORDER BY NOMBRE"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT ID_UNIDAD, DESCRIPCION " & _
                   "  FROM EQUIPOS_UNIDADES " & _
                   " ORDER BY DESCRIPCION"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function llenar_combo(combo As miCombo, FK As Long, FORMULARIO As Form, FILTRO As String)
    With combo
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "EQUIPOS_UNIDADES"
        .setDESCRIPCION = "Unidades Equipos"
        .setPK = "ID_UNIDAD"
        .setCAMPO = "DESCRIPCION"
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function
