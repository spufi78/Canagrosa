VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlodine_Etiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsALODINE_ETIQUETAS
'***************************************************************
Private ID_ETIQUETA As Long
Private DESCRIPCION As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsALODINE_ETIQUETAS
'***************************************************************
Public Property Let setID_ETIQUETA(ByVal dato As Long)
        ID_ETIQUETA = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsALODINE_ETIQUETAS
'***************************************************************
Public Property Get getID_ETIQUETA() As Long
        getID_ETIQUETA = ID_ETIQUETA
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsALODINE_ETIQUETAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ALODINE_ETIQUETAS WHERE ID_ETIQUETA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_ETIQUETA = rs("ID_ETIQUETA")
                DESCRIPCION = rs("DESCRIPCION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsALODINE_ETIQUETAS)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_ETIQUETA) FROM ALODINE_ETIQUETAS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_ETIQUETA = 1
        Else
                ID_ETIQUETA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO ALODINE_ETIQUETAS " & _
                   "  VALUES (" & _
                        ID_ETIQUETA & "," & "'" & DESCRIPCION & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_ETIQUETA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsALODINE_ETIQUETAS)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ALODINE_ETIQUETAS SET " & _
                        " DESCRIPCION = '" & DESCRIPCION & "'" & _
                " WHERE ID_ETIQUETA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsALODINE_ETIQUETAS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ALODINE_ETIQUETAS " & _
                "    WHERE ID_ETIQUETA = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsALODINE_ETIQUETAS)", vbCritical, Err.Description
End Function
Public Function Listado() As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ALODINE_ETIQUETAS ORDER BY ID_ETIQUETA"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As adodb.Recordset
    Dim consulta As String
    consulta = "SELECT ID_ETIQUETA, DESCRIPCION FROM ALODINE_ETIQUETAS " & _
               " order by DESCRIPCION"
    Set Listado_Combo = datos_bd(consulta)
End Function

