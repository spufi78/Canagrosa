VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTamanos_etiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTAMANOS_ETIQUETA
'***************************************************************
Private ID_TAMANO_ETIQUETA As Long
Private nombre As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTAMANOS_ETIQUETA
'***************************************************************
Public Property Let setID_TAMANO_ETIQUETA(ByVal dato As Long)
        ID_TAMANO_ETIQUETA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        nombre = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTAMANOS_ETIQUETA
'***************************************************************
Public Property Get getID_TAMANO_ETIQUETA() As Long
        getID_TAMANO_ETIQUETA = ID_TAMANO_ETIQUETA
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTAMANOS_ETIQUETA
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TAMANOS_ETIQUETA WHERE ID_TAMANO_ETIQUETA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_TAMANO_ETIQUETA = rs("ID_TAMANO_ETIQUETA")
                nombre = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTamanos_etiqueta)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_TAMANO_ETIQUETA) FROM TAMANOS_ETIQUETA"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_TAMANO_ETIQUETA = 1
        Else
                ID_TAMANO_ETIQUETA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO TAMANOS_ETIQUETA " & _
                   "  VALUES (" & _
                        ID_TAMANO_ETIQUETA & "," & "'" & nombre & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_TAMANO_ETIQUETA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTamanos_etiqueta)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE TAMANOS_ETIQUETA SET " & _
                        " NOMBRE = '" & nombre & "'" & _
                " WHERE ID_TAMANO_ETIQUETA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTamanos_etiqueta)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM TAMANOS_ETIQUETA " & _
                "    WHERE ID_TAMANO_ETIQUETA = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTamanos_etiqueta)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TAMANOS_ETIQUETA ORDER BY ID_TAMANO_ETIQUETA"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_TAMANO_ETIQUETA,NOMBRE FROM TAMANOS_ETIQUETA ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
