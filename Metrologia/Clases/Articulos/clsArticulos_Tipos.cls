VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArticulos_Tipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsARTICULOS_TIPOS
'***************************************************************
Private ID_TIPO_ARTICULO As Integer
Private NOMBRE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsARTICULOS_TIPOS
'***************************************************************
Public Property Let setID_TIPO_ARTICULO(ByVal dato As Integer)
        ID_TIPO_ARTICULO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsARTICULOS_TIPOS
'***************************************************************
Public Property Get getID_TIPO_ARTICULO() As Integer
        getID_TIPO_ARTICULO = ID_TIPO_ARTICULO
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsARTICULOS_TIPOS
'***************************************************************
Public Function Carga(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ARTICULOS_TIPOS WHERE ID_TIPO_ARTICULO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_TIPO_ARTICULO = rs("ID_TIPO_ARTICULO")
                NOMBRE = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsARTICULOS_TIPOS)", vbCritical, Err.Description
End Function
Public Function Insertar() As Integer
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO ARTICULOS_TIPOS " & _
                   "  VALUES (" & _
                        ID_TIPO_ARTICULO & "," & "'" & NOMBRE & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_TIPO_ARTICULO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsARTICULOS_TIPOS)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ARTICULOS_TIPOS SET " & _
                        " NOMBRE = '" & NOMBRE & "'" & _
                " WHERE ID_TIPO_ARTICULO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsARTICULOS_TIPOS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ARTICULOS_TIPOS " & _
                "    WHERE ID_TIPO_ARTICULO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsARTICULOS_TIPOS)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ARTICULOS_TIPOS WHERE ID_TIPO_ARTICULO <> 0 ORDER BY NOMBRE"
        Set Listado = datos_bd(consulta)
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_TIPO_ARTICULO) FROM ARTICULOS_TIPOS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_TIPO_ARTICULO = 1
    Else
        ID_TIPO_ARTICULO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_TIPO_ARTICULO, NOMBRE FROM ARTICULOS_TIPOS " & _
               " order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function





