VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFamilias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsFamilias
'***************************************************************
Private ID_FAMILIA As Integer
Private NOMBRE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsFamilias
'***************************************************************
Public Property Let setID_FAMILIA(ByVal dato As Integer)
        ID_FAMILIA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsFamilias
'***************************************************************
Public Property Get getID_FAMILIA() As Integer
        getID_FAMILIA = ID_FAMILIA
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsFamilias
'***************************************************************
Public Function Carga(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM familias WHERE ID_FAMILIA = " & ID
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
        MsgBox "Error al cargar los datos(clsFamilias)", vbCritical, Err.Description
End Function
Public Function Insertar() As Integer
        On Error GoTo fallo
        Dim consulta As String
        If Me.duplicado = False Then
            Me.CrearID
            consulta = "INSERT INTO familias " & _
                       "  VALUES (" & _
                            ID_FAMILIA & "," & "'" & NOMBRE & "'" & _
                    ")"
            execute_bd consulta
            Insertar = ID_FAMILIA
        Else
            MsgBox "Ya existe una familia con esa descripción.", vbExclamation, App.Title
        End If
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsFamilias)", vbCritical, Err.Description
End Function
Public Function duplicado() As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_FAMILIA FROM FAMILIAS WHERE NOMBRE = '" & Me.getNOMBRE & "'"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        duplicado = False
    Else
        duplicado = True
    End If
    Set rs = Nothing
End Function
Public Function Modificar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE familias SET " & _
                        " NOMBRE = '" & NOMBRE & "'" & _
                " WHERE ID_FAMILIA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFamilias)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        ' Validar subfamilias
        Dim osubfamilia As New clsSubfamilias
        Dim rs As ADODB.Recordset
        Set rs = osubfamilia.Listado_por_familia(ID)
        If rs.RecordCount > 0 Then
            MsgBox "Existen subfamilias asociadas a esta familia. No se puede eliminar.", vbInformation, App.Title
            Exit Function
        End If
        ' Validar articulo
'        Dim oArticulo As New clsArticulo
'        Set rs = oArticulo.Listado("", CStr(ID), "", "", False)
'        If rs.RecordCount > 0 Then
'            MsgBox "Existen artículos asociados a esta familia. No se puede eliminar.", vbInformation, App.Title
'            Exit Function
'        End If
        consulta = "DELETE FROM familias " & _
                "    WHERE ID_FAMILIA = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFamilias)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM familias WHERE ID_FAMILIA <> 0 ORDER BY NOMBRE"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_FAMILIA, NOMBRE FROM familias " & _
               " order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_FAMILIA) FROM familias"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_FAMILIA = 1
    Else
        ID_FAMILIA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub


