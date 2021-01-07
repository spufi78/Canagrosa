VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUnidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***********************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE unidades           *
'***********************************************************
Private id_unidad As Integer
Private nombre As String
Private CONV_NM As String
Public Property Let setID_UNIDAD(ByVal dato As Integer)
    id_unidad = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Let setCONV_NM(ByVal dato As String)
    CONV_NM = Trim(dato)
End Property
Public Property Get getID_UNIDAD() As Integer
    getID_UNIDAD = id_unidad
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getCONV_NM() As String
    getCONV_NM = CONV_NM
End Property
Public Function Listado(nombre As String) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM unidades "
    If Trim(nombre) <> "" Then
     consulta = consulta & " WHERE nombre like '%" & nombre & "%'"
    End If
    consulta = consulta & " order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function CARGAR(Unidad As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM unidades WHERE ID_UNIDAD=" & Unidad
    Set rs = datos_bd(consulta)
    Dim existe As Boolean
    existe = False
    If rs.RecordCount > 0 Then
        id_unidad = rs("ID_UNIDAD")
        nombre = rs("NOMBRE")
        If Not IsNull(rs("CONV_NM")) Then
            CONV_NM = rs("CONV_NM")
        End If
        existe = True
    End If
    CARGAR = existe
    Exit Function
fallo:
    CARGAR = False
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_UNIDAD) FROM unidades"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        id_unidad = 1
    Else
        id_unidad = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function InsertarUnidad() As Long
    On Error GoTo fallo
    Dim consulta As String
    InsertarUnidad = 0
    If duplicado = False Then
        Me.CrearID
        consulta = "Insert into unidades " & _
                   " values(" & _
                   id_unidad & ",'" & nombre & "'," & CONV_NM & ")"
        execute_bd consulta
        InsertarUnidad = id_unidad
    End If
    Exit Function
fallo:
    InsertarUnidad = 0
    MsgBox "Error al insertar la unidad (InsertarUnidad)", vbCritical, Err.Description
End Function
Public Function Eliminar(Unidad As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM formulas_campos WHERE UNIDAD_ID = " & Unidad
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        MsgBox "Existen fórmulas con esa unidad. No se puede eliminar.", vbExclamation, App.Title
        Exit Function
    End If
    consulta = "DELETE FROM unidades WHERE ID_UNIDAD=" & Unidad
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function Modificar(Unidad As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "UPDATE unidades " & _
               " SET NOMBRE = '" & nombre & "'" & _
               "    ,CONV_NM = " & CONV_NM & _
               " WHERE ID_UNIDAD=" & Unidad
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
End Function

Public Function duplicado() As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_UNIDAD FROM unidades WHERE NOMBRE = '" & nombre & "'"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        duplicado = False
    Else
        MsgBox "Existe una unidad con la misma descripción.", vbExclamation, App.Title
        duplicado = True
    End If
    Set rs = Nothing
End Function
Public Function Unidad_Campo_Resultado(deter As Long) As String
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT u.nombre " & _
               " FROM determinaciones as d, formulas as f,formulas_campos as cf, unidades as u " & _
               " where d.formula_id = f.id_formula " & _
               "   and f.campo_id_resultado = cf.id_campo " & _
               "   and cf.unidad_id = u.id_unidad " & _
               "   and d.id_determinacion = " & deter
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        Unidad_Campo_Resultado = ""
    Else
        Unidad_Campo_Resultado = rs.Fields(0)
    End If
    Set rs = Nothing
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_UNIDAD, NOMBRE FROM UNIDADES " & _
               " ORDER BY NOMBRE"
    Set Listado_Combo = datos_bd(consulta)
End Function

'E0047-I
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "UNIDADES"
        .setDESCRIPCION = "Unidades"
        .setPK = "ID_UNIDAD"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function
'E0047-F
