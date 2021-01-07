VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***********************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE formulas           *
'***********************************************************
Private ID_FORMULA As Integer
Private CODIGO As String
Private DESCRIPCION As String
Private EXPRESION As String
Private Antigua As Integer
Private NOMBRE As String
Private CAMPO_ID_RESULTADO As Integer
Private FORMULA_ID_ORIGEN As Integer
Public Property Let setID_FORMULA(ByVal dato As Integer)
    ID_FORMULA = dato
End Property
Public Property Let setFORMULA_ID_ORIGEN(ByVal dato As Integer)
    FORMULA_ID_ORIGEN = dato
End Property
Public Property Let setCODIGO(ByVal dato As String)
    CODIGO = Trim(dato)
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
    DESCRIPCION = Trim(dato)
End Property
Public Property Let setEXPRESION(ByVal dato As String)
    EXPRESION = Trim(dato)
End Property
Public Property Let setANTIGUA(ByVal dato As Integer)
    Antigua = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Let setCAMPO_ID_RESULTADO(ByVal dato As Integer)
    CAMPO_ID_RESULTADO = dato
End Property
Public Property Get getID_FORMULA() As Integer
    getID_FORMULA = ID_FORMULA
End Property
Public Property Get getFORMULA_ID_ORIGEN() As Integer
    getFORMULA_ID_ORIGEN = FORMULA_ID_ORIGEN
End Property
Public Property Get getCODIGO() As String
    getCODIGO = CODIGO
End Property
Public Property Get getDESCRIPCION() As String
    getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getEXPRESION() As String
    getEXPRESION = EXPRESION
End Property
Public Property Get getANTIGUA() As Integer
    getANTIGUA = Antigua
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property
Public Property Get getCAMPO_ID_RESULTADO() As Integer
    getCAMPO_ID_RESULTADO = CAMPO_ID_RESULTADO
End Property
Public Function Listado(filtro As String, TD As Long) As ADODB.Recordset
    Dim CONSULTA As String
    If TD = 0 Then
        CONSULTA = "SELECT f.nombre as n,c.nombre,f.id_formula, f.descripcion " & _
                   "  FROM formulas f, formulas_campos c" & _
                   " where f.id_formula = c.formula_id " & _
                   "   and f.campo_id_resultado = c.id_campo " & _
                   "   and f.antigua = 0 " & _
                   "   AND concat(f.NOMBRE,' ',f.descripcion) like '%" & filtro & "%'" & _
                   " order by f.nombre"
    Else
        CONSULTA = "SELECT f.nombre as n,c.nombre,f.id_formula, f.descripcion " & _
                   "  FROM formulas f, formulas_campos c,tipos_determinacion d " & _
                   " where f.id_formula = c.formula_id " & _
                   "   and f.campo_id_resultado = c.id_campo " & _
                   "   and f.antigua = 0 " & _
                   "   AND concat(f.NOMBRE,' ',f.descripcion) like '%" & filtro & "%'" & _
                   "   AND d.id_tipo_determinacion = " & TD & _
                   "   and d.formula_id = f.id_formula " & _
                   " order by f.nombre"
    End If
    Set Listado = datos_bd(CONSULTA)
End Function
Public Function Listado_no_utilizadas() As ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT f.nombre as n,c.nombre,f.id_formula " & _
               "  FROM formulas f, formulas_campos c" & _
               " where f.id_formula = c.formula_id " & _
               "   and f.campo_id_resultado = c.id_campo " & _
               "   and f.antigua = 0 " & _
               "   and id_formula not in (select formula_id from tipos_determinacion) " & _
               " order by f.nombre"
    Set Listado_no_utilizadas = datos_bd(CONSULTA)
End Function

Public Function CARGAR(ByVal id_for As Integer) As Boolean
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "select * from formulas where id_formula = " & id_for
    Set rs = datos_bd(CONSULTA)
    If rs.RecordCount = 0 Then
        CARGAR = False
    Else
        CARGAR = True
        ID_FORMULA = rs("ID_FORMULA")
        CODIGO = rs("CODIGO")
        DESCRIPCION = rs("DESCRIPCION")
        EXPRESION = rs("EXPRESION")
        Antigua = rs("ANTIGUA")
        NOMBRE = rs("NOMBRE")
        CAMPO_ID_RESULTADO = rs("CAMPO_ID_RESULTADO")
        FORMULA_ID_ORIGEN = rs("FORMULA_ID_ORIGEN")
    End If
    Set rs = Nothing
    Exit Function
fallo:
    CARGAR = False
    MsgBox "Error al cargar la formula", vbCritical, Err.Description
End Function
Public Function CARGAR_ORIGEN(ByVal id_for As Integer) As Boolean
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "select * from formulas where FORMULA_ID_ORIGEN = " & id_for
    Set rs = datos_bd(CONSULTA)
    If rs.RecordCount = 0 Then
        If CARGAR(id_for) Then
            CARGAR_ORIGEN = True
        Else
            CARGAR_ORIGEN = False
        End If
    Else
        CARGAR_ORIGEN = True
        ID_FORMULA = rs("ID_FORMULA")
        CODIGO = rs("CODIGO")
        DESCRIPCION = rs("DESCRIPCION")
        EXPRESION = rs("EXPRESION")
        Antigua = rs("ANTIGUA")
        NOMBRE = rs("NOMBRE")
        CAMPO_ID_RESULTADO = rs("CAMPO_ID_RESULTADO")
        FORMULA_ID_ORIGEN = rs("FORMULA_ID_ORIGEN")
    End If
    Set rs = Nothing
    Exit Function
fallo:
    CARGAR_ORIGEN = False
    MsgBox "Error al cargar la formula Origen", vbCritical, Err.Description
End Function

Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT MAX(ID_FORMULA) FROM FORMULAS"
    Set rs = datos_bd(CONSULTA)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_FORMULA = 1
    Else
        ID_FORMULA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function InsertarFormula() As Long
    On Error GoTo fallo
    Dim CONSULTA As String
    Me.CrearID
    CONSULTA = "Insert into formulas " & _
               " values(" & _
               ID_FORMULA & ",'" & CODIGO & "','" & DESCRIPCION & "'," & _
               CAMPO_ID_RESULTADO & ",'" & EXPRESION & "',0,'" & NOMBRE & "'," & _
               FORMULA_ID_ORIGEN & ")"
    execute_bd CONSULTA
    InsertarFormula = ID_FORMULA
    Exit Function
fallo:
    InsertarFormula = 0
    error_grave ("Error al insertar la formula (InsertarFormula)" & Err.Description)
End Function
Public Function Modificar(ID) As Long
    On Error GoTo fallo
    Dim CONSULTA As String
    CONSULTA = "update formulas " & _
               " set " & _
               " DESCRIPCION = '" & DESCRIPCION & "'," & _
               " CAMPO_ID_RESULTADO = " & CAMPO_ID_RESULTADO & "," & _
               " FORMULA_ID_ORIGEN = " & FORMULA_ID_ORIGEN & "," & _
               " NOMBRE = '" & NOMBRE & "'" & _
               " WHERE ID_FORMULA = " & ID
    execute_bd CONSULTA
    Modificar = ID
    Exit Function
fallo:
    Modificar = 0
    MsgBox "Error al Modificar la formula (Modificar)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim CONSULTA As String
   On Error GoTo Eliminar_Error

    CONSULTA = "select id_tipo_determinacion from tipos_determinacion where formula_id = " & ID & " and anulado = 0"
    Set rs = datos_bd(CONSULTA)
    If rs.RecordCount <> 0 Then
        MsgBox "No se puede eliminar la formula. Esta asociada a algunos tipos de determinacion. ID_TIPO_DETERMINACION = " & rs(0), vbExclamation, App.Title
        Exit Function
    End If
    ' Borramos la Formula
    execute_bd "UPDATE FORMULAS SET ANTIGUA = 1 WHERE ID_FORMULA =" & ID
    Eliminar = True
   On Error GoTo 0
   Exit Function

Eliminar_Error:
    Eliminar = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Eliminar of Módulo de clase clsFormulas"
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "FORMULAS"
        .setDESCRIPCION = "Fórmulas"
        .setPK = "ID_FORMULA"
        .setCAMPO = "CONCAT(NOMBRE, IF (DESCRIPCION='','',' -> '),DESCRIPCION)"
        .setFILTRO = " ANTIGUA=0 "
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = FORMULARIO
    End With
End Function
Public Function Pasar_Antigua(ID As Long) As Boolean
   On Error GoTo Antigua_Error

    execute_bd "UPDATE FORMULAS SET ANTIGUA = 1 WHERE ID_FORMULA =" & ID
    Pasar_Antigua = True
   On Error GoTo 0
   Exit Function

Antigua_Error:
    Pasar_Antigua = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Antigua of Módulo de clase clsFormulas"
End Function

