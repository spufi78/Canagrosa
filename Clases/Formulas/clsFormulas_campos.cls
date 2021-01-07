VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormulas_campos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE formulas_campos      *
'************************************************************
Private ID_CAMPO As Integer
Private FORMULA_ID As Integer
Private nombre As String
Private ENTEROS As Integer
Private DECIMALES As Integer
Private REQUERIDO As Integer
Private FORMULA_ID_REL As Integer
Private ES_SOLUCION As Integer
Private UNIDAD_ID As Integer

Public Property Let setID_CAMPO(ByVal dato As Integer)
    ID_CAMPO = dato
End Property
Public Property Let setFORMULA_ID(ByVal dato As Integer)
    FORMULA_ID = dato
End Property

Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property

Public Property Let setENTEROS(ByVal dato As Integer)
    ENTEROS = dato
End Property
Public Property Let setDECIMALES(ByVal dato As Integer)
    DECIMALES = dato
End Property
Public Property Let setREQUERIDO(ByVal dato As Integer)
    REQUERIDO = dato
End Property

Public Property Let setFORMULA_ID_REL(ByVal dato As Integer)
    FORMULA_ID_REL = dato
End Property
Public Property Let setES_SOLUCION(ByVal dato As Integer)
    ES_SOLUCION = dato
End Property
Public Property Let setUNIDAD_ID(ByVal dato As Integer)
    UNIDAD_ID = dato
End Property
Public Property Get getID_CAMPO() As Integer
    getID_CAMPO = ID_CAMPO
End Property
Public Property Get getFORMULA_ID() As Integer
    getFORMULA_ID = FORMULA_ID
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getENTEROS() As Integer
    getENTEROS = ENTEROS
End Property
Public Property Get getDECIMALES() As Integer
    getDECIMALES = DECIMALES
End Property
Public Property Get getREQUERIDO() As Integer
    getREQUERIDO = REQUERIDO
End Property
Public Property Get getFORMULA_ID_REL() As Integer
    getFORMULA_ID_REL = FORMULA_ID_REL
End Property
Public Property Get getES_SOLUCION() As Integer
    getES_SOLUCION = ES_SOLUCION
End Property
Public Property Get getUNIDAD_ID() As Integer
    getUNIDAD_ID = UNIDAD_ID
End Property
Public Function ListaFormulas(FORMULA As Long) As ADODB.RecordSet
    Dim consulta As String
    consulta = "SELECT * FROM formulas_campos WHERE formula_id=" & FORMULA & " order by requerido desc, id_campo asc"
    Set ListaFormulas = datos_bd(consulta)
End Function
Public Function Lista_Formulas_Unidades(FORMULA As Integer) As ADODB.RecordSet
    Dim consulta As String
    consulta = "SELECT cf.id_campo, cf.nombre, cf.es_solucion, u.nombre, cf.enteros, cf.decimales " & _
               "  FROM formulas_campos cf, unidades u " & _
               " WHERE cf.formula_id=" & FORMULA & _
               "   AND cf.unidad_id = u.id_unidad " & _
               " ORDER BY cf.requerido desc, cf.id_campo asc"
    Set Lista_Formulas_Unidades = datos_bd(consulta)
End Function
Public Function CARGAR(CAMPO As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.RecordSet
    Dim consulta As String
    consulta = "SELECT * FROM formulas_campos WHERE ID_CAMPO=" & CAMPO
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            CARGAR = False
        Else
            ID_CAMPO = rs("ID_CAMPO")
            FORMULA_ID = rs("FORMULA_ID")
            nombre = rs("NOMBRE")
            ENTEROS = rs("ENTEROS")
            DECIMALES = rs("DECIMALES")
            REQUERIDO = rs("REQUERIDO")
            FORMULA_ID_REL = rs("FORMULA_ID_REL")
            ES_SOLUCION = rs("ES_SOLUCION")
            UNIDAD_ID = rs("UNIDAD_ID")
            CARGAR = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        CARGAR = False
        MsgBox "Error al cargar los formulas_campos", vbCritical, Err.Description
End Function
Public Sub CrearID()
    Dim rs As New ADODB.RecordSet
    Dim consulta As String
    consulta = "SELECT MAX(ID_CAMPO) FROM formulas_campos"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_CAMPO = 1
    Else
        ID_CAMPO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function InsertarCamposFormula() As Long
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearID
    consulta = "Insert into formulas_campos " & _
               " values(" & _
               ID_CAMPO & "," & FORMULA_ID & ",'" & nombre & "'," & _
               ENTEROS & "," & DECIMALES & "," & REQUERIDO & "," & _
               UNIDAD_ID & "," & FORMULA_ID_REL & "," & ES_SOLUCION & ")"
    execute_bd consulta
    InsertarCamposFormula = ID_CAMPO
    Exit Function
fallo:
    InsertarCamposFormula = 0
    MsgBox "Error al insertar los formulas_campos (InsertarCamposFormula)", vbCritical, Err.Description
End Function

Public Function lista() As ADODB.RecordSet
    Dim consulta As String
    consulta = "SELECT a.id_campo as campo, a.nombre as nombre " & _
            " FROM formulas_campos a, formulas b" & _
            " WHERE a.formula_id=b.id_formula " & _
            " order by requerido desc, id_campo asc"
    Set lista = datos_bd(consulta)
End Function
Public Function Eliminar(FORMULA As Long) As Boolean
    On Error GoTo fallo
    execute_bd "DELETE FROM formulas_campos WHERE FORMULA_ID = " & FORMULA
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar formulas_campos.", vbCritical, Err.Description
End Function

Public Function CampoConUnidad(CAMPO As Long) As ADODB.RecordSet
    Dim consulta As String
    consulta = "SELECT a.nombre, u.nombre, a.id_campo, a.es_solucion, a.enteros,a.decimales " & _
               "  FROM formulas_campos a, unidades u " & _
               " WHERE a.id_campo=" & CAMPO & _
               "   AND a.unidad_id = u.id_unidad"
    Set CampoConUnidad = datos_bd(consulta)
End Function


