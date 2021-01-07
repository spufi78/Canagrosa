VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEq_Estados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***********************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE EQ_ESTADOS           *
'***********************************************************
Private ID_ESTADO As String
Private NOMBRE As String
Public Property Let setID_ESTADO(ByVal dato As String)
    ID_ESTADO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Get getID_ESTADO() As String
    getID_ESTADO = ID_ESTADO
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EQ_ESTADOS order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function CARGAR(ESTADO As String) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EQ_ESTADOS WHERE ID_ESTADO='" & ESTADO & "'"
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        ID_ESTADO = rs("ID_ESTADO")
        NOMBRE = rs("NOMBRE")
        CARGAR = True
    Else
        CARGAR = False
    End If
    Exit Function
fallo:
    CARGAR = False
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_ESTADO, CONCAT(ID_ESTADO,' - ',NOMBRE) FROM EQ_ESTADOS " & _
               " ORDER BY ID_ESTADO"
    Set Listado_Combo = datos_bd(consulta)
End Function
