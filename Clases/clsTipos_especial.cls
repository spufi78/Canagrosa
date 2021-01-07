VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_especial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_TIPO_ESPECIAL As Integer
Private nombre As String
Private DESCRIPCION As String
Private BANO As Integer
Public Property Let setID_TIPO_ESPECIAL(ByVal dato As Integer)
    ID_TIPO_ESPECIAL = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
    DESCRIPCION = Trim(dato)
End Property
Public Property Let setBANO(ByVal dato As Integer)
    BANO = dato
End Property
Public Property Get getID_TIPO_ESPECIAL() As Integer
    getID_TIPO_ESPECIAL = ID_TIPO_ESPECIAL
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getDESCRIPCION() As String
    getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getBANO() As Integer
    getBANO = BANO
End Property
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM tipos_especial order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_TIPO_ESPECIAL, NOMBRE FROM tipos_especial order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function
Public Function cargar(ID_TIPO As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM tipos_especial WHERE id_tipo_especial=" & ID_TIPO
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        ID_TIPO_ESPECIAL = rs("ID_TIPO_ESPECIAL")
        nombre = rs("NOMBRE")
        DESCRIPCION = rs("DESCRIPCION")
        BANO = rs("BANO")
        cargar = True
    Else
        cargar = False
    End If
    Exit Function
fallo:
    cargar = False
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "TIPOS_ESPECIAL"
        .setDESCRIPCION = "Tipos Especiales de Muestras"
        .setPK = "ID_TIPO_ESPECIAL"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = False
        .setFILTRO = filtro
        Set .FORMULARIO = FORMULARIO
    End With
End Function
