VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEntidades_Entrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_ENTIDAD_ENTREGA As Integer
Private DESCRIPCION As String
Public Property Let setID_ENTIDAD_ENTREGA(ByVal dato As Integer)
    ID_ENTIDAD_ENTREGA = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
    DESCRIPCION = Trim(dato)
End Property
Public Property Get getID_ENTIDAD_ENTREGA() As Integer
    getID_ENTIDAD_ENTREGA = ID_ENTIDAD_ENTREGA
End Property
Public Property Get getDESCRIPCION() As String
    getDESCRIPCION = DESCRIPCION
End Property
Public Function CargarEntidad(ByVal entidad As Integer) As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "select * from entidades_entrega where ID_ENTIDAD_ENTREGA = " & entidad
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        CargarEntidad = False
    Else
        CargarEntidad = True
        ID_ENTIDAD_ENTREGA = rs("ID_ENTIDAD_ENTREGA")
        DESCRIPCION = rs("DESCRIPCION")
    End If
    Set rs = Nothing
    Exit Function
fallo:
    CargarEntidad = False
    MsgBox "Error al cargar la entidad de entrada (clsCargarEntidad)", vbCritical, Err.Description
End Function

Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT descripcion,id_entidad_entrega FROM entidades_entrega order by descripcion"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT id_entidad_entrega,descripcion FROM entidades_entrega order by descripcion"
    Set Listado_Combo = datos_bd(consulta)
End Function


