VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecargas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private MUESTRA_ID As Long
Private OR_RUTA As String
Private OR_EMPLEADO_ID As Integer
Private OR_FECHA As String
Private OR_HORA As String
Private RR_RUTA As String
Private RR_EMPLEADO_ID As Integer
Private RR_FECHA As String
Private RR_HORA As String
Private AN_RUTA As String
Private AN_EMPLEADO_ID As Integer
Private AN_FECHA As String
Private AN_HORA As String
Public Property Let setMUESTRA_ID(ByVal dato As Long)
    MUESTRA_ID = dato
End Property
Public Property Let setOR_RUTA(ByVal dato As String)
    OR_RUTA = dato
End Property
Public Property Let setOR_EMPLEADO_ID(ByVal dato As Integer)
    OR_EMPLEADO_ID = Trim(dato)
End Property
Public Property Let setOR_FECHA(ByVal dato As String)
    OR_FECHA = dato
End Property
Public Property Let setOR_HORA(ByVal dato As String)
    OR_HORA = Trim(dato)
End Property
Public Property Let setRR_RUTA(ByVal dato As String)
    RR_RUTA = dato
End Property
Public Property Let setRR_EMPLEADO_ID(ByVal dato As Integer)
    RR_EMPLEADO_ID = Trim(dato)
End Property
Public Property Let setRR_FECHA(ByVal dato As String)
    RR_FECHA = dato
End Property
Public Property Let setRR_HORA(ByVal dato As String)
    RR_HORA = Trim(dato)
End Property
Public Property Let setAN_RUTA(ByVal dato As String)
    AN_RUTA = dato
End Property
Public Property Let setAN_EMPLEADO_ID(ByVal dato As Integer)
    AN_EMPLEADO_ID = Trim(dato)
End Property
Public Property Let setAN_FECHA(ByVal dato As String)
    AN_FECHA = dato
End Property
Public Property Let setAN_HORA(ByVal dato As String)
    AN_HORA = Trim(dato)
End Property
Public Property Get getMUESTRA_ID() As Long
    getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getOR_RUTA() As String
    getOR_RUTA = OR_RUTA
End Property
Public Property Get getOR_EMPLEADO_ID() As Integer
    getOR_EMPLEADO_ID = OR_EMPLEADO_ID
End Property
Public Property Get getOR_FECHA() As String
    getOR_FECHA = OR_FECHA
End Property
Public Property Get getOR_HORA() As String
    getOR_HORA = OR_HORA
End Property
Public Property Get getRR_RUTA() As String
    getRR_RUTA = RR_RUTA
End Property
Public Property Get getRR_EMPLEADO_ID() As Integer
    getRR_EMPLEADO_ID = RR_EMPLEADO_ID
End Property
Public Property Get getRR_FECHA() As String
    getRR_FECHA = RR_FECHA
End Property
Public Property Get getRR_HORA() As String
    getRR_HORA = RR_HORA
End Property
Public Property Get getAN_RUTA() As String
    getAN_RUTA = AN_RUTA
End Property
Public Property Get getAN_EMPLEADO_ID() As Integer
    getAN_EMPLEADO_ID = AN_EMPLEADO_ID
End Property
Public Property Get getAN_FECHA() As String
    getAN_FECHA = AN_FECHA
End Property
Public Property Get getAN_HORA() As String
    getAN_HORA = AN_HORA
End Property
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM RECARGAS order by MUESTRA_ID DESC"
    Set Listado = datos_bd(consulta)
End Function
Public Function Cargar(id As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM RECARGAS WHERE MUESTRA_ID=" & id
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        MUESTRA_ID = rs("MUESTRA_ID")
        OR_RUTA = rs("OR_RUTA")
        OR_EMPLEADO_ID = rs("OR_EMPLEADO_ID")
        OR_FECHA = rs("OR_FECHA")
        OR_HORA = rs("OR_HORA")
        RR_RUTA = rs("RR_RUTA")
        RR_EMPLEADO_ID = rs("RR_EMPLEADO_ID")
        RR_FECHA = rs("RR_FECHA")
        RR_HORA = rs("RR_HORA")
        AN_RUTA = rs("AN_RUTA")
        AN_EMPLEADO_ID = rs("AN_EMPLEADO_ID")
        AN_FECHA = rs("AN_FECHA")
        AN_HORA = rs("AN_HORA")
        Cargar = True
    Else
        Cargar = False
    End If
    Exit Function
fallo:
    Cargar = False
    MsgBox "Error al cargar la recarga.", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    consulta = "Insert into RECARGAS " & _
               " values(" & _
               MUESTRA_ID & ",'" & _
               OR_RUTA & "'," & OR_EMPLEADO_ID & ",'" & OR_FECHA & "','" & OR_HORA & "','" & _
               RR_RUTA & "'," & RR_EMPLEADO_ID & ",'" & RR_FECHA & "','" & RR_HORA & "','" & _
               AN_RUTA & "'," & AN_EMPLEADO_ID & ",'" & AN_FECHA & "','" & AN_HORA & "')"
    execute_bd consulta
    Insertar = MUESTRA_ID
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar la Recarga (Insertar)", vbCritical, Err.Description
End Function
Public Function Modificar(id As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE RECARGAS SET " & _
               " OR_RUTA = '" & OR_RUTA & "'," & _
               " OR_EMPLEADO_ID = " & OR_EMPLEADO_ID & "," & _
               " OR_FECHA = '" & OR_FECHA & "'," & _
               " OR_HORA = '" & OR_HORA & "'," & _
               " RR_RUTA = '" & RR_RUTA & "'," & _
               " RR_EMPLEADO_ID = " & RR_EMPLEADO_ID & "," & _
               " RR_FECHA = '" & RR_FECHA & "'," & _
               " RR_HORA = '" & RR_HORA & "'," & _
               " AN_RUTA = '" & AN_RUTA & "'," & _
               " AN_EMPLEADO_ID = " & AN_EMPLEADO_ID & "," & _
               " AN_FECHA = '" & AN_FECHA & "'," & _
               " AN_HORA = '" & AN_HORA & "'" & _
               " WHERE MUESTRA_ID = " & id
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al Modificar la Recarga (Modificar)", vbCritical, Err.Description
End Function
