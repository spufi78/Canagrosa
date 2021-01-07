VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRPR_Componentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private TIPO_REACTIVO_PR_ID As Integer
Private ORDEN As Integer
Private TIPO_REACTIVO_EX_ID As Integer
Private TIPO As String
Private PROCEDIMIENTO_REFERENCIA As String
Private cantidad As String
Private UNIDAD_ID As Integer

Public Property Let setTIPO_REACTIVO_PR_ID(ByVal dato As Integer)
    TIPO_REACTIVO_PR_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Integer)
    ORDEN = Trim(dato)
End Property
Public Property Let setTIPO_REACTIVO_EX_ID(ByVal dato As Integer)
    TIPO_REACTIVO_EX_ID = dato
End Property
Public Property Let setTIPO(ByVal dato As String)
    TIPO = dato
End Property
Public Property Let setPROCEDIMIENTO_REFERENCIA(ByVal dato As String)
    PROCEDIMIENTO_REFERENCIA = dato
End Property
Public Property Let setCANTIDAD(ByVal dato As String)
    cantidad = dato
End Property
Public Property Let setUNIDAD_ID(ByVal dato As Integer)
    UNIDAD_ID = dato
End Property
Public Property Get getTIPO_REACTIVO_PR_ID() As Integer
    getTIPO_REACTIVO_PR_ID = TIPO_REACTIVO_PR_ID
End Property
Public Property Get getORDEN() As Integer
    getORDEN = ORDEN
End Property
Public Property Get getTIPO_REACTIVO_EX_ID() As Integer
    getTIPO_REACTIVO_EX_ID = TIPO_REACTIVO_EX_ID
End Property
Public Property Get getTIPO() As String
    getTIPO = TIPO
End Property
Public Property Get getPROCEDIMIENTO_REFERENCIA() As String
    getPROCEDIMIENTO_REFERENCIA = PROCEDIMIENTO_REFERENCIA
End Property
Public Property Get getCANTIDAD() As String
    getCANTIDAD = cantidad
End Property
Public Property Get getUNIDAD_ID() As Integer
    getUNIDAD_ID = UNIDAD_ID
End Property
Public Function Listado() As ADODB.RecordSet
    Dim consulta As String
    consulta = "SELECT * FROM RPR_COMPONENTES order by NOMBRE"
    Set Listado = datos_bd(consulta)
End Function
Public Function Insertar() As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "Insert into RPR_COMPONENTES " & _
               " values(" & _
               TIPO_REACTIVO_PR_ID & "," & _
               TIPO_REACTIVO_EX_ID & ", " & _
               ORDEN & ",'" & TIPO & "'," & _
               "'" & PROCEDIMIENTO_REFERENCIA & "'," & _
               "'" & cantidad & "'," & _
               UNIDAD_ID & _
               ")"
    execute_bd consulta
    Insertar = True
    Exit Function
fallo:
    Insertar = False
    MsgBox "Error al insertar los componentes (Insertar)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.RecordSet
    Dim consulta As String
    consulta = "DELETE FROM RPR_COMPONENTES WHERE TIPO_REACTIVO_PR_ID=" & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function Componentes(ID As Long) As ADODB.RecordSet
    Dim consulta As String
    consulta = "select b.nombre,a.TIPO_REACTIVO_EX_ID,a.PROCEDIMIENTO_REFERENCIA,a.CANTIDAD,c.NOMBRE,a.UNIDAD_ID,a.TIPO,a.orden " & _
            " from rpr_componentes a " & _
            " left join tipos_reactivo_ex b on a.TIPO_REACTIVO_EX_ID = b.ID_TIPO_REACTIVO_EX " & _
            " left join unidades c on a.UNIDAD_ID = c.ID_UNIDAD " & _
            " Where a.TIPO_REACTIVO_PR_ID = " & ID & _
            " and a.TIPO = 'E' " & _
            " UNION "
    consulta = consulta & "select b.nombre,a.TIPO_REACTIVO_EX_ID,a.PROCEDIMIENTO_REFERENCIA,a.CANTIDAD,c.NOMBRE,a.UNIDAD_ID,a.TIPO,a.orden " & _
            " from rpr_componentes a " & _
            " left join rpr_tipos b on a.TIPO_REACTIVO_EX_ID = b.ID_TIPO_REACTIVO_PR " & _
            " left join unidades c on a.UNIDAD_ID = c.ID_UNIDAD " & _
            " Where a.TIPO_REACTIVO_PR_ID = " & ID & _
            " and a.TIPO = 'I' " & _
            " order by 8"
    Set Componentes = datos_bd(consulta)
End Function
