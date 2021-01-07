VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArticulos_Escandallo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*****************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLIENTE      *
'*****************************************************
Private ARTICULO_ID As Long
Private ARTICULO_ID_ESCANDALLO As Integer
Private CANTIDAD As Integer
Private ORDEN As Integer
Public Property Let setARTICULO_ID(ByVal dato As Long)
    ARTICULO_ID = dato
End Property
Public Property Let setARTICULO_ID_ESCANDALLO(ByVal dato As Integer)
    ARTICULO_ID_ESCANDALLO = dato
End Property
Public Property Let setCANTIDAD(ByVal dato As Integer)
    CANTIDAD = Trim(dato)
End Property
Public Property Let setORDEN(ByVal dato As Integer)
    ORDEN = Trim(dato)
End Property
Public Property Get getARTICULO_ID() As Long
    getARTICULO_ID = ARTICULO_ID
End Property
Public Property Get getARTICULO_ID_ESCANDALLO() As Integer
    getARTICULO_ID_ESCANDALLO = ARTICULO_ID_ESCANDALLO
End Property
Public Property Get getCANTIDAD() As Integer
    getCANTIDAD = CANTIDAD
End Property
Public Property Get getORDEN() As Integer
    getORDEN = ORDEN
End Property
Public Function Insertar() As Long
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "INSERT INTO ARTICULOS_ESCANDALLO" & _
               " VALUES(" & _
               ARTICULO_ID & "," & ARTICULO_ID_ESCANDALLO & "," & CANTIDAD & "," & _
               ORDEN & ")"
    execute_bd consulta
    Insertar = ARTICULO_ID
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el Articulo_Escandallo (Insertar)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    consulta = "DELETE FROM ARTICULOS_ESCANDALLO " & _
               " WHERE " & _
               " ARTICULO_ID = " & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar el artículo.", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.DESCRIPCION,AE.CANTIDAD,A.ID_ARTICULO,A.STOCK,AE.ARTICULO_ID_ESCANDALLO " & _
                   "  FROM ARTICULOS_ESCANDALLO AE,ARTICULOS A " & _
                   " WHERE AE.ARTICULO_ID = " & ID & _
                   "   AND AE.ARTICULO_ID_ESCANDALLO = A.ID_ARTICULO " & _
                   " ORDER BY AE.ORDEN"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_duplicar(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ARTICULO_ID_ESCANDALLO,CANTIDAD,ORDEN " & _
                   "  FROM ARTICULOS_ESCANDALLO " & _
                   " WHERE ARTICULO_ID = " & ID & _
                   " ORDER BY ORDEN"
        Set Listado_duplicar = datos_bd(consulta)
End Function


