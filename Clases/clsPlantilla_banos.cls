VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlantilla_banos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE plantilla_banos      *
'*************************************************************
Private PLANTILLA_ID As Integer
Private BANO_ID As Integer
Private ORDEN_MUESTRA As Integer
Public Property Let setPLANTILLA_ID(ByVal dato As Integer)
    PLANTILLA_ID = dato
End Property
Public Property Let setBANO_ID(ByVal dato As Integer)
    BANO_ID = dato
End Property
Public Property Let setORDEN_MUESTRA(ByVal dato As Integer)
    ORDEN_MUESTRA = dato
End Property
Public Property Get getPLANTILLA_ID() As Integer
    getPLANTILLA_ID = PLANTILLA_ID
End Property
Public Property Get getBANO_ID() As Integer
    getBANO_ID = BANO_ID
End Property
Public Property Get getORDEN_MUESTRA() As Integer
    getORDEN_MUESTRA = ORDEN_MUESTRA
End Property
Public Function Listado_bano(lplantilla As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT b.id_bano,b.nombre " & _
               " FROM plantilla_banos a, banos b " & _
               " WHERE a.plantilla_id = " & lplantilla & _
               "   AND a.bano_id = b.id_bano " & _
               " ORDER BY a.orden_muestra"
    Set Listado_bano = datos_bd(consulta)
End Function
Public Function Insertar() As Integer
    Dim consulta As String
    On Error GoTo fallo:
    '**************** TABLA MUESTRAS *********************
    consulta = "INSERT INTO PLANTILLA_BANOS" & _
               " values(" & _
               PLANTILLA_ID & "," & BANO_ID & "," & ORDEN_MUESTRA & ")"
    execute_bd consulta
    Insertar = PLANTILLA_ID
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al guardar los baños de la plantilla en la bd : " & Err.Description, vbCritical, Err.Number
End Function


