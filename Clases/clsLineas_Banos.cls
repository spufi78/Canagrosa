VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLineas_Banos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private LINEA_ID As Integer
Private BANO_ID As Integer
Private documento As Integer
Private ORDEN As Integer
Public Property Let setLINEA_ID(ByVal dato As Integer)
    LINEA_ID = dato
End Property
Public Property Let setBANO_ID(ByVal dato As Integer)
    BANO_ID = Trim(dato)
End Property
Public Property Let setDOCUMENTO(ByVal dato As Integer)
    documento = Trim(dato)
End Property
Public Property Let setORDEN(ByVal dato As Integer)
    ORDEN = Trim(dato)
End Property
Public Property Get getLINEA_ID() As Integer
    getLINEA_ID = LINEA_ID
End Property
Public Property Get getBANO_ID() As Integer
    getBANO_ID = BANO_ID
End Property
Public Property Get getDOCUMENTO() As Integer
    getDOCUMENTO = documento
End Property
Public Property Get getORDEN() As Integer
    getORDEN = ORDEN
End Property
Public Function Buscar_Bano(BANO As Integer) As Integer
    Dim rs As New ADODB.RecordSet
    Dim consulta As String
    consulta = "SELECT documento FROM LINEAS_BANOS WHERE BANO_ID=" & BANO
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        Buscar_Bano = 0
    Else
        Buscar_Bano = rs("documento")
    End If
End Function
Public Function Buscar_Documento(documento As Integer) As ADODB.RecordSet
    Dim consulta As String
    consulta = "SELECT * FROM LINEAS_BANOS " & _
               " WHERE DOCUMENTO=" & documento & _
               " ORDER BY ORDEN"
    Set Buscar_Documento = datos_bd(consulta)
End Function
