VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGenericClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ID As Long
Private NOMBRE As String
Private DESCRIPCION As String
Private OBSERVACIONES As String
Private ID_AUX As Integer


Public Property Get getID() As Long
   getID = ID
End Property

Public Property Get getID_AUX() As Integer
    getID_AUX = ID_AUX
End Property


Public Property Let setID(ByVal lngID As Long)
    ID = lngID
End Property

Public Property Get getDESCRIPCION() As String
    getDESCRIPCION = DESCRIPCION
End Property

Public Property Let setDESCRIPCION(ByVal strDESCRIPCION As String)
    DESCRIPCION = strDESCRIPCION
End Property

Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property

Public Property Let setID_AUX(ByVal valor As Integer)
    ID_AUX = valor
End Property

Public Property Let setNOMBRE(ByVal strNOMBRE As String)
    NOMBRE = strNOMBRE
End Property


Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
End Property

Public Property Let setOBSERVACIONES(ByVal strOBSERVACIONES As String)
    OBSERVACIONES = strOBSERVACIONES
End Property

