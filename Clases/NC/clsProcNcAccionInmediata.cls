VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcNcAccionInmediata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ID_ACCION_INMEDIATA As Long
Private DESCRIPCION As String
Private ID_AUX As Integer
Public Property Get getID_AUX() As Integer
    getID_AUX = ID_AUX
End Property

Public Property Let setID_AUX(ByVal intID_AUX As Integer)
    ID_AUX = intID_AUX
End Property

Public Property Get getID_ACCION_INMEDIATA() As Long

   On Error GoTo idAccionInmediata_Error

    getID_ACCION_INMEDIATA = ID_ACCION_INMEDIATA

   On Error GoTo 0
   Exit Property

idAccionInmediata_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure idAccionInmediata of Módulo de clase clsProcNcAccionInmediata"

End Property

Public Property Let setID_ACCION_INMEDIATA(ByVal lngidAccionInmediata As Long)

   On Error GoTo idAccionInmediata_Error

    ID_ACCION_INMEDIATA = lngidAccionInmediata

   On Error GoTo 0
   Exit Property

idAccionInmediata_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure idAccionInmediata of Módulo de clase clsProcNcAccionInmediata"

End Property

Public Property Get getDESCRIPCION() As String

   On Error GoTo DESCRIPCION_Error

    getDESCRIPCION = DESCRIPCION

   On Error GoTo 0
   Exit Property

DESCRIPCION_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Descripcion of Módulo de clase clsProcNcAccionInmediata"

End Property

Public Property Let setDESCRIPCION(ByVal strDESCRIPCION As String)

   On Error GoTo DESCRIPCION_Error

    DESCRIPCION = strDESCRIPCION

   On Error GoTo 0
   Exit Property

DESCRIPCION_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Descripcion of Módulo de clase clsProcNcAccionInmediata"

End Property
