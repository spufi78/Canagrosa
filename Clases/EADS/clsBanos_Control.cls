VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBanos_Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_CONTROL As Long
Private BANO_ID As Long
Private PLANTILLA As String
Public Property Let setID_CONTROL(ByVal dato As Long)
    ID_CONTROL = dato
End Property
Public Property Let setBANO_ID(ByVal dato As Long)
    BANO_ID = dato
End Property
Public Property Let setPLANTILLA(ByVal dato As String)
    PLANTILLA = dato
End Property
Public Property Get getID_CONTROL() As Integer
    getID_CONTROL = ID_CONTROL
End Property
Public Property Get getBANO_ID() As Integer
    getBANO_ID = BANO_ID
End Property
Public Property Get getPLANTILLA() As String
    getPLANTILLA = PLANTILLA
End Property
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM BANOS_CONTROL WHERE ID_CONTROL = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                BANO_ID = rs("BANO_ID")
                PLANTILLA = rs("PLANTILLA")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsAlodine)", vbCritical, Err.Description
End Function
Public Function Carga_por_BANO(BANO As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM BANOS_CONTROL WHERE BANO_ID = " & BANO
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga_por_BANO = False
        Else
                BANO_ID = rs("BANO_ID")
                PLANTILLA = rs("PLANTILLA")
                Carga_por_BANO = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga_por_BANO = False
        MsgBox "Error al cargar los datos(clsAlodine)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT MAX(ID_CONTROL) FROM BANOS_CONTROL"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_CONTROL = 1
        Else
                ID_CONTROL = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function

Public Function Insertar() As Boolean
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearID
    consulta = "Insert into BANOS_CONTROL " & _
               " values(" & _
               ID_CONTROL & "," & BANO_ID & ",'" & PLANTILLA & "')"
    execute_bd consulta
    Insertar = True
    Exit Function
fallo:
    Insertar = False
    MsgBox "Error al insertar los datos del baño (BANOS_CONTROL)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.RecordSet
    Dim consulta As String
    consulta = "DELETE FROM BANOS_CONTROL WHERE ID_CONTROL=" & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM BANOS_CONTROL ORDER BY ID_CONTROL"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT BANO_ID,NOMBRE " & _
                   " FROM BANOS_CONTROL BC, BANOS B " & _
                   " WHERE BC.BANO_ID = B.ID_BANO " & _
                   " ORDER BY B.NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
