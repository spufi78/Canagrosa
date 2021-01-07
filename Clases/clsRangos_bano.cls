VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRangos_bano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE campos_formula      *
'************************************************************
Private BANO_ID As Long
Private TIPO_DETERMINACION_ID As Long
Private ORDEN As Integer
Private MINIMO As String
Private MAXIMO As String
Private DIF_MINIMA As String
Private DIF_MAXIMA As String
Public Property Let setBANO_ID(ByVal dato As Long)
    BANO_ID = dato
End Property
Public Property Let setTIPO_DETERMINACION_ID(ByVal dato As Long)
    TIPO_DETERMINACION_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Integer)
    ORDEN = dato
End Property
Public Property Let setMINIMO(ByVal dato As String)
    MINIMO = Trim(dato)
End Property
Public Property Let setMAXIMO(ByVal dato As String)
    MAXIMO = Trim(dato)
End Property
Public Property Let setDIF_MINIMA(ByVal dato As String)
    DIF_MINIMA = Trim(dato)
End Property
Public Property Let setDIF_MAXIMA(ByVal dato As String)
    DIF_MAXIMA = Trim(dato)
End Property
Public Property Get getBANO_ID() As Integer
    getBANO_ID = BANO_ID
End Property
Public Property Get getTIPO_DETERMINACION_ID() As Integer
    getTIPO_DETERMINACION_ID = TIPO_DETERMINACION_ID
End Property
Public Property Get getORDEN() As Integer
    getORDEN = ORDEN
End Property
Public Property Get getMINIMO() As String
    getMINIMO = MINIMO
End Property
Public Property Get getMAXIMO() As String
    getMAXIMO = MAXIMO
End Property
Public Property Get getDIF_MINIMA() As String
    getDIF_MINIMA = DIF_MINIMA
End Property
Public Property Get getDIF_MAXIMA() As String
    getDIF_MAXIMA = DIF_MAXIMA
End Property
Public Function Insertar() As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "Insert into rangos_bano " & _
               " values(" & _
               BANO_ID & "," & TIPO_DETERMINACION_ID & "," & ORDEN & ",'" & MINIMO & "','" & _
               MAXIMO & "','" & DIF_MINIMA & "','" & DIF_MAXIMA & "')"
    execute_bd consulta
    Insertar = True
    Exit Function
fallo:
    Insertar = False
    MsgBox "Error al insertar los rangos del baño (Insertar)", vbCritical, Err.Description
End Function
Public Function Cargar(bano As Long, deter As Long) As Boolean
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "SELECT * FROM rangos_bano WHERE BANO_ID=" & bano & " AND TIPO_DETERMINACION_ID=" & deter
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        Cargar = False
    Else
        BANO_ID = rs("BANO_ID")
        TIPO_DETERMINACION_ID = rs("TIPO_DETERMINACION_ID")
        ORDEN = rs("ORDEN")
        MINIMO = rs("MINIMO")
        MAXIMO = rs("MAXIMO")
        DIF_MINIMA = rs("DIF_MINIMA")
        DIF_MAXIMA = rs("DIF_MAXIMA")
    End If
    Set rs = Nothing
    Cargar = True
    Exit Function
fallo:
    Cargar = False
    MsgBox "Error al cargar los rangos baño", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "DELETE FROM rangos_bano WHERE BANO_ID=" & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function determinaciones_bano(ID As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM RANGOS_BANO WHERE BANO_ID = " & ID & " ORDER BY ORDEN"
    Set determinaciones_bano = datos_bd(consulta)
End Function
Public Function Listado_determinaciones_bano(ID As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT TD.NOMBRE, TD.ID_TIPO_DETERMINACION " & _
               "  FROM RANGOS_BANO RB, TIPOS_DETERMINACION TD " & _
               " WHERE RB.BANO_ID = " & ID & _
               "   AND RB.TIPO_DETERMINACION_ID = TD.ID_TIPO_DETERMINACION " & _
               " ORDER BY RB.ORDEN"
    Set Listado_determinaciones_bano = datos_bd(consulta)
End Function
