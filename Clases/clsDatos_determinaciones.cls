VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatos_determinaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE formulas_campos      *
'************************************************************
Private DETERMINACION_ID As Long
Private CAMPO_ID As Integer
Private VALOR_1 As String
Private VALOR_2 As String

Public Property Let setDETERMINACION_ID(ByVal dato As Long)
    DETERMINACION_ID = dato
End Property
Public Property Let setCAMPO_ID(ByVal dato As Integer)
    CAMPO_ID = dato
End Property
Public Property Let setVALOR_1(ByVal dato As String)
    VALOR_1 = Trim(dato)
End Property
Public Property Let setVALOR_2(ByVal dato As String)
    VALOR_2 = Trim(dato)
End Property
Public Property Get getDETERMINACION_ID() As Long
    getDETERMINACION_ID = DETERMINACION_ID
End Property
Public Property Get getCAMPO_ID() As Integer
    getCAMPO_ID = CAMPO_ID
End Property
Public Property Get getVALOR_1() As String
    getVALOR_1 = VALOR_1
End Property
Public Property Get getVALOR_2() As String
    getVALOR_2 = VALOR_2
End Property
Public Function Insertar() As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "Insert into datos_determinaciones " & _
               " values(" & _
               DETERMINACION_ID & "," & CAMPO_ID & ",'" & VALOR_1 & "','" & _
               VALOR_2 & "')"
    execute_bd consulta
    Insertar = True
    Exit Function
fallo:
    Insertar = False
    MsgBox "Error al insertar los datos_determinaciones (clsDatos_determinaciones", vbCritical, Err.Description
End Function
Public Function CARGAR(deter As Long, CAMPO As Integer) As Boolean
    Dim consulta As String
    Dim rs As New ADODB.RecordSet
    consulta = "SELECT * FROM datos_determinaciones WHERE DETERMINACION_ID=" & deter & " AND CAMPO_ID=" & CAMPO
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
         CARGAR = False
    Else
        DETERMINACION_ID = rs("DETERMINACION_ID")
        CAMPO_ID = rs("CAMPO_ID")
        VALOR_1 = rs("VALOR_1")
        VALOR_2 = rs("VALOR_2")
    End If
    Set rs = Nothing
    CARGAR = True
    Exit Function
fallo:
    CARGAR = False
    MsgBox "Error al cargar los datos_determinacion", vbCritical, Err.Description
End Function

Public Function Insertar_Valores() As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "update datos_determinaciones " & _
               " set " & _
               " VALOR_1 = '" & VALOR_1 & "'," & _
               " VALOR_2 = '" & VALOR_2 & "'" & _
               " WHERE DETERMINACION_ID = " & DETERMINACION_ID & _
               "   AND CAMPO_ID = " & CAMPO_ID
    execute_bd consulta
'    If VALOR_1 <> "" And VALOR_1 <> "I-1" Then
    If VALOR_1 <> "" And VALOR_1 <> "I-1" Then
        Dim oDH As New clsDeterminaciones_historico
        With oDH
            .setDETERMINACION_ID = DETERMINACION_ID
            .setCAMPO_ID = CAMPO_ID
            .setVALOR_1 = VALOR_1
'            If VALOR_2 <> "" And VALOR_2 <> "I-2" Then
            If VALOR_2 <> "" Then
                .setVALOR_2 = VALOR_2
            Else
                .setVALOR_2 = ""
            End If
            .Insertar
        End With
    End If
    Insertar_Valores = True
    Exit Function
fallo:
    Insertar_Valores = False
    MsgBox "Error al Insertar_Valores (clsDatos_determinaciones", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.RecordSet
    Dim consulta As String
    consulta = "DELETE FROM DATOS_DETERMINACIONES WHERE DETERMINACION_ID=" & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function cargar_determinacion(deter As Long) As ADODB.RecordSet
    Dim consulta As String
    consulta = "SELECT * FROM datos_determinaciones WHERE DETERMINACION_ID=" & deter & " order by determinacion_id asc"
    Set cargar_determinacion = datos_bd(consulta)
End Function

Public Function Cargar_Campo(CAMPO As Integer) As Boolean
    Dim consulta As String
    Dim rs As New ADODB.RecordSet
    consulta = "SELECT * FROM datos_determinaciones WHERE CAMPO_ID=" & CAMPO
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        Cargar_Campo = False
    Else
        Cargar_Campo = True
    End If
    Set rs = Nothing
    Exit Function
fallo:
    Cargar_Campo = False
    MsgBox "Error al cargar el campo.", vbCritical, Err.Description
End Function

