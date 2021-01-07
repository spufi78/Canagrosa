VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDeterminaciones_historico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSDETERMINACIONES_HISTORICO
'***************************************************************
Private DETERMINACION_ID As Long
Private ORDEN As Long
Private CAMPO_ID As Long
Private VALOR_1 As String
Private VALOR_2 As String
Private fecha As String
Private hora As String
Private EMPLEADO_ID As Long
Private OBSERVACIONES As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSDETERMINACIONES_HISTORICO
'***************************************************************
Public Property Let setDETERMINACION_ID(ByVal dato As Long)
        DETERMINACION_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setCAMPO_ID(ByVal dato As Long)
        CAMPO_ID = dato
End Property
Public Property Let setVALOR_1(ByVal dato As String)
        VALOR_1 = dato
End Property
Public Property Let setVALOR_2(ByVal dato As String)
        VALOR_2 = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setHORA(ByVal dato As String)
        hora = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Long)
        EMPLEADO_ID = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSDETERMINACIONES_HISTORICO
'***************************************************************
Public Property Get getDETERMINACION_ID() As Long
        getDETERMINACION_ID = DETERMINACION_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getCAMPO_ID() As Long
        getCAMPO_ID = CAMPO_ID
End Property
Public Property Get getVALOR_1() As String
        getVALOR_1 = VALOR_1
End Property
Public Property Get getVALOR_2() As String
        getVALOR_2 = VALOR_2
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getHORA() As String
        getHORA = hora
End Property
Public Property Get getEMPLEADO_ID() As Long
        getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSDETERMINACIONES_HISTORICO
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.RecordSet
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM DETERMINACIONES_HISTORICO WHERE DETERMINACION_ID = " & ID
        Set RS = datos_bd(CONSULTA)
        If RS.RecordCount = 0 Then
                Carga = False
        Else
                DETERMINACION_ID = RS("DETERMINACION_ID")
                ORDEN = RS("ORDEN")
                CAMPO_ID = RS("CAMPO_ID")
                VALOR_1 = RS("VALOR_1")
                VALOR_2 = RS("VALOR_2")
                fecha = RS("FECHA")
                hora = RS("HORA")
                EMPLEADO_ID = RS("EMPLEADO_ID")
                OBSERVACIONES = RS("OBSERVACIONES")
                Carga = True
        End If
        Set RS = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsDeterminaciones_historico)", vbCritical, Err.Description
End Function
Public Function CalcularOrden(DETERMINACION_ID As Long) As Integer
    Dim RS As New ADODB.RecordSet
    Dim CONSULTA As String
    CONSULTA = "SELECT MAX(ORDEN) FROM DETERMINACIONES_HISTORICO WHERE DETERMINACION_ID = " & DETERMINACION_ID
    Set RS = datos_bd(CONSULTA)
    If IsNull(RS.Fields(0)) Or (RS.EOF And RS.BOF) Then  'si es nulo No se recupero ninguno
        CalcularOrden = 1
    Else
        CalcularOrden = RS.Fields(0) + 1
    End If
    Set RS = Nothing
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim CONSULTA As String
        Dim RS As ADODB.RecordSet
        ' Varificamos antes de insertar, si el resultado del campo es
        ' distinto del anterior
        CONSULTA = "SELECT VALOR_1,VALOR_2 FROM DETERMINACIONES_HISTORICO " & _
                   " WHERE DETERMINACION_ID = " & DETERMINACION_ID & _
                   "   AND CAMPO_ID = " & CAMPO_ID & _
                   " ORDER BY ORDEN DESC"
        Set RS = datos_bd(CONSULTA)
        If RS.RecordCount > 0 Then
            If RS(0) = VALOR_1 And RS(1) = VALOR_2 Then
                Insertar = DETERMINACION_ID
                Exit Function
            End If
        End If
        Me.setEMPLEADO_ID = usuario.getID_EMPLEADO
        Me.setFECHA = Format(Date, "yyyy-mm-dd")
        Me.setHORA = Format(Time, "hh:mm")
        Me.setORDEN = Me.CalcularOrden(DETERMINACION_ID)
        CONSULTA = "INSERT INTO DETERMINACIONES_HISTORICO " & _
                   "  VALUES (" & _
                        DETERMINACION_ID & "," & ORDEN & "," & CAMPO_ID & "," & _
                        "'" & VALOR_1 & "'" & "," & "'" & VALOR_2 & "'" & "," & "'" & fecha & "'" & "," & _
                        "'" & hora & "'" & "," & EMPLEADO_ID & "," & "'" & OBSERVACIONES & "'" & _
                ")"
        execute_bd CONSULTA
        Insertar = DETERMINACION_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsDeterminaciones_historico)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE DETERMINACIONES_HISTORICO SET " & _
                        " ORDEN = " & ORDEN & "," & _
                        " CAMPO_ID = " & CAMPO_ID & "," & _
                        " VALOR_1 = '" & VALOR_1 & "'," & _
                        " VALOR_2 = '" & VALOR_2 & "'," & _
                        " FECHA = '" & fecha & "'," & _
                        " HORA = '" & hora & "'," & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & "," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'" & _
                " WHERE DETERMINACION_ID = " & ID
        execute_bd CONSULTA
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsDeterminaciones_historico)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "DELETE FROM DETERMINACIONES_HISTORICO " & _
                "    WHERE DETERMINACION_ID = " & ID
        execute_bd CONSULTA
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDeterminaciones_historico)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.RecordSet
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM DETERMINACIONES_HISTORICO ORDER BY DETERMINACION_ID"
        Set Listado = datos_bd(CONSULTA)
End Function
Public Function ModificarVida(lDETERMINACION_ID As Long, lORDEN As Long, lCAMPO_ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE DETERMINACIONES_HISTORICO SET " & _
                        " FECHA = '" & fecha & "'," & _
                        " HORA = '" & hora & "'," & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & _
                " WHERE DETERMINACION_ID = " & lDETERMINACION_ID & _
                "   AND ORDEN = " & lORDEN & _
                "   AND CAMPO_ID = " & lCAMPO_ID
        execute_bd CONSULTA
        ModificarVida = True
        Exit Function
fallo:
        ModificarVida = False
        MsgBox "Error al ModificarVida (clsDeterminaciones_historico)", vbCritical, Err.Description
End Function


