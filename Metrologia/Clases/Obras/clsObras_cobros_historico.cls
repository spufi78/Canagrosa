VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObras_cobros_historico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSOBRAS_COBROS_HISTORICO
'***************************************************************
Private OBRA_ID As Long
Private orden As Long
Private fecha As String
Private OBSERVACIONES As String
Private FS As String
Private USUARIO_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSOBRAS_COBROS_HISTORICO
'***************************************************************
Public Property Let setOBRA_ID(ByVal dato As Long)
        OBRA_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        orden = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setFS(ByVal dato As String)
        FS = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSOBRAS_COBROS_HISTORICO
'***************************************************************
Public Property Get getOBRA_ID() As Long
        getOBRA_ID = OBRA_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = orden
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getFS() As String
        getFS = FS
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSOBRAS_COBROS_HISTORICO
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM OBRAS_COBROS_HISTORICO WHERE OBRA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                OBRA_ID = rs("OBRA_ID")
                orden = rs("ORDEN")
                fecha = rs("FECHA")
                OBSERVACIONES = rs("OBSERVACIONES")
                FS = rs("FS")
                USUARIO_ID = rs("USUARIO_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsObras_cobros_historico)", vbCritical, Err.Description
End Function
Public Function CrearORDEN(OBRA As Long)
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ORDEN) FROM OBRAS_COBROS_HISTORICO WHERE OBRA_ID = " & OBRA
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                orden = 1
        Else
                orden = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearORDEN (OBRA_ID)
        consulta = "INSERT INTO OBRAS_COBROS_HISTORICO " & _
                   "  VALUES (" & _
                        OBRA_ID & "," & orden & "," & "'" & fecha & "'" & "," & _
                        "'" & OBSERVACIONES & "'" & "," & "CURRENT_TIMESTAMP," & _
                        USUARIO_ID & _
                ")"
        execute_bd consulta
        Insertar = OBRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsObras_cobros_historico)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE OBRAS_COBROS_HISTORICO SET " & _
                        " ORDEN = " & orden & "," & _
                        " FECHA = '" & fecha & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " FS = '" & FS & "'," & _
                        " USUARIO_ID = " & USUARIO_ID & "" & _
                " WHERE OBRA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsObras_cobros_historico)", vbCritical, Err.Description
End Function
Public Function Eliminar(OBRA As Long, ORD As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM OBRAS_COBROS_HISTORICO " & _
                "    WHERE OBRA_ID = " & OBRA & " AND ORDEN = " & ORD
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsObras_cobros_historico)", vbCritical, Err.Description
End Function
Public Function Listado(OBRA As Long) As ADODB.Recordset
        Dim consulta As String
   On Error GoTo Listado_Error

        consulta = "SELECT O.ORDEN, O.FECHA,O.OBSERVACIONES, U.USUARIO " & _
                   "  FROM OBRAS_COBROS_HISTORICO O, USUARIOS U " & _
                   " WHERE O.USUARIO_ID = U.ID_EMPLEADO " & _
                   "   AND O.OBRA_ID = " & OBRA & _
                   " ORDER BY O.ORDEN DESC"
        Set Listado = datos_bd(consulta)

   On Error GoTo 0
   Exit Function

Listado_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Listado of Módulo de clase clsObras_cobros_historico"
End Function
