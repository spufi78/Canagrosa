VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFluidos_recepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSfluidos_recepcion
'***************************************************************
Private MUESTRA_ID As Long
Private NORMATIVA_APLICABLE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSfluidos_recepcion
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setNORMATIVA_APLICABLE(ByVal dato As String)
        NORMATIVA_APLICABLE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSfluidos_recepcion
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getNORMATIVA_APLICABLE() As String
        getNORMATIVA_APLICABLE = NORMATIVA_APLICABLE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSfluidos_recepcion
'***************************************************************
Public Function Carga(MUESTRA As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM fluidos_recepcion  WHERE MUESTRA_ID = " & MUESTRA
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                NORMATIVA_APLICABLE = rs("NORMATIVA_APLICABLE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsfluidos_recepcion)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO fluidos_recepcion  " & _
                   "  VALUES (" & _
                        MUESTRA_ID & ",'" & NORMATIVA_APLICABLE & "'" & _
                ") ON DUPLICATE KEY UPDATE NORMATIVA_APLICABLE = '" & NORMATIVA_APLICABLE & "'"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsfluidos_recepcion)", vbCritical, Err.Description
End Function
Public Function modificar(ID_MUESTRA As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE fluidos_recepcion " & _
                   "   SET NORMATIVA_APLICABLE = '" & NORMATIVA_APLICABLE & "'" & _
                   " WHERE MUESTRA_ID = " & ID_MUESTRA
        execute_bd consulta
        modificar = True
        Exit Function
fallo:
        modificar = False
        MsgBox "Error al insertar (clsfluidos_recepcion)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM fluidos_recepcion  " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsfluidos_recepcion)", vbCritical, Err.Description
End Function
