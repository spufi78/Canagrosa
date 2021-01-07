VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSellantes_determinaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSSELLANTES_DETERMINACIONES
'***************************************************************
Private MUESTRA_ID As Long
Private ORDEN As Long
Private CAMPO_ID As Long
Private VALOR_1 As String
Private VALOR_2 As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSSELLANTES_DETERMINACIONES
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
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
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSSELLANTES_DETERMINACIONES
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
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
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSSELLANTES_DETERMINACIONES
'***************************************************************

Public Function Carga(lMUESTRA As Long, lORDEN As Long, lCAMPO As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SELLANTES_DETERMINACIONES " & _
                   " WHERE MUESTRA_ID = " & lMUESTRA & _
                   "   AND ORDEN = " & lORDEN & _
                   "   AND CAMPO_ID = " & lCAMPO
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                VALOR_1 = rs("VALOR_1")
                VALOR_2 = rs("VALOR_2")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsSellantes_determinaciones)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO SELLANTES_DETERMINACIONES " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & ORDEN & "," & CAMPO_ID & "," & _
                        "'" & VALOR_1 & "'" & "," & "'" & VALOR_2 & "'" & _
                ")"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsSellantes_determinaciones)", vbCritical, Err.Description
End Function
Public Function Eliminar(MUESTRA As Long, ORDEN As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM SELLANTES_DETERMINACIONES " & _
                "    WHERE MUESTRA_ID = " & MUESTRA & " AND ORDEN = " & ORDEN
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsSellantes_determinaciones)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SELLANTES_DETERMINACIONES ORDER BY MUESTRA_ID"
        Set Listado = datos_bd(consulta)
End Function
'M0926-I
Public Function Listado_Muestra(idMuestra As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SELLANTES_DETERMINACIONES WHERE MUESTRA_ID=" & idMuestra
        Set Listado_Muestra = datos_bd(consulta)
End Function
'M0926-F
