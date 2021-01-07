VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlodine_normas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSALODINE_NORMAS
'***************************************************************
Private ALODINE_ID As Long
Private NORMA_ID As Long
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSALODINE_NORMAS
'***************************************************************
Public Property Let setALODINE_ID(ByVal dato As Long)
        ALODINE_ID = dato
End Property
Public Property Let setNORMA_ID(ByVal dato As Long)
        NORMA_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSALODINE_NORMAS
'***************************************************************
Public Property Get getALODINE_ID() As Long
        getALODINE_ID = ALODINE_ID
End Property
Public Property Get getNORMA_ID() As Long
        getNORMA_ID = NORMA_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSALODINE_NORMAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ALODINE_NORMAS WHERE ALODINE_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ALODINE_ID = rs("ALODINE_ID")
                NORMA_ID = rs("NORMA_ID")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsALODINE_NORMAS)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO ALODINE_NORMAS " & _
                   "  VALUES (" & _
                        ALODINE_ID & "," & NORMA_ID & "," & _
                        ORDEN & _
                ")"
        execute_bd consulta
        Insertar = ALODINE_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsALODINE_NORMAS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ALODINE_NORMAS " & _
                "    WHERE ALODINE_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsALODINE_NORMAS)", vbCritical, Err.Description
End Function
Public Function Duplicar(ALODINE_ID As Long, ALODINE_ID_NUEVO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO ALODINE_NORMAS " & _
                   " SELECT " & ALODINE_ID_NUEVO & ",NORMA_ID,ORDEN FROM ALODINE_NORMAS " & _
                   "  WHERE ALODINE_ID = " & ALODINE_ID
        execute_bd consulta
        Duplicar = True
        Exit Function
fallo:
        Duplicar = False
        MsgBox "Error al Duplicar (clsALODINE_NORMAS)", vbCritical, Err.Description
End Function
Public Function Listado(alodine As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.NORMA_ID, CONCAT(B.CODIGO, ' (',B.NOMBRE,')') " & _
                   " FROM ALODINE_NORMAS A,CA_NORMAS B " & _
                   " WHERE A.ALODINE_ID = " & alodine & _
                   "   AND A.NORMA_ID = B.ID_NORMA " & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function


