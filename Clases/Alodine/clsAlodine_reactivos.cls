VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlodine_reactivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsAlodine_normas
'***************************************************************
Private ALODINE_ID As Long
Private BOTE_EX_ID As Long
Private VOLUMEN As String
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsAlodine_normas
'***************************************************************
Public Property Let setALODINE_ID(ByVal dato As Long)
        ALODINE_ID = dato
End Property
Public Property Let setBOTE_EX_ID(ByVal dato As Long)
        BOTE_EX_ID = dato
End Property
Public Property Let setVOLUMEN(ByVal dato As String)
        VOLUMEN = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsAlodine_normas
'***************************************************************
Public Property Get getALODINE_ID() As Long
        getALODINE_ID = ALODINE_ID
End Property
Public Property Get getBOTE_EX_ID() As Long
        getBOTE_EX_ID = BOTE_EX_ID
End Property
Public Property Get getVOLUMEN() As String
        getVOLUMEN = VOLUMEN
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsAlodine_normas
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ALODINE_REACTIVOS WHERE ALODINE_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ALODINE_ID = rs("ALODINE_ID")
                BOTE_EX_ID = rs("BOTE_EX_ID")
                VOLUMEN = rs("VOLUMEN")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsAlodine_normas)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO ALODINE_REACTIVOS " & _
                   "  VALUES (" & _
                        ALODINE_ID & "," & BOTE_EX_ID & "," & "'" & VOLUMEN & "'" & "," & _
                        ORDEN & _
                ")"
        execute_bd consulta
        Insertar = ALODINE_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsAlodine_normas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ALODINE_REACTIVOS SET " & _
                        " BOTE_EX_ID = " & BOTE_EX_ID & "," & _
                        " VOLUMEN = '" & VOLUMEN & "'," & _
                        " ORDEN = " & ORDEN & "" & _
                " WHERE ALODINE_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsAlodine_normas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ALODINE_REACTIVOS " & _
                "    WHERE ALODINE_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsAlodine_normas)", vbCritical, Err.Description
End Function
Public Function Listado(alodine As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.BOTE_EX_ID, D.NOMBRE, A.VOLUMEN, B.FECHA_CADUCIDAD " & _
                   " FROM ALODINE_REACTIVOS A,BOTES_EX B, TIPOS_BOTE_EX C, TIPOS_REACTIVO_EX D " & _
                   " WHERE A.ALODINE_ID = " & alodine & _
                   "   AND A.BOTE_EX_ID = B.ID_BOTE_EX " & _
                   "   AND B.TIPO_BOTE_EX_ID = C.ID_TIPO_BOTE_EX " & _
                   "   AND C.TIPO_REACTIVO_EX_ID = D.ID_TIPO_REACTIVO_EX " & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function
