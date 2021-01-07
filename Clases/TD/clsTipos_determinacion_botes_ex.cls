VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_determinacion_botes_ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTIPOS_DETERMINACION_BOTES_EX
'***************************************************************
Private TIPO_DETERMINACION_ID As Long
Private BOTE_EX_ID As Long
Private TIPO As String
Private ORDEN As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTIPOS_DETERMINACION_BOTES_EX
'***************************************************************
Public Property Let setTIPO_DETERMINACION_ID(ByVal dato As Long)
        TIPO_DETERMINACION_ID = dato
End Property
Public Property Let setBOTE_EX_ID(ByVal dato As Long)
        BOTE_EX_ID = dato
End Property
Public Property Let setTIPO(ByVal dato As String)
        TIPO = dato
End Property
Public Property Let setORDEN(ByVal dato As Integer)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTIPOS_DETERMINACION_BOTES_EX
'***************************************************************
Public Property Get getTIPO_DETERMINACION_ID() As Long
        getTIPO_DETERMINACION_ID = TIPO_DETERMINACION_ID
End Property
Public Property Get getBOTE_EX_ID() As Long
        getBOTE_EX_ID = BOTE_EX_ID
End Property
Public Property Get getTIPO() As String
        getTIPO = TIPO
End Property
Public Property Get getORDEN() As Integer
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTIPOS_DETERMINACION_BOTES_EX
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM TIPOS_DETERMINACION_BOTES_EX WHERE TIPO_DETERMINACION_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                TIPO_DETERMINACION_ID = rs("TIPO_DETERMINACION_ID")
                BOTE_EX_ID = rs("BOTE_EX_ID")
                TIPO = rs("TIPO")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTipos_determinacion_botes_ex)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO TIPOS_DETERMINACION_BOTES_EX " & _
                   "  VALUES (" & _
                        TIPO_DETERMINACION_ID & "," & BOTE_EX_ID & ",'" & TIPO & "'," & ORDEN & _
                ")"
        execute_bd consulta
        Insertar = TIPO_DETERMINACION_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTipos_determinacion_botes_ex)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM TIPOS_DETERMINACION_BOTES_EX " & _
                "    WHERE TIPO_DETERMINACION_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTipos_determinacion_botes_ex)", vbCritical, Err.Description
End Function
Public Function Listado(TD As Long) As ADODB.RecordSet
        Dim consulta As String
'        consulta = "SELECT B.ID_BOTE_EX,D.NOMBRE,B.FECHA_CADUCIDAD,A.TIPO " & _
'                   " FROM TIPOS_DETERMINACION_BOTES_EX A, BOTES_EX B, TIPOS_BOTE_EX C, TIPOS_REACTIVO_EX D " & _
'                   " WHERE A.TIPO_DETERMINACION_ID = " & TD & _
'                   "   AND A.BOTE_EX_ID = B.ID_BOTE_EX " & _
'                   "   AND B.TIPO_BOTE_EX_ID = C.ID_TIPO_BOTE_EX " & _
'                   "   AND C.TIPO_REACTIVO_EX_ID = D.ID_TIPO_REACTIVO_EX " & _
'                   " ORDER BY A.ORDEN"
        consulta = "SELECT A.BOTE_EX_ID,A.TIPO " & _
                   " FROM TIPOS_DETERMINACION_BOTES_EX A " & _
                   " WHERE A.TIPO_DETERMINACION_ID = " & TD & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function
