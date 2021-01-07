VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlasma_Reactivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSplasma_reactivos
'***************************************************************
Private MUESTRA_ID As Long
Private BOTE_EX_ID As Long
Private tipo As String
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSplasma_reactivos
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setBOTE_EX_ID(ByVal dato As Long)
        BOTE_EX_ID = dato
End Property
Public Property Let setTIPO(ByVal dato As String)
        tipo = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSplasma_reactivos
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getBOTE_EX_ID() As Long
        getBOTE_EX_ID = BOTE_EX_ID
End Property
Public Property Get getTIPO() As String
        getTIPO = tipo
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSplasma_reactivos
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM plasma_reactivos WHERE MUESTRA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                BOTE_EX_ID = rs("BOTE_EX_ID")
                tipo = rs("TIPO")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsplasma_reactivos)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = " INSERT INTO plasma_reactivos VALUES (" & _
                     MUESTRA_ID & "," & BOTE_EX_ID & ",'" & tipo & "'," & ORDEN & ")"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al Insertar_Determinacion (clsDeterminaciones_equipos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM plasma_reactivos " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsplasma_reactivos)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.BOTE_EX_ID,A.TIPO,A.ORDEN " & _
                   " FROM plasma_reactivos A " & _
                   " WHERE A.MUESTRA_ID = " & ID & _
                   " ORDER BY A.ORDEN"
        
'        consulta = "SELECT B.ID_BOTE_EX,D.NOMBRE,B.FECHA_CADUCIDAD " & _
'                   " FROM plasma_reactivos A, BOTES_EX B, TIPOS_BOTE_EX C, TIPOS_REACTIVO_EX D " & _
'                   " WHERE A.MUESTRA_ID = " & ID & _
'                   "   AND A.BOTE_EX_ID = B.ID_BOTE_EX " & _
'                   "   AND B.TIPO_BOTE_EX_ID = C.ID_TIPO_BOTE_EX " & _
'                   "   AND C.TIPO_REACTIVO_EX_ID = D.ID_TIPO_REACTIVO_EX " & _
'                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function




